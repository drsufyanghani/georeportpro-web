from __future__ import annotations

import io
import os
import math
import tempfile
from dataclasses import dataclass, asdict
from datetime import date
from math import tan, radians, exp, pi
from typing import Optional, List, Dict, Any, Tuple

import numpy as np
import pandas as pd

from matplotlib.figure import Figure

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ============================================================
# STANDARDS / REFERENCES
# ============================================================

REFERENCE_STANDARDS = [
    ("Atterberg Limits", "Liquid limit, plastic limit, plasticity index", "IS 2720 (Part 5): Determination of liquid and plastic limit"),
    ("Soil Classification", "ISCS symbol using grain-size fractions and plasticity chart", "IS 1498: Classification and identification of soils for general engineering purposes"),
    ("Compaction", "Dry density, OMC and MDD from laboratory compaction data", "IS 2720 (Part 7 - light compaction) / IS 2720 (Part 8 - heavy compaction), as applicable to the test method"),
    ("CBR", "CBR at 2.5 mm and 5.0 mm using standard loads", "IS 2720 (Part 16): Laboratory determination of CBR"),
    ("SPT", "N-value from final 300 mm penetration, correction factor entered by engineer, dilatancy correction where applicable", "IS 2131: Standard penetration test for soils"),
    ("Bearing Capacity", "Preliminary Terzaghi-style shallow foundation bearing estimate; not a full substitute for final IS 6403 design assessment", "IS 6403 should govern final bearing capacity evaluation"),
    ("Figures", "Borehole log, PSD, plasticity chart, compaction curve, CBR curve, SPT profile", "Visualization of imported engineering data; not a separate design standard"),
]

ENGINEERING_DISCLAIMER = (
    "Automated outputs are intended for preliminary engineering review and reporting support only. "
    "All imported data, assumptions, correlations, plots, classifications, corrections, design values and "
    "recommendations must be checked, validated and approved by a qualified geotechnical engineer before client issue."
)

WSP_RED = "#E10600"
DARK = "#101820"
TEXT = "#111827"
GREY = "#667085"


# ============================================================
# DATA MODELS
# ============================================================

@dataclass
class ProjectInfo:
    project_name: str = ""
    client_name: str = ""
    location: str = ""
    report_number: str = ""
    prepared_by: str = ""
    checked_by: str = ""
    approved_by: str = ""
    report_date: str = date.today().isoformat()
    remarks: str = ""


@dataclass
class AtterbergResult:
    liquid_limit: float
    plastic_limit: float
    plasticity_index: float
    liquidity_index: Optional[float]
    consistency_index: Optional[float]
    plasticity_category: str
    interpretation: str


@dataclass
class SoilClassificationResult:
    soil_symbol: str
    soil_group: str
    description: str
    notes: str


@dataclass
class CompactionPoint:
    water_content_pct: float
    wet_density_g_cc: float
    dry_density_g_cc: float


@dataclass
class CompactionResult:
    points: List[CompactionPoint]
    observed_omc_pct: float
    observed_mdd_g_cc: float
    fitted_omc_pct: Optional[float]
    fitted_mdd_g_cc: Optional[float]
    interpretation: str
    qc_note: str


@dataclass
class CBRResult:
    cbr_2_5_mm: float
    cbr_5_0_mm: float
    design_cbr: float
    selected_penetration_mm: float
    load_2_5_kg: float
    load_5_0_kg: float
    interpretation: str
    qc_note: str


@dataclass
class SPTResult:
    raw_n: int
    overburden_corrected_n: float
    dilatancy_corrected_n: float
    description: str
    warning: str


@dataclass
class BearingCapacityResult:
    foundation_type: str
    nc: float
    nq: float
    ngamma: float
    ultimate_bearing_capacity_kpa: float
    safe_bearing_capacity_kpa: float
    interpretation: str
    limitation: str


# ============================================================
# CALCULATION BACKEND
# ============================================================

def safe_float(value, default=None):
    try:
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def to_float(text: str, field_name: str, minimum: Optional[float] = None) -> float:
    if text is None or str(text).strip() == "":
        raise ValueError(f"{field_name} is required.")
    try:
        value = float(str(text).strip())
    except Exception:
        raise ValueError(f"{field_name} must be numeric.")
    if minimum is not None and value < minimum:
        raise ValueError(f"{field_name} must be >= {minimum}.")
    return value


def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    cols = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in cols:
            return cols[key]
    for cand in candidates:
        key = cand.strip().lower().replace(" ", "").replace("_", "")
        for c in df.columns:
            ck = str(c).strip().lower().replace(" ", "").replace("_", "")
            if key == ck or key in ck or ck in key:
                return c
    return None


def calculate_atterberg(ll: float, pl: float, water_content: Optional[float] = None) -> AtterbergResult:
    if ll < 0 or pl < 0:
        raise ValueError("Liquid limit and plastic limit cannot be negative.")
    if ll < pl:
        raise ValueError("Liquid limit cannot be less than plastic limit.")
    pi_val = ll - pl
    li = None
    ci = None
    if water_content is not None and pi_val != 0:
        li = (water_content - pl) / pi_val
        ci = (ll - water_content) / pi_val

    if pi_val <= 0:
        category = "Non-plastic"
        text = "The soil is non-plastic based on the supplied Atterberg limits."
    elif ll < 35:
        category = "Low compressibility range"
        text = f"Atterberg limits indicate PI = {pi_val:.2f}% and LL = {ll:.2f}%."
    elif ll <= 50:
        category = "Intermediate compressibility range"
        text = f"Atterberg limits indicate PI = {pi_val:.2f}% and LL = {ll:.2f}%."
    else:
        category = "High compressibility range"
        text = f"Atterberg limits indicate PI = {pi_val:.2f}% and LL = {ll:.2f}%."

    return AtterbergResult(ll, pl, pi_val, li, ci, category, text)


def a_line_pi(ll: float) -> float:
    return 0.73 * (ll - 20.0)


def _compressibility_symbol(ll: float) -> Tuple[str, str]:
    if ll < 35:
        return "L", "low compressibility"
    if ll <= 50:
        return "I", "intermediate compressibility"
    return "H", "high compressibility"


def classify_fine_grained_soil(ll: float, pi_val: float) -> SoilClassificationResult:
    if pi_val <= 0:
        return SoilClassificationResult(
            "ML",
            "Silt of low compressibility / non-plastic fine-grained soil",
            "Fine-grained soil with negligible plasticity based on supplied Atterberg limits.",
            "Final classification should be checked against the full IS 1498 plasticity chart and grain-size data."
        )
    comp, comp_text = _compressibility_symbol(ll)
    is_clay = pi_val >= a_line_pi(ll)
    prefix = "C" if is_clay else "M"
    soil_word = "Clay" if is_clay else "Silt"
    symbol = f"{prefix}{comp}"
    group = f"{soil_word} of {comp_text}"
    desc = f"{group} based on LL = {ll:.1f}% and PI = {pi_val:.1f}% with comparison against the A-line."
    return SoilClassificationResult(symbol, group, desc, "Classification follows IS 1498-style plasticity chart logic; verify with full gradation and test records.")


def classify_soil(gravel: float, sand: float, fines: float, ll: float, pi_val: float) -> SoilClassificationResult:
    total = gravel + sand + fines
    if abs(total - 100.0) > 2.0:
        raise ValueError("Gravel, sand, and fines should approximately sum to 100%.")
    if fines >= 50:
        return classify_fine_grained_soil(ll, pi_val)

    dominant = "gravel" if gravel >= sand else "sand"
    prefix = "G" if dominant == "gravel" else "S"

    if fines < 5:
        return SoilClassificationResult(
            f"{prefix}W/{prefix}P",
            f"Clean {dominant}",
            f"Predominantly {dominant} with fines < 5%. Cu and Cc are required to distinguish well-graded and poorly graded classes.",
            "Do not finalize GW/GP or SW/SP without gradation coefficients."
        )

    fine = classify_fine_grained_soil(ll, pi_val)
    fine_letter = "C" if fine.soil_symbol.startswith("C") else "M"

    if 5 <= fines <= 12:
        symbol = f"{prefix}W-{prefix}{fine_letter} / {prefix}P-{prefix}{fine_letter}"
        group = f"{dominant.capitalize()} with 5-12% fines; dual symbol requires Cu/Cc confirmation"
    else:
        symbol = f"{prefix}{fine_letter}"
        group = f"{dominant.capitalize()} with {('clayey' if fine_letter == 'C' else 'silty')} fines"

    return SoilClassificationResult(
        symbol,
        group,
        f"Predominantly {dominant} with {fines:.1f}% fines. Fines plasticity indicates {fine.group}.",
        "Coarse-grained classification should be finalized using full IS 1498 procedure including Cu and Cc where required."
    )


def dry_density(wet_density: float, water_content: float) -> float:
    if wet_density <= 0:
        raise ValueError("Wet density must be positive.")
    return wet_density / (1 + water_content / 100.0)


def analyze_compaction(water_contents: List[float], wet_densities: List[float]) -> CompactionResult:
    if len(water_contents) != len(wet_densities):
        raise ValueError("Water content and wet density lists must have the same length.")
    if len(water_contents) < 3:
        raise ValueError("At least three compaction points are required.")
    points = [CompactionPoint(float(w), float(wet), dry_density(float(wet), float(w))) for w, wet in zip(water_contents, wet_densities)]
    points = sorted(points, key=lambda p: p.water_content_pct)
    dry_vals = [p.dry_density_g_cc for p in points]
    max_idx = int(np.argmax(dry_vals))
    observed_omc = points[max_idx].water_content_pct
    observed_mdd = points[max_idx].dry_density_g_cc

    fitted_omc = None
    fitted_mdd = None
    qc_notes = []
    try:
        x = np.array([p.water_content_pct for p in points], dtype=float)
        y = np.array([p.dry_density_g_cc for p in points], dtype=float)
        coeffs = np.polyfit(x, y, 2)
        a, b, _ = coeffs
        if a < 0:
            candidate_omc = -b / (2 * a)
            if min(x) <= candidate_omc <= max(x):
                fitted_omc = float(candidate_omc)
                fitted_mdd = float(np.polyval(coeffs, fitted_omc))
            else:
                qc_notes.append("Fitted OMC lies outside the measured water-content range and is not reported.")
        else:
            qc_notes.append("Quadratic fit is not concave downward; fitted OMC/MDD is not reported.")
    except Exception:
        qc_notes.append("Fitted compaction curve could not be generated.")

    text = f"Adopted observed MDD = {observed_mdd:.3f} g/cc at observed OMC = {observed_omc:.2f}%."
    if fitted_omc is not None:
        text += f" Indicative fitted MDD = {fitted_mdd:.3f} g/cc at fitted OMC = {fitted_omc:.2f}%."
        qc_notes.append("Observed OMC/MDD is adopted as the primary reported value; fitted value is indicative only.")

    return CompactionResult(points, observed_omc, observed_mdd, fitted_omc, fitted_mdd, text, " ".join(qc_notes))


def analyze_cbr_from_key_loads(load_2_5: float, load_5_0: float, adopt_5_if_higher_confirmed: bool = False) -> CBRResult:
    if load_2_5 < 0 or load_5_0 < 0:
        raise ValueError("CBR loads cannot be negative.")
    cbr_2_5 = load_2_5 / 1370.0 * 100.0
    cbr_5_0 = load_5_0 / 2055.0 * 100.0

    qc = "CBR computed using standard loads: 1370 kg at 2.5 mm and 2055 kg at 5.0 mm."
    if cbr_5_0 > cbr_2_5:
        if adopt_5_if_higher_confirmed:
            design, pen = cbr_5_0, 5.0
            qc += " CBR at 5.0 mm is higher and has been adopted only because repeat-test confirmation was selected."
        else:
            design, pen = cbr_2_5, 2.5
            qc += " CBR at 5.0 mm is higher; by default, the 2.5 mm value is retained unless repeat testing confirms adoption of 5.0 mm."
    else:
        design, pen = cbr_2_5, 2.5
        qc += " CBR at 2.5 mm governs."

    text = f"CBR at 2.5 mm = {cbr_2_5:.2f}%, CBR at 5.0 mm = {cbr_5_0:.2f}%. Adopted design CBR = {design:.2f}% at {pen:.1f} mm."
    return CBRResult(cbr_2_5, cbr_5_0, design, pen, load_2_5, load_5_0, text, qc)


def analyze_cbr_from_curve(df: pd.DataFrame, adopt_5_if_higher_confirmed: bool = False) -> Optional[CBRResult]:
    if df is None or df.empty:
        return None
    pen_col = find_col(df, ["Penetration (mm)", "Penetration"])
    load_col = find_col(df, ["Load (kg)", "Load"])
    if not pen_col or not load_col:
        return None
    d = df[[pen_col, load_col]].copy()
    d[pen_col] = pd.to_numeric(d[pen_col], errors="coerce")
    d[load_col] = pd.to_numeric(d[load_col], errors="coerce")
    d = d.dropna().sort_values(pen_col)
    if d.empty:
        return None
    if not (d[pen_col].min() <= 2.5 <= d[pen_col].max() and d[pen_col].min() <= 5.0 <= d[pen_col].max()):
        return None
    load25 = float(np.interp(2.5, d[pen_col], d[load_col]))
    load50 = float(np.interp(5.0, d[pen_col], d[load_col]))
    return analyze_cbr_from_key_loads(load25, load50, adopt_5_if_higher_confirmed)


def analyze_spt(b1: int, b2: int, b3: int, correction_factor: float = 1.0, saturated_fine_sand: bool = False) -> SPTResult:
    if min(b1, b2, b3) < 0:
        raise ValueError("SPT blow counts cannot be negative.")
    n = int(b2 + b3)
    n_corr = n * correction_factor
    n_dil = n_corr
    warning = ""
    if saturated_fine_sand and n_corr > 15:
        n_dil = 15 + 0.5 * (n_corr - 15)
        warning = "Dilatancy correction applied because corrected N-value exceeds 15 in saturated fine sand/silt."
    if b3 >= 50 or n >= 50:
        warning = (warning + " " if warning else "") + "Possible refusal or very dense/hard stratum; review field record."

    if n_dil < 4:
        desc = "Very loose"
    elif n_dil < 10:
        desc = "Loose"
    elif n_dil < 30:
        desc = "Medium dense"
    elif n_dil < 50:
        desc = "Dense"
    else:
        desc = "Very dense"

    return SPTResult(n, n_corr, n_dil, desc, warning)


def bearing_capacity_factors(phi_deg: float):
    if phi_deg < 0 or phi_deg > 50:
        raise ValueError("Friction angle should generally be between 0 and 50 degrees for this simplified calculation.")
    if abs(phi_deg) < 1e-9:
        return 5.7, 1.0, 0.0
    phi_rad = radians(phi_deg)
    nq = exp(pi * tan(phi_rad)) * (tan(radians(45) + phi_rad / 2) ** 2)
    nc = (nq - 1) / tan(phi_rad)
    ngamma = 1.5 * (nq - 1) * tan(phi_rad)
    return nc, nq, ngamma


def calculate_terzaghi(ftype: str, b: float, d: float, c: float, gamma: float, phi_deg: float, fos: float) -> BearingCapacityResult:
    if b <= 0:
        raise ValueError("Foundation width must be positive.")
    if d < 0:
        raise ValueError("Foundation depth cannot be negative.")
    if gamma <= 0:
        raise ValueError("Unit weight must be positive.")
    if fos <= 0:
        raise ValueError("FOS must be positive.")

    f = ftype.lower().strip()
    nc, nq, ngamma = bearing_capacity_factors(phi_deg)
    q = gamma * d

    if f == "strip":
        qult = c * nc + q * nq + 0.5 * gamma * b * ngamma
    elif f == "square":
        qult = 1.3 * c * nc + q * nq + 0.4 * gamma * b * ngamma
    elif f == "circular":
        qult = 1.3 * c * nc + q * nq + 0.3 * gamma * b * ngamma
    elif f == "rectangular":
        qult = 1.2 * c * nc + q * nq + 0.4 * gamma * b * ngamma
    else:
        raise ValueError("Foundation type must be Strip, Square, Rectangular, or Circular.")

    safe = qult / fos
    text = f"Preliminary Terzaghi calculation for {f} footing gives ultimate bearing capacity = {qult:.2f} kPa and safe bearing capacity = {safe:.2f} kPa with FOS = {fos:.2f}."
    limitation = "This is a preliminary shear-based estimate only. Final bearing capacity and allowable pressure must be checked using project-specific IS 6403 requirements, groundwater effects, inclination/eccentricity, settlement and engineering judgement."
    return BearingCapacityResult(f, nc, nq, ngamma, qult, safe, text, limitation)


def psd_qc_and_coefficients(df: pd.DataFrame) -> Tuple[List[str], Dict[str, Optional[float]]]:
    warnings = []
    coeffs = {"D10": None, "D30": None, "D60": None, "Cu": None, "Cc": None}
    if df is None or df.empty:
        return ["PSD data not available."], coeffs

    size_col = find_col(df, ["Sieve Size (mm)", "Particle Size (mm)", "Sieve Size"])
    pp_col = find_col(df, ["Percent Passing (%)", "Percent Passing", "% Passing"])
    if not size_col or not pp_col:
        return ["PSD required columns not found."], coeffs

    d = df[[size_col, pp_col]].copy()
    d[size_col] = pd.to_numeric(d[size_col], errors="coerce")
    d[pp_col] = pd.to_numeric(d[pp_col], errors="coerce")
    d = d.dropna().sort_values(size_col, ascending=False)

    if d.empty:
        return ["PSD data has no valid numeric rows."], coeffs
    if (d[size_col] <= 0).any():
        warnings.append("PSD contains zero or negative sieve size.")
    if ((d[pp_col] < 0) | (d[pp_col] > 100)).any():
        warnings.append("PSD contains percent passing outside 0-100%.")
    if len(d) >= 2 and (d[pp_col].diff().dropna() > 1e-6).any():
        warnings.append("PSD is non-monotonic: percent passing increases as sieve size decreases. Check cumulative sieve data.")

    # D-values by interpolation in log-size vs percent passing.
    # d is coarse-to-fine. For interpolation use percent passing ascending.
    try:
        g = d.drop_duplicates(subset=[pp_col]).sort_values(pp_col)
        pp = g[pp_col].values.astype(float)
        log_size = np.log10(g[size_col].values.astype(float))
        for target in [10, 30, 60]:
            if pp.min() <= target <= pp.max():
                val = 10 ** float(np.interp(target, pp, log_size))
                coeffs[f"D{target}"] = val
        if coeffs["D10"] and coeffs["D60"] and coeffs["D10"] > 0:
            coeffs["Cu"] = coeffs["D60"] / coeffs["D10"]
        if coeffs["D10"] and coeffs["D30"] and coeffs["D60"] and coeffs["D10"] > 0 and coeffs["D60"] > 0:
            coeffs["Cc"] = (coeffs["D30"] ** 2) / (coeffs["D10"] * coeffs["D60"])
    except Exception:
        warnings.append("Could not compute D10/D30/D60 from PSD data.")

    if not warnings:
        warnings.append("PSD passed basic monotonic and range checks.")
    return warnings, coeffs

# ============================================================
# WEB-APPLICATION HELPERS
# ============================================================

DEFAULT_INPUTS = {
    "classification": {
        "ll": 42.0, "pl": 21.0, "water_content": 25.0,
        "gravel": 5.0, "sand": 35.0, "fines": 60.0,
    },
    "compaction": {
        "points": [
            {"water_content": 8.0, "wet_density": 1.82},
            {"water_content": 10.0, "wet_density": 1.92},
            {"water_content": 12.0, "wet_density": 2.02},
            {"water_content": 14.0, "wet_density": 2.00},
            {"water_content": 16.0, "wet_density": 1.96},
        ]
    },
    "cbr": {"load_2_5": 105.0, "load_5_0": 160.0, "adopt_5_if_higher_confirmed": False},
    "spt": {"b1": 8, "b2": 15, "b3": 20, "correction_factor": 1.0, "saturated_fine_sand": False},
    "bearing": {"foundation_type": "Strip", "width": 1.5, "depth": 1.5, "cohesion": 25.0, "unit_weight": 18.0, "phi": 28.0, "fos": 3.0},
}


def new_state() -> Dict[str, Any]:
    import copy
    return {
        "project_info": ProjectInfo(),
        "atterberg_result": None,
        "classification_result": None,
        "compaction_result": None,
        "cbr_result": None,
        "spt_result": None,
        "spt_results": [],
        "bearing_result": None,
        "data": {},
        "excel_name": None,
        "inputs": copy.deepcopy(DEFAULT_INPUTS),
        "qc_warnings": [],
    }


def dataclass_or_none(value):
    return asdict(value) if value is not None else None


def df_preview(df: pd.DataFrame, max_rows: int = 40) -> Dict[str, Any]:
    if df is None or df.empty:
        return {"columns": [], "rows": []}
    show = df.head(max_rows).copy()
    # to_json safely converts NumPy scalars, dates and NaN values.
    import json
    rows = json.loads(show.to_json(orient="records", date_format="iso"))
    return {"columns": [str(c) for c in show.columns], "rows": rows}


def unique_boreholes(state: Dict[str, Any]) -> List[str]:
    values = []
    for sheet in ["Boreholes", "Lab_Summary", "PSD", "Compaction", "CBR", "SPT"]:
        df = state["data"].get(sheet)
        if df is not None and not df.empty:
            col = find_col(df, ["Borehole ID"])
            if col:
                values += [str(x) for x in df[col].dropna().unique()]
    return list(dict.fromkeys(values))


def unique_samples(state: Dict[str, Any], borehole_id: Optional[str] = None) -> List[str]:
    values = []
    for sheet in ["Lab_Summary", "PSD", "Compaction", "CBR"]:
        df = state["data"].get(sheet)
        if df is None or df.empty:
            continue
        d = df.copy()
        bh_col = find_col(d, ["Borehole ID"])
        sample_col = find_col(d, ["Sample ID", "Sample"])
        if borehole_id and bh_col:
            d = d[d[bh_col].astype(str) == str(borehole_id)]
        if sample_col:
            values += [str(x) for x in d[sample_col].dropna().unique()]
    return list(dict.fromkeys(values))


def filtered_sheet(state: Dict[str, Any], sheet: str, borehole_id: Optional[str] = None, sample_id: Optional[str] = None) -> pd.DataFrame:
    df = state["data"].get(sheet)
    if df is None or df.empty:
        return pd.DataFrame()
    d = df.copy()
    bh_col = find_col(d, ["Borehole ID"])
    sample_col = find_col(d, ["Sample ID", "Sample"])
    if borehole_id and bh_col:
        d = d[d[bh_col].astype(str) == str(borehole_id)]
    if sample_id and sample_col:
        d = d[d[sample_col].astype(str) == str(sample_id)]
    return d


def first_sample_for_sheet(state: Dict[str, Any], sheet: str, borehole_id: Optional[str] = None) -> Optional[str]:
    d = filtered_sheet(state, sheet, borehole_id, None)
    sample_col = find_col(d, ["Sample ID", "Sample"])
    if d.empty or not sample_col:
        return None
    vals = d[sample_col].dropna().astype(str).unique()
    return vals[0] if len(vals) else None


def compute_all_spt_results(state: Dict[str, Any]) -> None:
    state["spt_results"] = []
    df = state["data"].get("SPT")
    if df is None or df.empty:
        return
    c_b1 = find_col(df, ["Blows 0-150"])
    c_b2 = find_col(df, ["Blows 150-300"])
    c_b3 = find_col(df, ["Blows 300-450"])
    c_cf = find_col(df, ["Correction Factor"])
    c_sat = find_col(df, ["Saturated Fine Sand/Silt"])
    c_depth = find_col(df, ["Depth (m)"])
    c_bh = find_col(df, ["Borehole ID"])
    if not (c_b1 and c_b2 and c_b3):
        return
    for _, row in df.iterrows():
        try:
            sat = str(row[c_sat]).strip().lower() == "yes" if c_sat else False
            result = analyze_spt(
                int(row[c_b1]), int(row[c_b2]), int(row[c_b3]),
                safe_float(row[c_cf], 1.0) if c_cf else 1.0, sat,
            )
            state["spt_results"].append({
                "borehole": row[c_bh] if c_bh else "",
                "depth": safe_float(row[c_depth], None),
                "result": result,
            })
        except Exception:
            continue


def run_qc_checks(state: Dict[str, Any]) -> List[str]:
    warnings = []
    data = state["data"]
    required_sheets = ["Project", "Boreholes", "SPT", "Lab_Summary", "PSD", "Compaction", "CBR", "Bearing_Input"]
    for sheet in required_sheets:
        if sheet not in data:
            warnings.append(f"Missing sheet: {sheet}")

    expected_cols = {
        "Boreholes": ["Borehole ID", "From Depth (m)", "To Depth (m)", "Soil Description"],
        "Lab_Summary": ["LL (%)", "PL (%)", "Gravel (%)", "Sand (%)", "Fines (%)"],
        "Compaction": ["Sample ID", "Water Content (%)", "Wet Density (g/cc)"],
        "CBR": ["Sample ID", "Penetration (mm)", "Load (kg)"],
        "SPT": ["Borehole ID", "Depth (m)", "Blows 150-300", "Blows 300-450"],
        "PSD": ["Sample ID", "Sieve Size (mm)", "Percent Passing (%)"],
    }
    for sheet, cols in expected_cols.items():
        df = data.get(sheet)
        if df is None or df.empty:
            continue
        for col in cols:
            if not find_col(df, [col]):
                warnings.append(f"{sheet}: expected column missing or not recognized -> {col}")

    lab = data.get("Lab_Summary")
    if lab is not None and not lab.empty:
        llc = find_col(lab, ["LL (%)", "Liquid Limit"])
        plc = find_col(lab, ["PL (%)", "Plastic Limit"])
        if llc and plc:
            bad = lab[pd.to_numeric(lab[llc], errors="coerce") < pd.to_numeric(lab[plc], errors="coerce")]
            if not bad.empty:
                warnings.append("Lab_Summary: one or more rows have LL lower than PL.")
        gr = find_col(lab, ["Gravel (%)"])
        sa = find_col(lab, ["Sand (%)"])
        fi = find_col(lab, ["Fines (%)"])
        if gr and sa and fi:
            total = pd.to_numeric(lab[gr], errors="coerce") + pd.to_numeric(lab[sa], errors="coerce") + pd.to_numeric(lab[fi], errors="coerce")
            if ((total < 98) | (total > 102)).any():
                warnings.append("Lab_Summary: gravel+sand+fines is outside 100±2% for one or more samples.")

    comp = data.get("Compaction")
    if comp is not None and not comp.empty:
        wc = find_col(comp, ["Water Content (%)"])
        wet = find_col(comp, ["Wet Density (g/cc)"])
        if wc and wet and (pd.to_numeric(comp[wc], errors="coerce").isna().any() or pd.to_numeric(comp[wet], errors="coerce").isna().any()):
            warnings.append("Compaction: non-numeric water content or wet density values detected.")

    cbr = data.get("CBR")
    if cbr is not None and not cbr.empty:
        pen = find_col(cbr, ["Penetration (mm)", "Penetration"])
        load = find_col(cbr, ["Load (kg)", "Load"])
        if pen and load and cbr[[pen, load]].dropna().empty:
            warnings.append("CBR: no valid penetration-load rows found.")

    psd = data.get("PSD")
    if psd is not None and not psd.empty:
        sample_col = find_col(psd, ["Sample ID", "Sample"])
        if sample_col:
            for sid, grp in psd.groupby(sample_col):
                msgs, _ = psd_qc_and_coefficients(grp)
                for m in msgs:
                    if "passed" not in m.lower():
                        warnings.append(f"PSD {sid}: {m}")
        else:
            msgs, _ = psd_qc_and_coefficients(psd)
            for m in msgs:
                if "passed" not in m.lower():
                    warnings.append(f"PSD: {m}")
    state["qc_warnings"] = warnings
    return warnings


def auto_populate_from_workbook(state: Dict[str, Any]) -> None:
    data = state["data"]

    df = data.get("Project")
    if df is not None and not df.empty and df.shape[1] >= 2:
        try:
            mp = dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1]))
            state["project_info"] = ProjectInfo(
                project_name=str(mp.get("Project Name", "")),
                client_name=str(mp.get("Client Name", "")),
                location=str(mp.get("Location", "")),
                report_number=str(mp.get("Report Number", "")),
                prepared_by=str(mp.get("Prepared By", "")),
                checked_by=str(mp.get("Checked By", "")),
                approved_by=str(mp.get("Approved By", "")),
                report_date=str(mp.get("Report Date", date.today().isoformat())),
                remarks=str(mp.get("Remarks", "")),
            )
        except Exception:
            pass

    lab = data.get("Lab_Summary")
    if lab is not None and not lab.empty:
        row = lab.iloc[0]
        llc = find_col(lab, ["LL (%)", "Liquid Limit"])
        plc = find_col(lab, ["PL (%)", "Plastic Limit"])
        wc = find_col(lab, ["Water Content (%)", "Natural Water Content (%)"])
        gr = find_col(lab, ["Gravel (%)"])
        sa = find_col(lab, ["Sand (%)"])
        fi = find_col(lab, ["Fines (%)"])
        try:
            inp = state["inputs"]["classification"]
            if llc: inp["ll"] = safe_float(row[llc], inp["ll"])
            if plc: inp["pl"] = safe_float(row[plc], inp["pl"])
            if wc: inp["water_content"] = safe_float(row[wc], inp["water_content"])
            if gr: inp["gravel"] = safe_float(row[gr], inp["gravel"])
            if sa: inp["sand"] = safe_float(row[sa], inp["sand"])
            if fi: inp["fines"] = safe_float(row[fi], inp["fines"])
            att = calculate_atterberg(inp["ll"], inp["pl"], inp["water_content"])
            cls = classify_soil(inp["gravel"], inp["sand"], inp["fines"], inp["ll"], att.plasticity_index)
            state["atterberg_result"] = att
            state["classification_result"] = cls
        except Exception:
            pass

    comp = data.get("Compaction")
    if comp is not None and not comp.empty:
        try:
            sample = first_sample_for_sheet(state, "Compaction")
            d = filtered_sheet(state, "Compaction", None, sample)
            wc = find_col(d, ["Water Content (%)"])
            wet = find_col(d, ["Wet Density (g/cc)"])
            if wc and wet:
                d = d[[wc, wet]].dropna()
                points = [{"water_content": float(r[wc]), "wet_density": float(r[wet])} for _, r in d.iterrows()]
                state["inputs"]["compaction"] = {"points": points}
                state["compaction_result"] = analyze_compaction(
                    [p["water_content"] for p in points], [p["wet_density"] for p in points]
                )
        except Exception:
            pass

    cbr = data.get("CBR")
    if cbr is not None and not cbr.empty:
        try:
            sample = first_sample_for_sheet(state, "CBR")
            res = analyze_cbr_from_curve(filtered_sheet(state, "CBR", None, sample), False)
            if res:
                state["cbr_result"] = res
                state["inputs"]["cbr"].update({"load_2_5": res.load_2_5_kg, "load_5_0": res.load_5_0_kg})
        except Exception:
            pass

    spt = data.get("SPT")
    if spt is not None and not spt.empty:
        try:
            r = spt.iloc[0]
            c_b1 = find_col(spt, ["Blows 0-150"])
            c_b2 = find_col(spt, ["Blows 150-300"])
            c_b3 = find_col(spt, ["Blows 300-450"])
            c_cf = find_col(spt, ["Correction Factor"])
            c_sat = find_col(spt, ["Saturated Fine Sand/Silt"])
            inp = state["inputs"]["spt"]
            if c_b1: inp["b1"] = int(safe_float(r[c_b1], 0))
            if c_b2: inp["b2"] = int(safe_float(r[c_b2], 0))
            if c_b3: inp["b3"] = int(safe_float(r[c_b3], 0))
            if c_cf: inp["correction_factor"] = safe_float(r[c_cf], 1.0)
            if c_sat: inp["saturated_fine_sand"] = str(r[c_sat]).strip().lower() == "yes"
            state["spt_result"] = analyze_spt(**inp)
            compute_all_spt_results(state)
        except Exception:
            pass

    bi = data.get("Bearing_Input")
    if bi is not None and not bi.empty and bi.shape[1] >= 2:
        try:
            mp = dict(zip(bi.iloc[:, 0].astype(str), bi.iloc[:, 1]))
            inp = state["inputs"]["bearing"]
            inp.update({
                "foundation_type": str(mp.get("Foundation Type", "Strip")),
                "width": safe_float(mp.get("Width / Diameter B", 1.5), 1.5),
                "depth": safe_float(mp.get("Foundation Depth Df", 1.5), 1.5),
                "cohesion": safe_float(mp.get("Cohesion c", 25), 25),
                "unit_weight": safe_float(mp.get("Unit Weight gamma", 18), 18),
                "phi": safe_float(mp.get("Friction Angle phi", 28), 28),
                "fos": safe_float(mp.get("Factor of Safety", 3), 3),
            })
            state["bearing_result"] = calculate_terzaghi(
                inp["foundation_type"], inp["width"], inp["depth"], inp["cohesion"], inp["unit_weight"], inp["phi"], inp["fos"]
            )
        except Exception:
            pass

    run_qc_checks(state)


def standards_as_text() -> str:
    lines = ["REFERENCE STANDARDS / METHODS USED", "=" * 50]
    for module, logic, ref in REFERENCE_STANDARDS:
        lines.append(f"{module}: {ref}")
        lines.append(f"  Current tool logic: {logic}")
    lines += ["", "ENGINEERING DISCLAIMER", ENGINEERING_DISCLAIMER]
    return "\n".join(lines)


def build_summary_text(state: Dict[str, Any]) -> str:
    p = state["project_info"]
    lines = ["GEOTECHNICAL CALCULATION SUMMARY", "=" * 45]
    if p.project_name:
        lines += [f"Project: {p.project_name}", f"Client: {p.client_name}", f"Location: {p.location}", ""]
    bh = state["data"].get("Boreholes")
    if bh is not None:
        bid = find_col(bh, ["Borehole ID"])
        lines += ["BOREHOLES", f"Imported borehole records: {len(bh)}; Boreholes: {bh[bid].nunique() if bid else 'NA'}", ""]
    if state["atterberg_result"]:
        lines += ["ATTERBERG LIMITS", state["atterberg_result"].interpretation, ""]
    if state["classification_result"]:
        r = state["classification_result"]
        lines += ["SOIL CLASSIFICATION", f"{r.soil_symbol} - {r.soil_group}", r.description, ""]
    if state["compaction_result"]:
        r = state["compaction_result"]
        lines += ["COMPACTION", r.interpretation, r.qc_note, ""]
    if state["cbr_result"]:
        r = state["cbr_result"]
        lines += ["CBR", r.interpretation, r.qc_note, ""]
    if state["spt_result"]:
        r = state["spt_result"]
        lines += ["SPT", f"Single-test N = {r.raw_n}; corrected = {r.dilatancy_corrected_n:.2f}; condition = {r.description}", ""]
    if state["bearing_result"]:
        r = state["bearing_result"]
        lines += ["BEARING CAPACITY", r.interpretation, r.limitation, ""]
    if len(lines) <= 2:
        lines.append("No calculations completed yet.")
    lines += ["ENGINEERING REVIEW NOTE", ENGINEERING_DISCLAIMER]
    return "\n".join([str(x) for x in lines])


def state_payload(state: Dict[str, Any]) -> Dict[str, Any]:
    p = state["project_info"]
    bh = state["data"].get("Boreholes")
    bid = find_col(bh, ["Borehole ID"]) if bh is not None and not bh.empty else None
    borehole_count = int(bh[bid].nunique()) if bid else (int(len(bh)) if bh is not None else 0)
    qc = state.get("qc_warnings", [])
    cls = state["classification_result"]
    comp = state["compaction_result"]
    cbr = state["cbr_result"]
    spt = state["spt_result"]
    bearing = state["bearing_result"]
    metrics = {
        "project": p.project_name or "Not set",
        "project_subtitle": p.location or "Add details or import workbook",
        "boreholes": borehole_count if borehole_count else "--",
        "soil_class": cls.soil_symbol if cls else "--",
        "soil_subtitle": cls.soil_group if cls else "Classification pending",
        "design_cbr": f"{cbr.design_cbr:.2f}%" if cbr else "--",
        "safe_bearing": f"{bearing.safe_bearing_capacity_kpa:.1f} kPa" if bearing else "--",
        "qc_status": "Review" if qc else ("Passed" if state["data"] else "Pending"),
        "qc_subtitle": f"{len(qc)} issue(s) flagged" if qc else ("No basic issues flagged" if state["data"] else "Import data and run QC"),
    }
    return {
        "project": asdict(p),
        "inputs": state["inputs"],
        "metrics": metrics,
        "results": {
            "atterberg": dataclass_or_none(state["atterberg_result"]),
            "classification": dataclass_or_none(cls),
            "compaction": dataclass_or_none(comp),
            "cbr": dataclass_or_none(cbr),
            "spt": dataclass_or_none(spt),
            "bearing": dataclass_or_none(bearing),
        },
        "excel_name": state.get("excel_name"),
        "sheets": list(state["data"].keys()),
        "boreholes": unique_boreholes(state),
        "samples": unique_samples(state),
        "qc_warnings": qc,
        "summary_text": build_summary_text(state),
        "standards_text": standards_as_text(),
    }


def _style_axis(ax):
    ax.grid(True, alpha=0.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def _figure_to_png(fig: Figure) -> bytes:
    out = io.BytesIO()
    fig.savefig(out, format="png", dpi=180, bbox_inches="tight")
    return out.getvalue()


def _empty_plot(message: str, title: str = "") -> bytes:
    fig = Figure(figsize=(7.2, 4.2), dpi=100, tight_layout=True)
    ax = fig.add_subplot(111)
    ax.text(0.5, 0.5, message, ha="center", va="center", fontsize=11)
    if title:
        ax.set_title(title)
    ax.axis("off")
    return _figure_to_png(fig)


def plot_png(state: Dict[str, Any], kind: str, borehole: Optional[str] = None, sample: Optional[str] = None, spt_value: str = "Raw N") -> bytes:
    fig = Figure(figsize=(7.2, 4.2), dpi=100, tight_layout=True)
    ax = fig.add_subplot(111)

    if kind == "plasticity":
        ll = np.linspace(20, 100, 200)
        ax.plot(ll, 0.73 * (ll - 20), color=DARK, linewidth=1.8, label="A-line")
        ax.plot(ll, 0.9 * (ll - 8), color=WSP_RED, linestyle="--", linewidth=1.5, label="U-line")
        r = state["atterberg_result"]
        if r:
            ax.scatter([r.liquid_limit], [r.plasticity_index], s=80, color=WSP_RED, edgecolor=DARK, linewidth=0.7, label="Selected sample")
            cls = state["classification_result"]
            ax.annotate(cls.soil_symbol if cls else "Sample", (r.liquid_limit, r.plasticity_index), xytext=(6, 6), textcoords="offset points")
        ax.set_xlim(0, 100); ax.set_ylim(0, 70)
        ax.set_xlabel("Liquid Limit, LL (%)"); ax.set_ylabel("Plasticity Index, PI (%)")
        ax.set_title("Plasticity Chart | IS 1498")
        _style_axis(ax); ax.legend(loc="best", fontsize=8)
        return _figure_to_png(fig)

    if kind == "compaction":
        r = state["compaction_result"]
        if not r:
            return _empty_plot("No compaction data", "Compaction Curve")
        x = np.array([p.water_content_pct for p in r.points]); y = np.array([p.dry_density_g_cc for p in r.points])
        ax.scatter(x, y, color=DARK, s=45, label="Measured points")
        ax.plot(x, y, color=DARK, linewidth=1.2, alpha=0.8, label="Measured trend")
        ax.scatter([r.observed_omc_pct], [r.observed_mdd_g_cc], color=WSP_RED, s=85, zorder=4, label="Adopted observed OMC/MDD")
        if r.fitted_omc_pct is not None and len(x) >= 3:
            xs = np.linspace(min(x), max(x), 150); coeffs = np.polyfit(x, y, 2)
            ax.plot(xs, np.polyval(coeffs, xs), color=WSP_RED, linestyle="--", linewidth=1.3, label="Indicative fitted curve")
        ax.set_xlabel("Water Content (%)"); ax.set_ylabel("Dry Density (g/cc)")
        ax.set_title("Compaction Curve | Observed OMC/MDD Adopted")
        _style_axis(ax); ax.legend(fontsize=8)
        return _figure_to_png(fig)

    if kind == "cbr":
        if not sample:
            sample = first_sample_for_sheet(state, "CBR", borehole)
        d = filtered_sheet(state, "CBR", borehole, sample)
        pen_col = find_col(d, ["Penetration (mm)", "Penetration"])
        load_col = find_col(d, ["Load (kg)", "Load"])
        title_sample = sample or "Manual"
        if d is not None and not d.empty and pen_col and load_col:
            d = d[[pen_col, load_col]].dropna().copy()
            d[pen_col] = pd.to_numeric(d[pen_col], errors="coerce"); d[load_col] = pd.to_numeric(d[load_col], errors="coerce")
            d = d.dropna().sort_values(pen_col)
            ax.plot(d[pen_col], d[load_col], marker="o", color=DARK, label=str(title_sample))
        elif state["cbr_result"]:
            r = state["cbr_result"]
            ax.scatter([2.5, 5.0], [r.load_2_5_kg, r.load_5_0_kg], color=DARK, label="Key loads")
        else:
            return _empty_plot("No CBR data", "CBR Curve")
        ax.axvline(2.5, color=WSP_RED, linestyle="--", linewidth=1.0, alpha=0.75)
        ax.axvline(5.0, color=WSP_RED, linestyle="--", linewidth=1.0, alpha=0.75)
        ax.set_xlabel("Penetration (mm)"); ax.set_ylabel("Load (kg)")
        ax.set_title(f"CBR Load-Penetration Curve | {title_sample}")
        _style_axis(ax); ax.legend(fontsize=8)
        return _figure_to_png(fig)

    if kind == "spt":
        df = state["data"].get("SPT")
        if df is None or df.empty:
            return _empty_plot("No imported SPT data", "SPT Profile")
        depth_col = find_col(df, ["Depth (m)"]); b2 = find_col(df, ["Blows 150-300"]); b3 = find_col(df, ["Blows 300-450"])
        cf = find_col(df, ["Correction Factor"]); sat = find_col(df, ["Saturated Fine Sand/Silt"]); bh_col = find_col(df, ["Borehole ID"])
        if not (depth_col and b2 and b3):
            return _empty_plot("SPT columns missing", "SPT Profile")
        d = df.copy(); d["Raw N"] = pd.to_numeric(d[b2], errors="coerce") + pd.to_numeric(d[b3], errors="coerce")
        d["Corrected N"] = d["Raw N"] * pd.to_numeric(d[cf], errors="coerce") if cf else d["Raw N"]
        d["Dilatancy-corrected N"] = d["Corrected N"]
        if sat:
            mask = (d[sat].astype(str).str.lower().str.strip() == "yes") & (d["Corrected N"] > 15)
            d.loc[mask, "Dilatancy-corrected N"] = 15 + 0.5 * (d.loc[mask, "Corrected N"] - 15)
        d[depth_col] = pd.to_numeric(d[depth_col], errors="coerce"); d = d.dropna(subset=[depth_col, "Raw N"])
        plot_col = spt_value if spt_value in d.columns else "Raw N"
        if borehole and bh_col:
            d = d[d[bh_col].astype(str) == str(borehole)]
        if d.empty:
            return _empty_plot("No SPT data for selected filter", "SPT Profile")
        if bh_col:
            groups = list(d.groupby(bh_col))
            if not borehole and len(groups) > 6:
                groups = groups[:6]
                ax.text(0.02, 0.02, "Showing first 6 boreholes for readability", transform=ax.transAxes, fontsize=8, color=WSP_RED)
            for bh, grp in groups:
                grp = grp.sort_values(depth_col); ax.plot(grp[plot_col], grp[depth_col], marker="o", linewidth=1.4, label=str(bh))
        else:
            ax.plot(d[plot_col], d[depth_col], marker="o", linewidth=1.4, label=plot_col)
        ax.invert_yaxis(); ax.set_xlabel(plot_col); ax.set_ylabel("Depth (m)"); ax.set_title(f"SPT Profile | {plot_col} | IS 2131")
        _style_axis(ax); ax.legend(fontsize=8)
        return _figure_to_png(fig)

    if kind == "psd":
        df = state["data"].get("PSD")
        if df is None or df.empty:
            return _empty_plot("No PSD data", "Particle Size Distribution")
        if not sample:
            sample = first_sample_for_sheet(state, "PSD", borehole)
        d = filtered_sheet(state, "PSD", borehole, sample)
        sample_col = find_col(d, ["Sample ID", "Sample"]); size_col = find_col(d, ["Sieve Size (mm)", "Particle Size (mm)", "Sieve Size"]); pp_col = find_col(d, ["Percent Passing (%)", "Percent Passing", "% Passing"])
        if not (size_col and pp_col):
            return _empty_plot("PSD columns missing", "Particle Size Distribution")
        d = d.copy(); d[size_col] = pd.to_numeric(d[size_col], errors="coerce"); d[pp_col] = pd.to_numeric(d[pp_col], errors="coerce"); d = d.dropna(subset=[size_col, pp_col])
        if d.empty:
            return _empty_plot("No valid PSD rows", "Particle Size Distribution")
        groups = [(sample or "PSD", d)] if sample or not sample_col else list(d.groupby(sample_col))[:3]
        for sid, grp in groups:
            grp = grp.sort_values(size_col, ascending=False); ax.semilogx(grp[size_col], grp[pp_col], marker="o", linewidth=1.6, label=str(sid))
            _, coeffs = psd_qc_and_coefficients(grp); coeff_text = [f"{key}={coeffs[key]:.3g}" for key in ["D10", "D30", "D60", "Cu", "Cc"] if coeffs.get(key) is not None]
            if coeff_text:
                ax.text(0.02, 0.04, "\n".join(coeff_text), transform=ax.transAxes, fontsize=8, bbox={"facecolor": "white", "alpha": 0.8, "edgecolor": "#D0D5DD"})
        ax.axvline(4.75, color=WSP_RED, linestyle="--", linewidth=1.0, alpha=0.75); ax.axvline(0.075, color=WSP_RED, linestyle="--", linewidth=1.0, alpha=0.75)
        ymax = max(105, float(d[pp_col].max()) + 5); ax.set_ylim(0, ymax); ax.set_xlim(float(d[size_col].max()) * 1.15, max(float(d[size_col].min()) * 0.75, 0.001))
        ax.set_xlabel("Particle Size / Sieve Size (mm)"); ax.set_ylabel("Percent Passing (%)"); ax.set_title(f"Particle Size Distribution | {sample or 'Selected samples'}")
        _style_axis(ax); ax.grid(True, which="both", alpha=0.22); ax.legend(fontsize=8)
        return _figure_to_png(fig)

    if kind == "borehole":
        df = state["data"].get("Boreholes")
        if df is None or df.empty:
            return _empty_plot("No borehole data", "Borehole Log / Soil Profile")
        if borehole:
            bhc = find_col(df, ["Borehole ID"])
            if bhc:
                df = df[df[bhc].astype(str) == str(borehole)]
        bh_col = find_col(df, ["Borehole ID"]); frm = find_col(df, ["From Depth (m)"]); to = find_col(df, ["To Depth (m)"])
        desc = find_col(df, ["Soil Description"]); sym = find_col(df, ["USCS/IS Symbol", "Soil Symbol"])
        if not (bh_col and frm and to):
            return _empty_plot("Borehole columns missing", "Borehole Log / Soil Profile")
        boreholes = list(df[bh_col].dropna().unique())
        if not borehole and len(boreholes) > 6:
            boreholes = boreholes[:6]; ax.text(0.02, 0.02, "Showing first 6 boreholes; select one for detail", transform=ax.transAxes, fontsize=8, color=WSP_RED)
        colors = {"CL": "#FAD7A0", "CI": "#F8C471", "CH": "#E59866", "ML": "#D6EAF8", "MI": "#AED6F1", "MH": "#85C1E9", "SM": "#D5F5E3", "SC": "#ABEBC6", "SP": "#FCF3CF", "SW": "#F9E79F", "GM": "#E8DAEF", "GC": "#D7BDE2", "GP": "#D5DBDB", "GW": "#AEB6BF"}
        hatches = ["/", "\\", "xx", "..", "--", "++", "oo", "**"]; max_depth = 0
        for i, bh in enumerate(boreholes):
            grp = df[df[bh_col] == bh].copy()
            for j, (_, row) in enumerate(grp.iterrows()):
                z1 = safe_float(row[frm], 0); z2 = safe_float(row[to], z1); max_depth = max(max_depth, z2)
                symbol = str(row[sym]).strip() if sym and not pd.isna(row[sym]) else ""; key = symbol[:2].upper(); color = colors.get(key, "#E5E7EB")
                ax.bar(i, z2 - z1, bottom=z1, width=0.58, edgecolor="#111827", color=color, alpha=0.85, hatch=hatches[j % len(hatches)])
                label = symbol if symbol else (str(row[desc])[:14] if desc else "Layer"); ax.text(i, (z1 + z2) / 2, label, ha="center", va="center", fontsize=7)
        ax.set_xticks(range(len(boreholes))); ax.set_xticklabels(boreholes, rotation=0); ax.set_ylim(max_depth + 0.5, 0); ax.set_ylabel("Depth (m)"); ax.set_title("Borehole Log / Soil Profile")
        _style_axis(ax)
        return _figure_to_png(fig)

    return _empty_plot("Unknown figure type", kind)


def _docx_set_cell_shading(cell, fill="F2F2F2"):
    tc_pr = cell._tc.get_or_add_tcPr(); shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), fill); tc_pr.append(shd)


def _docx_set_cell_text(cell, text, bold=False, color=None, size=8):
    cell.text = ""; p = cell.paragraphs[0]; r = p.add_run("" if text is None else str(text)); r.bold = bold; r.font.size = Pt(size)
    if color: r.font.color.rgb = RGBColor(*color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _docx_add_section_heading(doc, number, title):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(14); p.paragraph_format.space_after = Pt(5)
    run = p.add_run(f"{number}    {title.upper()}"); run.bold = True; run.font.size = Pt(15); run.font.color.rgb = RGBColor(225, 6, 0)
    return p


def _docx_add_body_para(doc, text, bold=False):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6); p.paragraph_format.line_spacing = 1.08
    r = p.add_run("" if text is None else str(text)); r.font.size = Pt(10); r.bold = bold
    return p


def _docx_add_dataframe_table(doc, df, max_rows=35, max_cols=8):
    if df is None or df.empty:
        _docx_add_body_para(doc, "No data available."); return
    show = df.head(max_rows).copy(); cols = list(show.columns)[:max_cols]
    t = doc.add_table(rows=1, cols=len(cols)); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, c in enumerate(cols):
        _docx_set_cell_text(t.rows[0].cells[i], c, bold=True, color=(255, 255, 255), size=8); _docx_set_cell_shading(t.rows[0].cells[i], "E10600")
    for _, row in show.iterrows():
        cells = t.add_row().cells
        for i, c in enumerate(cols):
            val = "" if pd.isna(row[c]) else row[c]
            if isinstance(val, float): val = f"{val:.3f}".rstrip("0").rstrip(".")
            _docx_set_cell_text(cells[i], val, size=8)
    if len(df) > max_rows: _docx_add_body_para(doc, f"Showing {max_rows} of {len(df)} records.")


def build_word_report(state: Dict[str, Any]) -> bytes:
    qc_warnings = run_qc_checks(state)
    doc = Document(); sec = doc.sections[0]
    sec.top_margin = Inches(0.65); sec.bottom_margin = Inches(0.65); sec.left_margin = Inches(0.72); sec.right_margin = Inches(0.72)

    header = sec.header; hp = header.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = hp.add_run("GeoReportPro"); r.bold = True; r.font.size = Pt(18); r.font.color.rgb = RGBColor(225, 6, 0)
    footer = sec.footer; fp = footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = fp.add_run("GeoReportPro India | IS-referenced automated geotechnical report package | Engineering review required"); fr.font.size = Pt(8); fr.font.color.rgb = RGBColor(100, 100, 100)

    p = state["project_info"]
    table = doc.add_table(rows=1, cols=2); table.style = "Table Grid"
    table.cell(0, 0).text = p.report_date or date.today().isoformat(); table.cell(0, 1).text = f"Reference No. {p.report_number or 'GIR-XXX-Rev0'}"
    _docx_add_body_para(doc, p.client_name or "Client Name", bold=True); _docx_add_body_para(doc, p.location or "Project Location"); _docx_add_body_para(doc, "")
    _docx_add_body_para(doc, (p.project_name or "Geotechnical Investigation Report").upper(), bold=True); _docx_add_body_para(doc, "Dear Sir/Madam,")

    narrative = "This report package has been generated from the imported borehole and laboratory dataset. "
    bh = state["data"].get("Boreholes")
    if bh is not None:
        bid = find_col(bh, ["Borehole ID"]); to = find_col(bh, ["To Depth (m)"])
        narrative += f"The imported dataset comprises {bh[bid].nunique() if bid else 'multiple'} borehole(s)"
        if to:
            narrative += f" with maximum interpreted depth of approximately {pd.to_numeric(bh[to], errors='coerce').max():.1f} m"
        narrative += ". "
    if state["bearing_result"]:
        narrative += f"For the selected preliminary foundation case, the computed safe bearing capacity is {state['bearing_result'].safe_bearing_capacity_kpa:.1f} kPa; this is preliminary and not a substitute for final IS 6403 design evaluation. "
    narrative += ENGINEERING_DISCLAIMER
    _docx_add_body_para(doc, narrative)

    _docx_add_section_heading(doc, "1.0", "Background")
    _docx_add_body_para(doc, "GeoReportPro is used here to automate repetitive geotechnical reporting workflows including data checking, summary table preparation, borehole visualization, engineering plots and preliminary calculation summaries.")

    _docx_add_section_heading(doc, "2.0", "Reference Standards and Method Limitations")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(["Module", "Tool Logic", "Reference / Limitation"]):
        _docx_set_cell_text(t.rows[0].cells[i], h, bold=True, color=(255, 255, 255), size=8); _docx_set_cell_shading(t.rows[0].cells[i], "E10600")
    for module, logic, ref in REFERENCE_STANDARDS:
        cells = t.add_row().cells; _docx_set_cell_text(cells[0], module, size=8); _docx_set_cell_text(cells[1], logic, size=8); _docx_set_cell_text(cells[2], ref, size=8)

    _docx_add_section_heading(doc, "3.0", "Data Import and Quality Control")
    if qc_warnings:
        _docx_add_body_para(doc, f"QC status: Review required. {len(qc_warnings)} issue(s) flagged.", bold=True)
        for w in qc_warnings[:30]: _docx_add_body_para(doc, f"- {w}")
    else:
        _docx_add_body_para(doc, "QC status: Passed basic completeness, range and monotonicity checks.", bold=True)

    if state["data"].get("Boreholes") is not None:
        _docx_add_section_heading(doc, "4.0", "Borehole Summary"); _docx_add_dataframe_table(doc, state["data"].get("Boreholes"), max_rows=35, max_cols=8)
    if state["data"].get("SPT") is not None:
        _docx_add_section_heading(doc, "5.0", "SPT Summary"); _docx_add_body_para(doc, "SPT N-values are borehole-wise and depth-wise. Raw N is based on the final 300 mm penetration; correction factors are entered by the engineer where available.")
        _docx_add_dataframe_table(doc, state["data"].get("SPT"), max_rows=35, max_cols=8)
    if state["data"].get("Lab_Summary") is not None:
        _docx_add_section_heading(doc, "6.0", "Laboratory Summary"); _docx_add_dataframe_table(doc, state["data"].get("Lab_Summary"), max_rows=35, max_cols=9)

    _docx_add_section_heading(doc, "7.0", "Engineering Calculation Summary")
    if state["classification_result"]:
        rr = state["classification_result"]; _docx_add_body_para(doc, f"Soil classification: {rr.soil_symbol} - {rr.soil_group}. {rr.description}")
    if state["compaction_result"]:
        rr = state["compaction_result"]; _docx_add_body_para(doc, rr.interpretation); rr.qc_note and _docx_add_body_para(doc, rr.qc_note)
    if state["cbr_result"]:
        rr = state["cbr_result"]; _docx_add_body_para(doc, rr.interpretation); _docx_add_body_para(doc, rr.qc_note)
    if state["spt_result"]:
        rr = state["spt_result"]; _docx_add_body_para(doc, f"Selected manual SPT check: raw N = {rr.raw_n}; corrected value = {rr.dilatancy_corrected_n:.2f}; indicative condition = {rr.description}.")
    if state["bearing_result"]:
        rr = state["bearing_result"]; _docx_add_body_para(doc, rr.interpretation); _docx_add_body_para(doc, rr.limitation)

    _docx_add_section_heading(doc, "8.0", "Engineering Figures")
    for name, kind in [("Borehole Log", "borehole"), ("Particle Size Distribution", "psd"), ("Plasticity Chart", "plasticity"), ("Compaction Curve", "compaction"), ("CBR Curve", "cbr"), ("SPT Profile", "spt")]:
        _docx_add_body_para(doc, name, bold=True)
        png = plot_png(state, kind)
        doc.add_picture(io.BytesIO(png), width=Inches(6.6))

    _docx_add_section_heading(doc, "9.0", "Engineering Review Note"); _docx_add_body_para(doc, ENGINEERING_DISCLAIMER)
    out = io.BytesIO(); doc.save(out); return out.getvalue()
