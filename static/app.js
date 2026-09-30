let APP_STATE = null;
let PREVIEW = {};

const $ = (id) => document.getElementById(id);
const num = (id) => {
  const v = parseFloat($(id).value);
  return Number.isFinite(v) ? v : null;
};
const intVal = (id) => {
  const v = parseInt($(id).value, 10);
  return Number.isFinite(v) ? v : null;
};

function setLoading(show, text='Working...') {
  $('loader-text').textContent = text;
  $('loader').classList.toggle('show', show);
}

function toast(message) {
  const el = $('toast');
  el.textContent = message;
  el.classList.add('show');
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(() => el.classList.remove('show'), 2800);
}

async function api(url, options={}) {
  const res = await fetch(url, options);
  if (!res.ok) {
    let msg = `Request failed (${res.status})`;
    try { const j = await res.json(); msg = j.detail || msg; } catch (_) {}
    throw new Error(msg);
  }
  return res.json();
}

function showPage(name) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  const page = $(`page-${name}`);
  if (page) page.classList.add('active');
  const btn = document.querySelector(`.nav-btn[data-page="${name}"]`);
  if (btn) btn.classList.add('active');
  window.scrollTo({top:0, behavior:'smooth'});
  if (name === 'figures') refreshAllPlots();
  if (name === 'summary') refreshState();
}

for (const btn of document.querySelectorAll('.nav-btn')) {
  btn.addEventListener('click', () => showPage(btn.dataset.page));
}

function setInput(id, value) {
  if ($(id) && value !== undefined && value !== null) $(id).value = value;
}

function fillForms(state) {
  const p = state.project || {};
  ['project_name','client_name','location','report_number','prepared_by','checked_by','approved_by','report_date','remarks'].forEach(k => setInput(k, p[k] || ''));

  const c = state.inputs?.classification || {};
  setInput('ll', c.ll); setInput('pl', c.pl); setInput('water_content', c.water_content);
  setInput('gravel', c.gravel); setInput('sand', c.sand); setInput('fines', c.fines);

  const cb = state.inputs?.cbr || {};
  setInput('load_2_5', cb.load_2_5); setInput('load_5_0', cb.load_5_0);
  if ($('cbr_confirm')) $('cbr_confirm').checked = !!cb.adopt_5_if_higher_confirmed;

  const s = state.inputs?.spt || {};
  setInput('spt_b1', s.b1); setInput('spt_b2', s.b2); setInput('spt_b3', s.b3); setInput('spt_cf', s.correction_factor);
  if ($('spt_sat')) $('spt_sat').value = String(!!s.saturated_fine_sand);

  const b = state.inputs?.bearing || {};
  setInput('foundation_type', b.foundation_type); setInput('width', b.width); setInput('depth', b.depth);
  setInput('cohesion', b.cohesion); setInput('unit_weight', b.unit_weight); setInput('phi', b.phi); setInput('fos', b.fos);

  fillCompactionRows(state.inputs?.compaction?.points || []);
}

function updateDashboard(state) {
  const m = state.metrics || {};
  $('m-project').textContent = m.project ?? 'Not set';
  $('m-project-sub').textContent = m.project_subtitle ?? '';
  $('m-boreholes').textContent = m.boreholes ?? '--';
  $('m-soil').textContent = m.soil_class ?? '--';
  $('m-soil-sub').textContent = m.soil_subtitle ?? '';
  $('m-cbr').textContent = m.design_cbr ?? '--';
  $('m-sbc').textContent = m.safe_bearing ?? '--';
  $('m-qc').textContent = m.qc_status ?? 'Pending';
  $('m-qc-sub').textContent = m.qc_subtitle ?? '';
  $('excel-name').textContent = state.excel_name || 'No workbook loaded';
}

function resultText(state) {
  const a = state.results?.atterberg;
  const c = state.results?.classification;
  if (a && c) {
    let text = `STANDARD REFERENCE\nIS 2720 Part 5: Liquid limit and plastic limit test data.\nIS 1498: Soil classification based on grain-size fractions and plasticity chart.\n\n`;
    text += `ATTERBERG LIMITS\nLL = ${a.liquid_limit.toFixed(2)}%\nPL = ${a.plastic_limit.toFixed(2)}%\nPI = ${a.plasticity_index.toFixed(2)}%\nCategory = ${a.plasticity_category}`;
    if (a.liquidity_index !== null && a.liquidity_index !== undefined) text += `\nLiquidity Index = ${a.liquidity_index.toFixed(3)}\nConsistency Index = ${a.consistency_index.toFixed(3)}`;
    text += `\n\nSOIL CLASSIFICATION\nSymbol = ${c.soil_symbol}\nGroup = ${c.soil_group}\nDescription = ${c.description}\nEngineering note = ${c.notes}`;
    $('classification-result').textContent = text;
  }

  const comp = state.results?.compaction;
  if (comp) {
    let text = `STANDARD REFERENCE\nIS 2720 Part 7 / Part 8 depending on whether light or heavy compaction was performed.\n\nADOPTED REPORTING POSITION\nObserved OMC/MDD from measured laboratory points is treated as the primary reported value.\n\nRESULT\n${comp.interpretation}`;
    if (comp.qc_note) text += `\n\nQC / ENGINEERING NOTE\n${comp.qc_note}`;
    text += '\n\nPOINTS';
    for (const p of comp.points || []) text += `\nw = ${p.water_content_pct.toFixed(2)}%, Wet density = ${p.wet_density_g_cc.toFixed(3)} g/cc, Dry density = ${p.dry_density_g_cc.toFixed(3)} g/cc`;
    $('compaction-result').textContent = text;
  }

  const cbr = state.results?.cbr;
  if (cbr) {
    $('cbr-result').textContent = `STANDARD REFERENCE\nIS 2720 Part 16: CBR calculation using standard loads at 2.5 mm and 5.0 mm.\n\nLoad at 2.5 mm = ${cbr.load_2_5_kg.toFixed(2)} kg\nLoad at 5.0 mm = ${cbr.load_5_0_kg.toFixed(2)} kg\nCBR at 2.5 mm = ${cbr.cbr_2_5_mm.toFixed(2)}%\nCBR at 5.0 mm = ${cbr.cbr_5_0_mm.toFixed(2)}%\nAdopted Design CBR = ${cbr.design_cbr.toFixed(2)}% at ${cbr.selected_penetration_mm.toFixed(1)} mm\n\nAdoption note: ${cbr.qc_note}`;
  }

  const spt = state.results?.spt;
  if (spt) {
    let text = `STANDARD REFERENCE\nIS 2131: N-value is taken as the number of blows for the final 300 mm penetration.\n\nSINGLE TEST RESULT\nRaw N-value = ${spt.raw_n}\nEngineer-corrected N-value = ${spt.overburden_corrected_n.toFixed(2)}\nDilatancy-corrected N-value = ${spt.dilatancy_corrected_n.toFixed(2)}\nIndicative density interpretation = ${spt.description}\n\nIMPORTANT\nThis single-test calculator is not a borehole profile. Imported SPT profiles are plotted separately using the Excel SPT sheet.`;
    if (spt.warning) text += `\n\nWarning\n${spt.warning}`;
    $('spt-result').textContent = text;
  }

  const br = state.results?.bearing;
  if (br) {
    $('bearing-result').textContent = `METHOD STATUS\nThis is a preliminary Terzaghi-style bearing check. It must not be represented as a complete IS 6403 final design output.\n\nNc = ${br.nc.toFixed(3)}\nNq = ${br.nq.toFixed(3)}\nNγ = ${br.ngamma.toFixed(3)}\n\nUltimate Bearing Capacity = ${br.ultimate_bearing_capacity_kpa.toFixed(2)} kPa\nSafe Bearing Capacity = ${br.safe_bearing_capacity_kpa.toFixed(2)} kPa\n\n${br.interpretation}\n\nLimitation: ${br.limitation}`;
  }
}

function updateSummary(state) {
  const a = state.results?.atterberg, c = state.results?.classification, comp = state.results?.compaction;
  const cbr = state.results?.cbr, spt = state.results?.spt, br = state.results?.bearing;
  $('s-pi').textContent = a ? `${a.plasticity_index.toFixed(2)}%` : '--'; $('s-pi-sub').textContent = a?.plasticity_category || 'Pending';
  $('s-soil').textContent = c?.soil_symbol || '--'; $('s-soil-sub').textContent = c?.soil_group || 'Pending';
  $('s-mdd').textContent = comp ? `${comp.observed_mdd_g_cc.toFixed(3)} g/cc` : '--'; $('s-mdd-sub').textContent = comp ? `OMC = ${comp.observed_omc_pct.toFixed(2)}%` : 'Pending';
  $('s-cbr').textContent = cbr ? `${cbr.design_cbr.toFixed(2)}%` : '--'; $('s-cbr-sub').textContent = cbr ? `Selected at ${cbr.selected_penetration_mm.toFixed(1)} mm` : 'Pending';
  $('s-spt').textContent = spt ? spt.dilatancy_corrected_n.toFixed(1) : '--'; $('s-spt-sub').textContent = spt?.description || 'Pending';
  $('s-sbc').textContent = br ? `${br.safe_bearing_capacity_kpa.toFixed(1)} kPa` : '--';
  $('summary-text').textContent = state.summary_text || '';
  $('standards-text').textContent = state.standards_text || '';
  const warnings = state.qc_warnings || [];
  $('qc-text').textContent = warnings.length ? `QC CHECKS COMPLETED\n=============================================\nStatus: Review Required\n\n${warnings.map(w => '- ' + w).join('\n')}` : (state.sheets?.length ? 'QC CHECKS COMPLETED\n=============================================\nStatus: Passed basic completeness, range and monotonicity checks.' : 'Run QC after importing a workbook.');
}

function updateFilters(state) {
  const b = $('figure_borehole'), s = $('figure_sample');
  const currentB = b.value, currentS = s.value;
  b.innerHTML = '<option value="">All Boreholes</option>' + (state.boreholes || []).map(x => `<option>${escapeHtml(x)}</option>`).join('');
  s.innerHTML = '<option value="">Auto / First Sample</option>' + (state.samples || []).map(x => `<option>${escapeHtml(x)}</option>`).join('');
  if ([...b.options].some(o => o.value === currentB)) b.value = currentB;
  if ([...s.options].some(o => o.value === currentS)) s.value = currentS;
}

function applyState(state, refill=true) {
  APP_STATE = state;
  updateDashboard(state); resultText(state); updateSummary(state); updateFilters(state);
  if (refill) fillForms(state);
}

async function refreshState() {
  try { applyState(await api('/api/state'), true); } catch (e) { toast(e.message); }
}

async function saveProject() {
  const payload = {};
  for (const k of ['project_name','client_name','location','report_number','prepared_by','checked_by','approved_by','report_date','remarks']) payload[k] = $(k).value.trim();
  setLoading(true, 'Saving project details...');
  try { applyState(await api('/api/project', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)}), false); toast('Project details saved.'); }
  catch (e) { toast(e.message); }
  finally { setLoading(false); }
}

async function uploadWorkbook() {
  const file = $('excel-file').files[0];
  if (!file) return toast('Select an Excel workbook first.');
  const fd = new FormData(); fd.append('file', file);
  setLoading(true, 'Uploading and processing workbook...');
  try {
    const result = await api('/api/upload', {method:'POST', body:fd});
    PREVIEW = result.preview || {};
    renderPreview();
    applyState(result.state, true);
    refreshAllPlots();
    toast('Workbook imported successfully.');
  } catch (e) { toast(e.message); }
  finally { setLoading(false); }
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
}

function renderPreview() {
  const names = Object.keys(PREVIEW);
  const tabs = $('preview-tabs');
  if (!names.length) { tabs.innerHTML = ''; $('preview-container').innerHTML = '<div style="padding:24px" class="muted">No sheets to preview.</div>'; return; }
  tabs.innerHTML = names.map((n,i) => `<button class="tab-btn ${i===0?'active':''}" onclick="showPreviewSheet('${escapeHtml(n)}', this)">${escapeHtml(n)}</button>`).join('');
  showPreviewSheet(names[0], tabs.querySelector('.tab-btn'));
}

function showPreviewSheet(name, button) {
  document.querySelectorAll('#preview-tabs .tab-btn').forEach(b => b.classList.remove('active'));
  if (button) button.classList.add('active');
  const p = PREVIEW[name];
  if (!p || !p.columns?.length) { $('preview-container').innerHTML = '<div style="padding:24px" class="muted">Sheet is empty.</div>'; return; }
  let html = '<table class="data-table"><thead><tr>' + p.columns.map(c => `<th>${escapeHtml(c)}</th>`).join('') + '</tr></thead><tbody>';
  html += p.rows.map(row => '<tr>' + p.columns.map(c => `<td>${escapeHtml(row[c])}</td>`).join('') + '</tr>').join('') + '</tbody></table>';
  $('preview-container').innerHTML = html;
}

async function runClassification() {
  const payload = {ll:num('ll'), pl:num('pl'), water_content:num('water_content'), gravel:num('gravel'), sand:num('sand'), fines:num('fines')};
  setLoading(true, 'Calculating soil classification...');
  try { applyState(await api('/api/classification', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}), false); refreshPlot('plasticity'); toast('Classification updated.'); }
  catch (e) { toast(e.message); }
  finally { setLoading(false); }
}

function fillCompactionRows(points) {
  const body = $('compaction-rows'); body.innerHTML = '';
  const src = points.length ? points : [{water_content:'',wet_density:''}];
  for (const p of src) addCompactionRow(p.water_content, p.wet_density);
}

function addCompactionRow(w='', wet='') {
  const tr = document.createElement('tr');
  tr.innerHTML = `<td><input class="comp-w" type="number" step="any" value="${escapeHtml(w)}"></td><td><input class="comp-wet" type="number" step="any" value="${escapeHtml(wet)}"></td><td><button class="btn small danger" onclick="this.closest('tr').remove()">Remove</button></td>`;
  $('compaction-rows').appendChild(tr);
}

async function runCompaction() {
  const rows = [...document.querySelectorAll('#compaction-rows tr')];
  const points = rows.map(r => ({water_content:parseFloat(r.querySelector('.comp-w').value), wet_density:parseFloat(r.querySelector('.comp-wet').value)})).filter(p => Number.isFinite(p.water_content) && Number.isFinite(p.wet_density));
  setLoading(true, 'Calculating OMC and MDD...');
  try { applyState(await api('/api/compaction', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({points})}), false); refreshPlot('compaction'); toast('Compaction analysis updated.'); }
  catch (e) { toast(e.message); }
  finally { setLoading(false); }
}

async function runCBR() {
  const payload = {load_2_5:num('load_2_5'), load_5_0:num('load_5_0'), adopt_5_if_higher_confirmed:$('cbr_confirm').checked};
  setLoading(true, 'Calculating CBR...');
  try { applyState(await api('/api/cbr', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}), false); refreshPlot('cbr'); toast('CBR updated.'); }
  catch (e) { toast(e.message); }
  finally { setLoading(false); }
}

async function runSPT() {
  const payload = {b1:intVal('spt_b1'), b2:intVal('spt_b2'), b3:intVal('spt_b3'), correction_factor:num('spt_cf'), saturated_fine_sand:$('spt_sat').value==='true'};
  setLoading(true, 'Calculating SPT...');
  try { applyState(await api('/api/spt', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}), false); refreshPlot('spt'); toast('SPT result updated.'); }
  catch (e) { toast(e.message); }
  finally { setLoading(false); }
}

async function runBearing() {
  const payload = {foundation_type:$('foundation_type').value, width:num('width'), depth:num('depth'), cohesion:num('cohesion'), unit_weight:num('unit_weight'), phi:num('phi'), fos:num('fos')};
  setLoading(true, 'Calculating bearing capacity...');
  try { applyState(await api('/api/bearing', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}), false); toast('Bearing calculation updated.'); }
  catch (e) { toast(e.message); }
  finally { setLoading(false); }
}

function plotUrl(kind) {
  const params = new URLSearchParams();
  const bh = $('figure_borehole')?.value || ''; const sample = $('figure_sample')?.value || ''; const sptVal = $('figure_spt_value')?.value || 'Raw N';
  if (bh) params.set('borehole', bh); if (sample) params.set('sample', sample); if (kind === 'spt') params.set('spt_value', sptVal);
  params.set('_', Date.now());
  return `/api/plot/${kind}?${params.toString()}`;
}

function refreshPlot(kind) {
  const map = {plasticity:'plot-plasticity', compaction:'plot-compaction', cbr:'plot-cbr', spt:'plot-spt'};
  const id = map[kind]; if (id) $(id).src = plotUrl(kind);
  const figMap = {plasticity:'plot-fig-plasticity', compaction:'plot-fig-compaction', cbr:'plot-fig-cbr', spt:'plot-fig-spt'};
  if (figMap[kind]) $(figMap[kind]).src = plotUrl(kind);
}

function refreshAllPlots() {
  $('plot-borehole').src = plotUrl('borehole');
  $('plot-psd').src = plotUrl('psd');
  $('plot-fig-plasticity').src = plotUrl('plasticity');
  $('plot-fig-compaction').src = plotUrl('compaction');
  $('plot-fig-cbr').src = plotUrl('cbr');
  $('plot-fig-spt').src = plotUrl('spt');
}

async function runQC() {
  setLoading(true, 'Running QC checks...');
  try { applyState(await api('/api/qc', {method:'POST'}), false); toast('QC checks completed.'); }
  catch (e) { toast(e.message); }
  finally { setLoading(false); }
}

function downloadReport() {
  window.location.href = '/api/report';
}

$('figure_borehole').addEventListener('change', refreshAllPlots);
$('figure_sample').addEventListener('change', refreshAllPlots);
$('figure_spt_value').addEventListener('change', refreshAllPlots);

refreshState().then(() => {
  refreshPlot('plasticity'); refreshPlot('compaction'); refreshPlot('cbr'); refreshPlot('spt');
});
