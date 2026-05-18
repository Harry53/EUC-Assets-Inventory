document.addEventListener('DOMContentLoaded', () => {
  const sidebar = document.querySelector('.sidebar');
  document.getElementById('menuToggle')?.addEventListener('click', () => { sidebar.classList.toggle('open'); document.body.classList.toggle('sidebar-collapsed'); });
  document.querySelectorAll('[data-toggle]').forEach(btn => btn.addEventListener('click', () => document.getElementById(btn.dataset.toggle)?.classList.toggle('hidden')));
  document.getElementById('themeToggle')?.addEventListener('click', () => {
    const html = document.documentElement;
    html.dataset.theme = html.dataset.theme === 'light' ? 'dark' : 'light';
  });
  document.querySelectorAll('.gen-pass').forEach(btn => btn.addEventListener('click', () => {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%';
    btn.closest('td,form').querySelector('.password-field').value = Array.from({length: 14}, () => chars[Math.floor(Math.random() * chars.length)]).join('');
  }));
  const modal = document.getElementById('historyModal');
  const historyText = document.getElementById('historyText');
  document.querySelectorAll('.history-btn').forEach(btn => btn.addEventListener('click', () => {
    historyText.textContent = btn.dataset.history || 'No details available.';
    modal.classList.add('show');
  }));
  document.querySelector('.modal-close')?.addEventListener('click', () => modal.classList.remove('show'));
  modal?.addEventListener('click', e => { if (e.target === modal) modal.classList.remove('show'); });
  document.querySelectorAll('.live-search').forEach(input => input.addEventListener('input', () => {
    if (input.value === '' && new URLSearchParams(location.search).has(input.name)) location.href = location.pathname;
    const table = document.querySelector('.filterable tbody');
    if (!table) return;
    const term = input.value.toLowerCase();
    table.querySelectorAll('tr').forEach(row => row.style.display = row.textContent.toLowerCase().includes(term) ? '' : 'none');
  }));
  document.querySelectorAll('.table-filter').forEach(input => input.addEventListener('input', () => {
    const term = input.value.toLowerCase();
    input.parentElement.querySelectorAll('tbody tr').forEach(row => row.style.display = row.textContent.toLowerCase().includes(term) ? '' : 'none');
  }));
  document.querySelectorAll('.nav a').forEach(a => { if (a.pathname === location.pathname || (location.pathname.startsWith(a.pathname) && a.pathname !== '/')) a.classList.add('active'); });
  document.querySelectorAll('.chart').forEach(canvas => drawChart(canvas));
  document.querySelectorAll('[data-checklist-tab]').forEach(btn => btn.addEventListener('click', e => {
    e.preventDefault();
    const type = btn.dataset.checklistTab;
    document.getElementById('checklistType').value = type;
    document.querySelectorAll('.joining,.replacement,.rebuild,.exit').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('.' + type.toLowerCase()).forEach(el => el.classList.remove('hidden'));
  }));
});
function drawChart(canvas) {
  const data = JSON.parse(canvas.dataset.chart || '{}');
  const labels = Object.keys(data), values = Object.values(data).map(Number);
  const ctx = canvas.getContext('2d'), w = canvas.width = canvas.offsetWidth * devicePixelRatio, h = canvas.height = canvas.offsetHeight * devicePixelRatio;
  const max = Math.max(...values, 1), barW = w / Math.max(labels.length, 1) * .58;
  ctx.clearRect(0,0,w,h); ctx.font = `${12*devicePixelRatio}px Segoe UI`; ctx.textAlign = 'center';
  if (canvas.classList.contains('pie')) {
    let total = values.reduce((a,b)=>a+b,0) || 1, start = -Math.PI/2;
    labels.forEach((label,i)=>{ const slice = values[i]/total*Math.PI*2; ctx.beginPath(); ctx.moveTo(w/2,h/2); ctx.fillStyle=['#2fc3ff','#38e6aa','#ffd166','#ff5d7d','#8a7dff','#00c2a8'][i%6]; ctx.arc(w/2,h/2,Math.min(w,h)*.32,start,start+slice); ctx.fill(); start+=slice; ctx.fillText(label.slice(0,12), 70*devicePixelRatio, (24+i*18)*devicePixelRatio); }); return;
  }
  labels.forEach((label, i) => {
    const x = (i + .5) * (w / labels.length), bh = (values[i] / max) * (h * .62), y = h - bh - 42 * devicePixelRatio;
    const grad = ctx.createLinearGradient(0,y,0,h); grad.addColorStop(0,'#2fc3ff'); grad.addColorStop(1,'#38e6aa');
    ctx.fillStyle = grad; ctx.roundRect(x - barW/2, y, barW, bh, 12 * devicePixelRatio); ctx.fill();
    ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--text'); ctx.fillText(values[i], x, y - 8 * devicePixelRatio);
    ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--muted'); ctx.fillText(label.slice(0,12), x, h - 14 * devicePixelRatio);
  });
}

function parseJsonScript(id) {
  try { return JSON.parse(document.getElementById(id)?.textContent || '[]'); } catch { return []; }
}
function setNamed(form, name, value) {
  form.querySelectorAll(`[name="${CSS.escape(name)}"]`).forEach(el => {
    if (el.disabled && !name.endsWith('_display')) return;
    el.value = value || '';
  });
}
function buildHistoryTable(tag) {
  const rows = parseJsonScript('assetHistoryData').filter(r => r.device_tag === tag);
  const body = document.querySelector('#assetHistoryTable tbody');
  if (!body) return;
  body.innerHTML = rows.length ? rows.map(r => `<tr><td>${r.device_tag || ''}</td><td>${r.employee_name || r.username || ''}</td><td>${r.e_code || ''}</td><td>${r.action || ''}</td><td>${r.status || r.approval_status || ''}</td><td>${r.updated_at || r.created_at || ''}</td></tr>`).join('') : '<tr><td colspan="6">Empty</td></tr>';
}
function buildAllocatedSoftware(tag) {
  const rows = parseJsonScript('assignmentData').filter(r => r.desktop_laptop_tag === tag);
  const box = document.getElementById('allocatedSoftware');
  if (!box) return;
  box.innerHTML = rows.length ? `<div class="table-wrap small"><table><thead><tr><th>Software</th><th>Version</th><th>License</th><th>Status</th></tr></thead><tbody>${rows.map(r => `<tr><td>${r.software_name || ''}</td><td>${r.version || ''}</td><td>${r.license_key || ''}</td><td>${r.status || ''}</td></tr>`).join('')}</tbody></table></div>` : 'No software allocated.';
}
document.addEventListener('DOMContentLoaded', () => {
  const assets = parseJsonScript('assetData'), employees = parseJsonScript('employeeData'), vendors = parseJsonScript('vendorData');
  document.querySelectorAll('[data-autofill-asset]').forEach(input => input.addEventListener('change', () => {
    const form = input.closest('form');
    const asset = assets.find(a => a.device_tag === input.value || a.serial_number === input.value || a.system_name === input.value);
    if (!asset || !form) return;
    Object.entries(asset).forEach(([k, v]) => setNamed(form, k, v));
    setNamed(form, 'desktop_laptop_tag', asset.device_tag || '');
    setNamed(form, 'desktop_laptop_tag_display', asset.device_tag || '');
    const emp = employees.find(e => e.e_code === asset.e_code || e.employee_name === asset.username);
    if (emp) Object.entries(emp).forEach(([k, v]) => setNamed(form, k, v));
    buildHistoryTable(asset.device_tag);
    buildAllocatedSoftware(asset.device_tag);
  }));
  document.querySelectorAll('[data-autofill-employee]').forEach(input => input.addEventListener('change', () => {
    const form = input.closest('form');
    const emp = employees.find(e => e.e_code === input.value || e.employee_name === input.value);
    if (!emp || !form) return;
    Object.entries(emp).forEach(([k, v]) => setNamed(form, k, v));
    setNamed(form, 'username', emp.employee_name || '');
  }));
  document.querySelectorAll('[data-autofill-vendor]').forEach(input => input.addEventListener('change', () => {
    const form = input.closest('form');
    const vendor = vendors.find(v => v.vendor_name === input.value || v.contact_person === input.value);
    if (!vendor || !form) return;
    setNamed(form, 'supplier_name', vendor.vendor_name || '');
  }));
  document.querySelectorAll('fieldset legend input[type="checkbox"]').forEach(cb => cb.addEventListener('change', () => {
    cb.closest('fieldset').querySelectorAll('input, select, textarea').forEach(el => { if (el !== cb) el.disabled = !cb.checked; });
  }));
});
document.addEventListener('DOMContentLoaded', () => {
  const assets = parseJsonScript('assetData').concat(Array.from(document.querySelectorAll('#assetList option')).map(o => ({device_tag: o.value})));
  document.querySelectorAll('[data-checklist-asset]').forEach(input => input.addEventListener('change', () => {
    const allAssets = parseJsonScript('assetData');
    const asset = allAssets.find(a => a.device_tag === input.value || a.serial_number === input.value || a.system_name === input.value);
    const form = input.closest('form');
    if (!asset || !form) return;
    setNamed(form, 'machine_asset_tag', asset.device_tag || '');
    setNamed(form, 'system_serial_number', asset.serial_number || '');
    setNamed(form, 'model', asset.device_model || '');
    setNamed(form, 'ip_address', asset.ip_address || '');
    setNamed(form, 'host_name', asset.system_name || '');
    setNamed(form, 'employee_name', asset.username || '');
    setNamed(form, 'e_code', asset.e_code || '');
  }));
});
