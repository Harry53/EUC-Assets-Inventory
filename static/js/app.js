document.addEventListener('DOMContentLoaded', () => {
  const sidebar = document.querySelector('.sidebar');
  document.getElementById('menuToggle')?.addEventListener('click', () => sidebar.classList.toggle('open'));
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
  labels.forEach((label, i) => {
    const x = (i + .5) * (w / labels.length), bh = (values[i] / max) * (h * .62), y = h - bh - 42 * devicePixelRatio;
    const grad = ctx.createLinearGradient(0,y,0,h); grad.addColorStop(0,'#2fc3ff'); grad.addColorStop(1,'#38e6aa');
    ctx.fillStyle = grad; ctx.roundRect(x - barW/2, y, barW, bh, 12 * devicePixelRatio); ctx.fill();
    ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--text'); ctx.fillText(values[i], x, y - 8 * devicePixelRatio);
    ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--muted'); ctx.fillText(label.slice(0,12), x, h - 14 * devicePixelRatio);
  });
}

