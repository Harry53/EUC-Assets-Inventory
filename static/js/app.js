document.addEventListener('DOMContentLoaded', () => {
  const sidebar = document.querySelector('.sidebar');
  document.getElementById('menuToggle')?.addEventListener('click', () => sidebar.classList.toggle('open'));
  document.querySelectorAll('[data-toggle]').forEach(btn => btn.addEventListener('click', () => document.getElementById(btn.dataset.toggle)?.classList.toggle('hidden')));
  document.querySelectorAll('.gen-pass').forEach(btn => btn.addEventListener('click', () => {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%';
    const pass = Array.from({length: 14}, () => chars[Math.floor(Math.random() * chars.length)]).join('');
    btn.closest('td').querySelector('.password-field').value = pass;
  }));
  const modal = document.getElementById('historyModal');
  const historyText = document.getElementById('historyText');
  document.querySelectorAll('.history-btn').forEach(btn => btn.addEventListener('click', () => {
    historyText.textContent = btn.dataset.history || 'No previous user history available.';
    modal.classList.add('show');
  }));
  document.querySelector('.modal-close')?.addEventListener('click', () => modal.classList.remove('show'));
  modal?.addEventListener('click', e => { if (e.target === modal) modal.classList.remove('show'); });
  document.querySelectorAll('.table-filter').forEach(input => input.addEventListener('input', () => {
    const term = input.value.toLowerCase();
    input.parentElement.querySelectorAll('tbody tr').forEach(row => row.style.display = row.textContent.toLowerCase().includes(term) ? '' : 'none');
  }));
});

