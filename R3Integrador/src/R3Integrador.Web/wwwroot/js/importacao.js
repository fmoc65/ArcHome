document.addEventListener('dragover', e => e.preventDefault());
document.addEventListener('drop', e => {
  const zone = e.target.closest('[data-upload-zone]');
  if (!zone) return;
  e.preventDefault();
  const input = zone.querySelector('input[type=file]');
  if (!input || input.disabled || !e.dataTransfer.files.length) return;
  const transfer = new DataTransfer();
  for (const file of e.dataTransfer.files) transfer.items.add(file);
  input.files = transfer.files;
  input.dispatchEvent(new Event('change', { bubbles: true }));
});
