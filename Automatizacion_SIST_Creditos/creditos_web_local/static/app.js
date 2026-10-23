document.addEventListener('DOMContentLoaded', () => {
  const draftForm = document.querySelector('.draft-form');
  const draftResult = document.getElementById('draft-result');

  const showResult = (ok, text) => {
    if (!draftResult) {
      alert(text);
      return;
    }
    draftResult.style.display = 'block';
    draftResult.className = ok ? 'panel flash flash-success' : 'panel flash flash-error';
    draftResult.textContent = text;
  };

  if (draftForm) {
    draftForm.addEventListener('submit', async (event) => {
      event.preventDefault();
      const formData = new FormData(draftForm);
      const response = await fetch('/draft', {
        method: 'POST',
        body: formData,
      });
      const result = await response.json();
      if (response.ok && result.success) {
        if (result.duplicate) {
          showResult(true, result.message || 'Ya existía un borrador para este documento.');
          return;
        }
        showResult(true, `Borrador creado correctamente. ID: ${result.message_id || 'N/D'}`);
        return;
      }
      showResult(false, result.error || result.message || 'No se pudo crear el borrador.');
    });
  }
});
