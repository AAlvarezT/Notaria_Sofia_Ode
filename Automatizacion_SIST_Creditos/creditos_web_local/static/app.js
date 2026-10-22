document.addEventListener('DOMContentLoaded', () => {
  const draftForm = document.querySelector('.draft-form');
  if (draftForm) {
    draftForm.addEventListener('submit', async (event) => {
      event.preventDefault();
      const formData = new FormData(draftForm);
      const response = await fetch('/draft', {
        method: 'POST',
        body: formData,
      });
      const result = await response.json();
      alert(result.success ? 'Borrador creado correctamente.' : result.error || 'No se pudo crear el borrador.');
    });
  }
});
