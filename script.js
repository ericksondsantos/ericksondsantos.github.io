document.getElementById('siteForm').addEventListener('submit', function(e) {
  e.preventDefault();

  const formData = new FormData(this);

  fetch('process.php', {
    method: 'POST',
    body: formData
  })
  .then(response => response.text())
  .then(data => {
    document.getElementById('output').value = data;
  })
  .catch(error => {
    document.getElementById('output').value = 'Error: ' + error;
  });
});
