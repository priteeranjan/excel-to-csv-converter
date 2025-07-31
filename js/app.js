let csvOutput = '';
let selectedFile = null;

document.getElementById('upload').addEventListener('change', function (e) {
  selectedFile = e.target.files[0];
  document.getElementById('status').textContent = selectedFile
    ? `📁 Selected: ${selectedFile.name}`
    : '';
});

function convertExcel() {
  if (!selectedFile) {
    document.getElementById('status').textContent = "❗ Please select an Excel file first.";
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    csvOutput = XLSX.utils.sheet_to_csv(sheet);

    document.getElementById('status').textContent = "✅ File converted! You can now download the CSV.";
    document.getElementById('downloadBtn').disabled = false;
  };
  reader.readAsArrayBuffer(selectedFile);
}

function downloadCSV() {
  if (!csvOutput || !selectedFile) {
    document.getElementById('status').textContent = "❗ Please convert a file first.";
    return;
  }

  // Use original file name without extension
  const originalName = selectedFile.name;
  const baseName = originalName.substring(0, originalName.lastIndexOf('.')) || originalName;

  const blob = new Blob([csvOutput], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `${baseName}.csv`;
  link.click();
}
