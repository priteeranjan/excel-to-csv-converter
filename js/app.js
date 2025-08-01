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
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert sheet to JSON for processing
    const rawData = XLSX.utils.sheet_to_json(sheet, { raw: false });

    // Format Date column to ISO (YYYY-MM-DD)
    const processed = rawData.map(row => {
      if (row.Date) {
        const d = new Date(row.Date);
        if (!isNaN(d)) {
          row.Date = d.toISOString().split('T')[0]; // "YYYY-MM-DD"
        }
      }
      return row;
    });

    // Create a sheet and force-quote values for Excel-safe output
    const newSheet = XLSX.utils.json_to_sheet(processed);
    csvOutput = XLSX.utils.sheet_to_csv(newSheet, {
      FS: ",",
      forceQuotes: true
    });

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

  const originalName = selectedFile.name;
  const baseName = originalName.substring(0, originalName.lastIndexOf('.')) || originalName;

  const blob = new Blob([csvOutput], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `${baseName}.csv`;
  link.click();
}
