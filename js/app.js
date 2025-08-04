let csvOutput = '';
let selectedFile = null;
let workbook = null;
let visibleColumns = [];

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('upload');
const sheetSelector = document.getElementById('sheetSelector');
const previewTable = document.getElementById('previewTable');
const previewContainer = document.getElementById('previewContainer');
const searchInput = document.getElementById('searchInput');
const tableStats = document.getElementById('tableStats');
const matchCount = document.getElementById('matchCount');
const totalCount = document.getElementById('totalCount');
const searchWrapper = document.getElementById('searchWrapper');
const columnToggles = document.getElementById('columnToggles');
const alertBox = document.getElementById('alertBox');
const downloadBtn = document.getElementById('downloadBtn');

fileInput.addEventListener('change', e => {
  selectedFile = e.target.files[0];
  clearAlert();
  if (selectedFile) loadWorkbook(selectedFile);
});

dropzone.addEventListener('click', () => fileInput.click());

['dragenter', 'dragover'].forEach(evt =>
  dropzone.addEventListener(evt, e => {
    e.preventDefault(); e.stopPropagation();
    dropzone.classList.add('dragover');
  })
);

['dragleave', 'drop'].forEach(evt =>
  dropzone.addEventListener(evt, e => {
    e.preventDefault(); e.stopPropagation();
    dropzone.classList.remove('dragover');
  })
);

dropzone.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if (file && /\.(xlsx|xls|xlsm)$/i.test(file.name)) {
    selectedFile = file;
    fileInput.files = e.dataTransfer.files;
    clearAlert();
    loadWorkbook(file);
  } else {
    showAlert("❗ Invalid file type. Please upload an Excel file.");
  }
});

function loadWorkbook(file) {
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const sheetNames = workbook.SheetNames;

      if (sheetNames.length === 0) {
        showAlert("❗ No sheets found in the uploaded file.");
        return;
      }

      sheetSelector.innerHTML = '';
      sheetNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        sheetSelector.appendChild(option);
      });

      document.getElementById('sheetSelectorWrapper').style.display = 'block';
    } catch {
      showAlert("❗ Error reading Excel file.");
    }
  };
  reader.readAsArrayBuffer(file);
}

function convertExcel() {
  if (!selectedFile || !workbook) {
    showAlert("❗ Please select a valid Excel file.");
    return;
  }

  const selectedSheetName = sheetSelector.value;
  const sheet = workbook.Sheets[selectedSheetName];
  const rawData = XLSX.utils.sheet_to_json(sheet, { raw: false });

  if (!rawData.length) {
    showAlert("❗ Selected sheet is empty.");
    return;
  }

  const headers = Object.keys(rawData[0]);
  visibleColumns = [...headers];

  const processed = rawData.map(row => {
    if (row.Date) {
      const d = new Date(row.Date);
      if (!isNaN(d)) row.Date = d.toISOString().split('T')[0];
    }
    return row;
  });

  renderTable(processed, headers);
  renderColumnToggles(headers);

  csvOutput = generateCSV(processed, visibleColumns);
  downloadBtn.disabled = false;
  downloadBtn.innerHTML = '<i class="bi bi-download"></i> Download CSV';
  document.getElementById('status').textContent = `✅ "${selectedSheetName}" converted! You can now download the CSV.`;
}

function renderTable(data, headers) {
  previewTable.innerHTML = '';
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');

  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    th.dataset.col = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  previewTable.appendChild(thead);

  const tbody = document.createElement('tbody');
  data.forEach(row => {
    const tr = document.createElement('tr');
    headers.forEach(h => {
      const td = document.createElement('td');
      td.textContent = row[h] ?? '';
      td.dataset.col = h;
      td.setAttribute('contenteditable', 'true');

      td.addEventListener('keydown', e => {
        if (e.key === 'Enter') {
          e.preventDefault();
          td.blur();
        }
      });

      td.addEventListener('blur', () => {
        updateCSVOutput();
        downloadBtn.innerHTML = '<i class="bi bi-download"></i> Download Edited CSV <span class="text-success">✅</span>';
      });

      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  previewTable.appendChild(tbody);

  previewContainer.style.display = 'block';
  searchWrapper.style.display = 'block';
  tableStats.style.display = 'block';
  columnToggles.style.display = 'block';
  totalCount.textContent = data.length;
  matchCount.textContent = data.length;
}

function renderColumnToggles(headers) {
  columnToggles.innerHTML = 'Show columns: ';
  headers.forEach(h => {
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = true;
    checkbox.classList.add('form-check-input', 'me-1');
    checkbox.dataset.col = h;
    checkbox.addEventListener('change', () => toggleColumn(h, checkbox.checked));
    const label = document.createElement('label');
    label.classList.add('me-3');
    label.appendChild(checkbox);
    label.append(h);
    columnToggles.appendChild(label);
  });
}

function toggleColumn(colName, visible) {
  const cells = previewTable.querySelectorAll(`[data-col="${colName}"]`);
  cells.forEach(cell => {
    cell.style.display = visible ? '' : 'none';
  });

  if (visible && !visibleColumns.includes(colName)) {
    visibleColumns.push(colName);
  } else if (!visible) {
    visibleColumns = visibleColumns.filter(c => c !== colName);
  }

  updateCSVOutput();
}

function generateCSV(data, headers) {
  const rows = [headers.join(',')];
  data.forEach(row => {
    const line = headers.map(h => `"${(row[h] ?? '').replace(/"/g, '""')}"`).join(',');
    rows.push(line);
  });
  return rows.join('\n');
}

function getVisibleData() {
  const rows = [];
  const tbody = previewTable.querySelector('tbody');
  tbody.querySelectorAll('tr').forEach(tr => {
    if (tr.style.display !== 'none') {
      const row = {};
      tr.querySelectorAll('td').forEach(td => {
        const col = td.dataset.col;
        if (visibleColumns.includes(col)) {
          row[col] = td.textContent;
        }
      });
      rows.push(row);
    }
  });
  return rows;
}

function updateCSVOutput() {
  const data = getVisibleData();
  csvOutput = generateCSV(data, visibleColumns);
}

function downloadCSV() {
  if (!csvOutput || !selectedFile) {
    showAlert("❗ Nothing to download.");
    return;
  }
  const baseName = selectedFile.name.replace(/\.[^/.]+$/, '');
  const blob = new Blob([csvOutput], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `${baseName}_${sheetSelector.value}_visible.csv`;
  link.click();
}

searchInput.addEventListener('input', function () {
  const filter = this.value.toLowerCase();
  const rows = previewTable.querySelectorAll('tbody tr');
  let matchCounter = 0;

  rows.forEach(row => {
    let match = false;
    [...row.cells].forEach(cell => {
      const text = cell.textContent;
      if (filter && text.toLowerCase().includes(filter)) {
        const start = text.toLowerCase().indexOf(filter);
        const end = start + filter.length;
        cell.innerHTML =
          text.substring(0, start) +
          '<mark>' +
          text.substring(start, end) +
          '</mark>' +
          text.substring(end);
        match = true;
      } else {
        cell.innerHTML = cell.textContent;
      }
    });
    row.style.display = match || !filter ? '' : 'none';
    if (match || !filter) matchCounter++;
  });

  matchCount.textContent = matchCounter;
});

function resetApp() {
  selectedFile = null;
  workbook = null;
  csvOutput = '';
  visibleColumns = [];
  fileInput.value = '';
  previewTable.innerHTML = '';
  previewContainer.style.display = 'none';
  searchWrapper.style.display = 'none';
  columnToggles.style.display = 'none';
  tableStats.style.display = 'none';
  sheetSelector.innerHTML = '';
  document.getElementById('sheetSelectorWrapper').style.display = 'none';
  document.getElementById('status').textContent = '';
  downloadBtn.disabled = true;
  downloadBtn.innerHTML = '<i class="bi bi-download"></i> Download CSV';
  alertBox.classList.add('d-none');
  searchInput.value = '';
}

function showAlert(message) {
  alertBox.textContent = message;
  alertBox.classList.remove('d-none');
}

function clearAlert() {
  alertBox.classList.add('d-none');
  alertBox.textContent = '';
}
