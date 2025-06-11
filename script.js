let sheetCount = 0;
let cellData = {};
let currentCell = null;

function createSpreadsheet(containerId) {
  const container = document.getElementById(containerId);
  const spreadsheet = document.createElement('div');
  spreadsheet.className = 'spreadsheet';
  spreadsheet.dataset.sheetId = sheetCount;

  const topLeftCorner = document.createElement('div');
  topLeftCorner.className = 'header';
  spreadsheet.appendChild(topLeftCorner);

  // Column headers A-T
  for (let col = 0; col < 20; col++) {
    const colHeader = document.createElement('div');
    colHeader.className = 'header';
    colHeader.textContent = String.fromCharCode(65 + col);
    spreadsheet.appendChild(colHeader);
  }

  for (let row = 1; row <= 25; row++) {
    const rowHeader = document.createElement('div');
    rowHeader.className = 'row-header';
    rowHeader.textContent = row;
    spreadsheet.appendChild(rowHeader);

    for (let col = 0; col < 20; col++) {
      const cell = document.createElement('div');
      cell.className = 'cell';
      cell.contentEditable = true;

      const cellId = `S${sheetCount}_${String.fromCharCode(65 + col)}${row}`;
      cell.dataset.id = cellId;

      cell.addEventListener('click', () => {
        currentCell = cell;
      });

      cell.addEventListener('input', () => {
        cellData[cellId] = {
          value: cell.innerText,
          color: cell.style.color,
          bg: cell.style.backgroundColor
        };
      });

      spreadsheet.appendChild(cell);
    }
  }

  container.appendChild(spreadsheet);
  sheetCount++;
}

function format(command) {
  document.execCommand(command, false, null);
}

function applyColor(color, type) {
  if (currentCell) {
    currentCell.style[type] = color;
    const cellId = currentCell.dataset.id;
    cellData[cellId] = cellData[cellId] || {};
    cellData[cellId][type === 'color' ? 'color' : 'bg'] = color;
  }
}

document.getElementById('fontSelect').addEventListener('change', function () {
  document.execCommand('fontName', false, this.value);
});

document.getElementById('fontSizeSelect').addEventListener('change', function () {
  document.execCommand('fontSize', false, this.value);
});

function addNewSheet() {
  createSpreadsheet('sheetContainer');
}

function exportToCSV() {
  const rows = 25;
  const cols = 20;
  let csv = '';

  for (let row = 1; row <= rows; row++) {
    let rowData = [];
    for (let col = 0; col < cols; col++) {
      const cellId = `S0_${String.fromCharCode(65 + col)}${row}`;
      rowData.push((cellData[cellId] && cellData[cellId].value) || '');
    }
    csv += rowData.join(',') + '\n';
  }

  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'spreadsheet.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function importFromCSV() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.csv';

  input.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (event) {
      const text = event.target.result;
      const rows = text.trim().split('\n');
      const spreadsheet = document.querySelector('.spreadsheet');
      const cells = spreadsheet.querySelectorAll('.cell');

      let index = 0;
      for (let row = 0; row < rows.length; row++) {
        const cols = rows[row].split(',');
        for (let col = 0; col < cols.length; col++) {
          const cell = cells[index];
          if (cell) {
            cell.innerText = cols[col];
            const cellId = cell.dataset.id;
            cellData[cellId] = { value: cols[col], color: '', bg: '' };
          }
          index++;
        }
      }
    };
    reader.readAsText(file);
  });

  input.click();
}

createSpreadsheet('sheetContainer');
