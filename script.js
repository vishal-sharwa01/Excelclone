let sheetCount = 0;
let cellData = {};
let currentCell = null;

function createSpreadsheet(containerId) {
  const container = document.getElementById(containerId);
  const spreadsheet = document.createElement('div');
  spreadsheet.className = 'spreadsheet';
  spreadsheet.dataset.sheetId = sheetCount;
  sheetCount++;

  // Create header row (top-left corner empty + A to T)
  const topLeft = document.createElement('div');
  topLeft.className = 'header';
  spreadsheet.appendChild(topLeft);

  for (let col = 0; col < 20; col++) {
    const colHeader = document.createElement('div');
    colHeader.className = 'header';
    colHeader.textContent = String.fromCharCode(65 + col);
    spreadsheet.appendChild(colHeader);
  }

  // Rows and cells
  for (let row = 1; row <= 25; row++) {
    const rowHeader = document.createElement('div');
    rowHeader.className = 'row-header';
    rowHeader.textContent = row;
    spreadsheet.appendChild(rowHeader);

    for (let col = 0; col < 20; col++) {
      const cell = document.createElement('div');
      cell.className = 'cell';
      cell.contentEditable = true;

      const cellId = 'S0_${String.fromCharCode(65 + col)}${row};'
      cell.dataset.cellId = cellId;

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
}

function format(command) {
  document.execCommand(command, false, null);
}

function applyColor(color, type) {
  if (currentCell) {
    if (type === 'color') {
      currentCell.style.color = color;
    } else {
      currentCell.style.backgroundColor = color;
    }
    const cellId = currentCell.dataset.cellId;
    cellData[cellId] = cellData[cellId] || {};
    cellData[cellId][type] = color;
  }
}

document.getElementById('fontSelect').addEventListener('change', (e) => {
  document.execCommand('fontName', false, e.target.value);
});

document.getElementById('fontSizeSelect').addEventListener('change', (e) => {
  document.execCommand('fontSize', false, e.target.value);
});

function addNewSheet() {
  createSpreadsheet('sheetContainer');
}

function exportToCSV() {
  let csv = '';
  for (let row = 1; row <= 25; row++) {
    let rowData = [];
    for (let col = 0; col < 20; col++) {
      let cellId = 'S0_${String.fromCharCode(65 + col)}${row};'
      rowData.push(cellData[cellId]?.value || '');
    }
    csv += rowData.join(',') + '\n';
  }

  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'spreadsheet.csv';
  a.click();
  URL.revokeObjectURL(url);
}

// Initialize first sheet
createSpreadsheet('sheetContainer');