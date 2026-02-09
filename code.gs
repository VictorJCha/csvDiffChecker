const REQUIRED_HEADERS = {
  NETWORK: ["OVRC Name", "IP", "MAC Address", "VLAN"],
  UNIFI: ["Name", "IP Address", "MAC Address", "Expiration Time"],
  OVRC: ["Device Name", "IP Address", "MAC Address", "Status", "Monitored"]
};

/***********************
 * Web App
 ***********************/
const DATA_SPREADSHEET_ID = '1vPnmjI65gultvQBvvZie7ftjzuWSs-SxHGBudvA0kd8'; // replace with your copied spreadsheet ID

function getDataSheet() {
  return SpreadsheetApp.openById(DATA_SPREADSHEET_ID);
}

function doGet(e) {
  // Serve the same HTML dialog you use for the file upload
  return HtmlService.createHtmlOutputFromFile('UploadDialog')
      .setTitle('Network CSV Comparison');
}

/***********************
 * MENU
 ***********************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Network Tools')
    .addItem('Upload CSVs', 'showUploadDialog')
    .addSeparator()
    .addItem('Reset Network Spreadsheet', 'resetNetworkSheet')
    .addToUi();

}

/***********************
 * UI
 ***********************/
function showUploadDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body {
        font-family: Arial, sans-serif;
        background: #f5f7fa;
        margin: 0;
        padding: 0;
      }

      .container {
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100%;
        padding: 20px;
      }

      .card {
        background: #ffffff;
        border-radius: 8px;
        padding: 24px 28px;
        width: 100%;
        max-width: 360px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        text-align: center;
      }

      h2 {
        margin-top: 0;
        margin-bottom: 12px;
        font-size: 18px;
      }

      p {
        margin: 0 0 16px;
        color: #555;
        font-size: 13px;
      }

      input[type="file"] {
        margin: 12px 0 20px;
        width: 100%;
      }

      button {
        background: #1a73e8;
        color: white;
        border: none;
        padding: 10px 16px;
        border-radius: 4px;
        font-size: 14px;
        cursor: pointer;
        width: 100%;
      }

      button:hover {
        background: #1558b0;
      }

      #loading {
        display: none;
        margin-top: 20px;
      }

      .spinner {
        margin: 0 auto 10px;
        width: 36px;
        height: 36px;
        border: 4px solid #ddd;
        border-top: 4px solid #1a73e8;
        border-radius: 50%;
        animation: spin 1s linear infinite;
      }

      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }

      .loading-text {
        font-size: 13px;
        color: #555;
      }
    </style>

    <div class="container">
      <div class="card">
        <h2>Network CSV Comparison</h2>
        <p>Select the 3 CSV files to compare</p>

        <input type="file" id="files" accept=".csv" multiple>

        <button onclick="submitFiles()">Run Comparison</button>

        <div id="loading">
          <div class="spinner"></div>
          <div class="loading-text">Processing files…</div>
        </div>
      </div>
    </div>

    <script>
      function submitFiles() {
        const input = document.getElementById('files');
        if (input.files.length !== 3) {
          alert('Please select exactly 3 CSV files.');
          return;
        }

        document.getElementById('loading').style.display = 'block';

        const readers = [];
        for (let i = 0; i < input.files.length; i++) {
          readers.push(new Promise(resolve => {
            const reader = new FileReader();
            reader.onload = e => resolve({
              name: input.files[i].name,
              content: e.target.result
            });
            reader.readAsText(input.files[i]);
          }));
        }

        Promise.all(readers).then(files => {
          google.script.run
            .withSuccessHandler(() => {
              document.getElementById('loading').style.display = 'none';
              google.script.host.close();
            })
            .withFailureHandler(err => {
              document.getElementById('loading').style.display = 'none';
              alert(err.message);
            })
            .processFiles(files);
        });
      }
    </script>
  `)
  .setWidth(420)
  .setHeight(360);

  SpreadsheetApp.getUi().showModalDialog(html, 'Upload CSVs');
}

/***********************
 * WEB APP FILE PROCESSOR
 ***********************/
function processFilesForWeb(files) {
  const mismatches = processFiles(files); // your existing function
  return mismatches;
}

/***********************
 * MAIN ENTRY
 ***********************/
function processFiles(files) {
  if (!files || files.length !== 3) {
    throw new Error('Expected exactly 3 CSV files.');
  }

  const parsed = files.map(f => parseCsvText(f.content));
  const identified = identifyFiles(parsed);

  const mismatches = compareByIp(
    identified.master,
    identified.unifi,
    identified.ovrc
  );

  // Grab master CSV filename
  const masterFile = files.find(f => f.name && hasRequiredHeaders(parseCsvText(f.content).headers, REQUIRED_HEADERS.NETWORK));
  const masterFileName = masterFile ? masterFile.name : '';

  // Pass to writeMismatches
  writeMismatches(mismatches, masterFileName);
  return mismatches;
}

/***********************
 * CSV PARSING
 ***********************/
function parseCsvText(text) {
  const rows = Utilities.parseCsv(text);
  const headers = rows.shift();
  return { headers, rows };
}

/***********************
 * FILE IDENTIFICATION
 ***********************/
function identifyFiles(files) {
  const identified = {};

  files.forEach(file => {
    const headers = file.headers.map(h => h.trim());

    Logger.log("Headers found: " + headers.join(", "));

    if (hasRequiredHeaders(headers, REQUIRED_HEADERS.NETWORK)) {
      identified.master = file;
    }
    else if (hasRequiredHeaders(headers, REQUIRED_HEADERS.UNIFI)) {
      identified.unifi = file;
    }
    else if (hasRequiredHeaders(headers, REQUIRED_HEADERS.OVRC)) {
      identified.ovrc = file;
    }
    else {
      throw new Error(
        "Unrecognized CSV format.\n\nHeaders found:\n" + headers.join(", ")
      );
    }
  });

  if (!identified.master || !identified.unifi || !identified.ovrc) {
    throw new Error("All three CSV files must be uploaded.");
  }

  return identified;
}

/***********************
 * HEADERS
 ***********************/
function hasRequiredHeaders(headers, required) {
  return required.every(req =>
    headers.some(h => h.toLowerCase() === req.toLowerCase())
  );
}

/***********************
 * NORMALIZATION
 ***********************/
function normalizeMac(mac) {
  if (!mac) return '';

  const raw = mac.toString().trim();

  // Extract hex characters only
  const hex = raw.replace(/[^A-Fa-f0-9]/g, '').toUpperCase();

  // If it doesn't look like a real MAC (12 hex chars), keep original text
  if (hex.length !== 12) {
    return raw;
  }

  return hex;
}

// Strip non-alphanumeric chars from names for comparison
function normalizeName(name) {
  return name?.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
}

/***********************
 * MAP BUILDERS
 ***********************/
function toIpMap(csv, cols, type) {
  const map = {};

  csv.rows.forEach(r => {
    const ip = r[cols.ip]?.trim();
    if (!ip) return;

    const entry = {
      name: r[cols.name],
      mac: normalizeMac(r[cols.mac]),
      status: type === 'ovrc' ? r[csv.headers.indexOf('Status')]?.trim() : null
    };

    if (type === 'unifi') {
      map[ip] = map[ip] || [];
      map[ip].push(entry);
    } else if (type === 'ovrc') {
      map[ip] = map[ip] || [];
      map[ip].push(entry);
    } else {
      map[ip] = entry;
    }
  });

  // OVRC duplicate post-processing
  if (type === 'ovrc') {
    Object.keys(map).forEach(ip => {
      const rows = map[ip];
      if (rows.length === 1) {
        map[ip] = rows[0];
      } else {
        const healthyRows = rows.filter(e => e.status?.toLowerCase() !== 'critical');
        if (healthyRows.length > 0) map[ip] = healthyRows;
        else map[ip] = rows; // all critical
      }
    });
  }

  return map;
}

/***********************
 * COMPARISON
 ***********************/
function compareByIp(masterCsv, unifiCsv, ovrcCsv) {
  const master = toIpMap(masterCsv, {
    name: masterCsv.headers.indexOf('OVRC Name'),
    ip: masterCsv.headers.indexOf('IP'),
    mac: masterCsv.headers.indexOf('MAC Address')
  }, 'network');

  const unifi = toIpMap(unifiCsv, {
    name: unifiCsv.headers.indexOf('Name'),
    ip: unifiCsv.headers.indexOf('IP Address'),
    mac: unifiCsv.headers.indexOf('MAC Address')
  }, 'unifi');

  const ovrc = toIpMap(ovrcCsv, {
    name: ovrcCsv.headers.indexOf('Device Name'),
    ip: ovrcCsv.headers.indexOf('IP Address'),
    mac: ovrcCsv.headers.indexOf('MAC Address')
  }, 'ovrc');

  const mismatches = [];

  Object.keys(master).forEach(ip => {
    const m = master[ip];
    const uArray = unifi[ip] || [];
    let oArray = ovrc[ip];
    if (!Array.isArray(oArray)) oArray = oArray ? [oArray] : [];

    if (uArray.length === 0 && oArray.length === 0) {
      mismatches.push({ ip, m, u: null, o: null });
      return;
    }

    if (uArray.length === 0 && oArray.length > 0) {
      oArray.forEach(o => mismatches.push({
        ip, m, u: null, o,
        macMismatch: o.mac !== m.mac,
        nameMismatch: normalizeName(o.name) !== normalizeName(m.name)
      }));
      return;
    }

    if (uArray.length > 0 && oArray.length === 0) {
      uArray.forEach(u => mismatches.push({
        ip, m, u, o: null,
        macMismatch: u.mac !== m.mac,
        nameMismatch: normalizeName(u.name) !== normalizeName(m.name)
      }));
      return;
    }

    uArray.forEach(u => {
      oArray.forEach(o => {
        const macMismatch = u.mac !== m.mac || o.mac !== m.mac;
        const nameMismatch = normalizeName(u.name) !== normalizeName(m.name) ||
                             normalizeName(o.name) !== normalizeName(m.name);
        if (macMismatch || nameMismatch) {
          mismatches.push({ ip, m, u, o, macMismatch, nameMismatch });
        }
      });
    });
  });

  // Deduplicate
  const uniqueMismatches = [];
  const seen = new Set();
  mismatches.forEach(r => {
    const key = `${r.ip}|${r.u?.name ?? ''}|${r.o?.name ?? ''}`;
    if (!seen.has(key)) {
      seen.add(key);
      uniqueMismatches.push(r);
    }
  });

  return uniqueMismatches;
}

/***********************
 * OUTPUT
 ***********************/
 function getProjectSheetName(masterFileName) {
  try {
    if (masterFileName && masterFileName.includes('CSV')) {
      let name = masterFileName
        .split('CSV')[0]
        .trim()
        .replace(/_/g, ' ');

      // Google Sheets limit
      return name.length > 100 ? name.substring(0, 100) : name;
    }
  } catch (e) {}

  return 'Mismatches'; // safe fallback
}

function getOrCreateOutputSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
    sheet.setFrozenRows(0);
  } else {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function writeMismatches(rows, masterFileName) {
  const ss = getDataSheet();

  // Extract site name before "CSV" in filename
    const sheetName = getProjectSheetName(masterFileName);
    const sheet = getOrCreateOutputSheet(ss, sheetName);

  // Write header
  sheet.appendRow([
    'IP Address',
    'Master Name',
    'UniFi Name',
    'OVRC Name',
    'Master MAC',
    'UniFi MAC',
    'OVRC MAC'
  ]);

  // Style header row
  sheet.getRange(1, 1, 1, 7)
    .setBackground('#000000')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  //Force data rows into plain text
  sheet
  .getRange(2, 1, sheet.getMaxRows(), sheet.getMaxColumns())
  .setNumberFormat('@');

  // Write each mismatch row
  rows.forEach((r, i) => {
    const row = i + 2;

    sheet.getRange(row, 1, 1, 7).setValues([[
      r.ip,
      r.m.name,
      r.u?.name ?? '',
      r.o?.name ?? '',
      r.m.mac,
      r.u?.mac ?? '',
      r.o?.mac ?? ''
    ]]);

    // Master highlight (blue)
    setCellStatus(sheet.getRange(row, 2), 'master');
    setCellStatus(sheet.getRange(row, 5), 'master');

    // UniFi highlights
    if (!r.u) {
      setCellStatus(sheet.getRange(row, 3), 'missing');
      setCellStatus(sheet.getRange(row, 6), 'missing');
    } else {
      setCellStatus(sheet.getRange(row, 3), r.u.name === r.m.name ? 'match' : 'mismatch');
      setCellStatus(sheet.getRange(row, 6), r.u.mac === r.m.mac ? 'match' : 'mismatch');
    }

    // OVRC highlights
    if (!r.o) {
      setCellStatus(sheet.getRange(row, 4), 'missing');
      setCellStatus(sheet.getRange(row, 7), 'missing');
    } else {
      setCellStatus(sheet.getRange(row, 4), r.o.name === r.m.name ? 'match' : 'mismatch');
      setCellStatus(sheet.getRange(row, 7), r.o.mac === r.m.mac ? 'match' : 'mismatch');
    }
  });

  // Set fixed column widths
  sheet.setColumnWidth(1, 100);  // IP Address
  sheet.setColumnWidth(2, 300);  // Master Name
  sheet.setColumnWidth(3, 300);  // UniFi Name
  sheet.setColumnWidth(4, 300);  // OVRC Name
  sheet.setColumnWidth(5, 140);  // Master MAC
  sheet.setColumnWidth(6, 140);  // UniFi MAC
  sheet.setColumnWidth(7, 140);  // OVRC MAC

  // Wrap text and freeze header
  sheet.getRange(1, 1, sheet.getLastRow(), 7).setWrap(true);
  sheet.setFrozenRows(1);
}

/***********************
 * COLOR SCHEME
 ***********************/
function setCellStatus(range, status) {
  const COLORS = {
    master: '#cfe2f3',   // blue
    match:  '#cfe2f3',   // blue
    mismatch: '#f4cccc', // red
    missing: '#fff2cc'   // yellow
  };

  range.setBackground(COLORS[status] || null);
}

/***********************
 * RESET SPREADSHEET
 ***********************/
function resetNetworkSheet() {
  const ss = getDataSheet();
  let instructionsSheet = ss.getSheetByName('Instructions');

  // If Instructions sheet exists, just clear it
  if (instructionsSheet) {
    instructionsSheet.clear();
    instructionsSheet.setFrozenRows(0);
  } else {
    // If it doesn't exist, rename the first sheet to 'Instructions'
    instructionsSheet = ss.getSheets()[0];
    instructionsSheet.setName('Instructions');
    instructionsSheet.clear();
    instructionsSheet.setFrozenRows(0);
  }

  // Delete all other sheets
  ss.getSheets().forEach(sheet => {
    if (sheet.getName() !== 'Instructions') {
      ss.deleteSheet(sheet);
    }
  });

  // Write instructions
  writeInstructions(instructionsSheet);
}

/***********************
 * WRITE INSTRUCTIONS
 ***********************/
function writeInstructions(sheet) {
  const instructions = [
    ['Network CSV Comparison Tool'],
    [''],
    ['How to use:'],
    ['1. Click "Network Tools" > "Upload CSVs".'],
    ['2. Upload the three CSV files when prompted:'],
    ['   • Network Spreadsheet'],
    ['   • UniFi CSV export (Client Devices > DHCP Manager > Export)'],
    ['   • OVRC CSV export (Devices > Download CSV)'],
    ['3. The script will compare devices by IP address.'],
    ['4. A "Mismatches" table will be generated automatically.'],
    [''],
    ['Color legend:'],
    ['Matches Network Spreadsheet'],
    ['Mismatch Needs Correction'],
    ['Missing From Platform']
  ];

  sheet.getRange(1, 1, instructions.length, 1).setValues(instructions);

  // Formatting
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A3').setFontWeight('bold');
  sheet.getRange('A11').setFontWeight('bold');
  sheet.setColumnWidth(1, 600);

  // Apply legend colors using existing function
  setCellStatus(sheet.getRange('A13'), 'master');   // Matches
  setCellStatus(sheet.getRange('A14'), 'mismatch'); // Needs correction
  setCellStatus(sheet.getRange('A15'), 'missing');  // Missing
}
