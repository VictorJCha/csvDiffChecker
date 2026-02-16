const REQUIRED_HEADERS = {
  NETWORK: ["OVRC Name", "IP", "MAC Address"],
  UNIFI: ["Name", "IP Address", "MAC Address", "Expiration Time"],
  OVRC: ["Device Name", "IP Address", "MAC Address", "Status", "Monitored"]
};

const OUTPUT_FOLDER_ID = '1_w3Tj_fTvcwXJNAnpZ_uIHxZHstvG2fz';

/***********************
 * Web App
 ***********************/
function doGet(e) {
  // Serve the same HTML dialog you use for the file upload
  return HtmlService.createHtmlOutputFromFile('UploadDialog')
      .setTitle('Network CSV Comparison');
}

/***********************
 * WEB APP FILE PROCESSOR
 ***********************/
function processFilesForWeb(files) {
  const result = processFiles(files);

  const props = PropertiesService.getScriptProperties();

  props.setProperty('LAST_RESULTS', JSON.stringify(result.mismatches));
  props.setProperty('LAST_PROJECT', result.projectName || 'Current Project');

  return result;
}

/***********************
 * MAIN ENTRY
 ***********************/
function processFiles(files) {
  if (!files || files.length !== 3) {
    throw new Error('Expected exactly 3 CSV files.');
  }

  // Helper to parse project name from master file
  function getProjectNameFromFileName(filename) {
    if (!filename) return 'Current Project';

    // Remove file extension
    const baseName = filename.replace(/\.[^/.]+$/, ""); 

    // Find 'CSV' in filename (case-insensitive)
    const idx = baseName.toUpperCase().indexOf('CSV');

    if (idx <= 0) {
      // 'CSV' not found or nothing before CSV
      return 'Current Project';
    }

    // Take everything before CSV
    const beforeCSV = baseName.slice(0, idx);

    // Replace underscores and apostrophes with spaces, trim, collapse multiple spaces
    const clean = beforeCSV.replace(/[_']/g, ' ').trim().replace(/\s+/g, ' ');

    return clean || 'Current Project';
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

  // Extract project name
  const projectName = getProjectNameFromFileName(masterFileName);

  return {
    mismatches,
    projectName
  };
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
    headers.some(h => h.trim().toLowerCase() === req.toLowerCase())
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
 * EXPORT
 ***********************/
function exportResultsToSheet() {
  const props = PropertiesService.getScriptProperties();

  const rawResults = props.getProperty('LAST_RESULTS');
  if (!rawResults) {
    throw new Error('No comparison results found. Run the comparison first.');
  }

  const results = JSON.parse(rawResults);
  const projectName = props.getProperty('LAST_PROJECT') || 'Current Project';

  /************* TIMESTAMP *************/
  const now = new Date();
  const tz = Session.getScriptTimeZone();

  const datePart = Utilities.formatDate(now, tz, "MMddyyyy");
  const timePart = Utilities.formatDate(now, tz, "HHmmss");

  const timestamp = `${datePart} ${timePart}`;

  const ss = SpreadsheetApp.create(
    `Comparison Results - ${projectName} (${timestamp})`
  );


  const sheet = ss.getActiveSheet();

  const COLORS = {
    MASTER: '#cfe2f3',
    MATCH: '#cfe2f3',
    MISMATCH: '#f4cccc',
    MISSING: '#fff2cc'
  };

  const headers = [
    'IP Address',
    'Master Name',
    'UniFi Name',
    'OVRC Name',
    'Master MAC',
    'UniFi MAC',
    'OVRC MAC'
  ];

  const values = results.map(r => [
    r.ip,
    r.m.name,
    r.u?.name || '',
    r.o?.name || '',
    r.m.mac,
    r.u?.mac || '',
    r.o?.mac || ''
  ]);

  const totalRows = values.length + 1;
  const totalCols = headers.length;

  /************* FORCE PLAIN TEXT *************/
  sheet.getRange(1, 1, totalRows, totalCols).setNumberFormat('@');

  /************* WRITE HEADER *************/
  sheet.getRange(1, 1, 1, headers.length)
       .setValues([headers])
       .setFontWeight('bold')
       .setBackground('#000000')
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center'); // CENTER HEADER ONLY

  /************* WRITE DATA *************/
  sheet.getRange(2, 1, values.length, values[0].length)
       .setValues(values);

  /************* COLUMN WIDTHS *************/
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 255);
  sheet.setColumnWidth(6, 255);
  sheet.setColumnWidth(7, 255);

  sheet.setFrozenRows(1);

  /************* APPLY COLORS *************/
  results.forEach((r, i) => {
    const row = i + 2;

    sheet.getRange(row, 2).setBackground(COLORS.MASTER);
    sheet.getRange(row, 5).setBackground(COLORS.MASTER);

    applyCellColor(sheet.getRange(row, 3), r.m.name, r.u?.name);
    applyCellColor(sheet.getRange(row, 4), r.m.name, r.o?.name);

    applyMacColor(sheet.getRange(row, 6), r.m.mac, r.u?.mac);
    applyMacColor(sheet.getRange(row, 7), r.m.mac, r.o?.mac);
  });

  /************* MOVE FILE TO OUTPUT FOLDER *************/
  const file = DriveApp.getFileById(ss.getId());
  const folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return ss.getUrl();
}


function applyCellColor(cell, masterValue, compareValue) {
  const COLORS = {
    MATCH: '#cfe2f3',
    MISMATCH: '#f4cccc',
    MISSING: '#fff2cc'
  };

  if (!compareValue) {
    cell.setBackground(COLORS.MISSING);
    return;
  }

  const a = normalizeName(masterValue);
  const b = normalizeName(compareValue);

  cell.setBackground(a === b ? COLORS.MATCH : COLORS.MISMATCH);
}

function applyMacColor(cell, masterMac, compareMac) {
  const COLORS = {
    MATCH: '#cfe2f3',
    MISMATCH: '#f4cccc',
    MISSING: '#fff2cc'
  };

  if (!compareMac) {
    cell.setBackground(COLORS.MISSING);
    return;
  }

  const m = normalizeMac(masterMac);
  const c = normalizeMac(compareMac);

  cell.setBackground(m === c ? COLORS.MATCH : COLORS.MISMATCH);
}
