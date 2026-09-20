// ============================================================
// WLP Audit Script v2 — Valley School
// Outputs three JSON files to the WLP root folder:
//   1. WLP_FileTree.json     — full recursive folder/file tree
//   2. WLP_FileIndex.json    — per-subject-folder file inventory
//                              with naming checks + count checks
//   3. WLP_SheetData.json    — per-sheet content extraction
//                              (headers, lesson entries, holidays)
// ============================================================


// ── CONFIG ──────────────────────────────────────────────────
const WLP_FOLDER_ID = '1k5EJb5Zy1D8DX1H1Sj7FDaCTnbQ0majn';

// All folders 1_vzIv8_oO6Sota1i3ZoQT13i3esiVIEQ
// Foundation 1k5EJb5Zy1D8DX1H1Sj7FDaCTnbQ0majn
// Primary folder 1QMxCwZ-wIVblJz2-zsDRqdedTjzBDw-f
// Lower primary 1RF9Te6qY25tZMF8dU_W0Q02sRpkeo6Tk
// Upper primary 1UG-4jPlheLwpoB0uY-VrMV19UdF9RXJE
// Secondary folder 1sQXG3S5HENujFsG6xziaIYqpHnbxYVmO
// Lower secondary 1lpOk4FAgobqSyu_iiUvxpAXqBUnsxh-u
// Upper secondary 1XwR2VspZpmxqOzaPDfV3BlqwaWNCMRWq

// Just ICT 1nDJ7lSpODrsgxi-tXt7WjG5ppbew7HUe

//'1_vzIv8_oO6Sota1i3ZoQT13i3esiVIEQ';

// Naming convention: [Class]_WLP_[Subject]_T[N]_AY25-26_VIS
// Both .xlsx and Google Sheets native format are acceptable.
const NAMING_REGEX = /^[A-Za-z0-9]+_WLP_[A-Za-z0-9\-]+_T[1-4]_AY25-26_VIS$/;

// Valid file MIME types
const VALID_MIME_TYPES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
  'application/vnd.ms-excel',                                           // .xls (flag as warning)
  'application/vnd.google-apps.spreadsheet'                             // Google Sheets native
];

// Week sheet detection: matches "Week 1", "Week 3", "W5", etc.
const WEEK_SHEET_REGEX = /^(Week\s*\d+|W\d+)$/i;

// Labels we search for in the header block (rows 1–15, flexible)
const HEADER_LABELS = ['Name', 'Subject', 'Class', 'Week', 'Term'];

// Min non-empty cells in a row to be considered the table header row
// (avoids false positives from sparse info rows)
const TABLE_HEADER_MIN_CELLS = 1;

// Expected term file count per year/subject folder
const EXPECTED_TERM_COUNT = 3;


// ════════════════════════════════════════════════════════════
// MAIN ENTRY POINT
// ════════════════════════════════════════════════════════════
function runFullAudit() {
  const rootFolder = DriveApp.getFolderById(WLP_FOLDER_ID);
  Logger.log('Starting WLP Audit v2...');

  // ── Output 1: File tree ──
  Logger.log('--- Building file tree ---');
  const fileTree = buildFileTree(rootFolder);
  writeJsonToDrive(rootFolder, 'WLP_FileTree.json', fileTree);
  Logger.log('File tree written.');

  // ── Output 2: File index (flat, per-subject-folder inventory) ──
  Logger.log('--- Building file index ---');
  const fileIndex = buildFileIndex(rootFolder);
  writeJsonToDrive(rootFolder, 'WLP_FileIndex.json', fileIndex);
  Logger.log(`File index written. ${fileIndex.length} subject folders indexed.`);

  // ── Output 3: Sheet data extraction ──
  Logger.log('--- Extracting sheet data ---');
  const sheetData = extractAllSheetData(fileIndex);
  writeJsonToDrive(rootFolder, 'WLP_SheetData.json', sheetData);
  Logger.log(`Sheet data written. ${sheetData.length} files processed.`);

  Logger.log('=== Audit complete. Three JSON files written to WLP root folder. ===');
}


// ════════════════════════════════════════════════════════════
// OUTPUT 1 — File Tree
// Produces a nested JSON tree mirroring the Drive folder
// hierarchy. Claude uses this to check folder structure
// against WLP_Folder_Structure_Reference.md.
// ════════════════════════════════════════════════════════════

/**
 * Recursively builds a nested folder/file tree object.
 * Each node: { name, type: "folder"|"file", id, mimeType?, children? }
 */
function buildFileTree(folder) {
  const node = {
    name: folder.getName(),
    type: 'folder',
    id: folder.getId(),
    children: []
  };

  // Subfolders first
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    node.children.push(buildFileTree(subfolders.next()));
  }

  // Files
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    node.children.push({
      name: f.getName(),
      type: 'file',
      id: f.getId(),
      mimeType: f.getMimeType()
    });
  }

  return node;
}


// ════════════════════════════════════════════════════════════
// OUTPUT 2 — File Index
// A flat list of "leaf" folders (the deepest folder level
// that contains files directly), with per-file naming checks
// and a count-vs-expectation check.
//
// "Leaf folder" = a folder that contains at least one file.
// We assume these are the subject/year folders where teachers
// deposit their WLP files.
//
// Claude uses this to:
//   - Flag naming convention violations
//   - Flag folders with wrong number of term files (≠ 3)
//   - Flag unexpected extra files
// ════════════════════════════════════════════════════════════

function buildFileIndex(rootFolder) {
  const index = [];
  collectLeafFolders(rootFolder, [], index);
  return index;
}

function collectLeafFolders(folder, pathStack, index) {
  pathStack.push(folder.getName());

  const fileIter = folder.getFiles();
  const filesHere = [];

  while (fileIter.hasNext()) {
    const f = fileIter.next();
    filesHere.push(f);
  }

  // If this folder has files, record it as a leaf
  if (filesHere.length > 0) {
    const folderPath = pathStack.join(' / ');
    const fileRecords = filesHere.map(function(f) {
      return assessFile(f, pathStack);
    });

    // Count only WLP-intended files for the term count check
    // (ignore JSON outputs, non-spreadsheet files)
    const wlpFiles = fileRecords.filter(r => r.looksLikeWlpFile);
    const termCount = wlpFiles.length;

    index.push({
      folderPath: folderPath,
      folderId: folder.getId(),
      totalFilesFound: fileRecords.length,
      wlpFileCount: termCount,
      expectedTermCount: EXPECTED_TERM_COUNT,
      termCountCorrect: termCount === EXPECTED_TERM_COUNT,
      termCountNote: termCount < EXPECTED_TERM_COUNT
        ? `Missing ${EXPECTED_TERM_COUNT - termCount} term file(s)`
        : termCount > EXPECTED_TERM_COUNT
        ? `${termCount - EXPECTED_TERM_COUNT} unexpected extra file(s) found`
        : 'OK',
      files: fileRecords
    });
  }

  // Always recurse into subfolders
  const subfolderIter = folder.getFolders();
  while (subfolderIter.hasNext()) {
    collectLeafFolders(subfolderIter.next(), pathStack.slice(), index);
  }

  pathStack.pop();
}

/**
 * Assesses a single file: MIME type, naming convention, parsed tokens,
 * and whether the class token matches its parent folder name.
 */
function assessFile(file, pathStack) {
  const fileName = file.getName();
  const fileId   = file.getId();
  const mimeType = file.getMimeType();
  const baseName = stripExtension(fileName);

  // Is this plausibly a WLP file at all?
  // (Excludes JSON audit outputs, PDFs, images, etc.)
  const isSpreadsheet = VALID_MIME_TYPES.includes(mimeType);
  const looksLikeWlpFile = isSpreadsheet || /WLP/i.test(fileName);

  // MIME type check
  let mimeNote = null;
  if (!isSpreadsheet) {
    mimeNote = `Unexpected file type: ${mimeType}. Only .xlsx and Google Sheets are acceptable.`;
  } else if (mimeType === 'application/vnd.ms-excel') {
    mimeNote = 'Warning: .xls format (legacy). Should be .xlsx or Google Sheets.';
  }

  // Naming convention check
  const namingValid = NAMING_REGEX.test(baseName);
  const namingViolation = namingValid ? null : describeNamingViolation(baseName);

  // Parse tokens from filename if naming is valid
  let parsedClass = null, parsedSubject = null, parsedTerm = null;
  if (namingValid) {
    const parts  = baseName.split('_');
    parsedClass   = parts[0];
    parsedSubject = parts[2];
    parsedTerm    = parts[3];
  }

  // Class-location consistency check:
  // The immediate parent folder name should match parsedClass,
  // OR the grandparent should (for Upper Secondary where no year folder exists).
  let locationConsistencyNote = null;
  if (parsedClass) {
    const parentFolder = pathStack[pathStack.length - 1]; // current deepest
    if (parentFolder !== parsedClass) {
      locationConsistencyNote =
        `Class token "${parsedClass}" in filename does not match parent folder "${parentFolder}". ` +
        `Full path: ${pathStack.join(' / ')}`;
    }
  }

  return {
    fileId,
    fileName,
    baseName,
    mimeType,
    isSpreadsheet,
    looksLikeWlpFile,
    mimeNote,
    namingValid,
    namingViolation,
    parsedClass,
    parsedSubject,
    parsedTerm,
    locationConsistencyNote
  };
}


// ════════════════════════════════════════════════════════════
// OUTPUT 3 — Sheet Data Extraction  (REVISED)
//
// KEY CHANGE: Column headers are no longer matched against a
// fixed allowlist. Instead we:
//   1. Find the table's top-left anchor row (the first row
//      after the header block that has ≥ 3 non-empty cells).
//   2. Read the next 1–2 rows to build a two-tier header map
//      (parent label + child label → column index).
//   3. For each lesson block, emit { "<HeaderName>": value }
//      for every column that has a header — no hardcoded field
//      names, so template drift is handled automatically.
// ════════════════════════════════════════════════════════════




// The first column label is always the "anchor" — it identifies
// the Day/Date cell. We use it to find the left edge of the table.
// Any non-empty cell in row 1 of the header band qualifies.


function extractAllSheetData(fileIndex) {
  const results = [];

  fileIndex.forEach(function(folder) {
    folder.files.forEach(function(fileRecord) {
      if (!fileRecord.isSpreadsheet) return;

      Logger.log(`Processing: ${fileRecord.fileName}`);

      const driveMeta = getDriveFileMeta(fileRecord.fileId);  // ← Get DriveMeta

      let spreadsheetId = null;
      let tempCreated   = false;

      try {
        if (fileRecord.mimeType === 'application/vnd.google-apps.spreadsheet') {
          spreadsheetId = fileRecord.fileId;
        } else {
          spreadsheetId = convertToGoogleSheets(fileRecord.fileId, fileRecord.fileName);
          tempCreated   = true;
        }

        const ss     = SpreadsheetApp.openById(spreadsheetId);
        const sheets = ss.getSheets();

        const allSheetNames  = sheets.map(s => s.getName());
        const weekSheets     = sheets.filter(s => WEEK_SHEET_REGEX.test(s.getName()));
        const otherSheets    = sheets.filter(s => !WEEK_SHEET_REGEX.test(s.getName()));

        const weekData = weekSheets.map(function(sheet) {
          return extractWeekSheet(sheet, spreadsheetId);
        }).filter(Boolean);

        results.push({
          fileId:          fileRecord.fileId,
          fileName:        fileRecord.fileName,
          folderPath:      folder.folderPath,
          lastModifyingUser: driveMeta.lastModifyingUser,
          skipped:         false,
          allSheetNames:   allSheetNames,
          weekSheetNames:  weekSheets.map(s => s.getName()),
          otherSheetNames: otherSheets.map(s => s.getName()),
          weekSheetCount:  weekSheets.length,
          weekData:        weekData
        });

      } catch (e) {
        results.push({
          fileId:   fileRecord.fileId,
          fileName: fileRecord.fileName,
          lastModifyingUser: driveMeta.lastModifyingUser,
          skipped:  true,
          reason:   `Extraction error: ${e.message}`
        });
        Logger.log(`Extraction error: ${e.message}`);
      } finally {
        if (tempCreated && spreadsheetId) {
          try { DriveApp.getFileById(spreadsheetId).setTrashed(true); } catch (e) { 
            Logger.log(`Error trashing file: ${e.message}`);}
        }
      }
    });
  });

  return results;
}


function extractWeekSheet(sheet, spreadsheetId) {
  const sheetName = sheet.getName();
  const dataRange = sheet.getDataRange();
  const data      = dataRange.getValues();
  const merges    = dataRange.getMergedRanges();

  // Blank check
  const allValues = data.flat().filter(v => v !== null && v !== '' && String(v).trim() !== '');
  if (allValues.length < 3) return null;

  const mergeMap = buildMergeMap(merges, data);
  const headers  = extractHeaderFields(data, mergeMap);

  // ── Find table: anchor on the first row with ≥ TABLE_HEADER_MIN_CELLS non-empty cells
  //    that appears AFTER the info header block (after row 8 at earliest,
  //    but we scan dynamically).
  const tableInfo = findTableDynamic(data, mergeMap);

  if (!tableInfo) {
    return {
      sheetName: sheetName,
      headers:   headers,
      columnHeaders: [],
      lessonCount:   0,
      holidayCount:  0,
      lessons:   [],
      holidays:  [],
      notes: 'Could not locate lesson table — no row with ≥ 3 non-empty cells found below header block'
    };
  }

  const { columnHeaders, tableDataStartRow } = tableInfo;
  const { lessons, holidays } = extractLessonBlocksDynamic(data, mergeMap, tableInfo, spreadsheetId, sheet.getName());
  Logger.log(`Number of lessons: ${lessons.length})}, Number of holidays ${holidays.length}`)


  return {
    sheetName:     sheetName,
    headers:       headers,
    columnHeaders: columnHeaders,   // [ { label, col }, ... ] — for transparency
    lessonCount:   lessons.length,
    holidayCount:  holidays.length,
    lessons:       lessons,         // each lesson: { rowStart, rowEnd, fields: { "<header>": value } }
    holidays:      holidays,
    notes:         null
  };
}


/**
 * Dynamically locates the lesson table by finding the first row
 * (after an initial skip zone, rows 0–7) that has ≥ TABLE_HEADER_MIN_CELLS
 * non-empty cells. That's the top header row.
 *
 * Then checks whether the very next row is a "sub-header" row
 * (also has several non-empty cells but the first anchor column is empty).
 * If so, merges the two rows into a combined label per column.
 *
 * Returns:
 *   {
 *     tableHeaderRow,       // 0-indexed row of first header row
 *     tableDataStartRow,    // 0-indexed first data row
 *     anchorCol,            // column index of the leftmost non-empty header cell
 *     columnHeaders,        // [ { label: string, col: number }, ... ]
 *   }
 * or null if not found.
 */
function findTableDynamic(data, mergeMap) {
  const SKIP_ROWS = 7;   // rows 0–7 are the info header block; skip them

  let tableHeaderRow = null;
  let anchorCol      = null;

  // Find first candidate row
  for (let r = SKIP_ROWS; r < data.length; r++) {
    const row       = data[r];
    const nonEmpty  = row.filter(v => v !== null && v !== '' && String(v).trim() !== '');
    if (nonEmpty.length >= TABLE_HEADER_MIN_CELLS) {
      tableHeaderRow = r;
      //Logger.log(`Anchor row: ${r}`);
      // Find leftmost non-empty cell = anchor column
      for (let c = 0; c < row.length; c++) {
        if (row[c] !== null && row[c] !== '' && String(row[c]).trim() !== '') {
          anchorCol = c;
          //Logger.log(`Anchor column: ${c}`)
          break;
        }
      }
      break;
    }
  }

  if (tableHeaderRow === null) return null;

  // Build tier-1 labels from tableHeaderRow
  const tier1 = buildHeaderTier(data[tableHeaderRow], tableHeaderRow, mergeMap);
  //Logger.log(tier1);

  // Check if the next row is a sub-header tier
  // Heuristic: next row has ≥ 2 non-empty cells AND the anchor column is empty
  let tableDataStartRow = tableHeaderRow + 1;
  const nextRow = data[tableHeaderRow + 1] || [];
  const nextNonEmpty = nextRow.filter(v => v !== null && v !== '' && String(v).trim() !== '');
  const nextAnchorEmpty = !nextRow[anchorCol] || String(nextRow[anchorCol]).trim() === '';

  let columnHeaders = [];

  if (nextNonEmpty.length >= 2 && nextAnchorEmpty) {
    // Two-tier header: merge tier1 + tier2
    const tier2 = buildHeaderTier(nextRow, tableHeaderRow + 1, mergeMap);
    //Logger.log(tier2);
    tableDataStartRow = tableHeaderRow + 2;

    // For each column, label = tier1 label if it has one, else tier1 of its parent merge + "/" + tier2
    // Simpler: prefer the most specific (tier2) label; if tier2 is empty, use tier1.
    const allCols = new Set([
      ...tier1.map(h => h.col),
      ...tier2.map(h => h.col)
    ]);

    allCols.forEach(function(col) {
      const t1 = tier1.find(h => h.col === col);
      const t2 = tier2.find(h => h.col === col);

      let label = null;
      if (t2 && t1) {
        // Both tiers have something at this column
        label = t1.label + ' / ' + t2.label;
      } else if (t2) {
        // Only tier2 — find the tier1 parent (the most recent tier1 to the left)
        const parentT1 = tier1.filter(h => h.col <= col).pop();
        label = parentT1 ? parentT1.label + ' / ' + t2.label : t2.label;
      } else if (t1) {
        // Only tier1 — and no sub-columns override it
        // Include only if tier2 doesn't have any children that would supersede it
        const hasChildren = tier2.some(h => h.col > t1.col);
        label = t1.label;
      }

      //Logger.log(label);
      if (label) {
        columnHeaders.push({ label: label.trim(), col: col });
      }
    });

    // Sort by column position
    columnHeaders.sort((a, b) => a.col - b.col);

  } else {
    // Single-tier header
    columnHeaders = tier1;
  }

  return { tableHeaderRow, tableDataStartRow, anchorCol, columnHeaders };
}


/**
 * Reads one row and returns non-empty cells as { label, col } objects.
 * Resolves merged cell values where needed.
 */
function buildHeaderTier(row, rowIdx, mergeMap) {
  const result = [];
  for (let c = 0; c < row.length; c++) {
    const val = getMergeTopLeft(rowIdx, c, mergeMap) ?? row[c];
    const str = String(val || '').trim();
    if (str !== '') {
      result.push({ label: str, col: c });
    }
  }
  return result;
}


/**
 * Returns the top-left value of the merge that owns cell (row, col),
 * or null if the cell is not the top-left of any merge.
 * (Only returns a value if (row, col) IS the top-left — not a covered member.)
 */
function getMergeTopLeft(row, col, mergeMap) {
  if (mergeMap[row] && mergeMap[row][col]) {
    return mergeMap[row][col].value;
  }
  return null;
}


/**
 * Extracts lesson blocks using the dynamic column header map.
 * Each lesson's `fields` object uses the column header label as the key.
 */
function extractLessonBlocksDynamic(data, mergeMap, tableInfo, spreadsheetId, sheetName) {
  const { tableDataStartRow, anchorCol, columnHeaders } = tableInfo;
  const lessons  = [];
  const holidays = [];

  // Fetch bottom borders for the anchor/date column across all data rows
  const borderedRows = fetchBottomBorderedRows(
    spreadsheetId,
    sheetName,
    anchorCol,
    tableDataStartRow,
    data.length - 1
  );

/**
  // ── Fallback signal: pre-compute which rows have a NEW non-empty
  //    anchor value (i.e. the day/date changes).
  //
  //    Because anchor cells are often merged, getCellValue returns the
  //    merge's top-left value for every row in the merge — meaning the
  //    value looks the same for all rows in one day's block and then
  //    changes on the first row of the next block.
  //
  //    anchorChangeRows[r] = true  means row r starts a NEW anchor value.
  //    We use this to infer that row (r - 1) was the last row of the
  //    previous block.
  const anchorChangeRows = new Set();
  let lastSeenAnchor = null;

  for (let r = tableDataStartRow; r < data.length; r++) {
    const val = clean(getCellValue(r, anchorCol, mergeMap, data));
    if (val !== null && val !== lastSeenAnchor) {
      anchorChangeRows.add(r);
      lastSeenAnchor = val;
    }
  }
  **/

  // Group rows into lesson blocks.
  // A block ends when its last row has a bottom border in the date column. OR see lastSeenAnchor fallback.
  let blockRows = [];

  for (let r = tableDataStartRow; r < data.length; r++) {
 
 /**
    // Before adding this row: if it starts a new anchor value AND we
    // already have rows collected, the previous block just ended.
    // (This is the fallback — fires when no bottom border was present.)
    if (anchorChangeRows.has(r) && blockRows.length > 0) {
      const hasBorderOnLastRow = borderedRows.has(blockRows[blockRows.length - 1]);
      if (!hasBorderOnLastRow) {
        // Flush the current block early via fallback
        flushBlock(blockRows, lessons, holidays, columnHeaders, data, mergeMap);
        blockRows = [];
      }
      // If border signal already closed this block, blockRows is already
      // empty and we just fall through normally.
    }
**/
    
    blockRows.push(r);

    const isBlockEnd = borderedRows.has(r);  // bottom border = end of this lesson
    const isLastRow  = r === data.length - 1;

    if (isBlockEnd || isLastRow) {
      flushBlock(blockRows, lessons, holidays, columnHeaders, data, mergeMap);
      blockRows = [];  // reset for next block
    }
  }

  
  return { lessons, holidays };
}

/**
 * Finalises one collected block of rows into a lesson or holiday entry.
 * Extracted into its own function so both the border path and the
 * fallback path share identical output logic.
 */
function flushBlock(blockRows, lessons, holidays, columnHeaders, data, mergeMap) {
  if (blockRows.length === 0) return;

  const startRow = blockRows[0];
  const endRow   = blockRows[blockRows.length - 1];

  // Holiday detection (unchanged — scan all rows in block)
  const isHoliday = checkForHoliday(startRow, blockRows.length, data, mergeMap);

  if (isHoliday) {
    holidays.push({
      rowStart: startRow + 1,
      rowEnd:   endRow + 1,
      rawNote:  extractHolidayText(startRow, blockRows.length, data, mergeMap)
    });
  } else {
    // Build fields — for multi-row blocks, collect ALL non-null values per column
    const fields = {};
    columnHeaders.forEach(function(header) {
      const vals = blockRows
        .map(r2 => clean(getCellValue(r2, header.col, mergeMap, data)))
        .filter(v => v !== null);
      // Prefer merged/first value; join multiples if they differ
      const unique = [...new Set(vals)];
      fields[header.label] = unique.length === 1 ? unique[0]
                            : unique.length > 1  ? unique.join(' | ')
                            : null;
    });

    lessons.push({
      rowStart: startRow + 1,
      rowEnd:   endRow + 1,
      fields:   fields
    });
  }
}


/**
 * Builds a map of merged cell regions.
 * Returns an object: mergeMap[row][col] = { topLeftRow, topLeftCol, rowSpan, colSpan, value }
 * Only populated for the top-left cell of each merged range.
 */
function buildMergeMap(merges, data) {
  const map = {};

  merges.forEach(function(range) {
    const r1 = range.getRow() - 1;          // 0-indexed
    const c1 = range.getColumn() - 1;       // 0-indexed
    const r2 = range.getLastRow() - 1;
    const c2 = range.getLastColumn() - 1;

    const value = (r1 < data.length && c1 < data[r1].length)
      ? data[r1][c1]
      : null;

    if (!map[r1]) map[r1] = {};
    map[r1][c1] = {
      topLeftRow: r1,
      topLeftCol: c1,
      rowSpan: r2 - r1 + 1,
      colSpan: c2 - c1 + 1,
      value: value
    };
  });

  return map;
}

/**
 * Scans first 20 rows for header labels (Name, Subject, Class, Week, Term)
 * and reads the value in the adjacent cell to the right.
 */
function extractHeaderFields(data, mergeMap) {
  const result = {};
  const scanRows = Math.min(20, data.length);

  for (let r = 0; r < scanRows; r++) {
    const row = data[r];
    for (let c = 0; c < row.length; c++) {
      const cellStr = String(row[c] || '').trim();
      HEADER_LABELS.forEach(function(label) {
        if (cellStr === label && !result[label]) {
          // Value is in the cell to the right
          const valCell = row[c + 1];
          // If that cell is part of a merge, get the merge's top-left value
          const mergeVal = getMergeValue(r, c + 1, mergeMap, data);
          result[label] = String(mergeVal !== null ? mergeVal : (valCell || '')).trim();
        }
      });
    }
  }

  return result;
}

/**
 * Gets the effective value of a cell, resolving merged cell top-left values.
 */
function getCellValue(row, col, mergeMap, data) {
  if (col === null || col === undefined) return null;

  // Check if this cell is the top-left of a merge
  if (mergeMap[row] && mergeMap[row][col]) {
    return mergeMap[row][col].value;
  }

  // Otherwise return raw cell value
  if (row < data.length && col < data[row].length) {
    return data[row][col];
  }
  return null;
}

/**
 * Same as getCellValue but used in extractHeaderFields where
 * the cell is not necessarily the top-left (value may be in adjacent).
 */
function getMergeValue(row, col, mergeMap, data) {
  // Find if this cell belongs to any merge range as a non-top-left member
  // mergeMap only stores top-left cells, so we need to check if (row,col) is covered
  for (const r in mergeMap) {
    for (const c in mergeMap[r]) {
      const m = mergeMap[r][c];
      if (
        row >= m.topLeftRow &&
        row < m.topLeftRow + m.rowSpan &&
        col >= m.topLeftCol &&
        col < m.topLeftCol + m.colSpan
      ) {
        return m.value;
      }
    }
  }
  // Not in any merge — return raw
  if (row < data.length && col < data[row].length) {
    return data[row][col];
  }
  return null;
}

/**
 * Checks if a lesson block (given its row range) is a holiday entry.
 * A block is a holiday if any non-null cell in the block contains "holiday".
 */
function checkForHoliday(startRow, span, data, mergeMap) {
  for (let r = startRow; r < startRow + span && r < data.length; r++) {
    const row = data[r];
    for (let c = 0; c < row.length; c++) {
      const v = String(row[c] || '').trim().toLowerCase();
      if (v === 'holiday' || v.startsWith('holiday')) return true;
    }
  }
  return false;
}

/**
 * Extracts the text from a holiday block (whichever non-empty cell describes it).
 */
function extractHolidayText(startRow, span, data, mergeMap) {
  for (let r = startRow; r < startRow + span && r < data.length; r++) {
    const row = data[r];
    for (let c = 0; c < row.length; c++) {
      const v = String(row[c] || '').trim();
      if (v !== '') return v;
    }
  }
  return 'Holiday';
}


/** Cleans a cell value: converts to string, trims whitespace. Returns null if empty. */
function clean(val) {
  if (val === null || val === undefined) return null;
  const s = String(val).trim();
  return s === '' ? null : s;
}


// ════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════

function convertToGoogleSheets(fileId, fileName) {
  const xlsxFile  = DriveApp.getFileById(fileId);
  const parentId  = xlsxFile.getParents().next().getId();
  const blob      = xlsxFile.getBlob();

  // Use the Drive API v3 multipart upload via UrlFetchApp — no Advanced Service needed.
  // Step 1: upload the file bytes with ?convert implied by setting mimeType to Google Sheets.
  const boundary  = 'WLP_BOUNDARY_' + Utilities.getUuid().replace(/-/g, '');
  const metadata  = JSON.stringify({
    name:     `TEMP_CONVERT_${fileName}`,
    mimeType: 'application/vnd.google-apps.spreadsheet',
    parents:  [parentId]
  });

  // Build a multipart/related body manually
  const metaPart = '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    metadata + '\r\n';
  const dataPart = '--' + boundary + '\r\n' +
    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n';
  const closing   = '\r\n--' + boundary + '--';

  const bodyBytes = Utilities.newBlob(metaPart).getBytes()
    .concat(Utilities.newBlob(dataPart).getBytes())
    .concat(blob.getBytes())
    .concat(Utilities.newBlob(closing).getBytes());

  const token    = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
    {
      method:  'POST',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type':  'multipart/related; boundary=' + boundary
      },
      payload:              Utilities.newBlob(bodyBytes).getBytes(),
      muteHttpExceptions:   true
    }
  );

  if (response.getResponseCode() !== 200) {
    throw new Error(
      `Drive upload failed (${response.getResponseCode()}): ${response.getContentText().substring(0, 300)}`
    );
  }

  const result = JSON.parse(response.getContentText());
  return result.id;
}

function stripExtension(fileName) {
  return fileName.replace(/\.[^/.]+$/, '');
}

function describeNamingViolation(baseName) {
  const parts = baseName.split('_');
  if (parts.length !== 6) {
    return `Expected 6 underscore-separated segments, found ${parts.length}. ` +
           `Expected: [Class]_WLP_[Subject]_T[N]_AY25-26_VIS`;
  }
  if (parts[1] !== 'WLP') {
    return `Segment 2 should be "WLP", found "${parts[1]}"`;
  }
  if (!/^T[1-4]$/.test(parts[3])) {
    return `Segment 4 should be T1–T4, found "${parts[3]}"`;
  }
  if (parts[4] !== 'AY25-26') {
    return `Segment 5 should be "AY25-26", found "${parts[4]}"`;
  }
  if (parts[5] !== 'VIS') {
    return `Segment 6 should be "VIS", found "${parts[5]}"`;
  }
  return 'Does not match expected pattern [Class]_WLP_[Subject]_T[N]_AY25-26_VIS';
}

function writeJsonToDrive(folder, filename, data) {
  const jsonString = JSON.stringify(data, null, 2);
  const blob       = Utilities.newBlob(jsonString, 'application/json', filename);
  const existing   = folder.getFilesByName(filename);
  while (existing.hasNext()) { existing.next().setTrashed(true); }
  const newFile = folder.createFile(blob);
  Logger.log(`Written: ${filename} (ID: ${newFile.getId()})`);
  return newFile;
}

/**
 * Fetches bottom-border presence for each row in a given column,
 * within a specified row range. Uses the Sheets Advanced Service.
 *
 * Returns a Set of 0-indexed row numbers that have a bottom border
 * in the target column.
 *
 * @param {string} spreadsheetId
 * @param {string} sheetName
 * @param {number} col  0-indexed column
 * @param {number} startRow  0-indexed
 * @param {number} endRow    0-indexed (inclusive)
 */
function fetchBottomBorderedRows(spreadsheetId, sheetName, col, startRow, endRow) {
  const borderedRows = new Set();

  // A1 notation for just the anchor column, full data range
  const colLetter = columnToLetter(col + 1);
  const range = `'${sheetName}'!${colLetter}${startRow + 1}:${colLetter}${endRow + 1}`;

  const response = Sheets.Spreadsheets.get(spreadsheetId, {
    ranges: [range],
    fields: 'sheets(data(rowData(values(userEnteredFormat/borders))))'
  });

  const rowData = response.sheets[0].data[0].rowData || [];

  rowData.forEach(function(row, i) {
    const cell = row.values && row.values[0];
    if (!cell) return;
    const borders = cell.userEnteredFormat && cell.userEnteredFormat.borders;
    if (borders && borders.bottom && borders.bottom.style &&
        borders.bottom.style !== 'NONE') {
      borderedRows.add(startRow + i);
    }
  });

  return borderedRows;
}

function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/**
 * Fetches lastModifyingUser display name and email from the Drive Advanced Service.
 * Falls back gracefully if the field is unavailable.
 */
function getDriveFileMeta(fileId) {
  try {
    const file = Drive.Files.get(fileId, { fields: 'lastModifyingUser' });
    const user = file.lastModifyingUser;
    return {
      lastModifyingUser: user
        ? { displayName: user.displayName || null, email: user.emailAddress || null }
        : null
    };
  } catch (e) {
    Logger.log(`Could not fetch Drive meta for ${fileId}: ${e.message}`);
    return { lastModifyingUser: null };
  }
}

// ════════════════════════════════════════════════════════════
// OPTIONAL: Run individual stages for testing
// ════════════════════════════════════════════════════════════
function testTreeOnly() {
  const root = DriveApp.getFolderById(WLP_FOLDER_ID);
  writeJsonToDrive(root, 'WLP_FileTree.json', buildFileTree(root));
  Logger.log('Tree done.');
}

function testIndexOnly() {
  const root = DriveApp.getFolderById(WLP_FOLDER_ID);
  const idx  = buildFileIndex(root);
  writeJsonToDrive(root, 'WLP_FileIndex.json', idx);
  Logger.log(`Index done. ${idx.length} leaf folders.`);
}

function testSheetDataOnly() {
  const root  = DriveApp.getFolderById(WLP_FOLDER_ID);
  const idx   = buildFileIndex(root);
  const data  = extractAllSheetData(idx);
  writeJsonToDrive(root, 'WLP_SheetData.json', data);
  Logger.log(`Sheet data done. ${data.length} files.`);
}
