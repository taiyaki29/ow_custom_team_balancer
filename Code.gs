/**
 * Overwatch カスタム用 TSV 変換スクリプト
 * 10人 / 12人 ボタンでランダム分割 → タブ出力
 */

const CONFIG = {
  SOURCE_SHEET: 'フォームの回答 1',
  OUTPUT_SHEET_PREFIX: 'TSV出力_',
};

const HEADER_MAP = {
  'OW上の名前': '名前',
  '名前': '名前',
  'タンク　ランク': 'タンクランク',
  'タンクランク': 'タンクランク',
  'タンク(ランク)': 'タンクランク',
  'dps　ランク': 'DPSランク',
  'DPSランク': 'DPSランク',
  'dps(ランク)': 'DPSランク',
  'サポート　ランク': 'サポートランク',
  'サポートランク': 'サポートランク',
  'サポ(ランク)': 'サポートランク',
  'タンク希望': 'タンク希望',
  'DPS希望': 'DPS希望',
  'サポート希望': 'サポート希望',
};

const TSV_HEADERS = [
  '名前',
  'タンクランク',
  'DPSランク',
  'サポートランク',
  'タンク希望',
  'DPS希望',
  'サポート希望',
];

const PREFERENCE_MAP = {
  'いいえ': '×',
  'やってもいい': '○',
  '×': '×',
  '○': '○',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TSV')
    .addItem('10人でランダム分割', 'createTsvSheet10')
    .addItem('12人でランダム分割', 'createTsvSheet12')
    .addSeparator()
    .addItem('TSVをDriveに保存（全タブ）', 'exportTsvToDrive')
    .addToUi();
}

/** 10人ずつランダム分割 */
function createTsvSheet10() {
  createTsvSheet_(10);
}

/** 12人ずつランダム分割 */
function createTsvSheet12() {
  createTsvSheet_(12);
}

function createTsvSheet_(teamSize) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!src) {
    throw new Error(`回答シート「${CONFIG.SOURCE_SHEET}」が見つかりません`);
  }

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('回答データがありません');
    return;
  }

  const allData = src.getRange(1, 1, lastRow, lastCol).getValues();
  const srcHeaders = allData[0].map(normalizeHeader);
  const colIndex = buildColumnIndex_(srcHeaders);

  const rows = allData
    .slice(1)
    .map(row => rowToTsv_(row, colIndex))
    .filter(row => row[0] !== '');

  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('有効な回答がありません');
    return;
  }

  const shuffled = shuffleRows_(rows);
  const chunks = chunkRows_(shuffled, teamSize);

  deleteOutputSheets_(ss);
  writeChunkSheets_(ss, chunks);

  const summary = chunks
    .map((chunk, i) => `${CONFIG.OUTPUT_SHEET_PREFIX}${i + 1}: ${chunk.length}人`)
    .join('\n');

  SpreadsheetApp.getUi().alert(
    `${teamSize}人ずつ・ランダムで ${chunks.length} タブを作成しました（合計 ${rows.length}人）\n\n${summary}\n\nもう一度押すと別の組み合わせになります`
  );
}

function exportTsvToDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getOutputSheets_(ss);
  if (sheets.length === 0) {
    SpreadsheetApp.getUi().alert('先に「10人」または「12人」ボタンで分割してください');
    return;
  }

  const dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmm');
  const saved = [];

  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    const tsv = data
      .map(row => row.map(cellToTsvCell_).join('\t'))
      .join('\n');

    const suffix = sheet.getName().replace(CONFIG.OUTPUT_SHEET_PREFIX, '');
    const fileName = `players_${dateStr}_tab${suffix}.tsv`;
    DriveApp.createFile(fileName, tsv, MimeType.PLAIN_TEXT);
    saved.push(fileName);
  });

  SpreadsheetApp.getUi().alert(
    `TSVを ${saved.length} 件保存しました:\n\n${saved.join('\n')}\n\nDriveのマイドライブを確認してください`
  );
}

/** Fisher–Yates シャッフル（実行のたびに別の並び） */
function shuffleRows_(rows) {
  const arr = rows.slice();
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = arr[i];
    arr[i] = arr[j];
    arr[j] = tmp;
  }
  return arr;
}

function chunkRows_(rows, size) {
  const chunks = [];
  for (let i = 0; i < rows.length; i += size) {
    chunks.push(rows.slice(i, i + size));
  }
  return chunks;
}

function deleteOutputSheets_(ss) {
  ss.getSheets()
    .filter(s => s.getName().startsWith(CONFIG.OUTPUT_SHEET_PREFIX))
    .forEach(s => ss.deleteSheet(s));
}

function writeChunkSheets_(ss, chunks) {
  chunks.forEach((chunk, index) => {
    const name = `${CONFIG.OUTPUT_SHEET_PREFIX}${index + 1}`;
    const sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, TSV_HEADERS.length).setValues([TSV_HEADERS]);
    if (chunk.length > 0) {
      // 第3引数 = 行数（chunk.length）、第4引数 = 列数
      sheet
        .getRange(2, 1, chunk.length, TSV_HEADERS.length)
        .setValues(chunk);
    }
  });
}

function getOutputSheets_(ss) {
  return ss
    .getSheets()
    .filter(s => s.getName().startsWith(CONFIG.OUTPUT_SHEET_PREFIX))
    .sort((a, b) => {
      const na = parseInt(a.getName().replace(CONFIG.OUTPUT_SHEET_PREFIX, ''), 10);
      const nb = parseInt(b.getName().replace(CONFIG.OUTPUT_SHEET_PREFIX, ''), 10);
      return na - nb;
    });
}

function normalizeHeader(header) {
  return String(header).trim().replace(/\s+/g, '　');
}

function buildColumnIndex_(srcHeaders) {
  const index = {};
  srcHeaders.forEach((header, i) => {
    const mapped = HEADER_MAP[header];
    if (mapped) index[mapped] = i;
  });
  if (index['名前'] === undefined) {
    throw new Error('「OW上の名前」列が見つかりません');
  }
  return index;
}

function rowToTsv_(row, colIndex) {
  return TSV_HEADERS.map(header => {
    const i = colIndex[header];
    if (i === undefined) return '';
    let value = row[i];
    if (value === null || value === undefined) return '';
    value = String(value).trim();
    if (value === '') return '';
    if (header.endsWith('希望')) {
      return PREFERENCE_MAP[value] || value;
    }
    return value;
  });
}

function cellToTsvCell_(cell) {
  if (cell === null || cell === undefined) return '';
  return String(cell).replace(/\t/g, ' ').replace(/\r?\n/g, ' ');
}