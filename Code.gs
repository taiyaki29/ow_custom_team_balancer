/**
 * Overwatch カスタム用 TSV 変換スクリプト
 * ランダム分割 / ランクバランス分割 → タブ出力
 */

const CONFIG = {
  SOURCE_SHEET: 'フォームの回答 1',
  OUTPUT_SHEET_PREFIX: 'TSV出力_',
  BALANCED_OUTPUT_SHEET_PREFIX: 'バランス出力_',
  LOBBY_SPREAD_WEIGHT: 2,
  TIE_EPSILON: 2,
  DEFAULT_RANK_POINTS: 18,
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

const BALANCED_HEADERS = TSV_HEADERS.concat(['チーム', 'スコア']);

const PREFERENCE_MAP = {
  'いいえ': '×',
  'やってもいい': '○',
  '×': '×',
  '○': '○',
};

const RANK_POINTS = buildRankPointsMap_();

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('チーム作成')
    .addItem('10人でランダム分割', 'createTsvSheet10')
    .addItem('12人でランダム分割', 'createTsvSheet12')
    .addSeparator()
    .addItem('10人でバランス分割', 'createBalancedTsvSheet10')
    .addItem('12人でバランス分割', 'createBalancedTsvSheet12')
    .addSeparator()
    .addItem('TSVをDriveに保存（全タブ）', 'exportTsvToDrive')
    .addToUi();
}

function createTsvSheet10() {
  createTsvSheet_(10);
}

function createTsvSheet12() {
  createTsvSheet_(12);
}

function createBalancedTsvSheet10() {
  createBalancedTsvSheet_(10);
}

function createBalancedTsvSheet12() {
  createBalancedTsvSheet_(12);
}

function createTsvSheet_(teamSize) {
  const rows = loadPlayerRows_();
  if (!rows) return;

  const shuffled = shuffleRows_(rows);
  const chunks = chunkRows_(shuffled, teamSize);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  deleteOutputSheets_(ss, CONFIG.OUTPUT_SHEET_PREFIX);
  writeChunkSheets_(ss, chunks, CONFIG.OUTPUT_SHEET_PREFIX);

  const summary = chunks
    .map((chunk, i) => `${CONFIG.OUTPUT_SHEET_PREFIX}${i + 1}: ${chunk.length}人`)
    .join('\n');

  SpreadsheetApp.getUi().alert(
    `${teamSize}人ずつ・ランダムで ${chunks.length} タブを作成しました（合計 ${rows.length}人）\n\n${summary}\n\nもう一度押すと別の組み合わせになります`
  );
}

function createBalancedTsvSheet_(lobbySize) {
  const rows = loadPlayerRows_();
  if (!rows) return;

  const halfSize = lobbySize / 2;
  const players = rows.map(row => toPlayer_(row));
  const shuffled = shuffleArray_(players);
  const lobbyChunks = chunkArray_(shuffled, lobbySize);

  const fullLobbies = lobbyChunks.filter(lobby => lobby.length === lobbySize);
  const partialLobbies = lobbyChunks.filter(lobby => lobby.length !== lobbySize);

  optimizeLobbies_(fullLobbies, getSwapAttempts_(players.length));

  const balancedLobbies = fullLobbies.map(lobby => {
    const split = bestTeamSplit_(lobby, halfSize);
    return {
      players: lobby,
      teamA: split.teamA,
      teamB: split.teamB,
      matchGap: split.cost,
      sumA: split.sumA,
      sumB: split.sumB,
    };
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  deleteOutputSheets_(ss, CONFIG.BALANCED_OUTPUT_SHEET_PREFIX);
  writeBalancedLobbies_(ss, balancedLobbies, partialLobbies, lobbySize);

  const lobbySpread =
    fullLobbies.length > 1 ? computeLobbySpread_(fullLobbies) : 0;
  const matchGaps = balancedLobbies.map(l => l.matchGap).join(', ');
  const summaryLines = balancedLobbies.map((lobby, i) => {
    const name = `${CONFIG.BALANCED_OUTPUT_SHEET_PREFIX}${i + 1}`;
    return `${name}: A=${lobby.sumA} B=${lobby.sumB} 差=${lobby.matchGap}`;
  });

  let message =
    `${lobbySize}人ロビー・バランス分割: ${balancedLobbies.length} タブ（合計 ${players.length}人）\n` +
    `ロビー間スプレッド: ${lobbySpread}（小さいほど均等）\n` +
    `試合ごとのランク差: ${matchGaps || '—'}\n\n` +
    summaryLines.join('\n') +
    '\n\nもう一度押すと別の組み合わせになります';

  if (partialLobbies.length > 0) {
    const partialCount = partialLobbies.reduce((n, l) => n + l.length, 0);
    message += `\n\n※ ${partialCount}人は人数が足りないためバランス分割していません（別タブ）`;
  }

  SpreadsheetApp.getUi().alert(message);
}

function loadPlayerRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!src) {
    throw new Error(`回答シート「${CONFIG.SOURCE_SHEET}」が見つかりません`);
  }

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('回答データがありません');
    return null;
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
    return null;
  }

  return rows;
}

function toPlayer_(row) {
  return {
    row: row,
    score: playerScoreFromRow_(row),
  };
}

function playerScoreFromRow_(row) {
  const tank = rankToPoints_(row[1]);
  const dps = rankToPoints_(row[2]);
  const support = rankToPoints_(row[3]);
  const values = [tank, dps, support].filter(p => p > 0);
  if (values.length === 0) return CONFIG.DEFAULT_RANK_POINTS;
  return values.reduce((a, b) => a + b, 0) / values.length;
}

function rankToPoints_(rankStr) {
  if (rankStr === null || rankStr === undefined || rankStr === '') return 0;
  const key = normalizeRankKey_(String(rankStr));
  if (RANK_POINTS[key] !== undefined) return RANK_POINTS[key];

  const parsed = parseRankKey_(key);
  if (parsed) {
    const composite = `${parsed.tier}${parsed.div}`;
    if (RANK_POINTS[composite] !== undefined) return RANK_POINTS[composite];
  }

  return 0;
}

function normalizeRankKey_(rankStr) {
  return rankStr
    .trim()
    .toLowerCase()
    .replace(/[\s　_\-・]/g, '')
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xfee0));
}

function parseRankKey_(key) {
  const divMatch = key.match(/(\d)$/);
  if (!divMatch) return null;
  const div = parseInt(divMatch[1], 10);
  if (div < 1 || div > 5) return null;
  const tier = key.slice(0, -1);
  if (!tier) return null;
  return { tier: tier, div: div };
}

function buildRankPointsMap_() {
  const tiers = [
    ['bronze', 'ブロンズ'],
    ['silver', 'シルバー', 'silver'],
    ['gold', 'ゴールド'],
    ['platinum', 'プラチナ', 'plat', 'プラ'],
    ['diamond', 'ダイヤ', 'ダイヤモンド', 'dia'],
    ['master', 'マスター'],
    ['grandmaster', 'グランドマスター', 'グラマス', 'gm', 'グランドマスタ'],
    ['champion', 'チャンピオン', 'チャン'],
  ];

  const map = {};
  tiers.forEach((aliases, tierIdx) => {
    for (let div = 5; div >= 1; div--) {
      const points = tierIdx * 5 + (6 - div);
      aliases.forEach(alias => {
        const normalized = normalizeRankKey_(alias);
        map[`${normalized}${div}`] = points;
      });
    }
  });
  return map;
}

function getSwapAttempts_(playerCount) {
  if (playerCount <= 10) return 2000;
  if (playerCount <= 20) return 5000;
  if (playerCount <= 30) return 10000;
  if (playerCount <= 40) return 15000;
  return 20000;
}

function lobbySum_(lobby) {
  return lobby.reduce((sum, p) => sum + p.score, 0);
}

function computeLobbySpread_(lobbies) {
  const sums = lobbies.map(lobbySum_);
  return Math.round((Math.max.apply(null, sums) - Math.min.apply(null, sums)) * 10) / 10;
}

function optimizeLobbies_(lobbies, attempts) {
  if (lobbies.length < 2) return;

  let bestSpread = computeLobbySpread_(lobbies);

  for (let t = 0; t < attempts; t++) {
    const i = Math.floor(Math.random() * lobbies.length);
    let j = Math.floor(Math.random() * lobbies.length);
    if (j === i) j = (j + 1) % lobbies.length;

    const lobbyA = lobbies[i];
    const lobbyB = lobbies[j];
    if (lobbyA.length === 0 || lobbyB.length === 0) continue;

    const idxA = Math.floor(Math.random() * lobbyA.length);
    const idxB = Math.floor(Math.random() * lobbyB.length);

    const tmp = lobbyA[idxA];
    lobbyA[idxA] = lobbyB[idxB];
    lobbyB[idxB] = tmp;

    const newSpread = computeLobbySpread_(lobbies);
    if (newSpread < bestSpread) {
      bestSpread = newSpread;
    } else {
      lobbyB[idxB] = lobbyA[idxA];
      lobbyA[idxA] = tmp;
    }
  }
}

function bestTeamSplit_(players, halfSize) {
  const n = players.length;
  if (n !== halfSize * 2) {
    return {
      teamA: players.slice(),
      teamB: [],
      cost: 0,
      sumA: lobbySum_(players),
      sumB: 0,
    };
  }

  const eps = CONFIG.TIE_EPSILON;
  let minCost = Infinity;
  const ties = [];
  const chosen = [];

  function evaluate() {
    let sumA = 0;
    let sumB = 0;
    const teamA = [];
    const teamB = [];
    const inA = {};
    chosen.forEach(i => {
      inA[i] = true;
    });

    for (let i = 0; i < n; i++) {
      if (inA[i]) {
        sumA += players[i].score;
        teamA.push(players[i]);
      } else {
        sumB += players[i].score;
        teamB.push(players[i]);
      }
    }

    const cost = Math.abs(sumA - sumB);
    if (cost < minCost - eps) {
      minCost = cost;
      ties.length = 0;
      ties.push({ teamA: teamA, teamB: teamB, cost: cost, sumA: sumA, sumB: sumB });
    } else if (cost <= minCost + eps) {
      if (cost < minCost) minCost = cost;
      ties.push({ teamA: teamA, teamB: teamB, cost: cost, sumA: sumA, sumB: sumB });
    }
  }

  function search(start, depth) {
    if (depth === halfSize) {
      evaluate();
      return;
    }
    for (let i = start; i <= n - (halfSize - depth); i++) {
      chosen[depth] = i;
      search(i + 1, depth + 1);
    }
  }

  search(0, 0);

  if (ties.length === 0) {
    return {
      teamA: players.slice(0, halfSize),
      teamB: players.slice(halfSize),
      cost: 0,
      sumA: lobbySum_(players.slice(0, halfSize)),
      sumB: lobbySum_(players.slice(halfSize)),
    };
  }

  return ties[Math.floor(Math.random() * ties.length)];
}

function writeBalancedLobbies_(ss, balancedLobbies, partialLobbies, lobbySize) {
  balancedLobbies.forEach((lobby, index) => {
    const name = `${CONFIG.BALANCED_OUTPUT_SHEET_PREFIX}${index + 1}`;
    const sheet = ss.insertSheet(name);
    const lobbyTotal = Math.round(lobbySum_(lobby.players) * 10) / 10;

    sheet
      .getRange(1, 1, 1, 4)
      .setValues([
        [
          `ロビー ${lobbySize}人`,
          `A合計: ${round1_(lobby.sumA)}`,
          `B合計: ${round1_(lobby.sumB)}`,
          `試合差: ${round1_(lobby.matchGap)} / ロビー合計: ${lobbyTotal}`,
        ],
      ]);

    sheet.getRange(2, 1, 1, BALANCED_HEADERS.length).setValues([BALANCED_HEADERS]);

    const teamARows = lobby.teamA.map(p => playerToOutputRow_(p, 'A'));
    const teamBRows = lobby.teamB.map(p => playerToOutputRow_(p, 'B'));
    const dataRows = teamARows.concat(teamBRows);

    if (dataRows.length > 0) {
      sheet.getRange(3, 1, dataRows.length, BALANCED_HEADERS.length).setValues(dataRows);
    }
  });

  partialLobbies.forEach((lobby, index) => {
    const name = `${CONFIG.BALANCED_OUTPUT_SHEET_PREFIX}余り${index + 1}`;
    const sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, TSV_HEADERS.length).setValues([TSV_HEADERS]);
    const dataRows = lobby.map(p => p.row);
    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, TSV_HEADERS.length).setValues(dataRows);
    }
  });
}

function playerToOutputRow_(player, team) {
  return player.row.concat([team, round1_(player.score)]);
}

function round1_(n) {
  return Math.round(n * 10) / 10;
}

function exportTsvToDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getAllOutputSheets_(ss);
  if (sheets.length === 0) {
    SpreadsheetApp.getUi().alert(
      '先に「ランダム分割」または「バランス分割」ボタンで分割してください'
    );
    return;
  }

  const dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmm');
  const saved = [];

  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    const tsv = data
      .map(row => row.map(cellToTsvCell_).join('\t'))
      .join('\n');

    const suffix = sheet
      .getName()
      .replace(CONFIG.OUTPUT_SHEET_PREFIX, '')
      .replace(CONFIG.BALANCED_OUTPUT_SHEET_PREFIX, '');
    const prefix = sheet.getName().startsWith(CONFIG.BALANCED_OUTPUT_SHEET_PREFIX)
      ? 'balanced'
      : 'players';
    const fileName = `${prefix}_${dateStr}_tab${suffix}.tsv`;
    DriveApp.createFile(fileName, tsv, MimeType.PLAIN_TEXT);
    saved.push(fileName);
  });

  SpreadsheetApp.getUi().alert(
    `TSVを ${saved.length} 件保存しました:\n\n${saved.join('\n')}\n\nDriveのマイドライブを確認してください`
  );
}

function shuffleRows_(rows) {
  return shuffleArray_(rows.slice());
}

function shuffleArray_(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = arr[i];
    arr[i] = arr[j];
    arr[j] = tmp;
  }
  return arr;
}

function chunkRows_(rows, size) {
  return chunkArray_(rows, size);
}

function chunkArray_(arr, size) {
  const chunks = [];
  for (let i = 0; i < arr.length; i += size) {
    chunks.push(arr.slice(i, i + size));
  }
  return chunks;
}

function deleteOutputSheets_(ss, prefix) {
  ss.getSheets()
    .filter(s => s.getName().startsWith(prefix))
    .forEach(s => ss.deleteSheet(s));
}

function writeChunkSheets_(ss, chunks, prefix) {
  chunks.forEach((chunk, index) => {
    const name = `${prefix}${index + 1}`;
    const sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, TSV_HEADERS.length).setValues([TSV_HEADERS]);
    if (chunk.length > 0) {
      sheet.getRange(2, 1, chunk.length, TSV_HEADERS.length).setValues(chunk);
    }
  });
}

function getOutputSheets_(ss) {
  return getSheetsByPrefix_(ss, CONFIG.OUTPUT_SHEET_PREFIX);
}

function getAllOutputSheets_(ss) {
  return getSheetsByPrefix_(ss, CONFIG.OUTPUT_SHEET_PREFIX).concat(
    getSheetsByPrefix_(ss, CONFIG.BALANCED_OUTPUT_SHEET_PREFIX)
  );
}

function getSheetsByPrefix_(ss, prefix) {
  return ss
    .getSheets()
    .filter(s => s.getName().startsWith(prefix))
    .sort((a, b) => sheetSortKey_(a.getName(), prefix) - sheetSortKey_(b.getName(), prefix));
}

function sheetSortKey_(name, prefix) {
  const suffix = name.replace(prefix, '');
  const num = parseInt(suffix, 10);
  return isNaN(num) ? 9999 : num;
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
