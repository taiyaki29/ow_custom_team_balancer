/**
 * Overwatch カスタム用チーム作成スクリプト
 * ランク・ロールバランス分割 → タブ出力
 */

const CONFIG = {
  SOURCE_SHEET: 'フォームの回答 1',
  OUTPUT_SHEET: 'バランス出力',
  OUTPUT_SHEET_PREFIX: 'バランス出力',
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

const ROLE_LABELS = {
  tank: 'タンク',
  dps: 'DPS',
  support: 'サポート',
};

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
    .addItem('10人でチーム作成', 'createBalancedTsvSheet10')
    .addItem('12人でチーム作成', 'createBalancedTsvSheet12')
    .addSeparator()
    .addItem('TSVをDriveに保存', 'exportTsvToDrive')
    .addToUi();
}

function createBalancedTsvSheet10() {
  createBalancedTsvSheet_(10);
}

function createBalancedTsvSheet12() {
  createBalancedTsvSheet_(12);
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

  const lobbyRoleCounts = getLobbyRoleCounts_(lobbySize);
  const teamRoleSlots = getTeamRoleSlots_(halfSize);
  const balancedLobbies = [];
  const roleFailedLobbies = [];

  fullLobbies.forEach(lobby => {
    if (!canFormLobby_(lobby, lobbyRoleCounts)) {
      roleFailedLobbies.push(lobby);
      return;
    }
    if (!assignRolesInLobby_(lobby, lobbyRoleCounts)) {
      roleFailedLobbies.push(lobby);
      return;
    }
    const split = bestTeamSplitWithRoles_(lobby, teamRoleSlots);
    if (!split) {
      roleFailedLobbies.push(lobby);
      return;
    }
    balancedLobbies.push({
      players: lobby,
      teamA: split.teamA,
      teamB: split.teamB,
      matchGap: split.cost,
      sumA: split.sumA,
      sumB: split.sumB,
    });
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  deleteOutputSheets_(ss, CONFIG.OUTPUT_SHEET_PREFIX);
  writeBalancedLobbies_(ss, balancedLobbies, partialLobbies.concat(roleFailedLobbies), lobbySize);

  const lobbySpread =
    fullLobbies.length > 1 ? computeLobbySpread_(fullLobbies) : 0;
  const matchGaps = balancedLobbies.map(l => l.matchGap).join(', ');
  const summaryLines = balancedLobbies.map((lobby, i) => {
    return `ロビー${i + 1}: A=${round1_(lobby.sumA)} B=${round1_(lobby.sumB)} 差=${round1_(lobby.matchGap)}`;
  });

  let message =
    `「${CONFIG.OUTPUT_SHEET}」タブに ${balancedLobbies.length} 試合分を出力しました（合計 ${players.length}人）\n` +
    `ロビー間スプレッド: ${lobbySpread}（小さいほど均等）\n` +
    `試合ごとのランク差: ${matchGaps || '—'}\n\n` +
    summaryLines.join('\n') +
    '\n\nもう一度押すと別の組み合わせになります';

  const extraLobbies = partialLobbies.concat(roleFailedLobbies);
  if (extraLobbies.length > 0) {
    const extraCount = extraLobbies.reduce((n, l) => n + l.length, 0);
    message += `\n\n※ ${extraCount}人は人数不足またはロール割当不可のため「余り」に出力しています`;
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
  const roleScores = {
    tank: roleScoreFromRank_(row[1]),
    dps: roleScoreFromRank_(row[2]),
    support: roleScoreFromRank_(row[3]),
  };
  return {
    row: row,
    score: playerScoreFromRow_(roleScores),
    roleScores: roleScores,
    prefs: {
      tank: normalizePref_(row[4]),
      dps: normalizePref_(row[5]),
      support: normalizePref_(row[6]),
    },
    assignedRole: null,
  };
}

function normalizePref_(pref) {
  if (pref === '×') return '×';
  if (pref === '○') return '○';
  return '○';
}

function roleScoreFromRank_(rankStr) {
  const points = rankToPoints_(rankStr);
  return points > 0 ? points : CONFIG.DEFAULT_RANK_POINTS;
}

function playerScoreFromRow_(roleScoresOrRow) {
  let tank;
  let dps;
  let support;
  if (roleScoresOrRow.tank !== undefined) {
    tank = roleScoresOrRow.tank;
    dps = roleScoresOrRow.dps;
    support = roleScoresOrRow.support;
  } else {
    tank = rankToPoints_(roleScoresOrRow[1]);
    dps = rankToPoints_(roleScoresOrRow[2]);
    support = rankToPoints_(roleScoresOrRow[3]);
  }
  const values = [tank, dps, support].filter(p => p > 0);
  if (values.length === 0) return CONFIG.DEFAULT_RANK_POINTS;
  return values.reduce((a, b) => a + b, 0) / values.length;
}

function canPlayRole_(player, role) {
  return player.prefs[role] !== '×';
}

function effectiveScore_(player) {
  if (player.assignedRole) return player.roleScores[player.assignedRole];
  return player.score;
}

function getLobbyRoleCounts_(lobbySize) {
  const half = lobbySize / 2;
  const team = getTeamRoleSlots_(half);
  return {
    tank: team.tank * 2,
    dps: team.dps * 2,
    support: team.support * 2,
  };
}

function getTeamRoleSlots_(halfSize) {
  if (halfSize === 5) return { tank: 1, dps: 2, support: 2 };
  if (halfSize === 6) return { tank: 2, dps: 2, support: 2 };
  throw new Error(`未対応のチーム人数: ${halfSize}`);
}

function canFormLobby_(players, counts) {
  return ['tank', 'dps', 'support'].every(role => {
    const eligible = players.filter(p => canPlayRole_(p, role)).length;
    return eligible >= counts[role];
  });
}

function assignRolesInLobby_(players, counts) {
  for (let attempt = 0; attempt < 30; attempt++) {
    shuffleArray_(players);
    players.forEach(p => {
      p.assignedRole = null;
    });
    const ties = findRoleAssignments_(players, counts);
    if (ties.length === 0) continue;

    const pick = ties[Math.floor(Math.random() * ties.length)];
    players.forEach((p, i) => {
      p.assignedRole = pick[i];
    });
    return true;
  }
  return false;
}

function findRoleAssignments_(players, counts) {
  const roles = ['tank', 'dps', 'support'];
  const remaining = { tank: counts.tank, dps: counts.dps, support: counts.support };
  const eps = CONFIG.TIE_EPSILON;
  let minCost = Infinity;
  const ties = [];

  function assignmentCost_() {
    let sum = 0;
    let prefBonus = 0;
    players.forEach(p => {
      sum += p.roleScores[p.assignedRole];
      if (p.prefs[p.assignedRole] === '○') prefBonus += 0.5;
    });
    return -(sum + prefBonus);
  }

  function search(i) {
    if (i === players.length) {
      const cost = assignmentCost_();
      const snapshot = players.map(p => p.assignedRole);
      if (cost < minCost - eps) {
        minCost = cost;
        ties.length = 0;
        ties.push(snapshot);
      } else if (cost <= minCost + eps) {
        if (cost < minCost) minCost = cost;
        ties.push(snapshot);
      }
      return;
    }

    const player = players[i];
    for (let r = 0; r < roles.length; r++) {
      const role = roles[r];
      if (remaining[role] <= 0 || !canPlayRole_(player, role)) continue;
      player.assignedRole = role;
      remaining[role]--;
      search(i + 1);
      remaining[role]++;
      player.assignedRole = null;
    }
  }

  search(0);
  return ties;
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

function bestTeamSplitWithRoles_(players, teamSlots) {
  const byRole = { tank: [], dps: [], support: [] };
  players.forEach(p => {
    if (!p.assignedRole || !byRole[p.assignedRole]) return;
    byRole[p.assignedRole].push(p);
  });

  const needed = {
    tank: teamSlots.tank,
    dps: teamSlots.dps,
    support: teamSlots.support,
  };

  if (
    byRole.tank.length !== needed.tank * 2 ||
    byRole.dps.length !== needed.dps * 2 ||
    byRole.support.length !== needed.support * 2
  ) {
    return null;
  }

  const eps = CONFIG.TIE_EPSILON;
  let minCost = Infinity;
  const ties = [];

  function evaluateSplit_(teamA) {
    const teamASet = new Set(teamA);
    const teamB = [];
    let sumA = 0;
    let sumB = 0;
    players.forEach(p => {
      if (teamASet.has(p)) {
        sumA += effectiveScore_(p);
      } else {
        teamB.push(p);
        sumB += effectiveScore_(p);
      }
    });
    const cost = Math.abs(sumA - sumB);
    const result = {
      teamA: teamA.slice(),
      teamB: teamB,
      cost: cost,
      sumA: sumA,
      sumB: sumB,
    };
    if (cost < minCost - eps) {
      minCost = cost;
      ties.length = 0;
      ties.push(result);
    } else if (cost <= minCost + eps) {
      if (cost < minCost) minCost = cost;
      ties.push(result);
    }
  }

  forEachCombination_(byRole.tank, needed.tank, tankPick => {
    forEachCombination_(byRole.dps, needed.dps, dpsPick => {
      forEachCombination_(byRole.support, needed.support, supportPick => {
        evaluateSplit_(tankPick.concat(dpsPick, supportPick));
      });
    });
  });

  if (ties.length === 0) return null;
  return ties[Math.floor(Math.random() * ties.length)];
}

function forEachCombination_(arr, k, fn) {
  if (k === 0) {
    fn([]);
    return;
  }
  if (k > arr.length) return;

  const indices = [];
  function search(start, depth) {
    if (depth === k) {
      fn(indices.map(i => arr[i]));
      return;
    }
    for (let i = start; i <= arr.length - (k - depth); i++) {
      indices[depth] = i;
      search(i + 1, depth + 1);
    }
  }
  search(0, 0);
}

const ROLE_SORT_ORDER = { tank: 0, dps: 1, support: 2 };

function sortPlayersByRole_(players) {
  return players.slice().sort((a, b) => {
    const roleCmp =
      (ROLE_SORT_ORDER[a.assignedRole] ?? 9) - (ROLE_SORT_ORDER[b.assignedRole] ?? 9);
    if (roleCmp !== 0) return roleCmp;
    return String(a.row[0]).localeCompare(String(b.row[0]), 'ja');
  });
}

function getRoleRowSlots_(teamRoleSlots) {
  const slots = [];
  for (let i = 0; i < teamRoleSlots.tank; i++) slots.push({ role: 'tank', index: i });
  for (let i = 0; i < teamRoleSlots.dps; i++) slots.push({ role: 'dps', index: i });
  for (let i = 0; i < teamRoleSlots.support; i++) slots.push({ role: 'support', index: i });
  return slots;
}

function roleSlotLabel_(slot) {
  return ROLE_LABELS[slot.role];
}

function groupTeamByRole_(team) {
  const grouped = { tank: [], dps: [], support: [] };
  sortPlayersByRole_(team).forEach(p => {
    if (grouped[p.assignedRole]) grouped[p.assignedRole].push(p);
  });
  return grouped;
}

function getPlayerAtRoleSlot_(team, slot) {
  const grouped = groupTeamByRole_(team);
  const list = grouped[slot.role] || [];
  return list[slot.index] || null;
}

const ROLE_RANK_COL = { tank: 1, dps: 2, support: 3 };

function assignedRoleRank_(player) {
  if (!player || !player.assignedRole) return '';
  const col = ROLE_RANK_COL[player.assignedRole];
  return String(player.row[col] || '').trim();
}

function formatPlayerCell_(player) {
  if (!player) return '';
  const name = String(player.row[0]);
  const rank = assignedRoleRank_(player);
  if (rank) return `${name} (${rank})`;
  return name;
}

function countTeams_(balancedLobbies) {
  return balancedLobbies.length * 2;
}

function getTeamColumnHeaders_(teamCount) {
  const headers = [];
  for (let t = 1; t <= teamCount; t++) {
    headers.push(`チーム${t}`);
  }
  return headers;
}

function buildBalancedGrid_(balancedLobbies, lobbySize) {
  const teamRoleSlots = getTeamRoleSlots_(lobbySize / 2);
  const roleSlots = getRoleRowSlots_(teamRoleSlots);
  const teamCount = countTeams_(balancedLobbies);
  const headers = ['役割'].concat(getTeamColumnHeaders_(teamCount));

  const gridRows = roleSlots.map(slot => {
    const row = [roleSlotLabel_(slot)];
    balancedLobbies.forEach(lobby => {
      row.push(formatPlayerCell_(getPlayerAtRoleSlot_(lobby.teamA, slot)));
      row.push(formatPlayerCell_(getPlayerAtRoleSlot_(lobby.teamB, slot)));
    });
    return row;
  });

  return { headers: headers, rows: gridRows, teamCount: teamCount, matchCount: balancedLobbies.length };
}

function writeBalancedLobbies_(ss, balancedLobbies, partialLobbies, lobbySize) {
  const sheet = ss.insertSheet(CONFIG.OUTPUT_SHEET);
  const grid = buildBalancedGrid_(balancedLobbies, lobbySize);
  const colCount = grid.headers.length;
  const teamCount = grid.teamCount;
  const matchCount = grid.matchCount;

  const lobbySpread =
    balancedLobbies.length > 1
      ? computeLobbySpread_(balancedLobbies.map(l => l.players))
      : 0;
  const totalPlayers =
    balancedLobbies.reduce((n, l) => n + l.players.length, 0) +
    partialLobbies.reduce((n, l) => n + l.length, 0);

  const summaryRow = [
    `${lobbySize}人ロビー × ${matchCount}試合`,
    `${teamCount}チーム`,
    `合計 ${totalPlayers}人`,
    `ロビー間スプレッド: ${lobbySpread}`,
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
  ];
  while (summaryRow.length < colCount) summaryRow.push('');
  sheet.getRange(1, 1, 1, colCount).setValues([summaryRow.slice(0, colCount)]);

  sheet.getRange(2, 1, 1, colCount).setValues([grid.headers]);

  if (grid.rows.length > 0) {
    sheet.getRange(3, 1, grid.rows.length, colCount).setValues(grid.rows);
  }

  const extraNames = [];
  partialLobbies.forEach(lobby => {
    lobby.forEach(p => extraNames.push(String(p.row[0])));
  });
  let lastRow = 2 + grid.rows.length;
  if (extraNames.length > 0) {
    const extraRow = 3 + grid.rows.length + 1;
    sheet.getRange(extraRow, 1, 1, 2).setValues([['余り', extraNames.join(', ')]]);
    lastRow = extraRow;
  }

  formatBalancedOutputSheet_(sheet, colCount, lastRow, grid.rows.length);
}

function formatBalancedOutputSheet_(sheet, colCount, lastRow, dataRowCount) {
  if (colCount < 1) return;

  sheet.autoResizeColumns(1, colCount);

  for (let c = 1; c <= colCount; c++) {
    const min = c === 1 ? 72 : 150;
    const max = c === 1 ? 96 : 260;
    const w = Math.min(max, Math.max(min, sheet.getColumnWidth(c)));
    sheet.setColumnWidth(c, w);
  }

  sheet.getRange(1, 1, lastRow, colCount).setVerticalAlignment('middle');

  const header = sheet.getRange(2, 1, 1, colCount);
  header.setFontWeight('bold');
  header.setBackground('#f3f3f3');
  header.setHorizontalAlignment('center');

  sheet.getRange(1, 1, 1, colCount).setFontSize(9);
  sheet.getRange(1, 1, 1, 1).setFontWeight('bold');

  if (dataRowCount > 0) {
    const body = sheet.getRange(3, 1, dataRowCount, colCount);
    body.setVerticalAlignment('top');
    sheet.getRange(3, 1, dataRowCount, 1).setHorizontalAlignment('center');
    if (colCount > 1) {
      sheet.getRange(3, 2, dataRowCount, colCount - 1).setWrap(true);
    }
  }

  sheet.setFrozenRows(2);
}

function round1_(n) {
  return Math.round(n * 10) / 10;
}

function exportTsvToDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('先に「10人でチーム作成」または「12人でチーム作成」を実行してください');
    return;
  }

  const dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmm');
  const data = sheet.getDataRange().getValues();
  const tsv = data
    .map(row => row.map(cellToTsvCell_).join('\t'))
    .join('\n');
  const fileName = `teams_${dateStr}.tsv`;
  DriveApp.createFile(fileName, tsv, MimeType.PLAIN_TEXT);

  SpreadsheetApp.getUi().alert(
    `TSVを保存しました:\n\n${fileName}\n\nDriveのマイドライブを確認してください`
  );
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
