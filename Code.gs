/**
 * 麻雀点数記録アプリ - Google Apps Script バックエンド (15章準拠)
 * Backend Version: 2.1 (対応アプリ: v2.1.0以降 / 準拠する要件定義書・設計書バージョン: 2.1)
 *
 * v2.1 変更点(F-35 会機能対応):
 *   - 新規シート「会」を追加(1行=1会)。upsertSessions アクションでupsert。
 *   - 「対局一覧」シートに「会」列を追加(所属する会名。未所属は空欄)。
 *     不足していれば ensureSheet_ が右端に自動追記(既存シートも無改造でアップグレード可)。
 *   - GET ?action=list のレスポンスに sessions:[Session](全件・削除済み除く)を追加。
 *   - 会の対局数・参加者は _raw から都度導出して書き込む(アプリ側では持たない)。
 *
 * v2.0 変更点(F-34 場代精算対応):
 *   - 「対局一覧」シートに「場代」列を追加(未精算なら空欄、精算済みなら「3,000円/均等割り」のような表記)
 *   - 「対局結果明細」シートに「場代負担」「最終支払円」列を追加
 *   - ensureSheet_ 実行時にヘッダー行を検査し、既存シートに不足している列があれば
 *     右端に自動追記するようにした(既存データ行・既存列順は一切変更しない。MUST)。
 *     これにより v1.x で運用中のシートも、コード更新だけで新列が追加され、
 *     手動でのシート編集は不要。
 *   - upsertGames_ の行組み立てを「ヘッダー名→値」のマップ方式に変更し、
 *     シートの実際の列順(新規シートは定義順、移行済み旧シートは追記された列が右端)
 *     のどちらでも正しい列に値が入るようにした。
 *
 * v1.3 変更点: プレイヤーの削除(F-01)が同期されない設計上の抜けを修正。
 *   GameRecordの deleted:boolean と同じ論理削除方式をPlayerにも適用し、
 *   「プレイヤー」シートの状態列に「削除済み」を追加。以後、削除操作も
 *   upsertPlayers 経由で他端末に反映される。
 * v1.2 変更点(v1.4対応):
 *   - action=list のレスポンスに全プレイヤー配列を含めるように変更(複数端末間のプレイヤー同期。15.2/15.4)
 *   - 「プレイヤー」シートに「更新日時」列を追加し、last-write-winsの判定に使用
 *   - 対局一覧・対局結果明細・局履歴のプレイヤー名解決を「プレイヤー」シートを正とする方式に変更
 *     (旧バージョンでは座席情報のみに依存しておりズレの可能性がありました)
 * v1.1 変更点: playerNameOf_() が座席の name フィールドからプレイヤー名を正しく解決するよう修正
 *   (旧バージョンでは playerId をそのまま返してしまい、シート上に ID がそのまま表示される不具合がありました)
 *
 * ─── 既存デプロイをアップグレードする場合(新規デプロイなら読み飛ばしてOK) ───
 * a. このファイルの内容をApps Scriptエディタの Code.gs に上書き貼り付け(既存の
 *    スプレッドシートのセルデータは一切変更されない。コードとデータは別物)。
 * b. 「デプロイ」→「デプロイを管理」→ 既存デプロイの鉛筆(編集)アイコン →
 *    バージョン: 「新バージョン」を選択 → デプロイ。
 *    ※これをしないと保存しただけでは公開中のURLに新しいコードが反映されない。
 * c. 「プレイヤー」シートが v1.3 以前の4列(プレイヤーID/名前/登録日/状態)のまま
 *    残っている場合のみ、関数プルダウンで migratePlayersSheetV14 を選び1回実行する
 *    (5列レイアウトへ安全に移行。データは失われない。詳細は当該関数のコメント参照)。
 *    「対局一覧」「対局結果明細」「局履歴」「_raw」は v2.0 で不足列(場代関連)がある場合、
 *    次回のアクセス(doGet/doPost呼び出し)時に自動で右端に追記されるため、
 *    手動での移行作業は不要。
 *
 * ─── デプロイ手順(新規) ───
 * 1. 新しい Google スプレッドシートを作成する
 * 2. 「拡張機能」→「Apps Script」を開く
 * 3. このファイルの内容を Code.gs に貼り付ける
 * 4. 左メニューの「プロジェクトの設定」→「スクリプト プロパティ」で
 *    キー "TOKEN" に任意の合言葉(パスワード)を設定して保存する
 * 5. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」を選択
 *      - 実行ユーザー: 自分
 *      - アクセスできるユーザー: 全員
 * 6. デプロイ後に発行される URL (https://script.google.com/macros/s/xxxx/exec)
 *    をアプリの「その他 > スプレッドシート連携」に、手順4の合言葉とともに設定する
 * 7. 初回アクセス時に本スクリプトが自動でシート・ヘッダーを作成する
 *
 * シート1〜3は _raw から導出される閲覧用ビューであり、upsert のたびに
 * 該当対局IDの行を洗い替える。手動編集しても次回同期で上書きされる。
 */

const SHEET_NAMES = {
  LIST: '対局一覧',
  DETAIL: '対局結果明細',
  HISTORY: '局履歴',
  PLAYERS: 'プレイヤー',
  SESSIONS: '会',
  RAW: '_raw',
};

const HEADERS = {
  LIST: ['対局ID', '通しNo', '日時', 'モード', '所要分', '1位', '1位pt', '2位', '2位pt', '3位', '3位pt', '4位', '4位pt', 'フォルダー', '会', 'メモ', '場代', '更新日時', '削除'],
  DETAIL: ['対局ID', '日時', 'モード', 'プレイヤー', '順位', '最終持ち点', '素点', 'ウマ', 'オカ', '補正', '合計pt', '収支円', '焼き鳥', '場代負担', '最終支払円'],
  HISTORY: ['対局ID', '連番', '局', '本場', '種別', '内容', '点数移動', '時刻'],
  PLAYERS: ['プレイヤーID', '名前', '登録日', '更新日時', '状態'],
  SESSIONS: ['会ID', '会名', '開始日時', '終了日時', '対局数', '参加者', '場代', 'メモ', 'venue JSON', '更新日時', '削除'],
  RAW: ['対局ID', '更新日時', 'GameRecord JSON'],
};

const MODE_LABEL = { yonma: '4人麻雀', sanma: '3人麻雀', yonin_sanma: '4人3麻' };
const VENUE_METHOD_LABEL = { equal: '均等割り', weighted: '傾斜', topExempt: 'トップ免除' };

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function checkToken_(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('TOKEN');
  return expected && token === expected;
}

function getSs_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ---- v2.0: 既存シートのヘッダーに不足があれば右端に自動追記する(既存データ・既存列順は変更しない。MUST) ----
function migrateHeaderColumns_(sh, fullHeaders) {
  const lastCol = sh.getLastColumn();
  let headerRow = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const missing = fullHeaders.filter((h) => headerRow.indexOf(h) === -1);
  if (missing.length > 0) {
    sh.getRange(1, headerRow.length + 1, 1, missing.length).setValues([missing]);
    headerRow = headerRow.concat(missing);
  }
  return headerRow;
}

function ensureSheet_(name, headers) {
  const ss = getSs_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    if (name === SHEET_NAMES.LIST || name === SHEET_NAMES.DETAIL || name === SHEET_NAMES.HISTORY) {
      sh.getRange(2, 1).setNote('このシートは _raw から自動生成されるビューです。手動編集は次回同期で上書きされます。');
    }
  } else if (name === SHEET_NAMES.LIST || name === SHEET_NAMES.DETAIL) {
    // v2.0: 場代関連の新列が不足していれば右端に自動追記
    migrateHeaderColumns_(sh, headers);
  }
  return sh;
}

function ensureAllSheets_() {
  ensureSheet_(SHEET_NAMES.LIST, HEADERS.LIST);
  ensureSheet_(SHEET_NAMES.DETAIL, HEADERS.DETAIL);
  ensureSheet_(SHEET_NAMES.HISTORY, HEADERS.HISTORY);
  ensureSheet_(SHEET_NAMES.PLAYERS, HEADERS.PLAYERS);
  ensureSheet_(SHEET_NAMES.SESSIONS, HEADERS.SESSIONS);
  ensureSheet_(SHEET_NAMES.RAW, HEADERS.RAW);
}

// ---- v2.0: ヘッダー名→値のマップから、シートの実際の列順に合わせた行配列を組み立てる ----
function buildRowByHeader_(headerRow, valueMap) {
  return headerRow.map((h) => (Object.prototype.hasOwnProperty.call(valueMap, h) ? valueMap[h] : ''));
}

function fmtDate_(ts) {
  if (!ts) return '';
  return Utilities.formatDate(new Date(ts), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

function findRowByFirstCol_(sh, value) {
  const finder = sh.createTextFinder(String(value)).matchEntireCell(true);
  const range = finder.findNext();
  if (!range) return -1;
  if (range.getColumn() !== 1) return -1;
  return range.getRow();
}

function deleteRowsMatchingGameId_(sh, gameId) {
  const data = sh.getDataRange().getValues();
  for (let r = data.length - 1; r >= 1; r--) {
    if (data[r][0] === gameId) {
      sh.deleteRow(r + 1);
    }
  }
}

// ---- doGet: ping / list / get ----
function doGet(e) {
  ensureAllSheets_();
  const p = e.parameter || {};
  const action = p.action;
  if (action === 'ping') {
    if (!checkToken_(p.token)) return jsonOut_({ ok: false, error: 'unauthorized' });
    return jsonOut_({ ok: true, version: '1' });
  }
  if (action === 'list') {
    if (!checkToken_(p.token)) return jsonOut_({ ok: false, error: 'unauthorized' });
    const sh = ensureSheet_(SHEET_NAMES.RAW, HEADERS.RAW);
    const data = sh.getDataRange().getValues();
    const games = [];
    for (let r = 1; r < data.length; r++) {
      const id = data[r][0];
      if (!id) continue;
      let deleted = false;
      try {
        const rec = JSON.parse(data[r][2]);
        deleted = !!rec.deleted;
      } catch (err) {}
      games.push({ id, updatedAt: new Date(data[r][1]).getTime(), deleted });
    }
    const players = readAllPlayers_();
    const sessions = readAllSessions_().filter((s) => !s.deleted);
    return jsonOut_({ ok: true, games, players, sessions });
  }
  if (action === 'get') {
    if (!checkToken_(p.token)) return jsonOut_({ ok: false, error: 'unauthorized' });
    const ids = (p.ids || '').split(',').filter(Boolean);
    const sh = ensureSheet_(SHEET_NAMES.RAW, HEADERS.RAW);
    const data = sh.getDataRange().getValues();
    const games = [];
    for (let r = 1; r < data.length; r++) {
      if (ids.includes(data[r][0])) {
        try {
          games.push(JSON.parse(data[r][2]));
        } catch (err) {}
      }
    }
    return jsonOut_({ ok: true, games });
  }
  return jsonOut_({ ok: false, error: 'bad_request' });
}

// ---- doPost: upsertGames / upsertPlayers ----
function doPost(e) {
  ensureAllSheets_();
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'bad_request' });
  }
  if (!checkToken_(body.token)) return jsonOut_({ ok: false, error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'locked' });
  }
  try {
    if (body.action === 'upsertGames') {
      upsertGames_(body.games || []);
      return jsonOut_({ ok: true });
    }
    if (body.action === 'upsertPlayers') {
      upsertPlayers_(body.players || []);
      return jsonOut_({ ok: true });
    }
    if (body.action === 'upsertSessions') {
      upsertSessions_(body.sessions || []);
      return jsonOut_({ ok: true });
    }
    return jsonOut_({ ok: false, error: 'bad_request' });
  } finally {
    lock.releaseLock();
  }
}

function readAllPlayers_() {
  const sh = ensureSheet_(SHEET_NAMES.PLAYERS, HEADERS.PLAYERS);
  const data = sh.getDataRange().getValues();
  const players = [];
  for (let r = 1; r < data.length; r++) {
    const id = data[r][0];
    if (!id) continue;
    players.push({
      id: id,
      name: data[r][1],
      createdAt: data[r][2] ? Date.parse(String(data[r][2]).replace(/\//g, '-')) || null : null,
      updatedAt: data[r][3] ? Date.parse(String(data[r][3]).replace(/\//g, '-')) || null : null,
      archived: data[r][4] === '統合済み',
      deleted: data[r][4] === '削除済み',
    });
  }
  return players;
}

function loadPlayerNameMap_() {
  const map = {};
  readAllPlayers_().forEach((p) => {
    map[p.id] = p.name;
  });
  return map;
}

// ---- v2.1(F-35): 会(Session) ----
function loadSessionNameMap_() {
  const map = {};
  readAllSessions_().forEach((s) => {
    map[s.id] = s.name;
  });
  return map;
}

function readAllSessions_() {
  const sh = ensureSheet_(SHEET_NAMES.SESSIONS, HEADERS.SESSIONS);
  const data = sh.getDataRange().getValues();
  const sessions = [];
  for (let r = 1; r < data.length; r++) {
    const id = data[r][0];
    if (!id) continue;
    let venue;
    try {
      venue = data[r][8] ? JSON.parse(data[r][8]) : undefined;
    } catch (err) {
      venue = undefined;
    }
    sessions.push({
      id: id,
      name: data[r][1],
      startedAt: data[r][2] ? Date.parse(String(data[r][2]).replace(/\//g, '-')) || null : null,
      endedAt: data[r][3] ? Date.parse(String(data[r][3]).replace(/\//g, '-')) || null : undefined,
      memo: data[r][7] || undefined,
      venue: venue,
      updatedAt: data[r][9] ? Date.parse(String(data[r][9]).replace(/\//g, '-')) || null : null,
      deleted: data[r][10] === 'TRUE' || data[r][10] === true,
    });
  }
  return sessions;
}

// 対局数・参加者はGASが対局データ(_raw)から導出して書く(15.3)
function upsertSessions_(sessions) {
  const sh = ensureSheet_(SHEET_NAMES.SESSIONS, HEADERS.SESSIONS);
  const rawSh = ensureSheet_(SHEET_NAMES.RAW, HEADERS.RAW);
  const rawData = rawSh.getDataRange().getValues();
  const nameMap = loadPlayerNameMap_();
  sessions.forEach((s) => {
    const row = findRowByFirstCol_(sh, s.id);
    if (row > 0) {
      const existingUpdatedAt = Date.parse(String(sh.getRange(row, 10).getValue()).replace(/\//g, '-')) || 0;
      if (existingUpdatedAt > (s.updatedAt || 0)) return; // last-write-wins
    }
    // このセッションに属する対局を_rawから導出(対局数・参加者)
    let gameCount = 0;
    const playerIds = {};
    for (let r = 1; r < rawData.length; r++) {
      try {
        const rec = JSON.parse(rawData[r][2]);
        if (rec.sessionId === s.id && !rec.deleted) {
          gameCount++;
          (rec.seats || []).forEach((seat) => {
            playerIds[seat.playerId] = true;
          });
        }
      } catch (err) {
        /* skip malformed row */
      }
    }
    const participants = Object.keys(playerIds).map((pid) => nameMap[pid] || pid).join('・');
    const venueLabel = s.venue ? formatComma_(s.venue.totalYen) + '円/' + (VENUE_METHOD_LABEL[s.venue.method] || s.venue.method) : '';
    const rowData = [
      s.id,
      s.name,
      fmtDate_(s.startedAt),
      s.endedAt ? fmtDate_(s.endedAt) : '',
      gameCount,
      participants,
      venueLabel,
      s.memo || '',
      s.venue ? JSON.stringify(s.venue) : '',
      fmtDate_(s.updatedAt || Date.now()),
      s.deleted ? 'TRUE' : 'FALSE',
    ];
    if (row > 0) sh.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    else sh.appendRow(rowData);
  });
}

function formatComma_(n) {
  if (n == null || isNaN(n)) return '';
  return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function upsertGames_(games) {
  const rawSh = ensureSheet_(SHEET_NAMES.RAW, HEADERS.RAW);
  const listSh = ensureSheet_(SHEET_NAMES.LIST, HEADERS.LIST);
  const detailSh = ensureSheet_(SHEET_NAMES.DETAIL, HEADERS.DETAIL);
  const histSh = ensureSheet_(SHEET_NAMES.HISTORY, HEADERS.HISTORY);
  const nameMap = loadPlayerNameMap_(); // プレイヤーシートを正とする名前解決(v1.4)
  const sessionNameMap = loadSessionNameMap_(); // v2.1(F-35): 会名解決

  // v2.0: シートの実際のヘッダー行(新規は定義順、移行済み旧シートは追記列が右端)を取得
  const listHeaderRow = listSh.getRange(1, 1, 1, listSh.getLastColumn()).getValues()[0];
  const detailHeaderRow = detailSh.getRange(1, 1, 1, detailSh.getLastColumn()).getValues()[0];

  games.forEach((game) => {
    // ---- _raw upsert ----
    const row = findRowByFirstCol_(rawSh, game.id);
    const rawRow = [game.id, fmtDate_(game.updatedAt), JSON.stringify(game)];
    if (row > 0) rawSh.getRange(row, 1, 1, 3).setValues([rawRow]);
    else rawSh.appendRow(rawRow);

    // ---- 対局一覧・対局結果明細・局履歴は洗い替え ----
    deleteRowsMatchingGameId_(listSh, game.id);
    deleteRowsMatchingGameId_(detailSh, game.id);
    deleteRowsMatchingGameId_(histSh, game.id);

    if (game.deleted) return; // 論理削除: ビューには表示しない

    const ranked = (game.result || []).slice().sort((a, b) => a.rank - b.rank);
    const modeLabel = MODE_LABEL[game.mode] || game.mode;

    // v2.0(F-34): 場代精算済みなら「3,000円/均等割り」のような表記。未精算なら空欄。
    const venueLabel = game.venue
      ? formatComma_(game.venue.totalYen) + '円/' + (VENUE_METHOD_LABEL[game.venue.method] || game.venue.method)
      : '';
    const venueBySeat = {};
    if (game.venue && game.venue.settlements) {
      game.venue.settlements.forEach((s) => {
        venueBySeat[s.seat] = s;
      });
    }

    const listValueMap = {
      '対局ID': game.id,
      '通しNo': game.serialNo,
      '日時': fmtDate_(game.endedAt || game.startedAt),
      'モード': modeLabel,
      '所要分': game.durationMin || '',
      'フォルダー': game.folderId || '',
      '会': game.sessionId ? (sessionNameMap[game.sessionId] || '') : '',
      'メモ': game.memo || '',
      '場代': venueLabel,
      '更新日時': fmtDate_(game.updatedAt),
      '削除': game.deleted ? 'TRUE' : 'FALSE',
    };
    for (let i = 0; i < 4; i++) {
      const rankLabel = (i + 1) + '位';
      if (ranked[i]) {
        listValueMap[rankLabel] = playerNameOf_(game, ranked[i].playerId, nameMap);
        listValueMap[rankLabel + 'pt'] = ranked[i].totalPt;
      } else {
        listValueMap[rankLabel] = '';
        listValueMap[rankLabel + 'pt'] = '';
      }
    }
    listSh.appendRow(buildRowByHeader_(listHeaderRow, listValueMap));

    (game.result || []).forEach((r) => {
      const vs = venueBySeat[r.seat];
      const detailValueMap = {
        '対局ID': game.id,
        '日時': fmtDate_(game.endedAt || game.startedAt),
        'モード': modeLabel,
        'プレイヤー': playerNameOf_(game, r.playerId, nameMap),
        '順位': r.rank,
        '最終持ち点': r.finalPoints,
        '素点': r.rawPt,
        'ウマ': r.umaPt,
        'オカ': r.okaPt,
        '補正': r.bonusPt || 0,
        '合計pt': r.totalPt,
        '収支円': r.yen != null ? r.yen : '',
        '焼き鳥': r.yakitori ? 'TRUE' : 'FALSE',
        '場代負担': vs ? vs.burdenYen : '',
        '最終支払円': vs ? vs.finalYen : '',
      };
      detailSh.appendRow(buildRowByHeader_(detailHeaderRow, detailValueMap));
    });

    (game.events || []).forEach((ev, idx) => {
      histSh.appendRow([game.id, idx + 1, ev.kyoku || '', ev.honba != null ? ev.honba : '', ev.type, describeEvent_(game, ev, nameMap), pointsMoveDesc_(game, ev), fmtDate_(ev.at)]);
    });
  });
}

function playerNameOf_(game, playerId, nameMap) {
  if (nameMap && nameMap[playerId]) return nameMap[playerId];
  const seat = (game.seats || []).find((s) => s.playerId === playerId);
  if (seat && seat.name) return seat.name;
  return playerId || '';
}

function nameBySeat_(game, seatIdx, nameMap) {
  const s = (game.seats || [])[seatIdx];
  if (!s) return '座席' + seatIdx;
  return playerNameOf_(game, s.playerId, nameMap);
}

function describeEvent_(game, ev, nameMap) {
  switch (ev.type) {
    case 'RIICHI':
      return nameBySeat_(game, ev.seat, nameMap) + ' リーチ';
    case 'RIICHI_CANCEL':
      return nameBySeat_(game, ev.seat, nameMap) + ' リーチ取消';
    case 'DRAW':
      return ev.special ? '特殊流局' : '流局';
    case 'ADJUST':
      return nameBySeat_(game, ev.seat, nameMap) + ' 点数調整 ' + (ev.delta > 0 ? '+' : '') + ev.delta;
    case 'CHOMBO':
      return nameBySeat_(game, ev.seat, nameMap) + ' チョンボ';
    case 'WAREME_SET':
      return nameBySeat_(game, ev.seat, nameMap) + ' が割れ目に';
    case 'WIN': {
      const w = (ev.winners || [])[0] || {};
      const score = w.directPoints != null ? w.directPoints + '点' : (w.han || '') + '翻' + (w.fu || '') + '符';
      const honbaTxt = game.rules && game.rules.honbaValue && ev.honba ? '+' + game.rules.honbaValue * ev.honba : '';
      return ev.method === 'ron'
        ? nameBySeat_(game, w.seat, nameMap) + ' ロン←' + nameBySeat_(game, ev.loserSeat, nameMap) + ' ' + score + honbaTxt
        : nameBySeat_(game, w.seat, nameMap) + ' ツモ ' + score + honbaTxt;
    }
    default:
      return ev.type;
  }
}

function pointsMoveDesc_(game, ev) {
  return ''; // アプリ側で厳密な点数移動サマリを送る場合はここに反映
}

function upsertPlayers_(players) {
  const sh = ensureSheet_(SHEET_NAMES.PLAYERS, HEADERS.PLAYERS);
  players.forEach((p) => {
    const row = findRowByFirstCol_(sh, p.id);
    const status = p.deleted ? '削除済み' : p.archived ? '統合済み' : '有効';
    const rowData = [p.id, p.name, fmtDate_(p.createdAt), fmtDate_(p.updatedAt || Date.now()), status];
    if (row > 0) {
      // last-write-wins: シート側の更新日時が新しければ上書きしない
      const existingUpdatedAt = Date.parse(String(sh.getRange(row, 4).getValue()).replace(/\//g, '-')) || 0;
      if (existingUpdatedAt > (p.updatedAt || 0)) return;
      sh.getRange(row, 1, 1, 5).setValues([rowData]);
    } else {
      sh.appendRow(rowData);
    }
  });
}

/**
 * 【既存デプロイをv1.4以降にアップグレードする場合のみ・1回だけ手動実行】
 * v1.3以前の「プレイヤー」シート(プレイヤーID・名前・登録日・状態の4列)に
 * 「更新日時」列を追加し、v1.4以降の5列レイアウトへ安全に移行する。
 * 既存データは一切削除・上書きしない(既存の値はそのまま右にずれるだけで、
 * 新しい「更新日時」列には各行の登録日を暫定値として入れる。以後、
 * 各端末からの同期のたびに正しい更新日時へ自然に置き換わっていく)。
 *
 * 実行方法: Apps Scriptエディタ上部の関数選択プルダウンで
 *   "migratePlayersSheetV14" を選び、▷実行 ボタンを押すだけ。
 * 既に新形式(5列)のシートに対しては何もしない(安全に何度でも再実行可能)。
 * これから新規デプロイする場合は実行不要(自動生成が最初から5列で作られる)。
 */
function migratePlayersSheetV14() {
  const ss = getSs_();
  const sh = ss.getSheetByName(SHEET_NAMES.PLAYERS);
  if (!sh) {
    Logger.log('「プレイヤー」シートがまだ存在しません。初回アクセス時に自動作成されるため、移行の必要はありません。');
    return;
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headerRow[3] === '更新日時') {
    Logger.log('既に新形式(更新日時列あり)です。移行の必要はありません。');
    return;
  }
  sh.insertColumnBefore(4);
  sh.getRange(1, 4).setValue('更新日時');
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    for (let r = 2; r <= lastRow; r++) {
      const createdAtStr = sh.getRange(r, 3).getValue(); // 登録日を暫定の更新日時として使う
      sh.getRange(r, 4).setValue(createdAtStr || '');
    }
  }
  Logger.log('移行完了: 「プレイヤー」シートを5列レイアウト(更新日時列つき)にしました。既存データは失われていません。');
}

/**
 * 【本当にスプレッドシートのデータを全部消したい場合だけ・手動実行専用】
 * アプリの「その他 > 全データ削除」はこの端末のlocalStorageのみを消去し、
 * スプレッドシート側のデータには一切触れない(意図的な安全設計)。
 * 家族・友人で共有しているデータを誤って一括消去できてしまうと危険なため、
 * この操作はWeb API(doGet/doPost)からは呼び出せないようにしてあり、
 * シートの持ち主がApps Scriptエディタから手動実行した場合のみ動作する。
 *
 * 実行方法: 関数選択プルダウンで "wipeAllSpreadsheetDataDangerous" を選び実行。
 * 実行前に必ず「ファイル→版数履歴」等でバックアップを確認すること。
 * 5シートすべてのヘッダー行以外を削除する(ヘッダー行自体は残す)。
 */
function wipeAllSpreadsheetDataDangerous() {
  ensureAllSheets_();
  Object.values(SHEET_NAMES).forEach((name) => {
    const sh = getSs_().getSheetByName(name);
    if (!sh) return;
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      sh.deleteRows(2, lastRow - 1);
    }
  });
  Logger.log('全シートのデータ行を削除しました(ヘッダー行のみ残っています)。');
}
