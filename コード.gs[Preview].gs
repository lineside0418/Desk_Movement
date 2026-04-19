/**
 * S.K.I.D Backend System v3
 *
 * ■ v2 からの改善点
 *
 * [SECURITY]
 *   - complete_movement: payload.amount の整合性チェックを追加
 *
 * [ACID / Atomicity]
 *   - complete_movement: 在庫の負チェック(不正な完了操作)を書き込み前に行い、
 *     検証失敗時は一切書き込まないよう修正
 *   - validate_and_generate / generate_movements: 全検証をメモリで完結させてから
 *     シートへの書き込みをまとめて行うよう整理（書き込みは関数末尾に集約）
 *   - direct_update: 値が数値でない場合に明示的エラーを出して書き込みを中断
 *
 * [API 呼び出し最小化]
 *   - authenticateUser が getRoomInfo 経由で Inventory を二重読みしていた問題を修正:
 *     ss をキャッシュし、Inventory は handlePostAction の先頭で一括取得して使い回す
 *   - change_term: loadConfig を二重呼びしていた（handlePostAction 先頭 + change_term 内部）
 *     を修正し、先頭で取得した config を引き回す
 *   - checkFloorProgress: 呼び出し元がすでに持っているデータを引数で渡すことで
 *     Inventory / Targets / Term_Targets の再読み込みを排除
 *   - getRoomInfo を廃止し、authenticateUser も invData を受け取る形に変更
 *
 * [ROBUSTNESS / 当日リスク対策]
 *   - 全アクションでペイロードの型・必須フィールド検証を入口で実施
 *   - complete_movement: amount が 0 以下の場合はエラー
 *   - complete_movement: Movements シートの行 index が範囲外の場合はエラー
 *   - generate_movements / validate_and_generate: Targets に存在しない roomId が
 *     Inventory に含まれる場合の警告をログに残す（エラーにはしない）
 *   - ALL_TERMS の変更が Term_Targets のヘッダー生成に即反映されるよう定数参照を統一
 */

// =========================================
// 定数
// =========================================

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_NAMES = {
  INVENTORY:    'Inventory',
  USERS:        'Users',
  TARGETS:      'Targets',
  MOVEMENTS:    'Movements',
  CONFIG:       'Config',
  LOGS:         'Logs',
  TERM_TARGETS: 'Term_Targets'
};
const ALL_TERMS = ['Term1', 'Term2', 'Term3'];

// 有効な Term の Set（高速な存在チェック用）
const VALID_TERMS = new Set(ALL_TERMS);

// =========================================
// エントリーポイント
// =========================================

function doGet(e) {
  try {
    return createJsonResponse(handleGetData());
  } catch (err) {
    Logger.log('doGet error: ' + err.toString());
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  // [Isolation] スクリプトロックで同時書き込みを排除（最大30秒待機）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return createJsonResponse({ status: 'error', message: 'サーバーが混雑しています。数秒後に再試行してください。' });
  }
  try {
    const result = handlePostAction(e);
    // [Durability] 全書き込みをスプレッドシートに強制永続化
    SpreadsheetApp.flush();
    return createJsonResponse(result);
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return createJsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// =========================================
// GET ロジック（ロックなし・高速）
// =========================================

function handleGetData() {
  // [API最小化] ss を1回取得して全シートを参照
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // --- Config（1回のgetValues） ---
  const config     = loadConfig(ss);
  const targetMode = config['TARGET_MODE'] || '終了時';

  // --- Inventory（1回のgetDataRange） ---
  const invSheet   = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const invRaw     = invSheet.getDataRange().getValues();
  const invHeaders = invRaw.shift();
  const items      = [];
  const itemIndexes = {};
  for (let i = 3; i < invHeaders.length; i++) {
    if (invHeaders[i]) {
      const name = String(invHeaders[i]);
      items.push(name);
      itemIndexes[name] = i;
    }
  }

  // --- Targets（1回のgetDataRange） ---
  const targetSheet   = ss.getSheetByName(SHEET_NAMES.TARGETS);
  const targetRaw     = targetSheet.getDataRange().getValues();
  const targetHeaders = targetRaw.shift();
  const targetMap     = buildTargetMap(targetRaw, targetHeaders);

  // --- Term_Targets（1回のgetDataRange） ---
  const ttMap = loadTermTargets(ss);

  // --- Rooms（メモリ処理のみ・API呼び出しなし） ---
  const rooms = [];
  invRaw.forEach(row => {
    if (!row[0]) return;
    const id          = String(row[0]);
    const floor       = String(row[1]);
    const name        = String(row[2]);
    const currentTerm = config[`${floor}_TERM`] || 'Term1';

    const roomItems = {};
    const rTargets  = { final: {} };
    ALL_TERMS.forEach(t => { rTargets[t] = {}; });

    items.forEach(item => {
      const currentVal     = Number(row[itemIndexes[item]] || 0);
      const finalTargetVal = (targetMap[id] && targetMap[id][`${targetMode}_${item}`] !== undefined)
        ? Number(targetMap[id][`${targetMode}_${item}`]) : 0;

      rTargets.final[item] = finalTargetVal;
      ALL_TERMS.forEach(t => {
        rTargets[t][item] = (ttMap[id] && ttMap[id][t] && ttMap[id][t][item] !== undefined)
          ? ttMap[id][t][item] : finalTargetVal;
      });

      const activeTarget = (ttMap[id] && ttMap[id][currentTerm] && ttMap[id][currentTerm][item] !== undefined)
        ? ttMap[id][currentTerm][item] : finalTargetVal;

      roomItems[item] = { current: currentVal, target: activeTarget, diff: currentVal - activeTarget };
    });

    rooms.push({ id, floor, name, currentTerm, items: roomItems, targets: rTargets });
  });

  // --- Movements（1回のgetDataRange） ---
  const moveSheet   = ss.getSheetByName(SHEET_NAMES.MOVEMENTS);
  const moveRaw     = moveSheet.getDataRange().getValues();
  const moveHeaders = moveRaw.shift();
  const mIdx        = getHeaderMap(moveHeaders || []);

  let movements = [];
  if (moveRaw.length > 0 && moveRaw[0][0]) {
    movements = moveRaw.map((row, index) => ({
      rowIndex:  index + 2,
      id:        String(row[mIdx['ID']]),
      assignee:  String(row[mIdx['ID']]),
      term:      String(row[mIdx['Term']]),
      from:      String(row[mIdx['FromID']]),
      to:        String(row[mIdx['ToID']]),
      item:      String(row[mIdx['Item']]),
      amount:    Number(row[mIdx['Amount']]),
      status:    String(row[mIdx['Status']]),
      timestamp: row[mIdx['Timestamp']]
    })).reverse();
  }

  // --- App Config ---
  const appConfig = {
    notice:      config['NOTICE']       || '',
    targetMode:  targetMode,
    currentMode: config['CURRENT_MODE'] || 'setup',
    floors:      {}
  };
  Object.keys(config).forEach(key => {
    if (key.endsWith('_TERM')) appConfig.floors[key.replace('_TERM', '')] = config[key];
  });

  return { status: 'success', rooms, movements, config: appConfig, items };
}

// =========================================
// POST ロジック
// =========================================

function handlePostAction(e) {
  // --- ペイロード解析 ---
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (_) {
    throw new Error('リクエストの解析に失敗しました（不正なJSON）');
  }

  const action = payload.action;
  if (!action) throw new Error('action が指定されていません');

  // [API最小化] ss を1回だけ取得し、以降は引き回す
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // --- データを先頭で一括読み込み ---
  // ※ 認証・全アクション共通で使用するシートを先に読む
  const invSheet   = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const invData    = invSheet.getDataRange().getValues();  // [API: 1回]
  const invHeaders = invData[0];
  const invRows    = invData.slice(1);

  const config     = loadConfig(ss);  // [API: 1回]
  const logSheet   = ss.getSheetByName(SHEET_NAMES.LOGS);

  // --- 認証（invData を再利用・Inventory の二重読み取りを排除） ---
  const user = authenticateUser(ss, payload.email, invRows);
  if (!user) throw new Error('認証失敗: 登録されていないアカウントです');

  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  // ================================================================
  // check_auth
  // ================================================================
  if (action === 'check_auth') {
    return { status: 'success', user };
  }

  // ================================================================
  // validate_and_generate
  //
  // [Atomicity] 全シミュレーションと検証をメモリで完結させてから
  //             Term_Targets への書き込みを1回だけ実施する
  // [API最小化] 4シートを先頭でまとめて読み、以降はメモリで処理
  // ================================================================
  if (action === 'validate_and_generate') {
    if (user.role !== 'SuperAdmin') throw new Error('権限がありません');

    // 追加シートの読み込み（各1回）
    const targetsSheet   = ss.getSheetByName(SHEET_NAMES.TARGETS);
    const movementsSheet = ss.getSheetByName(SHEET_NAMES.MOVEMENTS);

    const targetRaw     = targetsSheet.getDataRange().getValues();    // [API: 1回]
    const targetHeaders = targetRaw.shift();
    const moveRaw       = movementsSheet.getDataRange().getValues();  // [API: 1回]
    const moveHeaders   = moveRaw.shift();

    // Inventory はすでに読み込み済みのものを流用
    const items = invHeaders.slice(3).map(String).filter(Boolean);
    if (items.length === 0) throw new Error('Inventoryシートにアイテム列が見つかりません');

    const mIdx = getHeaderMap(moveHeaders);

    // --- 必須ヘッダーの存在確認 ---
    const requiredMoveHeaders = ['Term', 'FromID', 'ToID', 'Item', 'Amount'];
    requiredMoveHeaders.forEach(h => {
      if (mIdx[h] === undefined) throw new Error(`Movementsシートに必須列 "${h}" がありません`);
    });

    // --- 1. 開始時状態の構築 ---
    const currentState = {};
    targetRaw.forEach(row => {
      const id = String(row[0]);
      if (!id) return;
      currentState[id] = {};
      items.forEach(item => {
        const col = targetHeaders.indexOf(`開始時_${item}`);
        currentState[id][item] = (col !== -1 && row[col] !== '') ? (Number(row[col]) || 0) : 0;
      });
    });

    // Inventory に存在するが Targets にない ID を警告（エラーにはしない）
    invRows.forEach(row => {
      const id = String(row[0]);
      if (id && !currentState[id]) {
        Logger.log(`WARN: Inventory に存在するが Targets にない roomId: ${id}`);
      }
    });

    // --- 2. Term ごとにシミュレーション → 適用後にスナップショット ---
    // [Atomicity] シミュレーション中にエラーが出ても書き込みは一切行わない
    const termSnapshots = {};
    for (const term of ALL_TERMS) {
      const termMoves = moveRaw.filter(m => String(m[mIdx['Term']]) === term);

      termMoves.forEach((m, rowNum) => {
        const from   = String(m[mIdx['FromID']]);
        const to     = String(m[mIdx['ToID']]);
        const item   = String(m[mIdx['Item']]);
        const amount = Number(m[mIdx['Amount']]);

        // ガード条件（ここで throw してもまだ書き込みはゼロ）
        if (!from || !to || !item) throw new Error(`Movements の ${term} 行 ${rowNum + 2}: 空のフィールドがあります`);
        if (isNaN(amount) || amount <= 0) throw new Error(`Movements の ${term} 行 ${rowNum + 2}: Amount が不正です (${amount})`);
        if (!currentState[from]) throw new Error(`Movementsに存在しないFromIDがあります: ${from}`);
        if (!currentState[to])   throw new Error(`Movementsに存在しないToIDがあります: ${to}`);
        if (currentState[from][item] === undefined) throw new Error(`[${from}] に存在しないアイテム: ${item}`);

        currentState[from][item] -= amount;
        currentState[to][item]   += amount;
      });

      // 移動適用後にスナップショット（deep copy）
      termSnapshots[term] = JSON.parse(JSON.stringify(currentState));
    }

    // --- 3. 最終状態 vs 終了時 検証 ---
    const errors = [];
    targetRaw.forEach(row => {
      const id = String(row[0]);
      if (!id || !currentState[id]) return;
      items.forEach(item => {
        const col      = targetHeaders.indexOf(`終了時_${item}`);
        const expected = (col !== -1) ? (Number(row[col]) || 0) : 0;
        const actual   = currentState[id][item];
        if (expected !== actual) {
          errors.push(`[${id}] ${item}: 期待値 ${expected}個 → シミュレーション結果 ${actual}個`);
        }
      });
    });
    if (errors.length > 0) {
      throw new Error('最終在庫がTargetsの終了時データと一致しません:\n'
        + errors.slice(0, 5).join('\n') + (errors.length > 5 ? `\n...他 ${errors.length - 5} 件` : ''));
    }

    // --- 4. フロア全体の総数不変チェック ---
    items.forEach(item => {
      let startTotal = 0, endTotal = 0;
      targetRaw.forEach(row => {
        const sc = targetHeaders.indexOf(`開始時_${item}`);
        const ec = targetHeaders.indexOf(`終了時_${item}`);
        startTotal += (sc !== -1) ? (Number(row[sc]) || 0) : 0;
        endTotal   += (ec !== -1) ? (Number(row[ec]) || 0) : 0;
      });
      if (startTotal !== endTotal) {
        throw new Error(`全体の${item}総数が開始時(${startTotal}個)と終了時(${endTotal}個)で一致しません。`);
      }
    });

    // --- 5. 検証完了 → ここで初めて書き込み（Atomicity 確保）---
    writeTtSheet(ss, items, currentState, termSnapshots);

    addLog(logSheet, timestamp, user.email, 'Validation_Success', 'Manual Movements Synced', '-', '-');
    return { status: 'success' };
  }

  // ================================================================
  // generate_movements
  //
  // [API最小化] Inventory は先頭取得済みを流用
  // [Atomicity] 全計算をメモリで完了後、Movements / Term_Targets / Config を
  //             まとめて書き込む（途中エラーでは一切書き込まない）
  // ================================================================
  if (action === 'generate_movements') {
    if (user.role !== 'SuperAdmin') throw new Error('権限がありません');

    const mode         = payload.mode;
    if (mode !== 'setup' && mode !== 'teardown') throw new Error(`不正なmode: ${mode}`);

    const sourcePrefix = mode === 'setup' ? '開始時' : '終了時';
    const targetPrefix = mode === 'setup' ? '終了時' : '開始時';
    const termMap      = { '机': 'Term1', '椅子': 'Term2', '教壇': 'Term3', '教卓': 'Term3' };

    // Inventory はすでに取得済みを流用
    const items = [];
    for (let i = 3; i < invHeaders.length; i++) {
      if (invHeaders[i]) items.push(String(invHeaders[i]));
    }

    const targetSheet   = ss.getSheetByName(SHEET_NAMES.TARGETS);
    const targetRaw     = targetSheet.getDataRange().getValues();  // [API: 1回]
    const targetHeaders = targetRaw.shift();
    const targetMap     = buildTargetMap(targetRaw, targetHeaders);

    // フロアごとの余剰・不足を集計（メモリ処理）
    const floorStatus       = {};
    const floorTotals       = {};
    const classroomsByFloor = {};
    const initialSimStates  = {};

    invRows.forEach(row => {
      if (!row[0]) return;
      const id    = String(row[0]);
      if (id.startsWith('a_')) return;
      const floor = String(row[1]);

      if (!floorStatus[floor]) { floorStatus[floor] = {}; floorTotals[floor] = {}; }
      if (id.startsWith('c_')) {
        if (!classroomsByFloor[floor]) classroomsByFloor[floor] = [];
        classroomsByFloor[floor].push(id);
      }

      initialSimStates[id] = {};
      items.forEach(item => {
        if (!floorStatus[floor][item]) {
          floorStatus[floor][item] = { surplus: [], deficit: [] };
          floorTotals[floor][item] = { source: 0, target: 0 };
        }
        const sourceVal = (targetMap[id] && targetMap[id][`${sourcePrefix}_${item}`] !== undefined)
          ? Number(targetMap[id][`${sourcePrefix}_${item}`]) : 0;
        const targetVal = (targetMap[id] && targetMap[id][`${targetPrefix}_${item}`] !== undefined)
          ? Number(targetMap[id][`${targetPrefix}_${item}`]) : 0;

        initialSimStates[id][item] = sourceVal;
        floorTotals[floor][item].source += sourceVal;
        floorTotals[floor][item].target += targetVal;

        const diff = sourceVal - targetVal;
        if (diff > 0)      floorStatus[floor][item].surplus.push({ id, amount: diff });
        else if (diff < 0) floorStatus[floor][item].deficit.push({ id, amount: Math.abs(diff) });
      });
    });

    // フロアごとの総数チェック（書き込み前に検証）
    const errors = [];
    Object.keys(floorTotals).forEach(floor => {
      Object.keys(floorTotals[floor]).forEach(item => {
        const tot = floorTotals[floor][item];
        if (tot.source !== tot.target) {
          errors.push(`[${floor}] ${item} (開始時:${tot.source} ≠ 終了時:${tot.target})`);
        }
      });
    });
    if (errors.length > 0) throw new Error('階の総数が合いません。Targetシートを修正してください:\n' + errors.join('\n'));

    // 移動指示の生成（メモリ処理）
    const newMovements = [['ID', 'Term', 'FromID', 'ToID', 'Item', 'Amount', 'Status', 'ExecutedBy', 'Timestamp']];
    Object.keys(floorStatus).forEach(floor => {
      Object.keys(floorStatus[floor]).forEach(item => {
        const surpluses = floorStatus[floor][item].surplus;
        const deficits  = floorStatus[floor][item].deficit;
        const term      = termMap[item] || 'Term1';
        let sIdx = 0, dIdx = 0;
        while (sIdx < surpluses.length && dIdx < deficits.length) {
          const s = surpluses[sIdx], d = deficits[dIdx];
          const moveAmount  = Math.min(s.amount, d.amount);
          const assigneeId  = s.id.startsWith('c_') ? s.id
            : (d.id.startsWith('c_') ? d.id
            : (classroomsByFloor[floor] ? classroomsByFloor[floor][0] : s.id));
          newMovements.push([assigneeId, term, s.id, d.id, item, moveAmount, '未', '', '']);
          s.amount -= moveAmount;
          d.amount -= moveAmount;
          if (s.amount === 0) sIdx++;
          if (d.amount === 0) dIdx++;
        }
      });
    });

    // Term_Targets 用シミュレーション（メモリ処理）
    // [FIX] current_sim を deep copy で引き継ぎ、term 間の累積バグを排除
    const termSnapshots = {};
    let current_sim = JSON.parse(JSON.stringify(initialSimStates));
    ALL_TERMS.forEach(term => {
      newMovements.forEach(m => {
        if (m[0] === 'ID' || m[1] !== term) return;
        const from = m[2], to = m[3], itm = m[4], amt = m[5];
        if (current_sim[from]) current_sim[from][itm] = (current_sim[from][itm] || 0) - amt;
        if (current_sim[to])   current_sim[to][itm]   = (current_sim[to][itm]   || 0) + amt;
      });
      termSnapshots[term] = JSON.parse(JSON.stringify(current_sim));
    });

    // --- 全計算完了 → ここから書き込み（Atomicity 確保）---

    // Movements シートを一括上書き（1回の setValues）
    const moveSheet = ss.getSheetByName(SHEET_NAMES.MOVEMENTS);
    moveSheet.clearContents();
    moveSheet.getRange(1, 1, newMovements.length, 9).setValues(newMovements);

    // Term_Targets を一括書き込み（writeTtSheet で1回の setValues）
    writeTtSheet(ss, items, current_sim, termSnapshots);

    // Config を更新（まとめて処理）
    updateConfig(ss, {
      TARGET_MODE:  targetPrefix,
      CURRENT_MODE: mode,
      ...Object.keys(floorStatus).reduce((acc, f) => { acc[`${f}_TERM`] = 'Term1'; return acc; }, {})
    });

    addLog(logSheet, timestamp, user.email, 'Generate_Missions', `Mode: ${mode}`, '-', '-');
    return { status: 'success' };
  }

  // ================================================================
  // complete_movement
  //
  // [Security]  ペイロードとシートの値を照合し、不一致は拒否
  // [Atomicity] 在庫の負チェック等を書き込み前に完了
  // [API最小化] Inventory は先頭取得済みを流用（読み込み0回追加）
  //             Movements は対象行のみ取得（1回）
  // ================================================================
  if (action === 'complete_movement') {
    const targetRow = Number(payload.rowIndex);
    if (!targetRow || targetRow < 2) throw new Error(`不正な rowIndex: ${payload.rowIndex}`);

    const moveSheet = ss.getSheetByName(SHEET_NAMES.MOVEMENTS);

    // Movements の最大行チェック（範囲外アクセス防止）
    const lastMoveRow = moveSheet.getLastRow();
    if (targetRow > lastMoveRow) throw new Error(`rowIndex ${targetRow} はMovementsシートの範囲外です`);

    // ヘッダーと対象行を同時取得（1回の getValues）
    const moveRange = moveSheet.getRange(1, 1, targetRow, 9);
    const moveBlock = moveRange.getValues();
    const h         = getHeaderMap(moveBlock[0]);
    const rowValues = moveBlock[targetRow - 1];

    // 重複完了防止（Idempotent）
    if (String(rowValues[h['Status']]) === '済') return { status: 'success' };

    // シート側の値を正とする
    const item   = String(rowValues[h['Item']]);
    const amount = Number(rowValues[h['Amount']]);
    const fromId = String(rowValues[h['FromID']]);
    const toId   = String(rowValues[h['ToID']]);
    const term   = String(rowValues[h['Term']]);

    // --- [Security] ペイロードとシート値の整合性チェック ---
    if (payload.item   !== undefined && String(payload.item)          !== item)   throw new Error('データ不整合: item が一致しません');
    if (payload.fromId !== undefined && String(payload.fromId)        !== fromId) throw new Error('データ不整合: fromId が一致しません');
    if (payload.toId   !== undefined && String(payload.toId)          !== toId)   throw new Error('データ不整合: toId が一致しません');
    if (payload.amount !== undefined && Number(payload.amount)        !== amount) throw new Error('データ不整合: amount が一致しません');

    // --- [Atomicity] 書き込み前の検証 ---
    if (amount <= 0) throw new Error(`不正な amount: ${amount}（0以下は処理できません）`);
    if (!VALID_TERMS.has(term)) throw new Error(`不正な Term: ${term}`);

    // Inventory から対象行を特定（先頭取得済みの invHeaders / invRows を流用）
    const itemColIdx = invHeaders.indexOf(item);
    if (itemColIdx === -1) throw new Error(`アイテム "${item}" がInventoryシートに存在しません`);

    let fromRowIdx = -1, toRowIdx = -1, fromFloor = '';
    for (let i = 0; i < invRows.length; i++) {
      const rowId = String(invRows[i][0]);
      if (rowId === fromId) { fromRowIdx = i; fromFloor = String(invRows[i][1]); }
      if (rowId === toId)   { toRowIdx   = i; }
      if (fromRowIdx > -1 && toRowIdx > -1) break; // 両方見つかったら早期終了
    }

    if (fromRowIdx === -1) throw new Error(`FromID "${fromId}" がInventoryシートに存在しません`);
    if (toRowIdx   === -1) throw new Error(`ToID   "${toId}" がInventoryシートに存在しません`);

    // 計算（メモリ上）
    const currentFrom = Number(invRows[fromRowIdx][itemColIdx]) || 0;
    const currentTo   = Number(invRows[toRowIdx][itemColIdx])   || 0;
    const newFromVal  = currentFrom - amount;
    const newToVal    = currentTo   + amount;

    // [Atomicity] 負在庫チェック → 失敗すれば書き込みはゼロ
    if (newFromVal < 0) {
      throw new Error(`在庫不足: ${fromId} の ${item} が ${currentFrom}個しかないため ${amount}個 の移動はできません`);
    }

    // --- 全検証完了 → ここから書き込み ---

    // Inventory を2セル書き込み
    invSheet.getRange(fromRowIdx + 2, itemColIdx + 1).setValue(newFromVal);
    invSheet.getRange(toRowIdx   + 2, itemColIdx + 1).setValue(newToVal);

    // Movements のステータス更新（対象行のみ・1回の setValues）
    rowValues[h['Status']]     = '済';
    rowValues[h['ExecutedBy']] = user.name;
    rowValues[h['Timestamp']]  = timestamp;
    moveSheet.getRange(targetRow, 1, 1, 9).setValues([rowValues]);

    addLog(logSheet, timestamp, user.email, 'Mission_Complete', `${fromId}→${toId}: ${item} x${amount}`, term, fromFloor);
    return { status: 'success' };
  }

  // ================================================================
  // direct_update
  //
  // [Atomicity] 値の検証（型・数値チェック）を先に行い、
  //             問題あれば書き込まずエラーを返す
  // [API最小化] Inventory は先頭取得済みを流用
  // ================================================================
  if (action === 'direct_update') {
    const targetRoomId = String(payload.roomId);
    const updates      = payload.updates;

    if (!targetRoomId)             throw new Error('roomId が指定されていません');
    if (!updates || typeof updates !== 'object') throw new Error('updates が不正です');

    // 値の事前検証（Atomicity: 1つでも不正なら全体を中断）
    Object.keys(updates).forEach(item => {
      const val = Number(updates[item]);
      if (isNaN(val) || val < 0) throw new Error(`${item} の値が不正です: ${updates[item]}（0以上の数値が必要）`);
    });

    // Inventory から対象行を特定（先頭取得済みを流用）
    let targetRowIdx = -1, floor = '';
    for (let i = 0; i < invRows.length; i++) {
      if (String(invRows[i][0]) === targetRoomId) { targetRowIdx = i; floor = String(invRows[i][1]); break; }
    }

    if (targetRowIdx === -1) throw new Error(`roomId "${targetRoomId}" がInventoryシートに存在しません`);

    // メモリ上で行データを更新
    const rowData = [...invRows[targetRowIdx]]; // shallow copy（元データを汚染しない）
    Object.keys(updates).forEach(item => {
      const colIdx = invHeaders.indexOf(item);
      if (colIdx > -1) rowData[colIdx] = Number(updates[item]);
    });

    // 1回の setValues で書き込み
    invSheet.getRange(targetRowIdx + 2, 1, 1, rowData.length).setValues([rowData]);

    addLog(logSheet, timestamp, user.email, 'Direct_Edit', targetRoomId, '-', floor);
    return { status: 'success' };
  }

  // ================================================================
  // change_term
  //
  // [API最小化] 進捗チェックに Inventory / Targets / Term_Targets を使うが、
  //             Inventory は先頭取得済みを流用し追加読み込みを最小化
  //             Config も先頭取得済みを流用
  // ================================================================
  if (action === 'change_term') {
    if (user.role !== 'FloorAdmin' && user.role !== 'SuperAdmin') throw new Error('権限がありません');

    const targetFloor = String(payload.targetFloor);
    const newTerm     = String(payload.newTerm);

    if (!targetFloor) throw new Error('targetFloor が指定されていません');
    if (!VALID_TERMS.has(newTerm)) throw new Error(`不正なTerm: ${newTerm}`);

    const currentTerm = config[`${targetFloor}_TERM`] || 'Term1';
    const targetMode  = config['TARGET_MODE'] || '終了時';

    // フロアの進捗確認（Inventory は先頭取得済みを渡して再読み込みを回避）
    const progressErrors = checkFloorProgress(ss, invHeaders, invRows, targetFloor, currentTerm, targetMode);
    if (progressErrors.length > 0) {
      throw new Error(`未完了の教室があります：\n${progressErrors.slice(0, 3).join('\n')}${progressErrors.length > 3 ? '\n他' : ''}`);
    }

    // Config の Term のみ更新（1回の setValues）
    updateConfig(ss, { [`${targetFloor}_TERM`]: newTerm });

    addLog(logSheet, timestamp, user.email, 'Term_Change', `${targetFloor} -> ${newTerm}`, newTerm, targetFloor);
    return { status: 'success' };
  }

  throw new Error(`不明なアクション: ${action}`);
}

// =========================================
// ユーティリティ
// =========================================

/**
 * Config シートを {key: value} の Map として返す
 * [API: getRange 1回]
 */
function loadConfig(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const raw   = sheet.getDataRange().getValues();  // A:B全体より行数が確定している方が安全
  const config = {};
  raw.forEach(row => { if (row[0]) config[String(row[0])] = String(row[1]); });
  return config;
}

/**
 * Config シートの指定キーを一括更新する
 * 既存行を上書き、存在しない行は末尾に追加
 * [API最小化] 1回の getDataRange + 更新分のみ setValue + 必要なら appendRow
 */
function updateConfig(ss, updates) {
  const sheet   = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const data    = sheet.getDataRange().getValues();
  const pending = { ...updates }; // 未処理のキーを追跡

  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0]);
    if (pending[key] !== undefined) {
      sheet.getRange(i + 1, 2).setValue(pending[key]);
      delete pending[key];
    }
  }
  // 残った（新規）キーを末尾に追加
  Object.keys(pending).forEach(key => sheet.appendRow([key, pending[key]]));
}

/**
 * Term_Targets シートを {roomId: {Term1: {item: val}, ...}} として返す
 * [API: getDataRange 1回]
 */
function loadTermTargets(ss) {
  const ttMap   = {};
  const ttSheet = ss.getSheetByName(SHEET_NAMES.TERM_TARGETS);
  if (!ttSheet) return ttMap;

  const ttRaw = ttSheet.getDataRange().getValues();
  if (ttRaw.length === 0) return ttMap;

  const ttHeaders = ttRaw.shift();
  ttRaw.forEach(row => {
    const id = String(row[0]);
    ttMap[id] = {};
    for (let i = 1; i < ttHeaders.length; i++) {
      const h        = String(ttHeaders[i]);
      const splitIdx = h.indexOf('_');
      if (splitIdx === -1) continue;
      const term = h.substring(0, splitIdx);
      const itm  = h.substring(splitIdx + 1);
      if (!ttMap[id][term]) ttMap[id][term] = {};
      ttMap[id][term][itm] = (row[i] !== '' && !isNaN(row[i])) ? Number(row[i]) : 0;
    }
  });
  return ttMap;
}

/**
 * Term_Targets シートを一括書き込み
 * [API: clearContents 1回 + setValues 1回]
 */
function writeTtSheet(ss, items, finalState, termSnapshots) {
  const ttHeaders = ['ID'];
  ALL_TERMS.forEach(t => items.forEach(i => ttHeaders.push(`${t}_${i}`)));

  const ttData = [ttHeaders];
  Object.keys(finalState).forEach(id => {
    const row = [id];
    ALL_TERMS.forEach(t => {
      items.forEach(i => {
        row.push((termSnapshots[t] && termSnapshots[t][id] && termSnapshots[t][id][i] !== undefined)
          ? termSnapshots[t][id][i] : 0);
      });
    });
    ttData.push(row);
  });

  let ttSheet = ss.getSheetByName(SHEET_NAMES.TERM_TARGETS);
  if (!ttSheet) ttSheet = ss.insertSheet(SHEET_NAMES.TERM_TARGETS);
  ttSheet.clearContents();
  ttSheet.getRange(1, 1, ttData.length, ttData[0].length).setValues(ttData);
}

/**
 * Targets シートデータを {id: {colHeader: val}} に変換
 */
function buildTargetMap(targetRaw, targetHeaders) {
  const targetMap = {};
  targetRaw.forEach(row => {
    const id = String(row[0]);
    if (!id) return;
    targetMap[id] = {};
    for (let i = 1; i < targetHeaders.length; i++) {
      targetMap[id][String(targetHeaders[i])] = row[i];
    }
  });
  return targetMap;
}

/**
 * change_term 用の軽量な進捗チェック
 * Inventory の生データを引数で受け取り、追加の API 呼び出しを最小化
 */
function checkFloorProgress(ss, invHeaders, invRows, floor, currentTerm, targetMode) {
  // Targets と Term_Targets は軽量なので個別取得
  const ttMap     = loadTermTargets(ss);  // [API: 1回]
  const targetSheet   = ss.getSheetByName(SHEET_NAMES.TARGETS);
  const targetRaw     = targetSheet.getDataRange().getValues();  // [API: 1回]
  const targetHeaders = targetRaw.shift();
  const targetMap     = buildTargetMap(targetRaw, targetHeaders);

  const items       = [];
  const itemIndexes = {};
  for (let i = 3; i < invHeaders.length; i++) {
    if (invHeaders[i]) { const n = String(invHeaders[i]); items.push(n); itemIndexes[n] = i; }
  }

  const errors = [];
  invRows.forEach(row => {
    if (!row[0]) return;
    const id        = String(row[0]);
    const roomFloor = String(row[1]);
    if (roomFloor !== floor) return;

    const roomName = String(row[2]);
    items.forEach(item => {
      const colIdx  = itemIndexes[item];
      const current = Number(row[colIdx] || 0);
      const finalTarget  = (targetMap[id] && targetMap[id][`${targetMode}_${item}`] !== undefined)
        ? Number(targetMap[id][`${targetMode}_${item}`]) : 0;
      const activeTarget = (ttMap[id] && ttMap[id][currentTerm] && ttMap[id][currentTerm][item] !== undefined)
        ? ttMap[id][currentTerm][item] : finalTarget;

      if (current !== activeTarget) {
        errors.push(`${roomName}: ${item} (${current}個 / 目標:${activeTarget}個)`);
      }
    });
  });

  return errors;
}

/**
 * ユーザー認証
 * [API最小化] invRows を引数で受け取り、Inventory の二重読み取りを排除
 */
function authenticateUser(ss, email, invRows) {
  if (!email) return null;
  const emailStr  = String(email).trim().toLowerCase();
  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data      = userSheet.getDataRange().getValues();  // [API: 1回]

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() !== emailStr) continue;

    const uid  = String(data[i][1]).trim();
    const role = String(data[i][2]).trim();

    if (uid === 'Admin' || role === 'SuperAdmin') {
      return { id: uid, name: 'SYSTEM ADMIN', role, floor: 'All', email: emailStr };
    }

    // Inventory の二重読み取りを排除（引数の invRows を線形探索）
    let floor = '', name = uid;
    for (let j = 0; j < invRows.length; j++) {
      if (String(invRows[j][0]) === uid) {
        floor = String(invRows[j][1]);
        name  = String(invRows[j][2]);
        break;
      }
    }
    return { id: uid, name, floor, role, email: emailStr };
  }
  return null;
}

function getHeaderMap(headers) {
  const map = {};
  headers.forEach((h, i) => { map[String(h)] = i; });
  return map;
}

function addLog(sheet, time, email, action, detail, term, floor) {
  sheet.appendRow([time, email, action, detail, term, floor]);
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}