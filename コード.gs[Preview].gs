/**
 * S.K.I.D Backend System (Google Account Auth Version)
 */

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

const SHEET_NAMES = {
  INVENTORY: 'Inventory',
  USERS: 'Users',
  TARGETS: 'Targets',
  MOVEMENTS: 'Movements',
  CONFIG: 'Config',
  LOGS: 'Logs',
  TERM_TARGETS: 'Term_Targets'
};

function doGet(e) { 
  const lock = LockService.getScriptLock();
  lock.tryLock(2000);
  try {
    return createJsonResponse(handleGetData(e));
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return createJsonResponse({ status: 'error', message: 'サーバーが混雑しています。数秒後に再試行してください。' });
  }
  try {
    const result = handlePostAction(e);
    SpreadsheetApp.flush(); 
    return createJsonResponse(result);
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// =========================================
// データ取得系処理 (READ)
// =========================================
function handleGetData(e) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const configRaw = configSheet.getRange("A:B").getValues();
  const config = {};
  configRaw.forEach(row => { if(row[0]) config[String(row[0])] = String(row[1]); });
  
  const targetMode = config['TARGET_MODE'] || '終了時'; 

  const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const invData = invSheet.getDataRange().getValues();
  const invHeaders = invData.shift(); 
  const items = [];
  const itemIndexes = {}; 
  
  for (let i = 3; i < invHeaders.length; i++) {
    if (invHeaders[i]) { 
      items.push(String(invHeaders[i])); 
      itemIndexes[String(invHeaders[i])] = i; 
    }
  }

  const targetSheet = ss.getSheetByName(SHEET_NAMES.TARGETS);
  const targetData = targetSheet.getDataRange().getValues();
  const targetHeaders = targetData.shift();
  const targetMap = {}; 
  targetData.forEach(row => {
    const id = String(row[0]); 
    targetMap[id] = {};
    for (let i = 1; i < targetHeaders.length; i++) { 
      targetMap[id][String(targetHeaders[i])] = row[i]; 
    }
  });

  const ttMap = {};
  const ttSheet = ss.getSheetByName(SHEET_NAMES.TERM_TARGETS);
  if (ttSheet) {
    const ttData = ttSheet.getDataRange().getValues();
    if (ttData.length > 0) {
      const ttHeaders = ttData.shift();
      ttData.forEach(row => {
        const id = String(row[0]);
        ttMap[id] = {};
        for(let i = 1; i < ttHeaders.length; i++) {
          const h = String(ttHeaders[i]);
          const splitIdx = h.indexOf('_');
          if (splitIdx > -1) {
            const term = h.substring(0, splitIdx);
            const itm = h.substring(splitIdx + 1);
            if(!ttMap[id][term]) ttMap[id][term] = {};
            ttMap[id][term][itm] = (row[i] !== '' && !isNaN(row[i])) ? Number(row[i]) : 0;
          }
        }
      });
    }
  }

  const rooms = [];
  const allTerms = ['Term1', 'Term2', 'Term3'];

  invData.forEach(row => {
    if(!row[0]) return;
    const id = String(row[0]); 
    const floor = String(row[1]);
    const name = String(row[2]);
    const currentTerm = config[`${floor}_TERM`] || 'Term1';
    
    const roomItems = {};
    const rTargets = { final: {} };
    allTerms.forEach(t => rTargets[t] = {});

    items.forEach(item => {
      const currentVal = Number(row[itemIndexes[item]] || 0);
      const targetKey = `${targetMode}_${item}`;
      const finalTargetVal = (targetMap[id] && targetMap[id][targetKey] !== undefined) ? Number(targetMap[id][targetKey]) : 0;
      
      rTargets.final[item] = finalTargetVal;
      allTerms.forEach(t => {
        rTargets[t][item] = (ttMap[id] && ttMap[id][t] && ttMap[id][t][item] !== undefined) ? ttMap[id][t][item] : finalTargetVal;
      });

      // 目標値は「現在のフロアのターム」に合わせて計算して返す
      const activeTarget = (ttMap[id] && ttMap[id][currentTerm] && ttMap[id][currentTerm][item] !== undefined) ? ttMap[id][currentTerm][item] : finalTargetVal;

      roomItems[item] = { current: currentVal, target: activeTarget, diff: currentVal - activeTarget };
    });
    
    rooms.push({ id, floor, name, currentTerm, items: roomItems, targets: rTargets });
  });

  const moveSheet = ss.getSheetByName(SHEET_NAMES.MOVEMENTS);
  const moveData = moveSheet.getDataRange().getValues();
  const moveHeaders = moveData.shift();
  const mIdx = {};
  if(moveHeaders) moveHeaders.forEach((h, i) => mIdx[String(h)] = i);
  
  let movements = [];
  if(moveData.length > 0 && moveData[0][0]) {
    movements = moveData.map((row, index) => ({
      rowIndex: index + 2, 
      id: String(row[mIdx['ID']]),
      assignee: String(row[mIdx['ID']]), 
      term: String(row[mIdx['Term']]),
      from: String(row[mIdx['FromID']]),
      to: String(row[mIdx['ToID']]),
      item: String(row[mIdx['Item']]),
      amount: Number(row[mIdx['Amount']]),
      status: String(row[mIdx['Status']]),
      timestamp: row[mIdx['Timestamp']]
    })).reverse(); 
  }

  const appConfig = { 
    notice: config['NOTICE'] || '', 
    targetMode: targetMode, 
    currentMode: config['CURRENT_MODE'] || 'setup',
    floors: {} 
  };
  Object.keys(config).forEach(key => { 
    if (key.endsWith('_TERM')) appConfig.floors[key.replace('_TERM', '')] = config[key]; 
  });

  return { status: 'success', rooms, movements, config: appConfig, items };
}

// =========================================
// データ更新系処理 (UPDATE)
// =========================================
function handlePostAction(e) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const payload = JSON.parse(e.postData.contents);
  const action = payload.action;
  
  const user = authenticateUser(ss, payload.email);
  if (!user) throw new Error('認証失敗: 登録されていないアカウントです');

  const timestamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  const logSheet = ss.getSheetByName(SHEET_NAMES.LOGS);

  if (action === 'check_auth') return { status: 'success', user: user };

  if (action === 'generate_movements') {
    if (user.role !== 'SuperAdmin') throw new Error('権限がありません');
    
    const mode = payload.mode; 
    const sourcePrefix = mode === 'setup' ? '開始時' : '終了時';
    const targetPrefix = mode === 'setup' ? '終了時' : '開始時';
    const termMap = { '机': 'Term1', '椅子': 'Term2', '教壇': 'Term3', '教卓': 'Term3' };

    const invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
    const targetSheet = ss.getSheetByName(SHEET_NAMES.TARGETS);
    const invData = invSheet.getDataRange().getValues();
    const invHeaders = invData.shift();
    const items = [];
    for (let i = 3; i < invHeaders.length; i++) { if (invHeaders[i]) items.push(String(invHeaders[i])); }
    const targetData = targetSheet.getDataRange().getValues();
    const targetHeaders = targetData.shift();
    const targetMap = {};
    targetData.forEach(row => {
      const id = String(row[0]); targetMap[id] = {};
      for (let i = 1; i < targetHeaders.length; i++) { targetMap[id][String(targetHeaders[i])] = Number(row[i]) || 0; }
    });

    const status = {}; 
    const floorTotals = {}; 
    const classroomsByFloor = {}; 
    const initialSimStates = {}; 

    invData.forEach(row => {
      if (!row[0]) return;
      const id = String(row[0]);
      if (id.startsWith('a_')) return; 
      const floor = String(row[1]);
      if (!status[floor]) { status[floor] = {}; floorTotals[floor] = {}; }
      if (id.startsWith('c_')) {
        if (!classroomsByFloor[floor]) classroomsByFloor[floor] = [];
        classroomsByFloor[floor].push(id);
      }
      initialSimStates[id] = {};
      items.forEach(item => {
        if (!status[floor][item]) { 
          status[floor][item] = { surplus: [], deficit: [] }; 
          floorTotals[floor][item] = { source: 0, target: 0 }; 
        }
        const sourceVal = targetMap[id][`${sourcePrefix}_${item}`] || 0;
        const targetVal = targetMap[id][`${targetPrefix}_${item}`] || 0;
        initialSimStates[id][item] = sourceVal;
        floorTotals[floor][item].source += sourceVal;
        floorTotals[floor][item].target += targetVal;
        const diff = sourceVal - targetVal;
        if (diff > 0) status[floor][item].surplus.push({ id: id, amount: diff });
        else if (diff < 0) status[floor][item].deficit.push({ id: id, amount: Math.abs(diff) });
      });
    });

    const errors = [];
    Object.keys(floorTotals).forEach(floor => {
      Object.keys(floorTotals[floor]).forEach(item => {
        const tot = floorTotals[floor][item];
        if (tot.source !== tot.target) errors.push(`[${floor}] ${item} (開始時:${tot.source} ≠ 終了時:${tot.target})`);
      });
    });
    if (errors.length > 0) throw new Error("階の総数が合いません。Targetシートを修正してください:\n" + errors.join("\n"));

    const newMovements = [];
    const taskCounts = {}; 
    newMovements.push(['ID', 'Term', 'FromID', 'ToID', 'Item', 'Amount', 'Status', 'ExecutedBy', 'Timestamp']);

    Object.keys(status).forEach(floor => {
      Object.keys(status[floor]).forEach(item => {
        let surpluses = status[floor][item].surplus;
        let deficits = status[floor][item].deficit;
        let term = termMap[item] || 'Term1'; 
        if (!taskCounts[term]) taskCounts[term] = {};
        let sIdx = 0, dIdx = 0;
        while (sIdx < surpluses.length && dIdx < deficits.length) {
          let s = surpluses[sIdx]; let d = deficits[dIdx];
          let moveAmount = Math.min(s.amount, d.amount);
          let assigneeId = (s.id.startsWith('c_')) ? s.id : (d.id.startsWith('c_') ? d.id : (classroomsByFloor[floor] ? classroomsByFloor[floor][0] : s.id));
          newMovements.push([assigneeId, term, s.id, d.id, item, moveAmount, '未', '', '']);
          s.amount -= moveAmount; d.amount -= moveAmount;
          if (s.amount === 0) sIdx++;
          if (d.amount === 0) dIdx++;
        }
      });
    });

    const moveSheet = ss.getSheetByName(SHEET_NAMES.MOVEMENTS);
    moveSheet.clearContents(); 
    if(newMovements.length > 0) moveSheet.getRange(1, 1, newMovements.length, 9).setValues(newMovements);

    const allTerms = ['Term1', 'Term2', 'Term3'];
    const termStates = {};
    allTerms.forEach(t => termStates[t] = {});
    let current_sim = JSON.parse(JSON.stringify(initialSimStates));
    allTerms.forEach(term => {
       newMovements.forEach(m => {
          if (m[0] === 'ID') return; 
          if (m[1] === term) {
             const from = m[2], to = m[3], itm = m[4], amt = m[5];
             if(current_sim[from]) current_sim[from][itm] -= amt;
             if(current_sim[to]) current_sim[to][itm] += amt;
          }
       });
       Object.keys(current_sim).forEach(id => { termStates[term][id] = { ...current_sim[id] }; });
    });

    const ttHeaders = ['ID'];
    allTerms.forEach(t => { items.forEach(i => ttHeaders.push(`${t}_${i}`)); });
    const ttData = [ttHeaders];
    Object.keys(initialSimStates).forEach(id => {
       const row = [id];
       allTerms.forEach(t => { items.forEach(i => { row.push(termStates[t][id] ? termStates[t][id][i] : 0); }); });
       ttData.push(row);
    });

    let ttSheet = ss.getSheetByName(SHEET_NAMES.TERM_TARGETS);
    if (!ttSheet) ttSheet = ss.insertSheet(SHEET_NAMES.TERM_TARGETS);
    ttSheet.clearContents();
    ttSheet.getRange(1, 1, ttData.length, ttData[0].length).setValues(ttData);

    const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
    const confData = configSheet.getDataRange().getValues();
    let foundTargetMode = false, foundCurrentMode = false;
    for (let i = 0; i < confData.length; i++) {
       if (confData[i][0] === 'TARGET_MODE') { configSheet.getRange(i+1, 2).setValue(targetPrefix); foundTargetMode = true; }
       if (confData[i][0] === 'CURRENT_MODE') { configSheet.getRange(i+1, 2).setValue(mode); foundCurrentMode = true; }
    }
    if (!foundTargetMode) configSheet.appendRow(['TARGET_MODE', targetPrefix]);
    if (!foundCurrentMode) configSheet.appendRow(['CURRENT_MODE', mode]);

    Object.keys(status).forEach(f => {
      let foundFloor = false;
      for (let i = 0; i < confData.length; i++) {
        if (String(confData[i][0]) === `${f}_TERM`) { configSheet.getRange(i + 1, 2).setValue('Term1'); foundFloor = true; break; }
      }
      if (!foundFloor) configSheet.appendRow([`${f}_TERM`, 'Term1']);
    });

    addLog(logSheet, timestamp, user.email, 'Generate_Missions', `Mode: ${mode}`, '-', '-');
    return { status: 'success' };
  }

  if (action === 'complete_movement') {
    const targetRow = Number(payload.rowIndex);
    const moveSheet = ss.getSheetByName(SHEET_NAMES.MOVEMENTS);
    const rowValues = moveSheet.getRange(targetRow, 1, 1, 9).getValues()[0];
    const h = getHeaderMap(moveSheet.getRange(1, 1, 1, 9).getValues()[0]);

    if (String(rowValues[h['Status']]) === '済') return { status: 'success' }; 

    const item = rowValues[h['Item']], amount = Number(rowValues[h['Amount']]), fromId = String(rowValues[h['FromID']]), toId = String(rowValues[h['ToID']]);
    updateInventoryCount(ss, fromId, item, -amount);
    updateInventoryCount(ss, toId, item, amount);

    moveSheet.getRange(targetRow, h['Status']+1).setValue('済');
    moveSheet.getRange(targetRow, h['ExecutedBy']+1).setValue(user.name);
    moveSheet.getRange(targetRow, h['Timestamp']+1).setValue(timestamp);

    addLog(logSheet, timestamp, user.email, 'Mission_Complete', `${fromId}→${toId}: ${item}`, rowValues[h['Term']], getRoomFloor(ss, fromId));
  }

  else if (action === 'direct_update') {
    const targetRoomId = String(payload.roomId);
    const updates = payload.updates;
    Object.keys(updates).forEach(item => {
      const newVal = updates[item];
      updateInventoryDirect(ss, targetRoomId, item, newVal);
    });
    addLog(logSheet, timestamp, user.email, 'Direct_Edit', targetRoomId, '-', getRoomFloor(ss, targetRoomId));
  }

  else if (action === 'change_term') {
    if (user.role !== 'FloorAdmin' && user.role !== 'SuperAdmin') throw new Error('権限なし');
    const targetFloor = String(payload.targetFloor);

    // 【新規追加：バリデーション】
    // そのフロアの全教室で、現在のタームの目標に達しているかチェックする
    const currentData = handleGetData(); // 全データを取得
    const floorRooms = currentData.rooms.filter(r => r.floor === targetFloor);
    const errors = [];

    floorRooms.forEach(room => {
      Object.keys(room.items).forEach(itemKey => {
        const item = room.items[itemKey];
        if (item.current !== item.target) {
          errors.push(`${room.name}: ${itemKey} (${item.current}個 / 目標:${item.target}個)`);
        }
      });
    });

    if (errors.length > 0) {
      throw new Error(`未完了の教室があります：\n${errors.slice(0, 3).join('\n')}${errors.length > 3 ? '\n他' : ''}`);
    }

    const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
    const data = configSheet.getDataRange().getValues();
    let found = false;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === `${targetFloor}_TERM`) { configSheet.getRange(i + 1, 2).setValue(payload.newTerm); found = true; break; }
    }
    if (!found) configSheet.appendRow([`${targetFloor}_TERM`, payload.newTerm]);
    addLog(logSheet, timestamp, user.email, 'Term_Change', `${targetFloor} -> ${payload.newTerm}`, payload.newTerm, targetFloor);
  }

  return { status: 'success' };
}

// =========================================
// ユーティリティ
// =========================================
function authenticateUser(ss, email) {
  if (!email) return null;
  const emailStr = String(email).trim().toLowerCase();
  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = userSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === emailStr) {
      const uid = String(data[i][1]).trim();
      const role = String(data[i][2]).trim();
      if (uid === 'Admin' || role === 'SuperAdmin') return { id: uid, name: 'SYSTEM ADMIN', role: role, floor: 'All', email: emailStr };
      const invInfo = getRoomInfo(ss, uid);
      return { id: uid, name: invInfo ? invInfo.name : uid, floor: invInfo ? invInfo.floor : '', role: role, email: emailStr };
    }
  }
  return null;
}

function getRoomInfo(ss, roomId) {
  const data = ss.getSheetByName(SHEET_NAMES.INVENTORY).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(roomId)) return { floor: String(data[i][1]), name: String(data[i][2]) }; }
  return null;
}

function getRoomFloor(ss, roomId) { const info = getRoomInfo(ss, roomId); return info ? info.floor : ''; }

function updateInventoryCount(ss, roomId, item, delta) {
  const sheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = sheet.getDataRange().getValues();
  const colIdx = data[0].indexOf(item);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(roomId)) {
      const cell = sheet.getRange(i+1, colIdx+1);
      cell.setValue((Number(cell.getValue()) || 0) + delta); 
      return;
    }
  }
}

function updateInventoryDirect(ss, roomId, item, value) {
  const sheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = sheet.getDataRange().getValues();
  const colIdx = data[0].indexOf(item);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(roomId)) { sheet.getRange(i+1, colIdx+1).setValue(value); return; }
  }
}

function addLog(sheet, time, user, action, detail, term, floor) { 
  const lastRow = Math.max(1, sheet.getLastRow());
  sheet.getRange(lastRow + 1, 1, 1, 6).setValues([[time, user, action, detail, term, floor]]);
}

function getHeaderMap(headers) { const map = {}; headers.forEach((h, i) => map[String(h)] = i); return map; }

function createJsonResponse(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }