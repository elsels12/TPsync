/**
 * TPSync v1.87 - Google Apps Script (GAS) バックエンド
 * 
 * 【特徴】
 * 1. LockService による同時書き込み排他制御（doPost: 最大15秒、doGet: 最大10秒待機で安全な読み書き分離）
 * 2. 初回のみテキスト列（A, G, K:L, V:X）に書式設定（@）を適用。圧力・気温列は数値型を維持し、スプレッドシート上での AVERAGE 計算やグラフ作成に対応
 * 3. bulk_upload の完全インメモリ一括処理（変更が40行以下の場合は変更行のみ書き込み、未変更行の書式崩れを完全防止）
 * 4. parseTime_ による安全なミリ秒時刻比較（NaN比較による意図しないレコード上書きを完全防止）
 * 5. doGet で getDisplayValues() を使用し、画面表示通りの文字列を正確に返却（1899年LMTタイムゾーンずれ問題を回避）
 * 6. 常に「1枚目のシート」を同期対象に固定（集計用シートやゴミ箱シートがあっても安全）
 * 7. bulk_upload 内の同一ID重複排除
 * 8. 手入力でIDが空の行に対しても、内容に基づく安定ID（rowKey_）を割り当て、編集・削除・重複防止に完全追従
 * 9. 「=,+,-,@」で始まる文字列セルに ' を付加し、数式化・勝手な数値変換を防止
 * 10. クライアント通信（30秒タイムアウト）とGASロック待機（15秒）の協調設計
 * 11. 【安心機能】削除データの「ゴミ箱」シート自動退避（90日間保持後に自動完全削除・悪意の全消し防止）
 */

/** ゴミ箱の自動保持日数（90日経過したゴミ箱データは自動で完全削除されます） */
var TRASH_RETENTION_DAYS = 90;

/** 同期対象シート = 1枚目のシート（集計・グラフ・ゴミ箱は別シートで安全に分離） */
function getDataSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

/** 削除一時保存（ゴミ箱）シートの取得・初期化 */
function getTrashSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trash = ss.getSheetByName('ゴミ箱');
  if (!trash) {
    trash = ss.insertSheet('ゴミ箱');
    var headers = ['日時','登録名','場所','路面状況','気温','路面温度','タイム','体感','銘柄_前','銘柄_後','サイズ_前','サイズ_後','冷間_FL','冷間_FR','冷間_RL','冷間_RR','温間_FL','温間_FR','温間_RL','温間_RR','メモ','ID','端末ID','更新日時','削除日時'];
    trash.appendRow(headers);
    trash.getRange("A:A").setNumberFormat("@");
    trash.getRange("G:G").setNumberFormat("@");
    trash.getRange("K:L").setNumberFormat("@");
    trash.getRange("V:Y").setNumberFormat("@");
  }
  return trash;
}

/** 指定日数を経過した古いゴミ箱行の自動クリーンアップ */
function purgeOldTrash_(days) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var trash = ss.getSheetByName('ゴミ箱');
    if (!trash) return;
    var lastRow = trash.getLastRow();
    if (lastRow <= 1) return;
    var limitMs = (days || 90) * 24 * 60 * 60 * 1000;
    var now = Date.now();
    var values = trash.getRange(2, 25, lastRow - 1, 1).getValues(); // 25列目 = 削除日時
    for (var i = values.length - 1; i >= 0; i--) {
      var dVal = values[i][0];
      if (dVal) {
        var t = new Date(dVal).getTime();
        if (!isNaN(t) && (now - t) > limitMs) {
          trash.deleteRow(i + 2);
        }
      }
    }
  } catch (err) {
    // クリーンアップエラーがメイン処理に影響しないよう保護
  }
}

/** 手動またはタイマー実行用クリーンアップ関数 */
function cleanupTrash() {
  purgeOldTrash_(TRASH_RETENTION_DAYS);
}

/** 手入力でIDが空の行の安定ID（doGet と同一式）。表示値の行から計算 */
function rowKey_(row) {
  function c(v) { return String(v || '').trim().replace(/^'/, ''); }
  var raw = c(row[21]);
  if (raw) return raw;
  return 'manual_' + c(row[0]).replace(/[^0-9]/g, '') + '_' + c(row[1]) + '_' + c(row[6]) + '_' + (c(row[12]) || c(row[16])) + '_' + (c(row[13]) || c(row[17]));
}

/** 文字列セル: 「=,+,-,@」始まりや数値に見える文字列が数式/数値に化けないよう ' を付けて文字列として保存 */
function txt_(v) {
  if (v === null || v === undefined) return '';
  var s = String(v);
  if (s === '') return '';
  if (s.charAt(0) === "'") return s; // 既に'がある場合は重ねて付加しない
  if (/^[=+\-@]/.test(s) || (s.trim() !== '' && !isNaN(Number(s)))) return "'" + s;
  return s;
}

/** 数値セル: 数値にできない値は空欄（NaN を書き込まない） */
function num_(v) {
  if (v === null || v === undefined || v === '') return '';
  var n = Number(v);
  return isFinite(n) ? n : '';
}

/** getValues() で読んだ行を別シートへ書く前に、文字列セルを再エスケープ（数式化防止。テキスト書式列 A,G,K,L,V,W,X,Y は除外） */
function escapeRowForWrite_(row) {
  var TEXT_COLS = {0:1, 6:1, 10:1, 11:1, 21:1, 22:1, 23:1, 24:1};
  return row.map(function(v, c) {
    return (!TEXT_COLS[c] && typeof v === 'string') ? txt_(v) : v;
  });
}

/** 安全なミリ秒時刻パース（パース失敗・NaN時は fallback または 0 を返却） */
function parseTime_(v, fallback) {
  if (!v) return (fallback !== undefined ? fallback : 0);
  var t = new Date(v).getTime();
  return isNaN(t) ? (fallback !== undefined ? fallback : 0) : t;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 最大15秒待機（クライアントタイムアウト30秒より短く設定し競合を解消）
    var sheet = getDataSheet_();
    var headers = ['日時','登録名','場所','路面状況','気温','路面温度','タイム','体感','銘柄_前','銘柄_後','サイズ_前','サイズ_後','冷間_FL','冷間_FR','冷間_RL','冷間_RR','温間_FL','温間_FR','温間_RL','温間_RR','メモ','ID','端末ID','更新日時'];
    
    // 初回シート初期化時のみ書式設定（テキスト列のみ@を設定し、圧力・気温列は数値型を保持）
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange("A:A").setNumberFormat("@"); // 日時 (ISO8601)
      sheet.getRange("G:G").setNumberFormat("@"); // タイム (01:23.456)
      sheet.getRange("K:L").setNumberFormat("@"); // サイズ
      sheet.getRange("V:X").setNumberFormat("@"); // ID, 端末ID, 更新日時
    }

    function buildRow(entry) {
      var cp = entry.coldPressures || (entry.mode === "cold" ? entry.pressures : null) || {};
      var hp = entry.hotPressures || (entry.mode === "hot" ? entry.pressures : null) || {};
      var safeUpdated = entry.updatedAt;
      if (!safeUpdated || isNaN(new Date(safeUpdated).getTime())) {
        safeUpdated = entry.timestamp;
        if (!safeUpdated || isNaN(new Date(safeUpdated).getTime())) {
          safeUpdated = new Date().toISOString();
        }
      }
      return [
        entry.timestamp || '',
        txt_(entry.car || 'マイカー'),
        txt_(entry.track || ''),
        txt_(entry.condition || ''),
        num_(entry.ambientTemp),
        num_(entry.trackTemp),
        entry.lapTime || '',
        txt_(entry.feel || ''),
        txt_(entry.tireBrandFront || ''),
        txt_(entry.tireBrandRear || ''),
        entry.sizeFront || '',
        entry.sizeRear || '',
        num_(cp.FL),
        num_(cp.FR),
        num_(cp.RL),
        num_(cp.RR),
        num_(hp.FL),
        num_(hp.FR),
        num_(hp.RL),
        num_(hp.RR),
        txt_(entry.notes || ''),
        String(entry.id || ''),
        String(entry.deviceId || ''),
        safeUpdated
      ];
    }

    var data = JSON.parse(e.postData.contents);

    // 1. 個別セッション保存・更新
    if (data.action === 'save' && data.entry) {
      var entry = data.entry;
      var newRow = buildRow(entry);
      var lastRow = sheet.getLastRow();
      var updated = false;
      var skipped = false;

      if (lastRow > 1 && entry.id) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, 24).getDisplayValues();
        var targetId = String(entry.id).trim().replace(/^'/, '');
        var newUpdateTime = parseTime_(entry.updatedAt || entry.timestamp, 0);

        for (var i = 0; i < existingData.length; i++) {
          if (rowKey_(existingData[i]) === targetId) {
            var existingUpdateTime = parseTime_(existingData[i][23] || existingData[i][0], 0);
            if (existingUpdateTime > 0 && (newUpdateTime === 0 || newUpdateTime < existingUpdateTime)) {
              // クラウド側の方が新しい編集（競合回避：古い未送信で上書きしない）
              skipped = true;
              updated = false;
              break;
            }
            sheet.getRange(i + 2, 1, 1, newRow.length).setValues([newRow]);
            updated = true;
            break;
          }
        }
      }
      if (!updated && !skipped) {
        sheet.appendRow(newRow);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', updated: updated, skipped: skipped })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. 一括アップロード（メモリ上で一括照合＆一括書き込み）
    if (data.action === 'bulk_upload' && Array.isArray(data.entries)) {
      // 送信データ内の同一IDは最新の updatedAt の1件に絞る（二重登録防止）
      var byId = {};
      var noId = [];
      data.entries.forEach(function(en) {
        if (!en) return;
        var k = String(en.id || '').trim().replace(/^'/, '');
        if (!k) { noId.push(en); return; }
        var cur = byId[k];
        var enTime = parseTime_(en.updatedAt || en.timestamp, 0);
        var curTime = cur ? parseTime_(cur.updatedAt || cur.timestamp, 0) : -1;
        if (!cur || enTime >= curTime) byId[k] = en;
      });
      data.entries = Object.keys(byId).map(function(k) { return byId[k]; }).concat(noId);
      var lastRow = sheet.getLastRow();
      var toAppend = [];
      var updatedCount = 0;
      var skippedCount = 0;

      if (lastRow > 1) {
        var allValues = sheet.getRange(2, 1, lastRow - 1, 24).getValues();
        var allDisplay = sheet.getRange(2, 1, lastRow - 1, 24).getDisplayValues();
        var idMap = {};
        for (var i = 0; i < allValues.length; i++) {
          var idStr = rowKey_(allDisplay[i]);
          if (idStr) idMap[idStr] = i;
        }

        var hasUpdates = false;
        var changedRows = {};
        data.entries.forEach(function(entry) {
          var row = buildRow(entry);
          var entryId = String(entry.id || '').trim().replace(/^'/, '');
          if (entryId && idMap[entryId] !== undefined) {
            var rowIdx = idMap[entryId];
            var existingRow = allDisplay[rowIdx];
            var existingUpdateTime = parseTime_(existingRow[23] || existingRow[0], 0);
            var incomingUpdateTime = parseTime_(entry.updatedAt || entry.timestamp, 0);

            if (incomingUpdateTime >= existingUpdateTime || existingUpdateTime === 0) {
              allValues[rowIdx] = row;
              allDisplay[rowIdx] = row.map(String);
              changedRows[rowIdx] = true;
              updatedCount++;
              hasUpdates = true;
            } else {
              skippedCount++;
            }
          } else {
            toAppend.push(row);
          }
        });

        var changedIdx = Object.keys(changedRows);
        if (hasUpdates && changedIdx.length <= 40) {
          // 変更が少ないときは変更行だけを書き込む（未変更行には一切触れない＝数式化・書式変化の恐れがない）
          changedIdx.forEach(function(k) {
            var ri = Number(k);
            sheet.getRange(ri + 2, 1, 1, 24).setValues([allValues[ri]]);
          });
        } else if (hasUpdates) {
          // 未変更行も一括で書き戻すため、文字列セルが数式に化けないよう再エスケープ（テキスト書式列 A,G,K,L,V,W,X 以外）
          var TEXT_COLS = {0:1, 6:1, 10:1, 11:1, 21:1, 22:1, 23:1};
          for (var r = 0; r < allValues.length; r++) {
            if (changedRows[r]) continue;
            for (var c = 0; c < 24; c++) {
              if (!TEXT_COLS[c] && typeof allValues[r][c] === 'string') allValues[r][c] = txt_(allValues[r][c]);
            }
          }
          sheet.getRange(2, 1, allValues.length, 24).setValues(allValues);
        }
      } else {
        data.entries.forEach(function(entry) {
          toAppend.push(buildRow(entry));
        });
      }

      if (toAppend.length > 0) {
        var startAppendRow = sheet.getLastRow() + 1;
        sheet.getRange(startAppendRow, 1, toAppend.length, 24).setValues(toAppend);
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        updated: updatedCount,
        appended: toAppend.length,
        skipped: skippedCount
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. 個別削除（ゴミ箱シートへ自動退避した上でメインシートから削除）
    if (data.action === 'delete' && data.id) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var idValues = sheet.getRange(2, 1, lastRow - 1, 24).getDisplayValues().map(function(r) { return [rowKey_(r)]; });
        var targetId = String(data.id).trim().replace(/^'/, '');
        for (var i = 0; i < idValues.length; i++) {
          if (String(idValues[i][0]).trim().replace(/^'/, '') === targetId) {
            // ゴミ箱へ退避
            var rowToTrash = sheet.getRange(i + 2, 1, 1, 24).getValues()[0];
            rowToTrash.push(new Date().toISOString()); // 25列目: 削除日時
            getTrashSheet_().appendRow(escapeRowForWrite_(rowToTrash));

            // メインシートから削除
            sheet.deleteRow(i + 2);
            break;
          }
        }
      }
      purgeOldTrash_(TRASH_RETENTION_DAYS);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
    }

    // 4. 一括削除（ゴミ箱シートへ一括退避した上でメインシートから削除）
    if (data.action === 'bulk_delete' && Array.isArray(data.ids) && data.ids.length > 0) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var idValues = sheet.getRange(2, 1, lastRow - 1, 24).getDisplayValues().map(function(r) { return [rowKey_(r)]; });
        var allRowValues = sheet.getRange(2, 1, lastRow - 1, 24).getValues();
        var idSet = {};
        data.ids.forEach(function(id) { idSet[String(id).trim().replace(/^'/, '')] = true; });

        var rowsToTrash = [];
        var rowsToDelete = [];
        var nowIso = new Date().toISOString();

        for (var i = 0; i < idValues.length; i++) {
          var rowId = String(idValues[i][0]).trim().replace(/^'/, '');
          if (idSet[rowId]) {
            var r = allRowValues[i];
            r.push(nowIso); // 25列目: 削除日時
            rowsToTrash.push(escapeRowForWrite_(r));
            rowsToDelete.push(i + 2);
          }
        }

        if (rowsToTrash.length > 0) {
          var trashSheet = getTrashSheet_();
          var startTrashRow = trashSheet.getLastRow() + 1;
          trashSheet.getRange(startTrashRow, 1, rowsToTrash.length, 25).setValues(rowsToTrash);
        }

        // 逆順で行削除
        for (var d = rowsToDelete.length - 1; d >= 0; d--) {
          sheet.deleteRow(rowsToDelete[d]);
        }
      }
      purgeOldTrash_(TRASH_RETENTION_DAYS);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'unknown_action' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function doGet(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 最大10秒待機（書き込み中の行入れ替わり・不完全データ取得を防止）
    var sheet = getDataSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', entries: [] })).setMimeType(ContentService.MimeType.JSON);
    }
    // getDisplayValues を使用して画面表示そのままの文字列を取得（Date型変換やタイムゾーン・LMTずれを完全回避）
    var values = sheet.getRange(2, 1, lastRow - 1, 24).getDisplayValues();

    function clean(v) {
      return String(v || '').trim().replace(/^'/, '');
    }

    var entries = values.map(function(row, idx) {
      var cfl = parseFloat(clean(row[12]));
      var cfr = parseFloat(clean(row[13]));
      var crl = parseFloat(clean(row[14]));
      var crr = parseFloat(clean(row[15]));
      var hasCold = !isNaN(cfl) || !isNaN(cfr) || !isNaN(crl) || !isNaN(crr);
      var hfl = parseFloat(clean(row[16]));
      var hfr = parseFloat(clean(row[17]));
      var hrl = parseFloat(clean(row[18]));
      var hrr = parseFloat(clean(row[19]));
      var hasHot = !isNaN(hfl) || !isNaN(hfr) || !isNaN(hrl) || !isNaN(hrr);
      var amb = parseFloat(clean(row[4]));
      var trk = parseFloat(clean(row[5]));
      var ts = clean(row[0]) || new Date().toISOString();
      var up = clean(row[23]) || ts;
      // 手入力でIDが空の行には内容に基づく安定IDを生成（削除で行番号が変わっても追従）
      var rowId = rowKey_(row);

      return {
        id: rowId,
        deviceId: clean(row[22]) || null,
        timestamp: ts,
        updatedAt: up,
        car: clean(row[1]) || 'マイカー',
        track: clean(row[2]) || null,
        condition: clean(row[3]) || 'Dry',
        ambientTemp: isNaN(amb) ? null : amb,
        trackTemp: isNaN(trk) ? null : trk,
        lapTime: clean(row[6]) || null,
        feel: clean(row[7]) || null,
        tireBrandFront: clean(row[8]) || null,
        tireBrandRear: clean(row[9]) || null,
        sizeFront: clean(row[10]) || null,
        sizeRear: clean(row[11]) || null,
        coldPressures: hasCold ? { FL: isNaN(cfl)?null:cfl, FR: isNaN(cfr)?null:cfr, RL: isNaN(crl)?null:crl, RR: isNaN(crr)?null:crr } : null,
        hotPressures: hasHot ? { FL: isNaN(hfl)?null:hfl, FR: isNaN(hfr)?null:hfr, RL: isNaN(hrl)?null:hrl, RR: isNaN(hrr)?null:hrr } : null,
        notes: clean(row[20]) || null
      };
    });
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', entries: entries })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}
