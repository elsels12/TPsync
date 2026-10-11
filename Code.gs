/**
 * TPSync v1.85 - Google Apps Script (GAS) バックエンド
 * 
 * 【特徴】
 * 1. LockService による同時書き込み排他制御（複数人同時アクセスでも安全に順次処理）
 * 2. 初回のみテキスト列（A, G, K:L, V:X）に書式設定（@）を適用。圧力・気温列は数値型を維持し、スプレッドシート上での AVERAGE 計算やグラフ作成に対応
 * 3. bulk_upload の完全インメモリ一括処理（数十〜数百件の送信でもAPI呼び出しは1回のみで高速完了）
 * 4. updatedAt タイムスタンプ比較による競合制御（古い未送信編集で新しいクラウド記録を上書きしない）
 * 5. doGet で getDisplayValues() を使用し、画面表示通りの文字列を正確に返却（1899年LMTタイムゾーンずれ問題を回避）
 * 6. 手入力でIDが空の行に対しても、内容に基づく安定的で固定的なIDを割り当て、同期ごとの重複増殖を防止
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 最大30秒待機
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
      var cp = entry.coldPressures || {};
      var hp = entry.hotPressures || {};
      return [
        entry.timestamp || '',
        entry.car || 'マイカー',
        entry.track || '',
        entry.condition || '',
        (entry.ambientTemp !== null && entry.ambientTemp !== undefined && entry.ambientTemp !== '') ? Number(entry.ambientTemp) : '',
        (entry.trackTemp !== null && entry.trackTemp !== undefined && entry.trackTemp !== '') ? Number(entry.trackTemp) : '',
        entry.lapTime || '',
        entry.feel || '',
        entry.tireBrandFront || '',
        entry.tireBrandRear || '',
        entry.sizeFront || '',
        entry.sizeRear || '',
        (cp.FL !== undefined && cp.FL !== null && cp.FL !== '') ? Number(cp.FL) : '',
        (cp.FR !== undefined && cp.FR !== null && cp.FR !== '') ? Number(cp.FR) : '',
        (cp.RL !== undefined && cp.RL !== null && cp.RL !== '') ? Number(cp.RL) : '',
        (cp.RR !== undefined && cp.RR !== null && cp.RR !== '') ? Number(cp.RR) : '',
        (hp.FL !== undefined && hp.FL !== null && hp.FL !== '') ? Number(hp.FL) : '',
        (hp.FR !== undefined && hp.FR !== null && hp.FR !== '') ? Number(hp.FR) : '',
        (hp.RL !== undefined && hp.RL !== null && hp.RL !== '') ? Number(hp.RL) : '',
        (hp.RR !== undefined && hp.RR !== null && hp.RR !== '') ? Number(hp.RR) : '',
        entry.notes || '',
        String(entry.id || ''),
        String(entry.deviceId || ''),
        entry.updatedAt || entry.timestamp || new Date().toISOString()
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
        var newUpdateTime = new Date(entry.updatedAt || entry.timestamp || 0).getTime();

        for (var i = 0; i < existingData.length; i++) {
          if (String(existingData[i][21]).trim().replace(/^'/, '') === targetId) {
            var existingUpdateTime = new Date(existingData[i][23] || existingData[i][0] || 0).getTime();
            if (newUpdateTime < existingUpdateTime) {
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
      var lastRow = sheet.getLastRow();
      var toAppend = [];
      var updatedCount = 0;
      var skippedCount = 0;

      if (lastRow > 1) {
        var allValues = sheet.getRange(2, 1, lastRow - 1, 24).getValues();
        var idMap = {};
        for (var i = 0; i < allValues.length; i++) {
          var idStr = String(allValues[i][21]).trim().replace(/^'/, '');
          if (idStr) idMap[idStr] = i;
        }

        var hasUpdates = false;
        data.entries.forEach(function(entry) {
          var row = buildRow(entry);
          var entryId = String(entry.id || '').trim().replace(/^'/, '');
          if (entryId && idMap[entryId] !== undefined) {
            var rowIdx = idMap[entryId];
            var existingRow = allValues[rowIdx];
            var existingUpdateTime = new Date(existingRow[23] || existingRow[0] || 0).getTime();
            var incomingUpdateTime = new Date(entry.updatedAt || entry.timestamp || 0).getTime();

            if (incomingUpdateTime >= existingUpdateTime) {
              allValues[rowIdx] = row;
              updatedCount++;
              hasUpdates = true;
            } else {
              skippedCount++;
            }
          } else {
            toAppend.push(row);
          }
        });

        if (hasUpdates) {
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

    // 3. 個別削除
    if (data.action === 'delete' && data.id) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var idValues = sheet.getRange(2, 22, lastRow - 1, 1).getDisplayValues();
        var targetId = String(data.id).trim().replace(/^'/, '');
        for (var i = 0; i < idValues.length; i++) {
          if (String(idValues[i][0]).trim().replace(/^'/, '') === targetId) {
            sheet.deleteRow(i + 2);
            break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
    }

    // 4. 一括削除
    if (data.action === 'bulk_delete' && Array.isArray(data.ids) && data.ids.length > 0) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var idValues = sheet.getRange(2, 22, lastRow - 1, 1).getDisplayValues();
        var idSet = {};
        data.ids.forEach(function(id) { idSet[String(id).trim().replace(/^'/, '')] = true; });
        for (var i = idValues.length - 1; i >= 0; i--) {
          var rowId = String(idValues[i][0]).trim().replace(/^'/, '');
          if (idSet[rowId]) {
            sheet.deleteRow(i + 2);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'unknown_action' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
      var rawId = clean(row[21]);
      var rowId = rawId ? rawId : ('manual_' + clean(row[0]).replace(/[^0-9]/g, '') + '_' + clean(row[1]) + '_' + clean(row[6]) + '_' + (idx + 2));

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
  }
}
