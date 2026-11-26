function verifyPasscode_(provided) {
  const expected = PropertiesService.getScriptProperties().getProperty('TEAM_PASSCODE');
  if (!expected) return true; // TEAM_PASSCODE未設定時は互換動作
  return String(provided || '').trim() === String(expected).trim();
}

function doGet(e) {
  const pass = (e && e.parameter && e.parameter.pass) || '';
  if (!verifyPasscode_(pass)) {
    return jsonOut_({ status: 'unauthorized', message: '合言葉（パスコード）が正しくありません' });
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scores = readSheetObjects_(getOrCreateScoresSheet_(ss)).reverse();
  const conditions = readSheetObjects_(getOrCreateConditionsSheet_(ss)).reverse();
  const centerMaster = readSheetObjects_(getOrCreateCenterMasterSheet_(ss));
  const centers = centerMaster.length ? centerMaster : readSheetObjects_(getOrCreateCentersSheet_(ss));
  const ballMaster = readSheetObjects_(getOrCreateBallMasterSheet_(ss));
  const players = readSheetObjects_(getOrCreatePlayersSheet_(ss));
  return jsonOut_({
    status: 'ok',
    records: scores,
    conditions: conditions,
    centers: centers,
    ballMaster: ballMaster,
    players: players
  });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (!verifyPasscode_(data.pass)) {
      return jsonOut_({ status: 'unauthorized', message: '合言葉（パスコード）が正しくありません' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.action === 'analyzeCondition') {
      return handleConditionAnalyzeOnly_(ss, data);
    }
    if (data.action === 'analyzeScoreSheet') {
      return handleScoreSheetAnalyze_(data);
    }
    if (data.action === 'updateBallMaster') {
      return withScriptLock_(() => handleUpdateBallMaster_(ss, data));
    }
    if (data.action === 'bulkImportBalls') {
      return withScriptLock_(() => handleBulkImportBalls_(ss, data));
    }
    if (data.action === 'addCenter') {
      return withScriptLock_(() => {
        const cSheet = getOrCreateCenterMasterSheet_(ss);
        cSheet.appendRow([data.pref || '', data.name || '', data.area || '', data.note || '']);
        return jsonOut_({ status: 'ok', centers: readSheetObjects_(cSheet) });
      });
    }
    if (data.action === 'bulkImportCenters') {
      return withScriptLock_(() => handleBulkImportCenters_(ss, data));
    }
    if (data.action === 'savePlayer') {
      return withScriptLock_(() => handleSavePlayer_(ss, data));
    }
    if (data.action === 'updateScore') {
      return withScriptLock_(() => handleUpdateScore_(ss, data));
    }
    if (data.action === 'deleteScore') {
      return withScriptLock_(() => handleDeleteScore_(ss, data));
    }
    if (data.action === 'parseCsv') {
      return handleCsvParse_(data);
    }
    if (data.action === 'bulkImportScores') {
      return withScriptLock_(() => {
        const sheet = getOrCreateScoresSheet_(ss);
        (data.records || []).forEach(r => appendScoreRow_(sheet, r));
        return jsonOut_({ status: 'ok', count: (data.records || []).length });
      });
    }

    return withScriptLock_(() => {
      const sheet = getOrCreateScoresSheet_(ss);
      appendScoreRow_(sheet, data);
      return jsonOut_({ status: 'ok' });
    });
  } catch (err) {
    return jsonOut_({ status: 'error', message: err.toString() });
  }
}

function handleBulkImportBalls_(ss, data) {
  const sheet = getOrCreateBallMasterSheet_(ss);
  const existingRows = sheet.getDataRange().getValues();
  const rowIndexByName = new Map();
  existingRows.slice(1).forEach((r, idx) => {
    const nm = String(r[2] || '').trim().toLowerCase();
    if (nm && !rowIndexByName.has(nm)) rowIndexByName.set(nm, idx + 2);
  });

  const toAdd = [];
  let updated = 0;
  (data.balls || []).forEach(b => {
    const maker = String(b.maker || '').trim();
    const name = String(b.name || '').trim();
    if (!maker || !name) return;
    const rowArr = [
      b.releaseDate || '',
      maker,
      name,
      b.cover || '',
      b.core || '',
      b.rg !== undefined ? b.rg : '',
      b.drg !== undefined ? b.drg : '',
      b.mb !== undefined ? b.mb : 'N/A',
      b.finish || ''
    ];
    const key = name.toLowerCase();
    if (rowIndexByName.has(key)) {
      if (data.upsert) {
        sheet.getRange(rowIndexByName.get(key), 1, 1, 9).setValues([rowArr]);
        updated++;
      }
      return;
    }
    rowIndexByName.set(key, -1);
    toAdd.push(rowArr);
  });
  if (toAdd.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 9).setValues(toAdd);
  }
  return jsonOut_({ status: 'ok', added: toAdd.length, updated: updated, total: rowIndexByName.size });
}

function handleUpdateBallMaster_(ss, data) {
  const sheet = getOrCreateBallMasterSheet_(ss);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return jsonOut_({ status: 'error', message: 'ボール一覧シートにデータがありません' });
  const headers = rows[0];
  const idxMaker = headers.indexOf('メーカー');
  const idxName = headers.indexOf('ボール名');
  const idxCover = headers.indexOf('カバーストック');
  const idxCore = headers.indexOf('コア');
  const idxRg = headers.indexOf('RG');
  const idxDrg = headers.indexOf('ΔRG');
  const idxMb = headers.indexOf('マスバイアス差');
  const idxFinish = headers.indexOf('表面仕上げ');

  const origMaker = String(data.origMaker || '').trim();
  const origName = String(data.origName || '').trim();
  let updated = false;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idxMaker]).trim() === origMaker && String(rows[i][idxName]).trim() === origName) {
      if (idxMaker >= 0 && data.maker !== undefined) sheet.getRange(i + 1, idxMaker + 1).setValue(data.maker);
      if (idxName >= 0 && data.name !== undefined) sheet.getRange(i + 1, idxName + 1).setValue(data.name);
      if (idxCover >= 0 && data.cover !== undefined) sheet.getRange(i + 1, idxCover + 1).setValue(data.cover);
      if (idxCore >= 0 && data.core !== undefined) sheet.getRange(i + 1, idxCore + 1).setValue(data.core);
      if (idxRg >= 0 && data.rg !== undefined) sheet.getRange(i + 1, idxRg + 1).setValue(data.rg);
      if (idxDrg >= 0 && data.drg !== undefined) sheet.getRange(i + 1, idxDrg + 1).setValue(data.drg);
      if (idxMb >= 0 && data.mb !== undefined) sheet.getRange(i + 1, idxMb + 1).setValue(data.mb);
      if (idxFinish >= 0 && data.finish !== undefined) sheet.getRange(i + 1, idxFinish + 1).setValue(data.finish);
      updated = true;
      break;
    }
  }
  return jsonOut_({
    status: updated ? 'ok' : 'error',
    message: updated ? '' : '対象のボールが見つかりませんでした',
    ballMaster: readSheetObjects_(sheet)
  });
}

function buildScoreRowValues_(data, existingTimestamp) {
  return [
    existingTimestamp || new Date(),
    data.date || '',
    data.player || 'ゲスト',
    data.center || '',
    data.event || '',
    Number(data.gamesCount || 0),
    Number(data.scratchTotal || 0),
    Number(data.hdcpPerGame || 0),
    Number(data.totalWithHdcp || 0),
    Number(data.average || 0),
    data.strikeRate || '',
    data.spareRate || '',
    data.tenPinRate || '',
    data.splitRate || '',
    data.punchOutRate || '',
    Number(data.highGame || 0),
    data.condName || '',
    data.condType || '',
    data.condLength || '',
    data.condVolume || '',
    data.condRatio || '',
    data.condNote || '',
    data.condImageUrl || '',
    typeof data.gamesJson === 'string' ? data.gamesJson : JSON.stringify(data.games || [])
  ];
}

function appendScoreRow_(sheet, data) {
  sheet.appendRow(buildScoreRowValues_(data, new Date()));
}

function findScoreRowIndex_(sheet, data) {
  const rowNum = Number(data.rowNumber || data._rowNumber || 0);
  const lastRow = sheet.getLastRow();
  if (rowNum >= 2 && rowNum <= lastRow) {
    return rowNum;
  }
  return -1;
}

function handleUpdateScore_(ss, data) {
  const sheet = getOrCreateScoresSheet_(ss);
  const rowIdx = findScoreRowIndex_(sheet, data);
  if (rowIdx < 2) {
    return jsonOut_({ status: 'error', message: '修正対象のスコア行が見つかりませんでした' });
  }
  const existingTs = sheet.getRange(rowIdx, 1).getValue();
  const rowValues = buildScoreRowValues_(data, existingTs);
  sheet.getRange(rowIdx, 1, 1, rowValues.length).setValues([rowValues]);
  return jsonOut_({
    status: 'ok',
    records: readSheetObjects_(sheet).reverse()
  });
}

function handleDeleteScore_(ss, data) {
  const sheet = getOrCreateScoresSheet_(ss);
  const rowIdx = findScoreRowIndex_(sheet, data);
  if (rowIdx < 2) {
    return jsonOut_({ status: 'error', message: '削除対象のスコア行が見つかりませんでした' });
  }
  sheet.deleteRow(rowIdx);
  return jsonOut_({
    status: 'ok',
    records: readSheetObjects_(sheet).reverse()
  });
}

function handleScoreSheetAnalyze_(data) {
  const base64 = (data.imageBase64 || '').split(',')[1];
  if (!base64) {
    return jsonOut_({ status: 'error', message: '画像データが空です' });
  }
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return jsonOut_({ status: 'error', message: 'スクリプトプロパティに GEMINI_API_KEY が設定されていません' });
  }

  const mode = String(data.mode || 'paper');
  const prompt = (mode === 'pins')
    ? [
        'この画像はボウリングのスコアシート（10本ピン図付きのアプリ画面・モニター画面等）です。',
        '画像から読み取れる情報を以下のJSON形式のみで返してください。',
        '{',
        '  "center": "ボウリング場名（不明なら空文字）",',
        '  "date": "YYYY-MM-DD形式の日付（不明なら空文字）",',
        '  "player": "プレイヤー名（不明なら空文字）",',
        '  "games": [',
        '    {',
        '      "reportedTotal": 合計スコア数値,',
        '      "frames": [',
        '        {',
        '          "rolls": [1投目倒ピン数, 2投目倒ピン数],',
        '          "leaves": [[1投目後に残ったピン番号の配列], [2投目後に残ったピン番号の配列]],',
        '          "splits": [1投目スプリットならtrue, false]',
        '        }',
        '      ]',
        '    }',
        '  ]',
        '}',
        '重要ルール:',
        '- framesは必ず1ゲームにつき10要素（1F〜10F）にする。',
        '- ストライク(X)は10、ガター(-/G/F)は0、スペア(/)は「10 - 1投目」の数値に変換する（例: 9と/なら [9, 1]）。',
        '- 1〜9Fでストライクのときは rolls:[10], leaves:[[]] の1要素のみ。',
        '- 10Fは最大3要素（例: X,9,/ なら rolls:[10,9,1], leaves:[[],[10],[]]）。',
        '- スプリット記号（丸囲み・赤字・Sマーク等）が付いている投球は splits を true にする。',
        '- 各フレーム下の10本ピン逆三角形ドット図（上段7,8,9,10 / 2段目4,5,6 / 3段目2,3 / 最下段1）を読み取り、黒点=1投目残りカバー(leaves[0]のみ)、赤点=ミス残り(leaves[0]とleaves[1]両方)に入れる。ピン図が無ければ空配列[]。',
        '- 出力前に各ゲームの印字合計点(reportedTotal)と各フレームの投球内容が一致するか検算する。'
      ].join('\n')
    : [
        'この画像はボウリング場の紙のスコアシート（レシート・印刷用紙）です。1人分の全ゲームを正確に読み取り、JSON形式のみで返してください。',
        '{',
        '  "center": "ボウリング場名（ロゴや印字から。不明なら空文字）",',
        '  "date": "YYYY-MM-DD形式の日付（不明なら空文字）",',
        '  "player": "プレイヤー名（不明なら空文字）",',
        '  "games": [',
        '    {',
        '      "reportedTotal": 合計スコア数値,',
        '      "frames": [',
        '        { "rolls": [1投目倒ピン数, 2投目倒ピン数], "splits": [1投目スプリットならtrue, false] }',
        '      ]',
        '    }',
        '  ]',
        '}',
        '厳守ルール:',
        '- framesは必ず1ゲームにつき10要素（1F〜10F）。leavesキーは不要。',
        '- X(ストライク)=10、- / G / F(ガター・ファウル)=0、/(スペア)=「10 - 1投目」（例: 8,/ → [8, 2]）。',
        '- 1〜9Fでストライクのときは rolls は [10] の1要素のみ。',
        '- 10Fは最大3要素（例: X,9,/ → [10,9,1]、9,/,X → [9,1,10]、9,- → [9,0]）。',
        '- スプリット記号（丸囲み・赤字・S・枠線等）は splits の該当位置を true。',
        '- 紙シートに印字されている各フレームの累計点および最終合計点(reportedTotal)と矛盾しないように、1投ごとの数値を必ず検算してから出力すること。'
      ].join('\n');

  const parsed = callGeminiJson_(apiKey, [
    { text: prompt },
    { inline_data: { mime_type: 'image/jpeg', data: base64 } }
  ]);

  return jsonOut_({
    status: 'ok',
    parsed: parsed
  });
}

function handleConditionAnalyzeOnly_(ss, data) {
  const base64 = data.imageBase64.split(',')[1];
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return jsonOut_({ status: 'error', message: 'スクリプトプロパティに GEMINI_API_KEY が設定されていません' });
  }

  const prompt = 'この画像はボウリング場のレーンコンディション表（オイルパターン表）です。画像から読み取れる情報をJSONのみで返してください。キー: patternName(パターン名), length(数字とft 例: 41ft), volume(数字とmL 例: 24.8mL), ratio(例: 1:6.5 や 1:2.8), note(オイル分布の特徴と攻略ポイントを日本語80文字以内で)';
  const parsed = callGeminiJson_(apiKey, [
    { text: prompt },
    { inline_data: { mime_type: 'image/jpeg', data: base64 } }
  ]);

  const condType = classifyCondition_(parsed.length || '', parsed.ratio || '');
  parsed.type = condType;

  const conditions = withScriptLock_(() => {
    const cSheet = getOrCreateConditionsSheet_(ss);
    cSheet.appendRow([
      new Date(),
      data.center || '',
      parsed.patternName || '無題パターン',
      condType,
      parsed.length || '',
      parsed.volume || '',
      parsed.ratio || '',
      parsed.note || '',
      ''
    ]);
    return readSheetObjects_(cSheet).reverse();
  });

  return jsonOut_({
    status: 'ok',
    imageUrl: '',
    parsed: parsed,
    conditions: conditions
  });
}

function classifyCondition_(lengthStr, ratioStr) {
  const lenMatch = String(lengthStr).match(/(\d+(?:\.\d+)?)/);
  const len = lenMatch ? parseFloat(lenMatch[1]) : 0;

  let ratioVal = 99;
  const rStr = String(ratioStr);
  if (rStr.includes(':')) {
    const parts = rStr.split(':').map(s => parseFloat(s.replace(/[^0-9.]/g, '')));
    if (parts.length === 2 && parts[0] > 0 && parts[1] > 0) {
      ratioVal = Math.max(parts[0], parts[1]) / Math.min(parts[0], parts[1]);
    }
  } else {
    const num = parseFloat(rStr.replace(/[^0-9.]/g, ''));
    if (!isNaN(num) && num > 0) ratioVal = num;
  }

  const isSport = ratioVal <= 3.0;
  if (!len) return isSport ? 'スポコンM' : 'ミドル';
  if (len <= 37) return isSport ? 'スポコンS' : 'ショート';
  if (len <= 42) return isSport ? 'スポコンM' : 'ミドル';
  return isSport ? 'スポコンL' : 'ロング';
}

function handleCsvParse_(data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return jsonOut_({ status: 'error', message: 'GEMINI_API_KEYが未設定です' });
  const sampleCsv = (data.csvText || '').split(/\r?\n/).slice(0, 40).join('\n');
  const prompt = '以下は他社ボウリングアプリやスプレッドシートのCSVデータです。各行を統一フォーマットのJSON配列に変換してください。キー: date(YYYY-MM-DD), player(プレイヤー名), center(ボウリング場), event(イベント名), gamesCount(ゲーム数 数値), scratchTotal(スクラッチ合計 数値), hdcpPerGame(1Gハンデ 数値), totalWithHdcp(ハンデ込合計 数値), average(アベレージ 数値), highGame(ハイスコア 数値), strikeRate(例: 45%), spareRate(例: 70%), tenPinRate(例: 80%), splitRate(例: 10%), punchOutRate(例: 20%), condName(パターン名), condLength(例: 41ft)\n\nCSV:\n' + sampleCsv;
  const records = callGeminiJson_(apiKey, [{ text: prompt }]);
  return jsonOut_({ status: 'ok', records: Array.isArray(records) ? records : [] });
}

function callGeminiJson_(apiKey, parts) {
  const cache = CacheService.getScriptCache();
  const cachedModel = cache.get('GEMINI_WORKING_MODEL');
  const baseCandidates = [
    'gemini-3.5-flash-lite',
    'gemini-3.5-flash',
    'gemini-3-flash-preview',
    'gemini-2.5-flash-lite',
    'gemini-2.5-flash',
    'gemini-2.0-flash',
    'gemini-flash-latest'
  ];
  const candidates = cachedModel
    ? [cachedModel, ...baseCandidates.filter(m => m !== cachedModel)]
    : baseCandidates;
  const payload = JSON.stringify({
    contents: [{ parts: parts }],
    generationConfig: { responseMimeType: 'application/json' }
  });

  const tried = {};
  let lastErr = '';
  let retryable503Model = '';

  for (let i = 0; i < candidates.length; i++) {
    const model = String(candidates[i]).replace(/^models\//, '');
    if (tried[model]) continue;
    tried[model] = true;
    const result = fetchGeminiModel_(apiKey, model, payload);
    if (result.ok) {
      if (model !== cachedModel) cache.put('GEMINI_WORKING_MODEL', model, 21600);
      return result.data;
    }
    lastErr = result.error;
    if (result.code === 503 && !retryable503Model) retryable503Model = model;
    if (![400, 404, 429, 500, 503].includes(result.code)) {
      throw new Error(lastErr);
    }
  }

  // 上記候補がすべて使えない場合、利用可能なflashモデル一覧を自動検出して順に試す
  const fallbackModels = findAvailableFlashModels_(apiKey);
  for (let i = 0; i < fallbackModels.length && i < 8; i++) {
    const model = String(fallbackModels[i]).replace(/^models\//, '');
    if (tried[model]) continue;
    tried[model] = true;
    const result = fetchGeminiModel_(apiKey, model, payload);
    if (result.ok) {
      cache.put('GEMINI_WORKING_MODEL', model, 21600);
      return result.data;
    }
    lastErr = result.error;
    if (result.code === 503 && !retryable503Model) retryable503Model = model;
    if (![400, 404, 429, 500, 503].includes(result.code)) {
      throw new Error(lastErr);
    }
  }

  // 一時混雑(503)のモデルがあった場合は少し待って1回だけ再試行
  if (retryable503Model) {
    Utilities.sleep(1800);
    const retryRes = fetchGeminiModel_(apiKey, retryable503Model, payload);
    if (retryRes.ok) {
      cache.put('GEMINI_WORKING_MODEL', retryable503Model, 21600);
      return retryRes.data;
    }
    lastErr = retryRes.error;
  }

  throw new Error(lastErr || 'Gemini API呼び出しに失敗しました');
}

function fetchGeminiModel_(apiKey, model, payload) {
  const cleanModel = String(model).replace(/^models\//, '');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + cleanModel + ':generateContent?key=' + apiKey;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const rawText = res.getContentText();
  let json;
  try {
    json = JSON.parse(rawText);
  } catch (e) {
    return { ok: false, code: code, error: 'Gemini APIから不正なレスポンスを受信しました (HTTP ' + code + ')' };
  }
  if (code < 200 || code >= 300) {
    const errMsg = (json.error && json.error.message) ? json.error.message : ('HTTP ' + code);
    return { ok: false, code: code, error: 'Gemini APIエラー (' + cleanModel + '): ' + errMsg };
  }
  const cand = json.candidates && json.candidates[0];
  const txt = cand && cand.content && cand.content.parts && cand.content.parts[0] && cand.content.parts[0].text;
  if (!txt) {
    return { ok: false, code: 500, error: 'Gemini APIから解析結果が返されませんでした' };
  }
  const cleaned = txt.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
  try {
    return { ok: true, code: code, data: JSON.parse(cleaned) };
  } catch (e) {
    return { ok: false, code: 500, error: 'Gemini APIのJSON解析に失敗しました: ' + e.message };
  }
}

function findAvailableFlashModels_(apiKey) {
  try {
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?key=' + apiKey, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return [];
    const json = JSON.parse(res.getContentText());
    const excludeRe = /(omni|tts|audio|image|live|robotics|computer-use|embedding)/i;
    const list = (json.models || []).filter(m =>
      m.name &&
      /^models\/gemini-/i.test(m.name) &&
      /flash/i.test(m.name) &&
      !excludeRe.test(m.name) &&
      Array.isArray(m.supportedGenerationMethods) &&
      m.supportedGenerationMethods.includes('generateContent')
    );
    if (!list.length) return [];
    // 新しい世代の標準flashモデルを優先（lite / preview / stable）
    list.sort((a, b) => String(b.name).localeCompare(String(a.name)));
    return list.map(m => String(m.name).replace(/^models\//, ''));
  } catch (e) {
    return [];
  }
}

function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('他のユーザーがデータを書き込み中です。数秒後にもう一度お試しください。');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateBallMasterSheet_(ss) {
  let sheet = ss.getSheetByName('ボール一覧');
  if (!sheet) {
    sheet = ss.insertSheet('ボール一覧');
    sheet.appendRow(['発売日','メーカー','ボール名','カバーストック','コア','RG','ΔRG','マスバイアス差','表面仕上げ']);
  }
  return sheet;
}

function getOrCreateScoresSheet_(ss) {
  let sheet = ss.getSheetByName('Scores');
  if (!sheet) {
    sheet = ss.insertSheet('Scores');
    sheet.appendRow(['timestamp','date','player','center','event','gamesCount','scratchTotal','hdcpPerGame','totalWithHdcp','average','strikeRate','spareRate','tenPinRate','splitRate','punchOutRate','highGame','condName','condType','condLength','condVolume','condRatio','condNote','condImageUrl','gamesJson']);
  }
  return sheet;
}

function getOrCreateConditionsSheet_(ss) {
  let sheet = ss.getSheetByName('Conditions');
  if (!sheet) {
    sheet = ss.insertSheet('Conditions');
    sheet.appendRow(['timestamp','center','patternName','type','length','volume','ratio','note','imageUrl']);
  }
  return sheet;
}

function handleSavePlayer_(ss, data) {
  const sheet = getOrCreatePlayersSheet_(ss);
  const name = String(data.player || '').trim();
  if (!name) return jsonOut_({ status: 'error', message: 'プレイヤー名を入力してください' });
  const hdcp = Number(data.hdcp || 0);
  const ballsJson = typeof data.ballsJson === 'string' ? data.ballsJson : JSON.stringify(data.balls || []);
  const now = new Date();

  const rows = sheet.getDataRange().getValues();
  let updated = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === name) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[hdcp, ballsJson, now]]);
      updated = true;
      break;
    }
  }
  if (!updated) {
    sheet.appendRow([name, hdcp, ballsJson, now]);
  }
  return jsonOut_({
    status: 'ok',
    players: readSheetObjects_(sheet)
  });
}

function getOrCreatePlayersSheet_(ss) {
  let sheet = ss.getSheetByName('Players');
  if (!sheet) {
    sheet = ss.insertSheet('Players');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['player', 'hdcp', 'ballsJson', 'updatedAt']);
  }
  if (sheet.getLastRow() === 1) {
    const defaultHoritaBalls = [
      { name: 'ブラックウィドー マニア', brand: 'HAMMER', cover: 'Pearl', core: 'Gas Mask', rg: '2.50 / 0.058', finish: '#1500 Polish' },
      { name: 'ブラックウィドー ツアー V1', brand: 'HAMMER', cover: 'Solid', core: 'Gas Mask', rg: '2.50 / 0.036', finish: '#2000 Abralon' },
      { name: 'エフェクトツアー', brand: 'HAMMER', cover: 'Solid', core: 'Huntsman', rg: '2.47 / 0.036', finish: '#2000 Abralon' },
      { name: 'アイオンマックスパール', brand: 'STORM', cover: 'Pearl', core: 'Element Max A.I.', rg: '2.47 / 0.055', finish: '#1000 Abralon' },
      { name: 'フェイズIII SE', brand: 'STORM', cover: 'Hybrid', core: 'Velocity', rg: '2.48 / 0.051', finish: '#1500 Polish' },
      { name: 'マーヴェル・マックス・ワールド', brand: 'STORM', cover: 'Solid', core: 'Centripetal', rg: '2.48 / 0.050', finish: '#2000 Abralon' },
      { name: 'ノットウレタン', brand: 'HAMMER', cover: 'Urethane', core: 'LED', rg: '2.65 / 0.015', finish: '#1000 Abralon' }
    ];
    sheet.appendRow(['堀田', 0, JSON.stringify(defaultHoritaBalls), new Date()]);
  }
  return sheet;
}

function handleBulkImportCenters_(ss, data) {
  const sheet = getOrCreateCenterMasterSheet_(ss);
  const existingRows = sheet.getDataRange().getValues();
  const existingNames = new Set(existingRows.slice(1).map(r => String(r[1] || '').trim()).filter(Boolean));
  const toAdd = [];
  (data.centers || []).forEach(c => {
    const name = String(c.name || '').trim();
    if (!name || existingNames.has(name)) return;
    existingNames.add(name);
    toAdd.push([c.pref || '', name, c.area || '', c.note || '']);
  });
  if (toAdd.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 4).setValues(toAdd);
  }
  return jsonOut_({ status: 'ok', added: toAdd.length, total: existingNames.size });
}

function getOrCreateCenterMasterSheet_(ss) {
  let sheet = ss.getSheetByName('センター一覧');
  if (!sheet) {
    sheet = ss.insertSheet('センター一覧');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['都道府県', 'センター名', '市区町村・エリア', 'メモ']);
  }
  return sheet;
}

function getOrCreateCentersSheet_(ss) {
  let sheet = ss.getSheetByName('Centers');
  if (!sheet) {
    sheet = ss.insertSheet('Centers');
    sheet.appendRow(['name','area','note']);
    sheet.appendRow(['ラウンドワン尼崎','兵庫県尼崎市','']);
    sheet.appendRow(['ボウルアロー松原店','大阪府松原市','']);
    sheet.appendRow(['エックスボウル東大阪','大阪府東大阪市','']);
    sheet.appendRow(['四條畷ボウル','大阪府四條畷市','']);
  }
  return sheet;
}

function readSheetObjects_(sheet) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  return rows.slice(1).map((r, idx) => {
    const obj = { _rowNumber: idx + 2 };
    headers.forEach((h, i) => {
      const val = r[i];
      if (val instanceof Date) {
        if (h === 'timestamp') {
          obj[h] = Utilities.formatDate(val, tz, 'yyyy-MM-dd HH:mm:ss');
        } else {
          obj[h] = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
        }
      } else {
        obj[h] = val;
      }
    });
    return obj;
  });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 初回のみApps Scriptエディタから「実行」して外部通信（Gemini API）の権限を許可するための関数
function authorizeExternalRequest() {
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/', { muteHttpExceptions: true });
  Logger.log('External request authorized. HTTP ' + res.getResponseCode());
}