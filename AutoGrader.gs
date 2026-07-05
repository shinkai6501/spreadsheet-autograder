/**
 * Spreadsheet Auto Grader for Google Apps Script.
 *
 * v2 changes:
 * - Grades one or more target sheets in one workbook-level score.
 * - Student sheet mapping is configured per answer sheet.
 * - Formula-only grading is the default, so pre-filled starter values do not
 *   reduce scores.
 * - Report-only/reference checks can be run without affecting the score.
 * - Free-form assignments can be checked without a model answer.
 */

const AUTO_GRADER = Object.freeze({
  sheets: Object.freeze({
    settings: '設定',
    submissions: '提出URL一覧',
    sheetRules: '採点対象シート',
    freeformRules: '自由課題チェック',
    spec: '採点仕様',
    results: '採点結果',
    summary: '採点サマリ'
  }),
  defaults: Object.freeze({
    useModelAnswer: true,
    sheetSelectionMode: '単一シート',
    defaultTargetMode: '数式のみ',
    totalScore: 100,
    scanMaxRows: 100,
    scanMaxCols: 30,
    numericTolerance: 1e-9,
    includeBlankCellsInSpec: false,
    includeBlankCellsInResults: false,
    includeInputLikeCellsInResults: false,
    includeReferenceInResults: true,
    treatDirectInputAsCorrect: false,
    treatResultMatchAsCorrect: true
  }),
  scope: Object.freeze({
    scored: 'SCORED',
    reference: 'REFERENCE'
  }),
  targetModes: Object.freeze({
    formulaOnly: 'FORMULA_ONLY',
    formulaAndValue: 'FORMULA_AND_VALUE',
    valueOnly: 'VALUE_ONLY',
    all: 'ALL',
    explicitOnly: 'EXPLICIT_ONLY'
  }),
  checks: Object.freeze({
    chartCountMin: 'CHART_COUNT_MIN',
    formulaCountMin: 'FORMULA_COUNT_MIN',
    nonEmptyCountMin: 'NON_EMPTY_COUNT_MIN',
    sheetCountMin: 'SHEET_COUNT_MIN',
    requiredSheetExists: 'REQUIRED_SHEET_EXISTS',
    textContains: 'TEXT_CONTAINS'
  }),
  types: Object.freeze({
    formula: 'FORMULA',
    value: 'VALUE',
    inputLike: 'INPUT_LIKE',
    blank: 'BLANK'
  })
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('自動採点')
    .addItem('初期シートを作成', 'setupAutoGrader')
    .addSeparator()
    .addItem('採点仕様を生成', 'buildGradingSpec')
    .addItem('提出URL一覧を採点', 'gradeSubmissions')
    .addItem('仕様生成と採点を実行', 'runAutoGrader')
    .addToUi();
}

function setupAutoGrader() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSettingsSheet_(ss);
  setupSubmissionSheet_(ss);
  setupSheetRuleSheet_(ss);
  setupFreeformRuleSheet_(ss);
  ensureSheet_(ss, AUTO_GRADER.sheets.spec);
  ensureSheet_(ss, AUTO_GRADER.sheets.results);
  ensureSheet_(ss, AUTO_GRADER.sheets.summary);
  notify_('初期シートを作成しました。設定、採点対象シート、提出URL一覧を入力してください。');
}

function buildGradingSpec() {
  const job = prepareJob_();
  writeSpecSheet_(job);
  notify_('採点仕様を生成しました。採点対象セル数: ' + job.targetCellCount + ' / 自由課題チェック数: ' + job.freeformRules.length);
}

function gradeSubmissions() {
  const job = prepareJob_();
  writeSpecSheet_(job);
  const submissions = readSubmissionRows_(job.controller);
  const grade = gradeSubmissionRows_(submissions, job);
  writeGradeOutputs_(job.controller, grade);
  notify_('採点が完了しました。提出数: ' + submissions.length);
}

function runAutoGrader() {
  gradeSubmissions();
}

function setupSettingsSheet_(ss) {
  const sheet = ensureSheet_(ss, AUTO_GRADER.sheets.settings);
  sheet.clear();

  const rows = [
    ['設定項目', '値', '説明'],
    ['模範解答を使う', AUTO_GRADER.defaults.useModelAnswer, 'FALSEにすると、自由課題チェックだけで採点できます。'],
    ['模範解答スプレッドシートURL', '', '空ならこの管理用スプレッドシートを模範解答として使います。'],
    ['模範解答シート名（単一シート用）', '', '単一シート採点のときだけ使います。空なら設定/出力用シート以外の先頭シートを使います。'],
    ['採点対象シート指定方法', AUTO_GRADER.defaults.sheetSelectionMode, '単一シート / すべてのシート / シート範囲 / 採点対象シート一覧。採点対象シートに有効行がある場合は一覧が優先されます。'],
    ['シート範囲開始', '', 'シート範囲モードで使います。ブック内の並び順で、このシートから採点します。'],
    ['シート範囲終了', '', 'シート範囲モードで使います。ブック内の並び順で、このシートまで採点します。'],
    ['既定の採点対象セル', AUTO_GRADER.defaults.defaultTargetMode, '数式のみ / 数式+値 / 値のみ / すべて / 明示範囲のみ。通常の穴埋め数式課題は「数式のみ」が安全です。'],
    ['総点', AUTO_GRADER.defaults.totalScore, '採点対象シートと自由課題チェックの配点が空欄の項目へ、自動で残り点を配分します。'],
    ['スキャン最大行数', AUTO_GRADER.defaults.scanMaxRows, '模範解答と自由課題チェックで読む最大行数です。'],
    ['スキャン最大列数', AUTO_GRADER.defaults.scanMaxCols, '模範解答と自由課題チェックで読む最大列数です。'],
    ['数値許容誤差', AUTO_GRADER.defaults.numericTolerance, '数値比較で許容する誤差です。'],
    ['仕様シートに空欄セルも出力', AUTO_GRADER.defaults.includeBlankCellsInSpec, 'TRUEなら解析範囲内の空欄セルも採点仕様に出します。採点対象になるわけではありません。'],
    ['採点結果に空欄セルも出力', AUTO_GRADER.defaults.includeBlankCellsInResults, 'TRUEなら空欄セルの比較も詳細結果に出します。通常はFALSE推奨です。'],
    ['採点結果に入力欄っぽいセルも出力', AUTO_GRADER.defaults.includeInputLikeCellsInResults, 'TRUEなら背景色、入力規則、メモ、表内空欄などを入力欄候補として詳細結果に出します。'],
    ['参考項目も採点結果に出力', AUTO_GRADER.defaults.includeReferenceInResults, 'TRUEなら区分が参考/採点対象外の項目も採点結果に出します。点数には入りません。'],
    ['直打ちを正答扱い', AUTO_GRADER.defaults.treatDirectInputAsCorrect, 'TRUEなら数式セルに正しい値を直打ちした場合も得点に含めます。数式を埋める課題ではFALSE推奨です。'],
    ['計算結果一致を正答扱い', AUTO_GRADER.defaults.treatResultMatchAsCorrect, 'TRUEなら数式は違っても計算結果が一致した場合を得点に含めます。']
  ];

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange(2, 2, 1, 1).insertCheckboxes();
  sheet.getRange(13, 2, 6, 1).insertCheckboxes();
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
}

function setupSubmissionSheet_(ss) {
  const sheet = ensureSheet_(ss, AUTO_GRADER.sheets.submissions);
  if (sheet.getLastRow() > 0) {
    return;
  }

  const rows = [
    ['生徒名（任意）', '提出スプレッドシートURL', '単一シート用の生徒シート名（任意）', '備考'],
    ['', 'https://docs.google.com/spreadsheets/d/...', '', '複数シート採点では「採点対象シート」シートの生徒シート列が優先されます。']
  ];
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#cfe2f3');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
}

function setupSheetRuleSheet_(ss) {
  const sheet = ensureSheet_(ss, AUTO_GRADER.sheets.sheetRules);
  if (sheet.getLastRow() > 0) {
    return;
  }

  const rows = [
    ['有効', '区分', '模範解答シート', '生徒シート', '採点対象セル', '明示範囲A1（任意）', '配点（任意）', '備考'],
    [false, '採点', 'Sheet1', '', '数式のみ', '', '', '生徒シートが空なら模範解答シートと同名のシートを採点します。'],
    [false, '参考', '確認用テスト', '', '数式のみ', '', 0, '区分を参考にすると点数には入りませんが、結果には出せます。']
  ];
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#fff2cc');
  sheet.getRange(2, 1, rows.length - 1, 1).insertCheckboxes();
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
}

function setupFreeformRuleSheet_(ss) {
  const sheet = ensureSheet_(ss, AUTO_GRADER.sheets.freeformRules);
  if (sheet.getLastRow() > 0) {
    return;
  }

  const rows = [
    ['有効', '区分', 'チェック種別', '対象シート指定', '条件値', '配点（任意）', '備考'],
    [false, '採点', 'グラフ数以上', 'すべて', 1, '', '正解が1つでない課題で、グラフが作られているかを確認します。'],
    [false, '採点', '入力セル数以上', 'すべて', 20, '', 'データ収集課題で、一定量のデータが入力されているかを確認します。'],
    [false, '参考', 'テキストを含む', 'すべて', '出典', 0, '出典など、点数外で確認したい条件に使います。']
  ];
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#eadcf8');
  sheet.getRange(2, 1, rows.length - 1, 1).insertCheckboxes();
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
}

function prepareJob_() {
  const controller = SpreadsheetApp.getActiveSpreadsheet();
  const settings = readSettings_(controller);
  const answerSpreadsheet = settings.useModelAnswer
    ? (settings.answerUrl ? SpreadsheetApp.openByUrl(settings.answerUrl) : controller)
    : null;
  const answerRules = answerSpreadsheet ? resolveAnswerRules_(controller, answerSpreadsheet, settings) : [];
  const analyses = answerRules.map(function(rule) {
    const answerSheet = chooseSheet_(answerSpreadsheet, rule.answerSheetName, null);
    return analyzeAnswerSheet_(answerSheet, settings, rule);
  });
  const freeformRules = readFreeformRules_(controller);
  allocatePoints_(analyses, freeformRules, settings);

  return {
    controller: controller,
    settings: settings,
    answerSpreadsheet: answerSpreadsheet,
    analyses: analyses,
    freeformRules: freeformRules,
    targetCellCount: analyses.reduce(function(sum, analysis) {
      return sum + analysis.targetCells.length;
    }, 0),
    maxPoints: analyses.reduce(function(sum, analysis) {
      return sum + analysis.allocatedPoints;
    }, 0) + freeformRules.reduce(function(sum, rule) {
      return sum + rule.allocatedPoints;
    }, 0)
  };
}

function readSettings_(ss) {
  const sheet = ss.getSheetByName(AUTO_GRADER.sheets.settings);
  if (!sheet || sheet.getLastRow() < 2) {
    return Object.assign({}, AUTO_GRADER.defaults, {
      answerUrl: '',
      answerSheetName: '',
      rangeStartSheetName: '',
      rangeEndSheetName: ''
    });
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const map = {};
  rows.forEach(function(row) {
    const key = String(row[0] || '').trim();
    if (key) {
      map[key] = row[1];
    }
  });

  return {
    useModelAnswer: toBool_(map['模範解答を使う'], AUTO_GRADER.defaults.useModelAnswer),
    answerUrl: String(map['模範解答スプレッドシートURL'] || '').trim(),
    answerSheetName: String(map['模範解答シート名（単一シート用）'] || map['模範解答シート名'] || '').trim(),
    sheetSelectionMode: String(map['採点対象シート指定方法'] || AUTO_GRADER.defaults.sheetSelectionMode).trim(),
    rangeStartSheetName: String(map['シート範囲開始'] || '').trim(),
    rangeEndSheetName: String(map['シート範囲終了'] || '').trim(),
    defaultTargetMode: String(map['既定の採点対象セル'] || AUTO_GRADER.defaults.defaultTargetMode).trim(),
    totalScore: toPositiveNumber_(map['総点'], AUTO_GRADER.defaults.totalScore),
    scanMaxRows: toPositiveInt_(map['スキャン最大行数'], AUTO_GRADER.defaults.scanMaxRows),
    scanMaxCols: toPositiveInt_(map['スキャン最大列数'], AUTO_GRADER.defaults.scanMaxCols),
    numericTolerance: toNumber_(map['数値許容誤差'], AUTO_GRADER.defaults.numericTolerance),
    includeBlankCellsInSpec: toBool_(map['仕様シートに空欄セルも出力'], AUTO_GRADER.defaults.includeBlankCellsInSpec),
    includeBlankCellsInResults: toBool_(map['採点結果に空欄セルも出力'], AUTO_GRADER.defaults.includeBlankCellsInResults),
    includeInputLikeCellsInResults: toBool_(map['採点結果に入力欄っぽいセルも出力'], AUTO_GRADER.defaults.includeInputLikeCellsInResults),
    includeReferenceInResults: toBool_(map['参考項目も採点結果に出力'], AUTO_GRADER.defaults.includeReferenceInResults),
    treatDirectInputAsCorrect: toBool_(map['直打ちを正答扱い'], AUTO_GRADER.defaults.treatDirectInputAsCorrect),
    treatResultMatchAsCorrect: toBool_(map['計算結果一致を正答扱い'], AUTO_GRADER.defaults.treatResultMatchAsCorrect)
  };
}

function resolveAnswerRules_(controller, answerSpreadsheet, settings) {
  const explicitRules = readSheetRules_(controller, settings);
  if (explicitRules.length > 0) {
    return explicitRules;
  }

  const sheets = resolveAnswerSheetsBySettings_(answerSpreadsheet, settings);
  return sheets.map(function(sheet, index) {
    return {
      id: 'AUTO-' + (index + 1),
      scope: AUTO_GRADER.scope.scored,
      scopeLabel: '採点',
      answerSheetName: sheet.getName(),
      studentSheetName: '',
      targetMode: normalizeTargetMode_(settings.defaultTargetMode),
      targetModeLabel: targetModeLabel_(normalizeTargetMode_(settings.defaultTargetMode)),
      explicitRangeText: '',
      specifiedPoints: null,
      note: '設定シートから自動生成'
    };
  });
}

function readSheetRules_(controller, settings) {
  const sheet = controller.getSheetByName(AUTO_GRADER.sheets.sheetRules);
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const rules = [];
  rows.forEach(function(row, index) {
    if (!toBool_(row[0], false)) {
      return;
    }

    const answerSheetName = String(row[2] || '').trim();
    if (!answerSheetName) {
      return;
    }

    const scope = normalizeScope_(row[1]);
    const targetMode = normalizeTargetMode_(row[4] || settings.defaultTargetMode);
    rules.push({
      id: 'SHEET-' + (index + 2),
      scope: scope,
      scopeLabel: scopeLabel_(scope),
      answerSheetName: answerSheetName,
      studentSheetName: String(row[3] || '').trim(),
      targetMode: targetMode,
      targetModeLabel: targetModeLabel_(targetMode),
      explicitRangeText: String(row[5] || '').trim(),
      specifiedPoints: toOptionalNumber_(row[6]),
      note: String(row[7] || '').trim()
    });
  });
  return rules;
}

function resolveAnswerSheetsBySettings_(spreadsheet, settings) {
  const mode = String(settings.sheetSelectionMode || '').trim();
  const nonSystemSheets = getNonSystemSheets_(spreadsheet);

  if (containsAny_(mode, ['すべて', '全シート', 'all'])) {
    return nonSystemSheets;
  }

  if (containsAny_(mode, ['範囲', 'range'])) {
    return resolveSheetRangeByName_(spreadsheet, settings.rangeStartSheetName, settings.rangeEndSheetName);
  }

  if (settings.answerSheetName) {
    return [chooseSheet_(spreadsheet, settings.answerSheetName, null)];
  }

  if (nonSystemSheets.length === 0) {
    throw new Error('模範解答として使えるシートが見つかりません。');
  }
  return [nonSystemSheets[0]];
}

function readFreeformRules_(controller) {
  const sheet = controller.getSheetByName(AUTO_GRADER.sheets.freeformRules);
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const rules = [];
  rows.forEach(function(row, index) {
    if (!toBool_(row[0], false)) {
      return;
    }

    const checkType = normalizeCheckType_(row[2]);
    if (!checkType) {
      return;
    }

    const scope = normalizeScope_(row[1]);
    rules.push({
      id: 'FREE-' + (index + 2),
      scope: scope,
      scopeLabel: scopeLabel_(scope),
      checkType: checkType,
      checkTypeLabel: checkTypeLabel_(checkType),
      targetSheetSpec: String(row[3] || '').trim(),
      conditionValue: row[4],
      specifiedPoints: toOptionalNumber_(row[5]),
      note: String(row[6] || '').trim(),
      allocatedPoints: 0
    });
  });
  return rules;
}

function analyzeAnswerSheet_(sheet, settings, rule) {
  const scanRows = Math.max(1, Math.min(sheet.getMaxRows(), settings.scanMaxRows));
  const scanCols = Math.max(1, Math.min(sheet.getMaxColumns(), settings.scanMaxCols));
  const grid = readGrid_(sheet, scanRows, scanCols);
  const bounds = determineEffectiveBounds_(sheet, grid);
  const explicitRanges = parseRangeSpecs_(sheet, rule.explicitRangeText);
  const cells = [];

  for (let r = 0; r < bounds.rowCount; r++) {
    for (let c = 0; c < bounds.colCount; c++) {
      const type = classifyCell_(r, c, grid, bounds);
      const cell = {
        row: r + 1,
        col: c + 1,
        a1: toA1_(r + 1, c + 1),
        type: type.type,
        typeLabel: type.label,
        inputReasons: type.reasons.join(', '),
        expectedFormulaR1C1: grid.formulasR1C1[r][c],
        expectedRawValue: grid.values[r][c],
        expectedDisplayValue: grid.displayValues[r][c],
        expectedNumberFormat: grid.numberFormats[r][c],
        note: grid.notes[r][c],
        inExplicitRange: explicitRanges.length === 0 || isCellInRanges_(r + 1, c + 1, explicitRanges),
        isTarget: false,
        cellPoint: 0
      };
      cell.isTarget = shouldTargetCell_(cell, rule, explicitRanges);
      cells.push(cell);
    }
  }

  const outputCells = cells.filter(function(cell) {
    if (cell.isTarget) {
      return true;
    }
    if (cell.type === AUTO_GRADER.types.formula || cell.type === AUTO_GRADER.types.value) {
      return true;
    }
    if (cell.type === AUTO_GRADER.types.inputLike) {
      return true;
    }
    return settings.includeBlankCellsInSpec;
  });
  const targetCells = cells.filter(function(cell) {
    return cell.isTarget;
  });
  const resultCells = cells.filter(function(cell) {
    return shouldReportCell_(cell, settings, rule);
  });

  return {
    rule: rule,
    sheet: sheet,
    sheetName: sheet.getName(),
    scanRows: scanRows,
    scanCols: scanCols,
    bounds: bounds,
    explicitRanges: explicitRanges,
    allCells: cells,
    outputCells: outputCells,
    targetCells: targetCells,
    resultCells: resultCells,
    warnings: buildAnalysisWarnings_(sheet, scanRows, scanCols),
    allocatedPoints: 0
  };
}

function readGrid_(sheet, rows, cols) {
  const range = sheet.getRange(1, 1, rows, cols);
  return {
    values: range.getValues(),
    displayValues: range.getDisplayValues(),
    formulasR1C1: range.getFormulasR1C1(),
    backgrounds: range.getBackgrounds(),
    dataValidations: range.getDataValidations(),
    notes: range.getNotes(),
    numberFormats: range.getNumberFormats(),
    rowCount: rows,
    colCount: cols
  };
}

function determineEffectiveBounds_(sheet, grid) {
  let maxRow = Math.max(1, Math.min(sheet.getLastRow() || 1, grid.rowCount));
  let maxCol = Math.max(1, Math.min(sheet.getLastColumn() || 1, grid.colCount));

  for (let r = 0; r < grid.rowCount; r++) {
    for (let c = 0; c < grid.colCount; c++) {
      if (hasCellSignal_(r, c, grid)) {
        maxRow = Math.max(maxRow, r + 1);
        maxCol = Math.max(maxCol, c + 1);
      }
    }
  }

  return {
    rowCount: maxRow,
    colCount: maxCol
  };
}

function classifyCell_(r, c, grid, bounds) {
  if (hasTex…5364 tokens truncated…'指定テキストを含むセルが見つかりません。'
    };
  }

  throw new Error('未対応の自由課題チェックです: ' + rule.checkTypeLabel);
}

function thresholdJudgement_(actual, expected, label) {
  const correct = actual >= expected;
  return {
    code: correct ? 'CHECK_PASS' : 'CHECK_FAIL',
    label: correct ? '条件達成' : '条件未達',
    correct: correct,
    expected: label + ' >= ' + expected,
    actual: actual,
    detail: label + 'は ' + actual + ' でした。'
  };
}

function judgement_(code, label, correct, detail) {
  return {
    code: code,
    label: label,
    correct: correct,
    detail: detail
  };
}

function createSummary_(submission, maxPoints) {
  return {
    timestamp: new Date(),
    studentName: submission.studentName,
    url: submission.url,
    possiblePoints: 0,
    earnedPoints: 0,
    maxPoints: maxPoints,
    targetCells: 0,
    correctCells: 0,
    targetChecks: 0,
    correctChecks: 0,
    formulaMatch: 0,
    resultMatch: 0,
    directInput: 0,
    valueMatch: 0,
    blank: 0,
    mismatch: 0,
    extraInput: 0,
    inputLike: 0,
    errors: []
  };
}

function addCellToSummary_(summary, expectedCell, judgement, shouldScore, points, earnedPoints) {
  if (shouldScore) {
    summary.targetCells += 1;
    summary.possiblePoints += points;
    summary.earnedPoints += earnedPoints;
    if (judgement.correct) {
      summary.correctCells += 1;
    }
  }

  if (judgement.code === 'FORMULA_MATCH') {
    summary.formulaMatch += 1;
  } else if (judgement.code === 'RESULT_MATCH') {
    summary.resultMatch += 1;
  } else if (judgement.code === 'DIRECT_INPUT') {
    summary.directInput += 1;
  } else if (judgement.code === 'VALUE_MATCH' || judgement.code === 'VALUE_MATCH_BY_FORMULA') {
    summary.valueMatch += 1;
  } else if (judgement.code === 'BLANK' || judgement.code === 'INPUT_BLANK') {
    summary.blank += 1;
  } else if (judgement.code === 'MISMATCH') {
    summary.mismatch += 1;
  } else if (judgement.code === 'EXTRA_INPUT') {
    summary.extraInput += 1;
  } else if (judgement.code === 'INPUT_WITH_FORMULA' || judgement.code === 'INPUT_WITH_VALUE') {
    summary.inputLike += 1;
  }
}

function summaryToRow_(summary) {
  const possiblePoints = summary.possiblePoints || summary.maxPoints;
  const rate = possiblePoints > 0 ? summary.earnedPoints / possiblePoints : '';
  const cellRate = summary.targetCells > 0 ? summary.correctCells / summary.targetCells : '';
  const checkRate = summary.targetChecks > 0 ? summary.correctChecks / summary.targetChecks : '';
  return [
    summary.timestamp,
    summary.studentName,
    summary.url,
    summary.earnedPoints,
    possiblePoints,
    rate,
    summary.targetCells,
    summary.correctCells,
    cellRate,
    summary.targetChecks,
    summary.correctChecks,
    checkRate,
    summary.formulaMatch,
    summary.resultMatch,
    summary.directInput,
    summary.valueMatch,
    summary.blank,
    summary.mismatch,
    summary.extraInput,
    summary.inputLike,
    summary.errors.join(' / ')
  ];
}

function resultRow_(data) {
  return [
    new Date(),
    data.studentName,
    data.url,
    data.scopeLabel,
    data.unitName,
    data.answerSheetName,
    data.studentSheetName,
    data.cell,
    data.category,
    data.judgementLabel,
    data.correct,
    data.points,
    data.earnedPoints,
    data.expected,
    data.actual,
    data.expectedFormula,
    data.actualFormula,
    data.detail
  ];
}

function writeGradeOutputs_(controller, grade) {
  const resultsSheet = ensureSheet_(controller, AUTO_GRADER.sheets.results);
  resultsSheet.clear();
  const resultHeaders = [
    '採点日時',
    '生徒名',
    '提出URL',
    '区分',
    '採点単位',
    '模範解答シート',
    '提出シート',
    'セル',
    '分類/チェック',
    '判定',
    '正答扱い',
    '配点',
    '得点',
    '期待値/条件',
    '提出値/実測',
    '期待数式R1C1',
    '提出数式R1C1',
    '詳細'
  ];
  resultsSheet.getRange(1, 1, 1, resultHeaders.length).setValues([resultHeaders]);
  if (grade.resultRows.length > 0) {
    resultsSheet.getRange(2, 1, grade.resultRows.length, resultHeaders.length).setValues(grade.resultRows);
  }
  resultsSheet.getRange(1, 1, 1, resultHeaders.length).setFontWeight('bold').setBackground('#d9ead3');
  resultsSheet.setFrozenRows(1);
  resultsSheet.autoResizeColumns(1, resultHeaders.length);

  const summarySheet = ensureSheet_(controller, AUTO_GRADER.sheets.summary);
  summarySheet.clear();
  const summaryHeaders = [
    '採点日時',
    '生徒名',
    '提出URL',
    '得点',
    '満点',
    '得点率',
    '採点対象セル数',
    '正答セル数',
    'セル正答率',
    '自由課題チェック数',
    '自由課題達成数',
    '自由課題達成率',
    '数式一致',
    '計算結果一致',
    '直打ち',
    '値一致',
    '空欄',
    '不一致',
    '余分な入力',
    '入力欄候補への入力',
    'エラー'
  ];
  summarySheet.getRange(1, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
  if (grade.summaryRows.length > 0) {
    summarySheet.getRange(2, 1, grade.summaryRows.length, summaryHeaders.length).setValues(grade.summaryRows);
    summarySheet.getRange(2, 6, grade.summaryRows.length, 1).setNumberFormat('0.0%');
    summarySheet.getRange(2, 9, grade.summaryRows.length, 1).setNumberFormat('0.0%');
    summarySheet.getRange(2, 12, grade.summaryRows.length, 1).setNumberFormat('0.0%');
  }
  summarySheet.getRange(1, 1, 1, summaryHeaders.length).setFontWeight('bold').setBackground('#cfe2f3');
  summarySheet.setFrozenRows(1);
  summarySheet.autoResizeColumns(1, summaryHeaders.length);
}

function chooseStudentSheetForRule_(spreadsheet, rule, submission) {
  if (rule.studentSheetName) {
    return chooseSheet_(spreadsheet, rule.studentSheetName, null);
  }

  const sameName = spreadsheet.getSheetByName(rule.answerSheetName);
  if (sameName) {
    return sameName;
  }

  if (submission.singleSheetOverride) {
    return chooseSheet_(spreadsheet, submission.singleSheetOverride, null);
  }

  throw new Error('生徒シートが見つかりません。同名シート「' + rule.answerSheetName + '」がない場合は、採点対象シートの「生徒シート」列で指定してください。');
}

function chooseSheet_(spreadsheet, preferredName, fallbackName) {
  if (preferredName) {
    const preferred = spreadsheet.getSheetByName(preferredName);
    if (!preferred) {
      throw new Error('シート「' + preferredName + '」が見つかりません: ' + spreadsheet.getName());
    }
    return preferred;
  }

  if (fallbackName) {
    const fallback = spreadsheet.getSheetByName(fallbackName);
    if (fallback) {
      return fallback;
    }
  }

  const nonSystem = getNonSystemSheets_(spreadsheet);
  return nonSystem[0] || spreadsheet.getSheets()[0];
}

function getNonSystemSheets_(spreadsheet) {
  const systemNames = Object.keys(AUTO_GRADER.sheets).map(function(key) {
    return AUTO_GRADER.sheets[key];
  });
  return spreadsheet.getSheets().filter(function(sheet) {
    return systemNames.indexOf(sheet.getName()) === -1;
  });
}

function resolveSheetRangeByName_(spreadsheet, startName, endName) {
  const sheets = getNonSystemSheets_(spreadsheet);
  if (sheets.length === 0) {
    return [];
  }
  if (!startName && !endName) {
    return sheets;
  }

  const names = sheets.map(function(sheet) {
    return sheet.getName();
  });
  const startIndex = startName ? names.indexOf(startName) : 0;
  const endIndex = endName ? names.indexOf(endName) : sheets.length - 1;
  if (startIndex === -1 || endIndex === -1) {
    throw new Error('シート範囲の開始または終了シートが見つかりません。');
  }
  const from = Math.min(startIndex, endIndex);
  const to = Math.max(startIndex, endIndex);
  return sheets.slice(from, to + 1);
}

function resolveSheetsBySpec_(spreadsheet, spec) {
  const text = String(spec || '').trim();
  if (!text || text === '*' || text === 'すべて' || text === '全シート') {
    return spreadsheet.getSheets();
  }

  if (text.indexOf(':') !== -1 && text.indexOf('!') === -1) {
    const parts = text.split(':');
    return resolveSheetRangeByName_(spreadsheet, parts[0].trim(), parts[1].trim());
  }

  const sheets = [];
  text.split(',').forEach(function(name) {
    const sheetName = name.trim();
    if (!sheetName) {
      return;
    }
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('対象シート「' + sheetName + '」が見つかりません。');
    }
    sheets.push(sheet);
  });
  return sheets;
}

function parseRangeSpecs_(sheet, rangeText) {
  const text = String(rangeText || '').trim();
  if (!text) {
    return [];
  }

  const specs = text.split(',').map(function(part) {
    return part.trim();
  }).filter(function(part) {
    return part !== '';
  });
  if (specs.length === 0) {
    return [];
  }

  return sheet.getRangeList(specs).getRanges().map(function(range) {
    return {
      rowStart: range.getRow(),
      rowEnd: range.getLastRow(),
      colStart: range.getColumn(),
      colEnd: range.getLastColumn()
    };
  });
}

function isCellInRanges_(row, col, ranges) {
  return ranges.some(function(range) {
    return row >= range.rowStart && row <= range.rowEnd && col >= range.colStart && col <= range.colEnd;
  });
}

function countFormulaCells_(sheets, settings) {
  return sheets.reduce(function(sum, sheet) {
    const rows = Math.max(1, Math.min(sheet.getMaxRows(), settings.scanMaxRows));
    const cols = Math.max(1, Math.min(sheet.getMaxColumns(), settings.scanMaxCols));
    const formulas = sheet.getRange(1, 1, rows, cols).getFormulasR1C1();
    let count = 0;
    formulas.forEach(function(row) {
      row.forEach(function(value) {
        if (hasText_(value)) {
          count += 1;
        }
      });
    });
    return sum + count;
  }, 0);
}

function countNonEmptyCells_(sheets, settings) {
  return sheets.reduce(function(sum, sheet) {
    const rows = Math.max(1, Math.min(sheet.getMaxRows(), settings.scanMaxRows));
    const cols = Math.max(1, Math.min(sheet.getMaxColumns(), settings.scanMaxCols));
    const values = sheet.getRange(1, 1, rows, cols).getDisplayValues();
    let count = 0;
    values.forEach(function(row) {
      row.forEach(function(value) {
        if (hasText_(value)) {
          count += 1;
        }
      });
    });
    return sum + count;
  }, 0);
}

function countTextContains_(sheets, settings, keyword) {
  return sheets.reduce(function(sum, sheet) {
    const rows = Math.max(1, Math.min(sheet.getMaxRows(), settings.scanMaxRows));
    const cols = Math.max(1, Math.min(sheet.getMaxColumns(), settings.scanMaxCols));
    const values = sheet.getRange(1, 1, rows, cols).getDisplayValues();
    let count = 0;
    values.forEach(function(row) {
      row.forEach(function(value) {
        if (String(value || '').indexOf(keyword) !== -1) {
          count += 1;
        }
      });
    });
    return sum + count;
  }, 0);
}

function buildAnalysisWarnings_(sheet, scanRows, scanCols) {
  const warnings = [];
  if ((sheet.getLastRow() || 1) > scanRows) {
    warnings.push('最終行がスキャン上限を超えています');
  }
  if ((sheet.getLastColumn() || 1) > scanCols) {
    warnings.push('最終列がスキャン上限を超えています');
  }
  return warnings;
}

function formulasEqual_(expectedFormula, actualFormula) {
  return normalizeFormulaR1C1_(expectedFormula) === normalizeFormulaR1C1_(actualFormula);
}

function normalizeFormulaR1C1_(formula) {
  const src = String(formula || '').trim();
  let out = '';
  let inString = false;
  let inSheetName = false;

  for (let i = 0; i < src.length; i++) {
    const ch = src.charAt(i);
    if (!inSheetName && ch === '"') {
      out += ch;
      if (inString && src.charAt(i + 1) === '"') {
        i++;
        out += '"';
        continue;
      }
      inString = !inString;
      continue;
    }
    if (!inString && ch === "'") {
      out += ch;
      if (inSheetName && src.charAt(i + 1) === "'") {
        i++;
        out += "'";
        continue;
      }
      inSheetName = !inSheetName;
      continue;
    }
    if (!inString && !inSheetName && /\s/.test(ch)) {
      continue;
    }
    out += ch;
  }

  return out;
}

function valuesEqual_(expectedRaw, expectedDisplay, actualRaw, actualDisplay, tolerance) {
  if (isBlankValue_(expectedRaw, expectedDisplay) && isBlankValue_(actualRaw, actualDisplay)) {
    return true;
  }

  const expectedNumber = comparableNumber_(expectedRaw, expectedDisplay);
  const actualNumber = comparableNumber_(actualRaw, actualDisplay);
  if (expectedNumber !== null && actualNumber !== null) {
    return Math.abs(expectedNumber - actualNumber) <= tolerance;
  }

  if (isDate_(expectedRaw) && isDate_(actualRaw)) {
    return expectedRaw.getTime() === actualRaw.getTime();
  }

  return normalizeDisplay_(expectedDisplay) === normalizeDisplay_(actualDisplay);
}

function comparableNumber_(rawValue, displayValue) {
  if (typeof rawValue === 'number' && isFinite(rawValue)) {
    return rawValue;
  }

  const text = normalizeDisplay_(displayValue).replace(/,/g, '');
  if (/^-?\d+(\.\d+)?%?$/.test(text)) {
    const value = parseFloat(text.replace('%', ''));
    return text.indexOf('%') !== -1 ? value / 100 : value;
  }
  return null;
}

function normalizeScope_(value) {
  const text = String(value || '').trim();
  if (containsAny_(text, ['参考', '確認', '対象外', 'report', 'reference'])) {
    return AUTO_GRADER.scope.reference;
  }
  return AUTO_GRADER.scope.scored;
}

function scopeLabel_(scope) {
  return scope === AUTO_GRADER.scope.reference ? '参考' : '採点';
}

function normalizeTargetMode_(value) {
  const text = String(value || '').trim();
  if (containsAny_(text, ['数式+値', '数式＋値', 'formula_and_value'])) {
    return AUTO_GRADER.targetModes.formulaAndValue;
  }
  if (containsAny_(text, ['値のみ', 'value'])) {
    return AUTO_GRADER.targetModes.valueOnly;
  }
  if (containsAny_(text, ['すべて', '全セル', 'all'])) {
    return AUTO_GRADER.targetModes.all;
  }
  if (containsAny_(text, ['明示', '範囲のみ', 'explicit'])) {
    return AUTO_GRADER.targetModes.explicitOnly;
  }
  return AUTO_GRADER.targetModes.formulaOnly;
}

function targetModeLabel_(mode) {
  if (mode === AUTO_GRADER.targetModes.formulaAndValue) {
    return '数式+値';
  }
  if (mode === AUTO_GRADER.targetModes.valueOnly) {
    return '値のみ';
  }
  if (mode === AUTO_GRADER.targetModes.all) {
    return 'すべて';
  }
  if (mode === AUTO_GRADER.targetModes.explicitOnly) {
    return '明示範囲のみ';
  }
  return '数式のみ';
}

function normalizeCheckType_(value) {
  const text = String(value || '').trim();
  if (containsAny_(text, ['グラフ', 'chart'])) {
    return AUTO_GRADER.checks.chartCountMin;
  }
  if (containsAny_(text, ['数式セル', 'formula'])) {
    return AUTO_GRADER.checks.formulaCountMin;
  }
  if (containsAny_(text, ['入力セル', '非空', 'non_empty'])) {
    return AUTO_GRADER.checks.nonEmptyCountMin;
  }
  if (containsAny_(text, ['シート数', 'sheet_count'])) {
    return AUTO_GRADER.checks.sheetCountMin;
  }
  if (containsAny_(text, ['指定シート', 'required_sheet'])) {
    return AUTO_GRADER.checks.requiredSheetExists;
  }
  if (containsAny_(text, ['テキスト', '文字', 'contains'])) {
    return AUTO_GRADER.checks.textContains;
  }
  return '';
}

function checkTypeLabel_(checkType) {
  if (checkType === AUTO_GRADER.checks.chartCountMin) {
    return 'グラフ数以上';
  }
  if (checkType === AUTO_GRADER.checks.formulaCountMin) {
    return '数式セル数以上';
  }
  if (checkType === AUTO_GRADER.checks.nonEmptyCountMin) {
    return '入力セル数以上';
  }
  if (checkType === AUTO_GRADER.checks.sheetCountMin) {
    return 'シート数以上';
  }
  if (checkType === AUTO_GRADER.checks.requiredSheetExists) {
    return '指定シートあり';
  }
  if (checkType === AUTO_GRADER.checks.textContains) {
    return 'テキストを含む';
  }
  return checkType;
}

function freeformExpectedLabel_(rule) {
  if (rule.checkType === AUTO_GRADER.checks.requiredSheetExists) {
    return '指定シート: ' + (rule.conditionValue || rule.targetSheetSpec);
  }
  return checkTypeLabel_(rule.checkType) + ': ' + rule.conditionValue;
}

function containsAny_(text, words) {
  const source = String(text || '').toLowerCase();
  return words.some(function(word) {
    return source.indexOf(String(word).toLowerCase()) !== -1;
  });
}

function ensureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function isBlankValue_(rawValue, displayValue) {
  return !hasText_(displayValue) && (rawValue === '' || rawValue === null || typeof rawValue === 'undefined');
}

function normalizeDisplay_(value) {
  return String(value === null || typeof value === 'undefined' ? '' : value).trim();
}

function isDate_(value) {
  return Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime());
}

function hasText_(value) {
  return String(value === null || typeof value === 'undefined' ? '' : value).trim() !== '';
}

function isNonDefaultBackground_(background) {
  const color = String(background || '').toLowerCase();
  return color !== '' && color !== '#ffffff' && color !== 'white';
}

function isNonDefaultNumberFormat_(format) {
  const value = String(format || '').trim();
  return value !== '' && value.toLowerCase() !== 'general';
}

function valueForOutput_(value) {
  if (isDate_(value)) {
    return value;
  }
  if (value === null || typeof value === 'undefined') {
    return '';
  }
  return value;
}

function toA1_(row, col) {
  let label = '';
  let n = col;
  while (n > 0) {
    const mod = (n - 1) % 26;
    label = String.fromCharCode(65 + mod) + label;
    n = Math.floor((n - mod - 1) / 26);
  }
  return label + row;
}

function extractSpreadsheetUrl_(text) {
  const source = String(text || '');
  const match = source.match(/https?:\/\/docs\.google\.com\/spreadsheets\/d\/[A-Za-z0-9_-]+[^\s]*/);
  return match ? match[0] : '';
}

function richTextUrl_(richText) {
  if (!richText) {
    return '';
  }

  const directUrl = richText.getLinkUrl();
  if (directUrl) {
    return extractSpreadsheetUrl_(directUrl);
  }

  const runs = richText.getRuns ? richText.getRuns() : [];
  for (let i = 0; i < runs.length; i++) {
    const runUrl = runs[i].getLinkUrl();
    if (runUrl) {
      const url = extractSpreadsheetUrl_(runUrl);
      if (url) {
        return url;
      }
    }
  }
  return '';
}

function cleanSpreadsheetUrl_(url) {
  return String(url || '').replace(/[),.。]+$/, '');
}

function toPositiveInt_(value, defaultValue) {
  const parsed = parseInt(value, 10);
  return parsed > 0 ? parsed : defaultValue;
}

function toPositiveNumber_(value, defaultValue) {
  const parsed = parseFloat(value);
  return isFinite(parsed) && parsed > 0 ? parsed : defaultValue;
}

function toNumber_(value, defaultValue) {
  const parsed = parseFloat(value);
  return isFinite(parsed) ? parsed : defaultValue;
}

function toOptionalNumber_(value) {
  if (value === null || typeof value === 'undefined' || value === '') {
    return null;
  }
  const parsed = parseFloat(value);
  return isFinite(parsed) ? parsed : null;
}

function toBool_(value, defaultValue) {
  if (typeof value === 'boolean') {
    return value;
  }
  if (value === null || typeof value === 'undefined' || value === '') {
    return defaultValue;
  }
  const text = String(value).trim().toLowerCase();
  if (['true', '1', 'yes', 'y', 'on', 'はい'].indexOf(text) !== -1) {
    return true;
  }
  if (['false', '0', 'no', 'n', 'off', 'いいえ'].indexOf(text) !== -1) {
    return false;
  }
  return defaultValue;
}

function notify_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    Logger.log(message);
  }
}

