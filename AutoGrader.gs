/**
 * Classroom Spreadsheet Folder Auto Grader.
 *
 * A teacher supplies:
 * - the Google Drive folder URL created for a Classroom assignment
 * - the model-answer spreadsheet URL
 *
 * The script scans every Google Sheets file in the folder, compares formula
 * cells on sheets with matching names, and creates a new result spreadsheet.
 */

const AUTO_GRADER = Object.freeze({
  version: '3.1.0',
  settingsSheet: '設定',
  specSheet: '採点仕様',
  resultSheets: Object.freeze({
    summary: '採点サマリ',
    details: '採点詳細',
    info: '実行情報'
  }),
  resultMarker: '[SpreadsheetAutoGraderResult]',
  defaults: Object.freeze({
    resultFileName: 'スプレッドシート課題_採点結果',
    totalScore: 50,
    chartScoreRatio: 0.3,
    searchSubfolders: true,
    includeHiddenAnswerSheets: true,
    excludedSheetNames: '',
    treatResultMatchAsCorrect: false,
    treatDirectInputAsCorrect: false,
    outputDetails: true,
    numericTolerance: 1e-9
  })
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('フォルダ自動採点')
    .addItem('初期設定シートを作成', 'setupAutoGrader')
    .addSeparator()
    .addItem('模範解答から採点仕様を確認', 'buildGradingSpec')
    .addItem('提出フォルダを採点', 'runAutoGrader')
    .addToUi();
}

function setupAutoGrader() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(spreadsheet, AUTO_GRADER.settingsSheet);
  sheet.clear();

  const rows = [
    ['設定項目', '値', '説明'],
    ['提出物フォルダURL', '', 'Classroomが課題用に作成したGoogle DriveフォルダのURLです。'],
    ['模範解答スプレッドシートURL', '', '数式が入力済みの模範解答スプレッドシートURLです。'],
    ['結果ファイル名', AUTO_GRADER.defaults.resultFileName, '実行日時を末尾に付けた新しいスプレッドシートを提出フォルダ内へ作ります。'],
    ['総点', AUTO_GRADER.defaults.totalScore, '数式とグラフを合わせた満点です。初期値は50点です。'],
    ['グラフ配点割合', AUTO_GRADER.defaults.chartScoreRatio, '模範解答にグラフがある場合の配点割合です。0.3なら50点中15点です。'],
    ['サブフォルダも検索', AUTO_GRADER.defaults.searchSubfolders, 'TRUEなら提出フォルダ配下も再帰的に検索します。'],
    ['非表示の模範解答シートも採点', AUTO_GRADER.defaults.includeHiddenAnswerSheets, 'TRUEなら非表示シートも同名シートと照合します。'],
    ['採点対象外シート名', AUTO_GRADER.defaults.excludedSheetNames, '採点しないシート名をカンマ区切りで入力します。空欄なら全シートです。'],
    ['計算結果一致を正答扱い', AUTO_GRADER.defaults.treatResultMatchAsCorrect, '数式が異なっても計算結果が一致した場合を正答にするか指定します。'],
    ['直打ちを正答扱い', AUTO_GRADER.defaults.treatDirectInputAsCorrect, '数式ではなく結果を直接入力した場合を正答にするか指定します。'],
    ['採点詳細を出力', AUTO_GRADER.defaults.outputDetails, 'TRUEならセルごとの判定を採点詳細シートへ出力します。'],
    ['数値許容誤差', AUTO_GRADER.defaults.numericTolerance, '計算結果の数値比較で許容する誤差です。']
  ];

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange(7, 2, 2, 1).insertCheckboxes();
  sheet.getRange(10, 2, 3, 1).insertCheckboxes();
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
  notify_('設定シートを作成しました。提出物フォルダURLと模範解答スプレッドシートURLを入力してください。');
}

function buildGradingSpec() {
  const controller = SpreadsheetApp.getActiveSpreadsheet();
  const settings = readSettings_(controller);
  const answerSpreadsheet = SpreadsheetApp.openById(settings.answerSpreadsheetId);
  const spec = analyzeAnswerSpreadsheet_(answerSpreadsheet, settings);
  writeSpecSheet_(controller, answerSpreadsheet, spec);
  notify_('採点仕様を生成しました。対象シート数: ' + spec.sheets.length + ' / 数式セル数: ' + spec.totalCells + ' / グラフ数: ' + spec.totalCharts);
}

function runAutoGrader() {
  const controller = SpreadsheetApp.getActiveSpreadsheet();
  const settings = readSettings_(controller);
  const folder = DriveApp.getFolderById(settings.submissionFolderId);
  const answerSpreadsheet = SpreadsheetApp.openById(settings.answerSpreadsheetId);
  const spec = analyzeAnswerSpreadsheet_(answerSpreadsheet, settings);

  if (spec.totalCells === 0 && spec.totalCharts === 0) {
    throw new Error('模範解答に採点対象となる数式セルまたはグラフがありません。');
  }

  const excludedIds = {};
  excludedIds[settings.answerSpreadsheetId] = true;
  excludedIds[controller.getId()] = true;
  const submissionFiles = collectSubmissionFiles_(folder, settings.searchSubfolders, excludedIds);
  if (submissionFiles.length === 0) {
    throw new Error('提出フォルダ内に採点可能なGoogleスプレッドシートが見つかりません。');
  }

  const grading = gradeSubmissionFiles_(submissionFiles, spec, settings);
  const result = createResultSpreadsheet_(folder, answerSpreadsheet, spec, grading, settings);
  writeSpecSheet_(controller, answerSpreadsheet, spec);
  notify_('採点が完了しました。提出数: ' + submissionFiles.length + '\n結果: ' + result.getUrl());
}

// Backward-compatible alias for the previous menu/function name.
function gradeSubmissions() {
  runAutoGrader();
}

function readSettings_(controller) {
  const sheet = controller.getSheetByName(AUTO_GRADER.settingsSheet);
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('最初に「フォルダ自動採点 > 初期設定シートを作成」を実行してください。');
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const map = {};
  values.forEach(function(row) {
    const key = String(row[0] || '').trim();
    if (key) {
      map[key] = row[1];
    }
  });

  const folderUrl = String(map['提出物フォルダURL'] || '').trim();
  const answerUrl = String(map['模範解答スプレッドシートURL'] || '').trim();
  if (!folderUrl) {
    throw new Error('設定シートの「提出物フォルダURL」を入力してください。');
  }
  if (!answerUrl) {
    throw new Error('設定シートの「模範解答スプレッドシートURL」を入力してください。');
  }

  return {
    folderUrl: folderUrl,
    answerUrl: answerUrl,
    submissionFolderId: extractDriveId_(folderUrl, 'folders'),
    answerSpreadsheetId: extractDriveId_(answerUrl, 'spreadsheets/d'),
    resultFileName: String(map['結果ファイル名'] || AUTO_GRADER.defaults.resultFileName).trim(),
    totalScore: toPositiveNumber_(map['総点'], AUTO_GRADER.defaults.totalScore),
    chartScoreRatio: toRatio_(map['グラフ配点割合'], AUTO_GRADER.defaults.chartScoreRatio),
    searchSubfolders: toBool_(map['サブフォルダも検索'], AUTO_GRADER.defaults.searchSubfolders),
    includeHiddenAnswerSheets: toBool_(map['非表示の模範解答シートも採点'], AUTO_GRADER.defaults.includeHiddenAnswerSheets),
    excludedSheetNames: parseNameList_(map['採点対象外シート名']),
    treatResultMatchAsCorrect: toBool_(map['計算結果一致を正答扱い'], AUTO_GRADER.defaults.treatResultMatchAsCorrect),
    treatDirectInputAsCorrect: toBool_(map['直打ちを正答扱い'], AUTO_GRADER.defaults.treatDirectInputAsCorrect),
    outputDetails: toBool_(map['採点詳細を出力'], AUTO_GRADER.defaults.outputDetails),
    numericTolerance: toNumber_(map['数値許容誤差'], AUTO_GRADER.defaults.numericTolerance)
  };
}

function analyzeAnswerSpreadsheet_(spreadsheet, settings) {
  const excluded = {};
  settings.excludedSheetNames.forEach(function(name) {
    excluded[name] = true;
  });

  const sheets = [];
  let totalCells = 0;
  let totalCharts = 0;
  spreadsheet.getSheets().forEach(function(sheet) {
    if (excluded[sheet.getName()]) {
      return;
    }
    if (!settings.includeHiddenAnswerSheets && sheet.isSheetHidden()) {
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const cells = [];
    if (lastRow >= 1 && lastCol >= 1) {
      const range = sheet.getRange(1, 1, lastRow, lastCol);
      const formulas = range.getFormulasR1C1();
      const values = range.getValues();
      const displayValues = range.getDisplayValues();

      for (let r = 0; r < lastRow; r++) {
        for (let c = 0; c < lastCol; c++) {
          if (!hasText_(formulas[r][c])) {
            continue;
          }
          cells.push({
            row: r + 1,
            col: c + 1,
            a1: toA1_(r + 1, c + 1),
            formulaR1C1: formulas[r][c],
            rawValue: values[r][c],
            displayValue: displayValues[r][c],
            point: 0
          });
        }
      }
    }

    const charts = analyzeCharts_(sheet);
    if (cells.length > 0 || charts.length > 0) {
      sheets.push({
        name: sheet.getName(),
        rowCount: Math.max(1, lastRow),
        colCount: Math.max(1, lastCol),
        cells: cells,
        charts: charts
      });
      totalCells += cells.length;
      totalCharts += charts.length;
    }
  });

  let formulaScore = settings.totalScore;
  let chartScore = 0;
  if (totalCells > 0 && totalCharts > 0) {
    chartScore = settings.totalScore * settings.chartScoreRatio;
    formulaScore = settings.totalScore - chartScore;
  } else if (totalCells === 0 && totalCharts > 0) {
    formulaScore = 0;
    chartScore = settings.totalScore;
  }

  const pointPerCell = totalCells > 0 ? formulaScore / totalCells : 0;
  const pointPerChart = totalCharts > 0 ? chartScore / totalCharts : 0;
  sheets.forEach(function(sheetSpec) {
    sheetSpec.cells.forEach(function(cell) {
      cell.point = pointPerCell;
    });
    sheetSpec.charts.forEach(function(chart) {
      chart.point = pointPerChart;
    });
  });

  return {
    sheets: sheets,
    totalCells: totalCells,
    totalCharts: totalCharts,
    pointPerCell: pointPerCell,
    pointPerChart: pointPerChart,
    formulaScore: formulaScore,
    chartScore: chartScore,
    totalScore: settings.totalScore
  };
}

function analyzeCharts_(sheet) {
  return sheet.getCharts().map(function(chart, index) {
    const options = chart.getOptions();
    return {
      index: index + 1,
      type: String(chart.modify().getChartType()),
      ranges: chart.getRanges().map(function(range) {
        return {
          rowCount: range.getNumRows(),
          colCount: range.getNumColumns(),
          data: range.getDisplayValues()
        };
      }),
      title: chartOptionText_(options, 'title'),
      horizontalAxisTitle: chartOptionText_(options, 'hAxis.title'),
      verticalAxisTitle: chartOptionText_(options, 'vAxis.title'),
      legendPosition: chartOptionText_(options, 'legend.position'),
      numHeaders: chart.getNumHeaders(),
      transpose: chart.getTransposeRowsAndColumns(),
      mergeStrategy: String(chart.getMergeStrategy()),
      point: 0
    };
  });
}

function chartOptionText_(options, key) {
  try {
    const value = options.get(key);
    if (value === null || typeof value === 'undefined') {
      return '';
    }
    return typeof value === 'object' ? JSON.stringify(value) : String(value).trim();
  } catch (error) {
    return '';
  }
}

function collectSubmissionFiles_(rootFolder, recursive, excludedIds) {
  const found = [];
  const seen = {};

  function visit(folder) {
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      const file = files.next();
      const id = file.getId();
      if (seen[id] || excludedIds[id]) {
        continue;
      }
      seen[id] = true;
      if (String(file.getDescription() || '').indexOf(AUTO_GRADER.resultMarker) !== -1) {
        continue;
      }
      found.push(file);
    }

    if (!recursive) {
      return;
    }
    const folders = folder.getFolders();
    while (folders.hasNext()) {
      visit(folders.next());
    }
  }

  visit(rootFolder);
  found.sort(function(a, b) {
    return a.getName().localeCompare(b.getName());
  });
  return found;
}

function gradeSubmissionFiles_(files, spec, settings) {
  const summaries = [];
  const details = [];

  files.forEach(function(file) {
    const summary = createSummary_(file, spec);
    try {
      const submission = SpreadsheetApp.openById(file.getId());
      spec.sheets.forEach(function(sheetSpec) {
        gradeSheet_(submission, file, sheetSpec, settings, summary, details);
      });
    } catch (error) {
      summary.errors.push(error.message);
    }
    summary.score = roundScore_(summary.score);
    summary.rate = spec.totalScore > 0 ? summary.score / spec.totalScore : '';
    summaries.push(summary);
  });

  return {
    summaries: summaries,
    details: details
  };
}

function createSummary_(file, spec) {
  return {
    fileName: file.getName(),
    fileUrl: file.getUrl(),
    score: 0,
    maxScore: spec.totalScore,
    rate: '',
    targetCells: spec.totalCells,
    correctCells: 0,
    targetCharts: spec.totalCharts,
    correctCharts: 0,
    formulaMatches: 0,
    resultMatches: 0,
    directInputs: 0,
    blanks: 0,
    mismatches: 0,
    chartMismatches: 0,
    missingSheets: [],
    errors: []
  };
}

function gradeSheet_(submission, file, sheetSpec, settings, summary, details) {
  const sheet = submission.getSheetByName(sheetSpec.name);
  if (!sheet) {
    summary.missingSheets.push(sheetSpec.name);
    summary.mismatches += sheetSpec.cells.length;
    summary.chartMismatches += sheetSpec.charts.length;
    sheetSpec.cells.forEach(function(expected) {
      if (settings.outputDetails) {
        details.push(detailRow_(file, sheetSpec.name, expected, emptyActualCell_(), {
          code: 'MISSING_SHEET',
          label: 'シートなし',
          correct: false,
          detail: '提出ファイルに同名シートがありません。'
        }));
      }
    });
    sheetSpec.charts.forEach(function(expectedChart) {
      if (settings.outputDetails) {
        details.push(chartDetailRow_(file, sheetSpec.name, expectedChart, null, {
          label: 'シートなし',
          correct: false,
          detail: '提出ファイルに同名シートがありません。'
        }));
      }
    });
    return;
  }

  const grid = readSubmissionGrid_(sheet, sheetSpec.rowCount, sheetSpec.colCount);
  sheetSpec.cells.forEach(function(expected) {
    const actual = getGridCell_(grid, expected.row, expected.col);
    const judgement = judgeFormulaCell_(expected, actual, settings);
    if (judgement.correct) {
      summary.correctCells += 1;
      summary.score += expected.point;
    }
    countJudgement_(summary, judgement.code);
    if (settings.outputDetails) {
      details.push(detailRow_(file, sheetSpec.name, expected, actual, judgement));
    }
  });
  gradeCharts_(sheet, file, sheetSpec, settings, summary, details);
}

function gradeCharts_(sheet, file, sheetSpec, settings, summary, details) {
  const actualCharts = analyzeCharts_(sheet);
  const unused = actualCharts.map(function(chart, index) {
    return { chart: chart, index: index };
  });

  sheetSpec.charts.forEach(function(expectedChart) {
    let bestPosition = -1;
    let bestComparison = null;
    unused.forEach(function(candidate, position) {
      const comparison = compareCharts_(expectedChart, candidate.chart);
      if (!bestComparison || comparison.matchCount > bestComparison.matchCount) {
        bestPosition = position;
        bestComparison = comparison;
      }
    });

    let actualChart = null;
    let judgement;
    if (bestPosition === -1) {
      judgement = {
        label: 'グラフなし',
        correct: false,
        detail: '対応する提出グラフがありません。'
      };
    } else {
      actualChart = unused[bestPosition].chart;
      unused.splice(bestPosition, 1);
      judgement = {
        label: bestComparison.correct ? 'グラフ一致' : 'グラフ不一致',
        correct: bestComparison.correct,
        detail: bestComparison.detail
      };
    }

    if (judgement.correct) {
      summary.correctCharts += 1;
      summary.score += expectedChart.point;
    } else {
      summary.chartMismatches += 1;
    }
    if (settings.outputDetails) {
      details.push(chartDetailRow_(file, sheetSpec.name, expectedChart, actualChart, judgement));
    }
  });
}

function compareCharts_(expected, actual) {
  const checks = [
    ['種類', expected.type === actual.type],
    ['データ範囲', chartRangesEqual_(expected.ranges, actual.ranges)],
    ['タイトル', expected.title === actual.title],
    ['横軸タイトル', expected.horizontalAxisTitle === actual.horizontalAxisTitle],
    ['縦軸タイトル', expected.verticalAxisTitle === actual.verticalAxisTitle],
    ['凡例位置', expected.legendPosition === actual.legendPosition],
    ['ヘッダー数', expected.numHeaders === actual.numHeaders],
    ['行列転置', expected.transpose === actual.transpose],
    ['範囲結合方法', expected.mergeStrategy === actual.mergeStrategy]
  ];
  const failures = checks.filter(function(check) { return !check[1]; }).map(function(check) { return check[0]; });
  return {
    correct: failures.length === 0,
    matchCount: checks.length - failures.length,
    detail: failures.length === 0 ? 'グラフ構造が模範解答と一致しました。' : '不一致項目: ' + failures.join(', ')
  };
}

function chartRangesEqual_(expectedRanges, actualRanges) {
  if (expectedRanges.length !== actualRanges.length) {
    return false;
  }
  for (let i = 0; i < expectedRanges.length; i++) {
    const expected = expectedRanges[i];
    const actual = actualRanges[i];
    if (expected.rowCount !== actual.rowCount || expected.colCount !== actual.colCount) {
      return false;
    }
    if (JSON.stringify(expected.data) !== JSON.stringify(actual.data)) {
      return false;
    }
  }
  return true;
}

function chartDetailRow_(file, sheetName, expected, actual, judgement) {
  return [
    file.getName(),
    file.getUrl(),
    sheetName,
    'グラフ' + expected.index,
    judgement.label,
    judgement.correct,
    expected.point,
    judgement.correct ? expected.point : 0,
    '',
    '',
    chartDescription_(expected),
    actual ? chartDescription_(actual) : '',
    judgement.detail
  ];
}

function chartDescription_(chart) {
  const rangeShapes = chart.ranges.map(function(range) {
    return range.rowCount + 'x' + range.colCount;
  }).join(',');
  return [
    '種類=' + chart.type,
    '範囲=' + rangeShapes,
    'タイトル=' + chart.title,
    '横軸=' + chart.horizontalAxisTitle,
    '縦軸=' + chart.verticalAxisTitle,
    '凡例=' + chart.legendPosition
  ].join('; ');
}

function readSubmissionGrid_(sheet, expectedRows, expectedCols) {
  const rowCount = Math.max(1, Math.min(expectedRows, sheet.getMaxRows()));
  const colCount = Math.max(1, Math.min(expectedCols, sheet.getMaxColumns()));
  const range = sheet.getRange(1, 1, rowCount, colCount);
  return {
    rowCount: rowCount,
    colCount: colCount,
    values: range.getValues(),
    displayValues: range.getDisplayValues(),
    formulasR1C1: range.getFormulasR1C1()
  };
}

function getGridCell_(grid, row, col) {
  const r = row - 1;
  const c = col - 1;
  if (r < 0 || c < 0 || r >= grid.rowCount || c >= grid.colCount) {
    return emptyActualCell_();
  }
  return {
    rawValue: grid.values[r][c],
    displayValue: grid.displayValues[r][c],
    formulaR1C1: grid.formulasR1C1[r][c]
  };
}

function emptyActualCell_() {
  return {
    rawValue: '',
    displayValue: '',
    formulaR1C1: ''
  };
}

function judgeFormulaCell_(expected, actual, settings) {
  const hasFormula = hasText_(actual.formulaR1C1);
  const isBlank = !hasFormula && isBlankValue_(actual.rawValue, actual.displayValue);
  if (isBlank) {
    return {
      code: 'BLANK',
      label: '空欄',
      correct: false,
      detail: '提出セルが空欄です。'
    };
  }

  if (hasFormula && formulasEqual_(expected.formulaR1C1, actual.formulaR1C1)) {
    return {
      code: 'FORMULA_MATCH',
      label: '数式一致',
      correct: true,
      detail: 'R1C1形式の数式が一致しました。'
    };
  }

  const resultMatches = valuesEqual_(
    expected.rawValue,
    expected.displayValue,
    actual.rawValue,
    actual.displayValue,
    settings.numericTolerance
  );
  if (hasFormula && resultMatches) {
    return {
      code: 'RESULT_MATCH',
      label: '計算結果一致（数式違い）',
      correct: settings.treatResultMatchAsCorrect,
      detail: '数式は異なりますが計算結果が一致しました。'
    };
  }
  if (!hasFormula && resultMatches) {
    return {
      code: 'DIRECT_INPUT',
      label: '直打ち',
      correct: settings.treatDirectInputAsCorrect,
      detail: '数式ではなく結果が直接入力されています。'
    };
  }
  return {
    code: 'MISMATCH',
    label: '不一致',
    correct: false,
    detail: '数式と計算結果のどちらも模範解答と一致しません。'
  };
}

function countJudgement_(summary, code) {
  if (code === 'FORMULA_MATCH') {
    summary.formulaMatches += 1;
  } else if (code === 'RESULT_MATCH') {
    summary.resultMatches += 1;
  } else if (code === 'DIRECT_INPUT') {
    summary.directInputs += 1;
  } else if (code === 'BLANK') {
    summary.blanks += 1;
  } else {
    summary.mismatches += 1;
  }
}

function detailRow_(file, sheetName, expected, actual, judgement) {
  return [
    file.getName(),
    file.getUrl(),
    sheetName,
    expected.a1,
    judgement.label,
    judgement.correct,
    expected.point,
    judgement.correct ? expected.point : 0,
    expected.formulaR1C1,
    actual.formulaR1C1,
    expected.displayValue,
    actual.displayValue,
    judgement.detail
  ];
}

function createResultSpreadsheet_(folder, answerSpreadsheet, spec, grading, settings) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const name = (settings.resultFileName || AUTO_GRADER.defaults.resultFileName) + '_' + timestamp;
  const result = SpreadsheetApp.create(name);
  const resultFile = DriveApp.getFileById(result.getId());
  resultFile.setDescription(AUTO_GRADER.resultMarker + ' version=' + AUTO_GRADER.version);
  resultFile.moveTo(folder);

  const summarySheet = result.getSheets()[0];
  summarySheet.setName(AUTO_GRADER.resultSheets.summary);
  writeSummarySheet_(summarySheet, grading.summaries);

  if (settings.outputDetails) {
    const detailSheet = result.insertSheet(AUTO_GRADER.resultSheets.details);
    writeDetailSheet_(detailSheet, grading.details);
  }

  const infoSheet = result.insertSheet(AUTO_GRADER.resultSheets.info);
  writeInfoSheet_(infoSheet, folder, answerSpreadsheet, spec, grading, settings);
  SpreadsheetApp.flush();
  return result;
}

function writeSummarySheet_(sheet, summaries) {
  const headers = [
    'ファイル名',
    '得点',
    '満点',
    '得点率',
    '採点対象セル数',
    '正答数',
    '採点対象グラフ数',
    'グラフ一致',
    'グラフ不一致',
    '数式一致',
    '計算結果一致',
    '直打ち',
    '空欄',
    '不一致',
    '不足シート',
    'エラー',
    '提出ファイルURL'
  ];
  const rows = summaries.map(function(summary) {
    return [
      summary.fileName,
      summary.score,
      summary.maxScore,
      summary.rate,
      summary.targetCells,
      summary.correctCells,
      summary.targetCharts,
      summary.correctCharts,
      summary.chartMismatches,
      summary.formulaMatches,
      summary.resultMatches,
      summary.directInputs,
      summary.blanks,
      summary.mismatches,
      summary.missingSheets.join(', '),
      summary.errors.join(' / '),
      summary.fileUrl
    ];
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 4, rows.length, 1).setNumberFormat('0.0%');
    sheet.getRange(2, 2, rows.length, 2).setNumberFormat('0.00');
  }
  formatOutputSheet_(sheet, headers.length, '#cfe2f3');
}

function writeDetailSheet_(sheet, rows) {
  const headers = [
    'ファイル名',
    '提出ファイルURL',
    'シート名',
    'セル',
    '判定',
    '正答扱い',
    '配点',
    '得点',
    '期待数式R1C1',
    '提出数式R1C1',
    '期待値',
    '提出値',
    '詳細'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 7, rows.length, 2).setNumberFormat('0.000');
  }
  formatOutputSheet_(sheet, headers.length, '#d9ead3');
}

function writeInfoSheet_(sheet, folder, answerSpreadsheet, spec, grading, settings) {
  const rows = [
    ['項目', '値'],
    ['実行日時', new Date()],
    ['プログラムバージョン', AUTO_GRADER.version],
    ['提出物フォルダ', folder.getName()],
    ['提出物フォルダURL', settings.folderUrl],
    ['模範解答', answerSpreadsheet.getName()],
    ['模範解答URL', settings.answerUrl],
    ['採点対象シート数', spec.sheets.length],
    ['採点対象数式セル数', spec.totalCells],
    ['採点対象グラフ数', spec.totalCharts],
    ['総点', spec.totalScore],
    ['数式配点', spec.formulaScore],
    ['グラフ配点', spec.chartScore],
    ['1セル配点', spec.pointPerCell],
    ['1グラフ配点', spec.pointPerChart],
    ['採点ファイル数', grading.summaries.length],
    ['対象シート名', spec.sheets.map(function(item) { return item.name; }).join(', ')]
  ];
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  formatOutputSheet_(sheet, 2, '#fff2cc');
}

function writeSpecSheet_(controller, answerSpreadsheet, spec) {
  const sheet = ensureSheet_(controller, AUTO_GRADER.specSheet);
  sheet.clear();
  const headers = ['模範解答', 'シート名', '対象種別', 'セル/グラフ', '期待数式R1C1', '期待値/グラフ条件', '配点'];
  const rows = [];
  spec.sheets.forEach(function(sheetSpec) {
    sheetSpec.cells.forEach(function(cell) {
      rows.push([
        answerSpreadsheet.getName(),
        sheetSpec.name,
        '数式',
        cell.a1,
        cell.formulaR1C1,
        cell.displayValue,
        cell.point
      ]);
    });
    sheetSpec.charts.forEach(function(chart) {
      rows.push([
        answerSpreadsheet.getName(),
        sheetSpec.name,
        'グラフ',
        'グラフ' + chart.index,
        '',
        chartDescription_(chart),
        chart.point
      ]);
    });
  });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 7, rows.length, 1).setNumberFormat('0.000');
  }
  formatOutputSheet_(sheet, headers.length, '#fff2cc');
}

function formatOutputSheet_(sheet, columnCount, headerColor) {
  sheet.getRange(1, 1, 1, columnCount).setFontWeight('bold').setBackground(headerColor);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, columnCount);
}

function formulasEqual_(expectedFormula, actualFormula) {
  return normalizeFormulaR1C1_(expectedFormula) === normalizeFormulaR1C1_(actualFormula);
}

function normalizeFormulaR1C1_(formula) {
  const source = String(formula || '').trim();
  let output = '';
  let inString = false;
  let inSheetName = false;

  for (let i = 0; i < source.length; i++) {
    const character = source.charAt(i);
    if (!inSheetName && character === '"') {
      output += character;
      if (inString && source.charAt(i + 1) === '"') {
        i++;
        output += '"';
        continue;
      }
      inString = !inString;
      continue;
    }
    if (!inString && character === "'") {
      output += character;
      if (inSheetName && source.charAt(i + 1) === "'") {
        i++;
        output += "'";
        continue;
      }
      inSheetName = !inSheetName;
      continue;
    }
    if (!inString && !inSheetName && /\s/.test(character)) {
      continue;
    }
    output += character;
  }
  return output;
}

function valuesEqual_(expectedRaw, expectedDisplay, actualRaw, actualDisplay, tolerance) {
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

function extractDriveId_(urlOrId, pathMarker) {
  const value = String(urlOrId || '').trim();
  const markerPattern = new RegExp(pathMarker.replace('/', '\\/') + '\\/([A-Za-z0-9_-]+)');
  const marked = value.match(markerPattern);
  if (marked) {
    return marked[1];
  }
  const idOnly = value.match(/^[A-Za-z0-9_-]{20,}$/);
  if (idOnly) {
    return idOnly[0];
  }
  throw new Error('URLからIDを取得できません: ' + value);
}

function parseNameList_(value) {
  return String(value || '')
    .split(',')
    .map(function(item) { return item.trim(); })
    .filter(function(item) { return item !== ''; });
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

function toA1_(row, col) {
  let label = '';
  let number = col;
  while (number > 0) {
    const mod = (number - 1) % 26;
    label = String.fromCharCode(65 + mod) + label;
    number = Math.floor((number - mod - 1) / 26);
  }
  return label + row;
}

function toBool_(value, defaultValue) {
  if (typeof value === 'boolean') {
    return value;
  }
  if (value === null || typeof value === 'undefined' || value === '') {
    return defaultValue;
  }
  const text = String(value).trim().toLowerCase();
  if (['true', '1', 'yes', 'on', 'はい'].indexOf(text) !== -1) {
    return true;
  }
  if (['false', '0', 'no', 'off', 'いいえ'].indexOf(text) !== -1) {
    return false;
  }
  return defaultValue;
}

function toPositiveNumber_(value, defaultValue) {
  const parsed = parseFloat(value);
  return isFinite(parsed) && parsed > 0 ? parsed : defaultValue;
}

function toRatio_(value, defaultValue) {
  const parsed = parseFloat(value);
  return isFinite(parsed) && parsed >= 0 && parsed <= 1 ? parsed : defaultValue;
}

function toNumber_(value, defaultValue) {
  const parsed = parseFloat(value);
  return isFinite(parsed) ? parsed : defaultValue;
}

function roundScore_(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function ensureSheet_(spreadsheet, name) {
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function notify_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    Logger.log(message);
  }
}
