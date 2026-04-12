/**
 * portfolio_formating.gs
 *
 * Post-processing / enrichment слой для уже загруженных листов портфеля.
 *
 * Текущая итерация:
 * - централизованное чтение Rules;
 * - полная реализация для Shares;
 * - bond analytics перенесена из update_bonds.gs сюда;
 * - служебные колонки с префиксом "A:" остаются post-process слоем.
 *
 * Принципы:
 * - не вызывает API;
 * - не меняет другие файлы;
 * - не использует clear();
 * - не затирает user/business данные вне целевых колонок;
 * - безопасен при повторном запуске.
 *
 * Rules:
 * GENERAL
 * B2 = enableFormatting
 * B3 = useServiceColumns
 * B4 = colorPl
 * B5 = showWarnings
 *
 * SHARES
 * E2  = strongProfitPct
 * E3  = strongLossPct
 * E4  = roeGood
 * E5  = roeBad
 * E6  = roaGood
 * E7  = roaBad
 * E8  = roicGood
 * E9  = roicBad
 * E10 = peCheap
 * E11 = peExpensive
 * E12 = psCheap
 * E13 = psExpensive
 * E14 = pbCheap
 * E15 = pbExpensive
 * E16 = evEbitdaCheap
 * E17 = evEbitdaExpensive
 * E18 = debtToEquityHigh
 * E19 = netDebtToEbitdaHigh
 * E20 = betaHigh
 * E21 = freeFloatLow
 * E22 = near52wHigh
 * E23 = near52wLow
 * E24 = marketCapSmall
 * E25 = marketCapLarge
 *
 * BONDS
 * I2  = yieldLowPct
 * I3  = yieldHighPct
 * I4  = yieldVeryHighPct
 * I5  = strongProfitPct
 * I6  = strongLossPct
 * I7  = manualRiskMedium
 * I8  = manualRiskHigh
 * I9  = shortMaturityYears
 * I10 = longMaturityYears
 * I11 = veryLongMaturityYears
 * I12 = nextCouponSoonDays
 * I13 = flagFloatingCoupon
 * I14 = flagVariableCoupon
 * I15 = riskDrawdownPct
 *
 * BONDS sector risk table
 * H18:H = sector
 * I18:I = score
 *
 * Дополнительно поддерживается legacy fallback:
 * - named range RISK_DD для riskDrawdownPct
 * - named range RISK_SECTORS для sector risk table
 */

const PORTFOLIO_FORMATING_SHEETS = {
  rules: 'Rules',
  shares: 'Shares',
  bonds: 'Bonds',
  funds: 'Funds',
  options: 'Options'
};

const PORTFOLIO_FORMATING_COLORS = {
  text: '#1f2937',
  mutedText: '#6b7280',
  neutralBg: '#ffffff',
  neutralSoft: '#f8fafc',
  headerBg: '#e8eefc',
  headerText: '#1e3a8a',
  green: '#e6f4ea',
  greenStrong: '#c8e6c9',
  red: '#fce8e6',
  redStrong: '#f4c7c3',
  amber: '#fff4ce',
  orange: '#fde7c8',
  blue: '#e8f0fe',
  purple: '#f3e8ff',
  grey: '#f1f3f4'
};

const PORTFOLIO_FORMATING_SHARES_SERVICE = {
  quality: 'A: Quality',
  valuation: 'A: Valuation',
  risk: 'A: Risk',
  position52w: 'A: 52w Position',
  flags: 'A: Flags'
};

const PORTFOLIO_FORMATING_SHARES_SERVICE_ORDER = [
  PORTFOLIO_FORMATING_SHARES_SERVICE.quality,
  PORTFOLIO_FORMATING_SHARES_SERVICE.valuation,
  PORTFOLIO_FORMATING_SHARES_SERVICE.risk,
  PORTFOLIO_FORMATING_SHARES_SERVICE.position52w,
  PORTFOLIO_FORMATING_SHARES_SERVICE.flags
];

const PORTFOLIO_FORMATING_BONDS_SERVICE = {
  yield: 'A: Yield',
  bondRisk: 'A: Bond Risk',
  maturity: 'A: Maturity',
  coupon: 'A: Coupon',
  flags: 'A: Flags'
};

const PORTFOLIO_FORMATING_BONDS_SERVICE_ORDER = [
  PORTFOLIO_FORMATING_BONDS_SERVICE.yield,
  PORTFOLIO_FORMATING_BONDS_SERVICE.bondRisk,
  PORTFOLIO_FORMATING_BONDS_SERVICE.maturity,
  PORTFOLIO_FORMATING_BONDS_SERVICE.coupon,
  PORTFOLIO_FORMATING_BONDS_SERVICE.flags
];

const PORTFOLIO_FORMATING_BONDS_DERIVED_ORDER = [
  'Риск (ручн.)',
  'Инвестировано',
  'Рыночная стоимость',
  'P/L (руб)',
  'P/L (%)',
  'Доходность купонная годовая (прибл.)'
];

const PORTFOLIO_FORMATING_RULES_LAYOUT = {
  sheetName: PORTFOLIO_FORMATING_SHEETS.rules,

  general: {
    anchor: 'A1',
    cells: {
      enableFormatting: { cell: 'B2', type: 'bool', defaultValue: true },
      useServiceColumns: { cell: 'B3', type: 'bool', defaultValue: true },
      colorPl: { cell: 'B4', type: 'bool', defaultValue: true },
      showWarnings: { cell: 'B5', type: 'bool', defaultValue: true }
    }
  },

  shares: {
    anchor: 'D1',
    cells: {
      strongProfitPct: { cell: 'E2', type: 'fraction', defaultValue: null },
      strongLossPct: { cell: 'E3', type: 'fraction', defaultValue: null },
      roeGood: { cell: 'E4', type: 'number', defaultValue: null },
      roeBad: { cell: 'E5', type: 'number', defaultValue: null },
      roaGood: { cell: 'E6', type: 'number', defaultValue: null },
      roaBad: { cell: 'E7', type: 'number', defaultValue: null },
      roicGood: { cell: 'E8', type: 'number', defaultValue: null },
      roicBad: { cell: 'E9', type: 'number', defaultValue: null },
      peCheap: { cell: 'E10', type: 'number', defaultValue: null },
      peExpensive: { cell: 'E11', type: 'number', defaultValue: null },
      psCheap: { cell: 'E12', type: 'number', defaultValue: null },
      psExpensive: { cell: 'E13', type: 'number', defaultValue: null },
      pbCheap: { cell: 'E14', type: 'number', defaultValue: null },
      pbExpensive: { cell: 'E15', type: 'number', defaultValue: null },
      evEbitdaCheap: { cell: 'E16', type: 'number', defaultValue: null },
      evEbitdaExpensive: { cell: 'E17', type: 'number', defaultValue: null },
      debtToEquityHigh: { cell: 'E18', type: 'number', defaultValue: null },
      netDebtToEbitdaHigh: { cell: 'E19', type: 'number', defaultValue: null },
      betaHigh: { cell: 'E20', type: 'number', defaultValue: null },
      freeFloatLow: { cell: 'E21', type: 'number', defaultValue: null },
      near52wHigh: { cell: 'E22', type: 'fraction', defaultValue: null },
      near52wLow: { cell: 'E23', type: 'fraction', defaultValue: null },
      marketCapSmall: { cell: 'E24', type: 'number', defaultValue: null },
      marketCapLarge: { cell: 'E25', type: 'number', defaultValue: null }
    }
  },

  bonds: {
    anchor: 'H1',
    cells: {
      yieldLowPct:           { cell: 'I2',  type: 'number', defaultValue: 12, useDefaultWhenBlank: false },
      yieldHighPct:          { cell: 'I3',  type: 'number', defaultValue: 18, useDefaultWhenBlank: false },
      yieldVeryHighPct:      { cell: 'I4',  type: 'number', defaultValue: 22, useDefaultWhenBlank: false },
      strongProfitPct:       { cell: 'I5',  type: 'number', defaultValue: 2,  useDefaultWhenBlank: false },
      strongLossPct:         { cell: 'I6',  type: 'number', defaultValue: 3,  useDefaultWhenBlank: false },
      manualRiskMedium:      { cell: 'I7',  type: 'number', defaultValue: 2,  useDefaultWhenBlank: false },
      manualRiskHigh:        { cell: 'I8',  type: 'number', defaultValue: 4,  useDefaultWhenBlank: false },
      shortMaturityYears:    { cell: 'I9',  type: 'number', defaultValue: 1,  useDefaultWhenBlank: false },
      longMaturityYears:     { cell: 'I10', type: 'number', defaultValue: 5,  useDefaultWhenBlank: false },
      veryLongMaturityYears: { cell: 'I11', type: 'number', defaultValue: 8,  useDefaultWhenBlank: false },
      nextCouponSoonDays:    { cell: 'I12', type: 'number', defaultValue: 14, useDefaultWhenBlank: false },
      flagFloatingCoupon:    { cell: 'I13', type: 'bool',   defaultValue: true, useDefaultWhenBlank: false },
      flagVariableCoupon:    { cell: 'I14', type: 'bool',   defaultValue: true, useDefaultWhenBlank: false },
      riskDrawdownPct:       { cell: 'I15', type: 'number', defaultValue: null, useDefaultWhenBlank: false }
    },
    tables: {
      sectorRiskScores: {
        startRow: 18,
        keyColumn: 'H',
        valueColumn: 'I',
        stopAfterBlankRows: 3,
        fallbackNamedRange: 'RISK_SECTORS'
      }
    }
  },

  funds: {
    anchor: 'K1',
    cells: {}
  },

  options: {
    anchor: 'N1',
    cells: {}
  }
};

// ==========================================================
// Public API
// ==========================================================

function runPortfolioFormating_() {
  return runPortfolioFormating();
}

function runPortfolioFormating() {
  var rules = readPortfolioFormatingRules_();

  if (!rules.general.enableFormatting) {
    Logger.log('[portfolio_formating] disabled by Rules!B2');
    return {
      ok: true,
      skipped: true,
      reason: 'Formatting disabled by Rules!B2'
    };
  }

  var summary = {
    ok: true,
    skipped: false,
    rulesSheetPresent: !!SpreadsheetApp.getActive().getSheetByName(PORTFOLIO_FORMATING_SHEETS.rules),
    shares: null,
    bonds: null,
    funds: { skipped: true, reason: 'Not implemented yet' },
    options: { skipped: true, reason: 'Not implemented yet' }
  };

  summary.shares = runSharesFormating(rules);
  summary.bonds = runBondsFormating(rules);

  return summary;
}

function runSharesFormating_(preloadedRules) {
  return runSharesFormating(preloadedRules);
}

function runSharesFormating(preloadedRules) {
  var rules = preloadedRules || readPortfolioFormatingRules_();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PORTFOLIO_FORMATING_SHEETS.shares);

  var summary = {
    ok: true,
    sheet: PORTFOLIO_FORMATING_SHEETS.shares,
    missingSheet: false,
    processedRows: 0,
    serviceColumnsAdded: 0,
    serviceColumnsPresent: 0,
    serviceColumnsUpdated: false
  };

  if (!sh) {
    summary.missingSheet = true;
    return summary;
  }

  var initialLastCol = Math.max(1, sh.getLastColumn());
  var headers = sh.getRange(1, 1, 1, initialLastCol).getValues()[0];
  var headerMap = pfBuildHeaderMap_(headers);

  var serviceSetup = pfEnsureServiceColumns_(
    sh,
    headerMap,
    PORTFOLIO_FORMATING_SHARES_SERVICE_ORDER,
    !!rules.general.useServiceColumns
  );

  headers = serviceSetup.headers;
  headerMap = serviceSetup.headerMap;
  summary.serviceColumnsAdded = serviceSetup.added;
  summary.serviceColumnsPresent = pfCountExistingHeaders_(headerMap, PORTFOLIO_FORMATING_SHARES_SERVICE_ORDER);

  pfFormatSharesServiceHeaders_(sh, headerMap);

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return summary;

  var lastCol = sh.getLastColumn();
  var numRows = lastRow - 1;
  var rows = sh.getRange(2, 1, numRows, lastCol).getValues();

  var sourceStyleHeaders = pfExistingHeaders_(headerMap, [
    'P/L (руб)', 'P/L (%)', 'ROE', 'ROA', 'ROIC', 'EBITDA TTM', 'Чистая прибыль TTM',
    'P/E TTM', 'P/S TTM', 'P/B TTM', 'EV/EBITDA', 'Beta', 'Debt/Equity',
    'NetDebt/EBITDA', 'Free Float', 'Заблокирован (TCS)', 'Шорт доступен',
    'Текущая цена', 'Капитализация'
  ]);

  var sourceStyles = {};
  sourceStyleHeaders.forEach(function(header) {
    sourceStyles[header] = pfCreateStyleMatrix_(numRows, PORTFOLIO_FORMATING_COLORS.neutralBg, PORTFOLIO_FORMATING_COLORS.text, 'normal');
  });

  var serviceValues = {};
  var serviceStyles = {};
  PORTFOLIO_FORMATING_SHARES_SERVICE_ORDER.forEach(function(header) {
    if (!pfHasHeader_(headerMap, header)) return;
    serviceValues[header] = [];
    serviceStyles[header] = pfCreateStyleMatrix_(numRows, PORTFOLIO_FORMATING_COLORS.neutralBg, PORTFOLIO_FORMATING_COLORS.text, 'normal');
  });

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var analysis = pfAnalyzeSharesRow_(row, headerMap, rules);

    Object.keys(analysis.sourceStyles).forEach(function(header) {
      if (!sourceStyles[header]) return;
      pfSetMatrixCell_(
        sourceStyles[header],
        i,
        analysis.sourceStyles[header].bg,
        analysis.sourceStyles[header].fontColor,
        analysis.sourceStyles[header].fontWeight
      );
    });

    PORTFOLIO_FORMATING_SHARES_SERVICE_ORDER.forEach(function(serviceHeader) {
      if (!serviceValues[serviceHeader]) return;

      if (!rules.general.useServiceColumns) {
        serviceValues[serviceHeader].push(['']);
        return;
      }

      var serviceKey = pfSharesServiceKeyByHeader_(serviceHeader);
      var serviceValue = analysis.service[serviceKey] || '';
      var serviceStyle = analysis.serviceStyles[serviceKey] || pfStyle_(PORTFOLIO_FORMATING_COLORS.neutralBg, PORTFOLIO_FORMATING_COLORS.text, 'normal');

      serviceValues[serviceHeader].push([serviceValue]);
      pfSetMatrixCell_(serviceStyles[serviceHeader], i, serviceStyle.bg, serviceStyle.fontColor, serviceStyle.fontWeight);
    });
  }

  Object.keys(sourceStyles).forEach(function(header) {
    var col = headerMap[header];
    if (!col) return;
    pfApplyColumnMatrix_(sh, col, sourceStyles[header], numRows);
  });

  Object.keys(serviceValues).forEach(function(header) {
    var col = headerMap[header];
    if (!col) return;
    sh.getRange(2, col, numRows, 1).setValues(serviceValues[header]);
    pfApplyColumnMatrix_(sh, col, serviceStyles[header], numRows);
  });

  pfFormatSharesServiceData_(sh, headerMap, numRows);

  summary.processedRows = numRows;
  summary.serviceColumnsUpdated = true;
  return summary;
}

function runBondsFormating_(preloadedRules) {
  return runBondsFormating(preloadedRules);
}

function runBondsFormating(preloadedRules) {
  var rules = preloadedRules || readPortfolioFormatingRules_();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PORTFOLIO_FORMATING_SHEETS.bonds);

  var summary = {
    ok: true,
    sheet: PORTFOLIO_FORMATING_SHEETS.bonds,
    missingSheet: false,
    processedRows: 0,
    serviceColumnsAdded: 0,
    serviceColumnsPresent: 0,
    serviceColumnsUpdated: false,
    derivedColumnsUpdated: false
  };

  if (!sh) {
    summary.missingSheet = true;
    return summary;
  }

  var initialLastCol = Math.max(1, sh.getLastColumn());
  var headers = sh.getRange(1, 1, 1, initialLastCol).getValues()[0];
  var headerMap = pfBuildHeaderMap_(headers);

  var serviceSetup = pfEnsureServiceColumns_(
    sh,
    headerMap,
    PORTFOLIO_FORMATING_BONDS_SERVICE_ORDER,
    !!rules.general.useServiceColumns
  );

  headers = serviceSetup.headers;
  headerMap = serviceSetup.headerMap;
  summary.serviceColumnsAdded = serviceSetup.added;
  summary.serviceColumnsPresent = pfCountExistingHeaders_(headerMap, PORTFOLIO_FORMATING_BONDS_SERVICE_ORDER);

  pfFormatBondsServiceHeaders_(sh, headerMap);

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return summary;

  var lastCol = sh.getLastColumn();
  var numRows = lastRow - 1;
  var rows = sh.getRange(2, 1, numRows, lastCol).getValues();

  var derivedHeaders = pfExistingHeaders_(headerMap, PORTFOLIO_FORMATING_BONDS_DERIVED_ORDER);
  var derivedValues = {};
  derivedHeaders.forEach(function(header) {
    derivedValues[header] = [];
  });

  var sourceStyleHeaders = pfExistingHeaders_(headerMap, [
    'P/L (руб)',
    'P/L (%)',
    'Доходность купонная годовая (прибл.)',
    'Риск (ручн.)',
    'Риск (уровень TCS)',
    'Класс риска TCS',
    'Дата погашения',
    'Следующий купон',
    'Тип купона (desc)'
  ]);

  var sourceStyles = {};
  sourceStyleHeaders.forEach(function(header) {
    sourceStyles[header] = pfCreateStyleMatrix_(numRows, PORTFOLIO_FORMATING_COLORS.neutralBg, PORTFOLIO_FORMATING_COLORS.text, 'normal');
  });

  var serviceValues = {};
  var serviceStyles = {};
  PORTFOLIO_FORMATING_BONDS_SERVICE_ORDER.forEach(function(header) {
    if (!pfHasHeader_(headerMap, header)) return;
    serviceValues[header] = [];
    serviceStyles[header] = pfCreateStyleMatrix_(numRows, PORTFOLIO_FORMATING_COLORS.neutralBg, PORTFOLIO_FORMATING_COLORS.text, 'normal');
  });

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var derived = pfPopulateBondDerivedRow_(row, headerMap, rules);

    Object.keys(derivedValues).forEach(function(header) {
      var v = Object.prototype.hasOwnProperty.call(derived.values, header) ? derived.values[header] : '';
      var safeValue = pfSheetValue_(v);
      derivedValues[header].push([safeValue]);
      if (pfHasHeader_(headerMap, header)) {
        row[headerMap[header] - 1] = safeValue;
      }
    });

    var analysis = pfAnalyzeBondRow_(row, headerMap, rules);

    Object.keys(analysis.sourceStyles).forEach(function(header) {
      if (!sourceStyles[header]) return;
      pfSetMatrixCell_(
        sourceStyles[header],
        i,
        analysis.sourceStyles[header].bg,
        analysis.sourceStyles[header].fontColor,
        analysis.sourceStyles[header].fontWeight
      );
    });

    PORTFOLIO_FORMATING_BONDS_SERVICE_ORDER.forEach(function(serviceHeader) {
      if (!serviceValues[serviceHeader]) return;

      if (!rules.general.useServiceColumns) {
        serviceValues[serviceHeader].push(['']);
        return;
      }

      var serviceKey = pfBondServiceKeyByHeader_(serviceHeader);
      var serviceValue = analysis.service[serviceKey] || '';
      var serviceStyle = analysis.serviceStyles[serviceKey] || pfStyle_(PORTFOLIO_FORMATING_COLORS.neutralBg, PORTFOLIO_FORMATING_COLORS.text, 'normal');

      serviceValues[serviceHeader].push([serviceValue]);
      pfSetMatrixCell_(serviceStyles[serviceHeader], i, serviceStyle.bg, serviceStyle.fontColor, serviceStyle.fontWeight);
    });
  }

  Object.keys(derivedValues).forEach(function(header) {
    var col = headerMap[header];
    if (!col) return;
    sh.getRange(2, col, numRows, 1).setValues(derivedValues[header]);
  });

  Object.keys(sourceStyles).forEach(function(header) {
    var col = headerMap[header];
    if (!col) return;
    pfApplyColumnMatrix_(sh, col, sourceStyles[header], numRows);
  });

  Object.keys(serviceValues).forEach(function(header) {
    var col = headerMap[header];
    if (!col) return;
    sh.getRange(2, col, numRows, 1).setValues(serviceValues[header]);
    pfApplyColumnMatrix_(sh, col, serviceStyles[header], numRows);
  });

  pfFormatBondsServiceData_(sh, headerMap, numRows);

  summary.processedRows = numRows;
  summary.derivedColumnsUpdated = Object.keys(derivedValues).length > 0;
  summary.serviceColumnsUpdated = true;
  return summary;
}

// ==========================================================
// Rules reader
// ==========================================================

function readPortfolioFormatingRules_() {
  var rules = pfDefaultRules_();
  var sh = SpreadsheetApp.getActive().getSheetByName(PORTFOLIO_FORMATING_RULES_LAYOUT.sheetName);
  if (!sh) return rules;

  rules.general = pfReadRulesBlock_(sh, PORTFOLIO_FORMATING_RULES_LAYOUT.general);
  rules.shares = pfReadRulesBlock_(sh, PORTFOLIO_FORMATING_RULES_LAYOUT.shares);
  rules.bonds = pfReadRulesBlock_(sh, PORTFOLIO_FORMATING_RULES_LAYOUT.bonds);
  rules.funds = pfReadRulesBlock_(sh, PORTFOLIO_FORMATING_RULES_LAYOUT.funds);
  rules.options = pfReadRulesBlock_(sh, PORTFOLIO_FORMATING_RULES_LAYOUT.options);

  pfApplyLegacyBondRulesFallbacks_(rules.bonds);
  return rules;
}

function pfDefaultRules_() {
  var out = {
    general: {},
    shares: {},
    bonds: {},
    funds: {},
    options: {}
  };

  Object.keys(PORTFOLIO_FORMATING_RULES_LAYOUT).forEach(function(blockName) {
    if (blockName === 'sheetName') return;

    var block = PORTFOLIO_FORMATING_RULES_LAYOUT[blockName];
    out[blockName] = out[blockName] || {};

    Object.keys(block.cells || {}).forEach(function(key) {
      out[blockName][key] = block.cells[key].defaultValue;
    });

    Object.keys(block.tables || {}).forEach(function(key) {
      out[blockName][key] = {};
    });
  });

  return out;
}

function pfReadRulesBlock_(sh, blockConfig) {
  var out = {};

  Object.keys(blockConfig.cells || {}).forEach(function(key) {
    var cfg = blockConfig.cells[key];
    var raw = sh.getRange(cfg.cell).getValue();
    out[key] = pfReadTypedRuleValue_(raw, cfg);
  });

  Object.keys(blockConfig.tables || {}).forEach(function(key) {
    out[key] = pfReadRulesTableMap_(sh, blockConfig.tables[key]);
  });

  return out;
}

function pfReadTypedRuleValue_(raw, cfg) {
  if (pfIsEmpty_(raw)) {
    return (cfg && cfg.useDefaultWhenBlank === false) ? null : cfg.defaultValue;
  }

  if (cfg.type === 'bool') {
    return pfParseBool_(raw, cfg.defaultValue);
  }

  if (cfg.type === 'fraction') {
    var frac = pfNormalizeFraction_(raw);
    return frac == null ? cfg.defaultValue : frac;
  }

  if (cfg.type === 'number') {
    var num = pfNormalizeNumber_(raw);
    return num == null ? cfg.defaultValue : num;
  }

  return raw;
}

function pfReadRulesTableMap_(sh, cfg) {
  var map = {};
  if (!cfg) return map;

  var keyCol = pfNormalizeColumnRef_(cfg.keyColumn);
  var valueCol = pfNormalizeColumnRef_(cfg.valueColumn);
  var startRow = Math.max(1, Number(cfg.startRow) || 1);
  var stopAfterBlankRows = Math.max(1, Number(cfg.stopAfterBlankRows) || 3);
  var maxRow = Math.max(startRow, sh.getLastRow());
  if (!keyCol || !valueCol || maxRow < startRow) return map;

  var startCol = Math.min(keyCol, valueCol);
  var width = Math.abs(keyCol - valueCol) + 1;
  var values = sh.getRange(startRow, startCol, maxRow - startRow + 1, width).getValues();
  var blankStreak = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var keyRaw = row[keyCol - startCol];
    var valueRaw = row[valueCol - startCol];
    var keyText = String(keyRaw == null ? '' : keyRaw).trim();
    var valueNum = pfNormalizeNumber_(valueRaw);

    if (!keyText && valueNum == null) {
      blankStreak++;
      if (blankStreak >= stopAfterBlankRows && Object.keys(map).length) break;
      continue;
    }

    blankStreak = 0;
    if (!keyText || valueNum == null) continue;
    map[pfNormalizeRuleMapKey_(keyText)] = valueNum;
  }

  if (!Object.keys(map).length && cfg.fallbackNamedRange) {
    map = pfReadNamedRangeMap_(cfg.fallbackNamedRange);
  }

  return map;
}

function pfApplyLegacyBondRulesFallbacks_(bondRules) {
  if (!bondRules) return;

  if (!pfHasValue_(bondRules.riskDrawdownPct)) {
    var legacyDd = pfReadNamedRangeScalar_('RISK_DD');
    if (legacyDd != null) {
      bondRules.riskDrawdownPct = legacyDd;
    } else if (pfHasValue_(bondRules.strongLossPct)) {
      bondRules.riskDrawdownPct = -Math.abs(Number(bondRules.strongLossPct));
    }
  }

  if (!bondRules.sectorRiskScores || !Object.keys(bondRules.sectorRiskScores).length) {
    bondRules.sectorRiskScores = pfReadNamedRangeMap_('RISK_SECTORS');
  }
}

// ==========================================================
// Shares derived metrics / signals
// ==========================================================

function pfAnalyzeSharesRow_(row, headerMap, rules) {
  var C = PORTFOLIO_FORMATING_COLORS;
  var general = rules.general || {};
  var shareRules = rules.shares || {};

  var out = {
    sourceStyles: {},
    service: {
      quality: '',
      valuation: '',
      risk: '',
      position52w: '',
      flags: ''
    },
    serviceStyles: {
      quality: pfStyle_(C.neutralBg, C.text, 'normal'),
      valuation: pfStyle_(C.neutralBg, C.text, 'normal'),
      risk: pfStyle_(C.neutralBg, C.text, 'normal'),
      position52w: pfStyle_(C.neutralBg, C.text, 'normal'),
      flags: pfStyle_(C.neutralBg, C.mutedText, 'normal')
    }
  };

  var figi = String(pfValueByHeader_(row, headerMap, 'FIGI') || '').trim();
  if (!figi) return out;

  var flags = [];

  var plRub = pfNumberByHeader_(row, headerMap, 'P/L (руб)');
  var plPct = pfNumberByHeader_(row, headerMap, 'P/L (%)');

  if (general.colorPl) {
    var plStyle = null;

    if (plRub != null || plPct != null) {
      var strongProfit = (plPct != null) && pfPositiveThresholdMet_(plPct, shareRules.strongProfitPct);
      var strongLoss = (plPct != null) && pfNegativeThresholdMet_(plPct, shareRules.strongLossPct);

      if ((plRub != null && plRub > 0) || (plPct != null && plPct > 0)) {
        plStyle = strongProfit
          ? pfStyle_(C.greenStrong, C.text, 'bold')
          : pfStyle_(C.green, C.text, 'normal');
      } else if ((plRub != null && plRub < 0) || (plPct != null && plPct < 0)) {
        plStyle = strongLoss
          ? pfStyle_(C.redStrong, C.text, 'bold')
          : pfStyle_(C.red, C.text, 'normal');
      }
    }

    if (plStyle) {
      out.sourceStyles['P/L (руб)'] = plStyle;
      out.sourceStyles['P/L (%)'] = plStyle;
    }
  }

  var qualityHasAnyData = false;
  var qualityPositiveCount = 0;
  var qualityNegativeCount = 0;
  var hardNegativeCount = 0;

  var roe = pfNumberByHeader_(row, headerMap, 'ROE');
  if (roe != null) qualityHasAnyData = true;
  if (pfIsMeaningfulQualityValue_(roe) && (pfHasValue_(shareRules.roeGood) || pfHasValue_(shareRules.roeBad))) {
    if (pfHasValue_(shareRules.roeGood) && roe >= shareRules.roeGood) {
      qualityPositiveCount++;
      out.sourceStyles['ROE'] = pfStyle_(C.green, C.text, 'bold');
    } else if (pfHasValue_(shareRules.roeBad) && roe <= shareRules.roeBad) {
      qualityNegativeCount++;
      out.sourceStyles['ROE'] = pfStyle_(C.red, C.text, 'bold');
    }
  }

  var roa = pfNumberByHeader_(row, headerMap, 'ROA');
  if (roa != null) qualityHasAnyData = true;
  if (pfIsMeaningfulQualityValue_(roa) && (pfHasValue_(shareRules.roaGood) || pfHasValue_(shareRules.roaBad))) {
    if (pfHasValue_(shareRules.roaGood) && roa >= shareRules.roaGood) {
      qualityPositiveCount++;
      out.sourceStyles['ROA'] = pfStyle_(C.green, C.text, 'bold');
    } else if (pfHasValue_(shareRules.roaBad) && roa <= shareRules.roaBad) {
      qualityNegativeCount++;
      out.sourceStyles['ROA'] = pfStyle_(C.red, C.text, 'bold');
    }
  }

  var roic = pfNumberByHeader_(row, headerMap, 'ROIC');
  if (roic != null) qualityHasAnyData = true;
  if (pfIsMeaningfulQualityValue_(roic) && (pfHasValue_(shareRules.roicGood) || pfHasValue_(shareRules.roicBad))) {
    if (pfHasValue_(shareRules.roicGood) && roic >= shareRules.roicGood) {
      qualityPositiveCount++;
      out.sourceStyles['ROIC'] = pfStyle_(C.green, C.text, 'bold');
    } else if (pfHasValue_(shareRules.roicBad) && roic <= shareRules.roicBad) {
      qualityNegativeCount++;
      out.sourceStyles['ROIC'] = pfStyle_(C.red, C.text, 'bold');
    }
  }

  var ebitda = pfNumberByHeader_(row, headerMap, 'EBITDA TTM');
  if (ebitda != null) qualityHasAnyData = true;
  if (ebitda != null && ebitda < 0) {
    hardNegativeCount++;
    pfPushUnique_(flags, 'NEG_EBITDA');
    out.sourceStyles['EBITDA TTM'] = pfStyle_(C.redStrong, C.text, 'bold');
  }

  var netIncome = pfNumberByHeader_(row, headerMap, 'Чистая прибыль TTM');
  if (netIncome != null) qualityHasAnyData = true;
  if (netIncome != null && netIncome < 0) {
    hardNegativeCount++;
    pfPushUnique_(flags, 'NEG_EARNINGS');
    out.sourceStyles['Чистая прибыль TTM'] = pfStyle_(C.redStrong, C.text, 'bold');
  }

  if (qualityHasAnyData) {
    if (hardNegativeCount > 0 || qualityNegativeCount >= 2) {
      out.service.quality = 'Weak';
      out.serviceStyles.quality = pfStyle_(C.red, C.text, 'bold');
    } else if (qualityPositiveCount >= 2 && qualityNegativeCount === 0) {
      out.service.quality = 'Strong';
      out.serviceStyles.quality = pfStyle_(C.green, C.text, 'bold');
    } else {
      out.service.quality = 'Neutral';
      out.serviceStyles.quality = pfStyle_(C.grey, C.text, 'normal');
    }
  }

  var cheapCount = 0;
  var expensiveCount = 0;
  var valuationInputs = 0;

  var pe = pfNumberByHeader_(row, headerMap, 'P/E TTM');
  if (pe != null && pe > 0 && (pfHasValue_(shareRules.peCheap) || pfHasValue_(shareRules.peExpensive))) {
    valuationInputs++;
    if (pfHasValue_(shareRules.peCheap) && pe <= shareRules.peCheap) {
      cheapCount++;
      out.sourceStyles['P/E TTM'] = pfStyle_(C.green, C.text, 'normal');
    } else if (pfHasValue_(shareRules.peExpensive) && pe >= shareRules.peExpensive) {
      expensiveCount++;
      out.sourceStyles['P/E TTM'] = pfStyle_(C.red, C.text, 'normal');
    }
  }

  var ps = pfNumberByHeader_(row, headerMap, 'P/S TTM');
  if (ps != null && ps > 0 && (pfHasValue_(shareRules.psCheap) || pfHasValue_(shareRules.psExpensive))) {
    valuationInputs++;
    if (pfHasValue_(shareRules.psCheap) && ps <= shareRules.psCheap) {
      cheapCount++;
      out.sourceStyles['P/S TTM'] = pfStyle_(C.green, C.text, 'normal');
    } else if (pfHasValue_(shareRules.psExpensive) && ps >= shareRules.psExpensive) {
      expensiveCount++;
      out.sourceStyles['P/S TTM'] = pfStyle_(C.red, C.text, 'normal');
    }
  }

  var pb = pfNumberByHeader_(row, headerMap, 'P/B TTM');
  if (pb != null && pb > 0 && (pfHasValue_(shareRules.pbCheap) || pfHasValue_(shareRules.pbExpensive))) {
    valuationInputs++;
    if (pfHasValue_(shareRules.pbCheap) && pb <= shareRules.pbCheap) {
      cheapCount++;
      out.sourceStyles['P/B TTM'] = pfStyle_(C.green, C.text, 'normal');
    } else if (pfHasValue_(shareRules.pbExpensive) && pb >= shareRules.pbExpensive) {
      expensiveCount++;
      out.sourceStyles['P/B TTM'] = pfStyle_(C.red, C.text, 'normal');
    }
  }

  var evEbitda = pfNumberByHeader_(row, headerMap, 'EV/EBITDA');
  if (evEbitda != null && evEbitda > 0 && (pfHasValue_(shareRules.evEbitdaCheap) || pfHasValue_(shareRules.evEbitdaExpensive))) {
    valuationInputs++;
    if (pfHasValue_(shareRules.evEbitdaCheap) && evEbitda <= shareRules.evEbitdaCheap) {
      cheapCount++;
      out.sourceStyles['EV/EBITDA'] = pfStyle_(C.green, C.text, 'normal');
    } else if (pfHasValue_(shareRules.evEbitdaExpensive) && evEbitda >= shareRules.evEbitdaExpensive) {
      expensiveCount++;
      out.sourceStyles['EV/EBITDA'] = pfStyle_(C.red, C.text, 'normal');
    }
  }

  if (valuationInputs > 0) {
    if (cheapCount > expensiveCount && cheapCount > 0) {
      out.service.valuation = 'Cheap';
      out.serviceStyles.valuation = pfStyle_(C.green, C.text, 'bold');
    } else if (expensiveCount > cheapCount && expensiveCount > 0) {
      out.service.valuation = 'Expensive';
      out.serviceStyles.valuation = pfStyle_(C.red, C.text, 'bold');
    } else {
      out.service.valuation = 'Fair';
      out.serviceStyles.valuation = pfStyle_(C.grey, C.text, 'normal');
    }
  }

  var riskPoints = 0;
  var riskInputs = 0;

  var beta = pfNumberByHeader_(row, headerMap, 'Beta');
  if (beta != null && pfHasValue_(shareRules.betaHigh)) {
    riskInputs++;
    if (beta >= shareRules.betaHigh) {
      riskPoints++;
      pfPushUnique_(flags, 'HIGH_BETA');
      out.sourceStyles['Beta'] = pfStyle_(C.amber, C.text, 'bold');
    }
  }

  var debtToEquity = pfNumberByHeader_(row, headerMap, 'Debt/Equity');
  if (debtToEquity != null && pfHasValue_(shareRules.debtToEquityHigh)) {
    riskInputs++;
    if (debtToEquity >= shareRules.debtToEquityHigh) {
      riskPoints++;
      pfPushUnique_(flags, 'HIGH_DEBT');
      out.sourceStyles['Debt/Equity'] = pfStyle_(C.amber, C.text, 'bold');
    }
  }

  var netDebtToEbitda = pfNumberByHeader_(row, headerMap, 'NetDebt/EBITDA');
  if (netDebtToEbitda != null && pfHasValue_(shareRules.netDebtToEbitdaHigh)) {
    riskInputs++;
    if (netDebtToEbitda >= shareRules.netDebtToEbitdaHigh) {
      riskPoints++;
      pfPushUnique_(flags, 'HIGH_NETDEBT');
      out.sourceStyles['NetDebt/EBITDA'] = pfStyle_(C.amber, C.text, 'bold');
    }
  }

  var freeFloat = pfNumberByHeader_(row, headerMap, 'Free Float');
  if (freeFloat != null && pfHasValue_(shareRules.freeFloatLow)) {
    riskInputs++;
    if (freeFloat <= shareRules.freeFloatLow) {
      riskPoints++;
      pfPushUnique_(flags, 'LOW_FLOAT');
      out.sourceStyles['Free Float'] = pfStyle_(C.orange, C.text, 'bold');
    }
  }

  var blockedText = pfTextByHeader_(row, headerMap, 'Заблокирован (TCS)');
  if (blockedText !== '') {
    riskInputs++;
    if (pfIsYes_(blockedText)) {
      riskPoints += 2;
      pfPushUnique_(flags, 'BLOCKED');
      out.sourceStyles['Заблокирован (TCS)'] = pfStyle_(C.grey, C.mutedText, 'bold');
    }
  }

  var shortText = pfTextByHeader_(row, headerMap, 'Шорт доступен');
  if (shortText !== '') {
    riskInputs++;
    if (pfIsYes_(shortText)) {
      out.sourceStyles['Шорт доступен'] = pfStyle_(C.blue, C.text, 'normal');
    }
  }

  if (riskInputs > 0) {
    if (pfIsYes_(blockedText) || riskPoints >= 2) {
      out.service.risk = 'High';
      out.serviceStyles.risk = pfStyle_(C.red, C.text, 'bold');
    } else if (riskPoints === 1) {
      out.service.risk = 'Medium';
      out.serviceStyles.risk = pfStyle_(C.amber, C.text, 'bold');
    } else {
      out.service.risk = 'Low';
      out.serviceStyles.risk = pfStyle_(C.green, C.text, 'normal');
    }
  }

  var currentPrice = pfNumberByHeader_(row, headerMap, 'Текущая цена');
  var high52 = pfNumberByHeader_(row, headerMap, '52w High');
  var low52 = pfNumberByHeader_(row, headerMap, '52w Low');

  if (currentPrice != null && high52 != null && low52 != null && high52 > low52 && (pfHasValue_(shareRules.near52wHigh) || pfHasValue_(shareRules.near52wLow))) {
    var range52 = high52 - low52;
    var distToHigh = (high52 - currentPrice) / range52;
    var distToLow = (currentPrice - low52) / range52;

    if (pfHasValue_(shareRules.near52wHigh) && distToHigh <= shareRules.near52wHigh) {
      out.service.position52w = 'Near High';
      out.serviceStyles.position52w = pfStyle_(C.blue, C.text, 'bold');
      out.sourceStyles['Текущая цена'] = pfStyle_(C.blue, C.text, 'bold');
    } else if (pfHasValue_(shareRules.near52wLow) && distToLow <= shareRules.near52wLow) {
      out.service.position52w = 'Near Low';
      out.serviceStyles.position52w = pfStyle_(C.purple, C.text, 'bold');
      out.sourceStyles['Текущая цена'] = pfStyle_(C.purple, C.text, 'bold');
    } else {
      out.service.position52w = 'Mid Range';
      out.serviceStyles.position52w = pfStyle_(C.grey, C.text, 'normal');
    }
  }

  var marketCap = pfNumberByHeader_(row, headerMap, 'Капитализация');
  if (marketCap != null) {
    if (pfHasValue_(shareRules.marketCapSmall) && marketCap <= shareRules.marketCapSmall) {
      pfPushUnique_(flags, 'SMALL_CAP');
      out.sourceStyles['Капитализация'] = pfStyle_(C.orange, C.text, 'bold');
    } else if (pfHasValue_(shareRules.marketCapLarge) && marketCap >= shareRules.marketCapLarge) {
      out.sourceStyles['Капитализация'] = pfStyle_(C.blue, C.text, 'normal');
    }
  }

  if (general.showWarnings) {
    out.service.flags = flags.join(', ');
    out.serviceStyles.flags = flags.length
      ? pfStyle_(C.amber, C.text, 'normal')
      : pfStyle_(C.neutralBg, C.mutedText, 'normal');
  } else {
    out.service.flags = '';
    out.serviceStyles.flags = pfStyle_(C.neutralBg, C.mutedText, 'normal');
  }

  return out;
}

// ==========================================================
// Bonds derived metrics / signals
// ==========================================================

function pfPopulateBondDerivedRow_(row, headerMap, rules) {
  var out = {
    values: {
      'Риск (ручн.)': '',
      'Инвестировано': '',
      'Рыночная стоимость': '',
      'P/L (руб)': '',
      'P/L (%)': '',
      'Доходность купонная годовая (прибл.)': ''
    }
  };

  var figi = String(pfValueByHeader_(row, headerMap, 'FIGI') || '').trim();
  if (!figi) return out;

  var qty = pfNumberByHeader_(row, headerMap, 'Кол-во');
  var avgPrice = pfNumberByHeader_(row, headerMap, 'Средняя цена');
  var currentPrice = pfNumberByHeader_(row, headerMap, 'Текущая цена');
  var couponValue = pfNumberByHeader_(row, headerMap, 'Размер купона');
  var couponsPerYear = pfNumberByHeader_(row, headerMap, 'купон/год');

  var invested = null;
  if (qty != null && avgPrice != null) {
    invested = pfRound2_(qty * avgPrice);
    out.values['Инвестировано'] = invested;
  }

  var marketValue = null;
  if (qty != null && currentPrice != null) {
    marketValue = pfRound2_(qty * currentPrice);
    out.values['Рыночная стоимость'] = marketValue;
  }

  var plRub = null;
  if (marketValue != null && invested != null) {
    plRub = pfRound2_(marketValue - invested);
    out.values['P/L (руб)'] = plRub;
  }

  var plPct = null;
  if (plRub != null && invested != null && Math.abs(invested) > 1e-9) {
    plPct = pfRound2_((plRub / invested) * 100);
    out.values['P/L (%)'] = plPct;
  }

  var annualYieldPct = null;
  if (couponValue != null && couponsPerYear != null && currentPrice != null && Math.abs(currentPrice) > 1e-9) {
    annualYieldPct = pfRound2_((couponValue * couponsPerYear / currentPrice) * 100);
    out.values['Доходность купонная годовая (прибл.)'] = annualYieldPct;
  }

  var manualRisk = pfCalculateBondManualRisk_(row, headerMap, rules, {
    plPct: plPct
  });

  if (manualRisk != null) {
    out.values['Риск (ручн.)'] = pfRoundRiskScore_(manualRisk);
  }

  return out;
}

function pfCalculateBondManualRisk_(row, headerMap, rules, derivedContext) {
  var bondRules = (rules && rules.bonds) || {};

  var tcsScore = pfBondTcsBaseRiskScore_(
    pfTextByHeader_(row, headerMap, 'Класс риска TCS'),
    pfTextByHeader_(row, headerMap, 'Риск (уровень TCS)')
  );

  var sectorScore = pfBondSectorRiskScore_(
    pfTextByHeader_(row, headerMap, 'Сектор'),
    bondRules.sectorRiskScores || {}
  );

  var drawdownScore = pfBondDrawdownRiskScore_(
    derivedContext ? derivedContext.plPct : null,
    bondRules.riskDrawdownPct,
    bondRules.strongLossPct
  );

  var shortMaturityScore = pfBondShortMaturityRiskScore_(
    pfDateByHeader_(row, headerMap, 'Дата погашения'),
    bondRules.shortMaturityYears
  );

  if (tcsScore == null && sectorScore == null && drawdownScore == null && shortMaturityScore == null) {
    return null;
  }

  return (tcsScore || 0) + (sectorScore || 0) + (drawdownScore || 0) + (shortMaturityScore || 0);
}

function pfAnalyzeBondRow_(row, headerMap, rules) {
  var C = PORTFOLIO_FORMATING_COLORS;
  var general = rules.general || {};
  var bondRules = rules.bonds || {};

  var out = {
    sourceStyles: {},
    service: {
      yield: '',
      bondRisk: '',
      maturity: '',
      coupon: '',
      flags: ''
    },
    serviceStyles: {
      yield: pfStyle_(C.neutralBg, C.text, 'normal'),
      bondRisk: pfStyle_(C.neutralBg, C.text, 'normal'),
      maturity: pfStyle_(C.neutralBg, C.text, 'normal'),
      coupon: pfStyle_(C.neutralBg, C.text, 'normal'),
      flags: pfStyle_(C.neutralBg, C.mutedText, 'normal')
    }
  };

  var figi = String(pfValueByHeader_(row, headerMap, 'FIGI') || '').trim();
  if (!figi) return out;

  var flags = [];

  var plRub = pfNumberByHeader_(row, headerMap, 'P/L (руб)');
  var plPct = pfNumberByHeader_(row, headerMap, 'P/L (%)');

  if (general.colorPl) {
    var plStyle = null;
    var strongProfit = (plPct != null) && pfPositiveThresholdMet_(plPct, bondRules.strongProfitPct);
    var strongLoss = (plPct != null) && pfNegativeThresholdMet_(plPct, bondRules.strongLossPct);

    if ((plRub != null && plRub > 0) || (plPct != null && plPct > 0)) {
      plStyle = strongProfit
        ? pfStyle_(C.greenStrong, C.text, 'bold')
        : pfStyle_(C.green, C.text, 'normal');
    } else if ((plRub != null && plRub < 0) || (plPct != null && plPct < 0)) {
      plStyle = strongLoss
        ? pfStyle_(C.redStrong, C.text, 'bold')
        : pfStyle_(C.red, C.text, 'normal');
      pfPushUnique_(flags, 'NEG_PL');
      if (strongLoss) pfPushUnique_(flags, 'STRONG_LOSS');
    }

    if (plStyle) {
      out.sourceStyles['P/L (руб)'] = plStyle;
      out.sourceStyles['P/L (%)'] = plStyle;
    }
  }

  var annualYieldPct = pfNumberByHeader_(row, headerMap, 'Доходность купонная годовая (прибл.)');
  var yieldInfo = pfClassifyBondYield_(annualYieldPct, bondRules);
  if (yieldInfo.label) {
    out.service.yield = yieldInfo.label;
    out.serviceStyles.yield = yieldInfo.serviceStyle;
    out.sourceStyles['Доходность купонная годовая (прибл.)'] = yieldInfo.sourceStyle;
    if (yieldInfo.flag) pfPushUnique_(flags, yieldInfo.flag);
  }

  var maturityDate = pfDateByHeader_(row, headerMap, 'Дата погашения');
  var couponTypeText = pfTextByHeader_(row, headerMap, 'Тип купона (desc)');

  var manualRisk = pfNumberByHeader_(row, headerMap, 'Риск (ручн.)');
  var rawTcsRiskClass = pfTextByHeader_(row, headerMap, 'Класс риска TCS');
  var tcsRiskText = pfTextByHeader_(row, headerMap, 'Риск (уровень TCS)');
  var riskInfo = pfClassifyBondRisk_(manualRisk, rawTcsRiskClass, tcsRiskText, bondRules);
  if (riskInfo.label) {
    out.service.bondRisk = riskInfo.label;
    out.serviceStyles.bondRisk = riskInfo.serviceStyle;

    if (riskInfo.source === 'manual') {
      out.sourceStyles['Риск (ручн.)'] = riskInfo.sourceStyle;
    } else if (riskInfo.source === 'tcsRaw') {
      out.sourceStyles['Класс риска TCS'] = riskInfo.sourceStyle;
    } else if (riskInfo.source === 'tcsText') {
      out.sourceStyles['Риск (уровень TCS)'] = riskInfo.sourceStyle;
    }

    if (riskInfo.label === 'High') pfPushUnique_(flags, 'HIGH_RISK');
  }

  var maturityInfo = pfClassifyBondMaturity_(maturityDate, bondRules);
  if (maturityInfo.label) {
    out.service.maturity = maturityInfo.label;
    out.serviceStyles.maturity = maturityInfo.serviceStyle;
    out.sourceStyles['Дата погашения'] = maturityInfo.sourceStyle;

    if (maturityInfo.label === 'Long') {
      pfPushUnique_(flags, 'LONG_DURATION');
    } else if (maturityInfo.label === 'Very Long') {
      pfPushUnique_(flags, 'VERY_LONG_DURATION');
    }
  }

  var nextCouponDate = pfDateByHeader_(row, headerMap, 'Следующий купон');
  var nextCouponSoonDays = pfNormalizeNumber_(bondRules.nextCouponSoonDays);
  var daysToCoupon = pfDaysUntilDate_(nextCouponDate);
  if (nextCouponDate && nextCouponSoonDays != null && daysToCoupon != null && daysToCoupon >= 0 && daysToCoupon <= nextCouponSoonDays) {
    out.sourceStyles['Следующий купон'] = pfStyle_(C.blue, C.text, 'bold');
    pfPushUnique_(flags, 'COUPON_SOON');
  }

  var couponInfo = pfClassifyBondCoupon_(couponTypeText);
  if (couponInfo.label) {
    out.service.coupon = couponInfo.label;
    out.serviceStyles.coupon = couponInfo.serviceStyle;
    out.sourceStyles['Тип купона (desc)'] = couponInfo.sourceStyle;

    if (couponInfo.key === 'floating' && pfIsTrueLike_(bondRules.flagFloatingCoupon)) {
      pfPushUnique_(flags, 'FLOATING_COUPON');
    }
    if (couponInfo.key === 'variable' && pfIsTrueLike_(bondRules.flagVariableCoupon)) {
      pfPushUnique_(flags, 'VARIABLE_COUPON');
    }
  }

  if (general.showWarnings) {
    out.service.flags = flags.join(', ');
    out.serviceStyles.flags = flags.length
      ? pfStyle_(C.amber, C.text, 'normal')
      : pfStyle_(C.neutralBg, C.mutedText, 'normal');
  } else {
    out.service.flags = '';
    out.serviceStyles.flags = pfStyle_(C.neutralBg, C.mutedText, 'normal');
  }

  return out;
}

function pfClassifyBondYield_(yieldPct, bondRules) {
  var C = PORTFOLIO_FORMATING_COLORS;
  if (yieldPct == null) return pfEmptyBondSignal_();

  var low = pfNormalizeNumber_(bondRules.yieldLowPct);
  var high = pfNormalizeNumber_(bondRules.yieldHighPct);
  var veryHigh = pfNormalizeNumber_(bondRules.yieldVeryHighPct);

  if (low == null && high == null && veryHigh == null) return pfEmptyBondSignal_();

  if (veryHigh != null && yieldPct >= veryHigh) {
    return {
      label: 'Very High',
      serviceStyle: pfStyle_(C.redStrong, C.text, 'bold'),
      sourceStyle: pfStyle_(C.redStrong, C.text, 'bold'),
      flag: 'VERY_HIGH_YIELD'
    };
  }

  if (high != null && yieldPct >= high) {
    return {
      label: 'High',
      serviceStyle: pfStyle_(C.amber, C.text, 'bold'),
      sourceStyle: pfStyle_(C.amber, C.text, 'bold'),
      flag: 'HIGH_YIELD'
    };
  }

  if (low != null && yieldPct < low) {
    return {
      label: 'Low',
      serviceStyle: pfStyle_(C.grey, C.mutedText, 'normal'),
      sourceStyle: pfStyle_(C.grey, C.mutedText, 'normal'),
      flag: ''
    };
  }

  return {
    label: 'Normal',
    serviceStyle: pfStyle_(C.neutralSoft, C.text, 'normal'),
    sourceStyle: pfStyle_(C.neutralSoft, C.text, 'normal'),
    flag: ''
  };
}

function pfClassifyBondRisk_(manualRisk, rawTcsRiskClass, tcsRiskText, bondRules) {
  var baseInfo = pfResolveBondRiskBase_(manualRisk, rawTcsRiskClass, tcsRiskText, bondRules);
  if (!baseInfo.label) return pfEmptyBondSignal_();

  var finalStyle = pfBondRiskStyleByLabel_(baseInfo.label);
  return {
    label: baseInfo.label,
    source: baseInfo.source,
    serviceStyle: finalStyle,
    sourceStyle: finalStyle
  };
}

function pfResolveBondRiskBase_(manualRisk, rawTcsRiskClass, tcsRiskText, bondRules) {
  var medium = pfNormalizeNumber_(bondRules.manualRiskMedium);
  var high = pfNormalizeNumber_(bondRules.manualRiskHigh);

  if (pfIsUsableBondManualRisk_(manualRisk) && (medium != null || high != null)) {
    if (high != null && manualRisk >= high) return { label: 'High', source: 'manual' };
    if (medium != null && manualRisk >= medium) return { label: 'Medium', source: 'manual' };
    return { label: 'Low', source: 'manual' };
  }

  var rawLabel = pfMapBondTcsRiskRaw_(rawTcsRiskClass);
  if (rawLabel) return { label: rawLabel, source: 'tcsRaw' };

  var textLabel = pfMapBondTcsRisk_(tcsRiskText);
  if (textLabel) return { label: textLabel, source: 'tcsText' };

  return { label: '', source: '' };
}

function pfBondRiskStyleByLabel_(label) {
  var C = PORTFOLIO_FORMATING_COLORS;
  if (label === 'High') return pfStyle_(C.red, C.text, 'bold');
  if (label === 'Medium') return pfStyle_(C.amber, C.text, 'bold');
  if (label === 'Low') return pfStyle_(C.green, C.text, 'normal');
  return pfStyle_(C.neutralBg, C.text, 'normal');
}

function pfClassifyBondMaturity_(maturityDate, bondRules) {
  var C = PORTFOLIO_FORMATING_COLORS;
  if (!maturityDate) return pfEmptyBondSignal_();

  var shortYears = pfNormalizeNumber_(bondRules.shortMaturityYears);
  var longYears = pfNormalizeNumber_(bondRules.longMaturityYears);
  var veryLongYears = pfNormalizeNumber_(bondRules.veryLongMaturityYears);
  if (shortYears == null && longYears == null && veryLongYears == null) return pfEmptyBondSignal_();

  var daysTo = pfDaysUntilDate_(maturityDate);
  if (daysTo == null) return pfEmptyBondSignal_();

  var yearsTo = daysTo / 365.25;
  var label = '';

  if (veryLongYears != null && yearsTo >= veryLongYears) {
    label = 'Very Long';
  } else if (longYears != null && yearsTo >= longYears) {
    label = 'Long';
  } else if (shortYears != null && yearsTo <= shortYears) {
    label = 'Near Maturity';
  } else if (shortYears != null && longYears != null && longYears > shortYears) {
    label = (yearsTo <= (shortYears + (longYears - shortYears) / 2)) ? 'Short' : 'Medium';
  } else if (shortYears != null) {
    label = (yearsTo <= shortYears * 3) ? 'Short' : 'Medium';
  } else if (longYears != null) {
    label = (yearsTo < longYears * 0.6) ? 'Short' : 'Medium';
  } else {
    label = 'Medium';
  }

  if (label === 'Very Long') {
    return {
      label: label,
      serviceStyle: pfStyle_(C.red, C.text, 'bold'),
      sourceStyle: pfStyle_(C.red, C.text, 'bold')
    };
  }

  if (label === 'Long') {
    return {
      label: label,
      serviceStyle: pfStyle_(C.amber, C.text, 'bold'),
      sourceStyle: pfStyle_(C.amber, C.text, 'bold')
    };
  }

  if (label === 'Near Maturity') {
    return {
      label: label,
      serviceStyle: pfStyle_(C.blue, C.text, 'bold'),
      sourceStyle: pfStyle_(C.blue, C.text, 'bold')
    };
  }

  if (label === 'Short') {
    return {
      label: label,
      serviceStyle: pfStyle_(C.neutralSoft, C.text, 'normal'),
      sourceStyle: pfStyle_(C.neutralSoft, C.text, 'normal')
    };
  }

  return {
    label: 'Medium',
    serviceStyle: pfStyle_(C.grey, C.text, 'normal'),
    sourceStyle: pfStyle_(C.grey, C.text, 'normal')
  };
}

function pfClassifyBondCoupon_(couponTypeText) {
  var C = PORTFOLIO_FORMATING_COLORS;
  var text = String(couponTypeText || '').trim().toLowerCase();
  if (!text) return pfEmptyBondSignal_();

  if (text.indexOf('фикс') !== -1 || text.indexOf('fixed') !== -1) {
    return {
      key: 'fixed',
      label: 'Fixed',
      serviceStyle: pfStyle_(C.neutralSoft, C.text, 'normal'),
      sourceStyle: pfStyle_(C.neutralSoft, C.text, 'normal')
    };
  }

  if (text.indexOf('плава') !== -1 || text.indexOf('floating') !== -1) {
    return {
      key: 'floating',
      label: 'Floating',
      serviceStyle: pfStyle_(C.purple, C.text, 'bold'),
      sourceStyle: pfStyle_(C.purple, C.text, 'bold')
    };
  }

  if (text.indexOf('перемен') !== -1 || text.indexOf('variable') !== -1) {
    return {
      key: 'variable',
      label: 'Variable',
      serviceStyle: pfStyle_(C.amber, C.text, 'bold'),
      sourceStyle: pfStyle_(C.amber, C.text, 'bold')
    };
  }

  return {
    key: 'other',
    label: 'Other',
    serviceStyle: pfStyle_(C.grey, C.text, 'normal'),
    sourceStyle: pfStyle_(C.grey, C.text, 'normal')
  };
}

function pfEmptyBondSignal_() {
  return {
    label: '',
    source: '',
    serviceStyle: pfStyle_(PORTFOLIO_FORMATING_COLORS.neutralBg, PORTFOLIO_FORMATING_COLORS.text, 'normal'),
    sourceStyle: pfStyle_(PORTFOLIO_FORMATING_COLORS.neutralBg, PORTFOLIO_FORMATING_COLORS.text, 'normal'),
    flag: ''
  };
}

function pfBondTcsBaseRiskScore_(rawRiskClass, riskText) {
  var raw = pfMapBondTcsRiskRaw_(rawRiskClass);
  var text = pfMapBondTcsRisk_(riskText);
  var label = raw || text;
  if (!label) return null;
  if (label === 'Low') return 0;
  if (label === 'Medium') return 2;
  if (label === 'High') return 4;
  return null;
}

function pfBondSectorRiskScore_(sectorText, sectorScoreMap) {
  var key = pfNormalizeRuleMapKey_(sectorText);
  if (!key) return null;
  if (!sectorScoreMap || !Object.prototype.hasOwnProperty.call(sectorScoreMap, key)) return 0;
  var score = pfNormalizeNumber_(sectorScoreMap[key]);
  return score == null ? 0 : score;
}

function pfBondDrawdownRiskScore_(plPct, riskDrawdownPct, strongLossPct) {
  if (plPct == null) return 0;

  var threshold = pfNormalizeNumber_(riskDrawdownPct);
  if (threshold == null && strongLossPct != null) {
    threshold = -Math.abs(Number(strongLossPct));
  }
  if (threshold == null) return 0;
  if (threshold > 0) threshold = -Math.abs(threshold);

  return plPct <= threshold ? 1 : 0;
}

function pfBondShortMaturityRiskScore_(maturityDate, shortMaturityYears) {
  var yearsLimit = pfNormalizeNumber_(shortMaturityYears);
  if (!maturityDate || yearsLimit == null) return 0;

  var yearsTo = pfYearsUntilDate_(maturityDate);
  if (yearsTo == null) return 0;
  return yearsTo <= yearsLimit ? 1 : 0;
}

function pfMapBondTcsRiskRaw_(text) {
  var s = String(text || '').trim().toUpperCase();
  if (!s) return '';
  if (s.indexOf('RISK_LEVEL_LOW') !== -1 || s === 'LOW') return 'Low';
  if (s.indexOf('RISK_LEVEL_MODERATE') !== -1 || s.indexOf('MODERATE') !== -1 || s === 'MEDIUM') return 'Medium';
  if (s.indexOf('RISK_LEVEL_HIGH') !== -1 || s === 'HIGH') return 'High';
  return '';
}

function pfMapBondTcsRisk_(text) {
  var s = String(text || '').trim().toLowerCase();
  if (!s) return '';
  if (s.indexOf('низ') !== -1 || s === 'low') return 'Low';
  if (s.indexOf('сред') !== -1 || s === 'medium') return 'Medium';
  if (s.indexOf('выс') !== -1 || s === 'high') return 'High';
  return '';
}

// ==========================================================
// Shares formatting
// ==========================================================

function pfFormatSharesServiceHeaders_(sh, headerMap) {
  PORTFOLIO_FORMATING_SHARES_SERVICE_ORDER.forEach(function(header) {
    var col = headerMap[header];
    if (!col) return;

    sh.getRange(1, col)
      .setBackground(PORTFOLIO_FORMATING_COLORS.headerBg)
      .setFontColor(PORTFOLIO_FORMATING_COLORS.headerText)
      .setFontWeight('bold');

    if (header === PORTFOLIO_FORMATING_SHARES_SERVICE.flags) {
      sh.setColumnWidth(col, 260);
    } else if (header === PORTFOLIO_FORMATING_SHARES_SERVICE.position52w) {
      sh.setColumnWidth(col, 120);
    } else {
      sh.setColumnWidth(col, 110);
    }
  });
}

function pfFormatSharesServiceData_(sh, headerMap, numRows) {
  if (numRows < 1) return;

  PORTFOLIO_FORMATING_SHARES_SERVICE_ORDER.forEach(function(header) {
    var col = headerMap[header];
    if (!col) return;

    var range = sh.getRange(2, col, numRows, 1);
    range.setVerticalAlignment('middle');

    if (header === PORTFOLIO_FORMATING_SHARES_SERVICE.flags) {
      range.setHorizontalAlignment('left');
      range.setWrap(true);
    } else {
      range.setHorizontalAlignment('center');
      range.setWrap(false);
    }
  });
}

// ==========================================================
// Bonds formatting
// ==========================================================

function pfFormatBondsServiceHeaders_(sh, headerMap) {
  PORTFOLIO_FORMATING_BONDS_SERVICE_ORDER.forEach(function(header) {
    var col = headerMap[header];
    if (!col) return;

    sh.getRange(1, col)
      .setBackground(PORTFOLIO_FORMATING_COLORS.headerBg)
      .setFontColor(PORTFOLIO_FORMATING_COLORS.headerText)
      .setFontWeight('bold');

    if (header === PORTFOLIO_FORMATING_BONDS_SERVICE.flags) {
      sh.setColumnWidth(col, 280);
    } else if (header === PORTFOLIO_FORMATING_BONDS_SERVICE.maturity) {
      sh.setColumnWidth(col, 125);
    } else if (header === PORTFOLIO_FORMATING_BONDS_SERVICE.bondRisk) {
      sh.setColumnWidth(col, 115);
    } else {
      sh.setColumnWidth(col, 110);
    }
  });
}

function pfFormatBondsServiceData_(sh, headerMap, numRows) {
  if (numRows < 1) return;

  PORTFOLIO_FORMATING_BONDS_SERVICE_ORDER.forEach(function(header) {
    var col = headerMap[header];
    if (!col) return;

    var range = sh.getRange(2, col, numRows, 1);
    range.setVerticalAlignment('middle');

    if (header === PORTFOLIO_FORMATING_BONDS_SERVICE.flags) {
      range.setHorizontalAlignment('left');
      range.setWrap(true);
    } else {
      range.setHorizontalAlignment('center');
      range.setWrap(false);
    }
  });
}

// ==========================================================
// Helpers for service columns
// ==========================================================

function pfEnsureServiceColumns_(sh, headerMap, serviceHeaders, createMissing) {
  var headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  var map = pfBuildHeaderMap_(headers);
  var missing = [];

  serviceHeaders.forEach(function(header) {
    if (!pfHasHeader_(map, header)) missing.push(header);
  });

  if (createMissing && missing.length) {
    var startCol = sh.getLastColumn() + 1;
    sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
    headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    map = pfBuildHeaderMap_(headers);
  }

  return {
    headers: headers,
    headerMap: map,
    added: createMissing ? missing.length : 0
  };
}

// ==========================================================
// Common helpers
// ==========================================================

function pfBuildHeaderMap_(headers) {
  var map = {};
  (headers || []).forEach(function(header, idx) {
    var key = String(header || '').trim();
    if (!key) return;
    if (!Object.prototype.hasOwnProperty.call(map, key)) {
      map[key] = idx + 1;
    }
  });
  return map;
}

function pfHasHeader_(headerMap, header) {
  return !!(headerMap && Object.prototype.hasOwnProperty.call(headerMap, header) && headerMap[header] > 0);
}

function pfExistingHeaders_(headerMap, headers) {
  return (headers || []).filter(function(header) {
    return pfHasHeader_(headerMap, header);
  });
}

function pfCountExistingHeaders_(headerMap, headers) {
  return pfExistingHeaders_(headerMap, headers).length;
}

function pfValueByHeader_(row, headerMap, header) {
  if (!pfHasHeader_(headerMap, header)) return '';
  return row[headerMap[header] - 1];
}

function pfTextByHeader_(row, headerMap, header) {
  var v = pfValueByHeader_(row, headerMap, header);
  return String(v == null ? '' : v).trim();
}

function pfNumberByHeader_(row, headerMap, header) {
  return pfNormalizeNumber_(pfValueByHeader_(row, headerMap, header));
}

function pfDateByHeader_(row, headerMap, header) {
  return pfNormalizeDate_(pfValueByHeader_(row, headerMap, header));
}

function pfNormalizeNumber_(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;

  var s = String(v).trim();
  if (!s) return null;

  s = s.replace(/[%\u00A0\s]/g, '');
  s = s.replace(',', '.');

  var n = Number(s);
  return isNaN(n) ? null : n;
}

function pfNormalizeFraction_(v) {
  if (v == null || v === '') return null;

  if (typeof v === 'string' && v.indexOf('%') !== -1) {
    var s = String(v).replace('%', '').trim().replace(',', '.');
    var nPct = Number(s);
    return isNaN(nPct) ? null : nPct / 100;
  }

  var n = pfNormalizeNumber_(v);
  if (n == null) return null;
  if (Math.abs(n) > 1) return n / 100;
  return n;
}

function pfNormalizeDate_(v) {
  if (v == null || v === '') return null;

  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  if (typeof v === 'number' && isFinite(v)) {
    var epoch = new Date(1899, 11, 30);
    var dt = new Date(epoch.getTime() + v * 24 * 60 * 60 * 1000);
    return isNaN(dt.getTime()) ? null : new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  }

  var s = String(v).trim();
  if (!s) return null;

  var mRu = s.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})$/);
  if (mRu) {
    var d1 = new Date(Number(mRu[3]), Number(mRu[2]) - 1, Number(mRu[1]));
    return isNaN(d1.getTime()) ? null : d1;
  }

  var mIso = s.match(/^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})/);
  if (mIso) {
    var d2 = new Date(Number(mIso[1]), Number(mIso[2]) - 1, Number(mIso[3]));
    return isNaN(d2.getTime()) ? null : d2;
  }

  var d = new Date(s);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function pfDaysUntilDate_(dateValue) {
  var dt = pfNormalizeDate_(dateValue);
  if (!dt) return null;

  var today = new Date();
  today = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return Math.round((dt.getTime() - today.getTime()) / (24 * 60 * 60 * 1000));
}

function pfYearsUntilDate_(dateValue) {
  var days = pfDaysUntilDate_(dateValue);
  if (days == null) return null;
  return days / 365.25;
}

function pfParseBool_(v, defaultValue) {
  if (v === true || v === false) return v;
  if (v == null || v === '') return defaultValue;

  var s = String(v).trim().toLowerCase();
  if (s === 'true' || s === '1' || s === 'yes' || s === 'y' || s === 'да' || s === 'on') return true;
  if (s === 'false' || s === '0' || s === 'no' || s === 'n' || s === 'нет' || s === 'off') return false;
  return defaultValue;
}

function pfIsYes_(v) {
  if (v === true) return true;
  if (v === false || v == null || v === '') return false;

  var s = String(v).trim().toLowerCase();
  return s === 'да' || s === 'true' || s === '1' || s === 'yes' || s === 'y';
}

function pfIsTrueLike_(v) {
  return pfParseBool_(v, false) === true;
}

function pfHasValue_(v) {
  return !(v == null || v === '');
}

function pfIsEmpty_(v) {
  return v == null || v === '';
}

function pfPositiveThresholdMet_(value, threshold) {
  if (value == null || threshold == null) return false;
  var t = Number(threshold);
  if (isNaN(t)) return false;
  if (t < 0) t = Math.abs(t);
  return value >= t;
}

function pfNegativeThresholdMet_(value, threshold) {
  if (value == null || threshold == null) return false;
  var t = Number(threshold);
  if (isNaN(t)) return false;
  if (t > 0) t = -Math.abs(t);
  return value <= t;
}

function pfPushUnique_(arr, value) {
  if (!value) return;
  if (arr.indexOf(value) === -1) arr.push(value);
}

function pfStyle_(bg, fontColor, fontWeight) {
  return {
    bg: bg || PORTFOLIO_FORMATING_COLORS.neutralBg,
    fontColor: fontColor || PORTFOLIO_FORMATING_COLORS.text,
    fontWeight: fontWeight || 'normal'
  };
}

function pfCreateStyleMatrix_(numRows, defaultBg, defaultFontColor, defaultFontWeight) {
  var backgrounds = [];
  var fontColors = [];
  var fontWeights = [];

  for (var i = 0; i < numRows; i++) {
    backgrounds.push([defaultBg]);
    fontColors.push([defaultFontColor]);
    fontWeights.push([defaultFontWeight]);
  }

  return {
    backgrounds: backgrounds,
    fontColors: fontColors,
    fontWeights: fontWeights
  };
}

function pfSetMatrixCell_(matrix, rowIndex, bg, fontColor, fontWeight) {
  if (!matrix || rowIndex < 0) return;
  if (bg != null) matrix.backgrounds[rowIndex][0] = bg;
  if (fontColor != null) matrix.fontColors[rowIndex][0] = fontColor;
  if (fontWeight != null) matrix.fontWeights[rowIndex][0] = fontWeight;
}

function pfApplyColumnMatrix_(sh, col, matrix, numRows) {
  if (!matrix || !col || !numRows) return;
  var range = sh.getRange(2, col, numRows, 1);
  range.setBackgrounds(matrix.backgrounds);
  range.setFontColors(matrix.fontColors);
  range.setFontWeights(matrix.fontWeights);
}

function pfSharesServiceKeyByHeader_(header) {
  if (header === PORTFOLIO_FORMATING_SHARES_SERVICE.quality) return 'quality';
  if (header === PORTFOLIO_FORMATING_SHARES_SERVICE.valuation) return 'valuation';
  if (header === PORTFOLIO_FORMATING_SHARES_SERVICE.risk) return 'risk';
  if (header === PORTFOLIO_FORMATING_SHARES_SERVICE.position52w) return 'position52w';
  if (header === PORTFOLIO_FORMATING_SHARES_SERVICE.flags) return 'flags';
  return '';
}

function pfBondServiceKeyByHeader_(header) {
  if (header === PORTFOLIO_FORMATING_BONDS_SERVICE.yield) return 'yield';
  if (header === PORTFOLIO_FORMATING_BONDS_SERVICE.bondRisk) return 'bondRisk';
  if (header === PORTFOLIO_FORMATING_BONDS_SERVICE.maturity) return 'maturity';
  if (header === PORTFOLIO_FORMATING_BONDS_SERVICE.coupon) return 'coupon';
  if (header === PORTFOLIO_FORMATING_BONDS_SERVICE.flags) return 'flags';
  return '';
}

function pfNormalizeColumnRef_(ref) {
  if (typeof ref === 'number') return ref > 0 ? Math.floor(ref) : null;
  var s = String(ref || '').trim().toUpperCase();
  if (!s) return null;
  if (/^\d+$/.test(s)) return Number(s);

  var col = 0;
  for (var i = 0; i < s.length; i++) {
    var code = s.charCodeAt(i);
    if (code < 65 || code > 90) return null;
    col = col * 26 + (code - 64);
  }
  return col || null;
}

function pfNormalizeRuleMapKey_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function pfReadNamedRangeScalar_(name) {
  try {
    var range = SpreadsheetApp.getActive().getRangeByName(name);
    if (!range) return null;
    return pfNormalizeNumber_(range.getValue());
  } catch (e) {
    return null;
  }
}

function pfReadNamedRangeMap_(name) {
  try {
    var range = SpreadsheetApp.getActive().getRangeByName(name);
    if (!range) return {};

    var values = range.getValues();
    var map = {};
    for (var i = 0; i < values.length; i++) {
      var key = String(values[i][0] == null ? '' : values[i][0]).trim();
      var value = pfNormalizeNumber_(values[i][1]);
      if (!key || value == null) continue;
      map[pfNormalizeRuleMapKey_(key)] = value;
    }
    return map;
  } catch (e) {
    return {};
  }
}

function pfRound2_(value) {
  if (value == null || !isFinite(value)) return null;
  return Math.round(value * 100) / 100;
}

function pfRoundRiskScore_(value) {
  if (value == null || !isFinite(value)) return null;
  var rounded = Math.round(value * 100) / 100;
  if (Math.abs(rounded - Math.round(rounded)) < 1e-9) return Math.round(rounded);
  return rounded;
}

function pfSheetValue_(value) {
  return value == null ? '' : value;
}

function pfIsZeroLike_(value) {
  if (value == null) return false;
  return Math.abs(Number(value) || 0) < 1e-9;
}

function pfIsUsableBondManualRisk_(value) {
  if (value == null || value === '') return false;
  var n = Number(value);
  return !isNaN(n);
}

function pfIsMeaningfulQualityValue_(value) {
  return value != null && !pfIsZeroLike_(value);
}
