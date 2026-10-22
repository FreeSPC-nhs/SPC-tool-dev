
/*
 * SimpleSPC English translations
 *
 * This is the reference language for the application.
 * Other languages will use the same translation keys.
 */

window.SimpleSPCi18n.registerLanguage("en", {

  // Application heading
  "app.title": "Simple SPC Web Tool",

  "app.subtitle": "Upload data and create SPC charts in your browser. Data never leaves your browser.",

  // Main buttons
  "buttons.generate": "Generate / Recalculate",
  "buttons.reset": "Reset chart",
  "buttons.help": "Help & guidance",
  "buttons.spcHelper": "SPC helper",

  // Main tabs
  "tabs.data": "Data",
  "tabs.chart": "Chart",
  "tabs.labels": "Labels",
  "tabs.axes": "Axes",
  "tabs.baseline": "Baseline",
  
  "tabs.rules": "Rules",

  // Language selector
  "language.label": "Language",

  // Introduction
  "app.intro": "Add your data, choose your chart and settings, then click 'Generate / Recalculate' when needed. You can upload CSV or Excel files, clean data in the editor, and save chart to reopen later. Not sure about SPC? Ask the helper.",

  // Tooltips
  "buttons.resetTitle": "Keep your loaded data and selected columns, but reset chart settings including targets, splits, annotations, rules and formatting.",

  
"tabs.baselineTitle": "Set an initial baseline period or a performance target",

// DATA PANEL
"data.addHeading": "Add data:",
"data.requirementsAria": "Data requirements help",
"data.requirementsTooltip": "The columns you need depend on the chart type. Most SPC charts need a time or sequence column and a measurement. Some also need counts and denominators. Pareto charts need a category column and can optionally use a frequency/count column.",
"data.uploadHint": "Upload a CSV/Excel file or paste data in the editor.",
"data.openEditor": "Open data editor (paste/type data)",
"data.clear": "Clear data & chart",
"data.clearTitle": "Remove the loaded data and reset the chart. Save your work first if you want to reopen it later.",

"data.saveHeading": "Save or open a chart",
"data.saveHint": "Save your chart, data and settings as a JSON file to reopen and edit later.",
"data.saveButton": "Save chart",
"data.openButton": "Open saved chart",

"data.portableHeading": "Save an offline copy",
"data.portableAria": "Portable chart help",
"data.portableTooltip": "Save a self-contained HTML copy of your chart. The file includes your data and a working copy of SimpleSPC. You can open it without an internet connection, make changes and save it again. Portable files contain your data. Store and share them in line with your organisation's information-governance requirements.",
"data.portableButton": "Save portable chart",

// CHART PANEL
"chart.noData": "Load data to enable chart and column selection.",
"chart.newDataHeading": "New data loaded",
"chart.newDataQuestion": "Would you like help choosing the most appropriate chart for these data?",
"chart.helpChoose": "Help me choose",
"chart.chooseManually": "Choose manually",
"chart.typeHeading": "Chart type:",
"chart.run": "Run chart",
"chart.xmr": "XmR chart",
"chart.moreTypes": "More chart types",

"chart.setup": "Chart setup",

// SECONDARY MR / S CHART CONTROLS
"chart.secondary.showMR": "Show MR",
"chart.secondary.showS": "Show S chart",
"chart.secondary.mrDisplay": "MR display",
"chart.secondary.sDisplay": "S chart display",
"chart.secondary.includesSplits": "(includes splits)",
"chart.secondary.current": "Current",
"chart.secondary.all": "All",
"chart.secondary.helpAria": "Secondary chart display help",
"chart.secondary.helpText": "XmR charts can show a Moving Range (MR) chart, which helps explain the variation between consecutive measurements. X-bar and S charts can show an S chart, which displays variation within subgroups. Current shows the latest analysis period only. All shows the full history, including split periods. Use the checkbox to show or hide the secondary chart.",

// COLUMN SELECTION
"chart.columns.heading": "Choose columns:",
"chart.columns.xDefault": "Date / X-axis column",
"chart.columns.yDefault": "Value / Y-axis column",
"chart.columns.thirdDefault": "Denominator / opportunities",
"chart.columns.xmrY": "Measure (used for XmR limits)",
"chart.columns.cY": "Count (c) per time period",
"chart.columns.pY": "Numerator: defectives (d)",
"chart.columns.pThird": "Denominator: total (n)",
"chart.columns.pHint": "P chart plots a proportion: d out of n.",
"chart.columns.pprimeHint": "P′ chart plots a proportion: d out of n, with Laney-adjusted control limits.",
"chart.columns.uY": "Numerator: defects (c)",
"chart.columns.uThird": "Denominator: opportunities (n)",
"chart.columns.uHint": "U chart plots defects per opportunity: c per n.",
"chart.columns.uprimeHint": "U′ chart plots defects per opportunity, with Laney-adjusted control limits.",
"chart.columns.xbarsY": "Measurement value",
"chart.columns.xbarsThird": "Subgroup ID (e.g. day / week / sample)",
"chart.columns.xbarsHint": "X̄–S needs multiple measurements per subgroup. The horizontal axis uses subgroup order automatically.",
"chart.columns.tGapsY": "Time between events (e.g. days)",
"chart.columns.tDatesY": "Value column not used (T chart uses dates)",
"chart.columns.tGapsHint": "T chart plots time between rare events (using your numeric gaps).",
"chart.columns.tDatesHint": "T chart plots time between rare events (calculated from event dates).",
"chart.columns.gY": "Opportunities between events",
"chart.columns.gHint": "G chart plots opportunities between rare events.",
"chart.columns.paretoX": "Category column",

"chart.columns.paretoY": "Frequency / count column",

// ADDITIONAL CHART TYPES
"chart.types.c": "C chart (counts)",
"chart.types.p": "P chart (proportion)",
"chart.types.pprime": "P′ chart (Laney)",
"chart.types.u": "U chart (rate)",
"chart.types.uprime": "U′ chart (Laney)",
"chart.types.xbars": "X̄–S chart (subgroups)",
"chart.types.t": "T chart (time between)",
"chart.types.g": "G chart (opportunities between)",
"chart.types.pareto": "Pareto chart (categories)",

// RATE DISPLAY
"chart.rate.label": "Display rate as:",
"chart.rate.per1": "Per 1",
"chart.rate.per10": "Per 10",
"chart.rate.per100": "Per 100",
"chart.rate.per1000": "Per 1,000",
"chart.rate.per10000": "Per 10,000",
"chart.rate.hint": "Changes the displayed rate only; the underlying U-chart calculation is unchanged.",

// X-AXIS TYPE
"chart.axis.type": "X-axis type",
"chart.axis.helpAria": "X-axis type help",
"chart.axis.help": "Use Date / time for time-series SPC. Use Sequence / category to plot rows in their original order without treating the X-axis values as dates.",
"chart.axis.date": "Date / time",
"chart.axis.sequence": "Sequence / category",

// DATE FORMAT
"chart.dateFormat.label": "Date format for numeric dates",
"chart.dateFormat.helpAria": "Date format help",
"chart.dateFormat.help": "Used for ambiguous numeric dates such as 01/02/2024. ISO dates such as 2024-02-01 and Excel dates are recognised automatically.",
"chart.dateFormat.uk": "UK / day-first (dd/mm/yyyy)",
"chart.dateFormat.us": "US / month-first (mm/dd/yyyy)",
"chart.dateFormat.auto": "Auto-detect",

"chart.dateFormat.iso": "ISO / Excel dates only",

// CHART AVAILABILITY MESSAGES
"chart.availability.noData": "Load data to assess which chart types are available.",
"chart.availability.selectedColumn": "the selected value column",
"chart.availability.xmr": "Less than 12 data points available in \"{column}\" — this is a minimum requirement for a valid XmR chart.",
"chart.availability.c": "No non-negative whole-number count column is available yet — C charts need counts per time period.",
"chart.availability.p": "No suitable numerator/denominator pair is available yet — P charts need two whole-number columns where the numerator is part of the denominator.",
"chart.availability.pprime": "No suitable numerator/denominator pair is available yet — P′ charts need two whole-number columns where the numerator is part of the denominator.",
"chart.availability.u": "No suitable count/opportunities pair is available yet — U charts need a whole-number count column plus a positive opportunities column.",
"chart.availability.uprime": "No suitable count/opportunities pair is available yet — U′ charts need a whole-number count column plus a positive opportunities column.",
"chart.availability.xbars": "X̄–S charts need a numeric measurement column and a subgroup column with repeated subgroup labels.",
"chart.availability.t": "T charts need either event dates or a numeric time-between-events column.",
"chart.availability.g": "G charts need a whole-number column where all values are at least 1, because they plot opportunities between rare events.",

// DATE-FORMAT WARNINGS
"chart.warning.mixedDates": "Warning: this column appears to contain a mixture of UK-style and US-style numeric dates. Please standardise the dates if possible.",
"chart.warning.usDatesUkSelected": "Warning: these dates look like US month-first dates (mm/dd/yyyy), but UK day-first is selected. Some rows may be ignored. Try switching to US month-first or Auto-detect.",
"chart.warning.ukDatesUsSelected": "Warning: these dates look like UK day-first dates (dd/mm/yyyy), but US month-first is selected. Some rows may be ignored. Try switching to UK day-first or Auto-detect.",
"chart.warning.ambiguousUk": "This column contains ambiguous numeric dates. They will currently be interpreted as UK day-first dates (dd/mm/yyyy).",
"chart.warning.ambiguousUs": "This column contains ambiguous numeric dates. They will currently be interpreted as US month-first dates (mm/dd/yyyy).",
"chart.warning.ambiguousAuto": "This column contains ambiguous numeric dates and the tool cannot confidently auto-detect the style. It will fall back to UK day-first dates unless you choose another option.",
"chart.warning.ambiguousIso": "This column contains numeric slash/hyphen dates. In 'ISO / Excel dates only' mode, ambiguous numeric dates may not be interpreted as dates.",

// DISABLED DATE-FORMAT CONTROL

"chart.dateFormat.disabled": "Date-format selection is only available when the X-axis type is Date / time.",

// SHARED FORMATTING LABELS
"common.autoDefault": "Auto / default",
"common.auto": "Auto",
"common.bold": "Bold",
"common.italic": "Italic",
"common.underline": "Underline",

// LABELS PANEL
"labels.heading": "Chart labels",
"labels.title": "Chart title",
"labels.xAxis": "X-axis label",
"labels.yAxis": "Y-axis label",
"labels.titleSettings": "Title settings",
"labels.fontFamily": "Title font family",
"labels.fontSize": "Title font size",
"labels.showTitle": "Show chart title",

// AXES PANEL
"axes.heading": "Axes",
"axes.xSection": "X axis",
"axes.ySection": "Y axis",
"axes.xFontFamily": "X-axis font family",
"axes.xFontSize": "X-axis font size",
"axes.yMin": "Y-axis minimum",
"axes.yMax": "Y-axis maximum",
"axes.yRangeAria": "Y-axis range help",
"axes.yRangeHelp": "Leave the minimum and maximum blank to let the chart choose its range automatically.",
"axes.yFontFamily": "Y-axis font family",
"axes.yFontSize": "Y-axis font size",
"axes.yFormat": "Y-axis format",
"axes.formatAuto": "Auto",
"axes.formatDecimal": "Decimal",
"axes.formatInteger": "Integer",
"axes.formatPercent": "Percent",
"axes.decimals": "Y-axis decimal places",
"axes.decimalsAuto": "Auto",
"axes.dp0": "0 dp",
"axes.dp1": "1 dp",
"axes.dp2": "2 dp",
"axes.dp3": "3 dp",
"axes.tickInterval": "Y-axis tick interval",
"axes.tickAria": "Y-axis tick interval help",
"axes.tickHelp": "Leave blank for automatic spacing. For percentages, enter percentage points: for example, 1 means 1%, and 0.5 means 0.5%.",

// BASELINE PANEL
"baseline.heading": "Baseline data",
"baseline.helpAria": "Baseline data help",
"baseline.help": "A baseline is an initial reference period used to understand how your process was performing before later observations. Enter the number of points to use as your starting reference. SimpleSPC uses these to calculate the initial centre line and, where applicable, control limits. Later observations remain visible for comparison. Leave blank if you do not need a separate initial baseline period. Choose a baseline based on your knowledge of the process, not simply to remove inconvenient signals.",
"baseline.useFirst": "Use first",
"baseline.pointsAsBaseline": "points as baseline",
"baseline.blankHint": "(leave blank if no baseline period)",
"baseline.targetHeading": "Target",
"baseline.targetHelpAria": "How to set a target",
"baseline.targetHelp": "Enter the performance level you want to achieve, then choose whether values at or above, or at or below, that level are good. A target is a performance goal, not a statistical control limit.",
"baseline.targetHint": "Enter your target value, then choose whether higher or lower values are better.",
"baseline.targetValue": "Target value",
"baseline.atOrAbove": "At or above target is good",
"baseline.atOrBelow": "At or below target is good",
"baseline.hideTarget": "Hide target line",

"baseline.showTarget": "Show target line",

// RULES PANEL
"rules.heading": "Detection rules",
"rules.explainer": "Rule explainer",

"rules.shift": "Shift (run) length",
"rules.shiftAria": "Shift rule help",
"rules.shiftHelp": "Number of consecutive points on the same side of the centre line required to flag a shift. Default: 8.",

"rules.trend": "Trend length",
"rules.trendAria": "Trend rule help",
"rules.trendHelp": "Number of consecutive increases or decreases required to flag a trend. Default: 6. This setting is only used when the trend rule is available and enabled.",

"rules.advanced": "Advanced detection options",
"rules.enableTrend": "Enable trend rule",
"rules.enableTrendHelp": "Looks for a sustained sequence of consecutive increases or decreases. The required length is set by the Trend length control above.",
"rules.advancedCaution": "Advanced option — may increase false alerts.",

"rules.rareLabel": "Run & trend rules for T/G charts",
"rules.rareAria": "Rare event rules help",
"rules.rareHelp": "Enables run and trend detection on T and G rare-event charts. These chart types are naturally irregular, so pattern rules should be used cautiously.",
"rules.rareCaution": "Advanced option — rare-event charts can produce false run/trend signals.",

"rules.clamp": "Clamp LCL at zero",
"rules.clampAria": "Clamp LCL help",
"rules.clampHelp": "If the calculated lower control limit is below zero, display the LCL at zero instead. Useful for measures such as counts or times that cannot be negative.",

"rules.zoneIntro": "Optional zone rules (off by default). Use only on chart types where they are supported.",
"rules.zone23": "Apply “2 out of 3 in the outer third” rule",
"rules.zone45": "Apply “4 out of 5 beyond 1-sigma” rule",
"rules.zoneCaution": "Additional pattern rules can increase false alerts on XmR / individuals charts. They are off by default and should only be used when you understand the trade-off.",

"rules.mrHeading": "XmR limit calculation",
"rules.mrAria": "XmR moving range screening help",
"rules.mrHelp": "A very large jump between two consecutive points can make XmR control limits too wide. When this is on, SimpleSPC keeps all data points and moving ranges on the chart, but does not use unusually large moving ranges to estimate routine variation. Uses the Nelson method: moving ranges greater than 3.5 times the initial average moving range are excluded from the variation estimate.",
"rules.mrScreen": "Screen unusually large jumps when setting limits",

"rules.conservative": "This chart uses a conservative rule set designed for reliable interpretation.",
"rules.flag": "Flag special-cause signals on the chart",
"rules.colourBlind": "Colour-blind friendly colours and symbols",

// DYNAMIC RULE CONTROL TOOLTIPS
"rules.disabledTrend": "Trend rule is not offered for this chart type.",
"rules.enableTrendFirst": "Tick 'Enable trend rule' to use this setting.",
"rules.enableRareFirst": "Enable run & trend rules for this chart type to use this setting.",
"rules.rareOnly": "Only used for T and G charts.",
"rules.zoneBlocked": "Zone rules are not available for this chart type because they may create misleading alerts.",

// CONFIRMATIONS WHEN ENABLING RARE-EVENT RULES
"rules.rareConfirmImmediate": "T and G charts are naturally irregular.\n\nRun and trend rules may create false signals on this chart type and are off by default.\n\nOnly enable these rules if you understand this trade-off.",
"rules.rareConfirmOnce": "Advanced option: Run & trend rules on {chart} charts\n\nThese patterns often happen by chance on T/G charts and can increase false alerts.\n\nEnable anyway?"







});
