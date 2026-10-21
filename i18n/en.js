
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
"chart.columns.paretoY": "Frequency / count column"



});
