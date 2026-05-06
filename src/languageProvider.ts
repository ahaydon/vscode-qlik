import * as vscode from 'vscode';
import { ScriptSection } from './scriptModel';
import { QlikScriptFS } from './scriptFS';
import { QlikScriptTreeProvider } from './treeProvider';

// ── Shared symbol extraction ───────────────────────────────────────────────

interface ScriptSymbols {
  variables: string[];
  subs: string[];
  tables: string[];
}

const VAR_RE   = /(?:^|;)\s*(?:set|let)\s+([\w$][\w$.]*)\s*=/gim;
const SUB_DEF_RE = /(?:^|;)\s*sub\s+([\w$][\w$.]*)\s*(?:\(|\r?\n)/gim;
const TABLE_RE = /^([\w$][\w$.]*)\s*:/gm;
const CALL_RE  = /\bcall\s+([\w$][\w$.]*)/gim;

const KEYWORD_BLOCKLIST = new Set([
  'if', 'then', 'else', 'for', 'next', 'do', 'while', 'loop', 'sub', 'end',
  'load', 'select', 'from', 'where', 'set', 'let', 'store', 'join', 'keep',
  'drop', 'rename', 'map', 'unmap', 'call', 'exit', 'switch', 'case', 'when',
  'unless', 'in', 'not', 'and', 'or', 'xor', 'like', 'match',
]);

function unique(names: string[]): string[] {
  return [...new Set(names)];
}

export function parseScriptSymbols(text: string): ScriptSymbols {
  const variables: string[] = [];
  const subs: string[] = [];
  const tables: string[] = [];

  let m: RegExpExecArray | null;

  VAR_RE.lastIndex = 0;
  while ((m = VAR_RE.exec(text)) !== null) variables.push(m[1]);

  SUB_DEF_RE.lastIndex = 0;
  while ((m = SUB_DEF_RE.exec(text)) !== null) subs.push(m[1]);

  TABLE_RE.lastIndex = 0;
  while ((m = TABLE_RE.exec(text)) !== null) {
    if (!KEYWORD_BLOCKLIST.has(m[1].toLowerCase())) tables.push(m[1]);
  }

  return { variables: unique(variables), subs: unique(subs), tables: unique(tables) };
}

// ── Static completion items ────────────────────────────────────────────────

const FUNCTION_NAMES = [
  // Aggregation - basic
  'FirstSortedValue', 'Max', 'Min', 'Mode', 'Only', 'Sum',
  // Aggregation - counter
  'Count', 'MissingCount', 'NullCount', 'NumericCount', 'TextCount',
  // Aggregation - financial
  'Irr', 'Npv', 'Xirr', 'Xnpv',
  // Aggregation - statistical
  'Avg', 'Correl', 'Fractile', 'FractileExc', 'Kurtosis',
  'LINEST_B', 'LINEST_DF', 'LINEST_F', 'LINEST_M', 'LINEST_R2',
  'LINEST_SEB', 'LINEST_SEM', 'LINEST_SEY', 'LINEST_SSREG', 'LINEST_SSRESID',
  'Median', 'MutualInfo', 'Skew', 'Stdev', 'Sterr', 'STEYX', 'Variance',
  // Aggregation - string
  'Concat', 'FirstValue', 'LastValue', 'MaxString', 'MinString',
  // Aggregation - statistical tests (chi2, ttest, ztest)
  'Chi2Test_chi2', 'Chi2Test_df', 'Chi2Test_p',
  'TTest_conf', 'TTest_df', 'TTest_dif', 'TTest_lower', 'TTest_sig', 'TTest_sterr', 'TTest_t', 'TTest_upper',
  'TTestw_conf', 'TTestw_df', 'TTestw_dif', 'TTestw_lower', 'TTestw_sig', 'TTestw_sterr', 'TTestw_t', 'TTestw_upper',
  'TTest1_conf', 'TTest1_df', 'TTest1_dif', 'TTest1_lower', 'TTest1_sig', 'TTest1_sterr', 'TTest1_t', 'TTest1_upper',
  'TTest1w_conf', 'TTest1w_df', 'TTest1w_dif', 'TTest1w_lower', 'TTest1w_sig', 'TTest1w_sterr', 'TTest1w_t', 'TTest1w_upper',
  'ZTest_conf', 'ZTest_dif', 'ZTest_lower', 'ZTest_sig', 'ZTest_sterr', 'ZTest_upper', 'ZTest_z',
  'ZTestw_conf', 'ZTestw_dif', 'ZTestw_lower', 'ZTestw_sig', 'ZTestw_sterr', 'ZTestw_upper', 'ZTestw_z',
  // Color
  'Argb', 'Hsl', 'IsValidColor', 'Rgb',
  // Conditional
  'Alt', 'Class', 'Coalesce', 'If', 'Match', 'MixMatch', 'Pick', 'WildMatch', 'WildMatch5',
  // Counter
  'AutoNumber', 'AutoNumberHash128', 'AutoNumberHash256', 'IterNo', 'RecNo', 'RowNo',
  // Date and time
  'AddMonths', 'AddYears', 'Age', 'ConvertToLocalTime',
  'Day', 'DaylightSaving', 'DayEnd', 'DayName', 'DayNumberOfQuarter', 'DayNumberOfYear', 'DayStart',
  'FirstWorkDate', 'Gmt', 'Hour',
  'InDay', 'InDayToTime', 'InLunarWeek', 'InLunarWeekToDate',
  'InMonth', 'InMonths', 'InMonthsToDate', 'InMonthToDate',
  'InQuarter', 'InQuarterToDate',
  'InWeek', 'InWeekToDate',
  'InYear', 'InYearToDate',
  'LastWorkDate', 'LocalTime',
  'LunarWeekEnd', 'LunarWeekName', 'LunarWeekStart',
  'MakeDate', 'MakeTime', 'MakeWeekDate', 'Minute',
  'Month', 'MonthEnd', 'MonthName', 'MonthsEnd', 'MonthsName', 'MonthsStart', 'MonthStart',
  'NetworkDays', 'Now',
  'QuarterEnd', 'QuarterName', 'QuarterStart',
  'Second', 'SetDateYear', 'SetDateYearMonth',
  'Timezone', 'Today', 'Utc',
  'Week', 'WeekDay', 'WeekEnd', 'WeekName', 'WeekStart', 'WeekYear',
  'Year', 'YearEnd', 'YearName', 'YearStart', 'YearToDate',
  // Exponential and logarithmic
  'E', 'Exp', 'Log', 'Log10', 'Pow', 'Sqrt',
  // Field / inter-record
  'Exists', 'FieldIndex', 'FieldValue', 'FieldValueCount', 'LookUp', 'NoOfRows', 'Peek', 'Previous',
  // File
  'Attribute', 'ConnectString',
  'FileBaseName', 'FileDir', 'FileExtension', 'FileName', 'FilePath', 'FileSize', 'FileTime',
  'GetFolderPath', 'GetRegistryString',
  'QvdCreateTime', 'QvdFieldName', 'QvdNoOfFields', 'QvdNoOfRecords', 'QvdTableName',
  // Financial
  'BlackAndSchole', 'Fv', 'Nper', 'Pmt', 'Pv', 'Rate',
  // Formatting
  'ApplyCodepage', 'Date', 'Dual', 'Interval', 'Money', 'Num', 'Time', 'Timestamp',
  // General numeric
  'Abs', 'BitCount', 'Ceil', 'Combin', 'Div', 'Even',
  'Fabs', 'Fact', 'Floor', 'Fmod', 'Frac',
  'Mod', 'Odd', 'Permut', 'Pi', 'Rand', 'Rem', 'Round', 'Sign',
  // Geospatial
  'GeoAggrGeometry', 'GeoBoundingBox', 'GeoCountVertex',
  'GeoGetBoundingBox', 'GeoGetPolygonCenter', 'GeoInvProjectGeometry',
  'GeoMakePoint', 'GeoProject', 'GeoProjectGeometry', 'GeoReduceGeometry',
  // Interpretation
  'Date#', 'Interval#', 'Money#', 'Num#', 'Text', 'Time#', 'Timestamp#',
  // JSON
  'IsJson', 'JsonGet', 'JsonSet', 'JsonSetEx',
  // Logical
  'IsNum', 'IsText',
  // Mapping
  'ApplyMap', 'MapSubString',
  // Null
  'EmptyIsNull', 'IsNull', 'Null',
  // Range
  'RangeApp', 'RangeAvg', 'RangeCorrel', 'RangeCount', 'RangeFractile',
  'RangeIrr', 'RangeKurtosis', 'RangeMax', 'RangeMaxString',
  'RangeMin', 'RangeMinString', 'RangeMissingCount', 'RangeMode',
  'RangeNullCount', 'RangeNumericCount', 'RangeNpv', 'RangeOnly',
  'RangeSkew', 'RangeStdev', 'RangeSum', 'RangeTextCount',
  'RangeXirr', 'RangeXnpv',
  // Relational / ML
  'HRank', 'KMeans2D', 'KMeansCentroid2D', 'KMeansCentroidND', 'KMeansND',
  'Rank', 'STL_Residual', 'STL_Seasonal', 'STL_Trend',
  // Statistical distribution
  'BetaDensity', 'BetaDist', 'BetaInv',
  'BinomDist', 'BinomFrequency', 'BinomInv',
  'ChiDensity', 'ChiDist', 'ChiInv',
  'FDensity', 'FDist', 'FInv',
  'GammaDensity', 'GammaDist', 'GammaInv',
  'NormDist', 'NormInv',
  'PoissonDist', 'PoissonFrequency', 'PoissonInv',
  'TDensity', 'TDist', 'TInv',
  // String
  'Capitalize', 'Chr', 'CountRegex', 'Evaluate',
  'ExtractRegex', 'ExtractRegexGroup', 'FindOneOfValue',
  'Hash128', 'Hash160', 'Hash256',
  'Index', 'IndexRegex', 'IndexRegexGroup',
  'IsEmpty', 'IsRegex',
  'KeepChar', 'Left', 'Len', 'LevenshteinDist', 'Lower', 'LTrim',
  'MatchRegex', 'Mid', 'MsgBox', 'Ord',
  'PurgeChar', 'Repeat', 'Replace', 'ReplaceRegex', 'ReplaceRegexGroup',
  'Right', 'RTrim',
  'SubField', 'SubFieldRegex', 'SubstringCount',
  'TextBetween', 'Trim', 'Upper',
  // System
  'Author', 'CalcDim', 'ClientPlatform', 'ComputerName',
  'DocumentName', 'DocumentPath', 'DocumentTitle',
  'EngineVersion', 'GetCollationLocale', 'GetObjectField',
  'GetSysAttr', 'GetUserAttr',
  'GroupDimensionIndex', 'GroupDimensionLabel',
  'InObject', 'IsPartialReload',
  'ObjectId', 'OSUser', 'ProductVersion', 'ReloadTime', 'StateName',
  // Table
  'FieldName', 'FieldNumber', 'NoOfFields',
  // Trigonometric and hyperbolic
  'Acos', 'Acosh', 'Asin', 'Asinh', 'Atan', 'Atan2', 'Atanh',
  'Cos', 'Cosh', 'Sin', 'Sinh', 'Tan', 'Tanh',
  // Window
  'Window', 'WRank',
];

const KEYWORD_NAMES = [
  // Control
  'IF', 'THEN', 'ELSEIF', 'ELSE', 'END IF',
  'FOR', 'FOR EACH', 'NEXT', 'TO', 'STEP',
  'DO', 'WHILE', 'LOOP', 'UNTIL',
  'EXIT', 'EXIT SCRIPT', 'EXIT FOR', 'EXIT DO',
  'SUB', 'END SUB', 'CALL',
  'SWITCH', 'CASE', 'DEFAULT', 'END SWITCH',
  'WHEN', 'UNLESS', 'IN',
  // Statement
  'LOAD', 'SELECT', 'FROM', 'WHERE', 'GROUP BY', 'ORDER BY',
  'HAVING', 'LIMIT', 'UNION', 'SET', 'LET', 'STORE', 'INTO',
  'RESIDENT', 'AUTOGENERATE', 'INLINE', 'EXTENSION',
  'CONNECT', 'USING', 'MAP', 'UNMAP', 'ALIAS', 'DIRECTORY', 'EXECUTE',
  'TAG', 'UNTAG', 'QUALIFY', 'UNQUALIFY',
  'RENAME FIELD', 'RENAME TABLE',
  'DROP FIELD', 'DROP TABLE', 'DROP FIELDS', 'DROP TABLES',
  'LOOSEN TABLE',
  'COMMENT FIELD', 'COMMENT TABLE',
  'BINARY', 'SECTION', 'ACCESS', 'APPLICATION',
  'DECLARE', 'DERIVE', 'SQL', 'DISCONNECT',
  'FIELD', 'FIELDS', 'DISTINCT', 'AS', 'WITH', 'FORMAT IS',
  'TRACE', 'SLEEP', 'FLUSH LOG',
  // Prefix
  'JOIN', 'INNER JOIN', 'LEFT JOIN', 'RIGHT JOIN', 'OUTER JOIN',
  'KEEP', 'INNER KEEP', 'LEFT KEEP', 'RIGHT KEEP',
  'CONCATENATE', 'NOCONCATENATE', 'ADD', 'REPLACE', 'MERGE',
  'PARTIAL RELOAD', 'FIRST', 'SAMPLE', 'BUFFER', 'MAPPING',
  'INTERVALMATCH', 'CROSSTABLE', 'GENERIC', 'SEMANTIC',
  'HIERARCHY', 'HIERARCHYBELONGSTO',
  // Clause / operators
  'AND', 'OR', 'NOT', 'XOR', 'LIKE', 'MATCH', 'MIXMATCH',
  'WILDMATCH', 'ASC', 'DESC', 'TOTAL', 'ALL',
];

const CONSTANT_NAMES = ['TRUE', 'FALSE', 'NULL', 'NAN'];

const FUNCTION_NAMES_LOWER = new Set(FUNCTION_NAMES.map(f => f.toLowerCase().replace('#', '')));

function makeFunction(name: string): vscode.CompletionItem {
  const item = new vscode.CompletionItem(name.toUpperCase(), vscode.CompletionItemKind.Function);
  item.insertText = new vscode.SnippetString(`${name.toUpperCase()}($1)`);
  item.detail = 'Built-in function';
  item.sortText = `2_${name.toLowerCase()}`;
  return item;
}

function makeKeyword(name: string): vscode.CompletionItem {
  const item = new vscode.CompletionItem(name, vscode.CompletionItemKind.Keyword);
  item.insertText = name;
  item.detail = 'Keyword';
  item.sortText = `3_${name.toLowerCase()}`;
  return item;
}

function makeConstant(name: string): vscode.CompletionItem {
  const item = new vscode.CompletionItem(name, vscode.CompletionItemKind.Constant);
  item.sortText = `3_${name.toLowerCase()}`;
  return item;
}

const FUNCTIONS: vscode.CompletionItem[] = FUNCTION_NAMES.map(makeFunction);
const KEYWORDS: vscode.CompletionItem[]  = KEYWORD_NAMES.map(makeKeyword);
const CONSTANTS: vscode.CompletionItem[] = CONSTANT_NAMES.map(makeConstant);

// ── Hover documentation ────────────────────────────────────────────────────

interface FunctionDoc { signature: string; doc: string; }

const FUNCTION_DOCS = new Map<string, FunctionDoc>([
  ['sum',         { signature: 'Sum([{set}] [DISTINCT] [TOTAL [<fld>]] expr)', doc: 'Returns the sum of values in the expression over the aggregated data.' }],
  ['count',       { signature: 'Count([{set}] [DISTINCT] [TOTAL [<fld>]] expr)', doc: 'Returns the count of values in the expression.' }],
  ['avg',         { signature: 'Avg([{set}] [DISTINCT] [TOTAL [<fld>]] expr)', doc: 'Returns the average of values in the expression.' }],
  ['min',         { signature: 'Min([{set}] [TOTAL [<fld>]] expr [, rank])', doc: 'Returns the lowest numeric value in the expression.' }],
  ['max',         { signature: 'Max([{set}] [TOTAL [<fld>]] expr [, rank])', doc: 'Returns the highest numeric value in the expression.' }],
  ['if',          { signature: 'If(condition, then [, else])', doc: 'Returns then-value if condition is true, else-value otherwise.' }],
  ['coalesce',    { signature: 'Coalesce(expr1, expr2 [, ...])', doc: 'Returns the first non-null value from the argument list.' }],
  ['only',        { signature: 'Only([{set}] [TOTAL [<fld>]] expr)', doc: 'Returns a value if there is exactly one possible result, otherwise null.' }],
  ['concat',      { signature: 'Concat([{set}] [DISTINCT] [TOTAL [<fld>]] string [, delimiter [, sort_weight]])', doc: 'Concatenates string values over aggregated rows.' }],
  ['maxstring',   { signature: 'MaxString([{set}] [TOTAL [<fld>]] expr)', doc: 'Returns the last text value in the expression sorted alphabetically.' }],
  ['minstring',   { signature: 'MinString([{set}] [TOTAL [<fld>]] expr)', doc: 'Returns the first text value in the expression sorted alphabetically.' }],
  ['firstsortedvalue', { signature: 'FirstSortedValue([{set}] [DISTINCT] [TOTAL [<fld>]] expr, sort_weight [, rank])', doc: 'Returns the value of expr corresponding to the lowest sort_weight.' }],
  ['stdev',       { signature: 'StDev([{set}] [DISTINCT] [TOTAL [<fld>]] expr)', doc: 'Returns the standard deviation of values in the expression.' }],
  ['median',      { signature: 'Median([{set}] [TOTAL [<fld>]] expr)', doc: 'Returns the median of values in the expression.' }],
  ['mode',        { signature: 'Mode([{set}] [TOTAL [<fld>]] expr)', doc: 'Returns the most common value in the expression.' }],
  ['fractile',    { signature: 'Fractile([{set}] [TOTAL [<fld>]] expr, fraction)', doc: 'Returns the value at the given fraction (0–1) in the sorted distribution.' }],
  ['left',        { signature: 'Left(text, count)', doc: 'Returns the leftmost count characters of text.' }],
  ['right',       { signature: 'Right(text, count)', doc: 'Returns the rightmost count characters of text.' }],
  ['mid',         { signature: 'Mid(text, start [, count])', doc: 'Returns count characters from text starting at start (1-based).' }],
  ['len',         { signature: 'Len(text)', doc: 'Returns the number of characters in text.' }],
  ['trim',        { signature: 'Trim(text)', doc: 'Removes leading and trailing spaces from text.' }],
  ['ltrim',       { signature: 'LTrim(text)', doc: 'Removes leading spaces from text.' }],
  ['rtrim',       { signature: 'RTrim(text)', doc: 'Removes trailing spaces from text.' }],
  ['upper',       { signature: 'Upper(text)', doc: 'Converts text to uppercase.' }],
  ['lower',       { signature: 'Lower(text)', doc: 'Converts text to lowercase.' }],
  ['capitalize',  { signature: 'Capitalize(text)', doc: 'Returns text with the first letter of each word capitalised.' }],
  ['index',       { signature: 'Index(text, substring [, count])', doc: 'Returns the position of the nth occurrence of substring in text. Negative count searches from the right.' }],
  ['replace',     { signature: 'Replace(text, from, to)', doc: 'Replaces all occurrences of from with to in text.' }],
  ['keepchar',    { signature: 'KeepChar(text, keep_chars)', doc: 'Returns text with all characters not in keep_chars removed.' }],
  ['purgechar',   { signature: 'PurgeChar(text, remove_chars)', doc: 'Returns text with all characters in remove_chars removed.' }],
  ['subfield',    { signature: 'SubField(text, delimiter [, field_no])', doc: 'Extracts substrings split by delimiter. Returns the nth field if field_no is given, otherwise generates one row per field.' }],
  ['substringcount', { signature: 'SubStringCount(text, substring)', doc: 'Returns the number of non-overlapping occurrences of substring in text.' }],
  ['repeat',      { signature: 'Repeat(text, count)', doc: 'Returns text repeated count times.' }],
  ['textbetween', { signature: 'TextBetween(text, start, end [, count])', doc: 'Returns the text between the nth occurrence of start and end in text.' }],
  ['chr',         { signature: 'Chr(int)', doc: 'Returns the Unicode character for the given code point.' }],
  ['ord',         { signature: 'Ord(text)', doc: 'Returns the Unicode code point of the first character of text.' }],
  ['num',         { signature: 'Num(number [, format [, dec_sep [, thou_sep]]])', doc: 'Formats a number using the given format string.' }],
  ['num#',        { signature: 'Num#(text [, format [, dec_sep [, thou_sep]]])', doc: 'Interprets text as a number using the given format.' }],
  ['date',        { signature: 'Date(number [, format])', doc: 'Formats a number as a date using the given format string.' }],
  ['date#',       { signature: 'Date#(text [, format])', doc: 'Interprets text as a date and returns a numeric value.' }],
  ['time',        { signature: 'Time(number [, format])', doc: 'Formats a number as a time using the given format string.' }],
  ['timestamp',   { signature: 'Timestamp(number [, format])', doc: 'Formats a number as a date and time.' }],
  ['interval',    { signature: 'Interval(number [, format])', doc: 'Formats a number as a time interval.' }],
  ['money',       { signature: 'Money(number [, format [, dec_sep [, thou_sep]]])', doc: 'Formats a number as a monetary value.' }],
  ['today',       { signature: 'Today([timer_mode])', doc: 'Returns the current date. timer_mode=1 returns a real-time value.' }],
  ['now',         { signature: 'Now([timer_mode])', doc: 'Returns the current timestamp. timer_mode=1 returns a real-time value.' }],
  ['makedate',    { signature: 'MakeDate(year [, month [, day]])', doc: 'Constructs a date value from year, month, and day components.' }],
  ['maketime',    { signature: 'MakeTime(hour [, minute [, second]])', doc: 'Constructs a time value from hour, minute, and second components.' }],
  ['addmonths',   { signature: 'AddMonths(startdate, months [, mode])', doc: 'Returns the date that is the given number of months after startdate.' }],
  ['addyears',    { signature: 'AddYears(startdate, years)', doc: 'Returns the date that is the given number of years after startdate.' }],
  ['monthstart',  { signature: 'MonthStart(date [, period_no])', doc: 'Returns the first millisecond of the month containing date.' }],
  ['monthend',    { signature: 'MonthEnd(date [, period_no])', doc: 'Returns the last millisecond of the month containing date.' }],
  ['yearstart',   { signature: 'YearStart(date [, period_no [, first_month_of_year]])', doc: 'Returns the first millisecond of the year containing date.' }],
  ['yearend',     { signature: 'YearEnd(date [, period_no [, first_month_of_year]])', doc: 'Returns the last millisecond of the year containing date.' }],
  ['weekstart',   { signature: 'WeekStart(date [, period_no [, first_week_day]])', doc: 'Returns the first millisecond of the week containing date.' }],
  ['weekend',     { signature: 'WeekEnd(date [, period_no [, first_week_day]])', doc: 'Returns the last millisecond of the week containing date.' }],
  ['daystart',    { signature: 'DayStart(date [, period_no [, day_start]])', doc: 'Returns the first millisecond of the day containing date.' }],
  ['dayend',      { signature: 'DayEnd(date [, period_no [, day_start]])', doc: 'Returns the last millisecond of the day containing date.' }],
  ['networkdays', { signature: 'NetworkDays(start_date, end_date [, holiday])', doc: 'Returns the number of working days between two dates, excluding weekends and optional holidays.' }],
  ['age',         { signature: 'Age(timestamp, date_of_birth)', doc: 'Returns the age in whole years at timestamp for someone born on date_of_birth.' }],
  ['year',        { signature: 'Year(date)', doc: 'Returns the year component of date.' }],
  ['month',       { signature: 'Month(date)', doc: 'Returns the month name of date.' }],
  ['day',         { signature: 'Day(date)', doc: 'Returns the day-of-month component of date.' }],
  ['hour',        { signature: 'Hour(time)', doc: 'Returns the hour component of time.' }],
  ['minute',      { signature: 'Minute(time)', doc: 'Returns the minute component of time.' }],
  ['second',      { signature: 'Second(time)', doc: 'Returns the second component of time.' }],
  ['week',        { signature: 'Week(date [, first_week_day [, broken_weeks [, reference_day]]])', doc: 'Returns the week number of date.' }],
  ['weekday',     { signature: 'WeekDay(date)', doc: 'Returns the day of week (0=Mon … 6=Sun) for date.' }],
  ['weekyear',    { signature: 'WeekYear(date [, first_week_day [, broken_weeks [, reference_day]]])', doc: 'Returns the year the ISO week number belongs to.' }],
  ['floor',       { signature: 'Floor(number [, step [, offset]])', doc: 'Rounds down to the nearest multiple of step.' }],
  ['ceil',        { signature: 'Ceil(number [, step [, offset]])', doc: 'Rounds up to the nearest multiple of step.' }],
  ['frac',        { signature: 'Frac(number)', doc: 'Returns the fractional part of number.' }],
  ['mod',         { signature: 'Mod(dividend, divisor)', doc: 'Returns the modulo (remainder) of dividend ÷ divisor, always ≥ 0.' }],
  ['div',         { signature: 'Div(dividend, divisor)', doc: 'Returns the integer quotient of dividend ÷ divisor.' }],
  ['abs',         { signature: 'Abs(number)', doc: 'Returns the absolute value of number.' }],
  ['sqrt',        { signature: 'Sqrt(number)', doc: 'Returns the square root of number.' }],
  ['even',        { signature: 'Even(number)', doc: 'Rounds number up to the nearest even integer.' }],
  ['odd',         { signature: 'Odd(number)', doc: 'Rounds number up to the nearest odd integer.' }],
  ['sign',        { signature: 'Sign(number)', doc: 'Returns 1, 0, or -1 depending on the sign of number.' }],
  ['pow',         { signature: 'Pow(base, exp)', doc: 'Returns base raised to the power exp.' }],
  ['exp',         { signature: 'Exp(number)', doc: 'Returns e raised to the power number.' }],
  ['log',         { signature: 'Log(number)', doc: 'Returns the natural logarithm of number.' }],
  ['log10',       { signature: 'Log10(number)', doc: 'Returns the base-10 logarithm of number.' }],
  ['sin',         { signature: 'Sin(number)', doc: 'Returns the sine of number (radians).' }],
  ['cos',         { signature: 'Cos(number)', doc: 'Returns the cosine of number (radians).' }],
  ['tan',         { signature: 'Tan(number)', doc: 'Returns the tangent of number (radians).' }],
  ['asin',        { signature: 'Asin(number)', doc: 'Returns the arcsine of number in radians.' }],
  ['acos',        { signature: 'Acos(number)', doc: 'Returns the arccosine of number in radians.' }],
  ['atan',        { signature: 'Atan(number)', doc: 'Returns the arctangent of number in radians.' }],
  ['atan2',       { signature: 'Atan2(y, x)', doc: 'Returns the angle in radians between the positive x-axis and the point (x, y).' }],
  ['rand',        { signature: 'Rand()', doc: 'Returns a pseudo-random number in the range [0, 1).' }],
  ['pi',          { signature: 'Pi()', doc: 'Returns the value of π.' }],
  ['e',           { signature: 'E()', doc: "Returns Euler's number e ≈ 2.71828." }],
  ['applymap',    { signature: 'ApplyMap(mapname, expr [, default])', doc: 'Looks up expr in a mapping table. Returns default if not found.' }],
  ['mapsubstring',{ signature: 'MapSubString(mapname, expr)', doc: 'Applies a mapping to all substrings of expr that match keys in the mapping table.' }],
  ['autonumber',  { signature: 'AutoNumber(expr [, AutoID])', doc: 'Assigns a unique sequential integer to each distinct value of expr.' }],
  ['fieldvalue',  { signature: 'FieldValue(field_name, row_no)', doc: 'Returns the value at position row_no in field_name (1-based).' }],
  ['fieldindex',  { signature: 'FieldIndex(field_name, value)', doc: 'Returns the row position of value in field_name, or 0 if not found.' }],
  ['fieldvaluecount', { signature: 'FieldValueCount(field_name)', doc: 'Returns the number of distinct values in field_name.' }],
  ['peek',        { signature: 'Peek(field_name [, row [, table_name]])', doc: 'Returns the value of a field in a previously loaded row. row=-1 is the last row.' }],
  ['recno',       { signature: 'RecNo()', doc: 'Returns the current record number within the table being read (resets per input source).' }],
  ['rowno',       { signature: 'RowNo([TOTAL])', doc: 'Returns the current row number in the internal table being generated.' }],
  ['exists',      { signature: 'Exists(field_name [, expr])', doc: 'Returns true if the value exists in the loaded data for field_name.' }],
  ['isnum',       { signature: 'IsNum(expr)', doc: 'Returns true if expr can be interpreted as a numeric value.' }],
  ['istext',      { signature: 'IsText(expr)', doc: 'Returns true if expr has a text representation.' }],
  ['isnull',      { signature: 'IsNull(expr)', doc: 'Returns true if expr is null.' }],
  ['null',        { signature: 'Null()', doc: 'Returns a null value.' }],
  ['hash128',     { signature: 'Hash128(expr1 [, expr2, ...])', doc: 'Returns a 128-bit hash as a string of the concatenated string values.' }],
  ['hash160',     { signature: 'Hash160(expr1 [, expr2, ...])', doc: 'Returns a 160-bit hash as a string.' }],
  ['hash256',     { signature: 'Hash256(expr1 [, expr2, ...])', doc: 'Returns a 256-bit hash as a string.' }],

  // Interpretation (# variants)
  ['time#',       { signature: 'Time#(text [, format])', doc: 'Interprets text as a time and returns its numeric representation.' }],
  ['timestamp#',  { signature: 'Timestamp#(text [, format])', doc: 'Interprets text as a date+time and returns its numeric representation.' }],
  ['interval#',   { signature: 'Interval#(text [, format])', doc: 'Interprets text as a time interval and returns its numeric representation.' }],
  ['money#',      { signature: 'Money#(text [, format [, dec_sep [, thou_sep]]])', doc: 'Interprets text as a monetary value and returns its numeric value.' }],
  ['text',        { signature: 'Text(expr)', doc: 'Forces expr to be treated as text, preventing numeric interpretation.' }],

  // Formatting
  ['dual',        { signature: 'Dual(text, number)', doc: 'Returns a dual value with both a text and a numeric representation.' }],
  ['applycodepage', { signature: 'ApplyCodepage(text, codepage)', doc: 'Applies the given character encoding (codepage number) to text and returns a Unicode string.' }],

  // Color
  ['argb',        { signature: 'ARGB(alpha, r, g, b)', doc: 'Returns a color representation from alpha and RGB components (0–255 each); alpha=0 is fully transparent.' }],
  ['rgb',         { signature: 'RGB(r, g, b)', doc: 'Returns a color representation from red, green, and blue components (0–255 each).' }],
  ['hsl',         { signature: 'HSL(hue, saturation, luminosity)', doc: 'Returns a color representation from hue (0–1), saturation (0–1), and luminosity (0–1).' }],
  ['isvalidcolor', { signature: 'IsValidColor(expr)', doc: 'Returns true if expr is a valid color representation (color name or ARGB code).' }],

  // Conditional
  ['alt',         { signature: 'Alt(expr1 [, expr2 [, ...]])', doc: 'Returns the first argument that evaluates to a valid (non-null, non-NaN) result.' }],
  ['class',       { signature: 'Class(value, size [, label [, offset]])', doc: 'Returns the class interval label containing value, where each interval has width size.' }],
  ['pick',        { signature: 'Pick(n, expr1 [, expr2, ...])', doc: 'Returns the nth expression in the argument list.' }],
  ['match',       { signature: 'Match(str, expr1 [, expr2, ...])', doc: 'Returns the 1-based index of the first case-sensitive match of str in the list; 0 if none.' }],
  ['mixmatch',    { signature: 'MixMatch(str, expr1 [, expr2, ...])', doc: 'Case-insensitive version of Match; returns the 1-based index of the first match.' }],
  ['wildmatch',   { signature: 'WildMatch(str, pattern1 [, pattern2, ...])', doc: 'Returns the 1-based index of the first wildcard pattern matching str (* and ?); 0 if none.' }],
  ['wildmatch5',  { signature: 'WildMatch5(str, pattern1 [, pattern2, ...])', doc: 'SQL-style wildcard match using % (any chars) and _ (one char); returns the 1-based matching index.' }],

  // Counter
  ['autonumberhash128', { signature: 'AutoNumberHash128(expr1 [, expr2, ...])', doc: 'Assigns a unique sequential integer to each distinct combination of values using a 128-bit hash.' }],
  ['autonumberhash256', { signature: 'AutoNumberHash256(expr1 [, expr2, ...])', doc: 'Assigns a unique sequential integer to each distinct combination of values using a 256-bit hash.' }],
  ['iterno',      { signature: 'IterNo()', doc: 'Returns the current iteration number inside a DO...LOOP with the WHILE or UNTIL clause; 1 on the first pass.' }],

  // Aggregation – counter
  ['missingcount', { signature: 'MissingCount([{set}] [DISTINCT] [TOTAL [<fld>]] expr)', doc: 'Returns the count of non-numeric, non-null (missing) values in the expression.' }],
  ['nullcount',   { signature: 'NullCount([{set}] [DISTINCT] [TOTAL [<fld>]] expr)', doc: 'Returns the count of null values in the expression.' }],
  ['numericcount', { signature: 'NumericCount([{set}] [DISTINCT] [TOTAL [<fld>]] expr)', doc: 'Returns the count of numeric values in the expression.' }],
  ['textcount',   { signature: 'TextCount([{set}] [DISTINCT] [TOTAL [<fld>]] expr)', doc: 'Returns the count of text (non-numeric, non-null) values in the expression.' }],

  // Aggregation – financial
  ['irr',         { signature: 'Irr([{set}] [TOTAL [<fld>]] value)', doc: 'Returns the internal rate of return for a series of cash flows represented by value.' }],
  ['npv',         { signature: 'Npv([{set}] [TOTAL [<fld>]] discount_rate, value)', doc: 'Returns the net present value of an investment given a discount rate and series of cash flows.' }],
  ['xirr',        { signature: 'Xirr([{set}] [TOTAL [<fld>]] value, date)', doc: 'Returns the internal rate of return for a schedule of non-periodic cash flows.' }],
  ['xnpv',        { signature: 'Xnpv([{set}] [TOTAL [<fld>]] discount_rate, value, date)', doc: 'Returns the net present value for a schedule of non-periodic cash flows.' }],

  // Aggregation – statistical
  ['correl',      { signature: 'Correl([{set}] [TOTAL [<fld>]] x_value, y_value)', doc: 'Returns the Pearson correlation coefficient between two aggregated series of values.' }],
  ['fractileexc', { signature: 'FractileExc([{set}] [TOTAL [<fld>]] expr, fraction)', doc: 'Returns the exclusive fractile at the given fraction (0–1) using interpolation between data points.' }],
  ['kurtosis',    { signature: 'Kurtosis([{set}] [TOTAL [<fld>]] expr)', doc: 'Returns the excess kurtosis of values in the expression.' }],
  ['skew',        { signature: 'Skew([{set}] [TOTAL [<fld>]] expr)', doc: 'Returns the skewness of values in the expression.' }],
  ['sterr',       { signature: 'Sterr([{set}] [TOTAL [<fld>]] expr)', doc: 'Returns the standard error of the mean for aggregated values.' }],
  ['steyx',       { signature: 'STEYX([{set}] [TOTAL [<fld>]] y_value, x_value)', doc: 'Returns the standard error of the predicted y-value for each x in a linear regression.' }],
  ['variance',    { signature: 'Variance([{set}] [TOTAL [<fld>]] expr)', doc: 'Returns the variance (square of standard deviation) of values in the expression.' }],
  ['mutualinfo',  { signature: 'MutualInfo([{set}] [TOTAL [<fld>]] col_expr, row_expr [, data_type [, sample [, epsilon]]])', doc: 'Returns the mutual information between two fields, measuring their statistical dependency.' }],

  // Aggregation – LINEST regression
  ['linest_b',    { signature: 'LINEST_B([{set}] [TOTAL [<fld>]] y_value, x_value [, y0 [, x0]])', doc: 'Returns the y-intercept (b) of the linear regression line y = m·x + b.' }],
  ['linest_df',   { signature: 'LINEST_DF([{set}] [TOTAL [<fld>]] y_value, x_value [, y0 [, x0]])', doc: 'Returns the degrees of freedom for the linear regression.' }],
  ['linest_f',    { signature: 'LINEST_F([{set}] [TOTAL [<fld>]] y_value, x_value [, y0 [, x0]])', doc: 'Returns the F statistic for the linear regression.' }],
  ['linest_m',    { signature: 'LINEST_M([{set}] [TOTAL [<fld>]] y_value, x_value [, y0 [, x0]])', doc: 'Returns the slope (m) of the linear regression line y = m·x + b.' }],
  ['linest_r2',   { signature: 'LINEST_R2([{set}] [TOTAL [<fld>]] y_value, x_value [, y0 [, x0]])', doc: 'Returns the coefficient of determination (R²) for the linear regression.' }],
  ['linest_seb',  { signature: 'LINEST_SEB([{set}] [TOTAL [<fld>]] y_value, x_value [, y0 [, x0]])', doc: 'Returns the standard error of the b (y-intercept) coefficient.' }],
  ['linest_sem',  { signature: 'LINEST_SEM([{set}] [TOTAL [<fld>]] y_value, x_value [, y0 [, x0]])', doc: 'Returns the standard error of the m (slope) coefficient.' }],
  ['linest_sey',  { signature: 'LINEST_SEY([{set}] [TOTAL [<fld>]] y_value, x_value [, y0 [, x0]])', doc: 'Returns the standard error of the y-estimate for the regression.' }],
  ['linest_ssreg', { signature: 'LINEST_SSREG([{set}] [TOTAL [<fld>]] y_value, x_value [, y0 [, x0]])', doc: 'Returns the regression sum of squares for the linear regression.' }],
  ['linest_ssresid', { signature: 'LINEST_SSRESID([{set}] [TOTAL [<fld>]] y_value, x_value [, y0 [, x0]])', doc: 'Returns the residual sum of squares for the linear regression.' }],

  // Aggregation – string
  ['firstvalue',  { signature: 'FirstValue([{set}] [TOTAL [<fld>]] expr)', doc: 'Returns the first value loaded for the expression in the aggregated data.' }],
  ['lastvalue',   { signature: 'LastValue([{set}] [TOTAL [<fld>]] expr)', doc: 'Returns the last value loaded for the expression in the aggregated data.' }],

  // Statistical test – chi2
  ['chi2test_chi2', { signature: 'Chi2Test_chi2(col, row, actual_value [, expected_value])', doc: 'Returns the chi-squared test statistic for a contingency table.' }],
  ['chi2test_df', { signature: 'Chi2Test_df(col, row, actual_value [, expected_value])', doc: 'Returns the degrees of freedom for the chi-squared test.' }],
  ['chi2test_p',  { signature: 'Chi2Test_p(col, row, actual_value [, expected_value])', doc: 'Returns the p-value (significance) for the chi-squared test.' }],

  // Statistical test – ttest (two-sample unpaired)
  ['ttest_conf',  { signature: 'TTest_conf(grp, value [, sig])', doc: 'Returns the confidence interval half-width for a two-sample unpaired t-test.' }],
  ['ttest_df',    { signature: 'TTest_df(grp, value)', doc: 'Returns the degrees of freedom for a two-sample unpaired t-test.' }],
  ['ttest_dif',   { signature: 'TTest_dif(grp, value)', doc: 'Returns the mean difference between groups for a two-sample unpaired t-test.' }],
  ['ttest_lower', { signature: 'TTest_lower(grp, value [, sig])', doc: 'Returns the lower confidence bound for a two-sample unpaired t-test.' }],
  ['ttest_sig',   { signature: 'TTest_sig(grp, value)', doc: 'Returns the p-value for a two-sample unpaired t-test.' }],
  ['ttest_sterr', { signature: 'TTest_sterr(grp, value)', doc: 'Returns the standard error of the mean difference for a two-sample unpaired t-test.' }],
  ['ttest_t',     { signature: 'TTest_t(grp, value)', doc: 'Returns the t-statistic for a two-sample unpaired t-test.' }],
  ['ttest_upper', { signature: 'TTest_upper(grp, value [, sig])', doc: 'Returns the upper confidence bound for a two-sample unpaired t-test.' }],

  // Statistical test – ttest weighted (two-sample unpaired)
  ['ttestw_conf', { signature: 'TTestw_conf(weight, grp, value [, sig])', doc: 'Returns the confidence interval half-width for a weighted two-sample unpaired t-test.' }],
  ['ttestw_df',   { signature: 'TTestw_df(weight, grp, value)', doc: 'Returns the degrees of freedom for a weighted two-sample unpaired t-test.' }],
  ['ttestw_dif',  { signature: 'TTestw_dif(weight, grp, value)', doc: 'Returns the mean difference for a weighted two-sample unpaired t-test.' }],
  ['ttestw_lower', { signature: 'TTestw_lower(weight, grp, value [, sig])', doc: 'Returns the lower confidence bound for a weighted two-sample unpaired t-test.' }],
  ['ttestw_sig',  { signature: 'TTestw_sig(weight, grp, value)', doc: 'Returns the p-value for a weighted two-sample unpaired t-test.' }],
  ['ttestw_sterr', { signature: 'TTestw_sterr(weight, grp, value)', doc: 'Returns the standard error of the mean difference for a weighted two-sample unpaired t-test.' }],
  ['ttestw_t',    { signature: 'TTestw_t(weight, grp, value)', doc: 'Returns the t-statistic for a weighted two-sample unpaired t-test.' }],
  ['ttestw_upper', { signature: 'TTestw_upper(weight, grp, value [, sig])', doc: 'Returns the upper confidence bound for a weighted two-sample unpaired t-test.' }],

  // Statistical test – ttest1 (one-sample)
  ['ttest1_conf', { signature: 'TTest1_conf(value [, sig])', doc: 'Returns the confidence interval half-width for a one-sample t-test.' }],
  ['ttest1_df',   { signature: 'TTest1_df(value)', doc: 'Returns the degrees of freedom for a one-sample t-test.' }],
  ['ttest1_dif',  { signature: 'TTest1_dif(value [, mu])', doc: 'Returns the difference between the sample mean and the population mean mu for a one-sample t-test.' }],
  ['ttest1_lower', { signature: 'TTest1_lower(value [, sig])', doc: 'Returns the lower confidence bound for a one-sample t-test.' }],
  ['ttest1_sig',  { signature: 'TTest1_sig(value)', doc: 'Returns the p-value for a one-sample t-test.' }],
  ['ttest1_sterr', { signature: 'TTest1_sterr(value)', doc: 'Returns the standard error of the mean for a one-sample t-test.' }],
  ['ttest1_t',    { signature: 'TTest1_t(value [, mu])', doc: 'Returns the t-statistic for a one-sample t-test.' }],
  ['ttest1_upper', { signature: 'TTest1_upper(value [, sig])', doc: 'Returns the upper confidence bound for a one-sample t-test.' }],

  // Statistical test – ttest1w (one-sample weighted)
  ['ttest1w_conf', { signature: 'TTest1w_conf(weight, value [, sig])', doc: 'Returns the confidence interval half-width for a weighted one-sample t-test.' }],
  ['ttest1w_df',  { signature: 'TTest1w_df(weight, value)', doc: 'Returns the degrees of freedom for a weighted one-sample t-test.' }],
  ['ttest1w_dif', { signature: 'TTest1w_dif(weight, value [, mu])', doc: 'Returns the mean difference from mu for a weighted one-sample t-test.' }],
  ['ttest1w_lower', { signature: 'TTest1w_lower(weight, value [, sig])', doc: 'Returns the lower confidence bound for a weighted one-sample t-test.' }],
  ['ttest1w_sig', { signature: 'TTest1w_sig(weight, value)', doc: 'Returns the p-value for a weighted one-sample t-test.' }],
  ['ttest1w_sterr', { signature: 'TTest1w_sterr(weight, value)', doc: 'Returns the standard error of the mean for a weighted one-sample t-test.' }],
  ['ttest1w_t',   { signature: 'TTest1w_t(weight, value [, mu])', doc: 'Returns the t-statistic for a weighted one-sample t-test.' }],
  ['ttest1w_upper', { signature: 'TTest1w_upper(weight, value [, sig])', doc: 'Returns the upper confidence bound for a weighted one-sample t-test.' }],

  // Statistical test – ztest
  ['ztest_conf',  { signature: 'ZTest_conf(value [, sigma [, sig]])', doc: 'Returns the confidence interval half-width for a z-test.' }],
  ['ztest_dif',   { signature: 'ZTest_dif(value)', doc: 'Returns the mean deviation for a z-test.' }],
  ['ztest_lower', { signature: 'ZTest_lower(value [, sigma [, sig]])', doc: 'Returns the lower confidence bound for a z-test.' }],
  ['ztest_sig',   { signature: 'ZTest_sig(value [, sigma])', doc: 'Returns the p-value for a z-test.' }],
  ['ztest_sterr', { signature: 'ZTest_sterr(value [, sigma])', doc: 'Returns the standard error of the mean for a z-test.' }],
  ['ztest_upper', { signature: 'ZTest_upper(value [, sigma [, sig]])', doc: 'Returns the upper confidence bound for a z-test.' }],
  ['ztest_z',     { signature: 'ZTest_z(value [, sigma])', doc: 'Returns the aggregated z-test statistic.' }],

  // Statistical test – ztest weighted
  ['ztestw_conf', { signature: 'ZTestw_conf(weight, value [, sigma [, sig]])', doc: 'Returns the confidence interval half-width for a weighted z-test.' }],
  ['ztestw_dif',  { signature: 'ZTestw_dif(weight, value)', doc: 'Returns the mean deviation for a weighted z-test.' }],
  ['ztestw_lower', { signature: 'ZTestw_lower(weight, value [, sigma [, sig]])', doc: 'Returns the lower confidence bound for a weighted z-test.' }],
  ['ztestw_sig',  { signature: 'ZTestw_sig(weight, value [, sigma])', doc: 'Returns the p-value for a weighted z-test.' }],
  ['ztestw_sterr', { signature: 'ZTestw_sterr(weight, value [, sigma])', doc: 'Returns the standard error of the mean for a weighted z-test.' }],
  ['ztestw_upper', { signature: 'ZTestw_upper(weight, value [, sigma [, sig]])', doc: 'Returns the upper confidence bound for a weighted z-test.' }],
  ['ztestw_z',    { signature: 'ZTestw_z(weight, value [, sigma])', doc: 'Returns the aggregated z-test statistic for a weighted z-test.' }],

  // Statistical distribution
  ['betadensity', { signature: 'BetaDensity(value, alpha, beta)', doc: 'Returns the probability density for the Beta distribution with shape parameters alpha and beta.' }],
  ['betadist',    { signature: 'BetaDist(value, alpha, beta)', doc: 'Returns the cumulative probability for the Beta distribution.' }],
  ['betainv',     { signature: 'BetaInv(prob, alpha, beta)', doc: 'Returns the inverse of BetaDist; the value at which the cumulative Beta probability equals prob.' }],
  ['binomdist',   { signature: 'BinomDist(value, trials, trial_probability)', doc: 'Returns the cumulative binomial probability of at most value successes in trials.' }],
  ['binomfrequency', { signature: 'BinomFrequency(value, trials, trial_probability)', doc: 'Returns the probability of exactly value successes in trials (binomial PDF).' }],
  ['binominv',    { signature: 'BinomInv(trials, trial_probability, probability_s)', doc: 'Returns the smallest integer for which the cumulative binomial probability is ≥ probability_s.' }],
  ['chidensity',  { signature: 'ChiDensity(value, df)', doc: 'Returns the probability density for the chi-squared distribution with df degrees of freedom.' }],
  ['chidist',     { signature: 'ChiDist(value, df)', doc: 'Returns the one-tailed probability of the chi-squared distribution.' }],
  ['chiinv',      { signature: 'ChiInv(prob, df)', doc: 'Returns the inverse of ChiDist; the chi-squared value at which the right-tail probability equals prob.' }],
  ['fdensity',    { signature: 'FDensity(value, df1, df2)', doc: 'Returns the probability density for the F-distribution with df1 and df2 degrees of freedom.' }],
  ['fdist',       { signature: 'FDist(value, df1, df2)', doc: 'Returns the cumulative F-distribution probability.' }],
  ['finv',        { signature: 'FInv(prob, df1, df2)', doc: 'Returns the inverse of FDist; the F-value at the given right-tail probability.' }],
  ['gammadensity', { signature: 'GammaDensity(value, k, theta)', doc: 'Returns the probability density for the Gamma distribution with shape k and scale theta.' }],
  ['gammadist',   { signature: 'GammaDist(value, k, theta)', doc: 'Returns the cumulative probability for the Gamma distribution.' }],
  ['gammainv',    { signature: 'GammaInv(prob, k, theta)', doc: 'Returns the inverse of GammaDist; the value at which cumulative Gamma probability equals prob.' }],
  ['normdist',    { signature: 'NormDist(value [, mean [, standard_dev]])', doc: 'Returns the cumulative normal distribution probability (default: standard normal).' }],
  ['norminv',     { signature: 'NormInv(prob [, mean [, standard_dev]])', doc: 'Returns the inverse normal distribution; the value at which the cumulative probability equals prob.' }],
  ['poissondist', { signature: 'PoissonDist(value, mean)', doc: 'Returns the cumulative Poisson probability of at most value events with the given mean.' }],
  ['poissonfrequency', { signature: 'PoissonFrequency(value, mean)', doc: 'Returns the probability of exactly value events in a Poisson distribution with the given mean.' }],
  ['poissoninv',  { signature: 'PoissonInv(prob, mean)', doc: 'Returns the smallest integer whose cumulative Poisson probability is ≥ prob.' }],
  ['tdensity',    { signature: 'TDensity(value, df)', doc: 'Returns the probability density for Student\'s t-distribution with df degrees of freedom.' }],
  ['tdist',       { signature: 'TDist(value, df, tails)', doc: 'Returns the probability for Student\'s t-distribution; tails is 1 (one-tailed) or 2 (two-tailed).' }],
  ['tinv',        { signature: 'TInv(prob, df)', doc: 'Returns the two-tailed inverse of Student\'s t-distribution at the given probability.' }],

  // Date & time (missing entries)
  ['converttolocaltime', { signature: 'ConvertToLocalTime(timestamp [, timezone [, ignore_DST]])', doc: 'Converts a UTC timestamp to local time for the given timezone name.' }],
  ['daylightsaving', { signature: 'DaylightSaving()', doc: 'Returns the current daylight saving time adjustment as set in Windows.' }],
  ['dayname',     { signature: 'DayName(date [, period_no])', doc: 'Returns a dual value showing the date of the start of the day containing date.' }],
  ['daynumberofquarter', { signature: 'DayNumberOfQuarter(date [, start_month])', doc: 'Returns the day number within the quarter containing date.' }],
  ['daynumberofyear', { signature: 'DayNumberOfYear(date [, start_month])', doc: 'Returns the day number within the year containing date.' }],
  ['firstworkdate', { signature: 'FirstWorkDate(end_date, no_of_workdays [, holiday])', doc: 'Returns the latest start date to achieve no_of_workdays before end_date, excluding weekends and holidays.' }],
  ['gmt',         { signature: 'Gmt()', doc: 'Returns the current UTC date and time.' }],
  ['inday',       { signature: 'InDay(timestamp, base_timestamp, period_no [, day_start])', doc: 'Returns true if timestamp falls in the same day as base_timestamp adjusted by period_no days.' }],
  ['indaytotime', { signature: 'InDayToTime(timestamp, base_timestamp, period_no [, day_start])', doc: 'Returns true if timestamp falls within the portion of the day up to and including base_timestamp.' }],
  ['inlunarweek', { signature: 'InLunarWeek(date, base_date, period_no [, first_week_day])', doc: 'Returns true if date falls in the same 7-day lunar week as base_date adjusted by period_no.' }],
  ['inlunarweektodate', { signature: 'InLunarWeekToDate(date, base_date, period_no [, first_week_day])', doc: 'Returns true if date falls within the portion of a lunar week up to and including base_date.' }],
  ['inmonth',     { signature: 'InMonth(date, base_date, period_no)', doc: 'Returns true if date falls in the same calendar month as base_date adjusted by period_no months.' }],
  ['inmonthtodate', { signature: 'InMonthToDate(date, base_date, period_no)', doc: 'Returns true if date falls within the portion of a month up to and including base_date.' }],
  ['inmonths',    { signature: 'InMonths(n, date, base_date, period_no [, first_month_of_year])', doc: 'Returns true if date falls within an n-month period containing base_date.' }],
  ['inmonthstodate', { signature: 'InMonthsToDate(n, date, base_date, period_no [, first_month_of_year])', doc: 'Returns true if date falls within the portion of an n-month period up to and including base_date.' }],
  ['inquarter',   { signature: 'InQuarter(date, base_date, period_no [, first_month_of_year])', doc: 'Returns true if date falls in the same calendar quarter as base_date adjusted by period_no quarters.' }],
  ['inquartertodate', { signature: 'InQuarterToDate(date, base_date, period_no [, first_month_of_year])', doc: 'Returns true if date falls within the portion of a quarter up to and including base_date.' }],
  ['inweek',      { signature: 'InWeek(date, base_date, period_no [, first_week_day])', doc: 'Returns true if date falls in the same week as base_date adjusted by period_no weeks.' }],
  ['inweektodate', { signature: 'InWeekToDate(date, base_date, period_no [, first_week_day])', doc: 'Returns true if date falls within the portion of a week up to and including base_date.' }],
  ['inyear',      { signature: 'InYear(date, base_date, period_no [, first_month_of_year])', doc: 'Returns true if date falls in the same year as base_date adjusted by period_no years.' }],
  ['inyeartodate', { signature: 'InYearToDate(date, base_date, period_no [, first_month_of_year])', doc: 'Returns true if date falls within the portion of a year up to and including base_date.' }],
  ['lastworkdate', { signature: 'LastWorkDate(start_date, no_of_workdays [, holiday])', doc: 'Returns the earliest end date to reach no_of_workdays after start_date, excluding weekends and holidays.' }],
  ['localtime',   { signature: 'LocalTime([timezone [, ignore_DST]])', doc: 'Returns the current time for the specified timezone (defaults to the system timezone).' }],
  ['lunarweekend', { signature: 'LunarWeekEnd(date [, period_no [, first_week_day]])', doc: 'Returns the last millisecond of the lunar week containing date.' }],
  ['lunarweekname', { signature: 'LunarWeekName(date [, period_no [, first_week_day]])', doc: 'Returns a dual value showing the year and lunar week number containing date.' }],
  ['lunarweekstart', { signature: 'LunarWeekStart(date [, period_no [, first_week_day]])', doc: 'Returns the first millisecond of the lunar week containing date.' }],
  ['makeweekdate', { signature: 'MakeWeekDate(year [, week [, day_of_week]])', doc: 'Constructs a date value from a year and ISO week number (and optional day-of-week 0=Mon).' }],
  ['monthname',   { signature: 'MonthName(date [, period_no])', doc: 'Returns a dual value showing the month name and year of date.' }],
  ['monthsend',   { signature: 'MonthsEnd(n, date [, period_no [, first_month_of_year]])', doc: 'Returns the last millisecond of the n-month period containing date.' }],
  ['monthsname',  { signature: 'MonthsName(n, date [, period_no [, first_month_of_year]])', doc: 'Returns a dual value showing the month range of an n-month period.' }],
  ['monthsstart', { signature: 'MonthsStart(n, date [, period_no [, first_month_of_year]])', doc: 'Returns the first millisecond of the n-month period containing date.' }],
  ['quarterend',  { signature: 'QuarterEnd(date [, period_no [, first_month_of_year]])', doc: 'Returns the last millisecond of the quarter containing date.' }],
  ['quartername', { signature: 'QuarterName(date [, period_no [, first_month_of_year]])', doc: 'Returns a dual value showing the month range of the quarter containing date.' }],
  ['quarterstart', { signature: 'QuarterStart(date [, period_no [, first_month_of_year]])', doc: 'Returns the first millisecond of the quarter containing date.' }],
  ['setdateyear', { signature: 'SetDateYear(timestamp, year)', doc: 'Returns a timestamp with the year component replaced by year.' }],
  ['setdateyearmonth', { signature: 'SetDateYearMonth(timestamp, year, month)', doc: 'Returns a timestamp with the year and month components replaced.' }],
  ['timezone',    { signature: 'Timezone()', doc: 'Returns the name of the current timezone as set in the operating system.' }],
  ['utc',         { signature: 'UTC()', doc: 'Returns the current UTC date and time.' }],
  ['weekname',    { signature: 'WeekName(date [, period_no [, first_week_day]])', doc: 'Returns a dual value showing the year and week number containing date.' }],
  ['yearname',    { signature: 'YearName(date [, period_no [, first_month_of_year]])', doc: 'Returns a dual value showing the year of date.' }],
  ['yeartodate',  { signature: 'YearToDate(date [, yearoffset [, firstmonth [, todaydate]]])', doc: 'Returns true if date falls within the current year up to and including today.' }],

  // Exponential / logarithmic (already in FUNCTION_DOCS but not FUNCTION_NAMES — now both covered)
  ['acosh',       { signature: 'Acosh(number)', doc: 'Returns the inverse hyperbolic cosine of number.' }],
  ['asinh',       { signature: 'Asinh(number)', doc: 'Returns the inverse hyperbolic sine of number.' }],
  ['atanh',       { signature: 'Atanh(number)', doc: 'Returns the inverse hyperbolic tangent of number.' }],
  ['cosh',        { signature: 'Cosh(number)', doc: 'Returns the hyperbolic cosine of number.' }],
  ['sinh',        { signature: 'Sinh(number)', doc: 'Returns the hyperbolic sine of number.' }],
  ['tanh',        { signature: 'Tanh(number)', doc: 'Returns the hyperbolic tangent of number.' }],

  // General numeric (missing)
  ['bitcount',    { signature: 'BitCount(integer)', doc: 'Returns the number of bits set to 1 in the binary representation of integer.' }],
  ['combin',      { signature: 'Combin(n, k)', doc: 'Returns the number of combinations of k items from n (order does not matter).' }],
  ['fabs',        { signature: 'Fabs(number)', doc: 'Returns the absolute value of number as a floating-point result.' }],
  ['fact',        { signature: 'Fact(n)', doc: 'Returns the factorial of non-negative integer n (n!).' }],
  ['fmod',        { signature: 'Fmod(dividend, divisor)', doc: 'Returns the floating-point remainder of dividend / divisor.' }],
  ['permut',      { signature: 'Permut(n, k)', doc: 'Returns the number of permutations of k items from n (order matters).' }],
  ['round',       { signature: 'Round(number [, step [, offset]])', doc: 'Rounds number to the nearest multiple of step (default 1); ties round up.' }],

  // Field / inter-record (missing)
  ['lookup',      { signature: 'LookUp(field_name, match_field_name, match_field_value [, table_name])', doc: 'Returns the value of field_name in the first row where match_field_name equals match_field_value.' }],
  ['noofrows',    { signature: 'NoOfRows([table_name])', doc: 'Returns the number of rows in the named table, or in the table currently being generated.' }],
  ['previous',    { signature: 'Previous(expr)', doc: 'Returns the value of expr evaluated using data from the previous input record (null on the first record).' }],

  // File (missing)
  ['attribute',   { signature: 'Attribute(filename, attribute_name)', doc: 'Returns the value of a meta-tag attribute from a media file.' }],
  ['connectstring', { signature: 'ConnectString()', doc: 'Returns the active connection string for the current ODBC or OLE DB connection.' }],
  ['filebasename', { signature: 'FileBaseName()', doc: 'Returns the base name (without extension) of the file currently being read by the script.' }],
  ['filedir',     { signature: 'FileDir()', doc: 'Returns the directory path of the file currently being read by the script.' }],
  ['fileextension', { signature: 'FileExtension()', doc: 'Returns the file extension of the file currently being read by the script.' }],
  ['filename',    { signature: 'FileName()', doc: 'Returns the name (with extension, without path) of the file currently being read by the script.' }],
  ['filepath',    { signature: 'FilePath()', doc: 'Returns the full path of the file currently being read by the script.' }],
  ['filesize',    { signature: 'FileSize([filename])', doc: 'Returns the size in bytes of the named file, or of the file being read if omitted.' }],
  ['filetime',    { signature: 'FileTime([filename])', doc: 'Returns the last-modified timestamp of the named file, or of the file being read if omitted.' }],
  ['getfolderpath', { signature: 'GetFolderPath(foldername)', doc: 'Returns the Windows path for a named special folder (e.g., "MyDocuments").' }],
  ['getregistrystring', { signature: 'GetRegistryString(path, key)', doc: 'Returns the value of a Windows registry string entry at the given path and key name.' }],
  ['rem',           { signature: 'Rem(dividend, divisor)', doc: 'Returns the remainder of dividend / divisor; the sign of the result matches the dividend (unlike Mod which is always non-negative).' }],
  ['qvdcreatetime', { signature: 'QvdCreateTime(filename)', doc: 'Returns the XML-header creation timestamp from a QVD file.' }],
  ['qvdfieldname', { signature: 'QvdFieldName(filename, field_no)', doc: 'Returns the name of the field at the given 1-based position in a QVD file.' }],
  ['qvdnooffields', { signature: 'QvdNoOfFields(filename)', doc: 'Returns the number of fields in a QVD file.' }],
  ['qvdnoofrecords', { signature: 'QvdNoOfRecords(filename)', doc: 'Returns the number of records stored in a QVD file.' }],
  ['qvdtablename', { signature: 'QvdTableName(filename)', doc: 'Returns the name of the table stored in a QVD file.' }],

  // Financial (non-aggregation)
  ['blackandschole', { signature: 'BlackAndSchole(strike, time_left, underlying_price, vol, rfr, call_put_flag)', doc: 'Returns the theoretical option price using the Black & Scholes model; call_put_flag is "c" or "p".' }],
  ['fv',          { signature: 'Fv(rate, nper, pmt [, pv [, type]])', doc: 'Returns the future value of an investment based on periodic constant payments and a constant interest rate.' }],
  ['nper',        { signature: 'Nper(rate, pmt, pv [, fv [, type]])', doc: 'Returns the number of periods for an investment given periodic payments and a constant interest rate.' }],
  ['pmt',         { signature: 'Pmt(rate, nper, pv [, fv [, type]])', doc: 'Returns the periodic payment for a loan based on constant payments and a constant interest rate.' }],
  ['pv',          { signature: 'Pv(rate, nper, pmt [, fv [, type]])', doc: 'Returns the present value of an investment — the total amount a series of future payments is worth today.' }],
  ['rate',        { signature: 'Rate(nper, pmt, pv [, fv [, type [, guess]]])', doc: 'Returns the interest rate per period for an annuity.' }],

  // Geospatial
  ['geoaggrgeometry', { signature: 'GeoAggrGeometry(geometry)', doc: 'Aggregates multiple geometry objects into a single combined geometry (script aggregation).' }],
  ['geoboundingbox', { signature: 'GeoBoundingBox(geometry)', doc: 'Returns the bounding box of a geometry as a GeoJSON polygon string.' }],
  ['geocountvertex', { signature: 'GeoCountVertex(geometry)', doc: 'Returns the number of vertices in the given geometry.' }],
  ['geogetboundingbox', { signature: 'GeoGetBoundingBox(geometry)', doc: 'Returns the bounding box coordinates [minLon, minLat, maxLon, maxLat] of a geometry.' }],
  ['geogetpolygoncenter', { signature: 'GeoGetPolygonCenter(geometry)', doc: 'Returns the center point of a polygon geometry as a GeoJSON point.' }],
  ['geoinvprojectgeometry', { signature: 'GeoInvProjectGeometry(projection, geometry)', doc: 'Re-projects geometry from the named projection back to WGS84 longitude/latitude.' }],
  ['geomakepoint', { signature: 'GeoMakePoint(latitude, longitude)', doc: 'Creates a GeoJSON point string from latitude and longitude values.' }],
  ['geoproject',  { signature: 'GeoProject(projection, geometry)', doc: 'Projects geometry to the named map projection and returns the result.' }],
  ['geoprojectgeometry', { signature: 'GeoProjectGeometry(projection, geometry)', doc: 'Projects geometry from WGS84 coordinates into the named map projection.' }],
  ['georeducegeometry', { signature: 'GeoReduceGeometry(geometry [, fraction])', doc: 'Reduces the vertex count of geometry by the given fraction (0–1) while preserving shape.' }],

  // JSON (missing)
  ['isjson',      { signature: 'IsJson(string [, json_type])', doc: 'Returns true if string is valid JSON, optionally constrained to a specific JSON type ("string", "array", etc.).' }],
  ['jsonget',     { signature: 'JsonGet(json, path)', doc: 'Returns the value at the JSON path (e.g. "$.key") from a JSON string.' }],
  ['jsonset',     { signature: 'JsonSet(json, path, value)', doc: 'Returns the JSON string with the value at the given path replaced by value.' }],
  ['jsonsetex',   { signature: 'JsonSetEx(json, path1, value1 [, path2, value2, ...])', doc: 'Returns the JSON string with multiple path/value pairs updated in one call.' }],

  // Null (missing)
  ['emptyisnull', { signature: 'EmptyIsNull(expr)', doc: 'Returns null if expr evaluates to an empty string; otherwise returns expr unchanged.' }],

  // String (missing entries)
  ['countregex',  { signature: 'CountRegex(text, regex)', doc: 'Returns the number of non-overlapping occurrences of the regular expression in text.' }],
  ['evaluate',    { signature: 'Evaluate(expression_string)', doc: 'Evaluates expression_string as a Qlik expression at runtime and returns the result.' }],
  ['extractregex', { signature: 'ExtractRegex(text, regex)', doc: 'Returns the first substring of text matched by the entire regex pattern.' }],
  ['extractregexgroup', { signature: 'ExtractRegexGroup(text, regex, group_no)', doc: 'Returns the substring matched by capture group group_no in the first regex match.' }],
  ['findoneofvalue', { signature: 'FindOneOfValue(str, char_set [, count])', doc: 'Returns the position of the count-th character in str found in char_set (default: first occurrence).' }],
  ['indexregex',  { signature: 'IndexRegex(text, regex [, count])', doc: 'Returns the character position of the count-th regex match in text (default: first).' }],
  ['indexregexgroup', { signature: 'IndexRegexGroup(text, regex, group_no [, count])', doc: 'Returns the character position of capture group group_no in the count-th regex match.' }],
  ['isempty',     { signature: 'IsEmpty(expr)', doc: 'Returns true if expr evaluates to an empty string.' }],
  ['isregex',     { signature: 'IsRegex(text, regex)', doc: 'Returns true if text matches the regular expression pattern.' }],
  ['levenshteindist', { signature: 'LevenshteinDist(str1, str2)', doc: 'Returns the Levenshtein edit distance between str1 and str2 (minimum single-character edits to transform one into the other).' }],
  ['matchregex',  { signature: 'MatchRegex(text, regex)', doc: 'Returns the character position of the first regex match in text, or 0 if no match.' }],
  ['msgbox',      { signature: 'MsgBox(message [, caption [, buttons [, icon]]])', doc: 'Displays a dialog box during script execution and returns the index of the button pressed.' }],
  ['replaceregex', { signature: 'ReplaceRegex(text, regex, replacement [, count])', doc: 'Replaces up to count regex matches in text with replacement (default: all).' }],
  ['replaceregexgroup', { signature: 'ReplaceRegexGroup(text, regex, group_no, replacement [, count])', doc: 'Replaces capture group group_no in regex matches with replacement.' }],
  ['subfieldregex', { signature: 'SubFieldRegex(text, regex [, field_no])', doc: 'Splits text by the regex delimiter and returns the field_no field; generates rows for each part if field_no is omitted.' }],

  // System
  ['author',      { signature: 'Author()', doc: 'Returns the name of the author of the current document as set in document properties.' }],
  ['calcdim',     { signature: 'CalcDim()', doc: 'Returns true when the current dimension is a calculated (synthetic) dimension.' }],
  ['clientplatform', { signature: 'ClientPlatform()', doc: 'Returns a string describing the client platform, including browser and operating system.' }],
  ['computername', { signature: 'ComputerName()', doc: 'Returns the network name of the computer running the Qlik Sense server.' }],
  ['documentname', { signature: 'DocumentName()', doc: 'Returns the name of the current document without path or extension.' }],
  ['documentpath', { signature: 'DocumentPath()', doc: 'Returns the full file path of the current document.' }],
  ['documenttitle', { signature: 'DocumentTitle()', doc: 'Returns the title of the current document.' }],
  ['engineversion', { signature: 'EngineVersion()', doc: 'Returns the version string of the Qlik associative engine.' }],
  ['getcollationlocale', { signature: 'GetCollationLocale()', doc: 'Returns the locale string currently used for sort-order collation.' }],
  ['getobjectfield', { signature: 'GetObjectField([chart_object_id [, field_no]])', doc: 'Returns the field name used by the specified dimension of a chart object.' }],
  ['getsysattr',  { signature: 'GetSysAttr(attribute_name)', doc: 'Returns the value of a named system attribute.' }],
  ['getuserattr', { signature: 'GetUserAttr(attribute_name)', doc: 'Returns the value of a named user attribute defined in the access section.' }],
  ['groupdimensionindex', { signature: 'GroupDimensionIndex()', doc: 'Returns the zero-based index of the currently active dimension within a cyclic or drill-down group.' }],
  ['groupdimensionlabel', { signature: 'GroupDimensionLabel([group_name])', doc: 'Returns the label of the currently displayed dimension within a cyclic or drill-down group.' }],
  ['inobject',    { signature: 'InObject()', doc: 'Returns true when the expression is evaluated in a chart context rather than the script.' }],
  ['ispartialreload', { signature: 'IsPartialReload()', doc: 'Returns true if the current reload is a partial (incremental) reload.' }],
  ['objectid',    { signature: 'ObjectId()', doc: 'Returns the object ID of the current chart object.' }],
  ['osuser',      { signature: 'OSUser()', doc: 'Returns the operating system username of the current user.' }],
  ['productversion', { signature: 'ProductVersion()', doc: 'Returns the full version string of Qlik Sense.' }],
  ['reloadtime',  { signature: 'ReloadTime()', doc: 'Returns the timestamp of when the last full reload completed.' }],
  ['statename',   { signature: 'StateName()', doc: 'Returns the name of the alternate state that the current chart object belongs to.' }],

  // Table
  ['fieldname',   { signature: 'FieldName(field_no [, table_name])', doc: 'Returns the name of the field at 1-based position field_no in the named table (or the current result table).' }],
  ['fieldnumber', { signature: 'FieldNumber(field_name [, table_name])', doc: 'Returns the 1-based position of field_name in the named table, or 0 if not found.' }],
  ['nooffields',  { signature: 'NoOfFields([table_name])', doc: 'Returns the number of fields in the named table.' }],

  // Relational / ML
  ['rank',        { signature: 'Rank([TOTAL] expr [, mode [, fmt]])', doc: 'Returns the rank of the current row value among all values in the segment; ties handled by mode.' }],
  ['hrank',       { signature: 'HRank([TOTAL] expr [, mode [, fmt]])', doc: 'Returns the rank of the current column value in a pivot table row; ties handled by mode.' }],
  ['kmeans2d',    { signature: 'KMeans2D(k, x_coord, y_coord [, iterations [, weight]])', doc: 'Clusters data into k groups using 2D coordinates; returns the cluster ID (0-based) for each data point.' }],
  ['kmeansnd',    { signature: 'KMeansND(k, coordinate_expr [, iterations [, weight]])', doc: 'Clusters data into k groups using n-dimensional coordinates; returns the cluster ID for each data point.' }],
  ['kmeanscentroid2d', { signature: 'KMeansCentroid2D(k, cluster_id, coord_no, x_coord, y_coord [, iterations [, weight]])', doc: 'Returns the centroid coordinate (coord_no: 0=x, 1=y) for the specified cluster ID.' }],
  ['kmeanscentroidnd', { signature: 'KMeansCentroidND(k, cluster_id, coord_no, coordinate_expr [, iterations [, weight]])', doc: 'Returns the n-dimensional centroid coordinate at coord_no for the specified cluster ID.' }],
  ['stl_trend',   { signature: 'STL_Trend([TOTAL] y_value [, period_length [, trend_smooth [, seasonal_smooth]]])', doc: 'Returns the trend component from STL (Seasonal-Trend decomposition using Loess) of a time series.' }],
  ['stl_seasonal', { signature: 'STL_Seasonal([TOTAL] y_value [, period_length [, trend_smooth [, seasonal_smooth]]])', doc: 'Returns the seasonal component from STL decomposition of a time series.' }],
  ['stl_residual', { signature: 'STL_Residual([TOTAL] y_value [, period_length [, trend_smooth [, seasonal_smooth]]])', doc: 'Returns the residual (irregular) component from STL decomposition of a time series.' }],

  // Window
  ['window',      { signature: 'Window(main_expr [, partition_exprs] [, sort_type [, sort_expr [, filter_expr [, start_expr [, end_expr]]]]])', doc: 'Performs a calculation over a sliding window of rows within a partition, returning one value per row.' }],
  ['wrank',       { signature: 'WRank(expr [, mode [, fmt]])', doc: 'Returns the rank of the current row value within its window partition.' }],

  // Range
  ['rangeapp',    { signature: 'RangeApp(first_value, step, last_value)', doc: 'Returns a list of values from first_value to last_value incremented by step.' }],
  ['rangeavg',    { signature: 'RangeAvg(first_expr [, ...])', doc: 'Returns the average of a range of values; null values are ignored.' }],
  ['rangecorrel', { signature: 'RangeCorrel(x_value, y_value [, x_value2, y_value2, ...])', doc: 'Returns the Pearson correlation coefficient for paired x and y values.' }],
  ['rangecount',  { signature: 'RangeCount(first_expr [, ...])', doc: 'Returns the count of all values (numeric and text) in the range.' }],
  ['rangefractile', { signature: 'RangeFractile(fraction, first_expr [, ...])', doc: 'Returns the value at the given fractile (0–1) in the range.' }],
  ['rangeirr',    { signature: 'RangeIrr(value [, value2, ...])', doc: 'Returns the internal rate of return for a range of cash flow values.' }],
  ['rangekurtosis', { signature: 'RangeKurtosis(first_expr [, ...])', doc: 'Returns the excess kurtosis of a range of values.' }],
  ['rangemax',    { signature: 'RangeMax(first_expr [, ...])', doc: 'Returns the highest numeric value in the range; null ignored.' }],
  ['rangemaxstring', { signature: 'RangeMaxString(first_expr [, ...])', doc: 'Returns the last value in the range when sorted alphabetically.' }],
  ['rangemin',    { signature: 'RangeMin(first_expr [, ...])', doc: 'Returns the lowest numeric value in the range; null ignored.' }],
  ['rangeminstring', { signature: 'RangeMinString(first_expr [, ...])', doc: 'Returns the first value in the range when sorted alphabetically.' }],
  ['rangemissingcount', { signature: 'RangeMissingCount(first_expr [, ...])', doc: 'Returns the count of non-numeric, non-null (missing) values in the range.' }],
  ['rangemode',   { signature: 'RangeMode(first_expr [, ...])', doc: 'Returns the most frequently occurring value in the range.' }],
  ['rangenullcount', { signature: 'RangeNullCount(first_expr [, ...])', doc: 'Returns the count of null values in the range.' }],
  ['rangenumericcount', { signature: 'RangeNumericCount(first_expr [, ...])', doc: 'Returns the count of numeric values in the range.' }],
  ['rangenpv',    { signature: 'RangeNpv(discount_rate, value [, value2, ...])', doc: 'Returns the net present value for a range of irregular cash flow values.' }],
  ['rangeonly',   { signature: 'RangeOnly(first_expr [, ...])', doc: 'Returns a value if all values in the range are identical; otherwise null.' }],
  ['rangeskew',   { signature: 'RangeSkew(first_expr [, ...])', doc: 'Returns the skewness of a range of values.' }],
  ['rangestdev',  { signature: 'RangeStdev(first_expr [, ...])', doc: 'Returns the sample standard deviation of a range of values.' }],
  ['rangesum',    { signature: 'RangeSum(first_expr [, ...])', doc: 'Returns the sum of a range of values; null treated as 0.' }],
  ['rangetextcount', { signature: 'RangeTextCount(first_expr [, ...])', doc: 'Returns the count of text (non-numeric, non-null) values in the range.' }],
  ['rangexirr',   { signature: 'RangeXirr(value, date [, value2, date2, ...])', doc: 'Returns the internal rate of return for a range of non-periodic cash flows.' }],
  ['rangexnpv',   { signature: 'RangeXnpv(discount_rate, value, date [, value2, date2, ...])', doc: 'Returns the net present value for a range of non-periodic cash flows.' }],
]);

// ── 1. Completion provider ─────────────────────────────────────────────────

export class QlikCompletionProvider implements vscode.CompletionItemProvider {
  constructor(private readonly treeProvider: QlikScriptTreeProvider) {}

  provideCompletionItems(document: vscode.TextDocument): vscode.CompletionItem[] {
    // Include the current document text plus any loaded cloud sections (deduplication via Set in parseScriptSymbols)
    const texts = [document.getText(), ...this.treeProvider.getSections().map(s => s.body)];
    const { variables, subs, tables } = parseScriptSymbols(texts.join('\n'));

    const dynamicItems: vscode.CompletionItem[] = [
      ...variables.map(name => {
        const item = new vscode.CompletionItem(name, vscode.CompletionItemKind.Variable);
        item.detail = 'Script variable';
        item.sortText = `0_${name.toLowerCase()}`;
        return item;
      }),
      ...subs.map(name => {
        const item = new vscode.CompletionItem(name, vscode.CompletionItemKind.Function);
        item.insertText = new vscode.SnippetString(`${name}($1)`);
        item.detail = 'Script subroutine';
        item.sortText = `0_${name.toLowerCase()}`;
        return item;
      }),
      ...tables.map(name => {
        const item = new vscode.CompletionItem(name, vscode.CompletionItemKind.Class);
        item.detail = 'Table label';
        item.sortText = `1_${name.toLowerCase()}`;
        return item;
      }),
    ];

    return [...dynamicItems, ...FUNCTIONS, ...KEYWORDS, ...CONSTANTS];
  }
}

// ── 2. Hover provider ─────────────────────────────────────────────────────

export class QlikHoverProvider implements vscode.HoverProvider {
  provideHover(document: vscode.TextDocument, position: vscode.Position): vscode.Hover | undefined {
    const range = document.getWordRangeAtPosition(position, /[\w$#]+/);
    if (!range) return undefined;

    const word = document.getText(range).toLowerCase();
    const entry = FUNCTION_DOCS.get(word);

    if (entry) {
      const md = new vscode.MarkdownString();
      md.appendCodeblock(entry.signature, 'qlikscript');
      md.appendMarkdown(entry.doc);
      return new vscode.Hover(md, range);
    }

    if (FUNCTION_NAMES_LOWER.has(word)) {
      return new vscode.Hover(new vscode.MarkdownString('Built-in Qlik function'), range);
    }

    return undefined;
  }
}

// ── 3. Diagnostics ────────────────────────────────────────────────────────

let _validateTimer: ReturnType<typeof setTimeout> | undefined;

export function validateSections(
  sections: ScriptSection[],
  scriptFS: QlikScriptFS,
  appId: string,
  diagCollection: vscode.DiagnosticCollection,
  debounceMs = 0,
): void {
  if (_validateTimer) clearTimeout(_validateTimer);
  _validateTimer = setTimeout(() => _runValidation(sections, scriptFS, appId, diagCollection), debounceMs);
}

function _runValidation(
  sections: ScriptSection[],
  scriptFS: QlikScriptFS,
  appId: string,
  diagCollection: vscode.DiagnosticCollection,
): void {
  const allText = sections.map(s => s.body).join('\n');
  const { subs } = parseScriptSymbols(allText);
  const definedSubs = new Set(subs.map(s => s.toLowerCase()));

  diagCollection.clear();

  for (const section of sections) {
    const diags: vscode.Diagnostic[] = [];
    CALL_RE.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = CALL_RE.exec(section.body)) !== null) {
      const callName = m[1];
      if (definedSubs.has(callName.toLowerCase())) continue;

      // Convert character offset to (line, char) within this section
      const beforeMatch = section.body.slice(0, m.index + m[0].length - callName.length);
      const lines = beforeMatch.split('\n');
      const line = lines.length - 1;
      const char = lines[line].length;
      const range = new vscode.Range(line, char, line, char + callName.length);

      diags.push(new vscode.Diagnostic(
        range,
        `Subroutine '${callName}' is not defined in any section`,
        vscode.DiagnosticSeverity.Warning,
      ));
    }

    const uri = QlikScriptFS.uri(appId, section);
    diagCollection.set(uri, diags);
  }
}

// ── Diagnostics for a single local document ───────────────────────────────

let _localValidateTimer: ReturnType<typeof setTimeout> | undefined;

export function validateDocument(
  document: vscode.TextDocument,
  diagCollection: vscode.DiagnosticCollection,
  debounceMs = 0,
): void {
  if (_localValidateTimer) clearTimeout(_localValidateTimer);
  _localValidateTimer = setTimeout(() => _runDocValidation(document, diagCollection), debounceMs);
}

function _runDocValidation(
  document: vscode.TextDocument,
  diagCollection: vscode.DiagnosticCollection,
): void {
  const text = document.getText();
  const { subs } = parseScriptSymbols(text);
  const definedSubs = new Set(subs.map(s => s.toLowerCase()));

  const diags: vscode.Diagnostic[] = [];
  CALL_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = CALL_RE.exec(text)) !== null) {
    const callName = m[1];
    if (definedSubs.has(callName.toLowerCase())) continue;
    const start = document.positionAt(m.index + m[0].length - callName.length);
    const range = new vscode.Range(start, start.translate(0, callName.length));
    diags.push(new vscode.Diagnostic(
      range,
      `Subroutine '${callName}' is not defined`,
      vscode.DiagnosticSeverity.Warning,
    ));
  }
  diagCollection.set(document.uri, diags);
}
