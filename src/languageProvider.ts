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
  'Abs', 'Acos', 'Acosh', 'AddMonths', 'AddYears', 'Age', 'ApplyCodepage', 'ApplyMap',
  'Asin', 'Asinh', 'Atan', 'Atan2', 'Atanh', 'AutoNumber', 'Avg',
  'BitCount', 'Capitalize', 'Ceil', 'Chr', 'Class', 'Coalesce', 'Combin', 'Concat',
  'ConvertToLocalTime', 'Correl', 'Cos', 'Cosh', 'Count', 'CountRegex',
  'Date', 'Date#', 'Day', 'DaylightSaving', 'DayEnd', 'DayName',
  'DayNumberOfQuarter', 'DayNumberOfYear', 'DayStart', 'Div', 'E', 'Evaluate', 'Even',
  'Exists', 'ExtractRegex', 'ExtractRegexGroup',
  'Fabs', 'Fact', 'FieldIndex', 'FieldValue', 'FieldValueCount',
  'FindOneOfValue', 'FirstSortedValue', 'FirstValue', 'FirstWorkDate',
  'Floor', 'Fmod', 'Frac', 'Fractile', 'Fv',
  'GetRegistryString', 'Gmt', 'Hash128', 'Hash160', 'Hash256',
  'Hour', 'If', 'InDay', 'InDayToTime', 'Index', 'IndexRegex', 'IndexRegexGroup',
  'Interval', 'Interval#', 'InLunarWeek', 'InLunarWeekToDate',
  'InMonth', 'InMonths', 'InMonthsToDate', 'InMonthToDate',
  'InQuarter', 'InQuarterToDate', 'InYear', 'InYearToDate',
  'Irr', 'IsEmpty', 'IsJson', 'IsNull', 'IsNum', 'IsRegex', 'IsText', 'IsValidColor',
  'JsonGet', 'JsonSet', 'JsonSetEx',
  'KeepChar', 'Kurtosis', 'LastValue', 'LastWorkDate', 'Left', 'Len',
  'LevenshteinDist', 'LocalTime', 'Lower', 'LTrim',
  'LunarWeekEnd', 'LunarWeekName', 'LunarWeekStart',
  'MakeDate', 'MakeWeekDate', 'MakeTime', 'MapSubString', 'Match', 'Max', 'MaxString',
  'Median', 'Mid', 'Min', 'MinString', 'Minute', 'MixMatch', 'Mod', 'Mode',
  'Money', 'Money#', 'Month', 'MonthEnd', 'MonthName',
  'MonthsEnd', 'MonthsName', 'MonthsStart', 'MonthStart',
  'MsgBox', 'NetworkDays', 'Now', 'Nper', 'Npv', 'Null', 'NullCount',
  'Num', 'Num#', 'NumericCount', 'Odd', 'Only', 'Ord',
  'Peek', 'Permut', 'Pi', 'Pmt', 'PurgeChar', 'Pv',
  'Rand', 'RangeApp', 'RangeAvg', 'RangeCount', 'RangeCorrel', 'RangeFractile',
  'RangeIrr', 'RangeKurtosis', 'RangeMax', 'RangeMaxString', 'RangeMin',
  'RangeMinString', 'RangeMissingCount', 'RangeMode', 'RangeNullCount',
  'RangeNumericCount', 'RangeNpv', 'RangeOnly', 'RangeSkew', 'RangeStdev',
  'RangeSum', 'RangeTextCount', 'RangeXirr', 'RangeXnpv', 'Rate', 'RecNo',
  'Rem', 'Repeat', 'Replace', 'ReplaceRegex', 'ReplaceRegexGroup', 'Right',
  'RowNo', 'RTrim', 'Second', 'SetDateYear', 'SetDateYearMonth',
  'Sign', 'Sin', 'Sinh', 'Skew', 'Sqrt', 'Stdev',
  'SubField', 'SubFieldRegex', 'SubstringCount', 'Sum',
  'Tan', 'Tanh', 'TextBetween', 'TextCount',
  'Time', 'Time#', 'Timestamp', 'Timestamp#', 'Timezone', 'Today', 'Trim',
  'Upper', 'Utc', 'Variance',
  'Week', 'WeekDay', 'WeekEnd', 'WeekName', 'WeekStart', 'WeekYear',
  'WildMatch', 'WildMatch5', 'Xirr', 'Xnpv',
  'Year', 'YearEnd', 'YearName', 'YearStart', 'YearToDate',
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
