/**
 * Google Apps Script for processing D2SDCE applications from Gmail into a spreadsheet.
 *
 * Usage order: execute setupApplicantSheet() once to prepare the sheet, then run() regularly.
 * This script is designed to be idempotent and relies on Gmail message IDs to avoid duplicates.
 */

const TIMEZONE = 'Asia/Tokyo';
const HEADER_ROW = [
  '応募ID',
  '氏名',
  '年齢',
  '性別',
  '電話番号',
  '同居家族（続柄・年齢を自由記述）',
  '業種',
  '職種',
  '学生フラグ',
  '学年（学生のみ）',
  '学部学科専攻（学生のみ）',
  'ミントタブレット購入頻度',
  'ソフトキャンディ購入頻度',
  '2025-10-01 15:00 参加可',
  '2025-10-01 19:00 参加可',
  '2025-10-02 15:00 参加可',
  '2025-10-02 19:00 参加可',
  '備考',
  'Date(ISO)',
  'From Name',
  'From Email',
  'To',
  'Cc',
  'Subject',
  'Summary(簡易)',
  'Body(plain,先頭1200字)',
  'Thread URL',
  'Message ID',
  'Thread ID'
];

const PHONE_REGEX = '^((\\+?\\d{1,4}[- ]?)?0\\d{1,4}[- ]?\\d{1,4}[- ]?\\d{3,4})$';
const DATA_VALIDATION_ROWS = 100;

/**
 * Fetch configuration from script properties with sensible defaults.
 * @return {{spreadsheetId: string, sheetName: string, query: string, processedLabel: string, summaryMaxLen: number, bodyHeadLen: number}}
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID property is required.');
  }
  const sheetName = props.getProperty('SHEET_NAME') || '応募一覧';
  const query = props.getProperty('QUERY') || 'in:inbox subject:#D2SDCE';
  const processedLabel = props.getProperty('PROCESSED_LABEL') || 'processed/D2SDCE';
  const summaryMaxLen = parseInt(props.getProperty('SUMMARY_MAX_LEN'), 10);
  const bodyHeadLen = parseInt(props.getProperty('BODY_HEAD_LEN'), 10);
  return {
    spreadsheetId: spreadsheetId,
    sheetName: sheetName,
    query: query,
    processedLabel: processedLabel,
    summaryMaxLen: isNaN(summaryMaxLen) ? 600 : summaryMaxLen,
    bodyHeadLen: isNaN(bodyHeadLen) ? 1200 : bodyHeadLen
  };
}

/**
 * Initialize the applicant sheet with headers, data validation, and basic formatting.
 * Idempotent and safe to run multiple times.
 */
function setupApplicantSheet() {
  const config = getConfig();
  const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = getOrCreateSheet_(spreadsheet, config.sheetName);

  sheet.clearContents();
  sheet.clearFormats();
  sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
  sheet.setFrozenRows(1);

  // Apply data validation and formatting to the first 100 rows beyond the header.
  applyDataValidations_(sheet);

  // Improve column widths for readability.
  sheet.setColumnWidths(1, 3, 80);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidths(6, 3, 160);
  sheet.setColumnWidths(9, 3, 110);
  sheet.setColumnWidth(12, 160);
  sheet.setColumnWidth(13, 160);
  sheet.setColumnWidths(14, 4, 130);
  sheet.setColumnWidth(18, 200);
  sheet.setColumnWidths(19, 11, 180);
  sheet.getRange('A1:AC1').setFontWeight('bold').setBackground('#f5f5f5');
}

/**
 * Main entry point to ingest Gmail messages and append to the sheet.
 * Searches for threads matching the query, filters out already processed messages by ID,
 * extracts structured data, appends rows, and labels threads as processed.
 */
function run() {
  const config = getConfig();
  const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = getOrCreateSheet_(spreadsheet, config.sheetName);
  ensureHeader_(sheet);

  const processedIds = loadExistingMessageIds_(sheet);
  const threads = GmailApp.search(config.query, 0, 500);
  const label = getOrCreateLabel_(config.processedLabel);

  const rowsToAppend = [];
  let skippedCount = 0;
  let processedCount = 0;

  threads.forEach(function(thread) {
    const messages = thread.getMessages();
    messages.forEach(function(message) {
      const messageId = message.getId();
      if (processedIds.has(messageId)) {
        skippedCount += 1;
        return;
      }
      const parsed = mapMessageToRow_(message, thread, config);
      rowsToAppend.push(parsed);
      processedIds.add(messageId);
      processedCount += 1;
    });
    thread.addLabel(label);
  });

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, HEADER_ROW.length)
      .setValues(rowsToAppend);
  }

  Logger.log('Processed %d new messages, skipped %d existing ones.', processedCount, skippedCount);
}

/**
 * Ensure the sheet exists, otherwise create it.
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet
 * @param {string} sheetName
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 * @private
 */
function getOrCreateSheet_(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * Guarantee the header row is present.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @private
 */
function ensureHeader_(sheet) {
  const currentHeader = sheet.getRange(1, 1, 1, HEADER_ROW.length).getValues()[0];
  const isHeaderDifferent = HEADER_ROW.some(function(value, index) {
    return currentHeader[index] !== value;
  });
  if (isHeaderDifferent) {
    sheet.insertRows(1);
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
  }
  sheet.setFrozenRows(1);
}

/**
 * Read existing message IDs from the sheet to avoid duplication.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Set<string>}
 * @private
 */
function loadExistingMessageIds_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return new Set();
  }
  const idRange = sheet.getRange(2, 28, lastRow - 1, 1).getValues();
  const ids = new Set();
  idRange.forEach(function(row) {
    const id = row[0];
    if (id) {
      ids.add(String(id));
    }
  });
  return ids;
}

/**
 * Build a single row array from a Gmail message.
 * @param {GoogleAppsScript.Gmail.GmailMessage} message
 * @param {GoogleAppsScript.Gmail.GmailThread} thread
 * @param {{summaryMaxLen: number, bodyHeadLen: number}} config
 * @return {Array<string>}
 * @private
 */
function mapMessageToRow_(message, thread, config) {
  const htmlBody = message.getBody();
  const plainBody = htmlBody ? cleanPlainText_(htmlToPlainText_(htmlBody)) : cleanPlainText_(message.getPlainBody());
  const summary = generateSummary_(plainBody, config.summaryMaxLen);
  const bodyHead = truncateText_(plainBody, config.bodyHeadLen);

  const from = parseNameAndEmail_(message.getFrom());
  const dateIso = Utilities.formatDate(message.getDate(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX");
  const row = new Array(HEADER_ROW.length).fill('');

  row[19 - 1] = dateIso; // Date(ISO)
  row[20 - 1] = from.name;
  row[21 - 1] = from.email;
  row[22 - 1] = message.getTo();
  row[23 - 1] = message.getCc();
  row[24 - 1] = message.getSubject();
  row[25 - 1] = summary;
  row[26 - 1] = bodyHead;
  row[27 - 1] = 'https://mail.google.com/mail/u/0/#inbox/' + thread.getId();
  row[28 - 1] = message.getId();
  row[29 - 1] = thread.getId();
  return row;
}

/**
 * Convert HTML to plain text by stripping tags, removing scripts/styles, and unescaping entities.
 * @param {string} html
 * @return {string}
 * @private
 */
function htmlToPlainText_(html) {
  if (!html) {
    return '';
  }
  let text = html;
  text = text.replace(/<\s*(script|style)[^>]*>[\s\S]*?<\/\s*\1\s*>/gi, '');
  text = text.replace(/<(br|hr)\s*\/?>(?=\s*<)/gi, '\n');
  text = text.replace(/<(br|hr)\s*\/?>(?!\n)/gi, '\n');
  text = text.replace(/<\/(p|div|tr|li|h[1-6]|table)>/gi, '\n');
  text = text.replace(/<(li)>/gi, '\n- ');
  text = text.replace(/<(p|div|tr|td|th|h[1-6])[^>]*>/gi, '\n');
  text = text.replace(/<\/?(span|strong|b|em|u)[^>]*>/gi, '');
  text = text.replace(/<[^>]+>/g, ' ');
  text = decodeHtmlEntities_(text);
  text = text.replace(/[\t ]+/g, ' ');
  text = text.replace(/\s*\n\s*/g, '\n');
  return text;
}

/**
 * Basic HTML entity decoding supporting common named and numeric entities.
 * @param {string} text
 * @return {string}
 * @private
 */
function decodeHtmlEntities_(text) {
  const namedEntities = {
    amp: '&',
    lt: '<',
    gt: '>',
    quot: '"',
    apos: '\'',
    nbsp: ' '
  };
  return text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, function(match, entity) {
    if (entity[0] === '#') {
      const code = entity[1] === 'x' || entity[1] === 'X'
        ? parseInt(entity.slice(2), 16)
        : parseInt(entity.slice(1), 10);
      if (!isNaN(code)) {
        return String.fromCharCode(code);
      }
      return match;
    }
    const lower = entity.toLowerCase();
    if (namedEntities.hasOwnProperty(lower)) {
      return namedEntities[lower];
    }
    return match;
  });
}

/**
 * Clean plain text by removing quoted replies, signatures, and compressing whitespace.
 * @param {string} text
 * @return {string}
 * @private
 */
function cleanPlainText_(text) {
  if (!text) {
    return '';
  }
  const lines = text.split(/\r?\n/);
  const result = [];
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trimEnd();
    if (!line) {
      if (result.length === 0 || result[result.length - 1] !== '') {
        result.push('');
      }
      continue;
    }
    const trimmed = line.trim();
    if (/^>/.test(trimmed)) {
      continue; // quoted reply
    }
    if (/^On\s.+wrote:\s*$/i.test(trimmed)) {
      break;
    }
    if (/^--\s*$/.test(trimmed)) {
      break;
    }
    if (/^From:\s.+$/i.test(trimmed)) {
      continue;
    }
    result.push(trimmed);
  }
  while (result.length > 0 && result[result.length - 1] === '') {
    result.pop();
  }
  return result.join('\n');
}

/**
 * Create a concise summary emphasizing bullet points and headings when available.
 * @param {string} text
 * @param {number} maxLen
 * @return {string}
 * @private
 */
function generateSummary_(text, maxLen) {
  if (!text) {
    return '';
  }
  const lines = text.split(/\r?\n/).map(function(line) { return line.trim(); }).filter(function(line) { return line; });
  const bulletLines = lines.filter(function(line) {
    return /^([\-*\u30fb\u2022\u25cf\d]+[\.)]?\s+)/.test(line);
  });
  const headingLines = lines.filter(function(line) {
    return /[:：]$/.test(line) && line.length <= 40;
  });
  const important = bulletLines.length > 0 ? bulletLines : headingLines;
  let summary = '';
  if (important.length > 0) {
    summary = important.slice(0, 3).join(' / ');
  } else {
    summary = lines.slice(0, 3).join(' ');
  }
  return truncateText_(summary, maxLen);
}

/**
 * Truncate text to the desired length while preserving whole Unicode characters.
 * @param {string} text
 * @param {number} maxLen
 * @return {string}
 * @private
 */
function truncateText_(text, maxLen) {
  if (!text || text.length <= maxLen) {
    return text || '';
  }
  const truncated = text.substring(0, maxLen);
  return truncated.replace(/[\s\u3000]+$/g, '') + '…';
}

/**
 * Parse a name and email address from a header string.
 * @param {string} fromStr
 * @return {{name: string, email: string}}
 * @private
 */
function parseNameAndEmail_(fromStr) {
  if (!fromStr) {
    return { name: '', email: '' };
  }
  const emailMatch = fromStr.match(/<([^>]+)>/);
  if (emailMatch) {
    const email = emailMatch[1].trim();
    const namePart = fromStr.replace(emailMatch[0], '').trim().replace(/^"|"$/g, '');
    return {
      name: namePart,
      email: email
    };
  }
  if (/^\S+@\S+$/.test(fromStr)) {
    return { name: '', email: fromStr };
  }
  const parts = fromStr.split(/\s+/);
  const possibleEmail = parts.find(function(part) { return /@/.test(part); });
  if (possibleEmail) {
    const name = fromStr.replace(possibleEmail, '').trim();
    return { name: name.replace(/^"|"$/g, ''), email: possibleEmail.replace(/["<>]/g, '') };
  }
  return { name: fromStr.replace(/^"|"$/g, ''), email: '' };
}

/**
 * Fetch (or create) the Gmail label used to mark processed threads.
 * @param {string} name
 * @return {GoogleAppsScript.Gmail.GmailLabel}
 * @private
 */
function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

/**
 * Apply all data validations (dropdowns, numeric ranges, checkboxes, phone regex) to the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @private
 */
function applyDataValidations_(sheet) {
  const startRow = 2;
  const numRows = DATA_VALIDATION_ROWS;

  // Gender validation
  const genderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['男性', '女性', 'その他', '回答しない'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, 4, numRows, 1).setDataValidation(genderRule);

  // Age validation
  const ageRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 120)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, 3, numRows, 1).setDataValidation(ageRule);

  // Student flag validation
  const studentRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['学生', '非学生'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, 9, numRows, 1).setDataValidation(studentRule);

  // Purchase frequency validation (two columns)
  const freqOptions = ['2-3ヶ月に1回', '1ヶ月に1回', '1ヶ月に2-3回', '週に1回', '週に1回以上'];
  const freqRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(freqOptions, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, 12, numRows, 1).setDataValidation(freqRule);
  sheet.getRange(startRow, 13, numRows, 1).setDataValidation(freqRule);

  // Availability checkboxes (four columns)
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(startRow, 14, numRows, 4).setDataValidation(checkboxRule);

  // Phone number validation using REGEXMATCH via requireFormulaSatisfied
  const formula = '=OR(ISBLANK(E2), REGEXMATCH(E2, "' + PHONE_REGEX.replace(/"/g, '""') + '"))';
  const phoneRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied(formula)
    .setAllowInvalid(true)
    .setHelpText('例: 090-1234-5678 や +81 90 1234 5678 などの形式で入力してください。')
    .build();
  sheet.getRange(startRow, 5, numRows, 1).setDataValidation(phoneRule);
}
