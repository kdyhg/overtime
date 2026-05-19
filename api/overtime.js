const { google } = require('googleapis');

const APP_PIN = process.env.APP_PIN || '0992';
const TEMPLATE_SHEET_NAME = process.env.TEMPLATE_SHEET_NAME || '양식';
const DB_SHEET_NAME = process.env.DB_SHEET_NAME || '통합DB';
const SETTINGS_SHEET_NAME = process.env.SETTINGS_SHEET_NAME || '설정';
const HISTORY_SHEET_NAME = process.env.HISTORY_SHEET_NAME || '수정이력';
const INTERNAL_SHEETS = new Set([
  TEMPLATE_SHEET_NAME,
  DB_SHEET_NAME,
  SETTINGS_SHEET_NAME,
  HISTORY_SHEET_NAME,
]);
const WEEKDAYS = ['월', '화', '수', '목', '금', '토', '일'];

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'POST 요청만 허용됩니다.' });
  }

  try {
    const body = parseBody(req.body);

    if (!isAuthorized(req, body)) {
      return res.status(401).json({ error: '비밀번호가 올바르지 않습니다.', code: 'UNAUTHORIZED' });
    }

    const { sheets, spreadsheetId } = createSheetsClient();
    const action = body.action;

    if (action === 'getInitialAppData') {
      const meta = await getSheetMeta(sheets, spreadsheetId);
      const sheetNames = getVisibleMonthSheets(meta.sheetNames);
      const currentSheet = pickDefaultSheet(sheetNames, body.sheetName);
      const data = currentSheet ? await getSheetData(sheets, spreadsheetId, currentSheet) : emptySheetData();
      const history = await getHistoryRows(sheets, spreadsheetId);

      return res.status(200).json({
        sheetNames,
        currentSheet,
        history,
        ...data,
      });
    }

    if (action === 'getSheetData') {
      requireField(body.sheetName, '시트 이름이 없습니다.');
      const data = await getSheetData(sheets, spreadsheetId, body.sheetName);
      const history = await getHistoryRows(sheets, spreadsheetId);
      return res.status(200).json({ history, ...data });
    }

    if (action === 'getHistory') {
      const history = await getHistoryRows(sheets, spreadsheetId);
      return res.status(200).json({ history });
    }

    if (action === 'saveOvertime') {
      requireField(body.sheetName, '시트 이름이 없습니다.');
      requireField(body.dateText, '날짜가 없습니다.');

      const target = await findDateRow(sheets, spreadsheetId, body.sheetName, body.dateText);
      if (!target) {
        return res.status(200).json({ success: false, message: '해당 날짜를 시트에서 찾을 수 없습니다.' });
      }

      const startTime = body.startTime || '';
      const endTime = body.endTime || '';
      const memo = body.memo || '';

      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: [
            {
              range: `${quoteSheetName(body.sheetName)}!C${target.rowNumber}:D${target.rowNumber}`,
              values: [[startTime, endTime]],
            },
            {
              range: `${quoteSheetName(body.sheetName)}!H${target.rowNumber}:H${target.rowNumber}`,
              values: [[memo]],
            },
          ],
        },
      });

      await appendHistory(sheets, spreadsheetId, {
        action: body.source === 'quick-start' ? '빠른 시작 저장' : body.source === 'quick-end' ? '빠른 종료 저장' : '근무 기록 저장',
        sheetName: body.sheetName,
        dateText: body.dateText,
        startTime,
        endTime,
        memo,
        detail: formatChangeDetail(target.row, { startTime, endTime, memo }),
      });

      const updatedData = await getSheetData(sheets, spreadsheetId, body.sheetName);
      const history = await getHistoryRows(sheets, spreadsheetId);
      return res.status(200).json({ success: true, history, ...updatedData });
    }

    if (action === 'saveAfterSchoolSettings') {
      requireField(body.sheetName, '시트 이름이 없습니다.');
      if (!Array.isArray(body.settings)) {
        return res.status(400).json({ success: false, message: '방과후 요일 설정 값이 올바르지 않습니다.' });
      }

      const writeData = body.settings.slice(0, 7).map((item) => [item ? 'O' : '']);
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${quoteSheetName(body.sheetName)}!L2:L8`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: writeData },
      });

      await appendHistory(sheets, spreadsheetId, {
        action: '방과후 요일 저장',
        sheetName: body.sheetName,
        detail: body.settings.map((checked, index) => `${WEEKDAYS[index] || index + 1}:${checked ? 'O' : '-'}`).join(', '),
      });

      const updatedData = await getSheetData(sheets, spreadsheetId, body.sheetName);
      const history = await getHistoryRows(sheets, spreadsheetId);
      return res.status(200).json({ success: true, history, ...updatedData });
    }

    if (action === 'createNewMonthSheet') {
      const targetName = normalizeMonthSheetName(body.targetName);
      if (!targetName) {
        return res.status(400).json({ success: false, message: '시트 이름은 26-5처럼 연도-월 형식으로 입력해 주세요.' });
      }

      const meta = await getSheetMeta(sheets, spreadsheetId);
      if (meta.sheetNames.includes(targetName)) {
        return res.status(409).json({ success: false, message: `${targetName} 시트가 이미 있습니다.` });
      }

      const templateSheet = meta.sheets.find((sheet) => sheet.properties.title === TEMPLATE_SHEET_NAME);
      if (!templateSheet) {
        return res.status(400).json({ success: false, message: `${TEMPLATE_SHEET_NAME} 시트를 찾을 수 없습니다.` });
      }

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              duplicateSheet: {
                sourceSheetId: templateSheet.properties.sheetId,
                insertSheetIndex: 0,
                newSheetName: targetName,
              },
            },
          ],
        },
      });

      const month = Number(targetName.split('-')[1]);
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${quoteSheetName(targetName)}!A2`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[`=FILTER('${DB_SHEET_NAME}'!A:A, MONTH('${DB_SHEET_NAME}'!A:A)=${month})`]],
        },
      });

      await appendHistory(sheets, spreadsheetId, {
        action: '새 월 시트 생성',
        sheetName: targetName,
        detail: `${TEMPLATE_SHEET_NAME} 시트를 복사했습니다.`,
      });

      const refreshedMeta = await getSheetMeta(sheets, spreadsheetId);
      const sheetNames = getVisibleMonthSheets(refreshedMeta.sheetNames);
      const data = await getSheetData(sheets, spreadsheetId, targetName);
      const history = await getHistoryRows(sheets, spreadsheetId);

      return res.status(200).json({
        success: true,
        newSheetName: targetName,
        sheetNames,
        currentSheet: targetName,
        history,
        ...data,
      });
    }

    return res.status(400).json({ error: '알 수 없는 요청입니다.' });
  } catch (error) {
    console.error(error);
    return res.status(error.statusCode || 500).json({
      error: error.userMessage || error.message || '서버 처리 중 오류가 발생했습니다.',
    });
  }
};

function parseBody(body) {
  if (!body) return {};
  if (typeof body === 'object') return body;
  if (typeof body !== 'string') return {};

  try {
    return JSON.parse(body);
  } catch {
    return {};
  }
}

function isAuthorized(req, body) {
  const headerPin = req.headers['x-app-pin'];
  const suppliedPin = Array.isArray(headerPin) ? headerPin[0] : headerPin || body.pin;
  return String(suppliedPin || '') === String(APP_PIN);
}

function createSheetsClient() {
  const required = ['GOOGLE_CLIENT_EMAIL', 'GOOGLE_PRIVATE_KEY', 'GOOGLE_SHEET_ID'];
  const missing = required.filter((name) => !process.env[name]);
  if (missing.length) {
    const error = new Error(`Vercel 환경변수 누락: ${missing.join(', ')}`);
    error.statusCode = 500;
    error.userMessage = `서버 환경변수 누락: ${missing.join(', ')}`;
    throw error;
  }

  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_CLIENT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  return {
    sheets: google.sheets({ version: 'v4', auth }),
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
  };
}

async function getSheetMeta(sheets, spreadsheetId) {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets.properties(sheetId,title)',
  });
  const sheetList = meta.data.sheets || [];
  return {
    sheets: sheetList,
    sheetNames: sheetList.map((sheet) => sheet.properties.title),
  };
}

function getVisibleMonthSheets(sheetNames) {
  return sheetNames
    .filter((name) => !INTERNAL_SHEETS.has(name))
    .sort(compareMonthSheetNames);
}

function pickDefaultSheet(sheetNames, requestedName) {
  if (requestedName && sheetNames.includes(requestedName)) return requestedName;

  const now = new Date();
  const currentMonthName = `${String(now.getFullYear()).slice(-2)}-${now.getMonth() + 1}`;
  if (sheetNames.includes(currentMonthName)) return currentMonthName;

  return sheetNames[0] || '';
}

function emptySheetData() {
  return {
    summary: { totalHours: '', approvedHours: '', hourlyRate: '', totalPay: '' },
    dates: [],
    afterSchool: WEEKDAYS.map((day) => ({ day, isSelected: false })),
    records: {},
  };
}

async function getSheetData(sheets, spreadsheetId, name) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${quoteSheetName(name)}!A:L`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const rows = response.data.values || [];

  const summary = {
    totalHours: rows[1]?.[9] || '',
    approvedHours: rows[2]?.[9] || '',
    hourlyRate: rows[3]?.[9] || '',
    totalPay: rows[4]?.[9] || '',
  };

  const dates = [];
  const records = {};
  const afterSchool = [];

  for (let index = 1; index < rows.length; index += 1) {
    const row = rows[index] || [];
    const displayDate = row[0] || '';
    if (!displayDate || displayDate.includes('날짜') || !displayDate.trim()) continue;

    const normalizedDate = normalizeDate(displayDate);
    const record = {
      dateText: displayDate,
      startTime: row[2] || '',
      endTime: row[3] || '',
      morningExtra: row[4] || '0:00',
      afternoonExtra: row[5] || '0:00',
      dailyTotal: row[6] || '0:00',
      memo: row[7] || '',
    };

    dates.push({ dateText: displayDate, key: normalizedDate });
    records[normalizedDate] = record;
  }

  for (let index = 1; index < 8; index += 1) {
    afterSchool.push({
      day: rows[index]?.[10] || WEEKDAYS[index - 1],
      isSelected: rows[index]?.[11] === 'O',
    });
  }

  return { summary, dates, afterSchool, records };
}

async function findDateRow(sheets, spreadsheetId, sheetName, dateText) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${quoteSheetName(sheetName)}!A:H`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const rows = response.data.values || [];
  const targetDate = normalizeDate(dateText);

  for (let index = 1; index < rows.length; index += 1) {
    const sheetDate = normalizeDate(rows[index]?.[0] || '');
    if (sheetDate === targetDate) {
      return {
        rowNumber: index + 1,
        row: rows[index] || [],
      };
    }
  }

  return null;
}

function formatChangeDetail(previousRow, next) {
  const beforeStart = previousRow?.[2] || '';
  const beforeEnd = previousRow?.[3] || '';
  const beforeMemo = previousRow?.[7] || '';
  return `시작 ${beforeStart || '-'} -> ${next.startTime || '-'}, 종료 ${beforeEnd || '-'} -> ${next.endTime || '-'}, 메모 ${beforeMemo || '-'} -> ${next.memo || '-'}`;
}

async function getHistoryRows(sheets, spreadsheetId) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${quoteSheetName(HISTORY_SHEET_NAME)}!A2:H`,
      valueRenderOption: 'FORMATTED_VALUE',
    });
    const rows = response.data.values || [];
    return rows
      .map((row) => ({
        timestamp: row[0] || '',
        action: row[1] || '',
        sheetName: row[2] || '',
        dateText: row[3] || '',
        startTime: row[4] || '',
        endTime: row[5] || '',
        memo: row[6] || '',
        detail: row[7] || '',
      }))
      .reverse()
      .slice(0, 30);
  } catch (error) {
    if (error.code === 400 || error.code === 404) return [];
    throw error;
  }
}

async function appendHistory(sheets, spreadsheetId, item) {
  try {
    await ensureHistorySheet(sheets, spreadsheetId);
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${quoteSheetName(HISTORY_SHEET_NAME)}!A:H`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }),
          item.action || '',
          item.sheetName || '',
          item.dateText || '',
          item.startTime || '',
          item.endTime || '',
          item.memo || '',
          item.detail || '',
        ]],
      },
    });
  } catch (error) {
    console.error('수정 이력 기록 실패:', error);
  }
}

async function ensureHistorySheet(sheets, spreadsheetId) {
  const meta = await getSheetMeta(sheets, spreadsheetId);
  if (!meta.sheetNames.includes(HISTORY_SHEET_NAME)) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: HISTORY_SHEET_NAME } } }],
      },
    });
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${quoteSheetName(HISTORY_SHEET_NAME)}!A1:H1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [['시간', '작업', '시트', '날짜', '시작', '종료', '메모', '상세']],
    },
  });
}

function normalizeMonthSheetName(value) {
  const match = String(value || '').trim().match(/^(\d{2}|\d{4})-(\d{1,2})$/);
  if (!match) return '';

  const year = match[1].slice(-2);
  const month = Number(match[2]);
  if (!Number.isInteger(month) || month < 1 || month > 12) return '';

  return `${year}-${month}`;
}

function compareMonthSheetNames(left, right) {
  const parse = (value) => {
    const normalized = normalizeMonthSheetName(value);
    if (!normalized) return 0;
    const [year, month] = normalized.split('-').map(Number);
    return year * 100 + month;
  };
  return parse(right) - parse(left);
}

function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function normalizeDate(value) {
  return String(value || '').replace(/\s+/g, '');
}

function requireField(value, message) {
  if (value === undefined || value === null || value === '') {
    const error = new Error(message);
    error.statusCode = 400;
    error.userMessage = message;
    throw error;
  }
}
