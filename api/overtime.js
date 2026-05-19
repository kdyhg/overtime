const crypto = require('crypto');
const { google } = require('googleapis');

const TEMPLATE_SHEET_NAME = process.env.TEMPLATE_SHEET_NAME || '양식';
const DB_SHEET_NAME = process.env.DB_SHEET_NAME || '통합DB';
const HISTORY_SHEET_NAME = process.env.HISTORY_SHEET_NAME || '수정이력';
const USERS_SHEET_NAME = process.env.USERS_SHEET_NAME || '사용자';
const LOCKS_SHEET_NAME = process.env.LOCKS_SHEET_NAME || '월마감';
const DEFAULT_AFTER_SCHOOL_DEDUCTION_MINUTES = 50;
const STANDARD_START_MINUTES = parseTimeToMinutes(process.env.STANDARD_START_TIME || '08:30') ?? 510;
const STANDARD_END_MINUTES = parseTimeToMinutes(process.env.STANDARD_END_TIME || '16:30') ?? 990;

const INTERNAL_SHEETS = new Set([
  TEMPLATE_SHEET_NAME,
  DB_SHEET_NAME,
  process.env.SETTINGS_SHEET_NAME || '설정',
  HISTORY_SHEET_NAME,
  USERS_SHEET_NAME,
  LOCKS_SHEET_NAME,
]);
const WEEKDAYS = ['월', '화', '수', '목', '금', '토', '일'];
const PUBLIC_ACTIONS = new Set(['login', 'signup', 'findUsername', 'resetPassword']);

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'POST 요청만 허용됩니다.' });

  try {
    const body = parseBody(req.body);
    const { sheets, spreadsheetId } = createSheetsClient();
    await ensureUserStore(sheets, spreadsheetId);

    if (PUBLIC_ACTIONS.has(body.action)) {
      const result = await handlePublicAction(body.action, body, sheets, spreadsheetId);
      return res.status(result.status || 200).json(result.body);
    }

    const user = await authenticateRequest(req, body, sheets, spreadsheetId);
    const result = await handlePrivateAction(body.action, body, sheets, spreadsheetId, user);
    return res.status(result.status || 200).json(result.body);
  } catch (error) {
    console.error(error);
    return res.status(error.statusCode || 500).json({
      error: error.userMessage || error.message || '서버 처리 중 오류가 발생했습니다.',
      code: error.code,
    });
  }
};

async function handlePublicAction(action, body, sheets, spreadsheetId) {
  if (action === 'login') {
    const user = await findUserByUsername(sheets, spreadsheetId, normalizeUsername(body.username));
    if (!user || user.active !== 'Y' || !verifyUserPassword(user, body.password || '')) {
      return { status: 401, body: { error: '아이디 또는 비밀번호가 올바르지 않습니다.' } };
    }
    return { body: { token: issueToken(user), user: publicUser(user), ...(await buildAppData(sheets, spreadsheetId, user)) } };
  }

  if (action === 'signup') {
    const users = await getUsers(sheets, spreadsheetId);
    const username = normalizeUsername(body.username);
    const displayName = normalizeDisplayName(body.displayName);
    const password = String(body.password || '');
    const recoveryCode = String(body.recoveryCode || '');
    const validationError = validateNewUser(username, displayName, password, recoveryCode);
    if (validationError) return { status: 400, body: { error: validationError } };
    if (users.some((user) => user.username === username)) return { status: 409, body: { error: '이미 사용 중인 아이디입니다.' } };

    const user = await appendUser(sheets, spreadsheetId, {
      username,
      displayName,
      password,
      recoveryCode,
      role: users.length ? 'user' : 'admin',
    });
    return { body: { token: issueToken(user), user: publicUser(user), ...(await buildAppData(sheets, spreadsheetId, user)) } };
  }

  if (action === 'findUsername') {
    const displayName = normalizeDisplayName(body.displayName);
    const recoveryCode = String(body.recoveryCode || '');
    if (!displayName || !recoveryCode) return { status: 400, body: { error: '이름과 복구코드를 입력해 주세요.' } };
    const usernames = (await getUsers(sheets, spreadsheetId))
      .filter((user) => user.active === 'Y' && user.displayName === displayName && verifyRecoveryCode(user, recoveryCode))
      .map((user) => user.username);
    return { body: { usernames } };
  }

  if (action === 'resetPassword') {
    const username = normalizeUsername(body.username);
    const recoveryCode = String(body.recoveryCode || '');
    const newPassword = String(body.newPassword || '');
    if (!username || !recoveryCode || !newPassword) return { status: 400, body: { error: '아이디, 복구코드, 새 비밀번호를 모두 입력해 주세요.' } };
    if (newPassword.length < 8) return { status: 400, body: { error: '새 비밀번호는 8자 이상이어야 합니다.' } };

    const user = await findUserByUsername(sheets, spreadsheetId, username);
    if (!user || user.active !== 'Y' || !verifyRecoveryCode(user, recoveryCode)) {
      return { status: 401, body: { error: '아이디 또는 복구코드가 올바르지 않습니다.' } };
    }

    const next = hashSecret(newPassword);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [
          { range: `${quoteSheetName(USERS_SHEET_NAME)}!C${user.rowNumber}:D${user.rowNumber}`, values: [[next.salt, next.hash]] },
          { range: `${quoteSheetName(USERS_SHEET_NAME)}!I${user.rowNumber}:I${user.rowNumber}`, values: [[nowText()]] },
        ],
      },
    });
    return { body: { success: true } };
  }

  return { status: 400, body: { error: '알 수 없는 요청입니다.' } };
}

async function handlePrivateAction(action, body, sheets, spreadsheetId, user) {
  if (action === 'getInitialAppData') return { body: await buildAppData(sheets, spreadsheetId, user, body.sheetName) };
  if (action === 'getSheetData') {
    requireField(body.sheetName, '시트 이름이 없습니다.');
    return { body: await buildAppData(sheets, spreadsheetId, user, body.sheetName) };
  }
  if (action === 'getHistory') return { body: { history: await getHistoryRows(sheets, spreadsheetId, user) } };
  if (action === 'saveOvertime') return saveOvertime(sheets, spreadsheetId, user, body);
  if (action === 'saveAfterSchoolSettings') return saveAfterSchoolSettings(sheets, spreadsheetId, user, body);
  if (action === 'createNewMonthSheet') return createNewMonthSheet(sheets, spreadsheetId, user, body);
  if (action === 'lockMonth') return setMonthLocked(sheets, spreadsheetId, user, body.sheetName, true);
  if (action === 'unlockMonth') {
    if (!verifyUserPassword(user, body.password || '')) return { status: 401, body: { error: '비밀번호가 올바르지 않습니다.' } };
    return setMonthLocked(sheets, spreadsheetId, user, body.sheetName, false);
  }
  return { status: 400, body: { error: '알 수 없는 요청입니다.' } };
}

async function buildAppData(sheets, spreadsheetId, user, requestedMonth = '') {
  const meta = await getSheetMeta(sheets, spreadsheetId);
  const sheetNames = getUserMonthSheetPairs(meta.sheetNames, user).map((pair) => pair.displayName);
  const currentSheet = pickDefaultSheet(sheetNames, requestedMonth);
  const physicalSheet = currentSheet ? resolvePhysicalSheetName(user, currentSheet, meta.sheetNames) : '';
  const data = physicalSheet ? await getSheetData(sheets, spreadsheetId, physicalSheet) : emptySheetData();
  const locks = await getLocksMap(sheets, spreadsheetId, user);
  return {
    user: publicUser(user),
    sheetNames,
    currentSheet,
    locked: Boolean(locks[currentSheet]),
    locks,
    history: await getHistoryRows(sheets, spreadsheetId, user),
    ...data,
  };
}

async function saveOvertime(sheets, spreadsheetId, user, body) {
  requireField(body.sheetName, '시트 이름이 없습니다.');
  requireField(body.dateText, '날짜가 없습니다.');
  const meta = await getSheetMeta(sheets, spreadsheetId);
  const physicalSheet = resolvePhysicalSheetName(user, body.sheetName, meta.sheetNames);
  if (!physicalSheet) return { status: 404, body: { success: false, message: '시트를 찾을 수 없습니다.' } };
  const lockCheck = await requireUnlockedOrOverride(sheets, spreadsheetId, user, body.sheetName, body.overridePassword);
  if (!lockCheck.ok) return lockCheck;

  const target = await findDateRow(sheets, spreadsheetId, physicalSheet, body.dateText);
  if (!target) return { body: { success: false, message: '해당 날짜를 시트에서 찾을 수 없습니다.' } };

  const startTime = body.startTime || '';
  const endTime = body.endTime || '';
  const memo = body.memo || '';
  const config = await getAfterSchoolConfig(sheets, spreadsheetId, physicalSheet);
  const afterSchoolOverride = normalizeAfterSchoolOverride(body.afterSchoolOverride ?? target.row[8] ?? '');
  const afterSchoolApplied = isAfterSchoolApplied(afterSchoolOverride, body.dateText, physicalSheet, config.settings);
  const computed = calculateOvertimeFields(startTime, endTime, afterSchoolApplied, config.deductionMinutes);
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [
        { range: `${quoteSheetName(physicalSheet)}!C${target.rowNumber}:I${target.rowNumber}`, values: [[startTime, endTime, computed.morningExtra, computed.afternoonExtra, computed.dailyTotal, memo, afterSchoolOverride]] },
      ],
    },
  });
  await appendHistory(sheets, spreadsheetId, user, {
    action: body.source === 'quick-start' ? '빠른 시작 저장' : body.source === 'quick-end' ? '빠른 종료 저장' : '근무 기록 저장',
    sheetName: body.sheetName,
    dateText: body.dateText,
    startTime,
    endTime,
    memo,
    detail: `${formatChangeDetail(target.row, { startTime, endTime, memo })}, 방과후 ${afterSchoolApplied ? `${config.deductionMinutes}분 차감` : '없음'}`,
  });
  return { body: { success: true, ...(await buildAppData(sheets, spreadsheetId, user, body.sheetName)) } };
}

async function saveAfterSchoolSettings(sheets, spreadsheetId, user, body) {
  requireField(body.sheetName, '시트 이름이 없습니다.');
  if (!Array.isArray(body.settings)) return { status: 400, body: { success: false, message: '방과후 요일 설정 값이 올바르지 않습니다.' } };
  const settings = WEEKDAYS.map((_, index) => Boolean(body.settings[index]));
  const deductionMinutes = normalizeDeductionMinutes(body.deductionMinutes);
  const meta = await getSheetMeta(sheets, spreadsheetId);
  const physicalSheet = resolvePhysicalSheetName(user, body.sheetName, meta.sheetNames);
  if (!physicalSheet) return { status: 404, body: { success: false, message: '시트를 찾을 수 없습니다.' } };
  const lockCheck = await requireUnlockedOrOverride(sheets, spreadsheetId, user, body.sheetName, body.overridePassword);
  if (!lockCheck.ok) return lockCheck;

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [
        { range: `${quoteSheetName(physicalSheet)}!I1`, values: [['방과후']] },
        { range: `${quoteSheetName(physicalSheet)}!L2:L8`, values: settings.map((item) => [item ? 'O' : '']) },
        { range: `${quoteSheetName(physicalSheet)}!M1:M2`, values: [['방과후 차감(분)'], [deductionMinutes]] },
      ],
    },
  });
  await refreshCalculatedRows(sheets, spreadsheetId, physicalSheet, settings, deductionMinutes);
  await appendHistory(sheets, spreadsheetId, user, {
    action: '방과후 설정 저장',
    sheetName: body.sheetName,
    detail: `${settings.map((checked, index) => `${WEEKDAYS[index] || index + 1}:${checked ? 'O' : '-'}`).join(', ')}, 차감 ${deductionMinutes}분`,
  });
  return { body: { success: true, ...(await buildAppData(sheets, spreadsheetId, user, body.sheetName)) } };
}

async function createNewMonthSheet(sheets, spreadsheetId, user, body) {
  const monthName = normalizeMonthSheetName(body.targetName);
  if (!monthName) return { status: 400, body: { success: false, message: '시트 이름은 26-5처럼 연도-월 형식으로 입력해 주세요.' } };
  const meta = await getSheetMeta(sheets, spreadsheetId);
  const visibleSheets = getUserMonthSheetPairs(meta.sheetNames, user).map((pair) => pair.displayName);
  if (visibleSheets.includes(monthName)) return { status: 409, body: { success: false, message: `${monthName} 시트가 이미 있습니다.` } };
  const physicalSheet = toPhysicalMonthSheetName(user, monthName);
  const templateSheet = meta.sheets.find((sheet) => sheet.properties.title === TEMPLATE_SHEET_NAME);
  if (!templateSheet) return { status: 400, body: { success: false, message: `${TEMPLATE_SHEET_NAME} 시트를 찾을 수 없습니다.` } };

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{ duplicateSheet: { sourceSheetId: templateSheet.properties.sheetId, insertSheetIndex: 0, newSheetName: physicalSheet } }],
    },
  });
  const month = Number(monthName.split('-')[1]);
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${quoteSheetName(physicalSheet)}!A2`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[`=FILTER('${DB_SHEET_NAME}'!A:A, MONTH('${DB_SHEET_NAME}'!A:A)=${month})`]] },
  });
  await appendHistory(sheets, spreadsheetId, user, { action: '새 월 시트 생성', sheetName: monthName, detail: `${TEMPLATE_SHEET_NAME} 시트를 복사했습니다.` });
  return { body: { success: true, newSheetName: monthName, ...(await buildAppData(sheets, spreadsheetId, user, monthName)) } };
}

async function setMonthLocked(sheets, spreadsheetId, user, sheetName, locked) {
  requireField(sheetName, '시트 이름이 없습니다.');
  const monthName = normalizeMonthSheetName(sheetName);
  if (!monthName) return { status: 400, body: { success: false, message: '월 시트 이름이 올바르지 않습니다.' } };
  await ensureLocksSheet(sheets, spreadsheetId);
  const rows = await getLockRows(sheets, spreadsheetId);
  const existing = rows.find((row) => row.username === user.username && row.sheetName === monthName);
  const values = [[user.username, monthName, locked ? 'Y' : 'N', nowText(), user.displayName]];
  if (existing) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${quoteSheetName(LOCKS_SHEET_NAME)}!A${existing.rowNumber}:E${existing.rowNumber}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values },
    });
  } else {
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${quoteSheetName(LOCKS_SHEET_NAME)}!A:E`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values },
    });
  }
  await appendHistory(sheets, spreadsheetId, user, { action: locked ? '월 마감' : '월 마감 해제', sheetName: monthName, detail: locked ? '월 기록을 잠갔습니다.' : '월 기록 잠금을 해제했습니다.' });
  return { body: { success: true, ...(await buildAppData(sheets, spreadsheetId, user, monthName)) } };
}

async function requireUnlockedOrOverride(sheets, spreadsheetId, user, sheetName, overridePassword) {
  if (!(await isMonthLocked(sheets, spreadsheetId, user, sheetName))) return { ok: true };
  if (overridePassword && verifyUserPassword(user, overridePassword)) return { ok: true };
  return {
    ok: false,
    status: 423,
    body: { success: false, code: 'MONTH_LOCKED', message: '마감된 월입니다. 수정하려면 비밀번호를 다시 입력해 주세요.' },
  };
}

async function getSheetData(sheets, spreadsheetId, physicalSheetName) {
  const response = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${quoteSheetName(physicalSheetName)}!A:M`, valueRenderOption: 'FORMATTED_VALUE' });
  const rows = response.data.values || [];
  const summary = { totalHours: rows[1]?.[9] || '', approvedHours: rows[2]?.[9] || '', hourlyRate: rows[3]?.[9] || '', totalPay: rows[4]?.[9] || '' };
  const afterSchool = getAfterSchoolSettingsFromRows(rows);
  const defaultSettings = afterSchool.map((item) => item.isSelected);
  const afterSchoolDeductionMinutes = normalizeDeductionMinutes(rows[1]?.[12]);
  const dates = [];
  const records = {};
  for (let index = 1; index < rows.length; index += 1) {
    const row = rows[index] || [];
    const displayDate = row[0] || '';
    if (!isDateText(displayDate)) continue;
    const normalizedDate = normalizeDate(displayDate);
    const afterSchoolOverride = normalizeAfterSchoolOverride(row[8] || '');
    const afterSchoolApplied = isAfterSchoolApplied(afterSchoolOverride, displayDate, physicalSheetName, defaultSettings);
    const record = { dateText: displayDate, startTime: row[2] || '', endTime: row[3] || '', morningExtra: row[4] || '0:00', afternoonExtra: row[5] || '0:00', dailyTotal: row[6] || '0:00', memo: row[7] || '', afterSchoolOverride, afterSchoolApplied };
    dates.push({ dateText: displayDate, key: normalizedDate, afterSchoolApplied });
    records[normalizedDate] = record;
  }
  return { summary, dates, afterSchool, afterSchoolDeductionMinutes, records };
}

async function findDateRow(sheets, spreadsheetId, physicalSheetName, dateText) {
  const response = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${quoteSheetName(physicalSheetName)}!A:I`, valueRenderOption: 'FORMATTED_VALUE' });
  const rows = response.data.values || [];
  const targetDate = normalizeDate(dateText);
  for (let index = 1; index < rows.length; index += 1) {
    if (normalizeDate(rows[index]?.[0] || '') === targetDate) return { rowNumber: index + 1, row: rows[index] || [] };
  }
  return null;
}

async function getAfterSchoolConfig(sheets, spreadsheetId, physicalSheetName) {
  const response = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${quoteSheetName(physicalSheetName)}!K1:M8`, valueRenderOption: 'FORMATTED_VALUE' });
  const rows = response.data.values || [];
  const settings = WEEKDAYS.map((_, index) => rows[index + 1]?.[1] === 'O');
  const deductionMinutes = normalizeDeductionMinutes(rows[1]?.[2]);
  return { settings, deductionMinutes };
}

function getAfterSchoolSettingsFromRows(rows) {
  return WEEKDAYS.map((day, index) => ({ day: rows[index + 1]?.[10] || day, isSelected: rows[index + 1]?.[11] === 'O' }));
}

async function refreshCalculatedRows(sheets, spreadsheetId, physicalSheetName, settings, deductionMinutes) {
  const response = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${quoteSheetName(physicalSheetName)}!A2:I`, valueRenderOption: 'FORMATTED_VALUE' });
  const rows = response.data.values || [];
  const data = rows
    .map((row, index) => ({ row, rowNumber: index + 2 }))
    .filter(({ row }) => isDateText(row[0]))
    .map(({ row, rowNumber }) => {
      const afterSchoolApplied = isAfterSchoolApplied(row[8] || '', row[0], physicalSheetName, settings);
      const computed = calculateOvertimeFields(row[2] || '', row[3] || '', afterSchoolApplied, deductionMinutes);
      return { range: `${quoteSheetName(physicalSheetName)}!E${rowNumber}:G${rowNumber}`, values: [[computed.morningExtra, computed.afternoonExtra, computed.dailyTotal]] };
    });
  if (!data.length) return;
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: 'USER_ENTERED', data },
  });
}

async function getHistoryRows(sheets, spreadsheetId, user) {
  try {
    const response = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${quoteSheetName(HISTORY_SHEET_NAME)}!A2:J`, valueRenderOption: 'FORMATTED_VALUE' });
    return (response.data.values || [])
      .map((row) => ({ timestamp: row[0] || '', action: row[1] || '', sheetName: row[2] || '', dateText: row[3] || '', startTime: row[4] || '', endTime: row[5] || '', memo: row[6] || '', detail: row[7] || '', username: row[8] || '', displayName: row[9] || '' }))
      .filter((item) => item.username === user.username || (!item.username && user.role === 'admin'))
      .reverse()
      .slice(0, 50);
  } catch (error) {
    if (error.code === 400 || error.code === 404) return [];
    throw error;
  }
}

async function appendHistory(sheets, spreadsheetId, user, item) {
  try {
    await ensureHistorySheet(sheets, spreadsheetId);
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${quoteSheetName(HISTORY_SHEET_NAME)}!A:J`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [[nowText(), item.action || '', item.sheetName || '', item.dateText || '', item.startTime || '', item.endTime || '', item.memo || '', item.detail || '', user.username, user.displayName]] },
    });
  } catch (error) {
    console.error('수정 이력 기록 실패:', error);
  }
}

async function ensureHistorySheet(sheets, spreadsheetId) {
  const meta = await getSheetMeta(sheets, spreadsheetId);
  if (!meta.sheetNames.includes(HISTORY_SHEET_NAME)) await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [{ addSheet: { properties: { title: HISTORY_SHEET_NAME } } }] } });
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${quoteSheetName(HISTORY_SHEET_NAME)}!A1:J1`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['시간', '작업', '시트', '날짜', '시작', '종료', '메모', '상세', '사용자ID', '사용자명']] } });
}

async function ensureUserStore(sheets, spreadsheetId) {
  const meta = await getSheetMeta(sheets, spreadsheetId);
  if (!meta.sheetNames.includes(USERS_SHEET_NAME)) await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [{ addSheet: { properties: { title: USERS_SHEET_NAME } } }] } });
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${quoteSheetName(USERS_SHEET_NAME)}!A1:J1`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['아이디', '이름', '비밀번호Salt', '비밀번호Hash', '복구Salt', '복구Hash', '역할', '생성일', '수정일', '활성']] } });
}

async function appendUser(sheets, spreadsheetId, userInput) {
  const passwordSecret = hashSecret(userInput.password);
  const recoverySecret = hashSecret(userInput.recoveryCode);
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${quoteSheetName(USERS_SHEET_NAME)}!A:J`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [[userInput.username, userInput.displayName, passwordSecret.salt, passwordSecret.hash, recoverySecret.salt, recoverySecret.hash, userInput.role || 'user', nowText(), nowText(), 'Y']] },
  });
  return findUserByUsername(sheets, spreadsheetId, userInput.username);
}

async function getUsers(sheets, spreadsheetId) {
  const response = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${quoteSheetName(USERS_SHEET_NAME)}!A2:J`, valueRenderOption: 'FORMATTED_VALUE' });
  return (response.data.values || [])
    .map((row, index) => ({ rowNumber: index + 2, username: row[0] || '', displayName: row[1] || '', passwordSalt: row[2] || '', passwordHash: row[3] || '', recoverySalt: row[4] || '', recoveryHash: row[5] || '', role: row[6] || 'user', createdAt: row[7] || '', updatedAt: row[8] || '', active: row[9] || 'Y' }))
    .filter((user) => user.username);
}

async function findUserByUsername(sheets, spreadsheetId, username) {
  return (await getUsers(sheets, spreadsheetId)).find((user) => user.username === username) || null;
}

async function authenticateRequest(req, body, sheets, spreadsheetId) {
  const authHeader = req.headers.authorization || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : body.token;
  const payload = verifyToken(token || '');
  if (!payload) {
    const error = new Error('로그인이 필요합니다.');
    error.statusCode = 401;
    error.code = 'UNAUTHORIZED';
    throw error;
  }
  const user = await findUserByUsername(sheets, spreadsheetId, payload.u);
  if (!user || user.active !== 'Y') {
    const error = new Error('사용자 계정을 찾을 수 없습니다.');
    error.statusCode = 401;
    error.code = 'UNAUTHORIZED';
    throw error;
  }
  return user;
}

async function ensureLocksSheet(sheets, spreadsheetId) {
  const meta = await getSheetMeta(sheets, spreadsheetId);
  if (!meta.sheetNames.includes(LOCKS_SHEET_NAME)) await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [{ addSheet: { properties: { title: LOCKS_SHEET_NAME } } }] } });
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${quoteSheetName(LOCKS_SHEET_NAME)}!A1:E1`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['사용자ID', '월', '마감', '변경일', '변경자']] } });
}

async function getLockRows(sheets, spreadsheetId) {
  await ensureLocksSheet(sheets, spreadsheetId);
  const response = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${quoteSheetName(LOCKS_SHEET_NAME)}!A2:E`, valueRenderOption: 'FORMATTED_VALUE' });
  return (response.data.values || [])
    .map((row, index) => ({ rowNumber: index + 2, username: row[0] || '', sheetName: row[1] || '', locked: row[2] === 'Y', updatedAt: row[3] || '', updatedBy: row[4] || '' }))
    .filter((row) => row.username && row.sheetName);
}

async function getLocksMap(sheets, spreadsheetId, user) {
  return (await getLockRows(sheets, spreadsheetId))
    .filter((row) => row.username === user.username)
    .reduce((map, row) => {
      map[row.sheetName] = Boolean(row.locked);
      return map;
    }, {});
}

async function isMonthLocked(sheets, spreadsheetId, user, sheetName) {
  return Boolean((await getLocksMap(sheets, spreadsheetId, user))[sheetName]);
}

function createSheetsClient() {
  const required = ['GOOGLE_CLIENT_EMAIL', 'GOOGLE_PRIVATE_KEY', 'GOOGLE_SHEET_ID'];
  const missing = required.filter((name) => !process.env[name]);
  if (missing.length) {
    const error = new Error(`서버 환경변수 누락: ${missing.join(', ')}`);
    error.statusCode = 500;
    error.userMessage = `서버 환경변수 누락: ${missing.join(', ')}`;
    throw error;
  }
  const auth = new google.auth.GoogleAuth({
    credentials: { client_email: process.env.GOOGLE_CLIENT_EMAIL, private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n') },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return { sheets: google.sheets({ version: 'v4', auth }), spreadsheetId: process.env.GOOGLE_SHEET_ID };
}

async function getSheetMeta(sheets, spreadsheetId) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties(sheetId,title)' });
  const sheetList = meta.data.sheets || [];
  return { sheets: sheetList, sheetNames: sheetList.map((sheet) => sheet.properties.title) };
}

function getUserMonthSheetPairs(sheetNames, user) {
  const prefix = `${user.username}__`;
  return sheetNames
    .filter((name) => !INTERNAL_SHEETS.has(name))
    .map((physicalName) => {
      if (physicalName.startsWith(prefix)) {
        const displayName = physicalName.slice(prefix.length);
        return normalizeMonthSheetName(displayName) ? { physicalName, displayName } : null;
      }
      if (user.role === 'admin' && normalizeMonthSheetName(physicalName)) return { physicalName, displayName: physicalName };
      return null;
    })
    .filter(Boolean)
    .sort((left, right) => compareMonthSheetNames(left.displayName, right.displayName));
}

function resolvePhysicalSheetName(user, displayName, sheetNames) {
  const normalized = normalizeMonthSheetName(displayName);
  if (!normalized) return '';
  const preferred = toPhysicalMonthSheetName(user, normalized);
  if (sheetNames.includes(preferred)) return preferred;
  if (user.role === 'admin' && sheetNames.includes(normalized)) return normalized;
  return '';
}

function toPhysicalMonthSheetName(user, displayName) {
  const normalized = normalizeMonthSheetName(displayName);
  return user.role === 'admin' ? normalized : `${user.username}__${normalized}`;
}

function pickDefaultSheet(sheetNames, requestedName) {
  if (requestedName && sheetNames.includes(requestedName)) return requestedName;
  const now = new Date();
  const currentMonthName = `${String(now.getFullYear()).slice(-2)}-${now.getMonth() + 1}`;
  return sheetNames.includes(currentMonthName) ? currentMonthName : sheetNames[0] || '';
}

function emptySheetData() {
  return { summary: { totalHours: '', approvedHours: '', hourlyRate: '', totalPay: '' }, dates: [], afterSchool: WEEKDAYS.map((day) => ({ day, isSelected: false })), afterSchoolDeductionMinutes: DEFAULT_AFTER_SCHOOL_DEDUCTION_MINUTES, records: {} };
}

function validateNewUser(username, displayName, password, recoveryCode) {
  if (!username || !/^[a-zA-Z0-9_.-]{3,30}$/.test(username)) return '아이디는 영문, 숫자, ., _, - 조합 3~30자로 입력해 주세요.';
  if (!displayName || displayName.length > 30) return '이름은 1~30자로 입력해 주세요.';
  if (!password || password.length < 8) return '비밀번호는 8자 이상이어야 합니다.';
  if (!recoveryCode || recoveryCode.length < 4) return '복구코드는 4자 이상이어야 합니다.';
  return '';
}

function normalizeUsername(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeDisplayName(value) {
  return String(value || '').trim();
}

function hashSecret(secret, salt = crypto.randomBytes(16).toString('hex')) {
  return { salt, hash: crypto.pbkdf2Sync(String(secret), salt, 120000, 32, 'sha256').toString('hex') };
}

function verifyHashedSecret(secret, salt, expectedHash) {
  if (!secret || !salt || !expectedHash) return false;
  const left = Buffer.from(hashSecret(secret, salt).hash, 'hex');
  const right = Buffer.from(expectedHash, 'hex');
  return left.length === right.length && crypto.timingSafeEqual(left, right);
}

function verifyUserPassword(user, password) {
  return verifyHashedSecret(password, user.passwordSalt, user.passwordHash);
}

function verifyRecoveryCode(user, recoveryCode) {
  return verifyHashedSecret(recoveryCode, user.recoverySalt, user.recoveryHash);
}

function issueToken(user) {
  const payload = Buffer.from(JSON.stringify({ u: user.username, r: user.role, exp: Date.now() + 1000 * 60 * 60 * 12 })).toString('base64url');
  return `${payload}.${sign(payload)}`;
}

function verifyToken(token) {
  const [payload, signature] = String(token || '').split('.');
  if (!payload || !signature || sign(payload) !== signature) return null;
  try {
    const data = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
    return data.u && Date.now() <= Number(data.exp || 0) ? data : null;
  } catch {
    return null;
  }
}

function sign(payload) {
  return crypto.createHmac('sha256', getSessionSecret()).update(payload).digest('base64url');
}

function getSessionSecret() {
  if (process.env.APP_SESSION_SECRET) return process.env.APP_SESSION_SECRET;
  if (process.env.GOOGLE_PRIVATE_KEY) return process.env.GOOGLE_PRIVATE_KEY;
  const error = new Error('APP_SESSION_SECRET 환경변수가 필요합니다.');
  error.statusCode = 500;
  throw error;
}

function publicUser(user) {
  return { username: user.username, displayName: user.displayName, role: user.role };
}

function normalizeMonthSheetName(value) {
  const match = String(value || '').trim().match(/^(\d{2}|\d{4})-(\d{1,2})$/);
  if (!match) return '';
  const year = match[1].slice(-2);
  const month = Number(match[2]);
  return Number.isInteger(month) && month >= 1 && month <= 12 ? `${year}-${month}` : '';
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

function formatChangeDetail(previousRow, next) {
  const beforeStart = previousRow?.[2] || '';
  const beforeEnd = previousRow?.[3] || '';
  const beforeMemo = previousRow?.[7] || '';
  return `시작 ${beforeStart || '-'} -> ${next.startTime || '-'}, 종료 ${beforeEnd || '-'} -> ${next.endTime || '-'}, 메모 ${beforeMemo || '-'} -> ${next.memo || '-'}`;
}

function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function normalizeDate(value) {
  return String(value || '').replace(/\s+/g, '');
}

function isDateText(value) {
  const text = String(value || '').trim();
  return Boolean(text && !text.includes('날짜') && parseMonthDay(text));
}

function normalizeAfterSchoolOverride(value) {
  if (value === true) return 'O';
  if (value === false) return 'X';
  const text = String(value ?? '').trim().toUpperCase();
  if (['O', 'Y', 'TRUE', '1', 'YES'].includes(text)) return 'O';
  if (['X', 'N', 'FALSE', '0', 'NO'].includes(text)) return 'X';
  return '';
}

function normalizeDeductionMinutes(value) {
  const text = String(value ?? '').replace(/[^\d.-]/g, '');
  if (!text) return DEFAULT_AFTER_SCHOOL_DEDUCTION_MINUTES;
  const minutes = Number(text);
  if (!Number.isFinite(minutes)) return DEFAULT_AFTER_SCHOOL_DEDUCTION_MINUTES;
  return Math.max(0, Math.min(240, Math.round(minutes)));
}

function isAfterSchoolApplied(override, dateText, physicalSheetName, defaultSettings) {
  const normalizedOverride = normalizeAfterSchoolOverride(override);
  if (normalizedOverride === 'O') return true;
  if (normalizedOverride === 'X') return false;
  const weekdayIndex = getWeekdayIndex(dateText, physicalSheetName);
  return weekdayIndex >= 0 ? Boolean(defaultSettings[weekdayIndex]) : false;
}

function getWeekdayIndex(dateText, physicalSheetName) {
  const parts = parseMonthDay(dateText);
  const monthSheet = parseMonthSheetName(physicalSheetName);
  if (!parts || !monthSheet) return -1;
  const year = monthSheet.year;
  const date = new Date(year, parts.month - 1, parts.day);
  if (Number.isNaN(date.getTime())) return -1;
  return (date.getDay() + 6) % 7;
}

function parseMonthDay(value) {
  const text = String(value || '').trim();
  const korean = text.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/);
  if (korean) return validMonthDay(Number(korean[1]), Number(korean[2]));
  const numeric = text.match(/(?:\d{2,4}\D+)?(\d{1,2})\D+(\d{1,2})/);
  if (numeric) return validMonthDay(Number(numeric[1]), Number(numeric[2]));
  return null;
}

function validMonthDay(month, day) {
  return month >= 1 && month <= 12 && day >= 1 && day <= 31 ? { month, day } : null;
}

function parseMonthSheetName(value) {
  const match = String(value || '').match(/(?:^|__)(\d{2}|\d{4})-(\d{1,2})$/);
  if (!match) return null;
  const shortYear = Number(match[1].slice(-2));
  const fullYear = match[1].length === 4 ? Number(match[1]) : 2000 + shortYear;
  const month = Number(match[2]);
  return Number.isInteger(month) && month >= 1 && month <= 12 ? { year: fullYear, month } : null;
}

function calculateOvertimeFields(startTime, endTime, afterSchoolApplied, deductionMinutes) {
  const startMinutes = parseTimeToMinutes(startTime);
  const endMinutes = parseTimeToMinutes(endTime);
  const morningMinutes = startMinutes === null ? 0 : Math.max(0, STANDARD_START_MINUTES - startMinutes);
  const rawAfternoonMinutes = endMinutes === null ? 0 : Math.max(0, endMinutes - STANDARD_END_MINUTES);
  const afternoonMinutes = Math.max(0, rawAfternoonMinutes - (afterSchoolApplied ? deductionMinutes : 0));
  return {
    morningExtra: formatDuration(morningMinutes),
    afternoonExtra: formatDuration(afternoonMinutes),
    dailyTotal: formatDuration(morningMinutes + afternoonMinutes),
  };
}

function parseTimeToMinutes(value) {
  const text = String(value || '').trim();
  const match = text.match(/(\d{1,2}):(\d{2})/);
  if (!match) return null;
  let hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (!Number.isInteger(hours) || !Number.isInteger(minutes)) return null;
  if (/오후|PM/i.test(text) && hours < 12) hours += 12;
  if (/오전|AM/i.test(text) && hours === 12) hours = 0;
  return hours * 60 + minutes;
}

function formatDuration(totalMinutes) {
  const minutes = Math.max(0, Math.round(totalMinutes));
  return `${Math.floor(minutes / 60)}:${String(minutes % 60).padStart(2, '0')}`;
}

function nowText() {
  return new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });
}

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

function requireField(value, message) {
  if (value === undefined || value === null || value === '') {
    const error = new Error(message);
    error.statusCode = 400;
    error.userMessage = message;
    throw error;
  }
}
