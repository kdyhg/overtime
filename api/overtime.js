const { google } = require('googleapis');

module.exports = async (req, res) => {
  // POST 요청만 허용
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { action, sheetName, dateText, startTime, endTime, settings, targetName } = req.body;

    // 1. 서비스 계정(봇) 인증
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.GOOGLE_CLIENT_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'), // 줄바꿈 문자 처리 필수
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;

    // (내부 헬퍼 함수) 시트 데이터 가져오기
    const getSheetData = async (name) => {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${name}!A:L`,
        valueRenderOption: 'FORMATTED_VALUE'
      });
      const rows = response.data.values || [];
      
      const summary = {
        totalHours: rows[1]?.[9] || "",
        approvedHours: rows[2]?.[9] || "",
        hourlyRate: rows[3]?.[9] || "",
        totalPay: rows[4]?.[9] || ""
      };

      const dates = [];
      const records = {};
      const afterSchool = [];

      for (let i = 1; i < rows.length; i++) {
        let displayDate = rows[i][0] || "";
        if (!displayDate || displayDate.includes("날짜") || displayDate.trim() === "") continue;
        
        dates.push({ dateText: displayDate });
        records[displayDate.replace(/\s+/g, '')] = {
          startTime: rows[i][2] || "",
          endTime: rows[i][3] || "",
          morningExtra: rows[i][4] || "0:00",
          afternoonExtra: rows[i][5] || "0:00",
          dailyTotal: rows[i][6] || "0:00"
        };
      }

      for (let i = 1; i < 8 && i < rows.length; i++) {
        afterSchool.push({
          day: rows[i][10] || "요일",
          isSelected: rows[i][11] === "O"
        });
      }
      return { summary, dates, afterSchool, records };
    };

    // 2. 클라이언트 요청(action)에 따른 분기 처리
    if (action === 'getInitialAppData') {
      const meta = await sheets.spreadsheets.get({ spreadsheetId });
      const sheetNames = meta.data.sheets.map(s => s.properties.title);
      
      const now = new Date();
      const yy = now.getFullYear().toString().slice(-2);
      const m = now.getMonth() + 1;
      const currentMonthName = `${yy}-${m}`;

      let defaultSheetName = sheetNames.includes(currentMonthName) ? currentMonthName : sheetNames[0];
      const filteredSheetNames = sheetNames.filter(name => name !== "양식" && name !== "통합DB" && name !== "설정");

      const data = await getSheetData(defaultSheetName);
      return res.status(200).json({
        sheetNames: filteredSheetNames,
        currentSheet: defaultSheetName,
        ...data
      });
    }
    
    else if (action === 'getSheetData') {
      const data = await getSheetData(sheetName);
      return res.status(200).json(data);
    }
    
    else if (action === 'saveOvertime') {
      // 날짜에 해당하는 행(Row) 찾기
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!A:A`,
        valueRenderOption: 'FORMATTED_VALUE'
      });
      const rows = response.data.values || [];
      const normalizedTarget = dateText.replace(/\s+/g, '');
      
      let targetRow = -1;
      for (let i = 1; i < rows.length; i++) {
        let sheetDate = (rows[i][0] || "").replace(/\s+/g, '');
        if (sheetDate === normalizedTarget) {
          targetRow = i + 1; // 구글 시트는 1부터 시작
          break;
        }
      }

      if (targetRow !== -1) {
        // C열(시작시간)과 D열(종료시간) 업데이트
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheetName}!C${targetRow}:D${targetRow}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[startTime, endTime]] }
        });
        const updatedData = await getSheetData(sheetName);
        return res.status(200).json({ success: true, ...updatedData });
      } else {
        return res.status(200).json({ success: false, message: "날짜를 찾을 수 없습니다." });
      }
    }
    
    else if (action === 'saveAfterSchoolSettings') {
      const writeData = settings.map(s => [s ? "O" : ""]);
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!L2:L8`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: writeData }
      });
      const updatedData = await getSheetData(sheetName);
      return res.status(200).json({ success: true, ...updatedData });
    }
    
    else if (action === 'createNewMonthSheet') {
       // 양식 시트 복사하여 새 월 시트 생성
       const meta = await sheets.spreadsheets.get({ spreadsheetId });
       const templateSheet = meta.data.sheets.find(s => s.properties.title === '양식');
       if (!templateSheet) return res.status(400).json({ success: false, message: "양식 시트 없음" });

       const request = {
         duplicateSheet: {
           sourceSheetId: templateSheet.properties.sheetId,
           insertSheetIndex: 0,
           newSheetName: targetName
         }
       };

       await sheets.spreadsheets.batchUpdate({
         spreadsheetId,
         requestBody: { requests: [request] }
       });

       // 수식 세팅
       const monthPart = targetName.split("-")[1];
       if (monthPart) {
         await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `${targetName}!A2`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [[`=FILTER('통합DB'!A:A, MONTH('통합DB'!A:A)=${monthPart})`]] }
         });
       }
       return res.status(200).json({ success: true, newSheetName: targetName });
    }

    res.status(400).json({ error: "알 수 없는 요청입니다." });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
};