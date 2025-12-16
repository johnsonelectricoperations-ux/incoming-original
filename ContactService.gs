/**
 * ContactService.gs
 * 비상연락망 정보 관리 서비스
 */

/**
 * 연락처 정보 저장/수정
 */
function saveContactInfo(token, contactData) {
  const startTime = Date.now();

  try {
    // 인증 확인
    const user = TokenService.validateToken(token);
    if (!user) {
      return {
        success: false,
        message: '인증되지 않은 사용자입니다.'
      };
    }

    // 일반 사용자는 자신의 업체 정보만 수정 가능
    if (user.role !== '관리자' && user.role !== 'JEO') {
      if (contactData.companyName !== user.companyName) {
        return {
          success: false,
          message: '다른 업체의 정보는 수정할 수 없습니다.'
        };
      }
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('ContactInfo');

    // ContactInfo 시트가 없으면 생성
    if (!sheet) {
      sheet = ss.insertSheet('ContactInfo');
      // 헤더 추가
      sheet.appendRow(['업체명', '품질담당자', '연락처', '이메일', '등록일', '수정일', '등록자']);
      sheet.getRange(1, 1, 1, 7).setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold');
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // 기존 데이터 찾기
    let existingRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === contactData.companyName) {
        existingRowIndex = i + 1;
        break;
      }
    }

    const now = new Date();
    const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    if (existingRowIndex > 0) {
      // 기존 데이터 수정
      sheet.getRange(existingRowIndex, 2).setValue(contactData.contactName || '');
      sheet.getRange(existingRowIndex, 3).setValue(contactData.phone || '');
      sheet.getRange(existingRowIndex, 4).setValue(contactData.email || '');
      sheet.getRange(existingRowIndex, 6).setValue(nowStr); // 수정일
    } else {
      // 새 데이터 추가
      sheet.appendRow([
        contactData.companyName,
        contactData.contactName || '',
        contactData.phone || '',
        contactData.email || '',
        nowStr, // 등록일
        nowStr, // 수정일
        user.name
      ]);
    }

    const totalTime = Date.now() - startTime;
    Logger.log(`연락처 정보 저장 완료 (${totalTime}ms): ${contactData.companyName}`);

    return {
      success: true,
      message: '연락처 정보가 저장되었습니다.',
      executionTime: totalTime
    };

  } catch (error) {
    const totalTime = Date.now() - startTime;
    Logger.log(`연락처 정보 저장 오류 (${totalTime}ms): ${error.toString()}`);
    return {
      success: false,
      message: '연락처 정보 저장 중 오류가 발생했습니다: ' + error.message,
      executionTime: totalTime
    };
  }
}

/**
 * 연락처 정보 조회
 */
function getContactInfo(token, companyName) {
  const startTime = Date.now();

  try {
    // 인증 확인
    const user = TokenService.validateToken(token);
    if (!user) {
      return {
        success: false,
        message: '인증되지 않은 사용자입니다.'
      };
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('ContactInfo');

    if (!sheet) {
      return {
        success: true,
        data: [],
        executionTime: Date.now() - startTime
      };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 일반 사용자는 자신의 업체만, 관리자/JEO는 전체 조회
      if (user.role === '관리자' || user.role === 'JEO') {
        // 전체 조회
        if (!companyName || row[0] === companyName) {
          result.push({
            companyName: row[0],
            contactName: row[1],
            phone: row[2],
            email: row[3],
            createdAt: row[4],
            updatedAt: row[5],
            createdBy: row[6]
          });
        }
      } else {
        // 자신의 업체만 조회
        if (row[0] === user.companyName) {
          result.push({
            companyName: row[0],
            contactName: row[1],
            phone: row[2],
            email: row[3],
            createdAt: row[4],
            updatedAt: row[5],
            createdBy: row[6]
          });
        }
      }
    }

    const totalTime = Date.now() - startTime;
    Logger.log(`연락처 정보 조회 완료 (${totalTime}ms): ${result.length}건`);

    return {
      success: true,
      data: result,
      executionTime: totalTime
    };

  } catch (error) {
    const totalTime = Date.now() - startTime;
    Logger.log(`연락처 정보 조회 오류 (${totalTime}ms): ${error.toString()}`);
    return {
      success: false,
      message: '연락처 정보 조회 중 오류가 발생했습니다: ' + error.message,
      executionTime: totalTime
    };
  }
}

/**
 * 모든 연락처 정보 조회 (관리자/JEO 전용)
 */
function getAllContactInfo(token) {
  const startTime = Date.now();

  try {
    // 인증 확인
    const user = TokenService.validateToken(token);
    if (!user) {
      return {
        success: false,
        message: '인증되지 않은 사용자입니다.'
      };
    }

    // 권한 확인
    if (user.role !== '관리자' && user.role !== 'JEO') {
      return {
        success: false,
        message: '권한이 없습니다.'
      };
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('ContactInfo');

    if (!sheet) {
      return {
        success: true,
        data: [],
        executionTime: Date.now() - startTime
      };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      result.push({
        companyName: row[0],
        contactName: row[1],
        phone: row[2],
        email: row[3],
        createdAt: row[4],
        updatedAt: row[5],
        createdBy: row[6]
      });
    }

    const totalTime = Date.now() - startTime;
    Logger.log(`전체 연락처 정보 조회 완료 (${totalTime}ms): ${result.length}건`);

    return {
      success: true,
      data: result,
      executionTime: totalTime
    };

  } catch (error) {
    const totalTime = Date.now() - startTime;
    Logger.log(`전체 연락처 정보 조회 오류 (${totalTime}ms): ${error.toString()}`);
    return {
      success: false,
      message: '연락처 정보 조회 중 오류가 발생했습니다: ' + error.message,
      executionTime: totalTime
    };
  }
}
