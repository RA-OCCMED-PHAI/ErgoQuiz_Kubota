const SHEET_NAME = 'ErgoGameshow';

const HEADERS = [
  'server_submitted_at',
  'client_submitted_at',

  'pain_lower_back_selected',
  'pain_lower_back_score',

  'pain_shoulder_selected',
  'pain_shoulder_score',

  'pain_knee_selected',
  'pain_knee_score',

  'q1_answer',
  'q1_correct',

  'q2_answer',
  'q2_correct',

  'q3_answer',
  'q3_correct',

  'q4_answer',
  'q4_correct',

  'q5_answer',
  'q5_correct',

  'q6_answer',
  'q6_correct',

  'q7_answer',
  'q7_correct',

  'q8_answer',
  'q8_correct',

  'q9_answer',
  'q9_correct',

  'q10_answer',
  'q10_correct',

  'correct_count',
  'score_percent',

  'page_url',
  'user_agent'
];


/**
 * Run ฟังก์ชันนี้ 1 ครั้ง
 * เพื่อสร้าง Sheet และ Header
 */
function setupSheet() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet =
      ss.insertSheet(SHEET_NAME);
  }

  sheet
    .getRange(
      1,
      1,
      1,
      HEADERS.length
    )
    .setValues([HEADERS]);

  sheet.setFrozenRows(1);

  sheet.autoResizeColumns(
    1,
    HEADERS.length
  );

  Logger.log(
    'Setup complete: ' +
    SHEET_NAME
  );
}


/**
 * รับข้อมูลจาก GitHub Pages
 */
function doPost(e) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(10000);

  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    let sheet =
      ss.getSheetByName(SHEET_NAME);


    if (!sheet) {

      sheet =
        ss.insertSheet(SHEET_NAME);

      sheet
        .getRange(
          1,
          1,
          1,
          HEADERS.length
        )
        .setValues([HEADERS]);

      sheet.setFrozenRows(1);
    }


    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {

      return jsonResponse_({
        ok: false,
        message:
          'No POST body received.'
      });
    }


    let data = {};


    try {

      data =
        JSON.parse(
          e.postData.contents
        );

    } catch (err) {

      return jsonResponse_({
        ok: false,
        message:
          'Invalid JSON body.',
        detail:
          String(err)
      });
    }


    const serverTimestamp =
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone()
          || 'Asia/Bangkok',
        'yyyy-MM-dd HH:mm:ss'
      );


    const rowObject =
      Object.assign(
        {},
        data,
        {
          server_submitted_at:
            serverTimestamp
        }
      );


    const row =
      HEADERS.map(header => {

        const value =
          rowObject[header];


        if (
          value === undefined ||
          value === null
        ) {
          return '';
        }


        // ป้องกัน formula injection
        if (
          typeof value === 'string' &&
          /^[=+\-@]/.test(value)
        ) {

          return "'" + value;
        }


        return value;
      });


    sheet.appendRow(row);


    return jsonResponse_({
      ok: true,
      message:
        'Saved',
      row:
        sheet.getLastRow()
    });


  } catch (error) {

    return jsonResponse_({
      ok: false,
      message:
        String(error)
    });


  } finally {

    lock.releaseLock();

  }
}


/**
 * ใช้ทดสอบว่า Web App เปิดได้หรือไม่
 */
function doGet() {

  return jsonResponse_({
    ok: true,
    app:
      'ErgoGameshow API',
    message:
      'Web App is running.'
  });
}


/**
 * JSON response
 */
function jsonResponse_(obj) {

  return ContentService
    .createTextOutput(
      JSON.stringify(obj)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
