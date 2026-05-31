const SHEET_NAME  = 'raw';
const ADMIN_EMAIL = 'discplus2010@naver.com';
const HEADER = [
    '문의번호', '문의일시',
    '성함', '소속/직함', '이메일', '전화번호',
    '문의 내용'
];

function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);

        // honeypot 차단
        if (data._honey) {
            return ContentService
                .createTextOutput(JSON.stringify({ result: 'blocked' }))
                .setMimeType(ContentService.MimeType.JSON);
        }

        const name  = String(data.name || '').trim();
        const org   = String(data.org || '').trim();
        const email = String(data.email || '').trim();
        const tel   = String(data.tel || '').trim();
        const memo  = String(data.memo || '').trim();

        const emailOk = /^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$/.test(email);
        const telOk = /^[0-9-]+$/.test(tel);

        if (!name || !memo) {
            return ContentService
                .createTextOutput(JSON.stringify({ result: 'invalid_required' }))
                .setMimeType(ContentService.MimeType.JSON);
        }

        if (!emailOk || !telOk) {
            return ContentService
                .createTextOutput(JSON.stringify({ result: 'invalid_contact' }))
                .setMimeType(ContentService.MimeType.JSON);
        }

        const ss    = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEET_NAME)
                     || ss.insertSheet(SHEET_NAME);

        ensureHeader_(sheet);

        // 문의번호 생성: CPQ-YYYYMMDD-NNN
        const today   = Utilities.formatDate(
                            new Date(), 'Asia/Seoul', 'yyyyMMdd');
        const lastRow = sheet.getLastRow();
        const seq     = String(lastRow).padStart(3, '0');
        const uid     = 'CPQ-' + today + '-' + seq;

        // 문의일시
        const timestamp = Utilities.formatDate(
                            new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

        // 시트에 기록
        sheet.appendRow([
            uid, timestamp,
            name,
            org,
            email,
            tel,
            memo
        ]);

        // 관리자 이메일 발송
        const subject = `📌[코칭플러스 문의] ${name}`;
        const body = [
            `■ 문의일시: ${timestamp}`,
            ``,
            `■ 성함: ${name}`,
            `■ 소속: ${org || '—'}`,
            `■ 이메일: ${email}`,
            `■ 전화번호: ${tel}`,
            ``,
            `■ 문의 내용:`,
            memo,
            ``,
            `─────────────────────`,
            `스프레드시트에서 확인: ${ss.getUrl()}`
        ].filter(Boolean).join('\n');

        MailApp.sendEmail(ADMIN_EMAIL, subject, body);

        return ContentService
            .createTextOutput(JSON.stringify({ result: 'ok', uid }))
            .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
        return ContentService
            .createTextOutput(JSON.stringify({ result: 'error', msg: err.message }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

function ensureHeader_(sheet) {
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(HEADER);
        sheet.setFrozenRows(1);
        return;
    }

    const firstRow = sheet.getRange(1, 1, 1, HEADER.length).getValues()[0];
    const hasInquiryHeader = firstRow[0] === HEADER[0] && firstRow[1] === HEADER[1];

    if (!hasInquiryHeader) {
        sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), HEADER.length)).clearContent();
        sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
        sheet.setFrozenRows(1);
    }

    const extraColumns = sheet.getMaxColumns() - HEADER.length;
    if (extraColumns > 0) {
        sheet.deleteColumns(HEADER.length + 1, extraColumns);
    }
}
