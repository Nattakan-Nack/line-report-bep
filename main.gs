function sendDailyLineMessageFromSheet() {
  const ACCESS_TOKEN = 'lWtakfUshjpo7SyZ8encvg1O8dwfC1REAflmmmSoxHkXqrhC6frTTrduRVY6DCIWzFCHxJBxasa8BHTPQEIVzbKzhEg450BqZMHo4WdhSkcpXbi/3gIk3jV1H5ggFzxjApwox6g59nLz4tiLhs0U+AdB04t89/1O/w1cDnyilFU='; // ← เปลี่ยนเป็น Token จริงของแน็ก
  const TO_GROUP_ID = 'C8891f4dc35552f0b83c228abfcfe232c';      // ← เปลี่ยนเป็น Group ID จริงของแน็ก

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data"); // ← ใส่ชื่อชีทให้ตรง
  const data = sheet.getDataRange().getValues();

  const today = new Date();
  const todayDay = today.getDate().toString(); // เช่น "18"

  // คอลัมน์ H คือ index 7 (นับจาก 0)
  const todayRow = data.find(row => String(row[7]) === todayDay);

  if (!todayRow) {
    throw new Error("❗ ไม่พบข้อมูลของวันที่ " + todayDay + " ในชีท กรุณาตรวจสอบอีกครั้ง");
  }

  const [
    , , , , , , , // ข้าม index 0-6 (A–G)
    day, caseCount, freeDates, status, totalSales,
    remaining, percentDone, toSellMore, lipo, gyne, suggestionText, percentOfTimePassed
  ] = todayRow;

  const message = 
`🧾 รายงานยอด Break Even ประจำวันที่ ${day} พ.ค.

📊 ยอดขายสะสม: ${formatNumber(totalSales)} บาท (${formatPercent(percentDone)} ของเป้าหมาย)
🎯 เป้าหมาย Break Even: 5,000,000 บาท
📉 ยอดที่ยังคงเหลือ: ${formatNumber(remaining)} บาท

📌 เพื่อเผื่อส่วนลดที่อาจมอบให้ลูกค้าและสร้างกำไรหลัง BEP
→ ควรมียอดขายเพิ่มเติมอีก: ${formatNumber(toSellMore)} บาท
(คิดจากยอดคงเหลือ + 57% เป็น buffer เพื่อความปลอดภัย)

แบ่งเป็น:
- 💉 Lipo (75%): ${formatNumber(lipo)} บาท
- 👨‍⚕️ Gyneco (25%): ${formatNumber(gyne)} บาท

📅 วันที่ว่างในคลินิก: ${freeDates}
✅ สถานะยอดขายวันนี้: ${status}

🔆 ยอดขายขณะนี้ ${formatPercent(percentDone)} เทียบกับเวลาเดือนนี้ ${formatPercent(percentOfTimePassed)}

📊 สถานะรวม: ${status}
${suggestionText}
`;

  const payload = {
    to: TO_GROUP_ID,
    messages: [{
      type: "text",
      text: message
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + ACCESS_TOKEN
    },
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", options);
}

function formatNumber(n) {
  return Number(n).toLocaleString('en-US');
}

function formatPercent(n) {
  return (Number(n) * 100).toFixed(2) + "%";
}

add main.gs (LINE auto-report script)
