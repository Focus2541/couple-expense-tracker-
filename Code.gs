/**
 * บันทึกรายจ่ายคู่รัก: แยกรายจ่ายของคนที่ 1 / คนที่ 2 / ร่วมกัน ตามหมวดหมู่
 * + แจ้งเตือน LINE รายวัน/รายเดือน ส่งแยกเข้าแชทส่วนตัวของแต่ละคน (ไม่ใช่เข้ากลุ่ม)
 * + พิมพ์ข้อความหรือส่งรูปสลิปในแชท "กลุ่ม" ให้ AI ("น้องออม") แยกหมวด/จำนวนเงิน/คนจ่ายให้อัตโนมัติ
 * -------------------------------------------------
 * วิธีติดตั้ง: ดูไฟล์ "คู่มือติดตั้ง.md" ที่แนบมาด้วย
 *
 * ต้องตั้งค่า Script Properties (Project Settings > Script Properties) ดังนี้:
 *   LINE_CHANNEL_ACCESS_TOKEN  = channel access token (long-lived) จาก LINE Developers Console
 *   CLAUDE_API_KEY             = API key จาก console.anthropic.com (สำหรับแยกวิเคราะห์ข้อความ/รูปสลิป)
 *   PARTNER1_LINE_ID           = userId ของคนที่ 1 (ทักแชทส่วนตัวกับ OA 1 ครั้ง แล้วดู ID จาก LineIDs sheet)
 *                                 — ใช้ทั้งระบุตัวตนตอนพิมพ์ในกลุ่ม และเป็นที่ส่งแจ้งเตือนส่วนตัว
 *   PARTNER2_LINE_ID           = userId ของคนที่ 2 (เหมือนกัน)
 *   LINE_TARGET_ID             = (ไม่บังคับ) groupId ที่จะส่งแจ้งเตือนเข้ากลุ่มด้วย ใช้เป็น fallback
 *                                 เฉพาะตอนที่ยังไม่ได้ตั้ง PARTNER1_LINE_ID/PARTNER2_LINE_ID เลย
 */

const SHEET_CATEGORIES = 'Categories';
const SHEET_EXPENSES = 'Expenses';
const SHEET_LINEIDS = 'LineIDs';

const WHO_SHARED = 'shared';

// ---------- เขตเวลา (รองรับตั้งเขตเวลาชั่วคราวตอนไปต่างประเทศ) ----------

function getActiveTimezone_() {
  const tz = PropertiesService.getScriptProperties().getProperty('TRAVEL_TIMEZONE');
  return tz || Session.getScriptTimeZone() || 'Asia/Bangkok';
}

// แปลง ISO string หรือ Date ให้เป็นวันที่/เดือนตามเขตเวลาที่ใช้งานอยู่ตอนนี้ (ไม่ใช่ค่า UTC ดิบ)
function localDatePrefix_(isoOrDate, pattern) {
  const d = (isoOrDate instanceof Date) ? isoOrDate : new Date(isoOrDate);
  return Utilities.formatDate(d, getActiveTimezone_(), pattern);
}

// สร้างลิสต์เดือน "YYYY-MM" สำหรับดรอปดาวน์ใน Dashboard (ย้อนหลัง/ล่วงหน้าตามที่ระบุ)
function generateMonthOptions_(monthsBack, monthsForward) {
  const options = [];
  const base = new Date();
  for (let i = -monthsBack; i <= monthsForward; i++) {
    const d = new Date(base.getFullYear(), base.getMonth() + i, 1);
    options.push(Utilities.formatDate(d, getActiveTimezone_(), 'yyyy-MM'));
  }
  return options;
}

// ---------- แปลงสกุลเงินต่างประเทศเป็นบาท ----------

function convertToTHB_(amount, currency) {
  const code = (currency || 'THB').toUpperCase().trim();
  if (!code || code === 'THB') return amount;
  try {
    const res = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/' + encodeURIComponent(code), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    const data = JSON.parse(res.getContentText());
    if (data.result !== 'success' || !data.rates || !data.rates.THB) return null;
    return amount * data.rates.THB;
  } catch (e) {
    Logger.log('convertToTHB_ fail: ' + e);
    return null;
  }
}

// ---------- ตั้งค่าเริ่มต้น ----------

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let categories = ss.getSheetByName(SHEET_CATEGORIES);
  if (!categories) {
    categories = ss.insertSheet(SHEET_CATEGORIES);
    categories.appendRow(['Category']);
    categories.appendRow(['ค่ากิน']);
    categories.appendRow(['ค่าน้ำ']);
    categories.appendRow(['ค่าไฟ']);
    categories.appendRow(['ค่าเที่ยว']);
    categories.appendRow(['ค่าที่พัก']);
  }

  let expenses = ss.getSheetByName(SHEET_EXPENSES);
  if (!expenses) {
    expenses = ss.insertSheet(SHEET_EXPENSES);
    expenses.appendRow(['ID', 'Date', 'Who', 'Category', 'Amount', 'Note', 'LocalDate', 'RecordedAt']);
  }
  // เผื่อชีตเก่าที่สร้างไว้ก่อนมีคอลัมน์ LocalDate ให้เติมหัวตารางให้ครบ (ไม่กระทบข้อมูลเดิม)
  if (expenses.getRange(1, 7).getValue() !== 'LocalDate') {
    expenses.getRange(1, 7).setValue('LocalDate');
  }
  // เผื่อชีตเก่าที่สร้างไว้ก่อนมีคอลัมน์ RecordedAt (เวลาพิมพ์จริง แยกจาก Date ที่ย้อนหลังได้)
  if (expenses.getRange(1, 8).getValue() !== 'RecordedAt') {
    expenses.getRange(1, 8).setValue('RecordedAt');
  }
  // บังคับคอลัมน์ LocalDate ให้เป็นข้อความล้วนเสมอ กัน Sheets แปลงเป็น Date object เองอัตโนมัติ
  expenses.getRange(1, 7, Math.max(expenses.getMaxRows(), 1), 1).setNumberFormat('@');
  expenses.getRange(1, 8, Math.max(expenses.getMaxRows(), 1), 1).setNumberFormat('@');

  let lineIds = ss.getSheetByName(SHEET_LINEIDS);
  if (!lineIds) {
    lineIds = ss.insertSheet(SHEET_LINEIDS);
    lineIds.appendRow(['Type', 'ID', 'DisplayName', 'CapturedAt']);
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PARTNER1')) props.setProperty('PARTNER1', 'คนที่ 1');
  if (!props.getProperty('PARTNER2')) props.setProperty('PARTNER2', 'คนที่ 2');

  return { ss, categories, expenses, lineIds };
}

// ---------- Web App entry points ----------

function doGet(e) {
  ensureSheets_();
  const p = e.parameter || {};
  const action = p.action || 'data';
  let result;

  // ล็อกหลังบ้าน: ทุกคำสั่งผ่าน Web app URL ต้องมีรหัส ADMIN_KEY ที่โฟกัสตั้งไว้เท่านั้น
  const adminKey = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!adminKey) {
    result = { error: 'ยังไม่ได้ตั้ง ADMIN_KEY ใน Script Properties ระบบหลังบ้านถูกล็อกไว้' };
  } else if (p.key !== adminKey) {
    result = { error: 'รหัสไม่ถูกต้อง' };
  } else if (action === 'data') {
    result = getAllData_();
  } else if (action === 'addExpense') {
    addExpense_(p.who, p.category, p.amount, p.note, p.date);
    result = getAllData_();
  } else if (action === 'deleteExpense') {
    deleteExpense_(p.id);
    result = getAllData_();
  } else if (action === 'addCategory') {
    addCategory_(p.category);
    result = getAllData_();
  } else if (action === 'deleteCategory') {
    deleteCategory_(p.category);
    result = getAllData_();
  } else if (action === 'setPartners') {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('PARTNER1', p.partner1 || 'คนที่ 1');
    props.setProperty('PARTNER2', p.partner2 || 'คนที่ 2');
    syncDashboardNames_();
    result = getAllData_();
  } else if (action === 'testSend') {
    dailyLineSummary();
    result = { ok: true };
  } else if (action === 'testSendMonthly') {
    monthlyLineSummary();
    result = { ok: true };
  } else {
    result = { error: 'unknown action' };
  }

  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOut_(result);
}

// doPost ใช้สำหรับ LINE webhook เท่านั้น (LINE server ยิงมาโดยตรง ไม่ผ่านเบราว์เซอร์ ไม่ติดปัญหา CORS)
function doPost(e) {
  ensureSheets_();
  const raw = e.postData ? e.postData.contents : '{}';
  let body;
  try {
    body = JSON.parse(raw);
  } catch (err) {
    return jsonOut_({ error: 'invalid json' });
  }

  if (body.events) {
    handleLineWebhook_(body.events);
  }
  return ContentService.createTextOutput('OK');
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- ข้อมูล ----------

function getAllData_() {
  const { categories, expenses } = ensureSheets_();
  const props = PropertiesService.getScriptProperties();

  const categoryRows = categories.getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => r[0]);

  const expenseRows = expenses.getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => {
      const isoDate = r[1] instanceof Date ? r[1].toISOString() : r[1];
      let localDate = r[6];
      if (localDate instanceof Date) {
        // Google Sheets แปลงข้อความ "YYYY-MM-DD" เป็น Date object ให้เองอัตโนมัติ ต้องแปลงกลับเป็นข้อความ
        localDate = Utilities.formatDate(localDate, getActiveTimezone_(), 'yyyy-MM-dd');
      } else if (!localDate) {
        // แถวเก่าก่อนมีคอลัมน์ LocalDate: เดาโดยสมมติว่าบันทึกด้วยเขตเวลาไทย (ค่าเริ่มต้นเดิมของระบบ)
        try {
          localDate = Utilities.formatDate(new Date(isoDate), 'Asia/Bangkok', 'yyyy-MM-dd');
        } catch (e) {
          localDate = String(isoDate).slice(0, 10);
        }
      } else {
        localDate = String(localDate);
      }
      let recordedAt = r[7];
      if (recordedAt instanceof Date) {
        recordedAt = recordedAt.toISOString();
      } else if (!recordedAt) {
        // แถวเก่าก่อนมีคอลัมน์ RecordedAt: ใช้ Date เดิมแทน (เผื่อรายการนั้นเคยย้อนหลัง จะไม่แม่นเป๊ะ แต่ดีกว่าไม่มีเลย)
        recordedAt = isoDate;
      } else {
        recordedAt = String(recordedAt);
      }
      return {
        id: r[0],
        date: isoDate,
        who: r[2],
        category: r[3],
        amount: Number(r[4]) || 0,
        note: r[5] || '',
        localDate: localDate,
        recordedAt: recordedAt
      };
    });

  return {
    categories: categoryRows,
    expenses: expenseRows,
    partners: [props.getProperty('PARTNER1'), props.getProperty('PARTNER2')]
  };
}

function addExpense_(who, category, amount, note, date, localDate, recordedAt) {
  const { expenses } = ensureSheets_();
  const id = Date.now();
  const isoDate = date || new Date().toISOString();
  const ld = localDate || localDatePrefix_(isoDate, 'yyyy-MM-dd');
  const ra = recordedAt || new Date().toISOString();
  expenses.appendRow([id, isoDate, who, category, Number(amount) || 0, note || '', ld, ra]);
}

function deleteExpense_(id) {
  const { expenses } = ensureSheets_();
  const data = expenses.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      expenses.deleteRow(i + 1);
      break;
    }
  }
}

function updateExpenseCategory_(id, category) {
  const { expenses } = ensureSheets_();
  const data = expenses.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      expenses.getRange(i + 1, 4).setValue(category); // column D = Category
      return true;
    }
  }
  return false;
}

function updateExpenseAmount_(id, amount) {
  const { expenses } = ensureSheets_();
  const data = expenses.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      expenses.getRange(i + 1, 5).setValue(Number(amount) || 0); // column E = Amount
      return true;
    }
  }
  return false;
}

function updateExpenseWho_(id, who) {
  const { expenses } = ensureSheets_();
  const data = expenses.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      expenses.getRange(i + 1, 3).setValue(who); // column C = Who
      return true;
    }
  }
  return false;
}

function renamePartnerInExpenses_(oldName, newName) {
  const { expenses } = ensureSheets_();
  const data = expenses.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === oldName) { // column C = Who
      expenses.getRange(i + 1, 3).setValue(newName);
    }
  }
}

function addCategory_(category) {
  if (!category) return;
  const { categories } = ensureSheets_();
  const existing = categories.getDataRange().getValues().slice(1).map(r => r[0]);
  if (existing.indexOf(category) === -1) categories.appendRow([category]);
}

function deleteCategory_(category) {
  const { categories } = ensureSheets_();
  const data = categories.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === category) {
      categories.deleteRow(i + 1);
      break;
    }
  }
}

// ---------- LINE Webhook: หา userId/groupId + แยกวิเคราะห์ข้อความ/รูปสลิปด้วย AI ----------

function handleLineWebhook_(events) {
  const { lineIds } = ensureSheets_();
  const props = PropertiesService.getScriptProperties();
  const setupDone = !!props.getProperty('CLAUDE_API_KEY');

  events.forEach(ev => {
    const source = ev.source || {};
    const srcType = source.groupId ? 'group' : (source.roomId ? 'room' : 'user');
    const id = source.groupId || source.roomId || source.userId;

    if (id) {
      const already = lineIds.getDataRange().getValues().slice(1).some(r => String(r[1]) === String(id));
      if (!already) lineIds.appendRow([srcType, id, '', new Date()]);
    }

    if (!setupDone) {
      if (ev.replyToken && id) {
        const propName = srcType === 'group' ? 'LINE_TARGET_ID (ไม่บังคับ ใช้ถ้าอยากแจ้งเตือนเข้ากลุ่มด้วย)' : 'PARTNER1_LINE_ID หรือ PARTNER2_LINE_ID (แล้วแต่ว่าเป็น ID ของใคร)';
        replyLine_(ev.replyToken, `รับทราบครับ ID ของ${srcType === 'group' ? 'กลุ่มนี้' : 'แชทนี้'}คือ:\n${id}\n\nนำไปใส่ใน Script Properties ชื่อ ${propName} ได้เลยครับ (และอย่าลืมใส่ CLAUDE_API_KEY ด้วยถ้ายังไม่ได้ใส่ ไม่งั้นบอทจะยังไม่เริ่มแยกวิเคราะห์ข้อความ)`);
      }
      return;
    }

    // รับเฉพาะข้อความจากโฟกัสกับอ๋อมเท่านั้น คนอื่นในกลุ่มพิมพ์อะไรมาจะไม่ถูกบันทึกและไม่ตอบ
    if (!isKnownMember_(source.userId, props)) {
      Logger.log('ข้ามข้อความจากผู้ใช้ที่ไม่รู้จัก: ' + source.userId);
      return;
    }

    if (ev.type === 'postback' && ev.postback) {
      try {
        handlePostback_(ev, source.userId, props);
      } catch (e) {
        Logger.log('handlePostback_ error: ' + e + '\n' + e.stack);
        if (ev.replyToken) replyLine_(ev.replyToken, '🔧 (โหมดตรวจสอบชั่วคราว)\nเกิดข้อผิดพลาดตอนกดปุ่ม:\n' + e);
      }
      return;
    }

    if (ev.type === 'message' && ev.message && ev.message.type === 'text') {
      try {
        handleExpenseMessage_(ev, source.userId, props);
      } catch (e) {
        Logger.log('handleExpenseMessage_ error: ' + e + '\n' + e.stack);
        if (ev.replyToken) {
          replyLine_(ev.replyToken, '🔧 (โหมดตรวจสอบชั่วคราว)\nเกิดข้อผิดพลาดตอนประมวลผลข้อความ:\n' + e + '\n' + (e.stack || '').slice(0, 500));
        }
      }
    } else if (ev.type === 'message' && ev.message && ev.message.type === 'image') {
      try {
        handleImageMessage_(ev, source.userId, props);
      } catch (e) {
        Logger.log('handleImageMessage_ error: ' + e + '\n' + e.stack);
        if (ev.replyToken) {
          replyLine_(ev.replyToken, '🔧 (โหมดตรวจสอบชั่วคราว)\nเกิดข้อผิดพลาดตอนประมวลผลรูป:\n' + e + '\n' + (e.stack || '').slice(0, 500));
        }
      }
    }
  });
}

// ---------- สิทธิ์การใช้งาน ----------

// สมาชิกที่บอทยอมรับ = โฟกัส (PARTNER1_LINE_ID) กับอ๋อม (PARTNER2_LINE_ID)
// ถ้ายังไม่ได้ตั้ง ID ทั้งสองคนเลย ให้ผ่านไปก่อน (กันล็อกตัวเองตอนติดตั้งครั้งแรก)
function isKnownMember_(userId, props) {
  const p1 = props.getProperty('PARTNER1_LINE_ID');
  const p2 = props.getProperty('PARTNER2_LINE_ID');
  if (!p1 && !p2) return true;
  return !!userId && (userId === p1 || userId === p2);
}

// ผู้ดูแลระบบ = โฟกัส (ใช้ ADMIN_LINE_ID ถ้าตั้งไว้ ไม่งั้นใช้ PARTNER1_LINE_ID)
function isAdmin_(userId, props) {
  const adminId = props.getProperty('ADMIN_LINE_ID') || props.getProperty('PARTNER1_LINE_ID');
  if (!adminId) return true;
  return !!userId && userId === adminId;
}

// คำสั่งที่แก้ไข/ลบ/ตั้งค่าข้อมูล ใช้ได้เฉพาะผู้ดูแลระบบ
const ADMIN_ACTIONS_ = ['delete_last', 'delete_pick', 'fix_category', 'fix_amount', 'set_shared', 'split_half', 'delete_category', 'rename_self', 'set_timezone'];

const HELP_TEXT_ =
  'พิมพ์บอกรายการได้เลยค่ะ เช่น:\n' +
  '• "กินข้าว 150" → บันทึกเป็นรายจ่ายของคนที่พิมพ์\n' +
  '• "ค่าเช่าบ้านร่วมกัน 3000", "หารกันค่าเน็ต 600" หรือ "ค่าน้ำมัน 700 หารครึ่ง" → ความหมายเดียวกันหมด แยกเป็น 2 รายการคนละครึ่งทันที (เช่น คนละ 1,500 บาท) ไม่ใช่ก้อนรวม\n' +
  '• "กินข้าว 150 เมื่อวาน" → บันทึกย้อนหลังเป็นเมื่อวาน (รองรับ "เมื่อวาน"/"เมื่อวานซืน"/"X วันก่อน")\n' +
  '• "กินราเมง 1200 เยน" → บันทึกเป็นเงินไทยให้อัตโนมัติ (แปลงสกุลเงินต่างประเทศให้เอง)\n' +
  'ส่งรูปสลิปเข้ามาได้เลย น้องออมอ่านยอดให้อัตโนมัติ (อ่านวันที่จากสลิปให้ตรงวันจริง และแปลงสกุลเงินต่างประเทศเป็นบาทให้ด้วย)\n' +
  '(ถ้าบันทึกหลัง 20:00 หลังแจ้งเตือนของวันนั้นออกไปแล้ว จะมาโผล่แถมในแจ้งเตือนของวันถัดไปให้อัตโนมัติ ไม่ตกหล่นค่ะ)\n' +
  'พิมพ์บอกได้ตามธรรมชาติเลยค่ะ ไม่ต้องพิมพ์ให้ตรงคำเป๊ะๆ เช่น:\n' +
  '• "ขอดูสรุปหน่อย" / "สรุปเดือนนี้" → การ์ดสรุปพร้อมเทียบเดือนก่อน\n' +
  '• "ดูค่าน้ำมันเดือนนี้" / "ค่ากินของอ๋อมเดือนที่แล้ว" / "เดือนนี้ฉันจ่ายอะไรไปบ้าง" → ค้นหารายการ\n' +
  'คำสั่งด้านล่างนี้ (แก้/ลบ/ตั้งค่า) ใช้ได้เฉพาะผู้ดูแลระบบ:\n' +
  '• "ลบ" → แสดงรายการวันนี้ให้กดปุ่มเลือกลบทีละรายการ ("ลบของเมื่อวาน" เพื่อดูวันอื่น)\n' +
  '• "ลบล่าสุด" / "เอารายการที่เพิ่งจดออก" → ลบรายการล่าสุดทันที\n' +
  '• "แก้ไขเป็นหมวดค่ากิน" / "เปลี่ยนหมวดล่าสุดเป็นค่าเที่ยว"\n' +
  '• "แก้ยอดเป็น 417.06" / "ยอดที่ถูกต้องคือ 500"\n' +
  '• "หารครึ่ง" / "คนละครึ่ง" / "ร่วมกัน" (พิมพ์เดี่ยวๆ หลังบันทึกไปแล้ว) → แยกรายการล่าสุดเป็นคนละครึ่งจริงๆ\n' +
  '• "ลบหมวดของใช้ออกไปเลย"\n' +
  '• "เปลี่ยนชื่อฉันเป็นปาล์ม"\n' +
  '• "อยู่ญี่ปุ่น" / "เปลี่ยนเขตเวลาเป็นโตเกียว" → ตั้งเขตเวลาตอนไปต่างประเทศ (พิมพ์ "กลับไทยแล้ว" เพื่อรีเซ็ตกลับ)';

function handleExpenseMessage_(ev, senderUserId, props) {
  const text = (ev.message.text || '').trim();

  // ทางลัดคำที่พิมพ์บ่อยที่สุดแบบตรงตัว (ตอบทันทีไม่ต้องรอ AI)
  if (text === 'สรุป' || text === 'สรุปเดือนนี้') {
    replyMonthlySummaryFlex_(ev.replyToken, new Date());
    return;
  }
  if (text === 'สรุปวันนี้') {
    replyLine_(ev.replyToken, buildDailySummaryText_());
    return;
  }
  if (text === 'วิธีใช้') {
    replyLine_(ev.replyToken, HELP_TEXT_);
    return;
  }
  if (text === 'ลบ' || text === 'ลบรายการ') {
    if (!isAdmin_(senderUserId, props)) {
      replyLine_(ev.replyToken, 'คำสั่งแก้ไข/ลบ/ตั้งค่า ใช้ได้เฉพาะผู้ดูแลระบบค่ะ ถ้าบันทึกผิด แจ้งโฟกัสให้ช่วยแก้ได้เลยนะคะ');
      return;
    }
    showDeletePicker_(ev.replyToken, localDatePrefix_(new Date(), 'yyyy-MM-dd'));
    return;
  }
  if (text === 'บันทึกรายจ่าย') {
    replyLine_(ev.replyToken, 'พิมพ์บอกรายการพร้อมจำนวนเงินได้เลยค่ะ เช่น "กินข้าว 150"');
    return;
  }

  const data = getAllData_();
  const parsed = parseWithClaude_(text, data.categories);

  if (!parsed || parsed.type === 'none') {
    if (ev.replyToken) {
      replyLine_(ev.replyToken, 'น้องออมยังไม่เข้าใจข้อความนี้ค่ะ ลองพิมพ์ เช่น "กินข้าว 150" หรือพิมพ์ "วิธีใช้" ดูตัวอย่างได้นะคะ');
    }
    return;
  }

  if (parsed.type === 'command') {
    dispatchCommand_(parsed, senderUserId, props, ev.replyToken);
    return;
  }

  if (parsed.type === 'expense' && parsed.amount) {
    logParsedResult_(parsed, senderUserId, props, ev.replyToken);
    return;
  }

  replyLine_(ev.replyToken, 'น้องออมยังไม่เข้าใจข้อความนี้ค่ะ ลองพิมพ์ เช่น "กินข้าว 150" ดูนะคะ');
}

// ---------- ตัวส่งต่อคำสั่งจัดการข้อมูลที่ AI ตีความมาให้ ----------

function dispatchCommand_(parsed, senderUserId, props, replyToken) {
  const action = parsed.action;
  const position = parsed.position && parsed.position > 0 ? Math.round(parsed.position) : 1;

  // คำสั่งแก้ไข/ลบ/ตั้งค่า ใช้ได้เฉพาะโฟกัส (ผู้ดูแลระบบ)
  if (ADMIN_ACTIONS_.indexOf(action) !== -1 && !isAdmin_(senderUserId, props)) {
    replyLine_(replyToken, 'คำสั่งแก้ไข/ลบ/ตั้งค่า ใช้ได้เฉพาะผู้ดูแลระบบค่ะ ถ้าบันทึกผิด แจ้งโฟกัสให้ช่วยแก้ได้เลยนะคะ');
    return;
  }

  if (action === 'delete_pick') {
    const d = /^\d{4}-\d{2}-\d{2}$/.test(parsed.pickDate || '') ? parsed.pickDate : localDatePrefix_(new Date(), 'yyyy-MM-dd');
    showDeletePicker_(replyToken, d);
    return;
  }
  if (action === 'search') {
    searchExpensesCommand_(parsed, senderUserId, props, replyToken);
  } else if (action === 'delete_last') {
    deleteExpenseAtPosition_(position, replyToken);
  } else if (action === 'fix_category') {
    fixCategoryAtPosition_(position, (parsed.category || '').trim(), replyToken);
  } else if (action === 'fix_amount') {
    fixAmountAtPosition_(position, Number(parsed.amount), replyToken);
  } else if (action === 'set_shared') {
    setSharedAtPosition_(position, !!parsed.shared, senderUserId, props, replyToken);
  } else if (action === 'split_half') {
    splitHalfAtPosition_(position, props, replyToken);
  } else if (action === 'delete_category') {
    deleteCategoryCommand_((parsed.category || '').trim(), replyToken);
  } else if (action === 'rename_self') {
    handleRenameSelf_((parsed.newName || '').trim(), senderUserId, props, replyToken);
  } else if (action === 'set_timezone') {
    setTravelTimezone_((parsed.timezone || '').trim(), (parsed.location || '').trim(), replyToken);
  } else if (action === 'summary_today') {
    replyLine_(replyToken, buildDailySummaryText_());
  } else if (action === 'summary_month') {
    replyMonthlySummaryFlex_(replyToken, new Date());
  } else if (action === 'help') {
    replyLine_(replyToken, HELP_TEXT_);
  } else {
    replyLine_(replyToken, 'น้องออมเข้าใจว่าอยากสั่งบางอย่าง แต่ยังไม่แน่ใจว่าอะไรค่ะ ลองพิมพ์ "วิธีใช้" ดูตัวอย่างคำสั่งได้นะคะ');
  }
}

// ---------- รับรูปสลิป: โหลดรูปจาก LINE แล้วให้ AI อ่านยอด/หมวดให้ ----------

function handleImageMessage_(ev, senderUserId, props) {
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token || !ev.replyToken) return;

  let blob;
  try {
    const res = UrlFetchApp.fetch(
      'https://api-data.line.me/v2/bot/message/' + ev.message.id + '/content',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) {
      replyLine_(ev.replyToken, 'น้องออมโหลดรูปไม่สำเร็จค่ะ ลองส่งใหม่อีกครั้งนะคะ');
      return;
    }
    blob = res.getBlob();
  } catch (e) {
    replyLine_(ev.replyToken, 'น้องออมโหลดรูปไม่สำเร็จค่ะ ลองส่งใหม่อีกครั้งนะคะ');
    return;
  }

  const base64Data = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType() || 'image/jpeg';

  const data = getAllData_();
  const parsed = parseImageWithClaude_(base64Data, mimeType, data.categories);

  if (parsed && parsed.type === 'debug_error') {
    replyLine_(ev.replyToken, '🔧 (โหมดตรวจสอบชั่วคราว)\n' + parsed.debug);
    return;
  }

  if (!parsed || parsed.type === 'none' || !parsed.amount) {
    replyLine_(ev.replyToken, 'น้องออมอ่านสลิปนี้ไม่ออกค่ะ ลองส่งรูปที่เห็นยอดเงินชัดๆ อีกครั้ง หรือพิมพ์บอกจำนวนเงินแทนก็ได้นะคะ');
    return;
  }

  logParsedResult_(parsed, senderUserId, props, ev.replyToken);
}

// ---------- บันทึกผลที่ AI แยกวิเคราะห์ได้ (ใช้ร่วมกันทั้งข้อความและรูปสลิป) ----------

function resolveExpenseDate_(parsed) {
  // สลิปรูป: ถ้าอ่านวันที่จริงจากสลิปได้ ให้ใช้วันนั้นตรงๆ เป็นวันที่ท้องถิ่นที่แท้จริง (ไม่ผ่านการตีความเขตเวลาซ้ำ)
  if (parsed.transactionDate && /^\d{4}-\d{2}-\d{2}$/.test(parsed.transactionDate)) {
    const d = new Date(parsed.transactionDate + 'T12:00:00');
    const iso = !isNaN(d.getTime()) ? d.toISOString() : new Date().toISOString();
    return { iso: iso, localDate: parsed.transactionDate };
  }
  // ข้อความ: ถ้าพูดถึงวันย้อนหลัง เช่น "เมื่อวาน" (daysAgo:1) นับตามเขตเวลาที่ใช้งานอยู่ตอนนี้
  const daysAgo = Number(parsed.daysAgo) || 0;
  if (daysAgo > 0) {
    const d = new Date(Date.now() - daysAgo * 86400000);
    return { iso: d.toISOString(), localDate: localDatePrefix_(d, 'yyyy-MM-dd') };
  }
  const now = new Date();
  return { iso: now.toISOString(), localDate: localDatePrefix_(now, 'yyyy-MM-dd') };
}

function logParsedResult_(parsed, senderUserId, props, replyToken) {
  if (parsed.splitHalf) {
    logSplitHalfExpense_(parsed, props, replyToken);
    return;
  }
  const who = parsed.shared ? WHO_SHARED : resolveWho_(senderUserId, props);
  const whoLabel = who === WHO_SHARED ? 'ร่วมกัน' : who;
  const category = parsed.category || 'อื่นๆ';
  const resolved = resolveExpenseDate_(parsed);
  const isBackdated = resolved.localDate !== localDatePrefix_(new Date(), 'yyyy-MM-dd');

  // แปลงสกุลเงินต่างประเทศเป็นบาทก่อนบันทึก (ถ้ามี)
  let amount = Number(parsed.amount);
  let note = parsed.note || '';
  const currency = (parsed.currency || 'THB').toUpperCase();
  if (currency !== 'THB') {
    const converted = convertToTHB_(amount, currency);
    if (converted !== null) {
      note = note
        ? `${note} (${amount.toLocaleString('th-TH')} ${currency} ≈ ${converted.toFixed(2)} บาท)`
        : `${amount.toLocaleString('th-TH')} ${currency} ≈ ${converted.toFixed(2)} บาท`;
      amount = Math.round(converted * 100) / 100;
    } else {
      note = note
        ? `${note} (${amount.toLocaleString('th-TH')} ${currency} - แปลงเป็นบาทไม่สำเร็จ ใช้ตัวเลขเดิม)`
        : `${amount.toLocaleString('th-TH')} ${currency} - แปลงเป็นบาทไม่สำเร็จ ใช้ตัวเลขเดิม`;
    }
  }

  addExpense_(who, category, amount, note, resolved.iso, resolved.localDate, new Date().toISOString());
  addCategory_(category);

  const fresh = getAllData_();
  const monthPrefix = resolved.localDate.slice(0, 7);
  const monthExpenses = fresh.expenses.filter(e => (e.localDate || '').slice(0, 7) === monthPrefix);
  const categoryTotal = monthExpenses
    .filter(e => e.category === category && e.who === who)
    .reduce((s, e) => s + e.amount, 0);

  const labelDate = new Date(resolved.localDate + 'T12:00:00');
  const dateLabel = isBackdated ? ` (บันทึกย้อนหลังวันที่ ${Utilities.formatDate(labelDate, getActiveTimezone_(), 'd MMM yyyy')})` : '';

  replyLine_(replyToken,
    `จดสำเร็จ ✅${dateLabel}\n-${Number(amount).toLocaleString('th-TH')} บาท · ${category} (${whoLabel})${note ? '\n"' + note + '"' : ''}\n\nเดือนนี้หมวด "${category}" ของ${whoLabel} รวม: ${categoryTotal.toLocaleString('th-TH')} บาท`);
}

// ---------- บันทึกรายจ่ายแบบ "หารครึ่ง" จริงๆ: แยกเป็น 2 รายการของแต่ละคนทันที ----------

function logSplitHalfExpense_(parsed, props, replyToken) {
  const category = parsed.category || 'อื่นๆ';
  const resolved = resolveExpenseDate_(parsed);
  const recordedAt = new Date().toISOString();

  let fullAmount = Number(parsed.amount);
  let note = parsed.note || '';
  const currency = (parsed.currency || 'THB').toUpperCase();
  if (currency !== 'THB') {
    const converted = convertToTHB_(fullAmount, currency);
    if (converted !== null) {
      note = note
        ? `${note} (${fullAmount.toLocaleString('th-TH')} ${currency} ≈ ${converted.toFixed(2)} บาท)`
        : `${fullAmount.toLocaleString('th-TH')} ${currency} ≈ ${converted.toFixed(2)} บาท`;
      fullAmount = Math.round(converted * 100) / 100;
    } else {
      note = note
        ? `${note} (${fullAmount.toLocaleString('th-TH')} ${currency} - แปลงเป็นบาทไม่สำเร็จ ใช้ตัวเลขเดิม)`
        : `${fullAmount.toLocaleString('th-TH')} ${currency} - แปลงเป็นบาทไม่สำเร็จ ใช้ตัวเลขเดิม`;
    }
  }

  const half = Math.round((fullAmount / 2) * 100) / 100;
  const [name1, name2] = [props.getProperty('PARTNER1') || 'คนที่ 1', props.getProperty('PARTNER2') || 'คนที่ 2'];

  addExpense_(name1, category, half, note, resolved.iso, resolved.localDate, recordedAt);
  addExpense_(name2, category, half, note, resolved.iso, resolved.localDate, recordedAt);
  addCategory_(category);

  const fresh = getAllData_();
  const monthPrefix = resolved.localDate.slice(0, 7);
  const monthExpenses = fresh.expenses.filter(e => (e.localDate || '').slice(0, 7) === monthPrefix && e.category === category);
  const total1 = monthExpenses.filter(e => e.who === name1).reduce((s, e) => s + e.amount, 0);
  const total2 = monthExpenses.filter(e => e.who === name2).reduce((s, e) => s + e.amount, 0);

  const isBackdated = resolved.localDate !== localDatePrefix_(new Date(), 'yyyy-MM-dd');
  const labelDate = new Date(resolved.localDate + 'T12:00:00');
  const dateLabel = isBackdated ? ` (บันทึกย้อนหลังวันที่ ${Utilities.formatDate(labelDate, getActiveTimezone_(), 'd MMM yyyy')})` : '';

  replyLine_(replyToken,
    `จดสำเร็จ ✅${dateLabel} (หารครึ่งคนละ ${half.toLocaleString('th-TH')} บาท)\n${category} (${name1}): -${half.toLocaleString('th-TH')} บาท\n${category} (${name2}): -${half.toLocaleString('th-TH')} บาท${note ? '\n"' + note + '"' : ''}\n\nเดือนนี้หมวด "${category}" ของ${name1} รวม: ${total1.toLocaleString('th-TH')} บาท\nของ${name2} รวม: ${total2.toLocaleString('th-TH')} บาท`);
}

function deleteExpenseAtPosition_(position, replyToken) {
  const data = getAllData_();
  if (data.expenses.length === 0) {
    replyLine_(replyToken, 'ยังไม่มีรายการให้ลบเลยค่ะ');
    return;
  }
  const sorted = [...data.expenses].sort((a, b) => b.id - a.id);
  const target = sorted[position - 1];
  if (!target) {
    replyLine_(replyToken, `ไม่พบรายการย้อนหลังที่ ${position} ค่ะ (ตอนนี้มีทั้งหมด ${sorted.length} รายการ)`);
    return;
  }
  deleteExpense_(target.id);
  const whoLabel = target.who === WHO_SHARED ? 'ร่วมกัน' : target.who;
  const posLabel = position === 1 ? '' : ` (ย้อนหลังที่ ${position})`;
  replyLine_(replyToken,
    `ลบแล้วค่ะ ✅${posLabel}\n-${Number(target.amount).toLocaleString('th-TH')} บาท · ${target.category} (${whoLabel})${target.note ? '\n"' + target.note + '"' : ''}`);
}

function fixCategoryAtPosition_(position, newCategory, replyToken) {
  if (!newCategory) {
    replyLine_(replyToken, 'พิมพ์ต่อท้ายด้วยชื่อหมวดค่ะ เช่น "แก้เป็น ค่ากิน"');
    return;
  }
  const data = getAllData_();
  if (data.expenses.length === 0) {
    replyLine_(replyToken, 'ยังไม่มีรายการให้แก้เลยค่ะ');
    return;
  }
  const sorted = [...data.expenses].sort((a, b) => b.id - a.id);
  const target = sorted[position - 1];
  if (!target) {
    replyLine_(replyToken, `ไม่พบรายการย้อนหลังที่ ${position} ค่ะ (ตอนนี้มีทั้งหมด ${sorted.length} รายการ)`);
    return;
  }
  const oldCategory = target.category;
  updateExpenseCategory_(target.id, newCategory);
  addCategory_(newCategory);

  const fresh = getAllData_();
  const monthPrefix = localDatePrefix_(new Date(), 'yyyy-MM');
  const categoryTotal = fresh.expenses
    .filter(e => e.category === newCategory && e.who === target.who && (e.localDate || '').slice(0, 7) === monthPrefix)
    .reduce((s, e) => s + e.amount, 0);

  const whoLabel = target.who === WHO_SHARED ? 'ร่วมกัน' : target.who;
  const posLabel = position === 1 ? '' : ` (ย้อนหลังที่ ${position})`;
  replyLine_(replyToken,
    `แก้แล้วค่ะ ✅${posLabel}\n-${Number(target.amount).toLocaleString('th-TH')} บาท (${whoLabel})\n${oldCategory} → ${newCategory}\n\nเดือนนี้หมวด "${newCategory}" ของ${whoLabel} รวม: ${categoryTotal.toLocaleString('th-TH')} บาท`);
}

function fixAmountAtPosition_(position, newAmount, replyToken) {
  if (!newAmount || newAmount <= 0) {
    replyLine_(replyToken, 'พิมพ์ต่อท้ายด้วยจำนวนเงินที่ถูกต้องค่ะ เช่น "แก้ยอดเป็น 417.06"');
    return;
  }
  const data = getAllData_();
  if (data.expenses.length === 0) {
    replyLine_(replyToken, 'ยังไม่มีรายการให้แก้เลยค่ะ');
    return;
  }
  const sorted = [...data.expenses].sort((a, b) => b.id - a.id);
  const target = sorted[position - 1];
  if (!target) {
    replyLine_(replyToken, `ไม่พบรายการย้อนหลังที่ ${position} ค่ะ (ตอนนี้มีทั้งหมด ${sorted.length} รายการ)`);
    return;
  }
  const oldAmount = target.amount;
  updateExpenseAmount_(target.id, newAmount);

  const fresh = getAllData_();
  const monthPrefix = localDatePrefix_(new Date(), 'yyyy-MM');
  const categoryTotal = fresh.expenses
    .filter(e => e.category === target.category && e.who === target.who && (e.localDate || '').slice(0, 7) === monthPrefix)
    .reduce((s, e) => s + e.amount, 0);

  const whoLabel = target.who === WHO_SHARED ? 'ร่วมกัน' : target.who;
  const posLabel = position === 1 ? '' : ` (ย้อนหลังที่ ${position})`;
  replyLine_(replyToken,
    `แก้แล้วค่ะ ✅${posLabel}\n${target.category} (${whoLabel})\n-${Number(oldAmount).toLocaleString('th-TH')} → -${Number(newAmount).toLocaleString('th-TH')} บาท\n\nเดือนนี้หมวด "${target.category}" ของ${whoLabel} รวม: ${categoryTotal.toLocaleString('th-TH')} บาท`);
}

function setSharedAtPosition_(position, makeShared, senderUserId, props, replyToken) {
  const data = getAllData_();
  if (data.expenses.length === 0) {
    replyLine_(replyToken, 'ยังไม่มีรายการให้แก้เลยค่ะ');
    return;
  }
  const sorted = [...data.expenses].sort((a, b) => b.id - a.id);
  const target = sorted[position - 1];
  if (!target) {
    replyLine_(replyToken, `ไม่พบรายการย้อนหลังที่ ${position} ค่ะ (ตอนนี้มีทั้งหมด ${sorted.length} รายการ)`);
    return;
  }
  const oldWho = target.who;
  const oldWhoLabel = oldWho === WHO_SHARED ? 'ร่วมกัน' : oldWho;
  const newWho = makeShared ? WHO_SHARED : resolveWho_(senderUserId, props);
  const newWhoLabel = newWho === WHO_SHARED ? 'ร่วมกัน' : newWho;

  if (oldWho === newWho) {
    replyLine_(replyToken, `รายการนี้เป็น "${newWhoLabel}" อยู่แล้วค่ะ ไม่ต้องแก้`);
    return;
  }

  updateExpenseWho_(target.id, newWho);
  const posLabel = position === 1 ? '' : ` (ย้อนหลังที่ ${position})`;
  replyLine_(replyToken,
    `แก้แล้วค่ะ ✅${posLabel}\n-${Number(target.amount).toLocaleString('th-TH')} บาท · ${target.category}\n${oldWhoLabel} → ${newWhoLabel}`);
}

function splitHalfAtPosition_(position, props, replyToken) {
  const data = getAllData_();
  if (data.expenses.length === 0) {
    replyLine_(replyToken, 'ยังไม่มีรายการให้แก้เลยค่ะ');
    return;
  }
  const sorted = [...data.expenses].sort((a, b) => b.id - a.id);
  const target = sorted[position - 1];
  if (!target) {
    replyLine_(replyToken, `ไม่พบรายการย้อนหลังที่ ${position} ค่ะ (ตอนนี้มีทั้งหมด ${sorted.length} รายการ)`);
    return;
  }
  if (target.who === WHO_SHARED) {
    replyLine_(replyToken, 'รายการนี้เป็นรายจ่ายร่วมกันแบบก้อนเดียวอยู่แล้วค่ะ ถ้าอยากแยกเป็นคนละครึ่งจริงๆ ต้องลบแล้วพิมพ์บันทึกใหม่พร้อมคำว่า "หารครึ่ง" นะคะ');
    return;
  }

  const half = Math.round((target.amount / 2) * 100) / 100;
  const [name1, name2] = [props.getProperty('PARTNER1') || 'คนที่ 1', props.getProperty('PARTNER2') || 'คนที่ 2'];

  deleteExpense_(target.id);
  addExpense_(name1, target.category, half, target.note, target.date, target.localDate, new Date().toISOString());
  addExpense_(name2, target.category, half, target.note, target.date, target.localDate, new Date().toISOString());

  const posLabel = position === 1 ? '' : ` (ย้อนหลังที่ ${position})`;
  replyLine_(replyToken,
    `แยกเป็นคนละครึ่งแล้วค่ะ ✅${posLabel}\n${target.category}: -${Number(target.amount).toLocaleString('th-TH')} บาท (${target.who})\n→ แยกเป็น ${name1} -${half.toLocaleString('th-TH')} บาท และ ${name2} -${half.toLocaleString('th-TH')} บาท`);
}

function resolveWho_(userId, props) {
  if (userId && userId === props.getProperty('PARTNER1_LINE_ID')) return props.getProperty('PARTNER1');
  if (userId && userId === props.getProperty('PARTNER2_LINE_ID')) return props.getProperty('PARTNER2');
  return props.getProperty('PARTNER1') || 'ไม่ระบุ';
}

function resolveSenderSlot_(userId, props) {
  if (userId && userId === props.getProperty('PARTNER1_LINE_ID')) return 'PARTNER1';
  if (userId && userId === props.getProperty('PARTNER2_LINE_ID')) return 'PARTNER2';
  return null;
}

function handleRenameSelf_(newName, senderUserId, props, replyToken) {
  if (!newName) {
    replyLine_(replyToken, 'พิมพ์ต่อท้ายด้วยชื่อใหม่ค่ะ เช่น "ชื่อฉันคือ ปาล์ม"');
    return;
  }
  const slot = resolveSenderSlot_(senderUserId, props);
  if (!slot) {
    replyLine_(replyToken, 'น้องออมยังไม่รู้ว่าคุณคือคนที่ 1 หรือคนที่ 2 ค่ะ ต้องตั้งค่า PARTNER1_LINE_ID/PARTNER2_LINE_ID ให้ตรงกับ userId ของคุณก่อนนะคะ (ดูวิธีในคู่มือติดตั้ง)');
    return;
  }
  const oldName = props.getProperty(slot) || (slot === 'PARTNER1' ? 'คนที่ 1' : 'คนที่ 2');
  if (oldName === newName) {
    replyLine_(replyToken, `ชื่อ "${newName}" เหมือนเดิมอยู่แล้วค่ะ`);
    return;
  }
  renamePartnerInExpenses_(oldName, newName);
  props.setProperty(slot, newName);
  syncDashboardNames_();
  replyLine_(replyToken,
    `เปลี่ยนชื่อแล้วค่ะ ✅\n${oldName} → ${newName}\n\nรายการเก่าที่เคยบันทึกไว้ในชื่อ "${oldName}" ปรับเป็น "${newName}" ให้หมดแล้วค่ะ`);
}

function setTravelTimezone_(timezone, location, replyToken) {
  if (!timezone) {
    replyLine_(replyToken, 'บอกชื่อประเทศ/เมืองที่อยากตั้งเขตเวลาด้วยค่ะ เช่น "อยู่ญี่ปุ่น" หรือ "เปลี่ยนเขตเวลาเป็นโตเกียว"');
    return;
  }
  const props = PropertiesService.getScriptProperties();
  // ทดสอบว่าเขตเวลาที่ AI ส่งมาใช้งานได้จริงก่อน กันตั้งค่าเป็นเขตเวลาที่ไม่มีอยู่จริง
  try {
    Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
  } catch (e) {
    replyLine_(replyToken, `น้องออมไม่รู้จักเขตเวลา "${timezone}" ค่ะ ลองระบุชื่อประเทศ/เมืองใหม่อีกครั้งนะคะ`);
    return;
  }
  props.setProperty('TRAVEL_TIMEZONE', timezone);
  const label = location || timezone;
  const isBangkok = timezone === 'Asia/Bangkok';
  replyLine_(replyToken,
    `ตั้งเขตเวลาเป็น ${label} แล้วค่ะ ✅\n${isBangkok ? 'กลับมาใช้เขตเวลาไทยตามปกติแล้วค่ะ' : 'ตอนนี้ระบบจะนับ "วันนี้"/"เมื่อวาน"/สรุปประจำวันตามเขตเวลานี้แทนเขตเวลาไทย จนกว่าจะเปลี่ยนกลับ (พิมพ์ "กลับไทยแล้ว" เพื่อรีเซ็ตกลับเมื่อไหร่ก็ได้)'}`);
}

function deleteCategoryCommand_(category, replyToken) {
  if (!category) {
    replyLine_(replyToken, 'พิมพ์ต่อท้ายด้วยชื่อหมวดค่ะ เช่น "ลบหมวด ของใช้"');
    return;
  }
  const data = getAllData_();
  if (data.categories.indexOf(category) === -1) {
    replyLine_(replyToken, `ไม่พบหมวด "${category}" ในระบบค่ะ ลองพิมพ์ "สรุป" เพื่อดูชื่อหมวดที่มีอยู่ตอนนี้ได้นะคะ`);
    return;
  }
  deleteCategory_(category);
  replyLine_(replyToken,
    `ลบหมวด "${category}" ออกจากระบบแล้วค่ะ ✅\n(รายการเก่าที่เคยบันทึกไว้ในหมวดนี้จะยังอยู่เหมือนเดิม แค่จะไม่มีในตัวเลือกหมวดสำหรับรายการใหม่แล้วค่ะ)`);
}

// ---------- เรียก Claude API เพื่อแยกวิเคราะห์ข้อความ ----------

function parseWithClaude_(text, categories) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return null;

  const sys = 'คุณคือน้องออม ผู้ช่วยของระบบบันทึกรายจ่ายคู่รักในแชท LINE วิเคราะห์ข้อความภาษาไทยที่คนพิมพ์มา แล้วตอบกลับเป็น JSON เท่านั้น ห้ามมีคำอธิบายอื่นหรือ markdown fence\n\n' +
    'วันนี้คือวันที่ ' + localDatePrefix_(new Date(), 'yyyy-MM-dd') + ' (ใช้คำนวณ "เดือนนี้"/"เดือนที่แล้ว" ได้)\n' +
    'หมวดรายจ่ายที่มีอยู่แล้ว: ' + (categories.join(', ') || '(ยังไม่มี)') + '\n' +
    'วิธีเลือกหมวดรายจ่าย ให้ตรวจตามลำดับนี้ (หยุดที่ข้อแรกที่ตรง อย่าข้ามไปข้อถัดไป):\n' +
    '1. ถ้าชื่อร้าน/รายละเอียดสื่อถึงอาหารหรือร้านอาหารชัดเจน (มีคำว่า ร้าน, ครัว, เฮือน, ก๋วยเตี๋ยว, หมู, ไก่, ข้าว, ส้มตำ, อาหารตามสั่ง ฯลฯ) → หมวด "ค่ากิน"\n' +
    '2. ถ้าเป็นชื่อร้าน/แบรนด์ปั๊มน้ำมันชัดเจน เช่น PT, ปตท., PTT, บางจาก, เชลล์ (Shell), เอสโซ่ (Esso), คาลเท็กซ์ (Caltex), ซัสโก้ (Susco), หรือคำว่า "Max Card"/"แมกซ์ การ์ด" → หมวด "ค่าน้ำมัน" (สร้างหมวดนี้ถ้ายังไม่มี)\n' +
    '3. ถ้าเป็นการซื้อเครื่องดื่มโดยตรงเท่านั้น (มีคำว่า ชานม กาแฟ น้ำหวาน น้ำอัดลม สมูทตี้ ชัดเจน) → หมวด "ค่าน้ำ" — ห้ามใช้หมวดนี้เป็นค่าเริ่มต้นเวลาไม่แน่ใจ ใช้เฉพาะกรณีชัดเจนว่าเป็นเครื่องดื่มเท่านั้น (หมวด "ค่าน้ำ" ไม่เกี่ยวกับค่าน้ำประปา ถ้าเป็นบิลน้ำประปา/ค่าน้ำมิเตอร์ ให้ใช้หมวด "ค่าน้ำประปา" แยกต่างหากแทน)\n' +
    '4. ถ้าไม่เข้าเงื่อนไขข้างต้นเลย ให้เลือกหมวดที่ใกล้เคียงที่สุดจากลิสต์ที่มีอยู่ หรือตั้งชื่อหมวดใหม่สั้นๆที่ตรงความหมายที่สุด\n\n' +
    'ข้อความที่ได้รับมาอาจเป็นได้ 3 แบบ:\n' +
    '1) การบันทึกรายจ่ายใหม่ เช่น "กินข้าว 150", "ชานมไข่มุก 65", "ค่าเช่าบ้านร่วมกัน 3000" (splitHalf), "หารกันค่าเน็ต 600" (splitHalf), "ค่าน้ำมัน 700 หารครึ่ง" (splitHalf), "กินข้าว 150 เมื่อวาน", "ค่าน้ำมัน 300 เมื่อวานซืน", "กินข้าว 100 เมื่อ 3 วันก่อน", "กินราเมง 1200 เยน", "ค่าที่พัก 80 ดอลลาร์"\n' +
    '2) คำสั่งจัดการข้อมูล — ไม่จำเป็นต้องพิมพ์ตรงคำเป๊ะ ให้ตีความเจตนาจากความหมายทั่วไป เช่น "ลบอันนี้ที", "เอารายการล่าสุดออก" (ลบรายการล่าสุดทันที → action delete_last), "ลบ", "อยากลบรายการวันนี้", "เลือกลบ", "ลบของเมื่อวาน", "ลบรายการวันที่ 3" (อยากเลือกเองว่าจะลบรายการไหนของวันนั้น → action delete_pick), "แก้ไขเป็นหมวดค่ากิน", "เปลี่ยนหมวดล่าสุดเป็นค่าเที่ยว", "แก้ยอดเป็น 417.06", "ยอดที่ถูกต้องคือ 500 ไม่ใช่ 300", "หารครึ่ง"/"คนละครึ่ง"/"ร่วมกัน"/"หารกัน" (พิมพ์เดี่ยวๆ หลังเพิ่งบันทึกไป ความหมายเดียวกันหมด คืออยากให้รายการล่าสุดแยกเป็นคนละครึ่งจริงๆ → action split_half), "ลบสลิปนี้ออก", "ไม่เอาหมวดของใช้แล้ว", "ลบหมวดของใช้ออกไปเลย", "เปลี่ยนชื่อฉันเป็นปาล์ม", "อยู่ญี่ปุ่น", "เปลี่ยนเขตเวลาเป็นโตเกียว", "กลับไทยแล้ว" (ตั้ง/รีเซ็ตเขตเวลาตอนไปต่างประเทศ), "ขอดูสรุปหน่อย", "สรุปวันนี้ยังไง", "วิธีใช้เป็นยังไง", "ดูค่าน้ำมันเดือนนี้", "ค่ากินของอ๋อมเดือนที่แล้ว", "เดือนนี้ฉันจ่ายอะไรไปบ้าง" (ค้นหารายการ → action search) ล้วนถือเป็นคำสั่งจัดการข้อมูลทั้งสิ้น\n' +
    '3) ข้อความทั่วไปที่ไม่เกี่ยวข้องเลย เช่น คำทักทาย คำถามทั่วไป\n\n' +
    'ตอบ JSON ตามแบบใดแบบหนึ่งนี้เท่านั้น (ห้ามใส่ฟิลด์อื่นนอกเหนือจากนี้):\n\n' +
    'แบบที่ 1 (บันทึกรายจ่าย):\n' +
    '{"type":"expense","amount":ตัวเลข,"category":"ชื่อหมวดที่ใกล้เคียงที่สุดจากลิสต์ หรือชื่อหมวดใหม่สั้นๆถ้าไม่มีหมวดไหนใกล้เคียง","shared":true หรือ false,"splitHalf":true หรือ false,"daysAgo":จำนวนวันย้อนหลังจากวันนี้ (0 ถ้าไม่ได้พูดถึงเวลาหรือหมายถึงวันนี้/เมื่อกี้; "เมื่อวาน"=1; "เมื่อวานซืน"=2; "X วันก่อน"/"เมื่อ X วันก่อน"=X),"currency":"รหัสสกุลเงิน ISO 3 ตัวอักษรตัวพิมพ์ใหญ่ เช่น THB, JPY, USD, EUR, KRW (ค่าเริ่มต้นคือ THB ถ้าไม่ได้พูดถึงสกุลเงินอื่นชัดเจน เช่นมีคำว่า เยน, ดอลลาร์, ยูโร, วอน หรือสัญลักษณ์ ¥ $ € ₩ ก็ให้ใส่รหัสตามนั้น ไม่ต้องคำนวณแปลงเป็นบาทเอง ใส่ตัวเลข amount ตามสกุลเงินเดิมตรงๆ ระบบจะแปลงให้เอง)","note":"รายละเอียดสั้นๆถ้ามี ไม่งั้นเป็นค่าว่าง"}\n' +
    'สำคัญ: ถ้าข้อความสื่อว่าเป็นค่าใช้จ่ายร่วมของทั้งคู่ในความหมายใดก็ตาม — ไม่ว่าจะมีคำว่า "ร่วมกัน", "หารกัน", "ด้วยกัน", "หารครึ่ง", "คนละครึ่ง", "แบ่งครึ่ง", "แบ่งจ่ายคนละครึ่ง" — ให้ตอบ "splitHalf":true เสมอ (ทุกคำเหล่านี้ถือว่าความหมายเดียวกันหมด คือต้องการแบ่งจ่ายคนละครึ่งจริงๆ แยกเป็นของแต่ละคน) และ "shared" ให้เป็น false เสมอ (ไม่ต้องใช้ "shared":true อีกต่อไปสำหรับการบันทึกรายจ่ายใหม่)\n' +
    'ห้ามคำนวณ/หารตัวเลขเองเด็ดขาดไม่ว่ากรณีใด ให้ใส่ "amount" เท่ากับจำนวนเงินเต็มที่ปรากฏในข้อความเสมอ เช่น "ค่าน้ำมัน 700 หารครึ่ง" หรือ "ค่าน้ำมัน 700 ร่วมกัน" ต้องได้ amount:700 (ไม่ใช่ 350) และ splitHalf:true เหมือนกัน — ระบบฝั่งหลังบ้านจะเป็นคนคำนวณหารเอง ไม่ใช่หน้าที่ของคุณ\n\n' +
    'แบบที่ 2 (คำสั่งจัดการข้อมูล):\n' +
    '{"type":"command","action":"delete_last หรือ fix_category หรือ fix_amount หรือ set_shared หรือ split_half หรือ delete_category หรือ rename_self หรือ set_timezone หรือ summary_today หรือ summary_month หรือ search หรือ delete_pick หรือ help","position":ตัวเลขลำดับย้อนหลังนับจากรายการล่าสุด (ไม่ระบุให้เป็น 1 ใช้กับ delete_last, fix_category, fix_amount, set_shared และ split_half เท่านั้น),"category":"ชื่อหมวด (ใช้กับ fix_category และ delete_category)","amount":ตัวเลขจำนวนเงินที่ถูกต้อง (ใช้กับ fix_amount เท่านั้น ห้ามคำนวณ/หารเอง ใส่ตามตัวเลขที่ผู้ใช้ระบุตรงๆ)","shared":true หรือ false (ใช้กับ set_shared เท่านั้น — true ถ้าอยากให้เป็นก้อนรวมของทั้งคู่ไม่แยกคน, false ถ้าอยากให้กลับเป็นรายจ่ายส่วนตัวของคนที่พิมพ์),"newName":"ชื่อใหม่ (ใช้กับ rename_self)","timezone":"ชื่อเขตเวลาแบบ IANA เช่น Asia/Tokyo, Asia/Seoul, Europe/Paris, America/New_York, Asia/Bangkok (ใช้กับ set_timezone เท่านั้น เดาจากชื่อประเทศ/เมืองที่พูดถึง เช่น \'อยู่ญี่ปุ่น\'→Asia/Tokyo, \'กลับไทยแล้ว\'→Asia/Bangkok)","location":"ชื่อประเทศ/เมืองที่มนุษย์อ่านเข้าใจง่าย (ใช้กับ set_timezone เท่านั้น เช่น \'ญี่ปุ่น\', \'ไทย\')","searchCategory":"ชื่อหมวดที่อยากค้นหา (ใช้กับ search เท่านั้น ใช้ชื่อจากลิสต์หมวดที่มีอยู่ ถ้าไม่ระบุหมวด/อยากดูทุกหมวดให้เป็นค่าว่าง)","searchMonth":"เดือนที่อยากค้นหา รูปแบบ YYYY-MM (ใช้กับ search เท่านั้น ถ้าไม่ระบุให้เป็นเดือนปัจจุบัน)","pickDate":"วันที่ของรายการที่อยากเลือกลบ รูปแบบ YYYY-MM-DD (ใช้กับ delete_pick เท่านั้น ถ้าไม่ระบุให้เป็นวันนี้ เช่น \'เมื่อวาน\' ให้คำนวณจากวันนี้)","searchPerson":"self ถ้าหมายถึงคนที่พิมพ์เอง เช่น \'ของฉัน\'/\'ฉันจ่าย\', หรือชื่อคนถ้าระบุชื่อ, หรือค่าว่างถ้าดูของทุกคน (ใช้กับ search เท่านั้น)"}\n\n' +
    'แบบที่ 3 (ไม่เกี่ยวข้อง):\n' +
    '{"type":"none"}\n\n' +
    'ข้อบังคับที่สำคัญที่สุด: ตอบกลับด้วย JSON object เดียวเท่านั้นในคำตอบทั้งหมด ห้ามมีข้อความอื่นใดๆ ก่อนหรือหลัง JSON เด็ดขาด (ห้ามทักทาย ห้ามอธิบาย ห้ามขอโทษ ห้ามแก้ไขคำตอบตัวเองกลางข้อความ แม้จะสังเกตเห็นว่าตอบผิดไปแล้วก็ตาม) ให้คิดคำตอบให้รอบคอบและถูกต้องในใจก่อน แล้วค่อยพิมพ์ JSON ออกมาเพียงครั้งเดียวเท่านั้น';

  const payload = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 300,
    system: sys,
    messages: [{ role: 'user', content: text }]
  };

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    Logger.log('Claude API error: ' + res.getContentText());
    return null;
  }

  try {
    const data = JSON.parse(res.getContentText());
    const textBlock = (data.content || []).find(b => b.type === 'text');
    const raw = textBlock ? textBlock.text : '';
    return extractJson_(raw);
  } catch (e) {
    Logger.log('parse fail: ' + e);
    return null;
  }
}

// ดึงเฉพาะส่วน JSON ออกมาจากคำตอบของ AI แม้จะมีข้อความอื่นห่อล้อมอยู่ (กันกรณีโมเดลตอบไม่ตรงรูปแบบเป๊ะๆ)
function extractJson_(raw) {
  const text = (raw || '').trim();

  // หา JSON ก้อนสุดท้ายเสมอ เผื่อโมเดลตอบมาหลายก้อน (เช่น ตอบผิดแล้วขอโทษ/แก้คำตอบใหม่กลางข้อความ)
  // จับแบบ non-greedy คู่ {...} ที่ไม่มีวงเล็บปีกกาซ้อนภายใน (สคีมาของเราเป็น flat object เสมอ)
  const matches = text.match(/\{[^{}]*\}/g);
  if (matches && matches.length > 0) {
    for (let i = matches.length - 1; i >= 0; i--) {
      try {
        return JSON.parse(matches[i]);
      } catch (e) {
        // ก้อนนี้พาร์สไม่ได้ ลองก้อนก่อนหน้าต่อ
      }
    }
  }

  Logger.log('extractJson_ ไม่พบ JSON ที่ใช้ได้ในคำตอบ: ' + text);
  return null;
}

// ---------- เรียก Claude API เพื่ออ่านรูปสลิปโอนเงิน ----------

function parseImageWithClaude_(base64Data, mimeType, categories) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return null;

  const sys = 'คุณคือตัวช่วยอ่านรูปสลิปโอนเงิน/ใบเสร็จภาษาไทยหรืออังกฤษ เพื่อบันทึกลงระบบรายจ่ายคู่รัก\n' +
    'หมวดรายจ่ายที่มีอยู่แล้ว: ' + (categories.join(', ') || '(ยังไม่มี)') + '\n' +
    'สำคัญที่สุด: ถ้าในสลิปมีข้อความ "บันทึกช่วยจำ" หรือโน้ต/หมายเหตุที่ผู้ใช้พิมพ์เองต่อท้ายสลิป (มักอยู่แถบล่างสุดของสลิป) ให้ใช้ข้อความนั้นเป็นสัญญาณหลักในการเลือกหมวดและเขียนลง "note" เสมอ เพราะเป็นสิ่งที่ผู้ใช้ตั้งใจระบุเองว่าจ่ายค่าอะไร สำคัญกว่าชื่อร้าน/ชื่อผู้รับเงินในสลิป\n' +
    'วิธีเลือกหมวดรายจ่าย ให้ตรวจตามลำดับนี้ (หยุดที่ข้อแรกที่ตรง อย่าข้ามไปข้อถัดไป):\n' +
    '0. ถ้ามี "บันทึกช่วยจำ"/โน้ตที่ผู้ใช้พิมพ์เองในสลิป ให้ใช้ข้อความนั้นตัดสินหมวดโดยตรง (เช่น เขียนว่า "ค่าข้าว" ก็คือหมวด "ค่ากิน")\n' +
    '1. ถ้าชื่อร้าน/รายละเอียดในสลิปสื่อถึงอาหารหรือร้านอาหารชัดเจน (มีคำว่า ร้าน, ครัว, เฮือน, ก๋วยเตี๋ยว, หมู, ไก่, ข้าว, ส้มตำ, อาหารตามสั่ง ฯลฯ) → หมวด "ค่ากิน"\n' +
    '2. ถ้าเป็นชื่อร้าน/แบรนด์ปั๊มน้ำมันชัดเจน เช่น PT, ปตท., PTT, บางจาก, เชลล์ (Shell), เอสโซ่ (Esso), คาลเท็กซ์ (Caltex), ซัสโก้ (Susco), หรือคำว่า "Max Card"/"แมกซ์ การ์ด" → หมวด "ค่าน้ำมัน" (สร้างหมวดนี้ถ้ายังไม่มี)\n' +
    '3. ถ้าเป็นการซื้อเครื่องดื่มโดยตรงเท่านั้น (มีคำว่า ชานม กาแฟ น้ำหวาน น้ำอัดลม สมูทตี้ ชัดเจน) → หมวด "ค่าน้ำ" — ห้ามใช้หมวดนี้เป็นค่าเริ่มต้นเวลาไม่แน่ใจ (หมวด "ค่าน้ำ" ไม่เกี่ยวกับค่าน้ำประปา ถ้าเป็นบิลน้ำประปา/ค่าน้ำมิเตอร์ ให้ใช้หมวด "ค่าน้ำประปา" แยกต่างหากแทน)\n' +
    '4. ถ้าไม่เข้าเงื่อนไขข้างต้นเลย ให้เลือกหมวดที่ใกล้เคียงที่สุดจากลิสต์ที่มีอยู่ หรือตั้งชื่อหมวดใหม่สั้นๆที่ตรงความหมายที่สุด\n' +
    'อ่านยอดเงินรวม (ยอดที่โอน/จ่ายจริง) จากรูป แล้วตอบกลับเป็น JSON เท่านั้น ห้ามมีคำอธิบายอื่นหรือ markdown fence รูปแบบ:\n' +
    '{"type":"expense" หรือ "none","amount": ตัวเลขหรือ null,"category":"ชื่อหมวดที่ใกล้เคียงที่สุดจากลิสต์ หรือชื่อหมวดใหม่สั้นๆถ้าไม่มีหมวดไหนใกล้เคียง","shared": false,"currency":"รหัสสกุลเงิน ISO 3 ตัวอักษรตัวพิมพ์ใหญ่ตามที่ปรากฏในสลิป เช่น THB, JPY, USD, EUR, KRW (สังเกตจากสัญลักษณ์ ¥ $ € ₩ หรือชื่อสกุลเงินที่พิมพ์ไว้ ถ้าไม่ระบุเลยหรือเป็นสลิปธนาคารไทยปกติให้ใส่ THB ใส่ตัวเลข amount ตามสกุลเงินเดิมในสลิปตรงๆ ไม่ต้องคำนวณแปลงเอง)","note":"ชื่อร้าน/ผู้รับเงิน หรือรายละเอียดสั้นๆจากสลิปถ้ามี ไม่งั้นเป็นค่าว่าง","transactionDate":"วันที่ทำรายการตามที่ปรากฏในสลิป รูปแบบ YYYY-MM-DD (ถ้าปีในสลิปเป็น พ.ศ. ให้แปลงเป็น ค.ศ. โดยลบ 543 ก่อน เช่น สลิประบุ 9 ก.ย. 69 ให้ตอบ 2026-09-09) ถ้าอ่านวันที่จากสลิปไม่ได้เลยให้เป็นค่าว่าง"}\n' +
    'ข้อสำคัญ: ให้ตอบ {"type":"none"} เฉพาะกรณีที่อ่าน "จำนวนเงิน" ในสลิปไม่ออกจริงๆ เท่านั้น (เช่น ภาพเบลอ/มืด/ไม่ใช่สลิปการเงิน) — ถ้าอ่านจำนวนเงินได้ชัดเจน ต้องตอบเป็น "expense" เสมอ ห้ามตอบ "none" เพียงเพราะไม่แน่ใจว่าเป็นหมวดอะไร (เช่น สลิปซื้อลอตเตอรี่ ค่าธรรมเนียม ค่าปรับ ของที่ไม่เข้าหมวดอาหาร/น้ำมัน/เครื่องดื่มเลย) ในกรณีแบบนี้ให้ใช้ข้อ 5 (ตั้งชื่อหมวดใหม่ที่ตรงความหมายที่สุด เช่น "หวย", "ค่าปรับ") แล้วบันทึกยอดเงินให้ถูกต้องเสมอ\n' +
    'ข้อบังคับที่สำคัญที่สุด: ตอบกลับด้วย JSON object เดียวเท่านั้นในคำตอบทั้งหมด ห้ามมีข้อความอื่นใดๆ ก่อนหรือหลัง JSON เด็ดขาด (ห้ามทักทาย ห้ามอธิบาย ห้ามขอโทษ ห้ามแก้ไขคำตอบตัวเองกลางข้อความ แม้จะสังเกตเห็นว่าตอบผิดไปแล้วก็ตาม) ให้คิดคำตอบให้รอบคอบและถูกต้องในใจก่อน แล้วค่อยพิมพ์ JSON ออกมาเพียงครั้งเดียวเท่านั้น';

  const payload = {
    model: 'claude-sonnet-5',
    max_tokens: 300,
    system: sys,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mimeType, data: base64Data } },
        { type: 'text', text: 'อ่านยอดเงินจากสลิปนี้ให้หน่อยค่ะ' }
      ]
    }]
  };

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    const errText = res.getContentText();
    Logger.log('Claude API (image) error: ' + errText);
    return { type: 'debug_error', debug: 'HTTP ' + res.getResponseCode() + ': ' + errText.slice(0, 400) };
  }

  try {
    const data = JSON.parse(res.getContentText());
    const textBlock = (data.content || []).find(b => b.type === 'text');
    const raw = textBlock ? textBlock.text : '';
    const result = extractJson_(raw);
    if (!result) {
      return { type: 'debug_error', debug: 'AI ตอบว่า: ' + raw.slice(0, 400) };
    }
    return result;
  } catch (e) {
    Logger.log('parse fail (image): ' + e);
    return { type: 'debug_error', debug: 'Exception: ' + e };
  }
}

// ---------- หาว่าจะส่งแจ้งเตือนไปหาใครบ้าง (ส่วนตัวก่อน, กลุ่มเป็นสำรอง) ----------

function getNotifyTargets_(props) {
  const p1 = props.getProperty('PARTNER1_LINE_ID');
  const p2 = props.getProperty('PARTNER2_LINE_ID');
  const personal = [p1, p2].filter(Boolean);
  const unique = personal.filter((id, i) => personal.indexOf(id) === i);
  if (unique.length > 0) return unique;

  const groupId = props.getProperty('LINE_TARGET_ID');
  return groupId ? [groupId] : [];
}

// ---------- LINE reply/push helpers ----------

function replyLine_(replyToken, text) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ replyToken, messages: [{ type: 'text', text }] }),
    muteHttpExceptions: true
  });
}

function pushLine_(token, to, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to, messages: [{ type: 'text', text }] }),
    muteHttpExceptions: true
  });
}

// ---------- ส่งข้อความหลายแบบ (รองรับการ์ด Flex) ----------

// mode = 'reply' (target = replyToken) หรือ 'push' (target = userId/groupId) คืนค่า true ถ้าส่งสำเร็จ
function sendLineMessages_(mode, target, messages) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) return false;
  const url = mode === 'reply' ? 'https://api.line.me/v2/bot/message/reply' : 'https://api.line.me/v2/bot/message/push';
  const body = mode === 'reply' ? { replyToken: target, messages: messages } : { to: target, messages: messages };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    Logger.log('sendLineMessages_ fail: ' + res.getResponseCode() + ' ' + res.getContentText());
    return false;
  }
  return true;
}

// ---------- สรุปรายเดือนแบบการ์ด พร้อมเทียบเดือนก่อน ----------

const TH_MONTHS_ = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

function monthLabelTh_(yyyyMM) {
  const parts = yyyyMM.split('-');
  return TH_MONTHS_[Number(parts[1]) - 1] + ' ' + parts[0];
}

function prevMonthKey_(yyyyMM) {
  const parts = yyyyMM.split('-');
  const d = new Date(Number(parts[0]), Number(parts[1]) - 2, 1);
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
}

function fmtBaht_(n) {
  return Number(n || 0).toLocaleString('th-TH', { maximumFractionDigits: 2 });
}

// คืน {text, color} ของเปอร์เซ็นต์เปลี่ยนแปลงเทียบเดือนก่อน (ใช้มากขึ้น = แดง, ใช้น้อยลง = เขียว)
function changeLabel_(cur, prev) {
  if (!prev) return cur ? { text: 'ใหม่', color: '#888888' } : { text: '-', color: '#888888' };
  const pct = Math.round(((cur - prev) / prev) * 100);
  if (pct === 0) return { text: 'เท่าเดิม', color: '#888888' };
  return pct > 0 ? { text: '▲ ' + pct + '%', color: '#C0392B' } : { text: '▼ ' + Math.abs(pct) + '%', color: '#1E8449' };
}

function flexRow_(label, value, change, bold) {
  return {
    type: 'box', layout: 'horizontal', margin: 'sm',
    contents: [
      { type: 'text', text: label || '-', size: 'sm', color: bold ? '#111111' : '#555555', weight: bold ? 'bold' : 'regular', flex: 4, wrap: true },
      { type: 'text', text: value || '-', size: 'sm', color: '#111111', weight: bold ? 'bold' : 'regular', align: 'end', flex: 3 },
      { type: 'text', text: change ? change.text : ' ', size: 'xs', color: change ? change.color : '#888888', align: 'end', gravity: 'center', flex: 2 }
    ]
  };
}

function buildMonthlySummaryFlex_(dateInMonth) {
  const data = getAllData_();
  const monthKey = localDatePrefix_(dateInMonth, 'yyyy-MM');
  const prevKey = prevMonthKey_(monthKey);
  const [name1, name2] = data.partners;

  const cur = data.expenses.filter(e => (e.localDate || '').slice(0, 7) === monthKey);
  const prev = data.expenses.filter(e => (e.localDate || '').slice(0, 7) === prevKey);
  const sumWho = (list, who) => list.filter(e => e.who === who).reduce((s, e) => s + e.amount, 0);
  const sumAll = list => list.reduce((s, e) => s + e.amount, 0);

  const body = [];
  if (cur.length === 0) {
    body.push({ type: 'text', text: 'เดือนนี้ยังไม่มีรายการบันทึกไว้เลย', size: 'sm', color: '#555555', wrap: true });
  } else {
    body.push(flexRow_(name1, fmtBaht_(sumWho(cur, name1)), changeLabel_(sumWho(cur, name1), sumWho(prev, name1))));
    body.push(flexRow_(name2, fmtBaht_(sumWho(cur, name2)), changeLabel_(sumWho(cur, name2), sumWho(prev, name2))));
    const sharedCur = sumWho(cur, WHO_SHARED);
    if (sharedCur > 0) body.push(flexRow_('ร่วมกัน', fmtBaht_(sharedCur), changeLabel_(sharedCur, sumWho(prev, WHO_SHARED))));
    body.push({ type: 'separator', margin: 'md' });
    body.push(flexRow_('รวมทั้งหมด', fmtBaht_(sumAll(cur)), changeLabel_(sumAll(cur), sumAll(prev)), true));

    const catCur = {}, catPrev = {};
    cur.forEach(e => { catCur[e.category] = (catCur[e.category] || 0) + e.amount; });
    prev.forEach(e => { catPrev[e.category] = (catPrev[e.category] || 0) + e.amount; });
    const cats = Object.keys(catCur).sort((a, b) => catCur[b] - catCur[a]).slice(0, 8);
    body.push({ type: 'text', text: 'แยกตามหมวด', size: 'xs', color: '#888888', margin: 'lg' });
    cats.forEach(c => body.push(flexRow_(c, fmtBaht_(catCur[c]), changeLabel_(catCur[c], catPrev[c] || 0))));
    body.push({ type: 'text', text: 'เทียบกับ ' + monthLabelTh_(prevKey), size: 'xxs', color: '#AAAAAA', margin: 'lg', align: 'end' });
  }

  return {
    type: 'flex',
    altText: 'สรุปรายจ่าย ' + monthLabelTh_(monthKey) + ' รวม ' + fmtBaht_(sumAll(cur)) + ' บาท',
    contents: {
      type: 'bubble', size: 'mega',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#173C36', paddingAll: 'lg',
        contents: [
          { type: 'text', text: 'สรุปรายจ่าย', size: 'xs', color: '#D9BE6E' },
          { type: 'text', text: monthLabelTh_(monthKey), size: 'lg', weight: 'bold', color: '#FFFFFF' }
        ]
      },
      body: { type: 'box', layout: 'vertical', paddingAll: 'lg', contents: body }
    }
  };
}

function replyMonthlySummaryFlex_(replyToken, dateInMonth) {
  const ok = sendLineMessages_('reply', replyToken, [buildMonthlySummaryFlex_(dateInMonth)]);
  if (!ok) replyLine_(replyToken, buildMonthlySummaryText_(dateInMonth)); // การ์ดส่งไม่ผ่าน ส่งเป็นข้อความแทน
}

// ---------- เลือกลบรายการของวันที่กำหนด (การ์ดมีปุ่มลบทีละรายการ) ----------

function showDeletePicker_(replyToken, dateStr) {
  const data = getAllData_();
  const list = data.expenses.filter(e => e.localDate === dateStr).sort((a, b) => a.id - b.id);
  const dateLabel = Number(dateStr.slice(8, 10)) + ' ' + monthLabelTh_(dateStr.slice(0, 7));

  if (list.length === 0) {
    replyLine_(replyToken, `วันที่ ${dateLabel} ยังไม่มีรายการให้ลบค่ะ\n(พิมพ์ "ลบของเมื่อวาน" หรือ "ลบรายการวันที่ 3" เพื่อดูวันอื่นได้)`);
    return;
  }

  const maxShow = 20;
  const shown = list.slice(-maxShow); // ถ้าเยอะเกิน แสดง 20 รายการล่าสุดของวัน
  const rows = [];
  shown.forEach((e, i) => {
    const whoLabel = e.who === WHO_SHARED ? 'ร่วมกัน' : e.who;
    if (i > 0) rows.push({ type: 'separator', margin: 'sm' });
    rows.push({
      type: 'box', layout: 'horizontal', margin: 'sm', alignItems: 'center',
      contents: [
        {
          type: 'box', layout: 'vertical', flex: 5,
          contents: [
            { type: 'text', text: `${e.category} ${fmtBaht_(e.amount)} บาท`, size: 'sm', weight: 'bold', color: '#111111', wrap: true },
            { type: 'text', text: whoLabel + (e.note ? ' · ' + String(e.note).slice(0, 40) : ''), size: 'xs', color: '#888888', wrap: true }
          ]
        },
        {
          type: 'button', flex: 2, height: 'sm', style: 'primary', color: '#C0392B',
          action: { type: 'postback', label: 'ลบ', data: 'del:' + e.id, displayText: `ลบ ${e.category} ${fmtBaht_(e.amount)} บาท (${whoLabel})` }
        }
      ]
    });
  });
  if (list.length > maxShow) {
    rows.push({ type: 'text', text: `แสดง ${maxShow} รายการล่าสุดจากทั้งหมด ${list.length} รายการ`, size: 'xxs', color: '#AAAAAA', margin: 'md' });
  }

  const flex = {
    type: 'flex',
    altText: `เลือกรายการที่จะลบ (${dateLabel}) ${list.length} รายการ`,
    contents: {
      type: 'bubble', size: 'mega',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#173C36', paddingAll: 'lg',
        contents: [
          { type: 'text', text: 'เลือกรายการที่จะลบ', size: 'xs', color: '#D9BE6E' },
          { type: 'text', text: dateLabel, size: 'lg', weight: 'bold', color: '#FFFFFF' }
        ]
      },
      body: { type: 'box', layout: 'vertical', paddingAll: 'lg', contents: rows }
    }
  };

  const ok = sendLineMessages_('reply', replyToken, [flex]);
  if (!ok) {
    // ส่งการ์ดไม่ผ่าน: ส่งเป็นข้อความพร้อมวิธีลบแบบระบุลำดับย้อนหลังแทน
    const sortedAll = [...data.expenses].sort((a, b) => b.id - a.id);
    const lines = [`รายการวันที่ ${dateLabel}:`];
    shown.forEach(e => {
      const pos = sortedAll.findIndex(x => x.id === e.id) + 1;
      const whoLabel = e.who === WHO_SHARED ? 'ร่วมกัน' : e.who;
      lines.push(`• ${e.category} ${fmtBaht_(e.amount)} (${whoLabel}) → พิมพ์ "ลบล่าสุด ${pos}"`);
    });
    replyLine_(replyToken, lines.join('\n'));
  }
}

function handlePostback_(ev, senderUserId, props) {
  const data = (ev.postback && ev.postback.data) || '';
  if (data.indexOf('del:') !== 0) return;

  if (!isAdmin_(senderUserId, props)) {
    replyLine_(ev.replyToken, 'คำสั่งแก้ไข/ลบ/ตั้งค่า ใช้ได้เฉพาะผู้ดูแลระบบค่ะ');
    return;
  }

  const id = data.slice(4);
  const target = getAllData_().expenses.find(e => String(e.id) === String(id));
  if (!target) {
    replyLine_(ev.replyToken, 'รายการนี้ถูกลบไปแล้ว หรือไม่มีอยู่ในระบบค่ะ');
    return;
  }
  deleteExpense_(target.id);
  const whoLabel = target.who === WHO_SHARED ? 'ร่วมกัน' : target.who;
  replyLine_(ev.replyToken,
    `ลบแล้วค่ะ ✅\n-${fmtBaht_(target.amount)} บาท · ${target.category} (${whoLabel})${target.note ? '\n"' + target.note + '"' : ''}\n\nพิมพ์ "ลบ" อีกครั้งถ้าต้องการลบรายการอื่น`);
}

// ---------- ค้นหารายการ ----------

function searchExpensesCommand_(parsed, senderUserId, props, replyToken) {
  const data = getAllData_();
  const monthKey = /^\d{4}-\d{2}$/.test(parsed.searchMonth || '') ? parsed.searchMonth : localDatePrefix_(new Date(), 'yyyy-MM');
  const category = (parsed.searchCategory || '').trim();
  let person = (parsed.searchPerson || '').trim();
  if (person === 'self') person = resolveWho_(senderUserId, props);

  let list = data.expenses.filter(e => (e.localDate || '').slice(0, 7) === monthKey);
  if (category) list = list.filter(e => e.category === category || String(e.category).indexOf(category) !== -1);
  if (person) list = list.filter(e => e.who === person);
  list.sort((a, b) => (a.localDate || '').localeCompare(b.localDate || '') || a.id - b.id);

  const title = (category ? category : 'ทุกหมวด') + (person ? ' ของ' + person : '') + ' · ' + monthLabelTh_(monthKey);
  if (list.length === 0) {
    replyLine_(replyToken, `ไม่พบรายการ ${title} ค่ะ`);
    return;
  }

  const lines = ['🔎 ' + title];
  const maxShow = 40;
  list.slice(0, maxShow).forEach(e => {
    const day = Number((e.localDate || '').slice(8, 10));
    const whoLabel = e.who === WHO_SHARED ? 'ร่วมกัน' : e.who;
    lines.push(`${day} · ${category ? '' : e.category + ' '}${fmtBaht_(e.amount)} (${whoLabel})${e.note ? ' ' + e.note : ''}`);
  });
  if (list.length > maxShow) lines.push(`…และอีก ${list.length - maxShow} รายการ`);
  lines.push(`รวม ${list.length} รายการ: ${fmtBaht_(list.reduce((s, e) => s + e.amount, 0))} บาท`);
  replyLine_(replyToken, lines.join('\n'));
}

// ---------- สรุปรายวันผ่าน LINE (ทุกวัน 20:00 น.) ----------

function dailyLineSummary() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const targets = getNotifyTargets_(props);
  if (!token || targets.length === 0) {
    Logger.log('ยังไม่ได้ตั้งค่า LINE_CHANNEL_ACCESS_TOKEN หรือยังไม่มี PARTNER1_LINE_ID/PARTNER2_LINE_ID/LINE_TARGET_ID เลย');
    return;
  }
  const sendTime = new Date();
  const text = buildDailySummaryText_(true);
  targets.forEach(id => pushLine_(token, id, text));
  // จำเวลาที่ส่งแจ้งเตือนครั้งนี้ไว้ เพื่อเช็คว่ารายการที่บันทึกหลังจากนี้เป็น "บันทึกหลังแจ้งเตือน" ของวันถัดไป
  props.setProperty('LAST_DAILY_NOTIFIED_AT', sendTime.toISOString());
}

function buildDailySummaryText_(includeCatchup) {
  const data = getAllData_();
  const now = new Date();
  const todayStr = localDatePrefix_(now, 'yyyy-MM-dd');

  const todayExpenses = data.expenses.filter(e => e.localDate === todayStr);

  const lines = [];
  lines.push(`📒 รายจ่ายวันนี้ (${localDatePrefix_(now, 'd MMM yyyy')})`);

  if (todayExpenses.length === 0) {
    lines.push('วันนี้ยังไม่มีรายการบันทึกไว้เลย');
  } else {
    let total = 0;
    todayExpenses.forEach(e => {
      total += e.amount;
      const whoLabel = e.who === WHO_SHARED ? 'ร่วมกัน' : e.who;
      lines.push(`• ${e.category}: ${e.amount.toLocaleString('th-TH')} บาท (${whoLabel}${e.note ? ' - ' + e.note : ''})`);
    });
    lines.push(`รวมวันนี้: ${total.toLocaleString('th-TH')} บาท`);
  }

  // รายการที่บันทึกหลังแจ้งเตือนครั้งก่อน (ปกติคือเมื่อวาน หลัง 20:00) ให้แถมมาด้วยกันลืม
  if (includeCatchup) {
    const lastNotifiedIso = PropertiesService.getScriptProperties().getProperty('LAST_DAILY_NOTIFIED_AT');
    if (lastNotifiedIso) {
      const lastNotifiedTime = new Date(lastNotifiedIso).getTime();
      const lateEntries = data.expenses.filter(e => {
        if (e.localDate === todayStr) return false; // นับไปแล้วด้านบน
        // ใช้ "เวลาที่พิมพ์จริง" (recordedAt) ไม่ใช่ Date ที่อาจถูกย้อนหลังไว้ ไม่งั้นรายการที่พิมพ์ "เมื่อวาน" จะหลุดจากทั้งสองที่
        const t = new Date(e.recordedAt || e.date).getTime();
        return !isNaN(t) && t > lastNotifiedTime;
      });
      if (lateEntries.length > 0) {
        const byDate = {};
        lateEntries.forEach(e => {
          (byDate[e.localDate] = byDate[e.localDate] || []).push(e);
        });
        lines.push('');
        lines.push('📌 บันทึกที่พลาดแจ้งเตือนครั้งก่อน (บันทึกหลัง 20:00):');
        Object.keys(byDate).sort().forEach(d => {
          const dLabel = Utilities.formatDate(new Date(d + 'T12:00:00'), getActiveTimezone_(), 'd MMM yyyy');
          lines.push(`(ของวันที่ ${dLabel})`);
          byDate[d].forEach(e => {
            const whoLabel = e.who === WHO_SHARED ? 'ร่วมกัน' : e.who;
            lines.push(`• ${e.category}: ${e.amount.toLocaleString('th-TH')} บาท (${whoLabel}${e.note ? ' - ' + e.note : ''})`);
          });
        });
      }
    }
  }

  return lines.join('\n');
}

// ---------- สรุปรายเดือนผ่าน LINE (อัตโนมัติทุกวันที่ 1 ของเดือน สรุปเดือนที่ผ่านมา) ----------

function monthlyLineSummary() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const targets = getNotifyTargets_(props);
  if (!token || targets.length === 0) {
    Logger.log('ยังไม่ได้ตั้งค่า LINE_CHANNEL_ACCESS_TOKEN หรือยังไม่มี PARTNER1_LINE_ID/PARTNER2_LINE_ID/LINE_TARGET_ID เลย');
    return;
  }
  const lastMonth = new Date();
  lastMonth.setDate(1);
  lastMonth.setMonth(lastMonth.getMonth() - 1);
  const flex = buildMonthlySummaryFlex_(lastMonth);
  targets.forEach(id => {
    const ok = sendLineMessages_('push', id, [flex]);
    if (!ok) pushLine_(token, id, buildMonthlySummaryText_(lastMonth)); // การ์ดส่งไม่ผ่าน ส่งเป็นข้อความแทน
  });
}

function buildMonthlySummaryText_(dateInMonth) {
  const data = getAllData_();
  const monthStr = localDatePrefix_(dateInMonth, 'yyyy-MM');
  const monthLabel = localDatePrefix_(dateInMonth, 'MMMM yyyy');
  const [name1, name2] = data.partners;

  const monthExpenses = data.expenses.filter(e => (e.localDate || '').slice(0, 7) === monthStr);

  const lines = [];
  lines.push(`📊 สรุปรายจ่ายเดือน ${monthLabel}`);

  if (monthExpenses.length === 0) {
    lines.push('เดือนนี้ยังไม่มีรายการบันทึกไว้เลย');
    return lines.join('\n');
  }

  const totalP1 = monthExpenses.filter(e => e.who === name1).reduce((s, e) => s + e.amount, 0);
  const totalP2 = monthExpenses.filter(e => e.who === name2).reduce((s, e) => s + e.amount, 0);
  const totalShared = monthExpenses.filter(e => e.who === WHO_SHARED).reduce((s, e) => s + e.amount, 0);
  const grandTotal = totalP1 + totalP2 + totalShared;

  lines.push('');
  lines.push(`${name1}: ${totalP1.toLocaleString('th-TH')} บาท`);
  lines.push(`${name2}: ${totalP2.toLocaleString('th-TH')} บาท`);
  lines.push(`ร่วมกัน: ${totalShared.toLocaleString('th-TH')} บาท`);
  lines.push(`รวมทั้งหมด: ${grandTotal.toLocaleString('th-TH')} บาท`);

  lines.push('');
  lines.push('แยกตามหมวด:');
  const byCategory = {};
  monthExpenses.forEach(e => {
    byCategory[e.category] = (byCategory[e.category] || 0) + e.amount;
  });
  Object.keys(byCategory)
    .sort((a, b) => byCategory[b] - byCategory[a])
    .forEach(cat => {
      lines.push(`• ${cat}: ${byCategory[cat].toLocaleString('th-TH')} บาท`);
    });

  return lines.join('\n');
}

// ---------- ตัวช่วยติดตั้ง ----------

function testSend() {
  dailyLineSummary();
}

function testSendMonthly() {
  monthlyLineSummary();
}

function testParseMessage() {
  const result = parseWithClaude_('กินข้าว 150', ['ค่ากิน', 'ค่าน้ำ', 'ค่าไฟ']);
  Logger.log(JSON.stringify(result));
}

/**
 * รันครั้งเดียวจาก Editor เพื่อ "รีเซ็ต" ล้างรายการรายจ่ายเก่าทั้งหมดออก
 * แล้วใส่รายการเริ่มต้นของวันนี้ให้ใหม่: โฟกัส ค่ากิน 70, อ๋อม ค่ากิน 125
 * (แก้ตัวเลข/ชื่อหมวดในฟังก์ชันนี้เองได้ถ้าอยากเปลี่ยนค่าเริ่มต้น)
 */
/**
 * รันครั้งเดียวจาก Editor หลังอัปเดตโค้ดเป็นเวอร์ชันที่มีคอลัมน์ "LocalDate"
 * เพื่อเติมวันที่ท้องถิ่นให้กับรายการเก่าที่บันทึกไว้ก่อนหน้านี้ (สมมติว่าบันทึกด้วยเขตเวลาไทย)
 * ไม่ทำก็ได้ถ้าไม่สนใจรายการเก่า แต่ถ้าไม่ทำ แท็บ Dashboard จะไม่เห็นรายการเก่าพวกนั้นเลย
 */
function backfillLocalDates() {
  const { expenses } = ensureSheets_(); // ตรงนี้บังคับฟอร์แมตคอลัมน์ G เป็นข้อความไว้แล้วด้วย
  const data = expenses.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; // แถวว่าง ข้าม
    const existing = data[i][6];
    const isDateObject = existing instanceof Date;
    if (existing && !isDateObject) continue; // มี LocalDate เป็นข้อความอยู่แล้ว ข้าม

    let localDate;
    if (isDateObject) {
      // แก้แถวที่ Sheets เคยแปลงเป็น Date object ไปแล้วให้กลับมาเป็นข้อความ
      localDate = Utilities.formatDate(existing, getActiveTimezone_(), 'yyyy-MM-dd');
    } else {
      const isoDate = data[i][1];
      try {
        localDate = Utilities.formatDate(new Date(isoDate), 'Asia/Bangkok', 'yyyy-MM-dd');
      } catch (e) {
        localDate = String(isoDate).slice(0, 10);
      }
    }
    expenses.getRange(i + 1, 7).setValue(localDate);
    updated++;
  }
  Logger.log('เติม/แก้ LocalDate ให้รายการเก่าเรียบร้อย: ' + updated + ' แถว');
}

function resetExpensesWithSeed() {
  const { expenses } = ensureSheets_();
  const lastRow = expenses.getLastRow();
  if (lastRow > 1) {
    expenses.deleteRows(2, lastRow - 1); // ลบทุกแถวยกเว้นหัวตาราง
  }

  const props = PropertiesService.getScriptProperties();
  const name1 = props.getProperty('PARTNER1') || 'คนที่ 1';
  const name2 = props.getProperty('PARTNER2') || 'คนที่ 2';
  const now = new Date().toISOString();

  addExpense_(name1, 'ค่ากิน', 70, '', now);
  addExpense_(name2, 'ค่ากิน', 125, '', now);

  Logger.log(`รีเซ็ตข้อมูลรายจ่ายเรียบร้อย เพิ่มรายการเริ่มต้นวันนี้: ${name1} ค่ากิน 70 บาท, ${name2} ค่ากิน 125 บาท`);
}

/** รันครั้งเดียวจาก Editor เพื่อตั้งเวลาแจ้งเตือนอัตโนมัติทุกวัน (ค่าเริ่มต้น 20:00) */
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyLineSummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyLineSummary')
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .create();
  Logger.log('ตั้งเวลาแจ้งเตือนรายวันทุกวัน 20:00 น. เรียบร้อย');
}

/** รันครั้งเดียวจาก Editor เพื่อตั้งเวลาสรุปรายเดือนอัตโนมัติ (ทุกวันที่ 1 ของเดือน เวลา 09:00) */
function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'monthlyLineSummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('monthlyLineSummary')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();
  Logger.log('ตั้งเวลาสรุปรายเดือนทุกวันที่ 1 เวลา 09:00 น. เรียบร้อย');
}

/** รันครั้งเดียวจาก Editor เพื่อสร้างแท็บ "Dashboard" สรุปรายจ่ายในตัว Google Sheet เอง (ดูสด + Export เป็น PDF ได้) */
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Dashboard');
  if (sheet) ss.deleteSheet(sheet); // สร้างใหม่ทุกครั้งที่รัน กันสูตรค้าง
  sheet = ss.insertSheet('Dashboard', 0);

  const props = PropertiesService.getScriptProperties();
  const name1 = props.getProperty('PARTNER1') || 'คนที่ 1';
  const name2 = props.getProperty('PARTNER2') || 'คนที่ 2';

  // ---------- โทนสี (ธีมเดียวกับแอป) ----------
  const COVER = '#173C36';
  const COVER_DARK = '#0F2B27';
  const GOLD = '#B4922E';
  const GOLD_LIGHT = '#D9BE6E';
  const PAPER = '#F4EEE0';
  const P1 = '#B0524A';
  const P1_SOFT = '#F2DAD6';
  const P2 = '#2F6E63';
  const P2_SOFT = '#DCEAE5';
  const SHARED_SOFT = '#F1E6C8';
  const SHARED_TEXT = '#8A6E1F';
  const INK = '#2B2A24';
  const INK_SOFT = '#6B6A5E';

  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:H200').setBackground(PAPER).setFontFamily('Prompt');

  // ---------- หัวเรื่อง ----------
  sheet.getRange('A1:E2').merge()
    .setValue('แดชบอร์ดรายจ่ายคู่รัก')
    .setBackground(COVER).setFontColor(GOLD_LIGHT)
    .setFontSize(20).setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sheet.setRowHeight(1, 32);
  sheet.setRowHeight(2, 32);
  sheet.getRange('A1:E2').setBorder(true, true, true, true, false, false, GOLD, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ---------- แถบเลือกเดือน (ดรอปดาวน์) ----------
  sheet.getRange('A4').setValue('เดือนที่ดู (YYYY-MM):')
    .setFontWeight('bold').setFontColor(INK).setVerticalAlignment('middle');
  const monthOptions = generateMonthOptions_(24, 6); // ย้อนหลัง 24 เดือน + ล่วงหน้า 6 เดือน
  const monthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(monthOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B4')
    .setNumberFormat('@')
    .setFormula('=TEXT(TODAY(),"YYYY-MM")')
    .setDataValidation(monthRule)
    .setBackground('#FFFDF5').setFontColor(COVER).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(4, 26);

  // ชื่อทั้งคู่ (ซ่อนไว้ที่คอลัมน์ H ใช้เป็นค่าอ้างอิงในสูตร ไม่ต้องแก้ตรงนี้ตรงๆ)
  sheet.getRange('H1').setValue(name1);
  sheet.getRange('H2').setValue(name2);
  sheet.hideColumns(8);

  // ---------- การ์ดสรุปยอด 4 ใบ ----------
  sheet.setRowHeight(6, 22);
  sheet.setRowHeight(7, 34);

  sheet.getRange('B6').setFormula('=$H$1');
  sheet.getRange('C6').setFormula('=$H$2');
  sheet.getRange('D6').setValue('ร่วมกัน');
  sheet.getRange('E6').setValue('รวมทั้งหมด');

  sheet.getRange('B7').setFormula('=SUMIFS(Expenses!E:E,Expenses!G:G,$B$4&"*",Expenses!C:C,$H$1)');
  sheet.getRange('C7').setFormula('=SUMIFS(Expenses!E:E,Expenses!G:G,$B$4&"*",Expenses!C:C,$H$2)');
  sheet.getRange('D7').setFormula('=SUMIFS(Expenses!E:E,Expenses!G:G,$B$4&"*",Expenses!C:C,"shared")');
  sheet.getRange('E7').setFormula('=B7+C7+D7');
  sheet.getRange('B7:E7').setNumberFormat('#,##0" บาท"').setFontSize(15).setFontWeight('bold');

  const cardStyles = [
    { range: 'B6:B7', bg: P1_SOFT, fg: P1 },
    { range: 'C6:C7', bg: P2_SOFT, fg: P2 },
    { range: 'D6:D7', bg: SHARED_SOFT, fg: SHARED_TEXT },
    { range: 'E6:E7', bg: COVER, fg: GOLD_LIGHT }
  ];
  cardStyles.forEach(c => {
    sheet.getRange(c.range)
      .setBackground(c.bg).setFontColor(c.fg)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, PAPER, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  sheet.getRange('B6:E6').setFontWeight('bold').setFontSize(11);

  // ---------- ภาระที่แท้จริง (รวมครึ่งค่าร่วม) ----------
  sheet.getRange('A8').setValue('ภาระที่แท้จริง (รวมครึ่งค่าร่วม)')
    .setFontColor(INK_SOFT).setFontSize(11);
  sheet.getRange('B8').setFormula('=B7+D7/2');
  sheet.getRange('C8').setFormula('=C7+D7/2');
  sheet.getRange('B8:C8').setNumberFormat('#,##0.00" บาท"').setFontColor(INK).setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(8, 22);

  // ---------- แยกตามหมวด ----------
  sheet.getRange('A10').setValue('แยกตามหมวด (เดือนที่เลือก)')
    .setFontWeight('bold').setFontSize(13).setFontColor(COVER);
  sheet.setRowHeight(10, 26);

  sheet.getRange('A11:C11')
    .setValues([['หมวด', 'ยอดรวม', 'สัดส่วน']])
    .setBackground(COVER).setFontColor(GOLD_LIGHT).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(11, 26);

  sheet.getRange('A12').setFormula(
    '=IFERROR(QUERY(Expenses!A:G,"select D, sum(E) where G matches \'"&$B$4&".*\' group by D order by sum(E) desc label sum(E) \'\'",1),"")'
  );
  sheet.getRange('C12').setFormula(
    '=ARRAYFORMULA(IF($B$12:$B$200="","",REPT("█",ROUND($B$12:$B$200/MAX($B$12:$B$200)*15,0))))'
  );
  sheet.getRange('A12:A200').setFontColor(INK);
  sheet.getRange('B12:B200').setNumberFormat('#,##0" บาท"').setFontColor(INK_SOFT).setHorizontalAlignment('right');
  sheet.getRange('C12:C200').setFontColor(P2).setFontWeight('bold');
  sheet.getRange('A12:C200').setBorder(false, false, true, false, false, true, '#E4DCC8', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 230);
  sheet.setColumnWidths(4, 2, 130);
  sheet.setFrozenRows(4);

  Logger.log('สร้างแท็บ Dashboard เรียบร้อย — เปิด Google Sheet แล้วไปที่แท็บ Dashboard ได้เลย');
}

/** ซิงก์ชื่อทั้งคู่เข้า Dashboard sheet (เรียกอัตโนมัติทุกครั้งที่แก้ชื่อผ่านแอป) */
function syncDashboardNames_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Dashboard');
  if (!sheet) return;
  const props = PropertiesService.getScriptProperties();
  sheet.getRange('H1').setValue(props.getProperty('PARTNER1') || 'คนที่ 1');
  sheet.getRange('H2').setValue(props.getProperty('PARTNER2') || 'คนที่ 2');
}
