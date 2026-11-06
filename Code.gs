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
 *   ADMIN_KEY                  = รหัสสำหรับเข้าหน้าแอป (Web app)
 *   BACKUP_FOLDER_ID           = (ระบบตั้งเอง) โฟลเดอร์ Google Drive ที่เก็บไฟล์สำรอง
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
  const adminKey = (PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '').trim();
  if (!adminKey) {
    result = { error: 'ยังไม่ได้ตั้ง ADMIN_KEY ใน Script Properties ระบบหลังบ้านถูกล็อกไว้' };
  } else if (String(p.key || '').trim() !== adminKey) {
    result = { error: 'รหัสไม่ถูกต้อง' };
  } else if (action === 'data') {
    result = getAllData_();
  } else if (action === 'addExpense') {
    result = addExpenseFromApp_(p) || getAllData_();
  } else if (action === 'resolvePending') {
    result = resolvePendingFromApp_(p) || getAllData_();
  } else if (action === 'deletePending') {
    deletePending_(p.id);
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
  } else if (action === 'backup') {
    try {
      backupNow_();
      result = getAllData_();
    } catch (err) {
      result = { error: 'สำรองข้อมูลไม่สำเร็จ: ' + err.message };
    }
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
    partners: [props.getProperty('PARTNER1'), props.getProperty('PARTNER2')],
    backup: getBackupStatus_(),
    pendings: getPendings_(null).map(x => ({
      id: x.id, amount: x.amount, item: x.item || '', localDate: x.localDate, who: resolveWho_(x.userId, props)
    }))
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

// ใครลบรายการไหนได้บ้าง: โฟกัสลบได้ทุกรายการ, อ๋อมลบได้เฉพาะรายการของตัวเอง
function canDeleteExpense_(expense, senderUserId, props) {
  if (isAdmin_(senderUserId, props)) return true;
  const slot = resolveSenderSlot_(senderUserId, props);
  if (!slot) return false;
  return expense.who === props.getProperty(slot);
}

// คำสั่งที่แก้ไข/ลบ/ตั้งค่าข้อมูล ใช้ได้เฉพาะผู้ดูแลระบบ
// (delete_pick ไม่อยู่ในลิสต์นี้ เพราะอ๋อมใช้ได้ด้วย แต่จะเห็นเฉพาะรายการของตัวเอง)
const ADMIN_ACTIONS_ = ['delete_last', 'fix_category', 'fix_amount', 'set_shared', 'split_half', 'delete_category', 'rename_self', 'set_timezone'];

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
  '💄 เครื่องสำอางของอ๋อม: ซื้อแล้วพิมพ์ตามปกติ เช่น "ลิป MAC 890" ระบบเพิ่มเข้าคลังให้เอง\n' +
  '• "เครื่องสำอาง" ดูคลัง · "สรุปเครื่องสำอาง" ดูยอดใช้จ่าย · "งบเครื่องสำอาง 1500" ตั้งงบรายเดือน\n' +
  '• "เปิดใช้ ลิป" · "ลิปหมดแล้ว" · "ทิ้ง มาสคาร่า" · "กันแดด หมดอายุ 05/2027" · "เพิ่มเครื่องสำอาง เซรั่ม Klairs" (ของที่มีอยู่แล้ว)\n' +
  '• ถ้าน้องออมไม่แน่ใจหมวด จะขึ้นปุ่มให้เลือก (เลือกครั้งเดียว ครั้งหน้าจำให้เอง) · ลืมกด พิมพ์ "รอเลือกหมวด" · ล้างที่ค้าง พิมพ์ "ล้างรอเลือกหมวด"\n' +
  '• "ลบ" → แสดงรายการวันนี้ให้กดปุ่มเลือกลบ (อ๋อมเห็นเฉพาะรายการของตัวเอง, "ลบของเมื่อวาน" เพื่อดูวันอื่น, "ยกเลิก" ถ้าไม่ลบแล้ว)\n' +
  'คำสั่งด้านล่างนี้ (แก้/ลบ/ตั้งค่า) ใช้ได้เฉพาะผู้ดูแลระบบ:\n' +
  '• "ลบล่าสุด" / "เอารายการที่เพิ่งจดออก" → ลบรายการล่าสุดทันที\n' +
  '• "เครดิต AI" → ดูว่าเดือนนี้ใช้ AI ไปกี่ครั้ง\n' +
  '• "สำรองข้อมูล" → สำรองข้อมูลทั้งหมดลง Google Drive ทันที (ปกติระบบสำรองให้เองทุกวันอาทิตย์ 03:00)\n' +
  '• "แก้ไขเป็นหมวดค่ากิน" / "เปลี่ยนหมวดล่าสุดเป็นค่าเที่ยว"\n' +
  '• "แก้ยอดเป็น 417.06" / "ยอดที่ถูกต้องคือ 500"\n' +
  '• "หารครึ่ง" / "คนละครึ่ง" / "ร่วมกัน" (พิมพ์เดี่ยวๆ หลังบันทึกไปแล้ว) → แยกรายการล่าสุดเป็นคนละครึ่งจริงๆ\n' +
  '• "ลบหมวดของใช้ออกไปเลย"\n' +
  '• "เปลี่ยนชื่อฉันเป็นปาล์ม"\n' +
  '• "อยู่ญี่ปุ่น" / "เปลี่ยนเขตเวลาเป็นโตเกียว" → ตั้งเขตเวลาตอนไปต่างประเทศ (พิมพ์ "กลับไทยแล้ว" เพื่อรีเซ็ตกลับ)';

function handleExpenseMessage_(ev, senderUserId, props) {
  const text = (ev.message.text || '').trim();

  // ยืมเงินตัวเอง (ส่วนตัว เห็นคนเดียว) ต้องเช็คก่อนทุกอย่าง กันไม่ให้หลุดไปบันทึกเป็นรายจ่ายปกติ
  if (handleSelfLoanMessage_(ev, text, senderUserId, props)) return;

  // กำลังรอให้พิมพ์ชื่อหมวดใหม่ (หลังกดปุ่ม "➕ หมวดใหม่")
  if (handleAwaitCategory_(ev, text, senderUserId, props)) return;

  // เครื่องสำอาง (คลัง/เปิดใช้/หมดแล้ว/สรุป/งบ)
  if (handleCosmeticsMessage_(ev, text, senderUserId, props)) return;

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
    // ใช้ได้ทั้งสองคน: โฟกัสเห็นทุกรายการ อ๋อมเห็นเฉพาะรายการของตัวเอง
    showDeletePicker_(ev.replyToken, localDatePrefix_(new Date(), 'yyyy-MM-dd'), senderUserId, props);
    return;
  }
  if (text === 'ยกเลิก' || text === 'ไม่ลบแล้ว') {
    replyLine_(ev.replyToken, 'โอเคค่ะ ไม่ลบอะไรค่ะ 👌');
    return;
  }
  if (text === 'สำรองข้อมูล' || text.toLowerCase() === 'backup') {
    if (!isAdmin_(senderUserId, props)) {
      replyLine_(ev.replyToken, '💾 ' + backupLabel_() + '\n(ระบบสำรองให้อัตโนมัติทุกวันอาทิตย์ค่ะ ส่วนการสั่งสำรองเองใช้ได้เฉพาะผู้ดูแลระบบ)');
      return;
    }
    try {
      const r = backupNow_();
      replyLine_(ev.replyToken,
        `💾 สำรองข้อมูลเสร็จแล้วค่ะ ✅\n${backupLabel_()}\n\nเก็บใน Google Drive โฟลเดอร์ "${BACKUP_FOLDER_NAME}" (เก็บ ${BACKUP_KEEP} ชุดล่าสุด)\n${r.url}`);
    } catch (err) {
      replyLine_(ev.replyToken, '🔧 (โหมดตรวจสอบชั่วคราว)\nสำรองข้อมูลไม่สำเร็จ:\n' + err);
    }
    return;
  }
  if (text === 'บันทึกรายจ่าย') {
    replyLine_(ev.replyToken, 'พิมพ์บอกรายการพร้อมจำนวนเงินได้เลยค่ะ เช่น "กินข้าว 150"');
    return;
  }
  if (/^(ล้าง|ยกเลิก)(รายการ)?รอเลือกหมวด(ทั้งหมด)?$/.test(text.replace(/\s+/g, ''))) {
    replyLine_(ev.replyToken, clearPendingsText_(clearPendings_(senderUserId)));
    return;
  }
  if (text === 'รอเลือกหมวด' || text === 'เลือกหมวด') {
    showNextPending_(ev.replyToken, senderUserId);
    return;
  }
  if (/^เครดิต\s*(ai)?$/i.test(text)) {
    replyLine_(ev.replyToken, isAdmin_(senderUserId, props) ? aiUsageText_() : 'คำสั่งนี้ใช้ได้เฉพาะผู้ดูแลระบบค่ะ');
    return;
  }

  // ข้อความแบบ "ชื่อ ตัวเลข" อ่านเองไม่ใช้เครดิต (ไม่แน่ใจหมวดจะขึ้นปุ่มให้เลือก)
  if (tryLocalExpense_(ev, text, senderUserId, props)) return;

  const data = getAllData_();
  const parsed = parseWithClaude_(text, data.categories);
  if (parsed) parsed._text = text;   // เก็บข้อความต้นฉบับไว้ตรวจซ้ำก่อนทำคำสั่งที่ลบข้อมูล

  if (!parsed && LAST_AI_ERROR_) {
    const kind = aiErrorKind_(LAST_AI_ERROR_);
    replyLine_(ev.replyToken, aiErrorUserText_(kind) +
      '\n\nลองพิมพ์แบบสั้นๆ "ชื่อ ตัวเลข" แทน เช่น "ค่าน้ำมัน 350" (แบบนี้บันทึกได้แม้ AI ใช้ไม่ได้)\nคำสั่ง สรุป · ลบ · วิธีใช้ ยังใช้ได้ตามปกติ');
    notifyAdminAiProblem_(kind);
    return;
  }

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
    replyLine_(replyToken, 'คำสั่งแก้ไข/ลบ/ตั้งค่า ใช้ได้เฉพาะผู้ดูแลระบบค่ะ ถ้าบันทึกผิด พิมพ์ "ลบ" เพื่อเลือกลบรายการของตัวเอง หรือแจ้งโฟกัสให้ช่วยแก้ได้เลยนะคะ');
    return;
  }

  // กันพลาด: คำสั่งที่ลบข้อมูลทันที ต้องมีคำที่สื่อว่า "ลบ" จริงในข้อความ
  // (เคยเกิดกรณี "ทิ้ง ลิป" แล้ว AI ตีความเป็นลบรายการล่าสุด ทำให้รายจ่ายหาย)
  const rawText = String(parsed._text || '');
  const DELETE_WORDS = /ลบ|เอา.*ออก|ยกเลิกรายการ|ไม่เอาหมวด|delete/i;
  if ((action === 'delete_last' || action === 'delete_category') && !DELETE_WORDS.test(rawText)) {
    replyLine_(replyToken, 'น้องออมไม่แน่ใจว่าต้องการลบรายการไหมค่ะ เลยยังไม่ได้ลบอะไร\nถ้าจะลบรายจ่าย พิมพ์ "ลบ" แล้วกดเลือก หรือพิมพ์ "ลบล่าสุด"\nถ้าจะเอาเครื่องสำอางออกจากคลัง พิมพ์ "ทิ้ง ชื่อของ"');
    return;
  }

  if (action === 'delete_pick') {
    const d = /^\d{4}-\d{2}-\d{2}$/.test(parsed.pickDate || '') ? parsed.pickDate : localDatePrefix_(new Date(), 'yyyy-MM-dd');
    showDeletePicker_(replyToken, d, senderUserId, props);
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
  countAiUsage_('image');
  const parsed = parseImageWithClaude_(base64Data, mimeType, data.categories);

  if (parsed && parsed.type === 'ai_error') {
    replyLine_(ev.replyToken, aiErrorUserText_(parsed.kind).replace('อ่านข้อความนี้', 'อ่านสลิปนี้') +
      '\n\nพิมพ์ยอดเองแทนได้เลยค่ะ เช่น "ค่ากิน 150"');
    notifyAdminAiProblem_(parsed.kind);
    return;
  }

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

function logParsedResult_(parsed, senderUserId, props, replyToken, opts) {
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

  // เครื่องสำอาง: ใช้ชื่อสินค้าเป็นหมายเหตุ จะได้รู้ว่าซื้ออะไรในสรุป
  if (category === COSMETIC_CATEGORY && parsed.itemName && note.indexOf(parsed.itemName) === -1) {
    note = note ? parsed.itemName + ' · ' + note : parsed.itemName;
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

  const baseText = `จดสำเร็จ ✅${dateLabel}\n-${Number(amount).toLocaleString('th-TH')} บาท · ${category} (${whoLabel})${note ? '\n"' + note + '"' : ''}\n\nเดือนนี้หมวด "${category}" ของ${whoLabel} รวม: ${categoryTotal.toLocaleString('th-TH')} บาท` + ((opts && opts.suffix) || '');

  // ซื้อเครื่องสำอาง (ของอ๋อม): เพิ่มเข้าคลังอัตโนมัติ + บอกงบ
  if (category === COSMETIC_CATEGORY && who === cosmeticOwnerName_(props)) {
    try {
      const extra = afterCosmeticPurchase_(parsed, amount, resolved.localDate, props);
      sendLineMessages_('reply', replyToken, [withQuickReply_({ type: 'text', text: baseText + extra.text }, extra.quick)]);
      return;
    } catch (e) {
      Logger.log('afterCosmeticPurchase_ error: ' + e);
    }
  }
  replyLine_(replyToken, baseText);
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
  // จำไว้ด้วย ครั้งหน้าพิมพ์ชื่อเดิมจะเข้าหมวดที่แก้ให้เอง
  if (target.note && target.note.length <= 30 && target.note.indexOf('≈') === -1) learnKeyword_(target.note, newCategory);

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
'4. ถ้าเป็นเครื่องสำอางหรือสกินแคร์ (ลิป ลิปสติก ทินท์ รองพื้น คุชชั่น แป้ง มาสคาร่า อายไลเนอร์ อายแชโดว์ บลัช คอนซีลเลอร์ ที่เขียนคิ้ว เซรั่ม โทนเนอร์ ครีมบำรุง มอยส์เจอไรเซอร์ ครีมกันแดด โฟม/คลีนซิ่งล้างหน้า มาส์กหน้า น้ำหอม) หรือซื้อจากร้านเครื่องสำอาง เช่น Eveandboy, Beautrium, Sephora, Konvy → หมวด "เครื่องสำอาง" เสมอ (สร้างหมวดนี้ถ้ายังไม่มี ห้ามใส่หมวด ของใช้)\n' +
    '5. ถ้าไม่เข้าเงื่อนไขข้างต้นเลย ให้เลือกหมวดที่ใกล้เคียงที่สุดจากลิสต์ที่มีอยู่ หรือตั้งชื่อหมวดใหม่สั้นๆที่ตรงความหมายที่สุด\n\n' +
    'ข้อความที่ได้รับมาอาจเป็นได้ 3 แบบ:\n' +
    '1) การบันทึกรายจ่ายใหม่ เช่น "กินข้าว 150", "ชานมไข่มุก 65", "ค่าเช่าบ้านร่วมกัน 3000" (splitHalf), "หารกันค่าเน็ต 600" (splitHalf), "ค่าน้ำมัน 700 หารครึ่ง" (splitHalf), "กินข้าว 150 เมื่อวาน", "ค่าน้ำมัน 300 เมื่อวานซืน", "กินข้าว 100 เมื่อ 3 วันก่อน", "กินราเมง 1200 เยน", "ค่าที่พัก 80 ดอลลาร์"\n' +
    '2) คำสั่งจัดการข้อมูล — ไม่จำเป็นต้องพิมพ์ตรงคำเป๊ะ ให้ตีความเจตนาจากความหมายทั่วไป เช่น "ลบอันนี้ที", "เอารายการล่าสุดออก" (ลบรายการล่าสุดทันที → action delete_last), "ลบ", "อยากลบรายการวันนี้", "เลือกลบ", "ลบของเมื่อวาน", "ลบรายการวันที่ 3" (อยากเลือกเองว่าจะลบรายการไหนของวันนั้น → action delete_pick), "แก้ไขเป็นหมวดค่ากิน", "เปลี่ยนหมวดล่าสุดเป็นค่าเที่ยว", "แก้ยอดเป็น 417.06", "ยอดที่ถูกต้องคือ 500 ไม่ใช่ 300", "หารครึ่ง"/"คนละครึ่ง"/"ร่วมกัน"/"หารกัน" (พิมพ์เดี่ยวๆ หลังเพิ่งบันทึกไป ความหมายเดียวกันหมด คืออยากให้รายการล่าสุดแยกเป็นคนละครึ่งจริงๆ → action split_half), "ลบสลิปนี้ออก", "ไม่เอาหมวดของใช้แล้ว", "ลบหมวดของใช้ออกไปเลย", "เปลี่ยนชื่อฉันเป็นปาล์ม", "อยู่ญี่ปุ่น", "เปลี่ยนเขตเวลาเป็นโตเกียว", "กลับไทยแล้ว" (ตั้ง/รีเซ็ตเขตเวลาตอนไปต่างประเทศ), "ขอดูสรุปหน่อย", "สรุปวันนี้ยังไง", "วิธีใช้เป็นยังไง", "ดูค่าน้ำมันเดือนนี้", "ค่ากินของอ๋อมเดือนที่แล้ว", "เดือนนี้ฉันจ่ายอะไรไปบ้าง" (ค้นหารายการ → action search) ล้วนถือเป็นคำสั่งจัดการข้อมูลทั้งสิ้น\n' +
    '3) ข้อความทั่วไปที่ไม่เกี่ยวข้องเลย เช่น คำทักทาย คำถามทั่วไป\n' +
    'ข้อห้าม: คำว่า "ทิ้ง" / "หมดแล้ว" / "หมดอายุ" เป็นคำสั่งของคลังเครื่องสำอาง ไม่ใช่การลบรายจ่าย ห้ามตอบเป็น delete_last หรือ delete_category เด็ดขาด ให้ตอบ {"type":"none"} แทน\n\n' +
    'ตอบ JSON ตามแบบใดแบบหนึ่งนี้เท่านั้น (ห้ามใส่ฟิลด์อื่นนอกเหนือจากนี้):\n\n' +
    'แบบที่ 1 (บันทึกรายจ่าย):\n' +
    '{"type":"expense","amount":ตัวเลข,"category":"ชื่อหมวดที่ใกล้เคียงที่สุดจากลิสต์ หรือชื่อหมวดใหม่สั้นๆถ้าไม่มีหมวดไหนใกล้เคียง","shared":true หรือ false,"splitHalf":true หรือ false,"daysAgo":จำนวนวันย้อนหลังจากวันนี้ (0 ถ้าไม่ได้พูดถึงเวลาหรือหมายถึงวันนี้/เมื่อกี้; "เมื่อวาน"=1; "เมื่อวานซืน"=2; "X วันก่อน"/"เมื่อ X วันก่อน"=X),"currency":"รหัสสกุลเงิน ISO 3 ตัวอักษรตัวพิมพ์ใหญ่ เช่น THB, JPY, USD, EUR, KRW (ค่าเริ่มต้นคือ THB ถ้าไม่ได้พูดถึงสกุลเงินอื่นชัดเจน เช่นมีคำว่า เยน, ดอลลาร์, ยูโร, วอน หรือสัญลักษณ์ ¥ $ € ₩ ก็ให้ใส่รหัสตามนั้น ไม่ต้องคำนวณแปลงเป็นบาทเอง ใส่ตัวเลข amount ตามสกุลเงินเดิมตรงๆ ระบบจะแปลงให้เอง)","note":"รายละเอียดสั้นๆถ้ามี ไม่งั้นเป็นค่าว่าง","itemName":"ชื่อสินค้าสั้นๆ พร้อมยี่ห้อถ้ามี (ใส่เฉพาะเมื่อหมวดเป็น เครื่องสำอาง เช่น ลิป MAC Ruby Woo หรือ กันแดด Biore ไม่งั้นเป็นค่าว่าง)"}\n' +
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
  countAiUsage_('text');
  LAST_AI_ERROR_ = null;

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    Logger.log('Claude API error: ' + res.getContentText());
    LAST_AI_ERROR_ = { code: res.getResponseCode(), text: res.getContentText().slice(0, 500) };
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
'4. ถ้าเป็นเครื่องสำอางหรือสกินแคร์ (ลิป ลิปสติก ทินท์ รองพื้น คุชชั่น แป้ง มาสคาร่า อายไลเนอร์ อายแชโดว์ บลัช คอนซีลเลอร์ ที่เขียนคิ้ว เซรั่ม โทนเนอร์ ครีมบำรุง มอยส์เจอไรเซอร์ ครีมกันแดด โฟม/คลีนซิ่งล้างหน้า มาส์กหน้า น้ำหอม) หรือซื้อจากร้านเครื่องสำอาง เช่น Eveandboy, Beautrium, Sephora, Konvy → หมวด "เครื่องสำอาง" เสมอ (สร้างหมวดนี้ถ้ายังไม่มี ห้ามใส่หมวด ของใช้)\n' +
    '5. ถ้าไม่เข้าเงื่อนไขข้างต้นเลย ให้เลือกหมวดที่ใกล้เคียงที่สุดจากลิสต์ที่มีอยู่ หรือตั้งชื่อหมวดใหม่สั้นๆที่ตรงความหมายที่สุด\n' +
    'อ่านยอดเงินรวม (ยอดที่โอน/จ่ายจริง) จากรูป แล้วตอบกลับเป็น JSON เท่านั้น ห้ามมีคำอธิบายอื่นหรือ markdown fence รูปแบบ:\n' +
    '{"type":"expense" หรือ "none","amount": ตัวเลขหรือ null,"category":"ชื่อหมวดที่ใกล้เคียงที่สุดจากลิสต์ หรือชื่อหมวดใหม่สั้นๆถ้าไม่มีหมวดไหนใกล้เคียง","shared": false,"currency":"รหัสสกุลเงิน ISO 3 ตัวอักษรตัวพิมพ์ใหญ่ตามที่ปรากฏในสลิป เช่น THB, JPY, USD, EUR, KRW (สังเกตจากสัญลักษณ์ ¥ $ € ₩ หรือชื่อสกุลเงินที่พิมพ์ไว้ ถ้าไม่ระบุเลยหรือเป็นสลิปธนาคารไทยปกติให้ใส่ THB ใส่ตัวเลข amount ตามสกุลเงินเดิมในสลิปตรงๆ ไม่ต้องคำนวณแปลงเอง)","note":"ชื่อร้าน/ผู้รับเงิน หรือรายละเอียดสั้นๆจากสลิปถ้ามี ไม่งั้นเป็นค่าว่าง","itemName":"ชื่อสินค้าสั้นๆ พร้อมยี่ห้อถ้ามี (ใส่เฉพาะเมื่อหมวดเป็น เครื่องสำอาง อ่านจากรายการสินค้าในใบเสร็จถ้ามี เช่น ลิป MAC Ruby Woo ไม่งั้นเป็นค่าว่าง)","transactionDate":"วันที่ทำรายการตามที่ปรากฏในสลิป รูปแบบ YYYY-MM-DD (ถ้าปีในสลิปเป็น พ.ศ. ให้แปลงเป็น ค.ศ. โดยลบ 543 ก่อน เช่น สลิประบุ 9 ก.ย. 69 ให้ตอบ 2026-09-09) ถ้าอ่านวันที่จากสลิปไม่ได้เลยให้เป็นค่าว่าง"}\n' +
    'ข้อสำคัญ: ให้ตอบ {"type":"none"} เฉพาะกรณีที่อ่าน "จำนวนเงิน" ในสลิปไม่ออกจริงๆ เท่านั้น (เช่น ภาพเบลอ/มืด/ไม่ใช่สลิปการเงิน) — ถ้าอ่านจำนวนเงินได้ชัดเจน ต้องตอบเป็น "expense" เสมอ ห้ามตอบ "none" เพียงเพราะไม่แน่ใจว่าเป็นหมวดอะไร (เช่น สลิปซื้อลอตเตอรี่ ค่าธรรมเนียม ค่าปรับ ของที่ไม่เข้าหมวดอาหาร/น้ำมัน/เครื่องดื่มเลย) ในกรณีแบบนี้ให้ใช้ข้อ 5 (ตั้งชื่อหมวดใหม่ที่ตรงความหมายที่สุด เช่น "หวย", "ค่าปรับ") แล้วบันทึกยอดเงินให้ถูกต้องเสมอ\n' +
    'ข้อบังคับที่สำคัญที่สุด: ตอบกลับด้วย JSON object เดียวเท่านั้นในคำตอบทั้งหมด ห้ามมีข้อความอื่นใดๆ ก่อนหรือหลัง JSON เด็ดขาด (ห้ามทักทาย ห้ามอธิบาย ห้ามขอโทษ ห้ามแก้ไขคำตอบตัวเองกลางข้อความ แม้จะสังเกตเห็นว่าตอบผิดไปแล้วก็ตาม) ให้คิดคำตอบให้รอบคอบและถูกต้องในใจก่อน แล้วค่อยพิมพ์ JSON ออกมาเพียงครั้งเดียวเท่านั้น';

  const payload = {
    model: 'claude-haiku-4-5-20251001',   // เดิมใช้ Sonnet: Haiku ถูกกว่ามากและอ่านสลิปธนาคารไทยได้ดีพอ
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
    LAST_AI_ERROR_ = { code: res.getResponseCode(), text: errText.slice(0, 500) };
    const kind = aiErrorKind_(LAST_AI_ERROR_);
    if (kind !== 'other') return { type: 'ai_error', kind: kind };
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
// โฟกัสเห็นทุกรายการ, อ๋อมเห็นเฉพาะรายการของตัวเอง

function showDeletePicker_(replyToken, dateStr, senderUserId, props) {
  const data = getAllData_();
  const isAdmin = isAdmin_(senderUserId, props);
  const slot = resolveSenderSlot_(senderUserId, props);
  const myName = slot ? props.getProperty(slot) : '';
  const list = data.expenses
    .filter(e => e.localDate === dateStr && canDeleteExpense_(e, senderUserId, props))
    .sort((a, b) => a.id - b.id);
  const dateLabel = Number(dateStr.slice(8, 10)) + ' ' + monthLabelTh_(dateStr.slice(0, 7));

  if (list.length === 0) {
    const whose = isAdmin ? '' : 'ของ' + myName;
    replyLine_(replyToken, `วันที่ ${dateLabel} ยังไม่มีรายการ${whose}ให้ลบค่ะ\n(พิมพ์ "ลบของเมื่อวาน" หรือ "ลบรายการวันที่ 3" เพื่อดูวันอื่นได้)`);
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
          { type: 'text', text: isAdmin ? 'เลือกรายการที่จะลบ' : `เลือกรายการของ${myName}ที่จะลบ`, size: 'xs', color: '#D9BE6E' },
          { type: 'text', text: dateLabel, size: 'lg', weight: 'bold', color: '#FFFFFF' }
        ]
      },
      body: { type: 'box', layout: 'vertical', paddingAll: 'lg', contents: rows },
      footer: {
        type: 'box', layout: 'vertical', paddingAll: 'md',
        contents: [{ type: 'button', height: 'sm', style: 'secondary', action: { type: 'message', label: 'ยกเลิก', text: 'ยกเลิก' } }]
      }
    }
  };

  const ok = sendLineMessages_('reply', replyToken, [flex]);
  if (!ok) {
    // ส่งการ์ดไม่ผ่าน: ส่งเป็นข้อความแทน
    const lines = [`รายการวันที่ ${dateLabel}:`];
    if (isAdmin) {
      const sortedAll = [...data.expenses].sort((a, b) => b.id - a.id);
      shown.forEach(e => {
        const pos = sortedAll.findIndex(x => x.id === e.id) + 1;
        const whoLabel = e.who === WHO_SHARED ? 'ร่วมกัน' : e.who;
        lines.push(`• ${e.category} ${fmtBaht_(e.amount)} (${whoLabel}) → พิมพ์ "ลบล่าสุด ${pos}"`);
      });
    } else {
      shown.forEach(e => lines.push(`• ${e.category} ${fmtBaht_(e.amount)}`));
      lines.push('(การ์ดแสดงไม่ได้ ลองพิมพ์ "ลบ" ใหม่อีกครั้ง หรือแจ้งโฟกัสให้ช่วยลบนะคะ)');
    }
    replyLine_(replyToken, lines.join('\n'));
  }
}

function handlePostback_(ev, senderUserId, props) {
  const data = (ev.postback && ev.postback.data) || '';
  if (data.indexOf('cos:') === 0) { handleCosmeticPostback_(ev, data, senderUserId, props); return; }
  if (/^cat(new|x|clear)?:/.test(data)) { handleCategoryPostback_(ev, data, senderUserId, props); return; }
  if (data.indexOf('del:') !== 0) return;

  const id = data.slice(4);
  const target = getAllData_().expenses.find(e => String(e.id) === String(id));
  if (!target) {
    replyLine_(ev.replyToken, 'รายการนี้ถูกลบไปแล้ว หรือไม่มีอยู่ในระบบค่ะ');
    return;
  }
  // กันกดปุ่มลบรายการของอีกคน (เช่นกดจากการ์ดที่ส่งในกลุ่ม)
  if (!canDeleteExpense_(target, senderUserId, props)) {
    replyLine_(ev.replyToken, 'ลบได้เฉพาะรายการของตัวเองค่ะ ถ้าต้องการลบรายการนี้ แจ้งโฟกัสให้ช่วยลบนะคะ');
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

  // วันสุดท้ายของเดือน: เตือนยอดยืมเงินตัวเองที่ค้าง (ส่งเฉพาะผู้ดูแลระบบคนเดียว)
  try {
    selfLoanMonthEndReminder_(false);
  } catch (e) {
    Logger.log('selfLoanMonthEndReminder_ error: ' + e);
  }

  // เครื่องสำอาง: เตือนใกล้หมดอายุ / น่าจะใกล้หมด (ส่งเฉพาะเจ้าของ)
  try {
    cosmeticsDailyCheck_();
  } catch (e) {
    Logger.log('cosmeticsDailyCheck_ error: ' + e);
  }
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

  pendingSummaryLines_().forEach(l => lines.push(l));

  lines.push('');
  lines.push('💾 ' + backupLabel_());

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

// ---------- สำรองข้อมูล (Backup) ----------
// ส่งออกทั้งชีตเป็นไฟล์ Excel (.xlsx) เก็บในโฟลเดอร์ Google Drive แยกต่างหาก
// เป็นไฟล์ข้อมูลล้วน ไม่มีโค้ดหรือรหัสลับติดไปด้วย เปิดดูได้ทั้งใน Excel และ Google Sheets

const BACKUP_FOLDER_NAME = 'สำรองข้อมูล - บันทึกรายจ่ายคู่รัก';
const BACKUP_KEEP = 8;           // เก็บไฟล์สำรองล่าสุดกี่ชุด (สัปดาห์ละชุด = ย้อนหลังราว 2 เดือน เก่ากว่านี้ย้ายลงถังขยะ)
const BACKUP_PREFIX = 'backup-รายจ่าย-';

function getBackupFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('BACKUP_FOLDER_ID');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) {
      // โฟลเดอร์ถูกลบไปแล้ว สร้างใหม่ด้านล่าง
    }
  }
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
  props.setProperty('BACKUP_FOLDER_ID', folder.getId());
  return folder;
}

// ลายนิ้วมือของข้อมูล ใช้เช็คว่ามีอะไรเปลี่ยนตั้งแต่สำรองครั้งก่อนไหม
function dataFingerprint_() {
  const { expenses, categories } = ensureSheets_();
  const raw = JSON.stringify(expenses.getDataRange().getValues()) + '|' + JSON.stringify(categories.getDataRange().getValues());
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8));
}

// สำรองข้อมูลทันที คืนค่า { url, rows, name, skipped } (โยน error ถ้าไม่สำเร็จ)
// onlyIfChanged = true → ถ้าข้อมูลไม่เปลี่ยนตั้งแต่ครั้งก่อน จะไม่สร้างไฟล์ใหม่ (ใช้กับรอบอัตโนมัติ)
function backupNow_(onlyIfChanged) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const { ss, expenses } = ensureSheets_();
    SpreadsheetApp.flush();
    const fp = dataFingerprint_();
    const props = PropertiesService.getScriptProperties();
    if (onlyIfChanged && fp === props.getProperty('LAST_BACKUP_HASH')) {
      return { skipped: true, rows: Math.max(expenses.getLastRow() - 1, 0) };
    }
    const res = UrlFetchApp.fetch(
      'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) {
      throw new Error('ส่งออกไฟล์ไม่สำเร็จ (HTTP ' + res.getResponseCode() + ')');
    }
    const now = new Date();
    const name = BACKUP_PREFIX + Utilities.formatDate(now, getActiveTimezone_(), 'yyyy-MM-dd_HHmm') + '.xlsx';
    const folder = getBackupFolder_();
    const file = folder.createFile(res.getBlob().setName(name));
    const rows = Math.max(expenses.getLastRow() - 1, 0);

    pruneOldBackups_(folder);

    props.setProperty('LAST_BACKUP_AT', now.toISOString());
    props.setProperty('LAST_BACKUP_HASH', fp);
    props.setProperty('LAST_BACKUP_INFO', JSON.stringify({ url: file.getUrl(), rows: rows, name: name }));
    writeBackupStatusToDashboard_();
    return { url: file.getUrl(), rows: rows, name: name };
  } finally {
    lock.releaseLock();
  }
}

function pruneOldBackups_(folder) {
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf(BACKUP_PREFIX) === 0) files.push(f);
  }
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  files.slice(BACKUP_KEEP).forEach(f => f.setTrashed(true));
}

function getBackupStatus_() {
  const props = PropertiesService.getScriptProperties();
  const at = props.getProperty('LAST_BACKUP_AT');
  let info = {};
  try { info = JSON.parse(props.getProperty('LAST_BACKUP_INFO') || '{}'); } catch (e) { info = {}; }
  return { at: at || '', label: backupLabel_(), url: info.url || '', rows: info.rows || 0 };
}

// ข้อความสถานะ เช่น "สำรองข้อมูลแล้ว: 6 ต.ค. 2026 03:00 (42 รายการ)"
function backupLabel_() {
  const props = PropertiesService.getScriptProperties();
  const at = props.getProperty('LAST_BACKUP_AT');
  if (!at) return 'ยังไม่เคยสำรองข้อมูล';
  let rows = '';
  try {
    const info = JSON.parse(props.getProperty('LAST_BACKUP_INFO') || '{}');
    if (info.rows !== undefined) rows = ` (${info.rows} รายการ)`;
  } catch (e) {}
  const d = new Date(at);
  const label = Number(Utilities.formatDate(d, getActiveTimezone_(), 'd')) + ' ' +
    TH_MONTHS_SHORT_[Number(Utilities.formatDate(d, getActiveTimezone_(), 'M')) - 1] + ' ' +
    Utilities.formatDate(d, getActiveTimezone_(), 'yyyy HH:mm');
  return 'สำรองข้อมูลแล้ว: ' + label + rows;
}

const TH_MONTHS_SHORT_ = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

// แสดงสถานะสำรองข้อมูลมุมขวาบนของแท็บ Dashboard ด้วย
function writeBackupStatusToDashboard_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (!sheet) return;
  sheet.getRange('D4').setValue('💾 ' + backupLabel_())
    .setFontSize(9).setFontColor('#6B6A5E').setVerticalAlignment('middle');
}

/** ทริกเกอร์สำรองข้อมูลอัตโนมัติทุกวันอาทิตย์ (ข้ามถ้าข้อมูลไม่เปลี่ยน) ถ้าไม่สำเร็จจะแจ้งโฟกัสทาง LINE */
function dailyBackup() {
  try {
    const r = backupNow_(true);
    if (r.skipped) { Logger.log('ข้อมูลไม่เปลี่ยนตั้งแต่สำรองครั้งก่อน ไม่ต้องสร้างไฟล์ใหม่'); return; }
    Logger.log('สำรองข้อมูลเรียบร้อย: ' + r.name + ' (' + r.rows + ' รายการ) ' + r.url);
  } catch (e) {
    Logger.log('สำรองข้อมูลไม่สำเร็จ: ' + e);
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    const adminId = props.getProperty('ADMIN_LINE_ID') || props.getProperty('PARTNER1_LINE_ID');
    if (token && adminId) pushLine_(token, adminId, '⚠️ สำรองข้อมูลอัตโนมัติไม่สำเร็จ\n' + e + '\nลองพิมพ์ "สำรองข้อมูล" เพื่อสั่งใหม่ได้ค่ะ');
  }
}

/** รันครั้งเดียวจาก Editor: ตั้งเวลาสำรองข้อมูลทุกวันอาทิตย์ 03:00 แล้วสำรองทันที 1 ชุด และย้ายไฟล์เก่าเกิน BACKUP_KEEP ชุดลงถังขยะ */
function setupBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBackup')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .create();
  const r = backupNow_();
  Logger.log('ตั้งเวลาสำรองข้อมูลทุกวันอาทิตย์ 03:00 เรียบร้อย และสำรองครั้งแรกแล้ว: ' + r.name + ' (' + r.rows + ' รายการ)');
  Logger.log('เปิดไฟล์: ' + r.url);
}

// ---------- ยืมเงินตัวเอง (ส่วนตัว เห็นเฉพาะผู้ดูแลระบบ) ----------
// เก็บแยกในแท็บ SelfLoans ไม่ปนกับ Expenses จึงไม่โผล่ในสรุปรายวัน/รายเดือน การค้นหา หรือ Dashboard
// ใช้ได้เฉพาะในแชทส่วนตัว 1:1 กับบอท ถ้าพิมพ์ในกลุ่ม บอทจะไม่ตอบในกลุ่มและไม่บันทึก
//   "ยืมเงินตัวเอง 2000 ซื้อของให้แฟน"   → บันทึกยอดยืม
//   "คืนเงินตัวเอง 500"                  → บันทึกการคืน (หักจากรายการที่ยืมเก่าสุดก่อน)
//   "ยอดยืม"                             → ดูยอดค้าง
//   "ลบยืมล่าสุด"                         → ลบรายการยืม/คืนล่าสุดที่บันทึกผิด

const SHEET_SELFLOANS = 'SelfLoans';

function getSelfLoanSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_SELFLOANS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SELFLOANS);
    sh.appendRow(['ID', 'LocalDate', 'Type', 'Amount', 'Note', 'RecordedAt']);
    sh.getRange(1, 2, sh.getMaxRows(), 1).setNumberFormat('@');
    sh.getRange(1, 6, sh.getMaxRows(), 1).setNumberFormat('@');
  }
  return sh;
}

function getSelfLoans_() {
  return getSelfLoanSheet_().getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => ({
      id: r[0],
      localDate: r[1] instanceof Date ? Utilities.formatDate(r[1], getActiveTimezone_(), 'yyyy-MM-dd') : String(r[1]),
      type: r[2],                       // 'borrow' หรือ 'repay'
      amount: Number(r[3]) || 0,
      note: r[4] || ''
    }));
}

// คำนวณยอดค้าง: เอายอดคืนทั้งหมดไปหักรายการยืมที่เก่าที่สุดก่อน
function selfLoanStatus_() {
  const all = getSelfLoans_().sort((a, b) => a.localDate.localeCompare(b.localDate) || a.id - b.id);
  const borrows = all.filter(x => x.type === 'borrow');
  let repaid = all.filter(x => x.type === 'repay').reduce((s, x) => s + x.amount, 0);
  const totalBorrowed = borrows.reduce((s, x) => s + x.amount, 0);
  const totalRepaid = repaid;
  const open = [];
  borrows.forEach(b => {
    const used = Math.min(repaid, b.amount);
    repaid -= used;
    const left = Math.round((b.amount - used) * 100) / 100;
    if (left > 0) open.push({ localDate: b.localDate, amount: b.amount, left: left, note: b.note });
  });
  const outstanding = Math.round((totalBorrowed - totalRepaid) * 100) / 100;
  // สรุปเฉพาะเดือนนี้ (ตามเขตเวลาที่ใช้งานอยู่)
  const monthKey = localDatePrefix_(new Date(), 'yyyy-MM');
  const inMonth = all.filter(x => (x.localDate || '').slice(0, 7) === monthKey);
  const monthBorrowed = inMonth.filter(x => x.type === 'borrow').reduce((s, x) => s + x.amount, 0);
  const monthBorrowCount = inMonth.filter(x => x.type === 'borrow').length;
  const monthRepaid = inMonth.filter(x => x.type === 'repay').reduce((s, x) => s + x.amount, 0);
  return { outstanding: outstanding, open: open, totalBorrowed: totalBorrowed, totalRepaid: totalRepaid,
    monthKey: monthKey, monthBorrowed: monthBorrowed, monthBorrowCount: monthBorrowCount, monthRepaid: monthRepaid };
}

function shortThDate_(yyyyMMdd) {
  const p = yyyyMMdd.split('-');
  return Number(p[2]) + ' ' + TH_MONTHS_SHORT_[Number(p[1]) - 1];
}

function selfLoanMonthLine_(st) {
  return `เดือน${monthLabelTh_(st.monthKey)}: ยืมไป ${fmtBaht_(st.monthBorrowed)} บาท (${st.monthBorrowCount} ครั้ง) · คืนแล้ว ${fmtBaht_(st.monthRepaid)} บาท`;
}

function selfLoanStatusText_(st) {
  if (st.outstanding <= 0) {
    return st.outstanding < 0
      ? `ไม่มียอดค้างค่ะ (คืนเกินไป ${fmtBaht_(-st.outstanding)} บาท)`
      : 'ไม่มียอดยืมเงินตัวเองค้างแล้วค่ะ 🎉';
  }
  const lines = [`ยอดค้างคืนทั้งหมด: ${fmtBaht_(st.outstanding)} บาท`];
  st.open.forEach(o => {
    const part = o.left < o.amount ? ` (เหลือ ${fmtBaht_(o.left)})` : '';
    lines.push(`• ${shortThDate_(o.localDate)} ยืม ${fmtBaht_(o.amount)}${part}${o.note ? ' · ' + o.note : ''}`);
  });
  return lines.join('\n');
}

// คืนค่า true ถ้าข้อความนี้เป็นคำสั่งยืมเงินตัวเอง (จัดการไปแล้ว ไม่ต้องทำต่อ)
function handleSelfLoanMessage_(ev, text, senderUserId, props) {
  const compact = text.replace(/\s+/g, '');
  let kind = null;
  // ประโยคคำถาม (เช่น "เดือนนี้ยืมเงินตัวเองไปเท่าไหร่") ต้องเช็คก่อน ไม่งั้นจะโดนจับเป็นคำสั่งยืม
  if (/(เท่าไหร่|เท่าไร|กี่บาท|เท่าไหร|ไปแล้ว\?|\?)$/.test(compact) && /ย[ืิึี]ม/.test(compact) && /ตัวเอง/.test(compact)) kind = 'status';
  else if (/^(วันนี้)?(ขอ)?ย[ืิึี]ม(เงิน|งิน)?ตัวเอง/.test(compact)) kind = 'borrow';   // รองรับพิมพ์ผิด ยิม/ยึม/ยีม
  else if (/^(วันนี้)?ค[ืิึี]น(เงิน|งิน)?(ให้)?ตัวเอง/.test(compact)) kind = 'repay';
  else if (/^(ดู|เช็ค|เช็ก)?ยอดย[ืิึี]ม(เงิน)?(ตัวเอง)?(เดือนนี้)?$/.test(compact) ||
           /^(เดือนนี้)?ย[ืิึี]ม(เงิน)?ตัวเอง(ไป)?(เท่าไหร่|เท่าไร|กี่บาท)/.test(compact)) kind = 'status';
  else if (/^(ลบ|ยกเลิก)(รายการ)?ย[ืิึี]ม(เงินตัวเอง)?ล่าสุด$/.test(compact)) kind = 'undo';
  if (!kind) return false;

  // คนอื่นพิมพ์ ปล่อยผ่านไปทำงานตามปกติ ไม่บอกใบ้ว่ามีฟังก์ชันนี้
  if (!isAdmin_(senderUserId, props)) return false;

  const src = ev.source || {};
  if (src.groupId || src.roomId) {
    // ห้ามตอบในกลุ่ม: แจ้งเตือนแยกเข้าแชทส่วนตัวแทน และไม่บันทึก
    const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    if (token && senderUserId) {
      pushLine_(token, senderUserId,
        '🔒 คำสั่ง "ยืมเงินตัวเอง" พิมพ์ในกลุ่มอยู่ค่ะ น้องต้มยำเลยไม่ตอบในกลุ่มและยังไม่ได้บันทึกให้\n\n' +
        'พิมพ์ใหม่ในแชทนี้ (แชทส่วนตัว) แทนนะคะ\n' +
        'ข้อความในกลุ่มกด "ยกเลิกส่ง" ได้ภายใน 24 ชม. (กดค้างที่ข้อความ → ยกเลิกส่ง)');
    }
    return true;
  }

  const sh = getSelfLoanSheet_();
  const today = localDatePrefix_(new Date(), 'yyyy-MM-dd');

  if (kind === 'status') {
    const st0 = selfLoanStatus_();
    replyLine_(ev.replyToken, '🔒 ยืมเงินตัวเอง (เห็นคนเดียว)\n' + selfLoanMonthLine_(st0) + '\n\n' + selfLoanStatusText_(st0) +
      '\n\nพิมพ์ "คืนเงินตัวเอง 500" เมื่อคืนเงิน');
    return true;
  }

  if (kind === 'undo') {
    const last = sh.getLastRow();
    if (last < 2) { replyLine_(ev.replyToken, 'ยังไม่มีรายการยืมเงินตัวเองให้ลบค่ะ'); return true; }
    const row = sh.getRange(last, 1, 1, 5).getValues()[0];
    sh.deleteRow(last);
    replyLine_(ev.replyToken,
      `🔒 ลบแล้วค่ะ ✅\n${row[2] === 'borrow' ? 'ยืม' : 'คืน'} ${fmtBaht_(row[3])} บาท${row[4] ? ' · ' + row[4] : ''}\n\n` +
      selfLoanStatusText_(selfLoanStatus_()));
    return true;
  }

  const m = text.match(/(\d[\d,]*(?:\.\d+)?)/);
  const amount = m ? Number(m[1].replace(/,/g, '')) : 0;
  if (!amount || amount <= 0) {
    replyLine_(ev.replyToken, kind === 'borrow'
      ? 'ใส่จำนวนเงินด้วยค่ะ เช่น "ยืมเงินตัวเอง 2000 ซื้อของให้แฟน"'
      : 'ใส่จำนวนเงินด้วยค่ะ เช่น "คืนเงินตัวเอง 500"');
    return true;
  }
  const note = text
    .replace(m[1], ' ')
    .replace(/^\s*(วันนี้\s*)?(ขอ\s*)?(ย[ืิึี]ม\s*(เงิน|งิน)?\s*ตัวเอง|ค[ืิึี]น\s*(เงิน|งิน)?\s*(ให้\s*)?ตัวเอง)/, ' ')
    .replace(/(^|\s)บาท(?=\s|$)/g, ' ')
    .replace(/\s+/g, ' ').trim();

  sh.appendRow([Date.now(), today, kind, amount, note, new Date().toISOString()]);
  const st = selfLoanStatus_();
  const head = kind === 'borrow'
    ? `🔒 บันทึกยืมเงินตัวเองแล้วค่ะ ✅\n-${fmtBaht_(amount)} บาท${note ? ' · ' + note : ''}`
    : `🔒 บันทึกคืนเงินตัวเองแล้วค่ะ ✅\n+${fmtBaht_(amount)} บาท${note ? ' · ' + note : ''}`;
  replyLine_(ev.replyToken, head + '\n\n' + selfLoanStatusText_(st) +
    (kind === 'borrow' ? '\n\n(ไม่นับเป็นรายจ่ายร่วม และไม่ขึ้นในสรุปที่ส่งให้อีกฝ่าย จะเตือนเฉพาะคุณทุกสิ้นเดือน)' : ''));
  return true;
}

// เรียกจาก dailyLineSummary ทุกวัน 20:00 แต่จะส่งจริงเฉพาะวันสุดท้ายของเดือน และมียอดค้าง
function selfLoanMonthEndReminder_(force) {
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const isLastDay = localDatePrefix_(new Date(now.getTime() + 86400000), 'dd') === '01';
  if (!force && !isLastDay) return;

  const st = selfLoanStatus_();
  if (!force && st.outstanding <= 0) return;

  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const adminId = props.getProperty('ADMIN_LINE_ID') || props.getProperty('PARTNER1_LINE_ID');
  if (!token || !adminId) return;

  pushLine_(token, adminId,
    `🔒 เตือนสิ้นเดือน: ยืมเงินตัวเอง (เห็นคนเดียว)\n${selfLoanMonthLine_(st)}\n\n` +
    selfLoanStatusText_(st) +
    '\n\nคืนแล้วพิมพ์ "คืนเงินตัวเอง 500" ในแชทนี้ได้เลยค่ะ');
}

/** ทดสอบ: ส่งเตือนยอดยืมเงินตัวเองเข้า LINE ของผู้ดูแลระบบทันที (ไม่ต้องรอสิ้นเดือน) */
function testSelfLoanReminder() {
  selfLoanMonthEndReminder_(true);
}

// ---------- เครื่องสำอาง (ของอ๋อม) ----------
// แท็บ Cosmetics เก็บคลังของที่มี: ซื้อเมื่อไหร่ เปิดใช้เมื่อไหร่ หมดอายุเมื่อไหร่ ใช้หมดในกี่วัน
// รายจ่ายยังอยู่ใน Expenses ตามปกติ (หมวด "เครื่องสำอาง") สรุปยอดจึงตรงกับระบบเดิม
// เจ้าของคลัง = PARTNER2 (อ๋อม) เท่านั้น: ถ้าโฟกัสบันทึกหมวดเครื่องสำอาง (เช่นซื้อเป็นของขวัญ) จะไม่เข้าคลังและไม่มีเตือนไปหาอ๋อม

const COSMETIC_CATEGORY = 'เครื่องสำอาง';
const SHEET_COSMETICS = 'Cosmetics';
const COSMETIC_HEADERS = ['ID', 'Name', 'Type', 'Price', 'BoughtDate', 'OpenedDate', 'PAOMonths', 'PrintedExpiry', 'CycleDays', 'Status', 'FinishedDate', 'ActualDays', 'Flags'];
const COSMETIC_EXPIRY_WARN_DAYS = 30;   // เตือนก่อนหมดอายุกี่วัน
const COSMETIC_RUNOUT_WARN_DAYS = 7;    // เตือนก่อนน่าจะใช้หมดกี่วัน
const COSMETIC_SNOOZE_DAYS = 14;        // กด "ยังเหลือ" แล้วเลื่อนเตือนไปกี่วัน

// อายุหลังเปิด (PAO, เดือน) และรอบใช้หมดโดยประมาณ (วัน) ตามประเภท เรียงจากคำที่เจาะจงก่อน
const COSMETIC_TYPES_ = [
  { type: 'ครีมกันแดด', kw: ['กันแดด', 'sunscreen', 'spf', 'sunblock'], pao: 12, cycle: 60 },
  { type: 'มาสคาร่า', kw: ['มาสคาร่า', 'mascara'], pao: 6, cycle: 90 },
  { type: 'อายไลเนอร์', kw: ['อายไลเนอร์', 'ไลเนอร์', 'eyeliner', 'liner'], pao: 6, cycle: 90 },
  { type: 'คิ้ว', kw: ['คิ้ว', 'brow'], pao: 12, cycle: 120 },
  { type: 'รองพื้น/คุชชั่น', kw: ['รองพื้น', 'foundation', 'คุชชั่น', 'cushion', 'บีบี', 'bb cream', 'cc cream'], pao: 12, cycle: 90 },
  { type: 'คอนซีลเลอร์', kw: ['คอนซีล', 'concealer'], pao: 12, cycle: 120 },
  { type: 'ลิป', kw: ['ลิป', 'lip', 'ทินท์', 'tint'], pao: 18, cycle: 150 },
  { type: 'แป้ง', kw: ['แป้ง', 'powder'], pao: 24, cycle: 150 },
  { type: 'อายแชโดว์/บลัช', kw: ['อายแชโดว์', 'eyeshadow', 'บลัช', 'blush', 'ปัดแก้ม', 'ไฮไลท์', 'highlight', 'คอนทัวร์', 'contour'], pao: 24, cycle: 240 },
  { type: 'คลีนเซอร์', kw: ['คลีนซิ่ง', 'cleansing', 'ล้างหน้า', 'โฟม', 'cleanser', 'ไมเซล่า', 'micellar', 'ล้างเครื่องสำอาง'], pao: 12, cycle: 60 },
  { type: 'มาส์ก', kw: ['มาส์ก', 'มาร์ก', 'mask'], pao: 12, cycle: 30 },
  { type: 'โทนเนอร์', kw: ['โทนเนอร์', 'toner'], pao: 12, cycle: 75 },
  { type: 'เซรั่ม', kw: ['เซรั่ม', 'serum', 'แอมพูล', 'ampoule', 'essence', 'เอสเซนส์'], pao: 12, cycle: 60 },
  { type: 'ครีมบำรุง', kw: ['มอยส์', 'moistur', 'ครีม', 'cream', 'เจล', 'gel', 'โลชั่น', 'lotion'], pao: 12, cycle: 60 },
  { type: 'น้ำหอม', kw: ['น้ำหอม', 'perfume', 'parfum', 'cologne', 'edt', 'edp'], pao: 36, cycle: 240 }
];
const COSMETIC_DEFAULT_TYPE_ = { type: 'อื่นๆ', pao: 12, cycle: 90 };

function cosmeticOwnerName_(props) {
  return (props || PropertiesService.getScriptProperties()).getProperty('PARTNER2') || 'คนที่ 2';
}

function getCosmeticSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_COSMETICS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_COSMETICS);
    sh.appendRow(COSMETIC_HEADERS);
    // คอลัมน์วันที่เก็บเป็นข้อความล้วน กัน Sheets แปลงเป็น Date object เอง
    [5, 6, 8, 11].forEach(c => sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'));
  }
  return sh;
}

function cosDateStr_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, getActiveTimezone_(), 'yyyy-MM-dd');
  return String(v);
}

function getCosmetics_() {
  return getCosmeticSheet_().getDataRange().getValues().slice(1)
    .map((r, i) => ({ row: i + 2, r: r }))
    .filter(x => x.r[0])
    .map(x => {
      let flags = {};
      try { flags = JSON.parse(x.r[12] || '{}'); } catch (e) { flags = {}; }
      return {
        row: x.row, id: x.r[0], name: String(x.r[1] || ''), type: String(x.r[2] || ''), price: Number(x.r[3]) || 0,
        bought: cosDateStr_(x.r[4]), opened: cosDateStr_(x.r[5]), pao: Number(x.r[6]) || 0,
        printedExpiry: cosDateStr_(x.r[7]), cycle: Number(x.r[8]) || 0, status: String(x.r[9] || 'active'),
        finished: cosDateStr_(x.r[10]), actualDays: Number(x.r[11]) || 0, flags: flags
      };
    });
}

function saveCosmetic_(it) {
  getCosmeticSheet_().getRange(it.row, 1, 1, COSMETIC_HEADERS.length).setValues([[
    it.id, it.name, it.type, it.price, it.bought, it.opened, it.pao, it.printedExpiry, it.cycle,
    it.status, it.finished, it.actualDays, JSON.stringify(it.flags || {})
  ]]);
}

// ---- วันที่ (ทำงานกับข้อความ yyyy-MM-dd ล้วน) ----
function cosToday_() { return localDatePrefix_(new Date(), 'yyyy-MM-dd'); }
function cosDayNum_(s) { const p = s.split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]) / 86400000; }
function cosDaysBetween_(a, b) { return Math.round(cosDayNum_(b) - cosDayNum_(a)); }
function cosAddDays_(s, n) { return new Date((cosDayNum_(s) + n) * 86400000).toISOString().slice(0, 10); }
function cosAddMonths_(s, n) {
  const p = s.split('-').map(Number);
  const last = new Date(Date.UTC(p[0], p[1] - 1 + n + 1, 0)).getUTCDate();
  return new Date(Date.UTC(p[0], p[1] - 1 + n, Math.min(p[2], last))).toISOString().slice(0, 10);
}
function cosThDate_(s) {
  if (!s) return '-';
  const p = s.split('-');
  const y = Number(p[0]);
  const curY = Number(cosToday_().slice(0, 4));
  return Number(p[2]) + ' ' + TH_MONTHS_SHORT_[Number(p[1]) - 1] + (y !== curY ? ' ' + String(y).slice(2) : '');
}

function cosNorm_(s) { return String(s || '').toLowerCase().replace(/\s+/g, ''); }

function detectCosmeticType_(name) {
  const n = cosNorm_(name);
  for (let i = 0; i < COSMETIC_TYPES_.length; i++) {
    const t = COSMETIC_TYPES_[i];
    if (t.kw.some(k => n.indexOf(cosNorm_(k)) !== -1)) return t;
  }
  return COSMETIC_DEFAULT_TYPE_;
}

// รอบใช้หมด: เรียนรู้จากของชิ้นเดิมที่เคยใช้หมด > ประเภทเดียวกัน > ค่าเริ่มต้นของประเภท
function estimateCycleDays_(name, typeInfo, all) {
  const done = all.filter(x => x.status === 'finished' && x.actualDays > 0);
  const avg = list => Math.round(list.reduce((s, x) => s + x.actualDays, 0) / list.length);
  const same = done.filter(x => cosNorm_(x.name) === cosNorm_(name));
  if (same.length) return { days: avg(same), source: 'same' };
  const sameType = done.filter(x => x.type === typeInfo.type);
  if (sameType.length) return { days: avg(sameType), source: 'type' };
  return { days: typeInfo.cycle, source: 'default' };
}

function cosExpiry_(it) {
  const byPao = it.opened && it.pao ? cosAddMonths_(it.opened, it.pao) : '';
  const cands = [byPao, it.printedExpiry].filter(Boolean).sort();
  return cands.length ? cands[0] : '';
}
function cosRunOut_(it) { return it.opened && it.cycle ? cosAddDays_(it.opened, it.cycle) : ''; }

function addCosmeticItem_(name, price, boughtDate, opened, printedExpiry) {
  const all = getCosmetics_();
  const t = detectCosmeticType_(name);
  const est = estimateCycleDays_(name, t, all);
  const sh = getCosmeticSheet_();
  // ID ต้องไม่ซ้ำ (ปุ่มกดอ้างถึงของด้วย ID) จึงใช้ค่าที่มากกว่า ID เดิมทุกตัวเสมอ
  const id = Math.max(Date.now(), all.reduce((mx, x) => Math.max(mx, Number(x.id) || 0), 0) + 1);
  sh.appendRow([id, name, t.type, price || 0, boughtDate, opened ? boughtDate : '', t.pao, printedExpiry || '', est.days, 'active', '', '', '{}']);
  const item = getCosmetics_().find(x => String(x.id) === String(id));
  // ซื้อชิ้นใหม่ของเดิม: ชิ้นเก่าที่ยังอยู่ในคลังไม่ต้องเตือนให้ซื้อซ้ำอีก
  all.filter(x => x.status === 'active' && cosNorm_(x.name) === cosNorm_(name)).forEach(x => {
    x.flags.runout = true; saveCosmetic_(x);
  });
  const prev = all.filter(x => cosNorm_(x.name) === cosNorm_(name) && x.price > 0)
    .sort((a, b) => (b.bought || '').localeCompare(a.bought || ''))[0];
  return { item: item, est: est, prev: prev };
}

function cosmeticMonthSpend_(owner, monthKey) {
  return getAllData_().expenses
    .filter(e => e.category === COSMETIC_CATEGORY && e.who === owner && (e.localDate || '').slice(0, 7) === monthKey);
}

function cosmeticBudgetLine_(owner, monthKey, props) {
  const budget = Number((props || PropertiesService.getScriptProperties()).getProperty('COSMETIC_BUDGET')) || 0;
  const spent = cosmeticMonthSpend_(owner, monthKey).reduce((s, e) => s + e.amount, 0);
  if (!budget) return { text: `เดือนนี้ใช้เครื่องสำอางไป ${fmtBaht_(spent)} บาท`, spent: spent, budget: 0 };
  const pct = Math.round(spent / budget * 100);
  let icon = '🟢';
  if (pct >= 100) icon = '🔴';
  else if (pct >= 80) icon = '🟡';
  const tail = spent > budget ? `เกินงบ ${fmtBaht_(spent - budget)} บาท` : `เหลือ ${fmtBaht_(budget - spent)} บาท`;
  return { text: `${icon} งบเครื่องสำอางเดือนนี้: ${fmtBaht_(spent)} / ${fmtBaht_(budget)} บาท (${pct}%) · ${tail}`, spent: spent, budget: budget };
}

// เรียกหลังบันทึกรายจ่ายหมวดเครื่องสำอางของอ๋อม คืน { text, quick }
function afterCosmeticPurchase_(parsed, amount, localDate, props) {
  const owner = cosmeticOwnerName_(props);
  const name = String(parsed.itemName || parsed.note || COSMETIC_CATEGORY).replace(/\s*\(.*?บาท\)\s*$/, '').trim().slice(0, 60);
  const r = addCosmeticItem_(name, amount, localDate, true, '');
  const it = r.item;
  const lines = [];
  lines.push(`💄 เพิ่มเข้าคลังแล้ว: ${it.name} (${it.type})`);
  const exp = cosExpiry_(it);
  if (exp) lines.push(`• เปิดแล้วใช้ได้ถึง ~${cosThDate_(exp)} (${it.pao} เดือนหลังเปิด)`);
  const src = r.est.source === 'same' ? 'จากรอบที่แล้ว' : (r.est.source === 'type' ? 'จากของประเภทเดียวกัน' : 'ค่าประมาณ');
  lines.push(`• น่าจะใช้หมดราว ${cosThDate_(cosRunOut_(it))} (~${it.cycle} วัน ${src})`);
  if (r.prev) {
    const diff = Math.round((amount - r.prev.price) * 100) / 100;
    const cmp = diff === 0 ? 'ราคาเท่าเดิม' : (diff > 0 ? `แพงขึ้น ${fmtBaht_(diff)}` : `ถูกลง ${fmtBaht_(-diff)}`);
    lines.push(`• ครั้งก่อน ${fmtBaht_(r.prev.price)} บาท (${cosThDate_(r.prev.bought)}) · ${cmp}`);
  }
  lines.push(cosmeticBudgetLine_(owner, localDate.slice(0, 7), props).text);
  return {
    text: '\n\n' + lines.join('\n'),
    quick: [
      { label: 'ยังไม่ได้เปิดใช้', data: 'cos:unopen:' + it.id, display: 'ยังไม่ได้เปิดใช้ ' + it.name },
      { label: 'ใส่วันหมดอายุ', text: it.name + ' หมดอายุ ' }
    ]
  };
}

// ---- ข้อความ ----
function withQuickReply_(msg, items) {
  if (!items || !items.length) return msg;
  msg.quickReply = {
    items: items.slice(0, 13).map(q => ({
      type: 'action',
      action: q.data
        ? { type: 'postback', label: q.label.slice(0, 20), data: q.data, displayText: (q.display || q.label).slice(0, 300) }
        : { type: 'message', label: q.label.slice(0, 20), text: q.text }
    }))
  };
  return msg;
}

function cosStatusOf_(it, today) {
  const exp = cosExpiry_(it);
  const run = cosRunOut_(it);
  return {
    exp: exp, run: run,
    expLeft: exp ? cosDaysBetween_(today, exp) : null,
    runLeft: run ? cosDaysBetween_(today, run) : null
  };
}

function buildCosmeticListText_(owner) {
  const today = cosToday_();
  const items = getCosmetics_().filter(x => x.status === 'active');
  if (!items.length) {
    return `💄 คลังเครื่องสำอางของ${owner} ยังว่างอยู่ค่ะ\n\nซื้อของใหม่พิมพ์ตามปกติ เช่น "ลิป MAC 890" ระบบจะเพิ่มเข้าคลังให้เอง\nของที่มีอยู่แล้ว พิมพ์ "เพิ่มเครื่องสำอาง เซรั่ม Klairs" (ไม่นับเป็นรายจ่าย)`;
  }
  const rows = items.map(it => ({ it: it, st: cosStatusOf_(it, today) }));
  const score = x => Math.min(x.st.expLeft === null ? 9999 : x.st.expLeft, x.st.runLeft === null ? 9999 : x.st.runLeft);
  rows.sort((a, b) => score(a) - score(b));

  const lines = [`💄 คลังเครื่องสำอางของ${owner} (${items.length} ชิ้น)`];
  rows.forEach(({ it, st }) => {
    let mark = '•';
    if (st.expLeft !== null && st.expLeft <= 0) mark = '⚠️';
    else if (st.expLeft !== null && st.expLeft <= COSMETIC_EXPIRY_WARN_DAYS) mark = '⏳';
    else if (st.runLeft !== null && st.runLeft <= COSMETIC_RUNOUT_WARN_DAYS) mark = '🛒';
    const parts = [];
    parts.push(it.opened ? 'เปิด ' + cosThDate_(it.opened) : 'ยังไม่เปิด');
    if (st.exp) parts.push(st.expLeft <= 0 ? 'หมดอายุแล้ว' : `หมดอายุ ${cosThDate_(st.exp)}${st.expLeft <= 60 ? ' (อีก ' + st.expLeft + ' วัน)' : ''}`);
    if (st.run) parts.push(st.runLeft <= 0 ? 'น่าจะใกล้หมดแล้ว' : `น่าจะหมด ~${cosThDate_(st.run)}`);
    lines.push(`${mark} ${it.name}\n   ${parts.join(' · ')}`);
  });
  lines.push('');
  lines.push('⚠️ หมดอายุแล้ว · ⏳ ใกล้หมดอายุ · 🛒 น่าจะใกล้หมด');
  lines.push('พิมพ์ "เปิดใช้ ชื่อ" · "ชื่อ หมดแล้ว" · "ทิ้ง ชื่อ"');
  return lines.join('\n');
}

function buildCosmeticSummaryText_(owner, props) {
  const monthKey = localDatePrefix_(new Date(), 'yyyy-MM');
  const prevKey = prevMonthKey_(monthKey);
  const cur = cosmeticMonthSpend_(owner, monthKey).sort((a, b) => (a.localDate || '').localeCompare(b.localDate || ''));
  const prev = cosmeticMonthSpend_(owner, prevKey);
  const sum = l => l.reduce((s, e) => s + e.amount, 0);
  const ch = changeLabel_(sum(cur), sum(prev));

  const lines = [`💄 เครื่องสำอางของ${owner} · ${monthLabelTh_(monthKey)}`];
  lines.push(`ใช้ไป ${fmtBaht_(sum(cur))} บาท (${cur.length} รายการ)` + (sum(prev) ? ` · ${ch.text} จากเดือนก่อน (${fmtBaht_(sum(prev))})` : ''));
  const b = cosmeticBudgetLine_(owner, monthKey, props);
  if (b.budget) lines.push(b.text);
  if (cur.length) {
    lines.push('');
    cur.slice(-15).forEach(e => lines.push(`• ${cosThDate_(e.localDate)} ${e.note ? e.note : COSMETIC_CATEGORY} ${fmtBaht_(e.amount)}`));
  }

  // ย้อนหลัง 6 เดือน
  const all = getAllData_().expenses.filter(e => e.category === COSMETIC_CATEGORY && e.who === owner);
  const months = [];
  let k = monthKey;
  for (let i = 0; i < 6; i++) { months.unshift(k); k = prevMonthKey_(k); }
  const trend = months.map(m => {
    const total = all.filter(e => (e.localDate || '').slice(0, 7) === m).reduce((s, e) => s + e.amount, 0);
    return TH_MONTHS_SHORT_[Number(m.slice(5, 7)) - 1] + ' ' + fmtBaht_(total);
  });
  lines.push('');
  lines.push('6 เดือนล่าสุด: ' + trend.join(' · '));

  const today = cosToday_();
  const active = getCosmetics_().filter(x => x.status === 'active').map(it => cosStatusOf_(it, today));
  const nearExp = active.filter(x => x.expLeft !== null && x.expLeft <= COSMETIC_EXPIRY_WARN_DAYS).length;
  const nearRun = active.filter(x => x.runLeft !== null && x.runLeft <= COSMETIC_RUNOUT_WARN_DAYS).length;
  lines.push(`คลัง: ${active.length} ชิ้น · ใกล้/หมดอายุ ${nearExp} · น่าจะใกล้หมด ${nearRun} (พิมพ์ "เครื่องสำอาง" ดูรายละเอียด)`);
  return lines.join('\n');
}

// ---- หาของจากชื่อที่พิมพ์ ----
function findCosmetics_(query) {
  const q = cosNorm_(query);
  if (!q) return [];
  const items = getCosmetics_().filter(x => x.status === 'active');
  const exact = items.filter(x => cosNorm_(x.name) === q);
  if (exact.length) return exact;
  const contains = items.filter(x => cosNorm_(x.name).indexOf(q) !== -1 || q.indexOf(cosNorm_(x.name)) !== -1);
  if (contains.length) return contains;
  const qt = detectCosmeticType_(query).type;
  if (qt === COSMETIC_DEFAULT_TYPE_.type) return [];
  return items.filter(x => x.type === qt);
}

function replyPickCosmetic_(replyToken, matches, action, verb) {
  const msg = withQuickReply_(
    { type: 'text', text: `มี ${matches.length} ชิ้นที่ตรงกันค่ะ กดเลือกชิ้นที่จะ${verb}:\n` + matches.map(m => `• ${m.name}${m.opened ? ' (เปิด ' + cosThDate_(m.opened) + ')' : ''}`).join('\n') },
    matches.map(m => ({ label: m.name, data: `cos:${action}:${m.id}`, display: `${verb} ${m.name}` }))
  );
  sendLineMessages_('reply', replyToken, [msg]);
}

function parseExpiryInput_(mm, yy) {
  let y = Number(yy);
  if (y < 100) y += 2000;
  if (y > 2400) y -= 543;   // ปี พ.ศ.
  const m = Number(mm);
  if (!(m >= 1 && m <= 12) || y < 2000) return '';
  const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + last).slice(-2);
}

// ---- การทำงานกับของ 1 ชิ้น (ใช้ร่วมกันทั้งพิมพ์และกดปุ่ม) ----
function cosmeticAction_(action, it, extra) {
  const today = cosToday_();
  if (action === 'open') {
    it.opened = today; it.flags = {}; saveCosmetic_(it);
    return `เปิดใช้ ${it.name} แล้วค่ะ ✅\n• ใช้ได้ถึง ~${cosThDate_(cosExpiry_(it))}\n• น่าจะหมดราว ${cosThDate_(cosRunOut_(it))} (~${it.cycle} วัน)`;
  }
  if (action === 'unopen') {
    it.opened = ''; it.flags = {}; saveCosmetic_(it);
    return `โอเคค่ะ ${it.name} ยังไม่เปิดใช้ ✅\nตอนเริ่มใช้ พิมพ์ "เปิดใช้ ${it.name}" ระบบจะเริ่มนับวันหมดอายุและรอบใช้หมดให้`;
  }
  if (action === 'done' || action === 'trash') {
    it.status = action === 'done' ? 'finished' : 'discarded';
    it.finished = today;
    if (action === 'done' && it.opened) it.actualDays = Math.max(cosDaysBetween_(it.opened, today), 1);
    saveCosmetic_(it);
    if (action === 'trash') return `ทิ้ง ${it.name} ออกจากคลังแล้วค่ะ 🗑️`;
    const lines = [`บันทึกว่า ${it.name} หมดแล้วค่ะ ✅`];
    if (it.actualDays) lines.push(`ใช้ไป ${it.actualDays} วัน (ครั้งหน้าจะคาดรอบตามนี้)`);
    if (it.price) lines.push(`ซื้อครั้งก่อน ${fmtBaht_(it.price)} บาท เมื่อ ${cosThDate_(it.bought)}`);
    lines.push(`ซื้อใหม่แล้วพิมพ์ เช่น "${it.name} ${it.price ? Math.round(it.price) : 390}" ได้เลยค่ะ`);
    return lines.join('\n');
  }
  if (action === 'snooze') {
    it.flags.snoozeUntil = cosAddDays_(today, COSMETIC_SNOOZE_DAYS);
    it.flags.runout = false;
    saveCosmetic_(it);
    return `โอเคค่ะ จะถามเรื่อง ${it.name} ใหม่อีก ${COSMETIC_SNOOZE_DAYS} วัน 👌`;
  }
  if (action === 'expiry') {
    it.printedExpiry = extra; it.flags.exp30 = false; it.flags.expired = false; saveCosmetic_(it);
    const exp = cosExpiry_(it);
    return `ตั้งวันหมดอายุ ${it.name} แล้วค่ะ ✅\nใช้ได้ถึง ${cosThDate_(exp)}${exp !== extra ? ' (นับจากวันเปิด ซึ่งมาก่อนวันที่พิมพ์บนกล่อง)' : ''}`;
  }
  return '';
}

// คืนค่า true ถ้าเป็นคำสั่งเครื่องสำอาง (จัดการแล้ว)
function handleCosmeticsMessage_(ev, text, senderUserId, props) {
  const t = text.trim();
  const c = cosNorm_(t).replace(/สํา/g, 'สำ');
  const owner = cosmeticOwnerName_(props);
  const reply = msg => replyLine_(ev.replyToken, msg);

  if (/^(คลัง|ดู)?เครื่องสำอาง(ของฉัน|ที่มี|ทั้งหมด)?$/.test(c)) { reply(buildCosmeticListText_(owner)); return true; }
  if (/^สรุป(ค่า|ยอด)?เครื่องสำอาง/.test(c)) { reply(buildCosmeticSummaryText_(owner, props)); return true; }

  let m = c.match(/^(ตั้ง)?งบ(ค่า)?เครื่องสำอาง(เดือนละ)?([\d,\.]+)(บาท)?$/);
  if (m) {
    const v = Number(m[4].replace(/,/g, '')) || 0;
    if (v > 0) props.setProperty('COSMETIC_BUDGET', String(v)); else props.deleteProperty('COSMETIC_BUDGET');
    reply(v > 0
      ? `ตั้งงบเครื่องสำอางเดือนละ ${fmtBaht_(v)} บาทแล้วค่ะ ✅\n${cosmeticBudgetLine_(owner, localDatePrefix_(new Date(), 'yyyy-MM'), props).text}\n(ตอนบันทึกรายการ ถ้าใช้ถึง 80% จะขึ้น 🟡 เกินงบขึ้น 🔴)`
      : 'ยกเลิกงบเครื่องสำอางแล้วค่ะ');
    return true;
  }

  m = t.match(/^เพิ่มเครื่องส[ำํา]+อาง\s*(.*)$/);
  if (m) {
    let name = m[1].trim(), printed = '';
    const em = name.match(/\s*หมดอายุ\s*(\d{1,2})\s*[\/\-.]\s*(\d{2,4})\s*$/);
    if (em) { printed = parseExpiryInput_(em[1], em[2]); name = name.slice(0, em.index).trim(); }
    if (!name) { reply('💄 พิมพ์ชื่อของต่อท้ายด้วยค่ะ เช่น\n"เพิ่มเครื่องสำอาง ลิป Romand"\n"เพิ่มเครื่องสำอาง กันแดด Biore หมดอายุ 05/2027"\n\n(ใช้กับของที่มีอยู่แล้ว ไม่นับเป็นรายจ่าย ถ้าเพิ่งซื้อ พิมพ์ "กันแดด Biore 390" ตามปกติได้เลย)'); return true; }
    const r = addCosmeticItem_(name, 0, cosToday_(), true, printed);
    const msg = withQuickReply_({ type: 'text', text:
      `💄 เพิ่ม ${r.item.name} (${r.item.type}) เข้าคลังแล้วค่ะ ✅ (ไม่นับเป็นรายจ่าย)\n• ถือว่าเปิดใช้วันนี้ ใช้ได้ถึง ~${cosThDate_(cosExpiry_(r.item))}\n• น่าจะหมดราว ${cosThDate_(cosRunOut_(r.item))}` },
      [{ label: 'ยังไม่ได้เปิดใช้', data: 'cos:unopen:' + r.item.id, display: 'ยังไม่ได้เปิดใช้ ' + r.item.name },
       { label: 'ใส่วันหมดอายุ', text: r.item.name + ' หมดอายุ ' }]);
    sendLineMessages_('reply', ev.replyToken, [msg]);
    return true;
  }

  // คำสั่งที่อ้างถึงของ 1 ชิ้น: strict = ถ้าไม่เจอของในคลัง ให้ปล่อยไปทำงานตามปกติ (เช่น "ข้าวหมดแล้ว")
  let action = null, query = '', extra = '', strict = true;
  if ((m = t.match(/^เปิดใช้\s*(.+)$/))) { action = 'open'; query = m[1]; strict = false; }
  else if ((m = t.match(/^ยังไม่(ได้)?เปิด(ใช้)?\s*(.+)$/))) { action = 'unopen'; query = m[3]; strict = false; }
  else if ((m = t.match(/^(.+?)\s*(ใช้)?หมดแล้ว$/))) {
    action = 'done'; query = m[1];
    strict = detectCosmeticType_(query).type === COSMETIC_DEFAULT_TYPE_.type;   // ชื่อเหมือนเครื่องสำอาง (ลิป ครีม ฯลฯ) ไม่ปล่อยไปให้ AI
  }
  else if ((m = t.match(/^ทิ้ง\s*(.+)$/))) { action = 'trash'; query = m[1]; strict = false; }
  else if ((m = t.match(/^(.+?)\s*(หมดอายุ|exp)\s*(\d{1,2})\s*[\/\-.]\s*(\d{2,4})$/i))) {
    action = 'expiry'; query = m[1]; strict = false; extra = parseExpiryInput_(m[3], m[4]);
    if (!extra) { reply('รูปแบบวันหมดอายุไม่ถูกค่ะ พิมพ์ เช่น "กันแดด หมดอายุ 05/2027"'); return true; }
  }
  if (!action) return false;

  const matches = findCosmetics_(query);
  if (!matches.length) {
    if (strict) return false;
    const names = getCosmetics_().filter(x => x.status === 'active').map(x => '• ' + x.name);
    reply(`ไม่พบ "${query}" ในคลังเครื่องสำอางค่ะ (ยังไม่ได้ลบอะไร)` +
      (names.length ? '\n\nของในคลังตอนนี้:\n' + names.slice(0, 20).join('\n') : '\n\nคลังยังว่างอยู่ค่ะ'));
    return true;
  }
  if (matches.length > 1) {
    if (action === 'expiry') {
      const dt = t.match(/(\d{1,2}\s*[\/\-.]\s*\d{2,4})$/)[1];
      reply('มีหลายชิ้นที่ชื่อคล้ายกันค่ะ พิมพ์ชื่อให้ชัดขึ้น เช่น:\n' + matches.map(x => `• ${x.name} หมดอายุ ${dt}`).join('\n'));
      return true;
    }
    const verbs = { open: 'เปิดใช้', unopen: 'ตั้งเป็นยังไม่เปิด', done: 'บันทึกว่าหมด', trash: 'ทิ้ง' };
    replyPickCosmetic_(ev.replyToken, matches, action, verbs[action]);
    return true;
  }
  reply(cosmeticAction_(action, matches[0], extra));
  return true;
}

function handleCosmeticPostback_(ev, data, senderUserId, props) {
  const parts = data.split(':');           // cos:<action>:<id>
  const action = parts[1], id = parts[2];
  if (['open', 'unopen', 'done', 'trash', 'snooze'].indexOf(action) === -1) return;
  const it = getCosmetics_().find(x => String(x.id) === String(id));
  if (!it) { replyLine_(ev.replyToken, 'ไม่พบรายการนี้ในคลังแล้วค่ะ'); return; }
  if (it.status !== 'active') {
    replyLine_(ev.replyToken, `${it.name} ถูกบันทึกว่า${it.status === 'finished' ? 'หมดแล้ว' : 'ทิ้งแล้ว'}ไปก่อนหน้านี้ค่ะ`);
    return;
  }
  replyLine_(ev.replyToken, cosmeticAction_(action, it, ''));
}

// ---- เตือนอัตโนมัติ (เรียกจาก dailyLineSummary ทุกวัน 20:00) ส่งเฉพาะอ๋อม รวมเป็นข้อความเดียว ----
// แต่ละเรื่องเตือนครั้งเดียว: ใกล้หมดอายุ (30 วันก่อน), หมดอายุแล้ว, น่าจะใกล้หมด (7 วันก่อนครบรอบ)
function cosmeticsDailyCheck_() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const to = props.getProperty('PARTNER2_LINE_ID');
  if (!token || !to) return;
  const today = cosToday_();
  const lines = [];
  const quick = [];

  getCosmetics_().filter(x => x.status === 'active').forEach(it => {
    const st = cosStatusOf_(it, today);
    let changed = false;
    if (st.expLeft !== null && st.expLeft <= 0 && !it.flags.expired) {
      lines.push(`⚠️ ${it.name} หมดอายุแล้ว (${cosThDate_(st.exp)}) ควรเลิกใช้ค่ะ`);
      quick.push({ kind: 'trash', name: it.name, label: 'ทิ้งเลย', data: 'cos:trash:' + it.id, display: 'ทิ้ง ' + it.name });
      it.flags.expired = true; it.flags.exp30 = true; changed = true;
    } else if (st.expLeft !== null && st.expLeft > 0 && st.expLeft <= COSMETIC_EXPIRY_WARN_DAYS && !it.flags.exp30) {
      lines.push(`⏳ ${it.name} จะหมดอายุ ${cosThDate_(st.exp)} (อีก ${st.expLeft} วัน)`);
      it.flags.exp30 = true; changed = true;
    }
    const snoozed = it.flags.snoozeUntil && today < it.flags.snoozeUntil;
    if (!it.flags.expired && st.runLeft !== null && st.runLeft <= COSMETIC_RUNOUT_WARN_DAYS && !it.flags.runout && !snoozed) {
      const used = cosDaysBetween_(it.opened, today);
      lines.push(`🛒 ${it.name} ใช้มา ${used} วันแล้ว น่าจะใกล้หมด` + (it.price ? ` · ครั้งก่อน ${fmtBaht_(it.price)} บาท` : ''));
      quick.push({ kind: 'run', name: it.name, label: 'หมดแล้ว', data: 'cos:done:' + it.id, display: it.name + ' หมดแล้ว' });
      quick.push({ kind: 'run', name: it.name, label: 'ยังเหลือ', data: 'cos:snooze:' + it.id, display: it.name + ' ยังเหลือ' });
      it.flags.runout = true; changed = true;
    }
    if (changed) saveCosmetic_(it);
  });

  if (!lines.length) return;
  // ปุ่ม LINE ยาวได้ไม่เกิน 20 ตัวอักษร: ถ้ามีของใกล้หมดชิ้นเดียว ใช้แค่ "หมดแล้ว"/"ยังเหลือ" ถ้าหลายชิ้นค่อยใส่ชื่อย่อ
  const runCount = quick.filter(q => q.kind === 'run').length / 2;
  if (runCount > 1) quick.filter(q => q.kind === 'run').forEach(q => { q.label = (q.label === 'หมดแล้ว' ? 'หมด: ' : 'เหลือ: ') + q.name; });
  if (quick.filter(q => q.kind === 'trash').length > 1) quick.filter(q => q.kind === 'trash').forEach(q => { q.label = 'ทิ้ง: ' + q.name; });
  const msg = withQuickReply_({ type: 'text', text: '💄 เตือนเครื่องสำอาง\n\n' + lines.join('\n') + '\n\nพิมพ์ "เครื่องสำอาง" เพื่อดูคลังทั้งหมด' }, quick);
  const ok = sendLineMessages_('push', to, [msg]);
  if (!ok) pushLine_(token, to, msg.text);
}

/** ทดสอบ: เช็คเครื่องสำอางแล้วส่งเตือนเข้า LINE อ๋อมทันที (เฉพาะรายการที่ถึงเกณฑ์และยังไม่เคยเตือน) */
function testCosmeticsCheck() {
  cosmeticsDailyCheck_();
}

// ---------- อ่านข้อความง่ายเอง + ปุ่มเลือกหมวด (ประหยัดเครดิต AI) ----------
// ข้อความแบบ "ชื่อ ตัวเลข" เช่น "ข้าว 50", "ค่าน้ำ 40", "ล้างรถ 60" ระบบอ่านเองไม่ต้องเรียก AI
// หาหมวดตามลำดับ: พิมพ์ชื่อหมวดมาตรงๆ > คำที่เคยเลือกไว้ (แท็บ Keywords) > รายการเก่าในชีต > คำพื้นฐาน
// ถ้ายังไม่แน่ใจ → ขึ้นปุ่มให้คนที่พิมพ์กดเลือก แล้วจำคำนั้นไว้ ครั้งหน้าบันทึกเองได้
// ข้อความที่ซับซ้อน (หารครึ่ง ย้อนวัน เงินต่างประเทศ คำสั่งแก้ไข ฯลฯ) ยังส่งให้ AI เหมือนเดิม

const SHEET_KEYWORDS = 'Keywords';
const PENDING_PREFIX_ = 'PENDING_';
const AWAIT_CAT_PREFIX_ = 'AWAIT_CAT_';
const AWAIT_CAT_MINUTES_ = 15;

// คำที่แปลว่าข้อความนี้ซับซ้อน ต้องให้ AI อ่าน
const COMPLEX_WORDS_ = /(^|\s)หาร|หารกัน|หารครึ่ง|ร่วมกัน|ด้วยกัน|ครึ่ง|แบ่ง|เมื่อวาน|วันก่อน|ที่แล้ว|เมื่อ\s*\d|วันที่|เดือน|(^|\s)แก้(?!ว)|(^|\s)ลบ|เปลี่ยน|สรุป|ค้นหา|ยอดที่|ย้อนหลัง|เขตเวลา|(^|\s)ตั้ง|(ไหม|มั้ย|เท่าไหร่|เท่าไร|กี่บาท)\s*\??$|\?|เยน|ดอลลาร์|ดอล|ยูโร|วอน|หยวน|ริงกิต|ปอนด์|ด่ง|กีบ|เปโซ|usd|jpy|eur|krw|cny|myr|gbp|sgd|\$|¥|€|₩|£/i;

// คำพื้นฐาน → หมวด (ใช้เมื่อยังไม่เคยเรียนรู้คำนั้น) เรียงคำเจาะจงก่อน
const BUILTIN_CATEGORY_WORDS_ = [
  { cat: 'ค่าน้ำมัน', words: ['น้ำมัน', 'ปตท', 'ptt', 'บางจาก', 'เชลล์', 'shell', 'คาลเท็กซ์', 'caltex', 'เอสโซ่', 'esso', 'ซัสโก้', 'susco'] },
  { cat: 'ค่าน้ำประปา', words: ['ค่าน้ำประปา', 'น้ำประปา', 'บิลน้ำ'] },
  { cat: 'ค่าไฟ', words: ['ค่าไฟ', 'บิลไฟ', 'การไฟฟ้า'] },
  { cat: 'เครื่องสำอาง', words: ['เครื่องสำอาง', 'ลิปสติก', 'ลิป', 'รองพื้น', 'คุชชั่น', 'มาสคาร่า', 'อายไลเนอร์', 'อายแชโดว์', 'บลัช', 'ปัดแก้ม', 'คอนซีลเลอร์', 'เซรั่ม', 'โทนเนอร์', 'กันแดด', 'มอยส์เจอไรเซอร์', 'ครีมบำรุง', 'คลีนซิ่ง', 'โฟมล้างหน้า', 'น้ำหอม', 'เขียนคิ้ว', 'มาส์กหน้า', 'แป้งพัฟ', 'แป้งฝุ่น'] },
  { cat: 'ค่าน้ำ', words: ['ชานม', 'ชาเย็น', 'ชาไทย', 'ชาเขียว', 'ชามะนาว', 'กาแฟ', 'โกโก้', 'น้ำเปล่า', 'น้ำดื่ม', 'น้ำแข็ง', 'น้ำอัดลม', 'น้ำหวาน', 'น้ำปั่น', 'น้ำผลไม้', 'โค้ก', 'เป๊ปซี่', 'สมูทตี้', 'นมสด', 'อเมซอน', 'คาเฟ่', 'ลาเต้', 'อเมริกาโน่', 'มอคค่า', 'คาปูชิโน่'] },
  { cat: 'ค่ากิน', words: ['ข้าว', 'กิน', 'อาหาร', 'ก๋วยเตี๋ยว', 'บะหมี่', 'เกี๊ยว', 'ส้มตำ', 'หมูกระทะ', 'ชาบู', 'สุกี้', 'ปิ้งย่าง', 'ขนม', 'โจ๊ก', 'ราดหน้า', 'ผัดไทย', 'ผัด', 'แกง', 'ต้มยำ', 'ไก่ทอด', 'หมูปิ้ง', 'ลูกชิ้น', 'ไอติม', 'ไอศกรีม', 'เค้ก', 'ขนมปัง', 'ผลไม้', 'มาม่า', 'พิซซ่า', 'เบอร์เกอร์', 'kfc', 'mk', 'สเต็ก', 'ซูชิ', 'มื้อ'] }
];

function normKey_(s) { return String(s || '').toLowerCase().replace(/\s+/g, '').replace(/สํา/g, 'สำ'); }

// แยกจำนวนเงินกับชื่อรายการ คืน null ถ้าไม่ใช่รูปแบบง่าย
function splitAmountAndItem_(text) {
  const t = String(text || '').trim();
  if (!t || t.length > 60 || /\n/.test(t)) return null;
  if (COMPLEX_WORDS_.test(t)) return null;
  const tokens = t.split(/\s+/);
  const isNum = tok => /^[\d,]+(\.\d+)?(บาท|บ\.?|฿)?$/.test(tok) && /\d/.test(tok);
  const nums = tokens.filter(isNum);
  let amountTok = null, rest = [];
  if (nums.length === 1) {
    amountTok = nums[0];
    rest = tokens.filter(tok => tok !== amountTok);
  } else if (nums.length === 0) {
    // ติดกัน เช่น "ข้าว50" หรือ "ข้าว50บาท"
    const m = t.match(/^(\D*?)\s*([\d,]+(?:\.\d+)?)\s*(บาท|บ\.?|฿)?\s*(\D*)$/);
    if (!m) return null;
    amountTok = m[2];
    rest = [m[1], m[4]].filter(Boolean);
  } else {
    return null;   // หลายตัวเลข เช่น "2 แก้ว 90" ให้ AI อ่าน
  }
  const amount = Number(String(amountTok).replace(/บาท|บ\.?|฿/g, '').replace(/,/g, ''));
  if (!(amount > 0) || amount > 1000000) return null;
  const item = rest.join(' ').replace(/(^|\s)(บาท|บ\.|฿)(?=\s|$)/g, ' ').replace(/\s+/g, ' ').trim();
  if (/\d/.test(item.replace(/\d+-\d+/g, ''))) return null;   // ยังมีตัวเลขอื่นปน (ยกเว้นชื่อแบบ 7-11)
  return { amount: amount, item: item };
}

// ---- คำที่เรียนรู้ไว้ (แท็บ Keywords) ----
function getKeywordSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_KEYWORDS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_KEYWORDS);
    sh.appendRow(['Keyword', 'Category', 'Count', 'UpdatedAt']);
  }
  return sh;
}

function getLearnedKeywords_() {
  return getKeywordSheet_().getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[1])
    .map(r => ({ key: normKey_(r[0]), category: String(r[1]), count: Number(r[2]) || 1 }));
}

function learnKeyword_(item, category) {
  const key = normKey_(item);
  if (!key || key.length < 2 || !category || /^[\d,.]+$/.test(key)) return;
  const sh = getKeywordSheet_();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (normKey_(rows[i][0]) === key) {
      sh.getRange(i + 1, 2, 1, 3).setValues([[category, (Number(rows[i][2]) || 0) + 1, new Date().toISOString()]]);
      return;
    }
  }
  sh.appendRow([String(item).trim(), category, 1, new Date().toISOString()]);
}

// คืน { category, source } หรือ { category: null, hints: [...] }
function guessCategory_(item, categories, expenses) {
  const key = normKey_(item);
  if (!key) return { category: null, hints: [] };

  // 1) พิมพ์ชื่อหมวดมาตรงๆ (เลือกชื่อที่ยาวที่สุดที่ตรง เช่น "ค่าน้ำมัน" ชนะ "ค่าน้ำ")
  const explicit = categories.filter(c => key.indexOf(normKey_(c)) === 0)
    .sort((a, b) => normKey_(b).length - normKey_(a).length)[0];
  if (explicit) return { category: explicit, source: 'explicit', note: item.replace(new RegExp('^\\s*' + explicit.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*'), '').trim() };

  // 2) คำที่เคยกดเลือกไว้: ตรงเป๊ะก่อน แล้วค่อยหาคำที่ยาวที่สุดที่อยู่ในข้อความ
  const learned = getLearnedKeywords_();
  const exact = learned.find(k => k.key === key);
  if (exact) return { category: exact.category, source: 'learned' };
  const partial = learned.filter(k => k.key.length >= 3 && key.indexOf(k.key) !== -1)
    .sort((a, b) => b.key.length - a.key.length);
  if (partial.length && partial.filter(p => p.key.length === partial[0].key.length && p.category !== partial[0].category).length === 0) {
    return { category: partial[0].category, source: 'learned' };
  }

  // 3) รายการเก่าในชีตที่หมายเหตุตรงกับข้อความนี้ (ต้องเป็นหมวดเดียวกันอย่างน้อย 70%)
  const counts = {};
  expenses.forEach(e => {
    const n = normKey_(e.note);
    if (!n || n.indexOf('≈') !== -1 || n.length > 40) return;
    if (n === key || (n.length >= 3 && key.indexOf(n) !== -1) || (key.length >= 5 && n.indexOf(key) !== -1)) counts[e.category] = (counts[e.category] || 0) + 1;
  });
  const total = Object.keys(counts).reduce((s, c) => s + counts[c], 0);
  const top = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0];
  if (top && counts[top] / total >= 0.7) return { category: top, source: 'history' };

  // 4) คำพื้นฐาน (ต้องตรงหมวดเดียวเท่านั้น ถ้าตรงหลายหมวดให้ถาม)
  const hits = [];
  BUILTIN_CATEGORY_WORDS_.forEach(g => {
    if (g.words.some(w => key.indexOf(normKey_(w)) !== -1)) hits.push(g.cat);
  });
  const usable = hits.filter(c => categories.indexOf(c) !== -1 || c === 'ค่าน้ำมัน' || c === COSMETIC_CATEGORY);
  if (usable.length === 1) return { category: usable[0], source: 'builtin' };
  return { category: null, hints: usable.length ? usable : (top ? [top] : []) };
}

// ---- ลองบันทึกเองโดยไม่ใช้ AI: คืน true ถ้าจัดการแล้ว (บันทึกหรือขึ้นปุ่ม) ----
function tryLocalExpense_(ev, text, senderUserId, props) {
  const sp = splitAmountAndItem_(text);
  if (!sp) return false;
  const data = getAllData_();
  const g = guessCategory_(sp.item, data.categories, data.expenses);

  if (g.category) {
    let note = g.source === 'explicit' ? g.note : sp.item;
    if (normKey_(note) === normKey_(g.category)) note = '';
    if (g.source === 'explicit' && note) learnKeyword_(note, g.category);   // "ค่ากิน 60 ร้านถุงเงิน" → จำว่าร้านถุงเงิน = ค่ากิน
    countAiUsage_('local');
    logParsedResult_({
      type: 'expense', amount: sp.amount, category: g.category, shared: false, splitHalf: false,
      daysAgo: 0, currency: 'THB', note: note, itemName: g.category === COSMETIC_CATEGORY ? (note || sp.item) : ''
    }, senderUserId, props, ev.replyToken, { suffix: pendingSuffix_(senderUserId) });
    return true;
  }

  // ไม่แน่ใจหมวด → เก็บไว้รอเลือก แล้วขึ้นปุ่ม
  const p = createPending_(senderUserId, sp.amount, sp.item);
  sendLineMessages_('reply', ev.replyToken, [buildPendingMessage_(p, data, g.hints, '')]);
  return true;
}

// ---- รายการรอเลือกหมวด ----
function createPending_(userId, amount, item) {
  // กันรายการซ้ำ: ยอดเท่ากันและชื่อเหมือน/คล้ายกัน (ชื่อหนึ่งอยู่ในอีกชื่อ) ใช้รายการเดิม
  const k = normKey_(item);
  const dup = getPendings_(userId).find(x => Math.abs(x.amount - amount) < 0.001 &&
    (normKey_(x.item) === k || (k && normKey_(x.item) && (normKey_(x.item).indexOf(k) !== -1 || k.indexOf(normKey_(x.item)) !== -1))));
  if (dup) return dup;
  const id = String(Date.now()) + Math.floor(Math.random() * 1000);
  const p = { id: id, userId: userId, amount: amount, item: item, localDate: localDatePrefix_(new Date(), 'yyyy-MM-dd'), at: new Date().toISOString() };
  PropertiesService.getScriptProperties().setProperty(PENDING_PREFIX_ + id, JSON.stringify(p));
  return p;
}

function getPendings_(userId) {
  const all = PropertiesService.getScriptProperties().getProperties();
  return Object.keys(all).filter(k => k.indexOf(PENDING_PREFIX_) === 0)
    .map(k => { try { return JSON.parse(all[k]); } catch (e) { return null; } })
    .filter(p => p && (!userId || p.userId === userId))
    .sort((a, b) => a.at.localeCompare(b.at));
}

function getPending_(id) {
  const raw = PropertiesService.getScriptProperties().getProperty(PENDING_PREFIX_ + id);
  try { return raw ? JSON.parse(raw) : null; } catch (e) { return null; }
}

function deletePending_(id) { PropertiesService.getScriptProperties().deleteProperty(PENDING_PREFIX_ + id); }

function pendingSuffix_(userId) {
  const n = getPendings_(userId).length;
  return n ? `\n\n⏳ ยังรอเลือกหมวดอีก ${n} รายการ พิมพ์ "รอเลือกหมวด"` : '';
}

// หมวดที่คนนี้ใช้บ่อยใน 90 วัน มาก่อน
function categoriesForButtons_(userId, data, hints) {
  const props = PropertiesService.getScriptProperties();
  const who = resolveWho_(userId, props);
  const since = localDatePrefix_(new Date(Date.now() - 90 * 86400000), 'yyyy-MM-dd');
  const freq = {};
  data.expenses.forEach(e => { if (e.who === who && (e.localDate || '') >= since) freq[e.category] = (freq[e.category] || 0) + 1; });
  const order = data.categories.slice().sort((a, b) => (freq[b] || 0) - (freq[a] || 0));
  const list = [];
  (hints || []).concat(order).forEach(c => { if (c && list.indexOf(c) === -1) list.push(c); });
  return list.slice(0, 11);
}

function buildPendingMessage_(p, data, hints, prefix, totalPending) {
  const many = (totalPending || getPendings_(p.userId).length) > 1;
  const cats = categoriesForButtons_(p.userId, data || getAllData_(), hints).slice(0, many ? 10 : 11);
  const label = p.item ? `${fmtBaht_(p.amount)} บาท · ${p.item}` : `${fmtBaht_(p.amount)} บาท`;
  const items = cats.map(c => ({ label: c, data: 'cat:' + p.id + ':' + c, display: c }));
  items.push({ label: '➕ หมวดใหม่', data: 'catnew:' + p.id, display: 'หมวดใหม่' });
  items.push({ label: 'ยกเลิก', data: 'catx:' + p.id, display: 'ยกเลิก' });
  if (many) items.push({ label: '🗑 ล้างทั้งหมด', data: 'catclear:' + p.id, display: 'ล้างรายการรอเลือกหมวดทั้งหมด' });
  const dayNote = p.localDate !== localDatePrefix_(new Date(), 'yyyy-MM-dd') ? ` (ของวันที่ ${cosThDate_(p.localDate)})` : '';
  return withQuickReply_({ type: 'text', text: (prefix || '') + `🗂️ ${label}${dayNote}\nเลือกหมวดค่ะ (เลือกครั้งเดียว ครั้งหน้าน้องออมจำให้เอง)` }, items);
}

// บันทึกรายการที่รอ เมื่อได้หมวดแล้ว
function savePendingWithCategory_(p, category, replyToken) {
  const props = PropertiesService.getScriptProperties();
  deletePending_(p.id);
  props.deleteProperty(AWAIT_CAT_PREFIX_ + p.userId);
  if (p.item) learnKeyword_(p.item, category);
  countAiUsage_('button');
  logParsedResult_({
    type: 'expense', amount: p.amount, category: category, shared: false, splitHalf: false,
    currency: 'THB', note: p.item || '', itemName: category === COSMETIC_CATEGORY ? (p.item || '') : '',
    transactionDate: p.localDate
  }, p.userId, props, replyToken, { suffix: (p.item ? `\n(จำแล้วค่ะ ครั้งหน้า "${p.item}" จะเข้าหมวด ${category} ให้เอง)` : '') + pendingSuffix_(p.userId) });
}

function clearPendings_(userId) {
  const list = getPendings_(userId);
  list.forEach(p => deletePending_(p.id));
  PropertiesService.getScriptProperties().deleteProperty(AWAIT_CAT_PREFIX_ + userId);
  return list;
}

function clearPendingsText_(list) {
  if (!list.length) return 'ไม่มีรายการรอเลือกหมวดค่ะ 👍';
  return `ล้างรายการรอเลือกหมวด ${list.length} รายการแล้วค่ะ (ไม่ได้บันทึก)\n` +
    list.slice(0, 15).map(p => `• ${fmtBaht_(p.amount)} บาท${p.item ? ' · ' + p.item : ''}`).join('\n') +
    '\n\nถ้ามีรายการไหนที่ต้องบันทึก พิมพ์ใหม่ได้เลยค่ะ';
}

function handleCategoryPostback_(ev, data, senderUserId, props) {
  const parts = data.split(':');
  const kind = parts[0], id = parts[1];
  if (kind === 'catclear') { replyLine_(ev.replyToken, clearPendingsText_(clearPendings_(senderUserId))); return; }
  const p = getPending_(id);
  if (!p) { replyLine_(ev.replyToken, 'รายการนี้บันทึกหรือยกเลิกไปแล้วค่ะ'); return; }
  if (p.userId !== senderUserId) {
    replyLine_(ev.replyToken, `ปุ่มนี้สำหรับ${resolveWho_(p.userId, props)}ค่ะ (คนที่พิมพ์รายการนี้เป็นคนเลือกหมวด)`);
    return;
  }
  if (kind === 'catx') {
    deletePending_(id);
    props.deleteProperty(AWAIT_CAT_PREFIX_ + senderUserId);
    replyLine_(ev.replyToken, `ยกเลิกแล้วค่ะ ไม่ได้บันทึก ${fmtBaht_(p.amount)} บาท${p.item ? ' · ' + p.item : ''}` + pendingSuffix_(senderUserId));
    return;
  }
  if (kind === 'catnew') {
    props.setProperty(AWAIT_CAT_PREFIX_ + senderUserId, JSON.stringify({ id: id, at: Date.now() }));
    replyLine_(ev.replyToken, `พิมพ์แค่ชื่อหมวดใหม่ (ไม่ต้องใส่ราคา) สำหรับ ${fmtBaht_(p.amount)} บาท${p.item ? ' · ' + p.item : ''} ได้เลยค่ะ\nเช่น "ค่าโทรศัพท์" หรือ "ค่าเดินทาง" (พิมพ์ "ยกเลิก" ถ้าไม่บันทึกแล้ว)`);
    return;
  }
  const category = parts.slice(2).join(':');
  if (!category) return;
  savePendingWithCategory_(p, category, ev.replyToken);
}

// กำลังรอให้พิมพ์ชื่อหมวดใหม่อยู่ไหม: คืน true ถ้าข้อความนี้ถูกใช้เป็นชื่อหมวดแล้ว
const RESERVED_TEXTS_ = ['ล้างรอเลือกหมวด', 'สรุป', 'สรุปวันนี้', 'สรุปเดือนนี้', 'วิธีใช้', 'ลบ', 'ลบรายการ', 'รอเลือกหมวด', 'เครดิตai', 'เครื่องสำอาง', 'ยอดยืม', 'สำรองข้อมูล', 'บันทึกรายจ่าย'];
function handleAwaitCategory_(ev, text, senderUserId, props) {
  const raw = props.getProperty(AWAIT_CAT_PREFIX_ + senderUserId);
  if (!raw) return false;
  let w = null;
  try { w = JSON.parse(raw); } catch (e) {}
  if (!w || Date.now() - w.at > AWAIT_CAT_MINUTES_ * 60000) { props.deleteProperty(AWAIT_CAT_PREFIX_ + senderUserId); return false; }
  const p = getPending_(w.id);
  if (!p) { props.deleteProperty(AWAIT_CAT_PREFIX_ + senderUserId); return false; }
  const t = text.trim();
  if (t === 'ยกเลิก') {
    deletePending_(p.id); props.deleteProperty(AWAIT_CAT_PREFIX_ + senderUserId);
    replyLine_(ev.replyToken, `ยกเลิกแล้วค่ะ ไม่ได้บันทึก ${fmtBaht_(p.amount)} บาท` + pendingSuffix_(senderUserId));
    return true;
  }
  if (/\d/.test(t)) {
    const sp = splitAmountAndItem_(t);
    if (sp && Math.abs(sp.amount - p.amount) < 0.001) {
      // พิมพ์ราคาเดิมซ้ำ = ยังเป็นรายการเดิม ถามหมวดอีกครั้ง ไม่สร้างรายการใหม่
      props.setProperty(AWAIT_CAT_PREFIX_ + senderUserId, JSON.stringify({ id: p.id, at: Date.now() }));
      sendLineMessages_('reply', ev.replyToken, [buildPendingMessage_(p, null, [],
        'ตอนนี้รอชื่อหมวดของรายการนี้อยู่ค่ะ ไม่ต้องพิมพ์ราคาซ้ำ\nกดเลือกหมวดด้านล่าง หรือพิมพ์แค่ชื่อหมวดใหม่ เช่น "ค่าโทรศัพท์"\n\n')]);
      return true;
    }
    props.deleteProperty(AWAIT_CAT_PREFIX_ + senderUserId);   // ยอดไม่ตรง = รายจ่ายใหม่ เลิกรอ (รายการเดิมยังค้างให้เลือกทีหลัง)
    return false;
  }
  if (t.length > 25 || RESERVED_TEXTS_.indexOf(normKey_(t)) !== -1) {
    props.deleteProperty(AWAIT_CAT_PREFIX_ + senderUserId);   // พิมพ์เรื่องอื่นแล้ว เลิกรอ (รายการยังค้างให้เลือกทีหลัง)
    return false;
  }
  savePendingWithCategory_(p, t, ev.replyToken);
  return true;
}

function showNextPending_(replyToken, userId) {
  const list = getPendings_(userId);
  if (!list.length) { replyLine_(replyToken, 'ไม่มีรายการรอเลือกหมวดค่ะ 👍'); return; }
  const prefix = list.length > 1 ? `รอเลือกหมวด ${list.length} รายการ เริ่มจากรายการแรกค่ะ\n\n` : '';
  sendLineMessages_('reply', replyToken, [buildPendingMessage_(list[0], null, [], prefix, list.length)]);
}

// สำหรับสรุป 20:00: บอกรายการที่ยังไม่ได้เลือกหมวด
function pendingSummaryLines_() {
  const props = PropertiesService.getScriptProperties();
  const list = getPendings_(null);
  if (!list.length) return [];
  const lines = ['', '⏳ รอเลือกหมวด (ยังไม่ได้บันทึก):'];
  list.slice(0, 10).forEach(p => lines.push(`• ${fmtBaht_(p.amount)} บาท${p.item ? ' · ' + p.item : ''} (${resolveWho_(p.userId, props)})`));
  lines.push('พิมพ์ "รอเลือกหมวด" เพื่อเลือกต่อค่ะ');
  return lines;
}

// ---------- นับการใช้ AI + จัดการเมื่อ AI ใช้ไม่ได้ ----------
let LAST_AI_ERROR_ = null;   // { code, text } ของการเรียก AI ครั้งล่าสุดที่ล้มเหลว (เฉพาะรอบการทำงานนี้)

function countAiUsage_(kind) {
  try {
    const props = PropertiesService.getScriptProperties();
    const key = 'AI_USAGE_' + localDatePrefix_(new Date(), 'yyyy-MM');
    let u = {};
    try { u = JSON.parse(props.getProperty(key) || '{}'); } catch (e) { u = {}; }
    u[kind] = (u[kind] || 0) + 1;
    props.setProperty(key, JSON.stringify(u));
  } catch (e) { Logger.log('countAiUsage_ error: ' + e); }
}

function aiErrorKind_(err) {
  if (!err) return null;
  const t = String(err.text || '');
  if (/credit balance|billing|purchase credits/i.test(t)) return 'credit';
  if (err.code === 401 || err.code === 403 || /invalid x-api-key|authentication/i.test(t)) return 'key';
  if (err.code === 429) return 'rate';
  if (err.code === 529 || err.code >= 500) return 'busy';
  return 'other';
}

function aiErrorUserText_(kind) {
  if (kind === 'credit') return '⚠️ ตอนนี้เครดิต AI หมด น้องออมเลยอ่านข้อความนี้ไม่ได้ค่ะ ยังไม่ได้บันทึก';
  if (kind === 'key') return '⚠️ ระบบ AI ใช้ไม่ได้ (API key มีปัญหา) น้องออมเลยอ่านข้อความนี้ไม่ได้ค่ะ ยังไม่ได้บันทึก';
  if (kind === 'rate' || kind === 'busy') return '⚠️ ระบบ AI ไม่ว่างชั่วคราว ลองส่งใหม่อีกครั้งในอีกสักครู่นะคะ ยังไม่ได้บันทึก';
  return '⚠️ ระบบ AI มีปัญหาชั่วคราว ยังไม่ได้บันทึกค่ะ';
}

// แจ้งโฟกัสวันละครั้งเมื่อเครดิตหมด/คีย์มีปัญหา
function notifyAdminAiProblem_(kind) {
  if (kind !== 'credit' && kind !== 'key') return;
  const props = PropertiesService.getScriptProperties();
  const today = localDatePrefix_(new Date(), 'yyyy-MM-dd');
  if (props.getProperty('AI_ALERT_DATE') === today) return;
  props.setProperty('AI_ALERT_DATE', today);
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const adminId = props.getProperty('ADMIN_LINE_ID') || props.getProperty('PARTNER1_LINE_ID');
  if (!token || !adminId) return;
  pushLine_(token, adminId, kind === 'credit'
    ? '⚠️ เครดิต Claude API หมดแล้ว\nข้อความที่ซับซ้อนและการอ่านสลิปจะใช้ไม่ได้จนกว่าจะเติมเครดิต\n(ข้อความแบบ "ชื่อ ตัวเลข" ยังบันทึกได้ตามปกติ)\nเติมได้ที่ console.anthropic.com → Billing'
    : '⚠️ Claude API key ใช้ไม่ได้ (401/403)\nเช็ค CLAUDE_API_KEY ใน Script Properties');
}

function aiUsageText_() {
  const props = PropertiesService.getScriptProperties();
  const mk = localDatePrefix_(new Date(), 'yyyy-MM');
  let u = {};
  try { u = JSON.parse(props.getProperty('AI_USAGE_' + mk) || '{}'); } catch (e) {}
  const ai = (u.text || 0) + (u.image || 0);
  const free = (u.local || 0) + (u.button || 0);
  const pct = ai + free ? Math.round(free / (ai + free) * 100) : 0;
  const learned = getLearnedKeywords_().length;
  return `🤖 การใช้ AI เดือน${monthLabelTh_(mk)}\n` +
    `• อ่านข้อความด้วย AI: ${u.text || 0} ครั้ง\n` +
    `• อ่านสลิปด้วย AI: ${u.image || 0} ครั้ง\n` +
    `• ระบบอ่านเอง: ${u.local || 0} ครั้ง · กดปุ่มเลือกหมวด: ${u.button || 0} ครั้ง\n` +
    `→ ไม่ใช้เครดิต ${pct}% ของรายการ\n` +
    `• คำที่จำหมวดไว้แล้ว: ${learned} คำ (แท็บ Keywords)`;
}

// ---------- บันทึกจากหน้าแอป (เฉพาะผู้ดูแลที่มี ADMIN_KEY) ----------
// who = ชื่อคน หรือ 'half' (หารครึ่ง แยกเป็น 2 รายการ) · localDate = YYYY-MM-DD (ย้อนหลังได้)

function appSaveExpense_(who, category, amount, note, localDate) {
  const props = PropertiesService.getScriptProperties();
  const [name1, name2] = [props.getProperty('PARTNER1') || 'คนที่ 1', props.getProperty('PARTNER2') || 'คนที่ 2'];
  amount = Number(amount);
  if (!category || !(amount > 0)) return { error: 'กรอกหมวดและจำนวนเงินให้ถูกต้อง' };
  const today = localDatePrefix_(new Date(), 'yyyy-MM-dd');
  const ld = /^\d{4}-\d{2}-\d{2}$/.test(localDate || '') ? localDate : today;
  if (ld > today) return { error: 'เลือกวันที่ในอนาคตไม่ได้' };
  const iso = ld === today ? new Date().toISOString() : new Date(ld + 'T12:00:00').toISOString();
  const recordedAt = new Date().toISOString();
  note = String(note || '').trim();

  if (who === 'half') {
    const half = Math.round((amount / 2) * 100) / 100;
    addExpense_(name1, category, half, note, iso, ld, recordedAt);
    addExpense_(name2, category, half, note, iso, ld, recordedAt);
  } else {
    const w = (who === name1 || who === name2 || who === WHO_SHARED) ? who : name1;
    addExpense_(w, category, amount, note, iso, ld, recordedAt);
    // เครื่องสำอางของอ๋อม: เข้าคลังด้วย
    if (category === COSMETIC_CATEGORY && w === cosmeticOwnerName_(props)) {
      try { addCosmeticItem_(note || COSMETIC_CATEGORY, amount, ld, true, ''); } catch (e) { Logger.log('app cosmetic add: ' + e); }
    }
  }
  addCategory_(category);
  if (note && note.length <= 30) learnKeyword_(note, category);   // ครั้งหน้าพิมพ์ใน LINE จะเข้าหมวดนี้เอง
  return null;
}

function addExpenseFromApp_(p) {
  const err = appSaveExpense_(p.who, p.category, p.amount, p.note, p.localDate);
  return err ? Object.assign(getAllData_(), err) : null;
}

function resolvePendingFromApp_(p) {
  const pend = getPending_(p.id);
  if (!pend) return Object.assign(getAllData_(), { error: 'รายการนี้บันทึกหรือยกเลิกไปแล้ว' });
  const props = PropertiesService.getScriptProperties();
  const who = p.who || resolveWho_(pend.userId, props);
  const err = appSaveExpense_(who, p.category, pend.amount, pend.item, p.localDate || pend.localDate);
  if (err) return Object.assign(getAllData_(), err);
  deletePending_(pend.id);
  props.deleteProperty(AWAIT_CAT_PREFIX_ + pend.userId);
  countAiUsage_('button');
  return null;
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

/**
 * รันครั้งเดียวจาก Editor เพื่อ "รีเซ็ต" ล้างรายการรายจ่ายเก่าทั้งหมดออก
 * แล้วใส่รายการเริ่มต้นของวันนี้ให้ใหม่: โฟกัส ค่ากิน 70, อ๋อม ค่ากิน 125
 * (แก้ตัวเลข/ชื่อหมวดในฟังก์ชันนี้เองได้ถ้าอยากเปลี่ยนค่าเริ่มต้น)
 */
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

  writeBackupStatusToDashboard_();

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
