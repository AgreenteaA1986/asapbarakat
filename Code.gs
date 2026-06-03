/**
 * Managers Pro - Business Intelligence (Server-side)
 * ระบบ Chain Update อัตโนมัติและเพิ่มประสิทธิภาพการเข้าถึงข้อมูล
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Managers Pro - Business Intelligence')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- 1. การดึงข้อมูลรายชื่อลูกค้า ---
function getAllCustomers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customers = new Set();
  
  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    // กรองดึงข้อมูลเฉพาะชีตที่เป็นปีตัวเลข 4 หลัก (เช่น "2024", "2025", "2026")
    if (!/^\d{4}$/.test(sheetName)) return; 

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // ดึงข้อมูลคอลัมน์ B (คอลัมน์ที่ 2) คือ ชื่อลูกค้า
      const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      data.forEach(row => {
        if (row[0]) customers.add(row[0].toString().trim());
      });
    }
  });
  
  return Array.from(customers).sort();
}

// --- 2. การดึงธุรกรรมทั้งหมดเพื่อแสดง Dashboard (อัปเดตดึงข้อมูลถึงคอลัมน์ X) ---
function getTransactions(customerName = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let transactions = [];
  const searchName = customerName ? customerName.trim().toLowerCase() : null;

  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (!/^\d{4}$/.test(sheetName)) return; 

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // 🆕 ดึงคอลัมน์ A ถึง X (รวม 24 คอลัมน์)
      const data = sheet.getRange(2, 1, lastRow - 1, 24).getValues();
      
      data.forEach((row, index) => {
        const rowCust = row[1] ? row[1].toString().trim() : ""; 
        
        if (!searchName || rowCust.toLowerCase() === searchName) {
          const smcNo = row[2] ? row[2].toString().trim() : ""; 
          const smcDate = row[3] || ""; 
          const invNo = row[4] ? row[4].toString().trim() : ""; 
          const invDate = row[5] || ""; 
          const blNo = row[6] ? row[6].toString().trim() : "";  
          const blDate = row[7] || "";  
          const rcNo = row[8] ? row[8].toString().trim() : "";  
          const rcDate = row[9] || "";  
          
          const subtotal = parseFloat(row[11]) || 0;
          const vat = parseFloat(row[12]) || 0;
          const wht = parseFloat(row[13]) || 0;
          const net = parseFloat(row[14]) || 0;

          const docWht = row[20] === true || row[20] === "TRUE"; 
          const sendAcc = row[21] === true || row[21] === "TRUE"; 
          const sendAccDate = row[22] || ""; 
          
          // 🆕 ดึงข้อมูล "เอกสารอื่นๆ / หมายเหตุ" จาก Col X (Index 23)
          const deliveryRemark = row[23] ? row[23].toString().trim() : "";
          
          let smcStatus = "", invStatus = "", blStatus = "", rcStatus = "";
          if (rcNo) {
            smcStatus = "ชำระแล้ว"; invStatus = "ชำระแล้ว"; blStatus = "ชำระแล้ว"; rcStatus = "ชำระแล้ว";
          } else if (blNo) {
            smcStatus = "วางบิล"; invStatus = "วางบิล"; blStatus = "ค้างชำระ";
          } else if (invNo) {
            smcStatus = "รอดำเนินการ"; invStatus = "รอดำเนินการ";
          } else if (smcNo) {
            smcStatus = "รอดำเนินการ";
          }

          // 🆕 ส่งตัวแปร deliveryRemark พ่วงไปด้วย
          if (smcNo) transactions.push(createTxObj('ใบส่งงาน/สรุปค่าใช้จ่าย', smcNo, smcDate, rowCust, subtotal, vat, wht, net, smcStatus, "-", index + 2, docWht, sendAcc, sendAccDate, deliveryRemark));
          if (invNo) transactions.push(createTxObj('ใบแจ้งหนี้', invNo, invDate, rowCust, subtotal, vat, wht, net, invStatus, smcNo, index + 2, docWht, sendAcc, sendAccDate, deliveryRemark));
          if (blNo) transactions.push(createTxObj('ใบวางบิล', blNo, blDate, rowCust, subtotal, vat, wht, net, blStatus, invNo, index + 2, docWht, sendAcc, sendAccDate, deliveryRemark));
          if (rcNo) transactions.push(createTxObj('ใบเสร็จรับเงิน/ใบกำกับภาษี', rcNo, rcDate, rowCust, subtotal, vat, wht, net, rcStatus, blNo, index + 2, docWht, sendAcc, sendAccDate, deliveryRemark));
        }
      });
    }
  });
  
  return transactions;
}

// 🆕 อัปเดตฟังก์ชันตัวช่วยให้รับค่า deliveryRemark
function createTxObj(type, docNo, date, customer, amount, vat, wht, net, status, reference, rowId, docWht, sendAcc, sendAccDate, deliveryRemark) {
  return {
    type: type, docNo: docNo, date: date instanceof Date ? Utilities.formatDate(date, "GMT+7", "yyyy-MM-dd") : date,
    customer: customer, amount: amount, vat: vat, wht: wht, net: net, status: status, reference: reference, rowId: rowId,
    docWht: docWht, sendAcc: sendAcc, sendAccDate: sendAccDate instanceof Date ? Utilities.formatDate(sendAccDate, "GMT+7", "yyyy-MM-dd") : sendAccDate,
    deliveryRemark: deliveryRemark
  };
}
function saveData(fd) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(10000); // รอคิวสูงสุด 10 วินาทีเพื่อป้องกัน Data Collision

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. ตรวจสอบและดึงชื่อชีตตาม "ปี" ของวันที่เอกสารอัตโนมัติ
    if (!fd.docDate) throw new Error("กรุณาระบุวันที่เอกสาร");
    const docDateObj = new Date(fd.docDate);
    const sheetName = docDateObj.getFullYear().toString(); // ได้เป็นข้อความเช่น "2025" หรือ "2026"
    
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      // 🆕 หากไม่พบแผ่นงานปีนี้ ให้สร้างขึ้นมาใหม่โดยอัตโนมัติ
      sheet = ss.insertSheet(sheetName);
    }

    // 2. คำนวณภาษีและยอดสุทธิประจำรายการ
    const amt = parseFloat(fd.amount) || 0;
    const v = amt * 0.07; // VAT 7%
    const w = amt * 0.03; // WHT 3%
    const n = amt + v - w;

    let targetRow = null;
    let rowValues = Array(24).fill(""); // 🆕 เตรียมพื้นที่คอลัมน์ A ถึง X (รวม 24 คอลัมน์)

    // 3. ตรวจสอบว่าเป็นการแก้ไข หรือเป็นการสร้างเอกสารใหม่มาเชื่อมโยงกับงานเดิม
    if (fd.rowId && !isNaN(fd.rowId)) {
      // เคส A: กดแก้ไขจากหน้าบ้านโดยตรง
      targetRow = parseInt(fd.rowId);
      rowValues = sheet.getRange(targetRow, 1, 1, 24).getValues()[0];
    } else if (fd.reference && fd.reference !== "-") {
      // เคส B: สร้างเอกสารใหม่แต่มีการใส่ "อ้างอิงเลขที่"
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        // 🆕 ดึงคอลัมน์ C ถึง I (ดัชนีคอลัมน์ที่ 3 ถึง 9) เพื่อค้นหาเลขที่เอกสารอ้างอิงในแถวเดิม
        const searchRange = sheet.getRange(2, 3, lastRow - 1, 7).getValues();
        const refUpper = fd.reference.trim().toUpperCase();
        
        for (let i = 0; i < searchRange.length; i++) {
          // ค้นหาว่าตรงกับเลขที่ ใบส่งงาน(0), ใบแจ้งหนี้(2), ใบวางบิล(4), หรือ ใบเสร็จ(6) หรือไม่
          if ((searchRange[i][0] && searchRange[i][0].toString().toUpperCase() === refUpper) ||
              (searchRange[i][2] && searchRange[i][2].toString().toUpperCase() === refUpper) ||
              (searchRange[i][4] && searchRange[i][4].toString().toUpperCase() === refUpper) ||
              (searchRange[i][6] && searchRange[i][6].toString().toUpperCase() === refUpper)) {
            targetRow = i + 2; // เจอแถวที่มีอยู่แล้ว
            rowValues = sheet.getRange(targetRow, 1, 1, 24).getValues()[0];
            break;
          }
        }
      }
    }

    // เคส C: หากไม่พบแถวเดิม หรือไม่มีการอ้างอิง ให้ต่อท้ายเป็นรายการแถวใหม่
    if (!targetRow) {
      targetRow = sheet.getLastRow() + 1;
      if (targetRow === 1) targetRow = 2; // ป้องกันการเขียนทับแถว Header
    }

    // 4. อัปเดตข้อมูลพื้นฐานลงใน Array แถว
    rowValues[1] = fd.customer.trim(); // คอลัมน์ B: ชื่อลูกค้า

    // 5. นำเลขที่เอกสารและวันที่ไปหยอดลงคอลัมน์ให้ตรงตามประเภทเอกสาร
    const docType = fd.docType.trim();
    if (docType.includes("ใบส่งงาน")) {
      rowValues[2] = fd.docNo.trim();  // คอลัมน์ C
      rowValues[3] = docDateObj;       // คอลัมน์ D
    } else if (docType.includes("ใบแจ้งหนี้")) {
      rowValues[4] = fd.docNo.trim();  // คอลัมน์ E
      rowValues[5] = docDateObj;       // คอลัมน์ F
    } else if (docType.includes("ใบวางบิล")) {
      rowValues[6] = fd.docNo.trim();  // คอลัมน์ G
      rowValues[7] = docDateObj;       // คอลัมน์ H
    } else if (docType.includes("ใบเสร็จ")) {
      rowValues[8] = fd.docNo.trim();  // คอลัมน์ I
      rowValues[9] = docDateObj;       // คอลัมน์ J
    }

    // 6. อัปเดตตัวเลขการเงิน (คอลัมน์ L, M, N, O) เป็นยอดล่าสุดที่มีการอัปเดตเข้ามา
    rowValues[11] = amt; // คอลัมน์ L: Subtotal
    rowValues[12] = v;   // คอลัมน์ M: Vat7
    rowValues[13] = w;   // คอลัมน์ N: WHT3
    rowValues[14] = n;   // คอลัมน์ O: Revenue

    // 7. บันทึกข้อมูลทั้งหมดกลับลงชีตในรอบเดียว (Bulk Update ความเร็วสูง)
    sheet.getRange(targetRow, 1, 1, 24).setValues([rowValues]);

    return "✅ บันทึกและเชื่อมโยงชุดข้อมูลในแถวเรียบร้อยแล้ว";
    
  } catch (e) {
    return "❌ ข้อผิดพลาด: " + e.message;
  } finally {
    lock.releaseLock(); // ปลดล็อคสคริปต์
  }
}
function saveAccDelivery(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. อัปเดตสถานะในชีตรายปี (Main Sheet)
    // 🆕 วนลูปและวิ่งไปอัปเดตชีตตาม "ปีของเอกสารนั้นๆ" (docYear)
    payload.items.forEach(item => {
      const row = parseInt(item.rowId);
      const docSheetName = item.docYear.toString();
      const docSheet = ss.getSheetByName(docSheetName);
      
      if (docSheet) {
        if (item.isChecked) {
          docSheet.getRange(row, 22).setValue(true);                  // Col V: ติ๊กว่านำส่งบัญชีแล้ว
          docSheet.getRange(row, 23).setValue(payload.deliveryDate);  // Col W: ใส่วันที่นำส่ง
          docSheet.getRange(row, 24).clearContent();                  // Col X: เคลียร์ทิ้ง
        } else {
          docSheet.getRange(row, 22).setValue(false); 
          docSheet.getRange(row, 23).clearContent();
          docSheet.getRange(row, 24).clearContent();
        }
      }
    });

    // 2. จัดการบันทึกแท็บ "ประวัตินำส่งบัญชี" (ใช้โค้ดเดิมต่อจากนี้ได้เลยครับ)
    let accSheet = ss.getSheetByName("ประวัตินำส่งบัญชี");
    if (!accSheet) throw new Error("ไม่พบแท็บ 'ประวัตินำส่งบัญชี'");

    const accData = accSheet.getDataRange().getValues();
    let targetAccRow = -1;
    const searchDate = payload.deliveryDate;
    
    for (let i = 1; i < accData.length; i++) {
      let rowDateStr = accData[i][0] instanceof Date
        ? Utilities.formatDate(accData[i][0], "GMT+7", "yyyy-MM-dd")
        : (accData[i][0] ? accData[i][0].toString().trim() : "");
      
      if (rowDateStr === searchDate) {
        targetAccRow = i + 1; 
        break;
      }
    }

    const recordRow = [ payload.deliveryDate, payload.docList, payload.remark ];

    if (targetAccRow > -1) {
      accSheet.getRange(targetAccRow, 1, 1, 3).setValues([recordRow]);
    } else {
      accSheet.appendRow(recordRow);
    }

    return "✅ บันทึกข้อมูลและจัดเก็บประวัติการนำส่งสำเร็จ";
  } catch (e) {
    throw new Error(e.message); 
  } finally {
    lock.releaseLock();
  }
}
function getAccHistoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ประวัตินำส่งบัญชี");
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const historyMap = {};
  
  // ลูปดึงข้อมูลมาทำเป็น Object โดยใช้วันที่ (Key) เพื่อให้หน้าบ้านดึงไปใช้ง่ายๆ
  for (let i = 1; i < data.length; i++) {
    let dateKey = data[i][0] instanceof Date 
      ? Utilities.formatDate(data[i][0], "GMT+7", "yyyy-MM-dd") 
      : (data[i][0] ? data[i][0].toString().trim() : "");
      
    if (dateKey) {
      historyMap[dateKey] = {
        docList: data[i][1] ? data[i][1].toString() : "",
        remark: data[i][2] ? data[i][2].toString() : ""
      };
    }
  }
  return historyMap;
}