/**
 * ==========================================
 * Panning System - Code Version 5.2 (Plan Status Table added)
 * ==========================================
 */

const SHEET_ID = '1TdSEJl7vidKFyqVYpuu1mMBjnbJIHjE-Kq7_bAMkU_8';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SCTC Planning')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getDepartments() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('ตั้งค่าระบบ');
    if(!sheet) return [];

    const data = sheet.getDataRange().getValues();
    let result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') {
        result.push({ name: data[i][0].toString().trim(), code: (data[i][1] || data[i][0]).toString().trim() });
      }
    }
    return result;
  } catch (err) {
    return [];
  }
}

function authenticate(username, password) {
  if (username === 'admin' && password === '1234') {
    return { status: 'success', role: 'admin', deptCode: 'ALL', deptName: 'ผู้ดูแลระบบ' };
  }
  if (username === 'เจ้าหน้าที่แผน' && password === 'plan1234') { 
    return { status: 'success', role: 'plan_officer', deptCode: 'ALL', deptName: 'เจ้าหน้าที่แผน' };
  }
  if (username === 'เจ้าหน้าที่การเงิน' && password === 'fin1234') { 
    return { status: 'success', role: 'finance_officer', deptCode: 'ALL', deptName: 'เจ้าหน้าที่การเงิน' };
  }

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('ตั้งค่าระบบ');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (username === data[i][0].toString().trim() && password === data[i][1].toString().trim()) {
      return { status: 'success', role: 'user', deptCode: data[i][1].toString().trim(), deptName: data[i][0].toString().trim() };
    }
  }
  return { status: 'error', message: 'รหัสผ่านไม่ถูกต้อง หรือคุณไม่มีสิทธิ์เข้าถึงบัญชีนี้' };
}

function getSummaryData() {
  const ss = getSpreadsheet();
  const summarySheet = ss.getSheetByName('Summary');
  if (!summarySheet) return [];
  
  let reqMap = {}, actMap = {}; 
  const reqData = ss.getSheetByName('BudgetReq')?.getDataRange().getValues() || [];
  const actData = ss.getSheetByName('BudgetActual')?.getDataRange().getValues() || [];

  for(let k=1; k<reqData.length; k++) {
    let deptCode = reqData[k][1]?.toString().split('-')[1];
    if(deptCode) reqMap[deptCode] = (reqMap[deptCode] || 0) + (parseFloat(reqData[k][5]) || 0);
  }
  for(let k=1; k<actData.length; k++) {
    let deptCode = actData[k][1]?.toString().split('-')[1];
    if(deptCode) actMap[deptCode] = (actMap[deptCode] || 0) + (parseFloat(actData[k][6]) || 0);
  }

  const data = summarySheet.getDataRange().getDisplayValues();
  const headers = data[0];
  let result = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][0].includes('รวม') || data[i][0] === '#N/A') continue;
    let rowObj = {};
    headers.forEach((h, j) => rowObj[h] = data[i][j]);
    
    let budgetTotal = parseFloat(rowObj['งบประมาณรวม']?.replace(/,/g, '')) || 0;
    rowObj['งบประมาณคงเหลือ (แผน)'] = budgetTotal - (reqMap[rowObj['รหัส']] || 0);
    rowObj['งบประมาณคงเหลือ (จริง)'] = budgetTotal - (actMap[rowObj['รหัส']] || 0);
    result.push(rowObj);
  }
  return result;
}

function getOutcomeDashboard(filterDeptCode, filterPlanId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('outcome');
  if(!sheet) throw new Error('ไม่พบชีตชื่อ "outcome" ใน Google Sheet');
  
  const data = sheet.getDataRange().getDisplayValues();
  if(data.length < 2) return { totalPlans: 0, totalBudget: 0, reqBudget: 0, remBudget: 0, typeArr: [], statusCount: {'สำเร็จ': 0, 'ปรับปรุง': 0}, planStatusCount: {'ยังไม่ได้ดำเนินการ':0, 'กำลังดำเนินการ':0, 'ดำเนินการแล้ว':0}, planOptions: [] };
  
  // --- โหลดข้อมูลสถานะแผนจาก BudgetReq และ Result ---
  let inProgressSet = new Set();
  const reqSheet = ss.getSheetByName('BudgetReq');
  if (reqSheet) {
    const reqData = reqSheet.getDataRange().getDisplayValues();
    const planIdIdx = reqData[0].indexOf('รหัสแผน');
    if (planIdIdx > -1) {
      for (let i = 1; i < reqData.length; i++) {
        if (reqData[i][planIdIdx]) inProgressSet.add(reqData[i][planIdIdx].toString().trim());
      }
    }
  }

  let completedSet = new Set();
  const resultSheet = ss.getSheetByName('Result');
  if (resultSheet) {
    const resData = resultSheet.getDataRange().getDisplayValues();
    const planIdIdx = resData[0].indexOf('รหัส');
    const recIdx = resData[0].indexOf('ได้รับเอกสาร');
    if (planIdIdx > -1 && recIdx > -1) {
      for (let i = 1; i < resData.length; i++) {
        if (resData[i][recIdx] === 'ได้รับแล้ว' && resData[i][planIdIdx]) completedSet.add(resData[i][planIdIdx].toString().trim());
      }
    }
  }
  // ------------------------------------------------

  const headers = data[0];
  const idIdx = headers.indexOf('รหัส');
  
  let totalPlans = 0;
  let totalBudget = 0;
  let reqBudget = 0;
  let remBudget = 0;
  let typeCount = {};
  let statusCount = { 'สำเร็จ': 0, 'ปรับปรุง': 0 };
  let planStatusCount = { 'ยังไม่ได้ดำเนินการ': 0, 'กำลังดำเนินการ': 0, 'ดำเนินการแล้ว': 0 };
  let planOptions = []; 

  for(let i=1; i<data.length; i++) {
    if(idIdx > -1 && (!data[i][idIdx] || data[i][idIdx].toString().trim() === '')) continue;

    let row = {};
    headers.forEach((h, j) => row[h] = data[i][j]);

    if(filterDeptCode !== 'ALL') {
         let deptStr = (row['ฝ่าย'] || '').toString();
         let idStr = (row['รหัส'] || '').toString();
         let isMatch = (deptStr === filterDeptCode) || (idStr.includes('-' + filterDeptCode + '-'));
         if(!isMatch) continue;
    }

    if (!planOptions.some(p => p.id === row['รหัส'])) {
        planOptions.push({ id: row['รหัส'], name: row['รายการ'] });
    }

    if(filterPlanId && filterPlanId !== 'ALL' && filterPlanId !== '') {
        if(row['รหัส'] !== filterPlanId) continue;
    }

    totalPlans++;
    totalBudget += parseFloat((row['งบประมาณ']||'0').toString().replace(/,/g, '').trim()) || 0;
    reqBudget += parseFloat((row['งบประมาณที่ขอใช้']||'0').toString().replace(/,/g, '').trim()) || 0;
    remBudget += parseFloat((row['งบประมาณคงเหลือ']||'0').toString().replace(/,/g, '').trim()) || 0;

    let pType = row['ประเภทแผน'] || 'ไม่ระบุ';
    typeCount[pType] = (typeCount[pType] || 0) + 1;

    let statusStr = (row['สถานะ'] || '').toString().trim();
    let resultScore = (row['บรรลุผล'] || '').toString().trim();

    if(statusStr === 'สำเร็จ' || statusStr === 'ผ่าน' || resultScore === '1' || resultScore === '1.00') {
      statusCount['สำเร็จ']++;
    } else if (statusStr !== '') {
      statusCount['ปรับปรุง']++;
    }

    // คำนวณสถานะดำเนินงาน
    let pId = (row['รหัส'] || '').toString().trim();
    if (completedSet.has(pId)) {
        planStatusCount['ดำเนินการแล้ว']++;
    } else if (inProgressSet.has(pId)) {
        planStatusCount['กำลังดำเนินการ']++;
    } else {
        planStatusCount['ยังไม่ได้ดำเนินการ']++;
    }
  }

  let typeArr = Object.keys(typeCount).map(k => ({name: k, count: typeCount[k]}));

  return { totalPlans, totalBudget, reqBudget, remBudget, typeArr, statusCount, planStatusCount, planOptions };
}

function fetchData(sheetName, role, deptCode) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "' + sheetName + '"');
  
  const data = sheet.getDataRange().getDisplayValues();
  if(data.length < 2) return [];
  
  const headers = data[0];
  let result = [];

  const startRow = 1; 
  for (let i = startRow; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().includes('รวม')) continue;
    let rowObj = { 'rowNumber': i + 1 };
    headers.forEach((h, j) => rowObj[h] = data[i][j]);
    
    if (deptCode === '') continue;
    
    if (role === 'user') {
        let deptStr = (rowObj['ฝ่าย'] || '').toString().trim();
        let idStr = (rowObj['รหัส'] || rowObj['รหัสแผน'] || '').toString().trim();
        
        let isMatch = false;
        if (deptCode && (deptStr === deptCode || deptStr.includes(deptCode))) {
            isMatch = true;
        }
        if (deptCode && (idStr.includes(deptCode) || idStr.includes(`-${deptCode}-`))) {
            isMatch = true;
        }
        if (!isMatch) continue;
    } else if (deptCode !== 'ALL') {
        let idStr = (rowObj['รหัส'] || rowObj['รหัสแผน'] || '').toString();
        if (deptCode && !idStr.includes(`-${deptCode}-`) && !idStr.includes(deptCode)) continue;
    }
    
    result.push(rowObj);
  }
  return result;
}

function saveData(sheetName, dataObj, rowNumber) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0]; 

    if (rowNumber) {
      headers.forEach((h, idx) => { if (dataObj[h] !== undefined) sheet.getRange(rowNumber, idx + 1).setValue(dataObj[h]); });
    } else {
      let rowData = headers.map(h => dataObj[h] || '');
      sheet.appendRow(rowData);
      const newRow = sheet.getLastRow();
      if (newRow > 2) { 
        sheet.getRange(newRow - 1, 1, 1, headers.length).copyTo(sheet.getRange(newRow, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
        headers.forEach((h, idx) => { if (dataObj[h] !== undefined) sheet.getRange(newRow, idx+1).setValue(dataObj[h]); });
      }
    }
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function deleteData(sheetName, rowNumber) {
  try {
    getSpreadsheet().getSheetByName(sheetName).deleteRow(rowNumber);
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function getPlanById(planId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Plan');
  if(!sheet) return null;
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0]; 
  const idIdx = headers.indexOf('รหัส') > -1 ? headers.indexOf('รหัส') : 0;
  
  for(let i=1; i<data.length; i++) { 
    if(data[i][idIdx] === planId) {
      let obj = {};
      for(let j=0; j<headers.length; j++) {
        if(headers[j]) obj[headers[j]] = data[i][j];
      }
      
      let originalBudget = parseFloat((obj['งบประมาณ']||'0').toString().replace(/,/g, '')) || 0;
      obj['งบประมาณ'] = originalBudget;
      
      let latestRemaining = originalBudget;
      const outcomeSheet = ss.getSheetByName('outcome');
      if(outcomeSheet) {
        const outcomeData = outcomeSheet.getDataRange().getDisplayValues();
        const outHeaders = outcomeData[0];
        const outIdIdx = outHeaders.indexOf('รหัส');
        const outReqIdx = outHeaders.indexOf('งบประมาณที่ขอใช้'); 
        const outRemIdx = outHeaders.indexOf('งบประมาณคงเหลือ');
        
        if(outIdIdx > -1) {
          for(let k=1; k<outcomeData.length; k++) {
            if(outcomeData[k][outIdIdx] === planId) {
              if (outRemIdx > -1 && outcomeData[k][outRemIdx] !== '') {
                latestRemaining = parseFloat((outcomeData[k][outRemIdx]||'0').toString().replace(/,/g, '')) || 0;
              } else if (outReqIdx > -1) {
                let reqAmt = parseFloat((outcomeData[k][outReqIdx]||'0').toString().replace(/,/g, '')) || 0;
                latestRemaining = originalBudget - reqAmt;
              }
              break;
            }
          }
        }
      }
      
      obj['งบประมาณคงเหลือล่าสุด'] = latestRemaining;
      return obj;
    }
  }
  return null;
}

function getBudgetHistory(type) {
  if (type === 'req') {
    return fetchData('BudgetReq', 'admin', 'ALL');
  } else if (type === 'act') {
    const reqData = fetchData('BudgetReq', 'admin', 'ALL');
    const actData = fetchData('BudgetActual', 'admin', 'ALL');
    
    let actList = [...actData];
    return reqData.map(req => {
      let reqAmt = parseFloat((req['จำนวนเงินที่ขอ']||'0').toString().replace(/,/g,'')) || 0;
      
      let actIdx = actList.findIndex(a => 
        (a['ID'] && req['ID'] && a['ID'] === req['ID']) || 
        (a['รหัสแผน'] === req['รหัสแผน'] && (parseFloat((a['จำนวนเงินที่ขอ']||'0').toString().replace(/,/g,'')) || 0) === reqAmt)
      );
      
      if (actIdx > -1) {
        let act = actList.splice(actIdx, 1)[0];
        return {
          ...act, 
          ...req, 
          'งบประมาณใช้จริง': act['งบประมาณใช้จริง'], 
          isActual: true, 
          actRowNumber: act.rowNumber,
          reqRowNumber: req.rowNumber
        };
      } else {
        return {
          ...req,
          isActual: false,
          reqRowNumber: req.rowNumber
        };
      }
    });
  }
  return [];
}

function saveBudgetRecord(type, dataObj, rowNumber) {
  return saveData(type === 'req' ? 'BudgetReq' : 'BudgetActual', dataObj, rowNumber);
}

function markResultReceived(rowNumber) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Result');
    const headers = sheet.getDataRange().getValues()[0];
    
    let colIdx = headers.indexOf('ได้รับเอกสาร');
    
    if (colIdx === -1) {
        colIdx = headers.length + 1;
        sheet.getRange(1, colIdx).setValue('ได้รับเอกสาร');
    } else {
        colIdx = colIdx + 1;
    }
    
    sheet.getRange(rowNumber, colIdx).setValue('ได้รับแล้ว');
    return { status: 'success' };
  } catch(e) { 
    return { status: 'error', message: e.toString() }; 
  }
}

// -----------------------------------------------------------------
// ฟังก์ชันใหม่: ดึงข้อมูลสถานะแผนงานแบบตาราง
// -----------------------------------------------------------------
function getPlanStatus(role, deptCode) {
  const ss = getSpreadsheet();
  const planSheet = ss.getSheetByName('Plan');
  const reqSheet = ss.getSheetByName('BudgetReq');
  const resultSheet = ss.getSheetByName('Result');

  if (!planSheet) return [];

  // 1. ดึงรหัสแผนที่ "กำลังดำเนินการ" (มีบันทึกใน BudgetReq)
  let inProgressSet = new Set();
  if (reqSheet) {
    const reqData = reqSheet.getDataRange().getDisplayValues();
    const reqHeaders = reqData[0];
    const planIdIdx = reqHeaders.indexOf('รหัสแผน');
    if (planIdIdx > -1) {
      for (let i = 1; i < reqData.length; i++) {
        if (reqData[i][planIdIdx] && reqData[i][planIdIdx].toString().trim() !== '') {
          inProgressSet.add(reqData[i][planIdIdx].toString().trim());
        }
      }
    }
  }

  // 2. ดึงรหัสแผนที่ "ดำเนินการแล้ว" (ส่งเอกสารใน Result และเจ้าหน้าที่ติ๊กรับแล้ว)
  let completedSet = new Set();
  if (resultSheet) {
    const resData = resultSheet.getDataRange().getDisplayValues();
    const resHeaders = resData[0];
    const planIdIdx = resHeaders.indexOf('รหัส');
    const recIdx = resHeaders.indexOf('ได้รับเอกสาร');
    if (planIdIdx > -1 && recIdx > -1) {
      for (let i = 1; i < resData.length; i++) {
        if (resData[i][recIdx] === 'ได้รับแล้ว' && resData[i][planIdIdx]) {
          completedSet.add(resData[i][planIdIdx].toString().trim());
        }
      }
    }
  }

  // 3. ประมวลผลจากชีต Plan หลัก เพื่อหาว่าแต่ละแผนอยู่สถานะไหน
  const planData = planSheet.getDataRange().getDisplayValues();
  if (planData.length < 2) return [];
  const pHeaders = planData[0];
  const idIdx = pHeaders.indexOf('รหัส');
  const nameIdx = pHeaders.indexOf('รายการ');
  const deptIdx = pHeaders.indexOf('ฝ่าย');

  let result = [];
  for (let i = 1; i < planData.length; i++) {
    if (!planData[i][0] || planData[i][0].toString().includes('รวม')) continue;
    
    let pId = (planData[i][idIdx] || '').toString().trim();
    let pName = planData[i][nameIdx] || '';
    let pDept = (planData[i][deptIdx] || '').toString().trim();
    
    // กรองข้อมูลแผนก ตามสิทธิ์ของผู้ใช้และ Filter
    if (deptCode && deptCode !== 'ALL') {
      let isMatch = false;
      if (role === 'user') {
          if (pDept === deptCode || pDept.includes(deptCode)) isMatch = true;
          if (pId.includes(deptCode) || pId.includes(`-${deptCode}-`)) isMatch = true;
      } else {
          if (pId.includes(`-${deptCode}-`) || pId.includes(deptCode)) isMatch = true;
      }
      if (!isMatch) continue;
    }

    // กำหนดเงื่อนไขสถานะแผน
    let status = 'ยังไม่ได้ดำเนินการ';
    if (completedSet.has(pId)) {
      status = 'ดำเนินการแล้ว';
    } else if (inProgressSet.has(pId)) {
      status = 'กำลังดำเนินการ';
    }

    result.push({
      id: pId,
      name: pName,
      dept: pDept,
      status: status
    });
  }
  
  return result;
}

// -----------------------------------------------------------------
// ฟังก์ชันใหม่: ดึงข้อมูลสรุปงบประมาณ
// -----------------------------------------------------------------
function getBudgetSummary(role, deptCode) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('outcome');
  if (!sheet) return [];

  const data = sheet.getDataRange().getDisplayValues();
  if(data.length < 2) return [];

  const headers = data[0];
  let result = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().includes('รวม')) continue;
    
    let rowObj = {};
    headers.forEach((h, j) => rowObj[h] = data[i][j]);

    let pId = (rowObj['รหัส'] || '').toString().trim();
    let pDept = (rowObj['ฝ่าย'] || '').toString().trim();

    // ตัวกรองฝ่าย
    if (deptCode && deptCode !== 'ALL') {
      let isMatch = false;
      if (role === 'user') {
          if (pDept === deptCode || pDept.includes(deptCode)) isMatch = true;
          if (pId.includes(deptCode) || pId.includes(`-${deptCode}-`)) isMatch = true;
      } else {
          if (pId.includes(`-${deptCode}-`) || pId.includes(deptCode)) isMatch = true;
      }
      if (!isMatch) continue;
    }

    result.push({
      id: pId,
      name: rowObj['รายการ'] || '-',
      budget: parseFloat((rowObj['งบประมาณ']||'0').toString().replace(/,/g, '')) || 0,
      latestReq: parseFloat((rowObj['งบประมาณที่ขอใช้']||'0').toString().replace(/,/g, '')) || 0,
      remaining: parseFloat((rowObj['งบประมาณคงเหลือ']||'0').toString().replace(/,/g, '')) || 0
    });
  }
  return result;
}