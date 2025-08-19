/**
 * @fileoverview
 * 實習志願分發程式 (V2 - 強化偵錯版)
 * 增強了志願代碼的解析能力，並在流程中加入更多防呆檢查。
 */

// --- 全域設定 ---
const FORM_RESPONSES_SHEET = '表單回應';
const SLOTS_SHEET = '實習名額表';
const ORDER_SHEET = '選填順序表';
const RESULTS_SHEET = '最終分發結果';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('自動分發工具')
    .addItem('>> 開始執行實習分發 (V2) <<', 'runAllocationById_V2')
    .addToUi();
}

function runAllocationById_V2() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('警告', '即將使用「學號」進行分發 (V2)，這會覆蓋「最終分發結果」的內容。確定嗎？', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    ui.alert('操作已取消。');
    return;
  }

  ui.alert('分發程序已開始...');

  try {
    // 步驟 1: 讀取資料
    const studentChoicesData = getDataFromSheet(FORM_RESPONSES_SHEET, 2, 8);
    const selectionOrderData = getDataFromSheet(ORDER_SHEET, 1, 3);
    const internshipSlotsData = getDataFromSheet(SLOTS_SHEET, 1, 2);

    if (studentChoicesData.length === 0 || selectionOrderData.length === 0 || internshipSlotsData.length === 0) {
      ui.alert('錯誤！', '一個或多個資料工作表是空的（沒有讀取到資料），請檢查工作表名稱是否正確，以及內容是否已填寫。', ui.ButtonSet.OK);
      return;
    }

    // 步驟 2: 建立資料 Map
    const choicesMap = new Map(studentChoicesData.map(row => [row[0].toString().trim(), row.slice(2)])); // key: 學號
    const orderMap = new Map(selectionOrderData.map(row => [row[1].toString().trim(), row[0]]));       // key: 學號
    const slotsMap = new Map(internshipSlotsData.map(row => [row[0].toString().trim(), { capacity: row[1], assigned: 0 }]));

    // 步驟 3: 排序
    const sortedStudents = [...studentChoicesData].sort((a, b) => {
      const studentIdA = a[0].toString().trim();
      const studentIdB = b[0].toString().trim();
      const orderA = orderMap.get(studentIdA) || 999;
      const orderB = orderMap.get(studentIdB) || 999;
      return orderA - orderB;
    });

    // 步驟 4: 執行分發
    const results = [];
    for (const studentRow of sortedStudents) {
      const studentId = studentRow[0].toString().trim();
      const studentName = studentRow[1];
      const preferences = choicesMap.get(studentId) || [];
      let assignedChoice = '所有志願已滿或資料不符，需人工處理';

      for (const choice of preferences) {
        if (!choice) continue;
        
        // ▼▼▼ 修改 ▼▼▼
        // 不論選項格式是 "H01" 還是 "H01 - 機構名稱"，都能準確取出 "H01"
        const choiceCode = choice.toString().split(' - ')[0].trim();
        
        const slot = slotsMap.get(choiceCode);
        if (slot && slot.assigned < slot.capacity) {
          slot.assigned++;
          assignedChoice = choiceCode;
          break;
        }
      }
      results.push([studentId, studentName, assignedChoice]);
    }

    // 步驟 5: 寫回結果
    const resultsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET);
    resultsSheet.clear();
    resultsSheet.getRange(1, 1, 1, 3).setValues([['學號', '姓名', '分發結果']]);
    if (results.length > 0) {
      resultsSheet.getRange(2, 1, results.length, results[0].length).setValues(results);
      ui.alert('分發完成！', `已成功處理 ${results.length} 位學生的資料，請至「最終分發結果」工作表查看。`, ui.ButtonSet.OK);
    } else {
      ui.alert('處理完畢', '沒有可以分發的學生資料。請檢查「フォームの回答」工作表是否有內容。', ui.ButtonSet.OK);
    }

  } catch (e) {
    ui.alert('執行失敗', '發生程式錯誤：' + e.toString(), ui.ButtonSet.OK);
  }
}

function getDataFromSheet(sheetName, startCol, endCol) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const numRows = sheet.getLastRow() - 1;
  const numCols = endCol - startCol + 1;
  return sheet.getRange(2, startCol, numRows, numCols).getValues();
}