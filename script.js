const k_date_index = 2;
const k_default_offset = 2;
const k_num_doc_sheets = 2; // yyyy/mmでないシートの数．submitsとsummary
const k_submits_sheet_id = 0;
const k_summary_sheet_id = 1;
const k_default_category_pos = 'H2';
const k_default_price_pos = 'I2';
const k_default_num_category_pos = 'J2';
const k_default_total_price_pos = 'K2';

function onSubmit() {
  const sheet_obj = SpreadsheetApp.getActiveSpreadsheet(); // スプレッドシートオブジェクトを得る
  
  insertSummarySheet(sheet_obj);
  
  var sheet_submits = sheet_obj.getSheetByName("submits");
  var lastRow = sheet_submits.getLastRow(); // submitされたデータの最後尾
  
  var row = lastRow; // 最新のデータの行
  var col = k_date_index; // 日付の行
  var range = sheet_submits.getRange(row, col);
  var date = Utilities.formatDate(range.getValue(), 'JST', 'yyyy/MM');
  
  var {sheet_id, create} = findInsertSheetInd(sheet_obj, date); // 得られたdateがどの月のシートに入るかを求める．createする必要があるならsheet_idで新規作成
  validateSheetDate(sheet_obj, date, sheet_id, create);
  
  form_data = getRowValues(sheet_submits, row, sheet_submits.getMaxColumns()); // 最新データの行のデータを取得
  
  insertFormData2SheetWithSortByDate(sheet_obj, sheet_id, form_data); // 最新のデータをsheet_idのシートにソートしてappendする．
  
  updateSheetStat(sheet_obj, sheet_id, form_data);　// sheet_idのsheetのstatをアップデートする
  
  register2Summary(sheet_obj, sheet_id); // sheet_idの更新したstatをsummaryページに登録
}

function insertSummarySheet(sheet_obj) {
  var sheet = sheet_obj.getSheets();
  if (sheet.length == 1) {
    sheet_obj.insertSheet("summary", k_summary_sheet_id);
    initSummarySheet(sheet_obj);
  }
}

function initSummarySheet(sheet_obj) {
  var sheet = sheet_obj.getSheets()[k_summary_sheet_id];
  sheet.getRange('A1').setValue('年月');
  sheet.getRange('B1').setValue('合計');
}

function findInsertSheetInd(sheet_obj, date) {
  var sheets = sheet_obj.getSheets();
  
  var initial = (sheets.length == 2)? true : false;
  if (initial) {
    // submitsのシートしかない状態
    return {sheet_id: 2, create: true};
  }
  
  sheet_dates = new Array();

  var pivot = 0;  
  var exists = false;
  
  for (iter = k_num_doc_sheets; iter < sheets.length; iter++) {
    if(sheets[iter].getName() == date) {
      exists = true;
      pivot = iter;
      break;
    }
    sheet_dates.push(new Date(sheets[iter].getName()))
  }
  if (exists) {
    return {sheet_id:pivot, create:false};
  }
  
  date = new Date(date);
  while ( !(date < sheet_dates[pivot])){
    if (pivot == (sheet_dates.length - 1)) {
      pivot = sheet_dates.length;
      break;
    }
    else
      pivot += 1;
  }
  var sheet_id = pivot + k_num_doc_sheets;
  return {sheet_id:sheet_id, create:true}
}

function validateSheetDate(sheet_obj, str_yyyy_mm, sheet_ind, create) {
  if (create) {
    sheet_obj.insertSheet(str_yyyy_mm, sheet_ind);
    initSheet(sheet_obj, sheet_ind);
  }
}

function initSheet(sheet_obj, sheet_ind) {
  // H1に「区分」，I1に「金額」と入力する
  sheet = sheet_obj.getSheets()[sheet_ind];
  sheet.getRange('H1').setValue('区分');
  sheet.getRange('I1').setValue('金額');
  sheet.getRange('J1').setValue('区分数');
  sheet.getRange('K1').setValue('合計金額');
  sheet.getRange('J2').setValue(0);
  sheet.getRange('K2').setValue(0);
}

function getRowValues(sheet_submit, row, num_cols) {
  return sheet_submit.getRange(row, 1, 1, num_cols).getValues();
}

function insertFormData2Sheet(sheet_obj, sheet, sheet_id, form_data, row_ind) {
  sheet.getRange(row_ind, 1, 1, form_data[0].length).setValues(form_data);
}

function insertFormData2SheetWithSortByDate(sheet_obj, sheet_id, form_data) {
  var sheet = sheet_obj.getSheets()[sheet_id];

  var num_rows = 0;
  const head_offset = k_default_offset;
  
  if (isEmptySheet(sheet)) {
    num_rows = 1 + head_offset;
  }
  else {
    // データがあるので，form_dataを挿入すべきnum_rowsを求める
    var num_data = sheet.getLastRow()-head_offset;
    var pivot = 1 + head_offset; // pivot行目の上に挿入する
    var form_day = form_data[0][1].getDate();
    var cnt_day = sheet.getRange(pivot, 2).getValue().getDate();
    while (form_day > cnt_day && ((pivot-head_offset) <= num_data)) {
      pivot += 1;
      if ((pivot - head_offset) == (num_data + 1)) {
        break;
      }
      cnt_day = sheet.getRange(pivot, 2).getValue().getDate();
    }
    num_rows = pivot;
  }
  
  sheet.insertRowBefore(num_rows); //pivot行目の上に行を挿入する
  insertFormData2Sheet(sheet_obj, sheet, sheet_id, form_data, num_rows); // 新たなpivot行にデータを挿入
}

function isEmptySheet(sheet) {
  return sheet.getLastRow() == k_default_offset;
}

function updateSheetStat(sheet_obj, sheet_id, form_data) {
  sheet = sheet_obj.getSheets()[sheet_id];
  
  var category = form_data[0][2];
  var price = form_data[0][4];
  
  // 新しいシートならカテゴリーを新規作成
  if (isNewSheet(sheet)) {
    addNewCategory(sheet, category, 0, k_default_category_pos, k_default_price_pos);
  }
  // そのカテゴリーがあるか調べる
  var {found, pos} = categoryExists(sheet, category);  
  // あればそこに合計を加算
  if (found)  {
    // H列からI列に
    price_pos = 'I' + pos.charAt(1);
    addPrice(sheet, price_pos, price);
  }
  else {
    price_pos = 'I' + pos.charAt(1);
    addNewCategory(sheet, category, price, pos, price_pos);
  }
  // 合計金額
  addTotalPrice(sheet, price);
}

function categoryExists(sheet, category) {
  var num_categories = sheet.getRange(k_default_num_category_pos).getValue();
  found = false;
  pos = k_default_category_pos;
  
  for (var iter = 1; iter <= num_categories; iter++) {
    if (sheet.getRange(pos).getValue() == category) {
      found = true;
      break;
    }
    pos = pos.charAt(0) + String(pos.charAt(1) * 1 + 1);
  }
  return {found:found, pos:pos};
}

function isNewSheet(sheet) {
  return sheet.getRange(k_default_category_pos).isBlank();
}

function addNewCategory(sheet, category, price, category_pos, price_pos) {
  sheet.getRange(category_pos).setValue(category);
  sheet.getRange(price_pos).setValue(price);
  var  num_category = sheet.getRange(k_default_num_category_pos).getValue()
  sheet.getRange(k_default_num_category_pos).setValue(num_category+1);
}

function addPrice(sheet, price_pos, price) {
  cur_price = sheet.getRange(price_pos).getValue();
  sheet.getRange(price_pos).setValue(cur_price + price);
}

function addTotalPrice(sheet, price) {
  cur_price = sheet.getRange(k_default_total_price_pos).getValue();
  sheet.getRange(k_default_total_price_pos).setValue(cur_price + price);
}

function register2Summary(sheet_obj, sheet_id) {
  sheet = sheet_obj.getSheets()[sheet_id];
  
}
