/**
 * ====================================================
 *  マイレシピ帳 — Google Apps Script バックエンド
 * ====================================================
 *
 * 【使い方】
 *  1. Googleスプレッドシートを開く
 *  2. 拡張機能 > Apps Script を開く
 *  3. このファイルの内容をすべてコピーして貼り付け
 *  4. 保存（Ctrl+S）
 *  5. デプロイ > 新しいデプロイ > ウェブアプリ
 *     ・アクセスできるユーザー：「全員」に設定
 *  6. デプロイ > ウェブアプリURLをコピー
 *  7. アプリの「設定」画面にURLを貼り付けて保存
 */

const SHEET_NAME = 'レシピ';
const HEADERS    = ['id', 'name', 'category', 'ingredients', 'steps', 'memo', 'createdAt'];

// ─────────────────────────────────────
//  エントリーポイント（GETリクエスト）
// ─────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';
  let result;

  try {
    switch (action) {
      case 'getAll': result = getAllRecipes();            break;
      case 'add':    result = addRecipe(e.parameter);    break;
      case 'update': result = updateRecipe(e.parameter); break;
      case 'delete': result = deleteRecipe(e.parameter.id); break;
      default:       result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────
//  シートの取得（なければ自動作成）
// ─────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // ヘッダー行を追加
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#FF8C42')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    // 列幅を調整
    sheet.setColumnWidth(1, 160); // id
    sheet.setColumnWidth(2, 140); // name
    sheet.setColumnWidth(3, 100); // category
    sheet.setColumnWidth(4, 220); // ingredients
    sheet.setColumnWidth(5, 260); // steps
    sheet.setColumnWidth(6, 180); // memo
    sheet.setColumnWidth(7, 140); // createdAt
  }

  return sheet;
}

// ─────────────────────────────────────
//  全レシピ取得
// ─────────────────────────────────────
function getAllRecipes() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return { recipes: [] };

  const values = sheet
    .getRange(2, 1, lastRow - 1, HEADERS.length)
    .getValues();

  const recipes = values
    .filter(row => row[0] !== '')   // 空行を除外
    .map(row => {
      const obj = {};
      HEADERS.forEach((h, i) => { obj[h] = String(row[i] ?? ''); });
      return obj;
    });

  return { recipes };
}

// ─────────────────────────────────────
//  レシピ追加
// ─────────────────────────────────────
function addRecipe(params) {
  const sheet = getSheet();
  const id    = String(Date.now());
  const now   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  sheet.appendRow([
    id,
    params.name        || '',
    params.category    || 'その他',
    params.ingredients || '',
    params.steps       || '',
    params.memo        || '',
    now
  ]);

  return { success: true, id };
}

// ─────────────────────────────────────
//  レシピ更新
// ─────────────────────────────────────
function updateRecipe(params) {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return { success: false, error: 'No data' };

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(params.id)) {
      const rowNum = i + 2;
      const existing = sheet.getRange(rowNum, 1, 1, HEADERS.length).getValues()[0];

      sheet.getRange(rowNum, 1, 1, HEADERS.length).setValues([[
        params.id,
        params.name        !== undefined ? params.name        : existing[1],
        params.category    !== undefined ? params.category    : existing[2],
        params.ingredients !== undefined ? params.ingredients : existing[3],
        params.steps       !== undefined ? params.steps       : existing[4],
        params.memo        !== undefined ? params.memo        : existing[5],
        existing[6]  // createdAt はそのまま
      ]]);

      return { success: true };
    }
  }

  return { success: false, error: 'Recipe not found' };
}

// ─────────────────────────────────────
//  レシピ削除
// ─────────────────────────────────────
function deleteRecipe(id) {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return { success: false, error: 'No data' };

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }

  return { success: false, error: 'Recipe not found' };
}
