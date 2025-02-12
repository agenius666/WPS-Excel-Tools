/*!
 * Copyright 2023 agenius666
 * GitHub: https://github.com/agenius666/WPS-Excel-Tools
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/* 复制当前单元格并向右移动一格 */
function copyCellAndMoveRight() {
    let activeCell = Application.ActiveCell; // 获取当前活动单元格
    activeCell.Copy(); // 复制当前单元格的内容
    activeCell.Offset(0, 1).Activate(); // 向右移动一格并激活该单元格
}

/* 复制当前单元格并向左移动一格 */
function copyCellAndMoveLeft() {
    let activeCell = Application.ActiveCell; // 获取当前活动单元格
    activeCell.Copy(); // 复制当前单元格的内容
    activeCell.Offset(0, -1).Activate(); // 向左移动一格并激活该单元格
}

/* 复制当前单元格并向上移动一格 */
function copyCellAndMoveUp() {
    let activeCell = Application.ActiveCell; // 获取当前活动单元格
    activeCell.Copy(); // 复制当前单元格的内容
    activeCell.Offset(-1, 0).Activate(); // 向上移动一格并激活该单元格
}

/* 复制当前单元格并向下移动一格 */
function copyCellAndMoveDown() {
    let activeCell = Application.ActiveCell; // 获取当前活动单元格
    activeCell.Copy(); // 复制当前单元格的内容
    activeCell.Offset(1, 0).Activate(); // 向下移动一格并激活该单元格
}

/* 批量设置会计格式（不带￥符号） */
function applyAccountingFormatWithoutYen() {
    let selection = Application.Selection; // 获取当前选中的区域
    selection.NumberFormatLocal = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * \"-\"??_ ;_ @_ "; 
    // 设置为会计格式：不带￥符号，保留两位小数，负数用括号表示
}

/* 插入表头模板 */
function insertHeaderTemplate() {
    // 获取当前活动工作表
    let originalActiveSheet = Application.ActiveSheet;
    // 打开指定的表头模板文件
    let workbook = Workbooks.Open("\\表头模板.xlsx");
    let sheet = workbook.Sheets("Sheet1");
    // 在当前工作表的第一行插入 12 行空白行
    originalActiveSheet.Rows("1:12").Insert();
    // 复制模板工作表的前 12 行
    sheet.Rows("1:12").Copy();
    // 将复制的内容粘贴到当前工作表的第一行
    originalActiveSheet.Rows("1:12").PasteSpecial(xlPasteAll);
    // 关闭模板工作簿（不保存更改）
    workbook.Close(false);
}
