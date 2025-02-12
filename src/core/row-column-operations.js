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

/* 删除当前工作表中的空白行 */
function deleteBlankRows() {
    try {
        var sheet = Application.ActiveSheet();
        var lastRow = Application.ActiveSheet.UsedRange.Rows.Count;
        for (var i = lastRow; i >= 1; i--) {
            if (WorksheetFunction.CountA(Rows(i)) === 0) {
                Application.Range(i + ":" + i).EntireRow.Delete();
            }
        }
    } catch (error) {
        // 忽略错误
    }
}

/* 删除当前工作表中和为0的行 */
function deleteRowsWithZeroSum() {
    try {
        var sheet = Application.ActiveSheet();
        var lastRow = Application.ActiveSheet.UsedRange.Rows.Count;
        for (var i = lastRow; i >= 1; i--) {
            if (WorksheetFunction.Sum(Rows(i)) === 0) {
                Application.Range(i + ":" + i).EntireRow.Delete();
            }
        }
    } catch (error) {
        // 忽略错误
    }
}

/* 设置所有工作表的行高为15，列宽为18 */
function setUniformRowHeightAndColumnWidthForAllSheets() {
    try {
        let rowHeight = 15; // 定义行高为15
        let columnWidth = 18; // 定义列宽为18
        let workbook = Application.ActiveWorkbook;

        // 遍历所有工作表
        for (let i = 1; i <= workbook.Sheets.Count; i++) {
            let sheet = workbook.Sheets(i);
            sheet.UsedRange.Rows.RowHeight = rowHeight; // 设置行高
            sheet.UsedRange.Columns.ColumnWidth = columnWidth; // 设置列宽
        }
    } catch (error) {
        // 忽略错误
    }
}

/* 设置当前工作表的行高为15，列宽为18 */
function setUniformRowHeightAndColumnWidthForActiveSheet() {
    try {
        let rowHeight = 15; // 定义行高为15
        let columnWidth = 18; // 定义列宽为18

        // 获取当前活动的工作表
        let sheet = Application.ActiveSheet;

        // 设置当前工作表的行高和列宽
        sheet.UsedRange.Rows.RowHeight = rowHeight;
        sheet.UsedRange.Columns.ColumnWidth = columnWidth;
    } catch (error) {
        // 忽略错误
    }
}

/* 批量设置所有工作表的行高和列宽 */
function batchSetRowHeightAndColumnWidthForAllSheets() {
    try {
        // 提示用户输入行高和列宽
        let rowHeight = InputBox("请输入行高", "行高");
        let columnWidth = InputBox("请输入列宽", "列宽");

        let workbook = Application.ActiveWorkbook;

        // 遍历所有工作表
        for (let i = 1; i <= workbook.Sheets.Count; i++) {
            let sheet = workbook.Sheets(i);
            sheet.UsedRange.Rows.RowHeight = rowHeight; // 设置行高
            sheet.UsedRange.Columns.ColumnWidth = columnWidth; // 设置列宽
        }
    } catch (error) {
        // 忽略错误
    }
}

/* 批量设置当前工作表的行高和列宽 */
function batchSetRowHeightAndColumnWidthForActiveSheet() {
    try {
        // 提示用户输入行高和列宽
        let rowHeight = InputBox("请输入行高", "行高");
        let columnWidth = InputBox("请输入列宽", "列宽");

        // 获取当前活动的工作表
        let sheet = Application.ActiveSheet;

        // 设置当前工作表的行高和列宽
        sheet.UsedRange.Rows.RowHeight = rowHeight;
        sheet.UsedRange.Columns.ColumnWidth = columnWidth;
    } catch (error) {
        // 忽略错误
    }
}

/* 自动调整当前工作表的所有列宽 */
function autoFitColumnsInActiveSheet() {
    try {
        // 获取当前活动工作表
        let sheet = Application.ActiveSheet;

        // 自动调整所有列的宽度
        sheet.Columns.AutoFit();
    } catch (error) {
        // 忽略错误
    }
}
