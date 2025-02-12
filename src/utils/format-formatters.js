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

/* 将选中区域的日期格式设置为 yyyy-m-d */
function setDateFormatToYMD() {
    let rangedate = Application.Selection;
    rangedate.NumberFormat = 'yyyy-m-d';
}

/* 将选中区域的日期格式设置为 yyyy-m */
function setDateFormatToYM() {
    let rangedate = Application.Selection;
    rangedate.NumberFormat = 'yyyy-m';
}

/* 设置选中区域的单元格为缩小字体填充 */
function setShrinkToFit() {
    var selection = Application.Selection;
    selection.Cells.HorizontalAlignment = xlHAlignCenter;
    selection.Cells.VerticalAlignment = xlVAlignCenter;
    selection.Cells.ReadingOrder = -5002;
    selection.Cells.AddIndent = false;
    selection.Cells.IndentLevel = 0;
    selection.Cells.WrapText = false;
    selection.Cells.ShrinkToFit = true;
}

/**
 * 批量修改所有工作表的行高和列宽
 */
function setAllSheetsRowHeightAndColumnWidth(height, width) {
    try {
        var workbook = Application.ActiveWorkbook;
        for (var i = 1; i <= workbook.Sheets.Count; i++) {
            var sheet = workbook.Sheets(i);
            sheet.UsedRange.Rows.RowHeight = height;
            sheet.UsedRange.Columns.ColumnWidth = width;
        }
    } catch (error) {
        // 忽略错误
    }
}

/**
 * 批量设置字体为微软雅黑，字号为9
 */
function setFontToMicrosoftYaHei() {
    var workbook = Application.ActiveWorkbook;
    for (var i = 1; i <= workbook.Sheets.Count; i++) {
        workbook.Sheets.Item(i).Activate();
        Application.Cells.Select();
        Selection.Font.Name = "微软雅黑";
        Selection.Font.Size = 9;
        Application.Range("A1").Select();
    }
    workbook.Sheets.Item(1).Activate();
}
