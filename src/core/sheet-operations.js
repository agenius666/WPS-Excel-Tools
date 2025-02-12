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

/* 获取所有工作表的名称并插入到指定位置 */
function getSheetNames(skipHiddenSheets, insertHere) {
    let IH = typeof (insertHere) == "string" ? Range(insertHere) : insertHere;
    let z = [];
    for (let i = 1; i <= Worksheets.Count; i++) {
        let iSh = Worksheets.Item(i);
        if (!skipHiddenSheets && !iSh.Visible) {
            z.push(iSh.Name);
        }
        if (iSh.Visible) {
            z.push(iSh.Name);
        }
    }
    for (let s = 0; s < z.length; s++) {
        IH.Offset(s, 0).Value2 = z[s];
    }
}

/* 创建新工作表并插入工作表名称 */
function createSheetWithNames() {
    try {
        if (typeof createSheetWithNames.counter === 'undefined') {
            createSheetWithNames.counter = 1;
        }
        let newSheetName = "工作表名称" + createSheetWithNames.counter;
        let nsh = Sheets.Add();
        nsh.Name = newSheetName;
        getSheetNames(true, newSheetName + "!A1");
        createSheetWithNames.counter++;
    } catch (error) {
        // 忽略错误
    }
}

/**
 * 删除工作簿中的所有水印（Shapes）
 */
function removeWorkbookWatermarks() {
    var workbook = Application.ActiveWorkbook;
    for (var i = 1; i <= workbook.Sheets.Count; i++) {
        var sheet = workbook.Sheets(i);
        var shapes = sheet.Shapes;
        for (var j = shapes.Count; j > 0; j--) {
            shapes.Item(j).Delete();
        }
    }
}

/**
 * 调整合并单元格的高度
 */
function adjustMergedCellHeight() {
    let mergedCells = {};
    for (let cell of ActiveSheet.UsedRange) {
        if (cell.MergeCells) {
            let mergeAreaAddress = cell.MergeArea.Address();
            if (!(mergeAreaAddress in mergedCells)) {
                mergedCells[mergeAreaAddress] = 1;
                if (cell.MergeArea.Rows.Count === 1) {
                    let totalWidth = 0;
                    for (let mergedCell of cell.MergeArea) {
                        totalWidth += mergedCell.ColumnWidth;
                    }
                    cell.RowHeight = calculateRowHeight(cell.Value2, totalWidth);
                }
            }
        }
    }
}

/**
 * 计算行高
 * @param {string} text - 单元格内容
 * @param {number} width - 单元格宽度
 * @returns {number} - 计算后的行高
 */
function calculateRowHeight(text, width) {
    Application.ScreenUpdating = false;
    Application.DisplayAlerts = false;
    let tempSheet = Worksheets.Add();
    let tempCell = tempSheet.Range("A1");
    tempCell.Value2 = text;
    tempCell.ColumnWidth = width;
    tempCell.WrapText = true;
    tempCell.EntireRow.AutoFit();
    let rowHeight = tempCell.RowHeight;
    tempSheet.Delete();
    Application.DisplayAlerts = true;
    Application.ScreenUpdating = true;
    return rowHeight;
}

/**
 * 设置所有工作表的打印区域为已使用区域
 */
function setPrintAreaToUsedRange() {
    var sheets = Application.ActiveWorkbook.Sheets;
    for (var i = 1; i <= sheets.Count; i++) {
        var sheet = sheets.Item(i);
        sheet.PageSetup.PrintArea = sheet.UsedRange.Address();
    }
}

/**
 * 移除工作簿的密码保护
 */
function removeWorkbookPassword() {
    ActiveWorkbook.SetPasswordEncryptionOptions("", "OfficeStandard", 0, false);
    ActiveWorkbook.Password = "";
    ActiveWorkbook.WritePassword = "";
}

/**
 * 隐藏空白工作表
 */
function hideBlankSheets() {
    for (var sheet of Sheets) {
        if (WorksheetFunction.CountA(sheet.UsedRange) == 0) {
            sheet.Visible = false;
        } else {
            sheet.Visible = true;
        }
    }
}

/**
 * 隐藏无值工作表
 */
function hideSheetsWithNoValues() {
    for (var sheet of Sheets) {
        try {
            if (WorksheetFunction.Sum(sheet.UsedRange) == 0) {
                sheet.Visible = false;
            } else {
                sheet.Visible = true;
            }
        } catch (error) {
            // 忽略错误
        }
    }
}

/* 规范化所有工作表的视图设置 */
function standardizeAllSheetsView() {
    try {
        // 获取当前活动工作簿
        let workbook = Application.ActiveWorkbook;

        // 遍历工作簿中的所有工作表
        for (let i = 1; i <= workbook.Sheets.Count; i++) {
            let sheet = workbook.Sheets(i);

            // 激活当前工作表
            sheet.Activate();

            // 确保当前窗口存在
            if (Application.ActiveWindow) {
                // 设置缩放比例为 100%
                Application.ActiveWindow.Zoom = 100;

                // 激活 A1 单元格
                sheet.Range("A1").Activate();
            } else {
                // 如果没有 ActiveWindow 对象，跳过该工作表并记录日志
                console.log(`跳过工作表: ${sheet.Name} (无 ActiveWindow 对象)`);
            }
        }

        // 返回到第一个工作表
        workbook.Sheets(1).Activate();
    } catch (error) {
        // 捕获错误并记录日志
        console.error("发生错误: ", error.message);
    }
}

/* 根据输入的工作表名选择工作表 */
function selectSheetsByName() {
    t = Application.InputBox("请输入需要筛选的工作表名");
    var sheets = Application.ActiveWorkbook.Sheets;
    var summarySheets = [];
    for (let i = 1; i <= Worksheets.Count; i++) {
        let iSh = Worksheets.Item(i);
        if (iSh.Visible && iSh.Name.indexOf(t) > -1) {
            summarySheets.push(Worksheets.Item(i).Name);
        }
    }
    if (summarySheets.length > 0) {
        Application.ActiveWorkbook.Sheets.Item(summarySheets[0]).Activate();
        for (var i = 1; i < summarySheets.length; i++) {
            Application.ActiveWorkbook.Sheets.Item(summarySheets[i]).Select(false);
        }
    } else {
        alert("没有找到指定的工作表");
    }
}

/* 根据输入的工作表名排除选择工作表 */
function selectSheetsExcludingName() {
    t = Application.InputBox("请输入需要排除筛选的工作表名");
    var sheets = Application.ActiveWorkbook.Sheets;
    var summarySheets = [];
    for (let i = 1; i <= Worksheets.Count; i++) {
        let iSh = Worksheets.Item(i);
        if (iSh.Visible && iSh.Name.indexOf(t) === -1) {
            summarySheets.push(Worksheets.Item(i).Name);
        }
    }
    if (summarySheets.length > 0) {
        Application.ActiveWorkbook.Sheets.Item(summarySheets[0]).Activate();
        for (var i = 1; i < summarySheets.length; i++) {
            Application.ActiveWorkbook.Sheets.Item(summarySheets[i]).Select(false);
        }
    } else {
        ("没有找到不包含指定内容的工作表");
    }
}

/* 取消所有工作表的保护 */
function unprotectAllSheets() {
    try {
        t = Application.InputBox("请输入工作表的保护密码\n所有工作表的密码必须一致，或没有；\n没有密码则不填");
        var sheets = Application.ActiveWorkbook.Sheets;
        for (let i = 1; i <= Worksheets.Count; i++) {
            let iSh = Worksheets.Item(i);
            iSh.Unprotect(t);
        }
    } catch (error) {
        // 忽略错误
    }
}
