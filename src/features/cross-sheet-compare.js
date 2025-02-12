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

/* 跨表对比两个工作簿中的工作表 */
function compareTwoSheetsAcrossWorkbooks() {
    try {
        // 打开第一个文件选择器，用于选择基准工作簿
        let fileDialog1 = Application.FileDialog(msoFileDialogFilePicker);
        fileDialog1.Filters.Clear(); // 清除默认过滤器
        fileDialog1.Filters.Add('Excel文件', '*.xls;*.xlsx;*.xlsm'); // 只显示 Excel 文件
        fileDialog1.Title = '请选择第一个需要对比的工作簿（基准表）';

        // 如果用户未选择任何文件，则提示并退出
        if (fileDialog1.Show() !== -1) {
            alert('你没有选择任何文件！');
            return;
        }

        // 获取用户选择的第一个文件路径，并打开该工作簿
        let filePath1 = fileDialog1.SelectedItems.Item(1);
        let workbook1 = Workbooks.Open(filePath1);

        // 提示用户输入基准工作表的名称
        let sheetName1 = Application.InputBox("请输入第一个需要对比的工作表表名\n（基准表，第二个表要与此表对比）");

        // 打开第二个文件选择器，用于选择对比工作簿
        let fileDialog2 = Application.FileDialog(msoFileDialogFilePicker);
        fileDialog2.Filters.Add('Excel文件', '*.xls;*.xlsx;*.xlsm'); // 只显示 Excel 文件
        fileDialog2.Title = '请选择第二个需要对比的工作簿（对比表）';

        // 如果用户未选择任何文件，则提示并退出
        if (fileDialog2.Show() !== -1) {
            alert('你没有选择任何文件！');
            return;
        }

        // 获取用户选择的第二个文件路径，并打开该工作簿
        let filePath2 = fileDialog2.SelectedItems.Item(1);
        let workbook2 = Workbooks.Open(filePath2);

        // 提示用户输入对比工作表的名称
        let sheetName2 = Application.InputBox("请输入第二个需要对比的工作表表名\n（对比表，用来与基准表对比）");

        // 获取基准工作表和对比工作表
        let sheet1 = workbook1.Worksheets(sheetName1);
        let sheet2 = workbook2.Worksheets(sheetName2);

        // 获取基准工作表的最大行数和列数
        let rowCount = sheet1.Cells(sheet1.Rows.Count, 1).End(xlUp).Row;
        let colCount = sheet1.Cells(1, sheet1.Columns.Count).End(xlToLeft).Column;

        // 遍历每个单元格进行对比
        for (let i = 1; i <= rowCount; i++) {
            for (let j = 1; j <= colCount; j++) {
                // 获取两个工作表中对应单元格的值
                let cellValue1 = sheet1.Cells(i, j).Value;
                let cellValue2 = sheet2.Cells(i, j).Value;

                // 如果两个单元格的值不同，则将差异高亮显示为黄色
                if (cellValue1 !== cellValue2) {
                    sheet1.Cells(i, j).Interior.ColorIndex = 6; // 黄色
                    sheet2.Cells(i, j).Interior.ColorIndex = 6; // 黄色
                }
            }
        }

        console.log("跨表对比完成！");
    } catch (error) {
        // 捕获错误并记录日志
        console.error("发生错误: ", error.message);
    }
}
