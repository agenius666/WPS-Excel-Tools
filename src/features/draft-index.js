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

/* 获取工作表名称列表 */
function getSheetNames(skipHiddenSheets, insertHere) {
    try {
        // 获取插入位置（支持字符串或范围对象）
        let insertRange = typeof (insertHere) === "string" ? Range(insertHere) : insertHere;

        // 初始化一个数组存储工作表名称
        let sheetNames = [];

        // 遍历所有工作表
        for (let i = 1; i <= Worksheets.Count; i++) {
            let sheet = Worksheets.Item(i);

            // 如果需要跳过隐藏工作表且当前工作表不可见，则跳过
            if (!skipHiddenSheets && !sheet.Visible) {
                continue;
            }

            // 添加可见工作表的名称到数组
            if (sheet.Visible) {
                sheetNames.push(sheet.Name);
            }
        }

        // 将工作表名称写入指定位置
        for (let s = 0; s < sheetNames.length; s++) {
            insertRange.Offset(s, 0).Value2 = sheetNames[s];
        }
    } catch (error) {
        console.error("获取工作表名称时发生错误: ", error.message);
    }
}

/* 编写底稿索引 - 生成模板表 */
function generateIndexTemplate() {
    try {
        // 定义一个全局变量来维护计数器
        if (typeof generateIndexTemplate.counter === 'undefined') {
            generateIndexTemplate.counter = 1;
        }

        // 生成新的工作表名称，带有序号
        let newSheetName = "半自动索引" + generateIndexTemplate.counter;

        // 新建工作表并设置名称
        let newSheet = Sheets.Add();
        newSheet.Name = newSheetName;

        // 在新工作表中设置标题行
        newSheet.Range("A1").Value2 = "工作表名称";
        newSheet.Range("B1").Value2 = "索引号";
        newSheet.Range("C1").Value2 = "写入的单元格";
        newSheet.Range("D1").Value2 = "底稿科目代码";

        // 在F1单元格提供说明
        newSheet.Range("F1").Value2 =
            "请将底稿科目代码写入E1单元格，例如“124001-7”。如果未指定具体的单元格位置，插件将在每个表格的第一行新增一行，并将生成代码写在A1单元格中。";

        // 调用函数获取工作表名称并插入到新工作表中
        getSheetNames(true, newSheetName + "!A2");

        // 计数器递增
        generateIndexTemplate.counter++;
    } catch (error) {
        console.error("生成模板表时发生错误: ", error.message);
    }
}

/* 编写底稿索引 - 添加索引 */
function addIndexesToSheets() {
    try {
        // 获取当前活动的工作表
        let activeSheet = Application.ActiveSheet;

        // 获取E1单元格的值（底稿科目代码）
        let subjectCode = activeSheet.Range("E1").Value();

        // 获取A列的最后一行行号
        let lastRow = activeSheet.Cells(activeSheet.Rows.Count, "A").End(xlUp).Row;

        // 遍历从第2行到最后一行的数据
        for (let i = 2; i <= lastRow; i++) {
            // 获取当前行的B列值（索引号）
            let indexNumber = activeSheet.Cells(i, "B").Value();
            // 获取当前行的A列值（工作表名称）
            let sheetName = activeSheet.Cells(i, "A").Value();
            // 获取当前行的C列值（目标单元格地址）
            let targetCellAddress = activeSheet.Cells(i, "C").Value();

            // 如果B列和A列都有值
            if (indexNumber && sheetName) {
                // 生成新的内容，格式为 "底稿科目代码-索引号"
                let newContent = subjectCode + "-" + indexNumber;

                // 获取目标工作表
                let targetSheet = Application.Worksheets.Item(sheetName);

                // 如果目标工作表存在
                if (targetSheet) {
                    // 定义目标单元格
                    let targetCell;
                    if (targetCellAddress) {
                        // 如果C列有值，则使用C列的值作为目标单元格地址
                        targetCell = targetSheet.Range(targetCellAddress);
                    } else {
                        // 如果C列没有值，则默认使用A1单元格
                        targetCell = targetSheet.Range("A1");
                    }

                    // 在目标单元格所在行插入新行
                    targetCell.EntireRow.Insert(xlDown);

                    // 将新内容写入目标单元格
                    if (targetCellAddress) {
                        targetSheet.Range(targetCellAddress).Value2 = newContent;
                    } else {
                        targetSheet.Range("A1").Value2 = newContent;
                    }

                    // 设置目标单元格的字体格式
                    targetCell.Font.Bold = true; // 加粗
                    targetCell.Font.Color = 255; // 红色 (RGB: 255, 0, 0)
                    targetCell.Font.Name = "微软雅黑"; // 字体名称
                    targetCell.Font.Size = 9; // 字体大小
                    targetCell.HorizontalAlignment = xlCenter; // 水平居中
                } else {
                    console.warn(`工作表 ${sheetName} 不存在`);
                }
            }
        }
    } catch (error) {
        console.error("添加索引时发生错误: ", error.message);
    }
}

/* 编写底稿索引 - 修改工作表名 */
function renameSheetsBasedOnIndex() {
    try {
        // 获取当前活动的工作表
        let activeSheet = ActiveSheet;

        // 获取A列最后一个有数据的行号
        let lastRow = activeSheet.Cells(activeSheet.Rows.Count, "A").End(xlUp).Row;

        // 从第二行开始遍历到最后一行
        for (let i = 2; i <= lastRow; i++) {
            // 获取当前行的B列（索引列）的值
            let index = activeSheet.Cells(i, "B").Value2;

            // 获取当前行的A列（工作表名列）的值
            let sheetName = activeSheet.Cells(i, "A").Value2;

            // 检查索引和工作表名是否都存在
            if (index && sheetName) {
                // 生成新的工作表名，格式为 "索引-工作表名"
                let newSheetName = index + "-" + sheetName;

                try {
                    // 尝试将指定名称的工作表重命名为新的名称
                    Worksheets(sheetName).Name = newSheetName;
                } catch (error) {
                    console.warn(`无法重命名工作表 ${sheetName}: ${error.message}`);
                }
            }
        }
    } catch (error) {
        console.error("修改工作表名时发生错误: ", error.message);
    }
}
