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

/* 将选中单元格的数值转换为万元显示 */
function convertToTenThousandDisplay() {
    try {
        // 获取当前选中的单元格区域
        let selection = Application.Selection;

        // 遍历选中的每个单元格
        for (let i = 1; i <= selection.Count; i++) {
            let cell = selection.Item(i);

            // 获取当前单元格的公式和值
            let currentFormula = cell.Formula;
            let currentValue = cell.Value;

            // 检查当前单元格是否有公式
            if (currentFormula) {
                // 如果单元格有公式，删除公式中的 "=" 符号
                let formulaWithoutEqual = currentFormula.startsWith("=") ? currentFormula.slice(1) : currentFormula;

                // 构建新的公式，嵌套 ROUND 函数，并将公式除以 10000
                let newFormula = "=ROUND((" + formulaWithoutEqual + ")/10000, 2)";
                cell.Formula = newFormula;
            } else {
                // 如果单元格没有公式，直接对值进行 ROUND 操作，并将结果以公式形式写入
                if (currentValue !== null && !isNaN(currentValue)) {
                    let roundedFormula = "=ROUND(" + currentValue + "/10000, 2)";
                    cell.Formula = roundedFormula;
                }
            }
        }

        console.log("已成功将选中单元格的数值转换为万元显示！");
    } catch (error) {
        // 捕获错误并记录日志
        console.error("发生错误: ", error.message);
    }
}

/* 还原选中单元格的具体数值显示 */
function restoreOriginalValueDisplay() {
    try {
        // 获取当前选中的单元格区域
        let selection = Application.Selection;

        // 遍历选中的每个单元格
        for (let i = 1; i <= selection.Count; i++) {
            // 获取当前遍历到的单元格
            let cell = selection.Item(i);

            // 获取当前单元格的公式
            let currentFormula = cell.Formula;

            // 如果当前单元格有公式
            if (currentFormula) {
                // 检查公式是否以 "=ROUND(" 开头
                if (currentFormula.startsWith("=ROUND(")) {
                    // 使用正则表达式匹配公式中的原始公式部分
                    let originalFormulaMatch = currentFormula.match(/=ROUND\(([^,]+,.*)\)/);
                    if (originalFormulaMatch) {
                        // 提取原始公式部分
                        let originalFormula = originalFormulaMatch[1];

                        // 如果原始公式以 "(ROUND" 开头
                        if (originalFormula.startsWith("(ROUND")) {
                            // 去除原始公式中的多余括号和 "/10000,2" 部分
                            originalFormula = originalFormula.replace(/^\(/, "").replace(/\)\/10000,2$/, "");
                            // 将修改后的公式写回单元格
                            cell.Formula = "=" + originalFormula;
                        } else {
                            // 如果原始公式不以 "(ROUND" 开头，仅去除 "/10000,2" 部分
                            originalFormula = originalFormula.replace(/\/10000,2$/, "");
                            // 将修改后的公式写回单元格
                            cell.Formula = "=" + originalFormula;
                        }
                    }
                }
            } else {
                // 如果当前单元格没有公式，则跳过该单元格
                continue;
            }
        }

        console.log("已成功还原选中单元格的具体数值显示！");
    } catch (error) {
        // 捕获错误并记录日志
        console.error("发生错误: ", error.message);
    }
}

/* 查询汇率数据并写入 Excel */
function fetchAndExportExchangeRates() {
    try {
        // 提示用户输入开始日期和结束日期
        let startDate = Application.InputBox("请输入开始日期（格式：YYYY-MM-DD）", "开始日期", "", 2);
        let endDate = Application.InputBox("请输入结束日期（格式：YYYY-MM-DD）", "结束日期", "", 2);

        // 检查用户输入是否有效
        if (!startDate || !endDate || startDate === "" || endDate === "") {
            MsgBox("请选择开始日期和结束日期");
            return;
        }

        // 构造请求 URL 和参数
        let url = "https://www.safe.gov.cn/AppStructured/hlw/RMBQuery.do";
        let params = `startDate=${encodeURIComponent(startDate)}&endDate=${encodeURIComponent(endDate)}&queryYN=true`;

        // 发送 POST 请求
        let http = new XMLHttpRequest();
        http.open("POST", url, false); // 同步请求
        http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
        http.send(params);

        // 检查请求状态
        if (http.status === 200) {
            let html = http.responseText;

            // 使用正则表达式提取表格数据
            let tableRegex = /<table[^>]*id="InfoTable"[^>]*>([\s\S]*?)<\/table>/i;
            let tableMatch = html.match(tableRegex);

            if (tableMatch && tableMatch[1]) {
                let tableHtml = tableMatch[1];

                // 提取行数据
                let rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
                let rowMatches = tableHtml.match(rowRegex);

                if (rowMatches) {
                    // 获取当前工作簿
                    let workbook = Application.ActiveWorkbook;

                    // 生成新工作表名称
                    let sheetName = `汇率统计_${startDate}_${endDate}`;

                    // 创建新工作表
                    let worksheet = workbook.Worksheets.Add();
                    worksheet.Name = sheetName; // 设置工作表名称

                    // 遍历每一行
                    for (let i = 0; i < rowMatches.length; i++) {
                        let rowHtml = rowMatches[i];

                        // 提取单元格数据
                        let cellRegex = /<(td|th)[^>]*>([\s\S]*?)<\/(td|th)>/gi;
                        let cellMatches = rowHtml.match(cellRegex);

                        if (cellMatches) {
                            for (let j = 0; j < cellMatches.length; j++) {
                                let cellHtml = cellMatches[j];

                                // 提取单元格内容，并清理服务器端标签和多余空白
                                let contentRegex = /<(td|th)[^>]*>([\s\S]*?)<\/(td|th)>/i;
                                let contentMatch = cellHtml.match(contentRegex);

                                if (contentMatch && contentMatch[2]) {
                                    let cellContent = contentMatch[2]
                                        .replace(/<s:[^>]*>([\s\S]*?)<\/s:[^>]*>/g, "") // 去除服务器端标签
                                        .replace(/<!--[\s\S]*?-->/g, "") // 去除注释
                                        .replace(/\s+/g, " ") // 合并多余空白
                                        .trim(); // 去除首尾空白

                                    // 将内容写入 Excel 单元格
                                    worksheet.Cells(i + 1, j + 1).Value2 = cellContent;
                                }
                            }
                        }
                    }

                    // 删除第二行（通常为无用数据）
                    worksheet.Rows(2).Delete();

                    MsgBox("数据已成功写入 Excel 表格！");
                } else {
                    MsgBox("未找到表格行数据！");
                }
            } else {
                MsgBox("未找到表格元素！");
            }
        } else {
            MsgBox(`请求失败，状态码：${http.status}`);
        }
    } catch (error) {
        console.error("发生错误: ", error.message);
        MsgBox("发生错误，请检查控制台日志！");
    }
}
