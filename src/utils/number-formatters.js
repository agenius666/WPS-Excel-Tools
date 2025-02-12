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

/* 为选中单元格添加 ROUND 函数 */
function addRoundFunctionToSelection() {
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

                // 构建新的公式，嵌套 ROUND 函数，并保留两位小数
                let newFormula = "=ROUND(" + formulaWithoutEqual + ", 2)";
                cell.Formula = newFormula;
            } else {
                // 如果单元格没有公式，直接对值进行 ROUND 操作，并将结果以公式形式写入
                if (currentValue !== null && !isNaN(currentValue)) {
                    let roundedFormula = "=ROUND(" + currentValue + ", 2)";
                    cell.Formula = roundedFormula;
                }
            }
        }

        console.log("已成功为选中单元格添加 ROUND 函数！");
    } catch (error) {
        // 捕获错误并记录日志
        console.error("发生错误: ", error.message);
    }
}

/* 删除选中单元格中的 ROUND 函数 */
function removeRoundFunctionFromSelection() {
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
                    // 使用正则表达式匹配公式中的内容
                    let originalFormulaMatch = currentFormula.match(/=ROUND\(([^,]+,.*)\)/);
                    if (originalFormulaMatch) {
                        // 提取原始公式部分（去掉外层的 ROUND 函数）
                        let originalFormula = originalFormulaMatch[1];

                        // 如果原始公式以 "ROUND(ROUND" 开头
                        if (originalFormula.startsWith("ROUND(ROUND")) {
                            // 去掉末尾的 ",2"，即去掉第二个参数（保留小数位数）
                            originalFormula = originalFormula.replace(/,2$/, "");
                            // 将修改后的公式重新赋值给单元格
                            cell.Formula = "=" + originalFormula;
                        } else if (originalFormula.startsWith("ROUND(")) {
                            // 如果原始公式以 "ROUND(" 开头
                            // 去掉末尾的 ",2"，即去掉第二个参数（保留小数位数）
                            originalFormula = originalFormula.replace(/,2$/, "");
                            // 将修改后的公式重新赋值给单元格
                            cell.Formula = "=" + originalFormula;
                        } else {
                            // 其他情况（非嵌套的 ROUND 函数）
                            // 去掉末尾的 ",2"，即去掉第二个参数（保留小数位数）
                            originalFormula = originalFormula.replace(/,2$/, "");
                            // 将修改后的公式重新赋值给单元格
                            cell.Formula = "=" + originalFormula;
                        }
                    }
                }
            } else {
                // 如果当前单元格没有公式，则跳过该单元格
                continue;
            }
        }

        console.log("已成功删除选中单元格中的 ROUND 函数！");
    } catch (error) {
        // 捕获错误并记录日志
        console.error("发生错误: ", error.message);
    }
}
