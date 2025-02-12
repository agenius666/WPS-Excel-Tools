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

/* 将选中区域中的相对引用转换为绝对引用 */
function convertRelativeReferencesToAbsolute() {
    try {
        // 获取当前选定的区域
        let selection = Application.Selection;

        // 检查选定区域是否为空
        if (!selection) {
            console.log("请选择一个区域！");
            return;
        }

        // 遍历选定区域中的每个单元格
        for (let i = 1; i <= selection.Cells.Count; i++) {
            let cell = selection.Cells.Item(i);

            // 检查单元格是否包含公式
            if (cell.HasFormula) {
                // 获取公式
                let formula = cell.Formula;

                // 使用正则表达式将相对引用（A1）转换为绝对引用（$A$1）
                let newFormula = formula.replace(/([A-Za-z]+)(\d+)/g, "$$$1$$$2");

                // 更新公式
                cell.Formula = newFormula;
            }
        }

        console.log("相对引用已成功转换为绝对引用！");
    } catch (error) {
        // 捕获错误并记录日志
        console.error("发生错误: ", error.message);
    }
}

/* 将选中区域中的绝对引用转换为相对引用 */
function convertAbsoluteReferencesToRelative() {
    try {
        // 获取当前选定的区域
        let selection = Application.Selection;

        // 检查选定区域是否为空
        if (!selection) {
            console.log("请选择一个区域！");
            return;
        }

        // 遍历选定区域中的每个单元格
        for (let i = 1; i <= selection.Cells.Count; i++) {
            let cell = selection.Cells.Item(i);

            // 检查单元格是否包含公式
            if (cell.HasFormula) {
                // 获取公式
                let formula = cell.Formula;

                // 使用正则表达式将绝对引用（$A$1）转换为相对引用（A1）
                let newFormula = formula.replace(/\$([A-Za-z]+)\$(\d+)/g, "$1$2");

                // 更新公式
                cell.Formula = newFormula;
            }
        }

        console.log("绝对引用已成功转换为相对引用！");
    } catch (error) {
        // 捕获错误并记录日志
        console.error("发生错误: ", error.message);
    }
}
