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
