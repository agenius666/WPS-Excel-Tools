# WPS Excel Tools | WPS Excel 工具集

> WPS Excel处理工具库。

## 功能速览
`src/core/cell-operations.js`包含：`复制当前单元格并向右移动一格`、`复制当前单元格并向左移动一格`、`复制当前单元格并向上移动一格`、`复制当前单元格并向下移动一格`、`批量设置会计格式（不带￥符号）`、`插入表头模板`。

`src/core/row-column-operations.js`包含：`删除当前工作表中的空白行`、`删除当前工作表中和为0的行`、`设置所有工作表的行高为15，列宽为18`、`设置当前工作表的行高为15，列宽为18`、`批量设置所有工作表的行高和列宽`、`批量设置当前工作表的行高和列宽`、`自动调整当前工作表的所有列宽`。

`src/core/sheet-operations.js`包含：`获取所有工作表的名称并插入到指定位置`、`创建新工作表并插入工作表名称`、`删除工作簿中的所有水印（Shapes）`、`调整合并单元格的高度`、`设置所有工作表的打印区域为已使用区域`、`移除工作簿的密码保护`、`隐藏空白工作表`、`隐藏无值工作表`、`规范化所有工作表的视图设置`、`根据输入的工作表名选择工作表`、`根据输入的工作表名排除选择工作表`、`取消所有工作表的保护`。

`src/features/ai-model-integration.js`包含：`AI大模型API调用接口设置及UI界面`。

`src/features/cross-sheet-compare.js`包含：`跨表对比两个工作簿中的工作表`。

`src/features/draft-index.js`包含：`编写底稿索引 - 生成模板表`、`编写底稿索引 - 添加索引`、`编写底稿索引 - 修改工作表名`。

`src/features/reference-converters.js`包含：`将选中区域中的相对引用转换为绝对引用`、`将选中区域中的绝对引用转换为相对引用`。

`src/utils/currency-formatters.js`包含：`将选中单元格的数值转换为万元显示`、`还原选中单元格的具体数值显示`、`查询汇率数据并写入 Excel`。

`src/utils/data-converters.js`包含：`将当前工作表的所有单元格转换为数值`、`将所有工作表的所有单元格转换为数值`。

`src/utils/format-formatters.js`包含：`将选中区域的日期格式设置为 yyyy-m-d`、`将选中区域的日期格式设置为 yyyy-m`、`设置选中区域的单元格为缩小字体填充`、`批量设置字体为微软雅黑，字号为9`。

`src/utils/number-formatters.js`包含：`为选中单元格添加 ROUND 函数`、`删除选中单元格中的 ROUND 函数`。

`src/utils/print-settings.js`包含：`设置所有工作表为A4横向打印，并调整为一页`、`设置所有工作表为A4横向打印，并将所有列调整为一页`、`设置所有工作表为A4横向打印，并将所有行调整为一页`、`设置所有工作表为A4纵向打印，并调整为一页`、`设置所有工作表为A4纵向打印，并将所有列调整为一页`、`设置所有工作表为A4纵向打印，并将所有行调整为一页`。

## 快速开始
只需复制代码，打开WPS宏编辑器并运行即可。

# WPS Excel Tools | WPS Excel Toolset

> A library of tools for handling WPS Excel.

## Feature Overview
`src/core/cell-operations.js` includes: `Copy the current cell and move one cell to the right`, `Copy the current cell and move one cell to the left`, `Copy the current cell and move one cell up`, `Copy the current cell and move one cell down`, `Batch set accounting format (without ￥ symbol)`, `Insert table header template`.

`src/core/row-column-operations.js` includes: `Delete blank rows in the current worksheet`, `Delete rows with a sum of 0 in the current worksheet`, `Set row height to 15 and column width to 18 for all worksheets`, `Set row height to 15 and column width to 18 for the current worksheet`, `Batch set row height and column width for all worksheets`, `Batch set row height and column width for the current worksheet`, `Auto-adjust all column widths in the current worksheet`.

`src/core/sheet-operations.js` includes: `Get names of all worksheets and insert them into a specified location`, `Create a new worksheet and insert the worksheet name`, `Remove all watermarks (Shapes) from the workbook`, `Adjust the height of merged cells`, `Set the print area of all worksheets to the used range`, `Remove password protection from the workbook`, `Hide blank worksheets`, `Hide worksheets without values`, `Standardize view settings for all worksheets`, `Select worksheets based on input worksheet names`, `Exclude worksheets based on input worksheet names`, `Unprotect all worksheets`.

`src/features/ai-model-integration.js` includes: `AI large model API call interface settings and UI`.

`src/features/cross-sheet-compare.js` includes: `Cross-sheet comparison between two workbooks`.

`src/features/draft-index.js` includes: `Draft Index - Generate Template Sheet`, `Draft Index - Add Index`, `Draft Index - Rename Worksheet`.

`src/features/reference-converters.js` includes: `Convert relative references in the selected range to absolute references`, `Convert absolute references in the selected range to relative references`.

`src/utils/currency-formatters.js` includes: `Convert numerical values in selected cells to ten thousand yuan display`, `Restore specific numerical value display in selected cells`, `Query exchange rate data and write it into Excel`.

`src/utils/data-converters.js` includes: `Convert all cells in the current worksheet to numeric values`, `Convert all cells in all worksheets to numeric values`.

`src/utils/format-formatters.js` includes: `Set date format in the selected range to yyyy-m-d`, `Set date format in the selected range to yyyy-m`, `Set cells in the selected range to shrink font fill`, `Batch set font to Microsoft YaHei, font size to 9`.

`src/utils/number-formatters.js` includes: `Add ROUND function to selected cells`, `Remove ROUND function from selected cells`.

`src/utils/print-settings.js` includes: `Set all worksheets to A4 landscape printing and fit to one page`, `Set all worksheets to A4 landscape printing and fit all columns to one page`, `Set all worksheets to A4 landscape printing and fit all rows to one page`, `Set all worksheets to A4 portrait printing and fit to one page`, `Set all worksheets to A4 portrait printing and fit all columns to one page`, `Set all worksheets to A4 portrait printing and fit all rows to one page`.

## Quick Start
Simply copy the code, open the WPS macro editor, and run it.
