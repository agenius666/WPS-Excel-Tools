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

/* 
1. Excel 文件 (AI-API.xlsx)
Sheet1：用于存储 API 信息。
B1 单元格：存储 API URL。
B2 单元格：存储 API Key。
B3 单元格：存储模型名称。
2. UserForm (ufm_大模型框架)
控件：
txtInputBox：文本框，用于用户输入问题。
txtDisplayBox：文本框，用于显示 API 的响应。
btnSetApi：按钮，用于设置 API 信息。
btnSendRequest：按钮，用于发送请求到 API。
*/

// 全局变量
let apiInfo;

// 初始化 API 信息
function initializeApiInfo() {
    return {
        apiUrl: "",
        apiKey: "",
        model: ""
    };
}

// 读取 API 信息和模型信息
function readApiInfo(apiInfo) {
    const filePath = "\\AI-API.xlsx";
    try {
        // 打开工作簿
        let workbook = Application.Workbooks.Open(filePath);
        console.log("文件已成功打开");
        let worksheet = workbook.Sheets(1);
        console.log("工作表已获取");

        // 读取 API 信息和模型信息
        apiInfo.apiUrl = worksheet.Cells(1, 2).Value2.toString(); // B1 单元格
        console.log("readApiInfo-API URL 已读取:", apiInfo.apiUrl);
        apiInfo.apiKey = worksheet.Cells(2, 2).Value2.toString(); // B2 单元格
        apiInfo.model = worksheet.Cells(3, 2).Value2.toString();  // B3 单元格

        // 关闭工作簿
        workbook.Close(false);
        console.log("readApiInfo-工作簿已关闭");

        console.log("readApiInfo-API URL: " + apiInfo.apiUrl);
        console.log("readApiInfo-API Key: " + apiInfo.apiKey);
        console.log("readApiInfo-Model: " + apiInfo.model);
    } catch (e) {
        console.error("读取文件时发生错误:", e);
        MsgBox("读取 API 信息失败，请检查文件路径和内容是否正确。");
        console.log(e);
    }
}

// 显示主对话框
function showMainDialog(apiInfo) {
    // 读取 API 信息和模型信息
    readApiInfo(apiInfo);
    console.log("showMainDialog-API URL: " + apiInfo.apiUrl);
    console.log("showMainDialog-API Key: " + apiInfo.apiKey);
    console.log("showMainDialog-Model: " + apiInfo.model);
    
        // 检查 UserForm 是否已定义
    if (typeof ufm_大模型框架 === 'undefined') {
        console.error("ufm_大模型框架 未定义");
            console.log("ufm_大模型框架 未定义");
        MsgBox("主框架未定义。");
        return;
    }
    
                console.log("ufm_大模型框架 已定义");
    // 直接展示 UserForm
    ufm_大模型框架.Show();
        console.log("表单已显示");
}

// 设置 API 按钮点击事件
function ufm_大模型框架_btnSetApi_Click() {
    console.log("设置 API 按钮被点击");

    // 使用 InputBox 弹窗获取 API 信息和模型信息
    let inputApiUrl = Application.InputBox("请输入 API URL:", "设置 API URL", );
    let inputApiKey = Application.InputBox("请输入 API Key:", "设置 API Key", );
    let inputModel = Application.InputBox("请输入模型名称:", "设置模型名称", );
    console.log("InputBox-API URL: " + inputApiUrl);
    console.log("InputBox-API Key: " + inputApiKey);
    console.log("InputBox-Model: " + inputModel);

    // 如果用户取消输入，直接返回
    if (inputApiUrl === false || inputApiKey === false || inputModel === false) {
        MsgBox("用户取消了输入，API 信息和模型信息未更新。");
        return;
    }

    // 保存 API 信息和模型信息到 Excel 文件
    const filePath = "\\AI-API.xlsx";
    try {
        // 打开或创建工作簿
        let workbook = Application.Workbooks.Open(filePath, false, false);
        let worksheet = workbook.Sheets(1);

        // 写入 API 信息和模型信息
        worksheet.Cells(1, 2).Value2 = inputApiUrl.toString(); // B1 单元格
        worksheet.Cells(2, 2).Value2 = inputApiKey.toString(); // B2 单元格
        worksheet.Cells(3, 2).Value2 = inputModel.toString();  // B3 单元格

        // 保存并关闭工作簿
        workbook.Save();
        workbook.Close();

        MsgBox("API 信息和模型信息已保存到文件。");
        console.log("API 信息和模型信息已保存到文件。");
    } catch (e) {
        MsgBox("保存 API 信息和模型信息失败，请检查文件路径是否可写。");
        console.log(e);
    }

    // 重新读取 Excel 文件中的信息并更新全局变量
    updateApiInfoFromExcel(apiInfo);
}

// 重新读取 Excel 文件中的信息并更新全局变量
function updateApiInfoFromExcel(apiInfo) {
    try {
        readApiInfo(apiInfo);

        console.log("update-API URL: " + apiInfo.apiUrl);
        console.log("update-API Key: " + apiInfo.apiKey);
        console.log("update-Model: " + apiInfo.model);
    } catch (e) {
        console.log("重新读取 API 信息失败:", e);
        MsgBox("重新读取 API 信息失败，请检查文件路径和内容是否正确。");
        console.error("重新读取 API 信息失败:", e);
    }
}

// 发送 HTTP 请求的封装函数
function sendHttpRequest(url, method, headers, body, callback) {
    var http = new XMLHttpRequest();
    http.open(method, url, true);

    // 设置请求头
    for (var key in headers) {
        http.setRequestHeader(key, headers[key]);
    }

    // 处理响应
    http.onreadystatechange = function () {
        if (http.readyState === 4) {
            if (http.status === 200) {
                callback(null, http.responseText);
            } else {
                callback(new Error(`请求失败，状态码：${http.status}`), null);
            }
        }
    };

    // 发送请求
    http.send(body);
}

// 发送请求按钮点击事件
function ufm_大模型框架_btnSendRequest_Click() {
    // 检查 UserForm 是否已定义
    if (typeof ufm_大模型框架 === 'undefined') {
        console.error("ufm_大模型框架 未定义");
        MsgBox("UserForm 未定义，请检查代码。");
        return;
    }

    // 获取控件的值
    var displayBox = ufm_大模型框架.txtDisplayBox;
    var inputBox = ufm_大模型框架.txtInputBox;

    // 获取用户输入的问题
    var userText = inputBox.Value;

    // 检查 API 信息是否已设置
    if (!apiInfo.apiUrl || !apiInfo.apiKey) {
        MsgBox("请先设置 API 接口和密钥！");
        return;
    }

    // 调用 API
    var url = apiInfo.apiUrl;
    var headers = {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiInfo.apiKey}`
    };
    var body = JSON.stringify({
        model: apiInfo.model, // 使用用户设置的模型
        messages: [{ role: "user", content: userText }]
    });

    // 发送 HTTP 请求
    sendHttpRequest(url, "POST", headers, body, function (error, response) {
        if (error) {
            MsgBox("请求失败：" + error.message);
        } else {
            try {
                var result = JSON.parse(response);
                var answer = result.choices[0].message.content;

                // 更新显示框内容
                var currentText = displayBox.Value;
                displayBox.Value = `用户：${userText}\r\n\r\n大模型：${answer}\r\n\r\n${currentText}`;
            } catch (e) {
                MsgBox("解析 API 响应失败，请检查返回的数据格式。");
            }
        }
    });
}


// 主函数，用于初始化并显示主对话框
function AImain() {
    apiInfo = initializeApiInfo(); // 初始化全局变量
    showMainDialog(apiInfo); // 传递
        console.log("主函数执行完毕，表单应已显示");
}
