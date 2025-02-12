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

/* 设置所有工作表为A4横向打印，并调整为一页 */
function setAllSheetsA4LandscapeOnePage() {
    var workbook = Application.ActiveWorkbook;
    for (var i = 1; i <= workbook.Sheets.Count; i++) {
        var sheet = workbook.Sheets(i);
        Application.Worksheets.Item(sheet).PageSetup.PaperSize = xlPaperA4;
        Application.Worksheets.Item(sheet).PageSetup.Orientation = xlLandscape;
        let pagesetup = Application.Worksheets.Item(sheet).PageSetup;
        pagesetup.Zoom = false;
        pagesetup.FitToPagesTall = 1;
        pagesetup.FitToPagesWide = 1;
    }
}

/* 设置所有工作表为A4横向打印，并将所有列调整为一页 */
function setAllSheetsA4LandscapeOnePageWide() {
    var workbook = Application.ActiveWorkbook;
    for (var i = 1; i <= workbook.Sheets.Count; i++) {
        var sheet = workbook.Sheets(i);
        Application.Worksheets.Item(sheet).PageSetup.PaperSize = xlPaperA4;
        Application.Worksheets.Item(sheet).PageSetup.Orientation = xlLandscape;
        let pagesetup = Application.Worksheets.Item(sheet).PageSetup;
        pagesetup.Zoom = false;
        pagesetup.FitToPagesTall = 0;
        pagesetup.FitToPagesWide = 1;
    }
}

/* 设置所有工作表为A4横向打印，并将所有行调整为一页 */
function setAllSheetsA4LandscapeOnePageTall() {
    var workbook = Application.ActiveWorkbook;
    for (var i = 1; i <= workbook.Sheets.Count; i++) {
        var sheet = workbook.Sheets(i);
        Application.Worksheets.Item(sheet).PageSetup.PaperSize = xlPaperA4;
        Application.Worksheets.Item(sheet).PageSetup.Orientation = xlLandscape;
        let pagesetup = Application.Worksheets.Item(sheet).PageSetup;
        pagesetup.Zoom = false;
        pagesetup.FitToPagesTall = 1;
        pagesetup.FitToPagesWide = 0;
    }
}

/* 设置所有工作表为A4纵向打印，并调整为一页 */
function setAllSheetsA4PortraitOnePage() {
    var workbook = Application.ActiveWorkbook;
    for (var i = 1; i <= workbook.Sheets.Count; i++) {
        var sheet = workbook.Sheets(i);
        Application.Worksheets.Item(sheet).PageSetup.PaperSize = xlPaperA4;
        Application.Worksheets.Item(sheet).PageSetup.Orientation = xlPortrait;
        let pagesetup = Application.Worksheets.Item(sheet).PageSetup;
        pagesetup.Zoom = false;
        pagesetup.FitToPagesTall = 1;
        pagesetup.FitToPagesWide = 1;
    }
}

/* 设置所有工作表为A4纵向打印，并将所有列调整为一页 */
function setAllSheetsA4PortraitOnePageWide() {
    var workbook = Application.ActiveWorkbook;
    for (var i = 1; i <= workbook.Sheets.Count; i++) {
        var sheet = workbook.Sheets(i);
        Application.Worksheets.Item(sheet).PageSetup.PaperSize = xlPaperA4;
        Application.Worksheets.Item(sheet).PageSetup.Orientation = xlPortrait;
        let pagesetup = Application.Worksheets.Item(sheet).PageSetup;
        pagesetup.Zoom = false;
        pagesetup.FitToPagesTall = 0;
        pagesetup.FitToPagesWide = 1;
    }
}

/* 设置所有工作表为A4纵向打印，并将所有行调整为一页 */
function setAllSheetsA4PortraitOnePageTall() {
    var workbook = Application.ActiveWorkbook;
    for (var i = 1; i <= workbook.Sheets.Count; i++) {
        var sheet = workbook.Sheets(i);
        Application.Worksheets.Item(sheet).PageSetup.PaperSize = xlPaperA4;
        Application.Worksheets.Item(sheet).PageSetup.Orientation = xlPortrait;
        let pagesetup = Application.Worksheets.Item(sheet).PageSetup;
        pagesetup.Zoom = false;
        pagesetup.FitToPagesTall = 1;
        pagesetup.FitToPagesWide = 0;
    }
}
