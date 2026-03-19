// ==========================================
// 核心指挥中心：点击按钮分配任务
// ==========================================
function OnAction(control) {
    const eleId = control.Id;

    switch (eleId) {
        case "btnCleanGtj": DoCleanDataAndFormat(); break;
        case "btnCopyID": DoCopySimple(); break;
        case "btnCleanFormat": DoCleanStyle(); break;
        case "btnToggleZero": DoToggleZero(); break;
        case "btnCalcFormula": DoShowCalcDialog(); break;
        case "btnTraceRef": DoShowTraceDialog(); break;
        case "btnFormatTool": DoShowFormatDialog(); break;
        case "btnPivotSummary": DoShowPivotDialog(); break;
        case "btnCustom1": alert("预留按钮，随时可以写新代码接上来！"); break;
    }
}

// ==========================================
// 功能：弹出【数据透析表】对话框 (已拆除硬锁)
// ==========================================
function DoShowPivotDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/pivot.html";
    wps.ShowDialog(htmlUrl, "数据透析与汇总", 350, 380, false);
}

// ==========================================
// 功能：弹出【格式清洗】对话框 (已拆除硬锁)
// ==========================================
function DoShowFormatDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/format.html";
    wps.ShowDialog(htmlUrl, "高级格式清洗", 350, 220, false);
}

// ==========================================
// 功能：弹出【录入计算公式】对话框 (已拆除硬锁)
// ==========================================
function DoShowCalcDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/calc.html";
    wps.ShowDialog(htmlUrl, "录入计算公式", 520, 200, false);
}

// ==========================================
// 🚨 功能：弹出【公式追踪】左侧任务窗格 
// ==========================================
function DoShowTraceDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/trace.html";

    let tpId = wps.PluginStorage.getItem("trace_taskpane_id");
    let tp = tpId ? wps.GetTaskPane(tpId) : null;

    if (!tp) {
        tp = wps.CreateTaskPane(htmlUrl, "公式追踪");
        if (tp) {
            wps.PluginStorage.setItem("trace_taskpane_id", tp.ID);
            tp.DockPosition = 0;
            // 🚨 解除强硬的宽度限制，设置为 200，让你可以往左边拉到极致！
            tp.Width = 200;
        }
    }
    if (tp) tp.Visible = true;
}

// ==========================================
// 其他功能执行函数 (保持不变)
// ==========================================
function DoCleanDataAndFormat() {
    const app = wps.Application;
    const sheet = app.ActiveSheet;
    if (!sheet) return;
    let targetRange = (app.Selection && app.Selection.Count > 1) ? app.Selection : sheet.UsedRange;
    if (!targetRange) return;
    targetRange.Interior.ColorIndex = 0;
    for (let i = 1; i <= targetRange.Areas.Count; i++) {
        let area = targetRange.Areas.Item(i);
        if (area.Count === 1) {
            let val = area.Value2;
            if (typeof val === 'string') {
                let cleanVal = val.replace(/\s+/g, '');
                area.Value2 = (!isNaN(cleanVal) && cleanVal !== '') ? Number(cleanVal) : cleanVal;
            }
            continue;
        }
        let dataArr = area.Value2;
        if (!dataArr) continue;
        for (let r = 0; r < dataArr.length; r++) {
            for (let c = 0; c < dataArr[r].length; c++) {
                let val = dataArr[r][c];
                if (typeof val === 'string') {
                    let cleanVal = val.replace(/\s+/g, '');
                    dataArr[r][c] = (!isNaN(cleanVal) && cleanVal !== '') ? Number(cleanVal) : cleanVal;
                }
            }
        }
        area.Value2 = dataArr;
    }
}

function DoToggleZero() {
    const app = wps.Application;
    try { app.ActiveWindow.DisplayZeros = !app.ActiveWindow.DisplayZeros; } catch (e) { }
}

function DoCleanStyle() {
    const app = wps.Application;
    const xlNone = -4142;
    const rng = app.ActiveSheet.UsedRange;
    if (!rng) return;
    rng.Interior.ColorIndex = xlNone;
    rng.Font.Size = 11;
}

function DoCopySimple() {
    const app = wps.Application;
    const sel = app.Selection;
    if (!sel || sel.Count === 0) return;
    let arr = [];
    let sheetName = app.ActiveSheet.Name;
    let safeSheetName = /^[a-zA-Z0-9_\u4e00-\u9fa5]+$/.test(sheetName) ? sheetName : "'" + sheetName + "'";
    for (let i = 1; i <= sel.Areas.Count; i++) {
        let area = sel.Areas.Item(i);
        for (let r = 1; r <= area.Rows.Count; r++) {
            for (let c = 1; c <= area.Columns.Count; c++) {
                let cell = area.Cells.Item(r, c);
                if (cell.EntireRow.Hidden === false && cell.EntireColumn.Hidden === false) {
                    arr.push(safeSheetName + "!" + cell.Address());
                }
            }
        }
    }
    if (arr.length === 0) return;
    const result = "=" + arr.join("+");
    try {
        let textArea = document.createElement("textarea");
        textArea.value = result;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand("copy");
        document.body.removeChild(textArea);
    } catch (e) { }
}

function OnAddinLoad(ribbonUI) { }
function OnGetVisible(control) { return true; }
function OnGetLabel(control) { return ""; }

function GetImage(control) {
    const eleId = control.Id;
    switch (eleId) {
        case "btnCleanGtj": return "images/clean_data.png";
        case "btnCopyID": return "images/copy_id.png";
        case "btnCleanFormat": return "images/clean_style.png";
        case "btnToggleZero": return "images/zero.png";
        case "btnCalcFormula": return "images/calc.png";
        case "btnTraceRef": return "images/trace.png";
        case "btnFormatTool": return "images/clean_style.png";
        case "btnPivotSummary": return "images/calc.png";
        case "btnCustom1": return "images/custom.png";
    }
    return "";
}