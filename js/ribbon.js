// ==========================================
// 核心指挥中心：点击按钮分配任务
// ==========================================
function OnAction(control) {
    const eleId = control.Id; 
    
    switch (eleId) {
        case "btnCleanGtj":
            DoCleanDataAndFormat(); // 去底色+清空格转数值 (极速版)
            break;
        case "btnCopyID":
            DoCopySimple(); // 复制拼接ID (静默版)
            break;
        case "btnCleanFormat":
            DoCleanStyle(); // 去底色+11号字
            break;
        case "btnToggleZero":
            DoToggleZero(); // 显示/隐藏0值
            break;
        case "btnCalcFormula": 
            DoShowCalcDialog(); // 唤起录入计算弹窗
            break;
        case "btnTraceRef":     
            DoShowTraceDialog(); // 公式追踪按钮指令
            break;
        case "btnFormatTool":   // 👈 新增：唤起格式清洗弹窗
            DoShowFormatDialog();
            break;
        case "btnCustom1":
            alert("预留按钮，随时可以写新代码接上来！");
            break;
    }
}

// ==========================================
// 功能：弹出【格式清洗】对话框 (新增)
// ==========================================
function DoShowFormatDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/format.html";
    // ⚠️ 宽度 350 不变，高度从 240 压缩到 180！
    wps.ShowDialog(htmlUrl, "高级格式清洗", 350, 180, false);
}

// ==========================================
// 功能：弹出【公式追踪】对话框
// ==========================================
function DoShowTraceDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/trace.html";
    wps.ShowDialog(htmlUrl, "公式追踪神器", 380, 480, false);
}

// ==========================================
// 功能：弹出【录入计算公式】对话框
// ==========================================
function DoShowCalcDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/calc.html";
    wps.ShowDialog(htmlUrl, "录入计算公式", 520, 200, false);
}

// ==========================================
// 功能：极速去底色 + 清空格转数值 (全自动智能判断版)
// ==========================================
function DoCleanDataAndFormat() {
    const app = wps.Application;
    const sheet = app.ActiveSheet;
    if (!sheet) return;

    let targetRange;
    
    if (app.Selection && app.Selection.Count > 1) {
        targetRange = app.Selection;
    } else {
        targetRange = sheet.UsedRange; 
    }

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

// ==========================================
// 功能：显示/隐藏0值 (一键切换)
// ==========================================
function DoToggleZero() {
    const app = wps.Application;
    try {
        app.ActiveWindow.DisplayZeros = !app.ActiveWindow.DisplayZeros;
    } catch(e) {
        alert("切换失败，请确保当前打开了表格文件！");
    }
}

// ==========================================
// 功能：去底色 + 11号字
// ==========================================
function DoCleanStyle() {
    const app = wps.Application;
    const xlNone = -4142;
    
    const rng = app.ActiveSheet.UsedRange;
    if (!rng) return;
    
    rng.Interior.ColorIndex = xlNone; 
    rng.Font.Size = 11;
}

// ==========================================
// 功能：复制拼接ID (避开隐藏行，静默无弹窗版)
// ==========================================
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
    } catch(e) {
    }
}

// 插件基础配置函数（勿删）
function OnGetVisible(control){ return true; }
function OnGetLabel(control){ return ""; }
function OnAddinLoad(ribbonUI){ }

// ==========================================
// 专门负责给按钮发送自定义图标的“快递员”
// ==========================================
function GetImage(control) {
    const eleId = control.Id;
    switch (eleId) {
        case "btnCleanGtj": return "images/clean_data.png";
        case "btnCopyID": return "images/copy_id.png";
        case "btnCleanFormat": return "images/clean_style.png";
        case "btnToggleZero": return "images/zero.png";
        case "btnCalcFormula": return "images/calc.png"; 
        case "btnTraceRef": return "images/trace.png"; 
        case "btnFormatTool": return "images/FormatTool_style.png"; // 👈 新增：给格式清洗发图标 (复用了之前的图标)
        case "btnCustom1": return "images/custom.png";
    }
    return "";
}