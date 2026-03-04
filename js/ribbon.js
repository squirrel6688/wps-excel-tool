// ==========================================
// 核心指挥中心：点击按钮分配任务
// ==========================================
function OnAction(control) {
    const eleId = control.Id; 
    
    switch (eleId) {
        case "btnCleanGtj":
            DoCleanData(); // 去底色+清空格转数值
            break;
        case "btnCopyID":
            DoCopySimple(); // 复制拼接ID
            break;
        case "btnCleanFormat":
            DoCleanStyle(); // 去底色+11号字
            break;
        case "btnToggleZero":
            DoToggleZero(); // 显示/隐藏0值
            break;
        case "btnCustom1":
            alert("预留按钮，随时可以写新代码接上来！");
            break;
    }
}

// ==========================================
// 功能：显示/隐藏0值 (一键切换)
// ==========================================
function DoToggleZero() {
    const app = wps.Application;
    try {
        // 读取当前窗口的 0 值显示状态，并直接反转它
        app.ActiveWindow.DisplayZeros = !app.ActiveWindow.DisplayZeros;
    } catch(e) {
        alert("切换失败，请确保当前打开了表格文件！");
    }
}

// ==========================================
// 功能：去底色 + 清首尾空格并转数值
// ==========================================
function DoCleanData() {
    const app = wps.Application;
    const xlNone = -4142; 
    const xlCellTypeConstants = 2;
    
    const rng = app.ActiveSheet.UsedRange;
    if (!rng) return;
    
    rng.Interior.ColorIndex = xlNone; 
    rng.NumberFormatLocal = "G/通用格式"; 
    
    try {
        const constRng = rng.SpecialCells(xlCellTypeConstants);
        if (!constRng) return;

        for (let i = 1; i <= constRng.Areas.Count; i++) {
            let cells = constRng.Areas.Item(i).Cells;
            for (let j = 1; j <= cells.Count; j++) {
                let c = cells.Item(j);
                let v = c.Value2;
                c.Value2 = (typeof v === "string") ? v.replace(/^[\s\u3000]+|[\s\u3000]+$/g, "") : v;
            }
        }
    } catch(e) {}
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
        // 这里的 alert 已经全部被移除了，点击后会默默把内容放进剪贴板
    } catch(e) {
        // 失败也不弹窗打扰
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
    // 根据按钮的身份证号，分别派送对应的图片
    switch (eleId) {
        case "btnCleanGtj": return "images/clean_data.png";
        case "btnCopyID": return "images/copy_id.png";
        case "btnCleanFormat": return "images/clean_style.png";
        case "btnToggleZero": return "images/zero.png";
        case "btnCustom1": return "images/custom.png";
    }
    return "";
}

