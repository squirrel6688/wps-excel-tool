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
            // ⚠️ 极其重要提醒：上面这个 "btnCalcFormula" 必须和你 ribbon.xml 里面
            // “录入计算公式”那个按钮的 id 完全一模一样！如果不对应，点按钮会没反应！
            DoShowCalcDialog(); // 唤起录入计算弹窗
            break;
        case "btnCustom1":
            alert("预留按钮，随时可以写新代码接上来！");
            break;
    }
}

// ==========================================
// 功能：弹出【录入计算公式】对话框
// ==========================================
function DoShowCalcDialog() {
    // 动态获取当前插件的路径 (兼容本地测试和线上环境)
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/calc.html";
    
    // 召唤弹窗魔法：wps.ShowDialog(网页地址, 窗口标题, 宽度, 高度, 是否强制拦截底层操作)
    // 宽度520，高度120，false代表弹窗打开时，依然可以点背后的Excel格子
    wps.ShowDialog(htmlUrl, "录入计算公式", 520, 200, false);
}

// ==========================================
// 功能：极速去底色 + 清空格转数值 (大数据秒级处理版)
// ==========================================
function DoCleanDataAndFormat() {
    const app = wps.Application;
    const sel = app.Selection;
    if (!sel || sel.Count === 0) return;

    // 1. 【极速去底色】
    sel.Interior.ColorIndex = 0; 

    // 2. 遍历选区内的每一个连续区域
    for (let i = 1; i <= sel.Areas.Count; i++) {
        let area = sel.Areas.Item(i);

        // 如果只选了一个格子
        if (area.Count === 1) {
            let val = area.Value2;
            if (typeof val === 'string') {
                let cleanVal = val.replace(/\s+/g, ''); 
                area.Value2 = (!isNaN(cleanVal) && cleanVal !== '') ? Number(cleanVal) : cleanVal;
            }
            continue;
        }

        // 3. 【核心提速】：吸入内存二维数组
        let dataArr = area.Value2; 

        // 4. 内存中极速清洗
        for (let r = 0; r < dataArr.length; r++) {
            for (let c = 0; c < dataArr[r].length; c++) {
                let val = dataArr[r][c];
                if (typeof val === 'string') {
                    let cleanVal = val.replace(/\s+/g, ''); 
                    dataArr[r][c] = (!isNaN(cleanVal) && cleanVal !== '') ? Number(cleanVal) : cleanVal;
                }
            }
        }

        // 5. 【瞬间复原】：拍回表格
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
        
        // 👇 这里就是给你新增的“录入计算公式”按钮派发图标的代码！
        // ⚠️ 确保 "btnCalcFormula" 和你在 ribbon.xml 里的 id 一模一样
        // ⚠️ 确保 "images/calc.png" 和你刚才放进文件夹里的图片名字一模一样
        case "btnCalcFormula": return "images/calc.png"; 
        
        case "btnCustom1": return "images/custom.png";
    }
    return "";
}