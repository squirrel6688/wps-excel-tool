// ==========================================
// 核心指挥中心：监听所有的按钮点击事件，并分配任务
// ==========================================
function OnAction(control) {
    const eleId = control.Id;

    switch (eleId) {
        // --- 格式处理组 ---
        case "btnCleanGtj": DoCleanDataAndFormat(); break; // GTJ报表清洗
        case "btnCleanFormat": DoCleanStyle(); break;      // 去底色+11号字
        case "btnFormatTool": DoShowFormatDialog(); break; // 高级格式清洗弹窗
        case "btnToggleZero": DoToggleZero(); break;       // 显示/隐藏0值

        // --- 数据处理组 ---
        case "btnCopyID": DoCopySimple(); break;           // 复制拼接ID
        case "btnCalcFormula": DoShowCalcDialog(); break;  // 录入计算公式弹窗
        case "btnTraceRef": DoShowTraceDialog(); break;    // 公式追踪侧边栏

        // --- 个性工具栏 ---
        case "btnPivotSummary": DoShowPivotDialog(); break;// 数据透视表弹窗

        // --- 用户组 (二级菜单) ---
        case "btnAbout": DoShowAboutDialog(); break;       // 关于与风险提示
        case "btnShare": DoShowShare(); break;             // 分享好友 (复制链接)
        case "btnGuide": DoShowGuide(); break;             // 使用说明弹窗
        case "btnContact": DoShowContact(); break;         // 联系作者 (预留)
    }
}

// ==========================================
// 【弹窗功能区】
// ==========================================

function DoShowPivotDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/pivot.html";
    wps.ShowDialog(htmlUrl, "数据透视与汇总", 350, 380, false);
}

function DoShowFormatDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/format.html";
    wps.ShowDialog(htmlUrl, "优化工程量表", 320, 580, false);
}

function DoShowCalcDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/calc.html";
    wps.ShowDialog(htmlUrl, "录入计算公式", 520, 200, false);
}

function DoShowAboutDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/about.html";
    wps.ShowDialog(htmlUrl, "关于量效助手", 400, 280, false);
}

function DoShowGuide() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/guide.html";
    wps.ShowDialog(htmlUrl, "量效助手 - 使用说明", 500, 480, false);
}

// ==========================================
// 【特殊操作功能区】
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
            tp.Width = 200;
        }
    }
    if (tp) tp.Visible = true;
}

function DoShowShare() {
    const shareUrl = "https://squirrel6688.github.io/wps-excel-tool/publish.html";
    try {
        let textArea = document.createElement("textarea");
        textArea.value = "推荐你使用【量效助手】WPS插件，提升算量效率！\n安装地址：" + shareUrl;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand("copy");
        document.body.removeChild(textArea);
        alert("🎉 复制成功！\n\n插件安装链接已复制到剪贴板，快去微信粘贴发送给好友吧！");
    } catch (e) {
        alert("复制失败，请手动分享网址：" + shareUrl);
    }
}

function DoShowContact() {
    alert("👨‍💻 作者正在日夜奋战开发新功能中...\n如需定制或反馈问题，此通道将在后续版本开放，敬请期待！");
}

// ==========================================
// 【底层数据处理执行区】 
// ==========================================

// 🚨 升级版：极速去底色 + 智能清空格转数值 (防止误伤文字和符号)
function DoCleanDataAndFormat() {
    const app = wps.Application;
    const sheet = app.ActiveSheet;
    if (!sheet) return;

    // 获取当前工作表中所有“有内容的区域”
    const area = sheet.UsedRange;
    if (!area) return;

    // 🚨 恢复你的神级功能：一键去除所有有内容区域的底色
    area.Interior.ColorIndex = 0;

    // 获取区域内的所有数据
    let dataArr = area.Value2;
    if (dataArr == null) return;

    // 兼容处理：如果整个表格只有一个单元格有内容，Value2 返回的是基础值而不是数组
    let isSingleCell = !Array.isArray(dataArr);
    if (isSingleCell) {
        dataArr = [[dataArr]];
    }

    // 遍历所有数据清洗空字符
    for (let r = 0; r < dataArr.length; r++) {
        for (let c = 0; c < dataArr[r].length; c++) {
            let val = dataArr[r][c];
            
            // 只有当单元格内容是文本类型时，才进行清洗判断
            if (typeof val === 'string') {
                
                // 核心逻辑 1：判断是否“单纯只有空字符”（包括空格、换行、零宽字符等）
                if (/^[\s\u00A0\u200B]*$/.test(val)) {
                    dataArr[r][c] = null; // 彻底清空，消除绿三角
                } else {
                    // 核心逻辑 2：如果包含文字，不误伤文字里的空格，但如果是带空格的纯数字则转换
                    let cleanVal = val.replace(/[\s\u00A0\u200B]+/g, '');
                    if (!isNaN(cleanVal) && cleanVal !== '') {
                        dataArr[r][c] = Number(cleanVal);
                    }
                }
            }
        }
    }

    // 将清洗后的干净数据，一次性写回工作表
    if (isSingleCell) {
        area.Value2 = dataArr[0][0];
    } else {
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

// ==========================================
// 【系统必备基础配置函数】
// ==========================================
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
        case "btnPivotSummary": return "images/pivot.png";

        case "menuUser": return "images/custom.png";
        case "btnAbout": return "images/clean_data.png";
        case "btnShare": return "images/copy_id.png";
        case "btnGuide": return "images/calc.png";
        case "btnContact": return "images/custom.png";
    }
    return "";
}