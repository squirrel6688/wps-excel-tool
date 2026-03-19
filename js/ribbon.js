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
    wps.ShowDialog(htmlUrl, "优化工程量表", 350, 500, false);
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

    let targetRange = (app.Selection && app.Selection.Count > 1) ? app.Selection : sheet.UsedRange;
    if (!targetRange) return;

    // 1. 去除所有底色
    targetRange.Interior.ColorIndex = 0;

    // 2. 遍历数据处理空格
    for (let i = 1; i <= targetRange.Areas.Count; i++) {
        let area = targetRange.Areas.Item(i);

        // 如果只选中了一个单元格
        if (area.Count === 1) {
            let val = area.Value2;
            if (typeof val === 'string') {
                let cleanVal = val.replace(/\s+/g, ''); // 尝试去掉空格
                // 核心逻辑：只有去掉空格后它是个“纯数字”，才转换！否则原封不动保留！
                if (!isNaN(cleanVal) && cleanVal !== '') {
                    area.Value2 = Number(cleanVal);
                }
            }
            continue;
        }

        // 如果是批量区域
        let dataArr = area.Value2;
        if (!dataArr) continue;

        for (let r = 0; r < dataArr.length; r++) {
            for (let c = 0; c < dataArr[r].length; c++) {
                let val = dataArr[r][c];
                if (typeof val === 'string') {
                    let cleanVal = val.replace(/\s+/g, '');
                    // 核心逻辑：遇到文字或符号，依然保留原状不误伤
                    if (!isNaN(cleanVal) && cleanVal !== '') {
                        dataArr[r][c] = Number(cleanVal);
                    }
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