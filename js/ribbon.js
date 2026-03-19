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

// 弹出：数据透视表对话框
function DoShowPivotDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/pivot.html";
    wps.ShowDialog(htmlUrl, "数据透视与汇总", 350, 380, false);
}

// 弹出：格式清洗对话框 
function DoShowFormatDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/format.html";
    wps.ShowDialog(htmlUrl, "高级格式清洗", 350, 220, false);
}

// 弹出：录入计算公式对话框
function DoShowCalcDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/calc.html";
    wps.ShowDialog(htmlUrl, "录入计算公式", 520, 200, false);
}

// 弹出：关于与风险提示对话框
function DoShowAboutDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/about.html";
    wps.ShowDialog(htmlUrl, "关于量效助手", 400, 280, false);
}

// 弹出：使用说明对话框
function DoShowGuide() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/guide.html";
    // 界面稍微大一点，方便阅读说明文字
    wps.ShowDialog(htmlUrl, "量效助手 - 使用说明", 500, 480, false);
}

// ==========================================
// 【特殊操作功能区】
// ==========================================

// 动作：弹出左侧【公式追踪】任务窗格 (侧边栏)
function DoShowTraceDialog() {
    let basePath = location.href.substring(0, location.href.lastIndexOf("/"));
    let htmlUrl = basePath + "/ui/trace.html";

    // 检查侧边栏是否已经创建过
    let tpId = wps.PluginStorage.getItem("trace_taskpane_id");
    let tp = tpId ? wps.GetTaskPane(tpId) : null;

    if (!tp) {
        // 创建新的侧边栏
        tp = wps.CreateTaskPane(htmlUrl, "公式追踪");
        if (tp) {
            wps.PluginStorage.setItem("trace_taskpane_id", tp.ID);
            tp.DockPosition = 0; // 停靠在左侧
            tp.Width = 200;      // 允许的极限最小宽度
        }
    }
    if (tp) tp.Visible = true; // 显示出来
}

// 动作：一键复制分享链接到剪贴板
function DoShowShare() {
    // 你的 GitHub 在线安装地址
    const shareUrl = "https://squirrel6688.github.io/wps-excel-tool/publish.html";

    try {
        let textArea = document.createElement("textarea");
        textArea.value = "推荐你使用【量效助手】WPS插件，提升算量效率！\n安装地址：" + shareUrl;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand("copy"); // 执行系统复制命令
        document.body.removeChild(textArea);

        alert("🎉 复制成功！\n\n插件安装链接已复制到剪贴板，快去微信粘贴发送给好友吧！");
    } catch (e) {
        alert("复制失败，请手动分享网址：" + shareUrl);
    }
}

// 动作：联系作者 (后期预留位置)
function DoShowContact() {
    alert("👨‍💻 作者正在日夜奋战开发新功能中...\n如需定制或反馈问题，此通道将在后续版本开放，敬请期待！");
}

// ==========================================
// 【底层数据处理执行区】 (原封不动保留)
// ==========================================

// 极速去底色 + 清空格转数值
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

// 显示/隐藏0值
function DoToggleZero() {
    const app = wps.Application;
    try { app.ActiveWindow.DisplayZeros = !app.ActiveWindow.DisplayZeros; } catch (e) { }
}

// 去底色 + 11号字
function DoCleanStyle() {
    const app = wps.Application;
    const xlNone = -4142;
    const rng = app.ActiveSheet.UsedRange;
    if (!rng) return;
    rng.Interior.ColorIndex = xlNone;
    rng.Font.Size = 11;
}

// 复制拼接ID
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

// ==========================================
// 图标分发器：专门负责给界面的按钮派发图片
// ==========================================
function GetImage(control) {
    const eleId = control.Id;
    switch (eleId) {
        // 主面板功能图标
        case "btnCleanGtj": return "images/clean_data.png";
        case "btnCopyID": return "images/copy_id.png";
        case "btnCleanFormat": return "images/clean_style.png";
        case "btnToggleZero": return "images/zero.png";
        case "btnCalcFormula": return "images/calc.png";
        case "btnTraceRef": return "images/trace.png";
        case "btnFormatTool": return "images/clean_style.png";
        case "btnPivotSummary": return "images/pivot.png"; // 这里改成了你自带的图标

        // 用户下拉菜单图标 (复用现有图标进行示意)
        case "menuUser": return "images/custom.png";
        case "btnAbout": return "images/clean_data.png";
        case "btnShare": return "images/copy_id.png";
        case "btnGuide": return "images/calc.png";
        case "btnContact": return "images/custom.png";
    }
    return "";
}