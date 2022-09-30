function OnAddinLoad(ribbonUI) {
    if (typeof (wps.ribbonUI) != "object") {
        wps.ribbonUI = ribbonUI
    }
    if (typeof (wps.Enum) != "object") { // 如果没有内置枚举值
        wps.Enum = WPS_Enum
    }
    wps.PluginStorage.setItem("EnableFlag", !1) //往PluginStorage中设置一个标记，用于控制两个按钮的置灰
    wps.PluginStorage.setItem("ApiEventFlag", !1) //往PluginStorage中设置一个标记，用于控制ApiEvent的按钮label
    return true

}
function Reg_Str(...regs) { let str = ''; regs.forEach(e => { str += e.toString().slice(1, -1) }); return str }
function Reg_exp(range, reg) {
    if (range === null) return 1;
    let str = '', islist = Array.isArray(range);
    if (islist) {
        if (reg.test((range.forEach(e => { str += e.Text }), str))) { range.forEach(e => { e.Font.Color = 0 }); return true }
        else { range.forEach(e => { e.Font.Color = 255 }); return false }
    }
    else {
        if (reg.test(range.Text)) { range.Font.Color = 0; return true }
        else { range.Font.Color = 255; return false }
    }
}

var WebNotifycount = 0, err, sel,
    cn_au = /(\[?[\u4e00-\u9fa5]{2,4}(,\s[\u4e00-\u9fa5]{2,4})*(,\s\&\s[\u4e00-\u9fa5]{2,4})?\.\s)/,
    en_au = /([a-zA-Z\-\s\']+,(\s[A-Z]\.){1,3}(,\s[a-zA-Z\-\s\']+,(\s[A-Z]\.){1,3})*(,\s\&\s[a-zA-Z\-\s\']+,(\s[A-Z]\.){1,3})?\s)/,
    au = new RegExp('^(' + Reg_Str(cn_au) + '|' + Reg_Str(en_au) + ')'),
    da = /^\(\d{4}[a-z]?\)\.\s/,
    ti = /^[^]+[\.\?]\s/,
    ge_so = /^([^,\.]{2,},\s\d*(\([\d\-]+\))?(,\s[a-z\d\-]+)+\.\s)/,
    th_so = /^[^,\.]{2,}[,\:]\s[^,\.]{2,}\.\s/,
    so = new RegExp('(' + Reg_Str(ge_so) + '|' + Reg_Str(th_so) + ')(https:\\/\\/[^\.]+\\s)?\\]?\\s*$'),
    reg = new RegExp(Reg_Str(au, da, ti, so));

function Refer() {
    sel = {};
    sel = wps.Selection;
    let ps = sel.Paragraphs, p_space = [],
        chr = ["（）【】，−－–：’？", "()[],---:'?"];
    //纠正符号错误
    for (let c = 0; c < chr[0].length; c++) { sel.Text = sel.Text.replace(RegExp(chr[0][c], 'g'), chr[1][c]) }
    //整体字体格式
    let f = sel.Font;
    f.Size = 9,
        f.Bold = !1,
        f.Italic = !1,
        f.Name = '',
        f.NameAscii = "Times New Roman",
        f.NameFarEast = "宋体";
    //遍历每段
    for (let n = 1; n <= ps.Count; n++) {
        let p = ps.Item(n);
        if (p.Range.Text.length == 1) { p_space.push(p); continue } //跳过空行
        function Modify() {
            //标点后空格
            let str = '';
            for (let w = 1; w <= p.Range.Words.Count; w++) {
                let wrd = p.Range.Words.Item(w).Text, chars = ',.&:?';
                for (let c of chars) { wrd.includes(c) && (wrd = c + ' ') }
                str += wrd
            }
            p.Range.Text = str.replace(/\.\s,/g, '.,').replace(/\s\[/g, ' \n[');
            //拆分元素
            function Rslice(rs, a, b) { let arr = []; for (let i = a; i <= b; i++) { arr.push(rs.Item(i)) } return arr }
            let stcs = p.Range.Sentences, stc_n = stcs.Count;
            if (/^]/.test(stcs.Item(stc_n).Text)) stc_n--; //最后一句为 ] 则取前一句
            let authors = Rslice(stcs, 1, stc_n - 3),
                date = stcs.Item(stc_n - 2),
                title = stcs.Item(stc_n - 1),
                source = stcs.Item(stc_n);
            wps._debug_p = [authors, date.Text, title.Text, source.Text]; //记录拆分信息
            if (stcs.Count < 4) p.Range.Font.Color = 255;
            else Reg_exp(authors, au), Reg_exp(date, da), Reg_exp(title, ti), Reg_exp(source, so); //正则检查标红
            //斜体
            if (source.Text.split(',').length == 2) source = stcs.Item(stcs.Count - 1); //学位论文title斜体
            let wrds = source.Words, has_trial = '';
            for (let w = 1; w <= wrds.Count; w++) {
                has_trial += wrds.Item(w).Text;
                let dots = has_trial.match(/,/g); //逗号数量
                if (wrds.Item(w).Text === '(' || (dots && dots.length == 2)) break;
                else wrds.Item(w).Italic = !0;
            }
            //整体段落格式
            p.Space1(), //*倍行距
                p.SpaceAfter = 0, //段后间距
                p.SpaceBefore = 0,
                p.CharacterUnitLeftIndent = 0, //左缩进量
                p.CharacterUnitRightIndent = 0,
                p.CharacterUnitFirstLineIndent = -2; //悬挂缩进2字符
        }
        try { Modify() } catch (e) { console.log(p.Range.Text + '\n', e); }
    }//*/
    p_space.forEach(e => { e.Range.Text = '' }) //删除空行
    // sel.SortAscending();
    return !0;
}
function test(control) { Refer(); return !0 }
function box_value(text, control) { return !0 }
function OnAction(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnIsEnbable":
            {
                let bFlag = wps.PluginStorage.getItem("EnableFlag")
                wps.PluginStorage.setItem("EnableFlag", !bFlag)
                //通知wps刷新以下几个按饰的状态
                wps.ribbonUI.InvalidateControl("btnIsEnbable")
                wps.ribbonUI.InvalidateControl("btnShowTaskPane")
                //wps.ribbonUI.Invalidate(); 这行代码打开则是刷新所有的按钮状态
                break
            }
        case "btnShowTaskPane":
            {
                let tsId = wps.PluginStorage.getItem("taskpane_id")
                if (!tsId) {
                    let tskpane = wps.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html")
                    let id = tskpane.ID
                    wps.PluginStorage.setItem("taskpane_id", id)
                    tskpane.Visible = true
                } else {
                    let tskpane = wps.GetTaskPane(tsId)
                    tskpane.Visible = !tskpane.Visible
                }
            }
            break
        default:
            break
    }
    return true
}

function GetImage(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            return "images/1.svg"
        case "btnShowDialog":
            return "images/2.svg"
        case "btnShowTaskPane":
            return "images/3.svg"
        default:
            ;
    }
    return "images/newFromTemp.svg"
}
