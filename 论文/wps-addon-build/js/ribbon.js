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
function Reg_exp(range, reg, ctrl = true) {
    if (range === null) return 1;
    let str = '', islist = Array.isArray(range);
    if (islist) {
        if (reg.test((range.forEach(e => { str += e.Text }), str))) { ctrl && range.forEach(e => { e.Font.Color = 0 }); return true }
        else { ctrl && range.forEach(e => { e.Font.Color = 255 }); return false }
    }
    else {
        if (reg.test(range.Text)) { ctrl && (range.Font.Color = 0); return true }
        else { ctrl && (range.Font.Color = 255); return false }
    }
}
function Rang_Str(range) {
    if (Array.isArray(range)) { let str = ''; return (range.forEach(e => { str += e.Text }), str) }
    else return range.Text;
}
function isNet(str) { let i = 0; net_head.forEach(e => { i += str.includes(e) }); return i }

var WebNotifycount = 0, err, sel,
    cn_au = /(\[?[\u4e00-\u9fa5]{2,4}(,\s[\u4e00-\u9fa5]{2,4})*(,\s(\&|\…|\.\s\.\s\.)\s[\u4e00-\u9fa5]{2,4})?\.\s)/,
    en_au = /([a-zA-Z\-\s\']+,(\s[A-Z]\.){1,3}(,\s[a-zA-Z\-\s\']+,(\s[A-Z]\.){1,3})*(,\s(\&|\…|\.\s\.\s\.)\s[a-zA-Z\-\s\']+,(\s[A-Z]\.){1,3})?\s)/,
    au = new RegExp('^(' + Reg_Str(cn_au) + '|' + Reg_Str(en_au) + ')'),
    da = /^\(\d{4}[a-z]?\)\.\s/, //(1234b). /
    ti = /^([^]+[\.\?]|[^]+\?\s[^]+\.)\s/,
    ge_so = /^([a-zA-Z\u4e00-\u9fa5,\-\&\s\:\(\)]{2,},\s\d+(\(\d+(\-\d+)?\))?,\s(Article\s)?[a-zA-Z]?\d+([\-\+][a-zA-Z]?\d+){0,2}\.\s)/,
    th_so = /^[a-zA-Z\u4e00-\u9fa5,\-\&\s\:\(\)]{2,}[,\:]\s[a-zA-Z\u4e00-\u9fa5\s\']{2,}\.\s/, //**, **. //**: **. /
    so = new RegExp('(' + Reg_Str(ge_so) + '|' + Reg_Str(th_so) + ')\\]?\[^\]*$'),
    reg = new RegExp(Reg_Str(au, da, ti, so)),
    net_head = ['doi', 'http'],
    _debug = {
        txt: '', reds: 0, num: 0,
        init: function () { this.txt = '', this.reds = 0 },
        split_p: function (a, d, t, s) {
            this.txt = '';
            this.txt = ['===分段信息===', '作者>' + Rang_Str(a), '日期>' + Rang_Str(d), '标题>' + Rang_Str(t), '来源>' + Rang_Str(s)].join('\n')
        },
        time: function (func, c) { let a = new Date().getTime(); func(); let d = new Date().getTime() - a; !c && console.log(d); return d }
    };

function Refer() {
    _debug.init(); //初始化
    sel = wps.Selection;
    let ps = sel.Paragraphs, p_space = [],
        chr_swap = ["（）【】，−－–：’？", "()[],---:'?"], chr_spac = ',.&:?…';
    _debug.num = sel.Words.Count; //初始行数
    sel.Text = sel.Text.replace(/\r{3,}/g, '\n'); //删除过多换行
    //纠正符号错误
    for (let c = 0; c < chr_swap[0].length; c++) { sel.Text = sel.Text.replace(RegExp(chr_swap[0][c], 'g'), chr_swap[1][c]) }
    //整体字体格式
    let f = sel.Font;
    f.Size = 9,
        f.Bold = !1,
        f.Italic = !1,
        f.Color = 0,
        f.Name = '',
        f.NameAscii = "Times New Roman",
        f.NameFarEast = "宋体";
    //遍历每段
    for (let n = 1; n <= ps.Count; n++) {
        let p = ps.Item(n);
        if (p.Range.Text.length == 1) { p_space.push(p); continue } //跳过空行
        function Modify() {
            //标点后空格
            function Add_spac(Cells) { //查找句子 > 查找单词
                let str = '';
                for (let s = 1; s <= Cells.Count; s++) {
                    let stc_t = Cells.Item(s).Text;
                    if (isNet(stc_t)) {
                        if (Cells.Item(s).Words.Count == 1) stop_add = true;
                        else stc_t = Add_spac(Cells.Item(s).Words);
                    }
                    else if (!stop_add) for (let c of chr_spac) { stc_t = stc_t.replace(new RegExp('\\' + c + '[ ]*', 'g'), c + ' ') }
                    str += stc_t;
                }
                return str;
            }
            let stop_add = false;
            p.Range.Text = Add_spac(p.Range.Sentences)
                .replace(/\.\s,/g, '.,')
                .replace(/\s\[/g, ' \n[')
                .replace(/\?\s\.\s/g, '? ')
                .replace(/\r{2,}/g, '\r');
            //拆分元素
            function Rslice(rs, a, b) { let arr = []; for (let i = a; i <= b; i++) { arr.push(rs.Item(i)) } return arr }
            let authors, date, title, source; //各成分
            let stcs = p.Range.Sentences, stc_n = stcs.Count, lastHeadText = stcs.Item(stc_n).Words.Item(1).Text;
            if (isNet(lastHeadText) || lastHeadText.includes(']')) stc_n--; //首词为 ] net_head 则source取前一句
            source = stcs.Item(stc_n);
            for (let i = 1; i <= stc_n; i++) {
                let s = stcs.Item(i);
                if (Reg_exp(s, da, false)) {
                    if (1 > i - 1 || i + 1 > stc_n - 1) console.log(p.Range.Text + '\ndate位置错误 ' + i + ' ' + stc_n);
                    date = s,
                        authors = Rslice(stcs, 1, i - 1),
                        title = Rslice(stcs, i + 1, stc_n - 1)
                }
            }
            if (date === undefined) { //date无匹配
                // p.Range.Text = '>>. (2333a). >>' + p.Range.Text,
                p.Range.Font.Bold = !0,
                    p.Range.Font.Color = 255;
            } else {
                //拆分成功
                if (stcs.Count < 4) p.Range.Font.Color = 255;
                else (
                    Reg_exp(authors, au) * Reg_exp(date, da) * Reg_exp(title, ti) * Reg_exp(source, so) == 0 && _debug.reds++,//正则错误标红
                    _debug.split_p(authors, date, title, source) //记录调试信息
                );
                //斜体
                if (Reg_exp(source, th_so, false)) stc_n--; //学位论文title斜体
                let trial_txt = stcs.Item(stc_n), wrds = trial_txt.Words, has_trial = '';
                for (let w = 1; w <= wrds.Count; w++) {
                    has_trial += wrds.Item(w).Text;
                    if (trial_txt.Text.includes('(')) {
                        if (wrds.Item(w).Text === '(') break;
                    } else {
                        if (/\s?\d,\s?/.test(has_trial.slice(-3))) break;
                    }
                    wrds.Item(w).Italic = !0;
                }
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

    // sel.SortAscending();
    return !0;
}
function test() { console.log(1000 / (_debug.time(() => { Refer() }, !1) / _debug.num)); return !0 }
function box_value(text, control) { return true }
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
