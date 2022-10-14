function Paper_onload(ribbonUI) {
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

var Rang_Str = function (range) {
    if (Array.isArray(range)) { let str = ''; return (range.forEach(e => { str += e.Text }), str) }
    else return range.Text;
}
var Data_Refs = { values: [], ranges: [], init: function () { this.values = [], this.ranges = [] } },
    _debug = {
        log: '', segments: {}, reds: { now: [], before: [], index: 0 }, num: 0, dbg: void 0,
        init: function () { this.log = '', this.segments = {}, this.reds.now = [], this.num = 0 },
        rcd_segs: function (segs) {
            this.segments = {}
            segs.forEach((e, i) => { this.segments[i + 'N_' + (e.length || 1)] = Rang_Str(e); })
            Data_Refs.values.push(Object.values(this.segments));
            Data_Refs.ranges.push(segs);
        },
        time: function (func, c) {
            let a = new Date().getTime(); func();
            let d = new Date().getTime() - a;
            c && console.log('运行时间：' + d);
            return d
        }
    }

/**
 * 参考文献格式化
 * @param {*} control 
 * @returns 
 */
function Refer(control) {
    {//功能组
        var Reg_Str = function (...regs) {
            let str = '';
            regs.forEach(e => {
                if (typeof e == 'string') str += e;
                else str += e.toString().slice(1, -1);
            });
            return str
        }
        var Reg_exp = function (range, reg, ctrl = true) {
            if (range === null) return 1;
            let str = '', islist = Array.isArray(range);
            if (islist) {
                if (reg.test((range.forEach(e => { str += e.Text }), str))) { ctrl && range.forEach(e => { e.Font.Color = 0 }); return true }
                else { ctrl && (range.forEach(e => { e.Font.Color = 255 }), range[0] && _debug.reds.now.push(range[0])); return false }
            } else {
                if (reg.test(range.Text)) { ctrl && (range.Font.Color = 0); return true }
                else { ctrl && (range.Font.Color = 255, _debug.reds.now.push(range)); return false }
            }
        }
        var isNetname = function (str) { let i = 0; net_head.forEach(e => { i += str.includes(e) }); return i }
        var formated = function () {
            _debug.init();
            _debug.num = sel.Words.Count; //初始词数
            sel.Text = sel.Text.replace(/[\r\n]{3,}/g, '\n'); //删除过多换行
            //纠正符号错误
            for (let c = 0; c < chr_swap[0].length; c++) { sel.Text = sel.Text.replace(RegExp(chr_swap[0][c], 'g'), chr_swap[1][c]) }
            //删除突出强调
            sel.Range.HighlightColorIndex = 0;
            //整体字体格式
            let f = sel.Font;
            f.Size = 9, f.Bold = !1, f.Italic = !1, f.Color = 0,
                f.Name = '', f.NameAscii = "Times New Roman", f.NameFarEast = "宋体";
            //遍历每段
            for (let n = 1; n <= pars.Count; n++) {
                let p = pars.Item(n);
                if (p.Range.Text.length == 1) { p_space.push(p); continue } //跳过空行
                function p_format(p) {
                    //非法符号替换，规定符号后加空格、纠正符号组合
                    function add_spac(Cells) { //查找句子 > 查找单词
                        let str = '';
                        for (let s = 1; s <= Cells.Count; s++) {
                            let stc_t = Cells.Item(s).Text;
                            if (isNetname(stc_t)) {
                                if (Cells.Item(s).Words.Count == 1) stop_add = true;
                                else stc_t = add_spac(Cells.Item(s).Words);
                            }
                            else if (!stop_add) for (let c of chr_spac) {
                                stc_t = stc_t.replace(new RegExp('[ ]*\\' + c + '[ ]*', 'g'), c == '&' ? ' & ' : c + ' ')
                            }
                            str += stc_t;
                        }
                        return str;
                    }
                    let stop_add = false;
                    p.Range.Text = add_spac(p.Range.Sentences)
                        .replace(/\. ,/g, '.,')
                        .replace(/\. \. \./g, '...') //省略号
                        .replace(/,\.\.\./g, ', ...') //逗号+省略号
                        .replace(/ {2,}/g, ' ') //删除过多空格
                        .replace(/\? \. /g, '? ')
                        .replace(/\r{2,}/g, '\r'); //删除空行
                    //分段
                    function p_slice(p) {
                        function r_merge(rs, a, b) { let arr = []; for (let i = a; i <= b; i++) { arr.push(rs.Item(i)) } return arr }
                        let authors, date, title, source; //各成分
                        let stcs = p.Range.Sentences, stc_n = stcs.Count;
                        if (stc_n < 4) return false;
                        //最后一句首词为 net_head ? link,so : so
                        [1, 3].forEach(() => {
                            let last = (n) => { return stcs.Item(n).Words.Item(1).Text }
                            if (isNetname(last(stc_n)) || last(stc_n)[0] == ']' ||
                                isNetname(last(stc_n - 1)) || last(stc_n - 1)[0] == ']') stc_n--;
                        });
                        source = stcs.Item(stc_n);
                        //匹配日期 ? date : 标红加粗
                        for (let i = 1; i < stc_n; i++) {
                            let s = stcs.Item(i);
                            if (Reg_exp(s, da, false)) {
                                if (1 > i - 1 || i + 1 > stc_n - 1) console.log('date位置错误 ' + i + ' ' + stc_n);
                                date = s;
                                authors = r_merge(stcs, 1, i - 1);
                                title = r_merge(stcs, i + 1, stc_n - 1);
                                break;
                            }
                        }
                        if (date === undefined)
                            return false;
                        else
                            return [authors, date, title, source]; //可能含有 []
                    }
                    let p_segs = p_slice(p), isTh = false;
                    p_segs && _debug.rcd_segs(p_segs);
                    //检测类型：期刊文献，学位论文
                    if (!p_segs) {
                        p.Range.Font.Bold = !0;
                        p.Range.Font.Color = 255;
                        _debug.reds.now.push(p.Range);
                    } else if (/((硕|博)士学位论文|Unpublished (master\'s thesis|doctorial dissertation))/.test(p.Range.Text)) {
                        Reg_exp(p_segs[0], au), Reg_exp(p_segs[1], da), Reg_exp(p_segs[2], th_ti), Reg_exp(p_segs[3], th_so);
                        _debug.log += 'thesis\n' + [au, da, th_ti, th_so].join('\n');
                        isTh = true;
                    } else {
                        Reg_exp(p_segs[0], au), Reg_exp(p_segs[1], da), Reg_exp(p_segs[2], ge_ti), Reg_exp(p_segs[3], ge_so);
                        _debug.log += 'journal\n' + [au, da, ge_ti, ge_so].join('\n');
                    }
                    //斜体
                    let trial_range = isTh ? p_segs[2].slice(-1)[0] : p_segs[3],
                        wrds = trial_range && trial_range.Words,
                        has_trial = '';
                    if (wrds) for (let w = 1; w <= wrds.Count; w++) {
                        has_trial += wrds.Item(w).Text;
                        if (trial_range.Text.includes('(') && /\($/.test(has_trial.slice(-1))) break;
                        else if (/ ?\d+, ?/.test(has_trial.slice(-3))) break;
                        wrds.Item(w).Italic = !0;
                    }
                    //段落格式
                    p.Space15(), //*倍行距
                        p.SpaceAfter = 0, //段后间距
                        p.SpaceBefore = 0,
                        p.CharacterUnitLeftIndent = 0, //左缩进量
                        p.CharacterUnitRightIndent = 0,
                        p.CharacterUnitFirstLineIndent = -2; //悬挂缩进2字符
                }
                try { p_format(p) } catch (e) { p.Range.HighlightColorIndex = 7; console.log(p.Range.Text + '\n', e) }
            }
            return !0;
        }
    }
    {//变量
        var cn_au = /(^\[?[\u4e00-\u9fa5]{2,4}(, [\u4e00-\u9fa5]{2,4}){0,5}(, (\.\.\. )?[\u4e00-\u9fa5]{2,4})?\. $)/,
            en_au = /(^[a-zA-Z\-\']+,( [A-Z]\.){1,3}(, [a-zA-Z\-\']+,( [A-Z]\.){1,3}){0,5}(, (\&|\.\.\.) [a-zA-Z\-\']+,( [A-Z]\.){1,3})? $)/,
            au = new RegExp('^(' + Reg_Str(cn_au) + '|' + Reg_Str(en_au) + ')'),
            da = /(\(\d{4}[a-z]?\)\. )/, //(1234b). /
            ti = /([^]+[\.\?]|[^]+\? [^]+\.) /,
            ge_so = /([a-zA-Z\u4e00-\u9fa5,\-\&\:\(\)\' ]{2,}, \d+(\(\d+(\-\d+)?\))?, (Article )?[a-zA-Z]?\d+([\-\+][a-zA-Z]?\d+){0,2}\. \r*)/,
            th_so = /(^[a-zA-Z\u4e00-\u9fa5\' ]{4,}(, [a-zA-Z\u4e00-\u9fa5\' ]{2,}){0,3}\. \r*$)/,
            so = new RegExp('^(' + Reg_Str(ge_so) + '|' + Reg_Str(th_so) + ')\]?\\r*'),
            link = /(^(https?|doi)\:\/\/[a-zA-Z\#\?\.\/]*\.\s*)?$/,
            reg = new RegExp(Reg_Str(au, da, ti, so)),
            net_head = ['doi', 'http'];
        var sel = wps.Selection,
            pars = sel.Paragraphs,
            p_space = [],
            chr_swap = ["（）【】，−－–：’？！ ．", "()[],---:'?! ."], chr_spac = ',.&:?…!',
            ge_ti = /[^\(\)]*. /,
            th_ti = /[^]*[a-zA-Z\u4e00-\u9fa5] \(((硕|博)士学位论文|Unpublished (master\'s thesis|doctorial dissertation))\). /;
    }
    switch (control.Id) {
        case 'format':
            if (pars.Count != 1) { //不止选择一段
                Data_Refs.init();
                console.log('每秒词数：' + 1000 / (_debug.time(() => { formated() }, !1) / _debug.num));
                _debug.log = _debug.reds.now.length + '个疑似错误^_^';
                alert('检查成功，' + _debug.reds.now.length + '个疑似错误');
                _debug.reds.now[0] && _debug.reds.now[0].Select(); //选中第一个red
                _debug.reds.before = _debug.reds.now; //更新reds库
            } else formated();
            break;
        case 'up_err':
            if (_debug.reds.index != undefined) {
                0 < _debug.reds.index && _debug.reds.index--;
                _debug.reds.before[0] && _debug.reds.before[_debug.reds.index].Select();
            } break;
        case 'down_err':
            if (_debug.reds.index != undefined) {
                _debug.reds.before.length - 1 > _debug.reds.index && _debug.reds.index++;
                _debug.reds.before[0] && _debug.reds.before[_debug.reds.index].Select();
            } break;
        default: break;
    }
    return !0;
}

var dependent_variables = [],
    independent_variables = [];
/**
 * 核对索引
 * @returns 
 */
function Discuss() {
    //字符串正则不匹配，返回空数组
    const _match = String.prototype.match;
    String.prototype.match = function (reg) {
        let res = _match.call(this, reg);
        return res == null ? [] : res;
    }
    //删除数组重复元素
    Array.prototype.delRepeatElements = function () {
        let _this = [];
        this.forEach((e, i) => { _this[i] = e + '' }); //元素转化为字符串
        _this = [...new Set(_this)[Symbol.iterator]()]; //删除元素
        _this.forEach((e, i) => { _this[i] = e.split(',') }); //元素转化为数组
        return _this;
    }
    var sel = wps.Selection, txt = sel.Text;
    //纠正错误字符
    chr_swap = ['（）', '()'];
    // for (let c = 0; c < chr_swap[0].length; c++) { sel.Text = sel.Text.replace(RegExp(chr_swap[0][c], 'g'), chr_swap[1][c]) }
    try {
        //提取正文索引
        var mt = /[^a-zA-Z\u4e00-\u9fa5\(\)]+[a-zA-Z\u4e00-\u9fa5]+(等人)?\(\d{4}\)/g, //。abc等人(2333)
            gt = /\([^\(\)]+\d{4}\)/g;
        var Date_mainTxt = [], arr = [];
        for (let i of txt.match(gt)) {
            arr.push(...i.split('; '))
        }
        Date_mainTxt = [...txt.match(mt), ...arr];
        console.log('正文中的索引:\n', Date_mainTxt)
        console.log('参考文献的索引:\n', Data_Refs.values);
        //比对正文索引和参考文献
        var match_ok = [], match_lack = [];
        Data_Refs.values.forEach((ref, i) => { //遍历参考文献，事先标蓝
            Data_Refs.ranges[i][1].Font.ColorIndex = 2;
        });
        Date_mainTxt.forEach((cite) => { //遍历正文索引
            let find_times = 0;
            Data_Refs.values.forEach((ref, i) => { //遍历参考文献
                let name = ref[0].split(', ')[0].replace(/\. /g, ''), //Dickman
                    date = ref[1].split('. ')[0].slice(1, -1); //2020
                if (cite.includes(name) && cite.includes(date)) {
                    find_times++;
                    match_ok.push([name, date]);
                    Data_Refs.ranges[i][1].Font.ColorIndex = 1;
                }
            });
            if (find_times == 0) match_lack.push(cite);
        });
        console.log('匹配成功的索引:\n', match_ok.delRepeatElements());
        console.log('缺少参考文献的正文索引:\n', match_lack);
    } catch (err) { console.log('\n', err) }
    return !0;
}

function box_value(text, control) { return '' + _debug.num }
function ShowTaskPane(control) {
    let tsId = wps.PluginStorage.getItem("taskpane_id")
    if (!tsId) {
        let tskpane = wps.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html");
        wps.PluginStorage.setItem("taskpane_id", tskpane.ID),
            tskpane.Visible = true
    } else {
        let tskpane = wps.GetTaskPane(tsId)
        tskpane.Visible = !tskpane.Visible
    }
    return true
}
