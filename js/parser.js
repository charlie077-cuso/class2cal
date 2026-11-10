/* ============================================================
 * 课表解析模块（Excel 课表 → 结构化课程）
 * 浏览器：window.TT.parseWorkbook / listSheets / defaultSlots / parseWeeks
 * Node：  module.exports = { ... }
 * 依赖：SheetJS（浏览器全局 XLSX，Node require('xlsx')）
 * ============================================================ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('xlsx'));
  } else {
    root.TT = root.TT || {};
    var api = factory(root.XLSX);
    for (var k in api) if (Object.prototype.hasOwnProperty.call(api, k)) root.TT[k] = api[k];
  }
})(typeof self !== 'undefined' ? self : this, function (XLSX) {
  'use strict';

  var VERSION = '1.0.0';
  var DAY_WORDS = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '日': 7, '天': 7 };
  var DAY_EN = {
    monday: 1, mon: 1, tuesday: 2, tue: 2, tues: 2, wednesday: 3, wed: 3,
    thursday: 4, thu: 4, thur: 4, thurs: 4, friday: 5, fri: 5,
    saturday: 6, sat: 6, sunday: 7, sun: 7
  };
  var CN_DIGITS = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 };

  /* ---------------- 基础工具 ---------------- */

  function str(v) {
    return v == null ? '' : String(v);
  }

  function trim(v) {
    return str(v).replace(/\u00a0/g, ' ').trim();
  }

  function normalizeText(v) {
    return trim(v).replace(/\r\n?/g, '\n');
  }

  function cnToNumber(s) {
    s = trim(s);
    if (!s) return NaN;
    if (/^\d+$/.test(s)) return +s;
    if (s === '十') return 10;
    if (/^十[一二三四五六七八九]$/.test(s)) return 10 + CN_DIGITS[s[1]];
    if (/^[一二三四五六七八九]十$/.test(s)) return CN_DIGITS[s[0]] * 10;
    if (/^[一二三四五六七八九]十[一二三四五六七八九]$/.test(s)) return CN_DIGITS[s[0]] * 10 + CN_DIGITS[s[2]];
    if (CN_DIGITS[s]) return CN_DIGITS[s];
    return NaN;
  }

  function uniqueSorted(arr) {
    var seen = {}, out = [];
    arr.forEach(function (n) {
      if (Number.isFinite(n) && n > 0 && !seen[n]) { seen[n] = 1; out.push(n); }
    });
    return out.sort(function (a, b) { return a - b; });
  }

  function rangeNumbers(a, b) {
    if (!(a > 0) || !(b > 0)) return [];
    var lo = Math.min(a, b), hi = Math.max(a, b), out = [];
    for (var i = lo; i <= hi && i < lo + 60; i++) out.push(i);
    return out;
  }

  /* ---------------- 读取工作簿 ---------------- */

  /* 是否为真正的二进制 Excel（xlsx=ZIP / xls=OLE）；否则按文本处理（CSV / HTML 伪装成 xls） */
  function isBinarySpreadsheet(buf) {
    if (!buf || !buf.length) return false;
    if (buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4b) return true;
    if (buf.length >= 4 && buf[0] === 0xd0 && buf[1] === 0xcf && buf[2] === 0x11 && buf[3] === 0xe0) return true;
    return false;
  }

  function decodeTextFile(buf) {
    var bytes = null;
    if (typeof Buffer !== 'undefined' && Buffer.isBuffer(buf)) bytes = buf;
    else if (buf instanceof Uint8Array) bytes = buf;
    else if (buf instanceof ArrayBuffer) bytes = new Uint8Array(buf);
    if (!bytes) return String(buf);
    if (typeof TextDecoder !== 'undefined') {
      try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^\uFEFF/, ''); } catch (e) {}
      try { return new TextDecoder('gbk').decode(bytes); } catch (e) {}
    }
    var out = '';
    for (var i = 0; i < bytes.length; i++) out += String.fromCharCode(bytes[i]);
    return out;
  }

  function readWorkbook(input) {
    if (input && typeof input === 'object' && input.SheetNames && input.Sheets) return input;
    var data = input;
    var type = 'array';
    if (typeof Buffer !== 'undefined' && Buffer.isBuffer(input)) {
      data = input;
      type = 'buffer';
    } else if (typeof ArrayBuffer !== 'undefined' && input instanceof ArrayBuffer) {
      data = new Uint8Array(input);
      type = 'array';
    } else if (ArrayBuffer.isView(input)) {
      data = new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
      type = 'array';
    }
    var bytes = type === 'buffer' ? data : (data instanceof Uint8Array ? data : new Uint8Array(data));
    if (!isBinarySpreadsheet(bytes)) {
      return XLSX.read(decodeTextFile(bytes), { type: 'string', cellDates: true, cellNF: false, cellText: false });
    }
    return XLSX.read(data, { type: type, cellDates: true, cellNF: false, cellText: false });
  }

  function sheetMatrix(ws) {
    if (!ws) return { matrix: [], merges: [] };
    var matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: true });
    var merges = (ws['!merges'] || []).map(function (m) {
      return { s: { r: m.s.r, c: m.s.c }, e: { r: m.e.r, c: m.e.c } };
    });
    // 用合并区域左上角的值填充空单元格（表头可能跨列）
    merges.forEach(function (m) {
      var v = matrix[m.s.r] && matrix[m.s.r][m.s.c];
      if (v === '' || v == null) return;
      for (var r = m.s.r; r <= m.e.r; r++) {
        if (!matrix[r]) matrix[r] = [];
        for (var c = m.s.c; c <= m.e.c; c++) {
          if (matrix[r][c] === '' || matrix[r][c] == null) matrix[r][c] = v;
        }
      }
    });
    return { matrix: matrix, merges: merges };
  }

  function findMerge(merges, r, c) {
    for (var i = 0; i < merges.length; i++) {
      var m = merges[i];
      if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) return m;
    }
    return null;
  }

  /* ---------------- 星期 / 表头 ---------------- */

  function weekdayFromDate(y, m, d) {
    if (!y || y < 1900) return 0;
    var dt = new Date(y, m - 1, d);
    if (isNaN(dt.getTime())) return 0;
    var w = dt.getDay();
    return w === 0 ? 7 : w;
  }

  function parseWeekdayCell(v, yearHint) {
    if (v == null || v === '') return 0;
    if (v instanceof Date) {
      var d = v.getDay();
      return d === 0 ? 7 : d;
    }
    var s = trim(v);
    if (!s) return 0;
    var low = s.toLowerCase();
    var tokens = low.replace(/[^a-z]+/g, ' ').split(/\s+/).filter(Boolean);
    for (var i = 0; i < tokens.length; i++) {
      if (DAY_EN[tokens[i]]) return DAY_EN[tokens[i]];
    }
    var m = /(?:星期|周|礼拜)\s*([一二三四五六日天1-7])/.exec(s);
    if (m) return /[1-7]/.test(m[1]) ? +m[1] : DAY_WORDS[m[1]];
    if (/^[一二三四五六日天]$/.test(s)) return DAY_WORDS[s];
    // 表头用日期表示：2026-09-07 / 2026/9/7 / 2026.9.7；没有年份时用学年学期提示
    var dm = /^(\d{4})\s*[-\/.年]\s*(\d{1,2})\s*[-\/.月]\s*(\d{1,2})\s*日?$/.exec(s);
    if (dm) return weekdayFromDate(+dm[1], +dm[2], +dm[3]);
    var md = /^(\d{1,2})\s*[月\/.\-]\s*(\d{1,2})\s*日?$/.exec(s);
    if (md && yearHint) {
      var mm = +md[1];
      return weekdayFromDate(mm >= 8 ? yearHint : yearHint + 1, mm, +md[2]);
    }
    return 0;
  }

  /* 从「2026-2027-1」这类学期文本取起始年份，用于无年份的日期表头 */
  function semesterYear(matrix) {
    var lim = Math.min(matrix.length, 12);
    for (var r = 0; r < lim; r++) {
      var row = matrix[r] || [];
      for (var c = 0; c < row.length; c++) {
        var s = String(cellAt(matrix, r, c) == null ? '' : cellAt(matrix, r, c));
        var m = /(20\d{2})\s*[-–—~至]\s*(20\d{2})/.exec(s);
        if (m) return +m[1];
      }
    }
    return 0;
  }

  function cellAt(matrix, r, c) {
    return matrix[r] ? matrix[r][c] : '';
  }

  function detectHeader(matrix, yearHint) {
    var best = null;
    var maxRow = Math.min(matrix.length, 25);
    for (var r = 0; r < maxRow; r++) {
      var map = {}, count = 0;
      var row = matrix[r] || [];
      for (var c = 0; c < row.length; c++) {
        var day = parseWeekdayCell(row[c], yearHint);
        if (day && !map[day]) { map[day] = c; count++; }
      }
      if (count >= 2 && (!best || count > best.count)) {
        best = { row: r, weekdayCols: map, count: count };
      }
    }
    return best;
  }

  /* 竖排课表：星期在一列里（纵向），节次在一行里（横向） */
  function detectTransposedHeader(matrix, yearHint) {
    var maxC = Math.min(matrix.reduce(function (m, row) { return Math.max(m, row ? row.length : 0); }, 0), 20);
    var best = null;
    for (var c = 0; c < maxC; c++) {
      var map = {}, count = 0;
      var lim = Math.min(matrix.length, 40);
      for (var r = 0; r < lim; r++) {
        var day = parseWeekdayCell(cellAt(matrix, r, c), yearHint);
        if (day && !map[day]) { map[day] = r; count++; }
      }
      if (count >= 3 && (!best || count > best.count)) {
        best = { col: c, weekdayRows: map, count: count };
      }
    }
    return best;
  }

  function transposeSheet(matrix, merges) {
    var maxC = matrix.reduce(function (m, row) { return Math.max(m, row ? row.length : 0); }, 0);
    var out = [];
    for (var c = 0; c < maxC; c++) {
      var row = [];
      for (var r = 0; r < matrix.length; r++) row.push(cellAt(matrix, r, c));
      out.push(row);
    }
    var tmerges = (merges || []).map(function (m) {
      return { s: { r: m.s.c, c: m.s.r }, e: { r: m.e.c, c: m.e.r } };
    });
    return { matrix: out, merges: tmerges };
  }

  /* ---------------- 节次 / 时段 ---------------- */

  function parsePeriods(text) {
    var raw = String(text == null ? '' : text).replace(/\r/g, '');
    if (!raw.trim()) return [];
    var s = raw.replace(/\u00a0/g, ' ');

    // 第1-2节 / 1-2节 / 第1~2节
    var m = /(?:第)?\s*([0-9一二三四五六七八九十]+)\s*[-~至—－]\s*([0-9一二三四五六七八九十]+)\s*节/.exec(s);
    if (m) {
      var a = cnToNumber(m[1]), b = cnToNumber(m[2]);
      if (a > 0 && b > 0) return uniqueSorted(rangeNumbers(a, b));
    }
    // 第1、2节 / 1,2节 / 第1，2节
    var mlist = /(?:第)?\s*([0-9]{1,2}(?:\s*[、,，]\s*[0-9]{1,2})+)\s*节/.exec(s);
    if (mlist) {
      var nums = mlist[1].split(/[、,，]/).map(function (t) { return +trim(t); }).filter(function (n) { return n > 0 && n <= 20; });
      if (nums.length) return uniqueSorted(nums);
    }
    // 第一大节 / 第1大节 / 大节一 → 每大节按 2 节算（[01-02]节、[03-04]节…）
    var mBig = /(?:第)?\s*([0-9一二三四五六七八九十]+)\s*大节/.exec(s);
    if (mBig) {
      var bn = cnToNumber(mBig[1]);
      if (bn > 0) return [bn * 2 - 1, bn * 2];
    }
    // 第3节
    var m2 = /第\s*([0-9一二三四五六七八九十]+)\s*节/.exec(s);
    if (m2) {
      var n = cnToNumber(m2[1]);
      if (n > 0) return [n];
    }
    // 0102 / 0910 / 1112 这类连续两位数
    var firstLine = s.split('\n')[0].replace(/[^0-9]/g, '');
    if (/^\d{4}$/.test(firstLine)) {
      var p1 = +firstLine.slice(0, 2), p2 = +firstLine.slice(2, 4);
      if (p1 > 0 && p2 > p1 && p2 - p1 === 1 && p2 <= 20) return [p1, p2];
    }
    if (/^\d{2}$/.test(firstLine)) {
      var p = +firstLine;
      if (p > 0 && p <= 20) return [p];
    }
    // 单独的 "1-2"（没有“节”字）
    var m3 = /^\s*([0-9]{1,2})\s*[-~至—－]\s*([0-9]{1,2})\s*$/.exec(trim(s));
    if (m3) {
      var x = +m3[1], y = +m3[2];
      if (x > 0 && y >= x && y <= 20) return uniqueSorted(rangeNumbers(x, y));
    }
    // 纯数字节次标签（如 "5"）
    var m4 = /^\s*([0-9]{1,2})\s*$/.exec(trim(s));
    if (m4) {
      var z = +m4[1];
      if (z > 0 && z <= 20) return [z];
    }
    return [];
  }

  function isSlotLabel(text) {
    var t = trim(text);
    if (!t) return false;
    if (parsePeriods(t).length) return true;
    return /^(上午|下午|晚上|中午)/.test(t);
  }

  function parseTimeRanges(text) {
    var s = String(text == null ? '' : text);
    var out = [];
    var re = /([01]?\d|2[0-3])\s*[:：]\s*([0-5]\d)\s*[-~至—－]\s*([01]?\d|2[0-3])\s*[:：]\s*([0-5]\d)/g;
    var m;
    while ((m = re.exec(s))) {
      out.push({
        start: pad2(+m[1]) + ':' + m[2],
        end: pad2(+m[3]) + ':' + m[4]
      });
    }
    return out;
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  function slotFromLabel(label, timeText) {
    var clean = trim(String(label).replace(/\s+/g, ' '));
    var ranges = parseTimeRanges(timeText || label);
    return {
      label: clean || trim(label),
      periods: parsePeriods(label),
      start: ranges.length ? ranges[0].start : null,
      end: ranges.length ? ranges[ranges.length - 1].end : null,
      raw: String(label == null ? '' : label)
    };
  }

  function detectSlotColumn(matrix, header, merges) {
    var minDayCol = 99;
    Object.keys(header.weekdayCols).forEach(function (d) {
      minDayCol = Math.min(minDayCol, header.weekdayCols[d]);
    });
    var maxC = matrix.reduce(function (m, row) { return Math.max(m, row ? row.length : 0); }, 0);
    var candidates = [];
    for (var c = 0; c < maxC; c++) {
      if (c >= minDayCol) continue;
      var hits = 0, timeHits = 0;
      for (var r = header.row + 1; r < matrix.length && r <= header.row + 60; r++) {
        var v = cellAt(matrix, r, c);
        if (isSlotLabel(v)) hits++;
        if (parseTimeRanges(v).length) timeHits++;
      }
      candidates.push({ col: c, hits: hits, timeHits: timeHits });
    }
    candidates.sort(function (a, b) { return (b.hits - a.hits) || (b.timeHits - a.timeHits); });
    var periodCol = candidates.length && candidates[0].hits > 0 ? candidates[0].col : -1;
    var timeCol = -1;
    candidates.forEach(function (x) {
      if (x.col !== periodCol && x.timeHits >= 2 && (timeCol === -1 || x.timeHits > candidates.filter(function (y) { return y.col === timeCol; })[0].timeHits)) timeCol = x.col;
    });
    return { periodCol: periodCol, timeCol: timeCol };
  }

  /* ---------------- 周次 ---------------- */

  function parseWeeks(text) {
    var s = String(text == null ? '' : text).replace(/\u00a0/g, ' ');
    if (!s.trim()) return [];
    var parity = 0;
    if (/单周/.test(s)) parity = 1;
    else if (/双周/.test(s)) parity = 2;
    // 去掉 [..] 与 (..) 内容（[周]、[01-02节]、([单周])）
    var cleaned = s.replace(/\[[^\]]*\]/g, ' ').replace(/\([^)]*\)/g, ' ').replace(/（[^）]*）/g, ' ');
    // 去掉“第X周”的“周”字与其它非数字
    cleaned = cleaned.replace(/[周]/g, ' ');
    var tokens = cleaned.split(/[^0-9\-~至—－]+/).map(trim).filter(Boolean);
    var weeks = [];
    tokens.forEach(function (tok) {
      var m = /^(\d{1,2})\s*[-~至—－]\s*(\d{1,2})$/.exec(tok);
      if (m) { weeks = weeks.concat(rangeNumbers(+m[1], +m[2])); return; }
      var n = /^(\d{1,2})$/.exec(tok);
      if (n) weeks.push(+n[1]);
    });
    weeks = uniqueSorted(weeks);
    if (parity === 1) weeks = weeks.filter(function (w) { return w % 2 === 1; });
    if (parity === 2) weeks = weeks.filter(function (w) { return w % 2 === 0; });
    if (!weeks.length && /^\s*\d+\s*$/.test(cleaned)) weeks = [+cleaned.trim()];
    return weeks;
  }

  /* ---------------- 单元格课程块 ---------------- */

  function parseCellBlocks(text) {
    var s = normalizeText(text);
    if (!s.trim()) return [];
    // 部分系统用一行横线/等号分隔同一格里的多门课
    s = s.replace(/\n[ \t]*[-=_*—－]{3,}[ \t]*\n/g, '\n\n');
    // 课程之间用空行分隔；先按空行切分，再清理
    var parts = s.split(/\n[ \t]*\n+/);
    if (parts.length === 1) {
      // 有些文件只用单个换行分隔多门课：以“周次行”作为块边界
      var lines = s.split('\n');
      var hasWeekLine = lines.some(function (l) { return /\[[^\]]*周[^\]]*\]/.test(l) || /\d+\s*周/.test(l); });
      if (hasWeekLine && lines.filter(function (l) { return trim(l); }).length > 6) {
        parts = [];
        var cur = [];
        lines.forEach(function (l) {
          cur.push(l);
          if (/\[[^\]]*周[^\]]*\]\s*(\[[^\]]*节\])?\s*$/.test(trim(l))) { parts.push(cur.join('\n')); cur = []; }
        });
        if (cur.length) parts.push(cur.join('\n'));
      }
    }
    return parts.map(function (p) { return trim(p); }).filter(Boolean);
  }

  function looksLikeLocation(line) {
    return /(楼|室|馆|场|校区|教室|机房|实训|基地|中心|操场|体育场|报告厅|实验室)/.test(line);
  }

  function looksLikeTeacher(line) {
    if (!line) return false;
    if (/\d/.test(line)) return false;
    if (/(周|节|楼|室|馆|场|校区|教室|分组)/.test(line)) return false;
    return line.length <= 24;
  }

  function parseCourseBlock(block) {
    var lines = block.split('\n').map(trim).filter(Boolean);
    if (!lines.length) return null;
    var out = {
      name: lines[0],
      teacher: '',
      location: '',
      weeksText: '',
      weeks: [],
      note: '',
      raw: block
    };
    var weekIdx = -1;
    for (var i = 1; i < lines.length; i++) {
      var line = lines[i];
      if (weekIdx === -1 && /周/.test(line) && /[0-9]/.test(line)) { weekIdx = i; out.weeksText = line; out.weeks = parseWeeks(line); continue; }
      if (/分组/.test(line)) { if (!out.note) out.note = line; }
    }
    function isPeriodOnly(line) {
      return /^\[?\s*\d{1,2}\s*[-~,，]\s*\d{1,2}\s*\]?\s*节?$/.test(trim(line)) ||
        /^\[?\s*\d{1,2}\s*\]?\s*节$/.test(trim(line));
    }
    if (weekIdx > 0) {
      // 教师：周次行之前、且不是分组备注的最后一行
      for (var j = weekIdx - 1; j >= 1; j--) {
        if (/^[（(]/.test(lines[j]) || /分组/.test(lines[j])) continue;
        if (looksLikeTeacher(lines[j]) || (lines[j] && !looksLikeLocation(lines[j]))) { out.teacher = lines[j]; break; }
      }
      // 地点：周次行之后的第一行（跳过 [01-02]节 这类纯节次行）
      for (var k = weekIdx + 1; k < lines.length; k++) {
        if (/^[（(]/.test(lines[k])) continue;
        if (isPeriodOnly(lines[k])) continue;
        out.location = lines[k];
        break;
      }
      if (!out.location && lines.length > weekIdx + 1) out.location = lines[lines.length - 1];
    } else {
      // 没有周次行：第二行当教师，最后一行若像地点则当地点
      if (lines.length >= 2 && !/^[（(]/.test(lines[1])) out.teacher = lines[1];
      if (lines.length >= 3 && looksLikeLocation(lines[lines.length - 1])) out.location = lines[lines.length - 1];
    }
    out.name = out.name.replace(/^[■□◆◇●○★☆▶►▷·•\s\-—]+/, '').trim();
    out.teacher = trim(String(out.teacher).replace(/^(任课教师|授课教师|教师|老师)\s*[：:]\s*/, ''));
    out.location = trim(String(out.location).replace(/^(上课地点|地点|教室)\s*[：:]\s*/, '').replace(/\[?\s*\d{1,2}\s*[-~,，]\s*\d{1,2}\s*\]?\s*节\s*$/, ''));
    // 仅当括号内容是分组信息时才从课程名中拆出，避免拆坏「形势与政策(含军事理论)」这类名字
    var gm = /^([^（(]+)[（(]\s*(分组[^）)]*)[）)]\s*$/.exec(out.name);
    if (gm) { out.name = trim(gm[1]); if (!out.note) out.note = gm[2]; }
    return out.name ? out : null;
  }

  /* ---------------- 元信息 ---------------- */

  function extractMeta(matrix, headerRow) {
    var meta = {
      title: '', school: '', semester: '', className: '', major: '', department: '', rawLines: []
    };
    var lines = [];
    for (var r = 0; r < headerRow; r++) {
      var row = matrix[r] || [];
      var texts = row.map(trim).filter(Boolean);
      if (!texts.length) continue;
      // 合并单元格可能把同一段文字重复填进多列，去重后再拼接
      var seen = {}, uniq = [];
      texts.forEach(function (t) {
        if (!seen[t]) { seen[t] = 1; uniq.push(t); }
      });
      var line = uniq.join('  ');
      lines.push(line);
    }
    meta.rawLines = lines;
    if (lines.length) {
      meta.title = lines[0];
      var first = lines[0].split(/\s+/)[0];
      if (/(学院|大学|学校|中学)/.test(first)) meta.school = first;
    }
    var all = lines.join('\n');
    function pick(re) {
      var m = re.exec(all);
      return m ? trim(m[1]) : '';
    }
    meta.semester = pick(/学年学期[：:]\s*([^\s]+)/) || pick(/学期[：:]\s*([^\s]+)/);
    meta.className = pick(/班级[：:]\s*([^\s]+)/);
    meta.major = pick(/专业[：:]\s*([^\s]+)/);
    meta.department = pick(/院系[：:]\s*([^\s]+)/) || pick(/系部[：:]\s*([^\s]+)/) || pick(/学院[：:]\s*([^\s]+)/);
    return meta;
  }

  /* ---------------- 列表式课表（一行一门课） ---------------- */

  var LIST_KEYS = [
    { key: 'name', re: /^(课程名称|课程名|课程|科目名称|科目|教学班名称|教学班|名称)$/ },
    { key: 'teacher', re: /^(任课教师|授课教师|上课教师|主讲教师|主讲老师|教师姓名|任课老师|教师|老师)$/ },
    { key: 'weekday', re: /^(星期|星期几|周几|上课星期|星期\/日期|星期[\(（]日期[\)）]|日期)$/ },
    { key: 'period', re: /^(节次|上课节次|节次范围|上课节数|课节|节)$/ },
    { key: 'weeks', re: /^(周次|上课周次|授课周次|教学周|起止周|周数)$/ },
    { key: 'location', re: /^(上课地点|授课地点|上课教室|地点|教室|场地)$/ },
    { key: 'time', re: /^(上课时间|上课时段|时间|上课时间地点|时间地点)$/ }
  ];

  function listHeaderRow(matrix) {
    var lim = Math.min(matrix.length, 20);
    for (var r = 0; r < lim; r++) {
      var map = {}, row = matrix[r] || [];
      for (var c = 0; c < row.length; c++) {
        var t = trim(row[c]).replace(/[\s:：]/g, '');
        if (!t) continue;
        for (var k = 0; k < LIST_KEYS.length; k++) {
          if (LIST_KEYS[k].re.test(t) && map[LIST_KEYS[k].key] == null) { map[LIST_KEYS[k].key] = c; break; }
        }
      }
      if (map.name != null && map.period != null && (map.weekday != null || map.time != null)) {
        return { row: r, cols: map };
      }
    }
    return null;
  }

  function weekdayFromText(s, yearHint) {
    var txt = trim(s);
    if (!txt) return 0;
    var d = parseWeekdayCell(txt, yearHint || 0);
    if (d) return d;
    var m = /(?:星期|周|礼拜)\s*([1-7])/.exec(txt);
    if (m) return +m[1];
    if (/^\s*[1-7]\s*$/.test(txt)) return +txt;
    return 0;
  }

  /* 课程名可能带装饰符或课程代码，如「■高等数学」「[123456]大学英语」「B1234-01 高等数学」 */
  function cleanCourseName(name) {
    var s = trim(name);
    s = s.replace(/^[■□◆◇●○★☆▶►▷\s]+/, '');
    s = s.replace(/^[\[\(【]?[A-Za-z]{0,4}\d{3,}[\]\)】]?[\s\-—:：、]*/, '');
    return trim(s) || trim(name);
  }

  function parseListSheet(matrix) {
    var head = listHeaderRow(matrix);
    if (!head) return null;
    var slots = [], slotMap = {}, courses = [], warnings = [], counter = 0;
    var yearHint = semesterYear(matrix);

    function cellOf(row, key) {
      var c = head.cols[key];
      return c == null ? '' : trim(row[c]);
    }

    for (var r = head.row + 1; r < matrix.length; r++) {
      var row = matrix[r] || [];
      var name = cellOf(row, 'name');
      if (!name) continue;
      if (/^(备注|说明|合计|总计)/.test(name)) break;
      name = cleanCourseName(name);

      var timeText = cellOf(row, 'time');
      var weekday = weekdayFromText(cellOf(row, 'weekday'), yearHint) || weekdayFromText(timeText, yearHint);
      var periods = parsePeriods(cellOf(row, 'period'));
      if (!periods.length) periods = parsePeriods(timeText);
      var weeksText = cellOf(row, 'weeks');
      var weeks = parseWeeks(weeksText);
      if (!weeks.length && timeText) weeks = parseWeeks(timeText);
      if (!weeks.length) weeksText = '';
      var ranges = parseTimeRanges(timeText);
      if (!ranges.length) ranges = parseTimeRanges(cellOf(row, 'period'));
      var label = cellOf(row, 'period');
      if (!label && timeText) label = trim(timeText.split('\n')[0]).slice(0, 24);

      if (!weekday) {
        warnings.push({ code: 'NO_WEEKDAY', message: '「' + name + '」没有识别出星期，已跳过', detail: row.join(' | ') });
        continue;
      }
      if (!periods.length) {
        warnings.push({ code: 'NO_PERIODS', message: '「' + name + '」没有识别出节次，已跳过', detail: row.join(' | ') });
        continue;
      }

      var slotKey = periods.join('-') + '@' + label;
      var si = slotMap[slotKey];
      if (si == null) {
        si = slots.length;
        slotMap[slotKey] = si;
        slots.push({
          index: si,
          label: label || ('第' + periods[0] + '-' + periods[periods.length - 1] + '节'),
          periods: periods.slice(),
          start: ranges.length ? ranges[0].start : null,
          end: ranges.length ? ranges[ranges.length - 1].end : null,
          raw: label
        });
      } else if (ranges.length && !slots[si].start) {
        slots[si].start = ranges[0].start;
        slots[si].end = ranges[ranges.length - 1].end;
      }

      counter++;
      var course = {
        id: 'c' + counter,
        name: name,
        teacher: cellOf(row, 'teacher'),
        location: cellOf(row, 'location'),
        weekday: weekday,
        slotIndex: si,
        slotEndIndex: si,
        periodLabel: slots[si].label,
        weeks: weeks,
        weeksText: weeksText,
        note: '',
        raw: row.join(' | ')
      };
      courses.push(course);
      if (!weeks.length) {
        warnings.push({ code: 'NO_WEEKS', message: '「' + name + '」没有识别出周次，请在核对页补充。', detail: course.raw });
      }
    }

    if (!courses.length) return null;
    return { head: head, slots: slots, courses: courses, warnings: warnings };
  }

  /* ---------------- 主解析 ---------------- */

  function noTimetableResult(sheetName, sheetNames, warnings) {
    return {
      sheetName: sheetName || '',
      sheetNames: sheetNames || [],
      meta: { title: '', school: '', semester: '', className: '', major: '', department: '', rawLines: [] },
      slots: [],
      courses: [],
      warnings: warnings || [{ code: 'NO_TIMETABLE', message: '没有识别出课表：找不到星期表头或节次列，可能是不支持的表格结构。可以点「手动创建课表」逐门录入，或把文件反馈给我们适配。' }],
      detected: { headerRow: -1, weekdayCols: {}, periodCol: -1, slotsFound: 0, confidence: 0 },
      stats: { courseCount: 0, slotCount: 0 }
    };
  }

  /* 周次作行标签：行 = 第X周，列 = 星期（每格是该周要上的课） */
  function detectWeekRowColumn(matrix) {
    var maxC = Math.min(matrix.reduce(function (m, row) { return Math.max(m, row ? row.length : 0); }, 0), 20);
    var best = null;
    for (var c = 0; c < maxC; c++) {
      var hits = 0;
      for (var r = 0; r < matrix.length && r < 60; r++) {
        var txt = trim(cellAt(matrix, r, c));
        if (!/周/.test(txt)) continue;
        if (parseWeeks(txt).length === 1) hits++;
      }
      if (hits >= 3 && (!best || hits > best.hits)) best = { col: c, hits: hits };
    }
    return best;
  }

  function buildWeekRowResult(matrix, header, weekCol, options) {
    var slots = [], slotMap = {}, courses = [], warnings = [], counter = 0;
    var weekRowStart = -1;
    for (var r = header.row + 1; r < matrix.length; r++) {
      if (parseWeeks(cellAt(matrix, r, weekCol)).length === 1) { weekRowStart = r; break; }
    }
    if (weekRowStart < 0) return null;

    for (var r2 = weekRowStart; r2 < matrix.length; r2++) {
      var weekNum = parseWeeks(cellAt(matrix, r2, weekCol));
      if (weekNum.length !== 1) continue;
      for (var d = 1; d <= 7; d++) {
        var col = header.weekdayCols[d];
        if (col == null) continue;
        var text = normalizeText(cellAt(matrix, r2, col));
        if (!trim(text)) continue;
        parseCellBlocks(text).forEach(function (block) {
          var parsed = parseCourseBlock(block);
          if (!parsed) return;
          var periods = parsePeriods(block);
          var label = '';
          var lines = String(block).split('\n');
          for (var i = 0; i < lines.length; i++) {
            if (parsePeriods(lines[i]).length) { label = trim(lines[i]).replace(/^\[|\]$/g, ''); break; }
          }
          if (!periods.length) {
            warnings.push({ code: 'NO_PERIODS', message: '「' + parsed.name + '」第' + weekNum[0] + '周没有识别出节次，已跳过', detail: block });
            return;
          }
          var slotKey = periods.join('-') + '@' + label;
          var si = slotMap[slotKey];
          if (si == null) {
            si = slots.length;
            slotMap[slotKey] = si;
            slots.push({
              index: si,
              label: label || ('第' + periods[0] + '-' + periods[periods.length - 1] + '节'),
              periods: periods.slice(),
              start: null,
              end: null,
              raw: label
            });
          }
          counter++;
          courses.push({
            id: 'c' + counter,
            name: parsed.name,
            teacher: parsed.teacher,
            location: parsed.location,
            weekday: d,
            slotIndex: si,
            slotEndIndex: si,
            periodLabel: slots[si].label,
            weeks: [weekNum[0]],
            weeksText: '第' + weekNum[0] + '周',
            note: parsed.note,
            raw: block
          });
        });
      }
    }
    if (!courses.length) return null;
    slots.sort(function (a, b) { return a.periods[0] - b.periods[0]; });
    var remap = {};
    slots.forEach(function (s, i) { remap[s.index] = i; s.index = i; });
    courses.forEach(function (c) { c.slotIndex = remap[c.slotIndex]; c.slotEndIndex = c.slotIndex; });
    return {
      sheetName: options.sheetName || '',
      sheetNames: options.sheetNames || [],
      meta: extractMeta(matrix, header.row),
      slots: slots,
      courses: courses,
      warnings: warnings,
      detected: { headerRow: header.row, weekdayCols: header.weekdayCols, periodCol: weekCol, slotsFound: slots.length, confidence: 0.6 },
      stats: { courseCount: courses.length, slotCount: slots.length }
    };
  }

  /* 暴力尝试「表头行 × 节次列」组合，抢救非常规布局 */
  function tryGridCombos(matrix, merges, yearHint) {
    var best = null;
    var maxC = Math.min(matrix.reduce(function (m, row) { return Math.max(m, row ? row.length : 0); }, 0), 30);
    var headerRows = [], slotCols = [];
    for (var r = 0; r < matrix.length && r < 40; r++) {
      var map = {}, count = 0, row = matrix[r] || [];
      for (var c = 0; c < row.length; c++) {
        var d = parseWeekdayCell(row[c], yearHint);
        if (d && !map[d]) { map[d] = c; count++; }
      }
      if (count >= 2) headerRows.push({ row: r, weekdayCols: map, count: count });
    }
    for (var cc = 0; cc < maxC; cc++) {
      var hits = 0;
      for (var rr = 0; rr < matrix.length && rr < 80; rr++) {
        if (isSlotLabel(cellAt(matrix, rr, cc))) hits++;
      }
      if (hits >= 2) slotCols.push({ col: cc, hits: hits });
    }
    headerRows.slice(0, 6).forEach(function (h) {
      var dayCols = Object.keys(h.weekdayCols).map(function (k) { return h.weekdayCols[k]; });
      slotCols.slice(0, 6).forEach(function (sc) {
        if (dayCols.indexOf(sc.col) >= 0) return;
        var ctx = { kind: 'grid', matrix: matrix, merges: merges, header: h, cols: { periodCol: sc.col, timeCol: -1 } };
        var res;
        try { res = buildResult(ctx, { sheetName: '', sheetNames: [] }); } catch (e) { return; }
        if (!res || !res.courses.length) return;
        var withWeeks = res.courses.filter(function (c) { return c.weeks.length; }).length;
        var score = res.courses.length * 2 + withWeeks * 2 + res.slots.length;
        if (!best || score > best.score) best = { score: score, result: res };
      });
    });
    return best;
  }

  function analyzeSheet(ws) {
    var sm = sheetMatrix(ws);
    var matrix = sm.matrix, merges = sm.merges;
    var yearHint = semesterYear(matrix);
    var header = detectHeader(matrix, yearHint);
    if (!header) {
      // 竖排课表：星期在行、节次在列
      var th = detectTransposedHeader(matrix, yearHint);
      if (th) {
        var t = transposeSheet(matrix, merges);
        var header2 = detectHeader(t.matrix, yearHint);
        if (header2 && header2.count >= 2) {
          var cols2 = detectSlotColumn(t.matrix, header2, t.merges);
          if (cols2.periodCol >= 0) {
            return {
              ok: true, kind: 'transposed', matrix: t.matrix, merges: t.merges,
              header: header2, cols: cols2,
              meta: extractMeta(matrix, Math.min(matrix.length, 6))
            };
          }
        }
      }
      // 列表式课表：一行一门课
      var list = parseListSheet(matrix);
      if (list) return { ok: true, kind: 'list', list: list, meta: extractMeta(matrix, list.head.row) };
      var cb0 = tryGridCombos(matrix, merges, yearHint);
      if (cb0) return { ok: true, kind: 'result', result: cb0.result };
      return { ok: false, reason: 'no-header', matrix: matrix, merges: merges, header: null };
    }
    var cols = detectSlotColumn(matrix, header, merges);
    if (cols.periodCol < 0) {
      // 周次作行标签的课表（行=第X周）
      var wk = detectWeekRowColumn(matrix);
      if (wk) {
        var wr = buildWeekRowResult(matrix, header, wk.col, {});
        if (wr) return { ok: true, kind: 'result', result: wr };
      }
      var list2 = parseListSheet(matrix);
      if (list2) return { ok: true, kind: 'list', list: list2, meta: extractMeta(matrix, list2.head.row) };
      var cb1 = tryGridCombos(matrix, merges, yearHint);
      if (cb1) return { ok: true, kind: 'result', result: cb1.result };
      return { ok: false, reason: 'no-slot-column', matrix: matrix, merges: merges, header: header, cols: cols };
    }
    return { ok: true, kind: 'grid', matrix: matrix, merges: merges, header: header, cols: cols };
  }

  function buildResult(ctx, options) {
    if (ctx.kind === 'list') {
      var l = ctx.list;
      return {
        sheetName: options.sheetName || '',
        sheetNames: options.sheetNames || [],
        meta: ctx.meta || { title: '', school: '', semester: '', className: '', major: '', department: '', rawLines: [] },
        slots: l.slots,
        courses: l.courses,
        warnings: l.warnings,
        detected: { headerRow: l.head.row, weekdayCols: {}, periodCol: -1, slotsFound: l.slots.length, confidence: 0.9 },
        stats: { courseCount: l.courses.length, slotCount: l.slots.length }
      };
    }
    var matrix = ctx.matrix, merges = ctx.merges, header = ctx.header, cols = ctx.cols;
    var maxC = matrix.reduce(function (m, row) { return Math.max(m, row ? row.length : 0); }, 0);
    var slots = [];
    var slotRows = [];
    var blankStreak = 0;
    for (var r = header.row + 1; r < matrix.length; r++) {
      var dayHasText = false;
      for (var d = 1; d <= 7; d++) {
        var cc = header.weekdayCols[d];
        if (cc == null) continue;
        if (trim(cellAt(matrix, r, cc))) { dayHasText = true; break; }
      }
      var label = cellAt(matrix, r, cols.periodCol);
      var slotLike = isSlotLabel(label);
      if (!slotLike && !dayHasText) {
        blankStreak++;
        if (blankStreak >= 2) break;
        continue;
      }
      blankStreak = 0;
      if (!slotLike) continue;
      var timeText = cols.timeCol >= 0 ? cellAt(matrix, r, cols.timeCol) : '';
      var slot = slotFromLabel(label, timeText);
      slot.index = slots.length;
      var ranges = parseTimeRanges(timeText || label);
      if (!ranges.length && cols.timeCol >= 0) ranges = parseTimeRanges(cellAt(matrix, r, cols.timeCol));
      if (ranges.length) { slot.start = ranges[0].start; slot.end = ranges[ranges.length - 1].end; }
      slots.push(slot);
      slotRows.push(r);
    }
    if (!slots.length) return noTimetableResult('', [], [{ code: 'NO_TIMETABLE', message: '识别到星期表头，但没有找到节次行。' }]);

    // 行号 -> slot 下标
    function slotIndexForRow(row) {
      var idx = -1;
      for (var i = 0; i < slotRows.length; i++) {
        if (slotRows[i] <= row) idx = i; else break;
      }
      return idx;
    }

    var courses = [];
    var warnings = [];
    var counter = 0;

    for (var si = 0; si < slotRows.length; si++) {
      var row = slotRows[si];
      for (var d2 = 1; d2 <= 7; d2++) {
        var col = header.weekdayCols[d2];
        if (col == null || col >= maxC) continue;
        var val = cellAt(matrix, row, col);
        var text = normalizeText(val);
        if (!trim(text)) continue;
        // 合并单元格：只处理左上角，结束行取合并末行
        var mg = findMerge(merges, row, col);
        var endRow = row;
        if (mg) {
          if (mg.s.r !== row || mg.s.c !== col) continue;
          endRow = mg.e.r;
        }
        var endSlot = slotIndexForRow(endRow);
        if (endSlot < si) endSlot = si;
        var blocks = parseCellBlocks(text);
        blocks.forEach(function (block) {
          var parsed = parseCourseBlock(block);
          if (!parsed) return;
          counter++;
          var course = {
            id: 'c' + counter,
            name: parsed.name,
            teacher: parsed.teacher,
            location: parsed.location,
            weekday: d2,
            slotIndex: si,
            slotEndIndex: endSlot,
            periodLabel: slots[si] ? slots[si].label : '',
            weeks: parsed.weeks,
            weeksText: parsed.weeksText,
            note: parsed.note,
            raw: parsed.raw
          };
          courses.push(course);
          if (!course.weeks.length) {
            warnings.push({
              code: 'NO_WEEKS',
              message: '「' + course.name + '」没有识别出周次（' + ['', '周一', '周二', '周三', '周四', '周五', '周六', '周日'][d2] + ' ' + (course.periodLabel || '') + '），请在核对页补充。',
              detail: block
            });
          }
        });
      }
    }

    if (!courses.length) {
      warnings.push({ code: 'NO_COURSES', message: '识别到表头和节次，但单元格里没有课程内容。' });
    }

    var meta = ctx.meta || extractMeta(matrix, header.row);
    var confidence = Math.min(1, 0.55 * Math.min(1, header.count / 5) + 0.45 * Math.min(1, slots.length / 6));

    return {
      sheetName: options.sheetName || '',
      sheetNames: options.sheetNames || [],
      meta: meta,
      slots: slots,
      courses: courses,
      warnings: warnings,
      detected: {
        headerRow: header.row,
        weekdayCols: header.weekdayCols,
        periodCol: cols.periodCol,
        slotsFound: slots.length,
        confidence: Math.round(confidence * 100) / 100
      },
      stats: { courseCount: courses.length, slotCount: slots.length }
    };
  }

  function sheetScore(ws) {
    var analysis;
    try { analysis = analyzeSheet(ws); } catch (e) { return { score: -1, isTimetable: false, reason: 'error', headerRow: -1, weekdayCount: 0 }; }
    if (analysis.kind === 'result') {
      var rr = analysis.result;
      return { score: 40 + Math.min(rr.courses.length, 15), isTimetable: rr.courses.length > 0, reason: 'recovered', headerRow: rr.detected ? rr.detected.headerRow : -1, weekdayCount: 7 };
    }
    if (!analysis.ok || !analysis.header) {
      return { score: 0, isTimetable: false, reason: analysis.reason || (analysis.kind === 'list' ? 'list' : 'no-header'), headerRow: analysis.header ? analysis.header.row : -1, weekdayCount: analysis.header ? analysis.header.count : 0 };
    }
    if (analysis.kind === 'list') {
      var ln = analysis.list.courses.length;
      return { score: 40 + Math.min(ln, 15), isTimetable: true, reason: 'list', headerRow: analysis.list.head.row, weekdayCount: 7 };
    }
    var header = analysis.header;
    var slots = 0;
    for (var r = header.row + 1; r < analysis.matrix.length && r <= header.row + 60; r++) {
      if (isSlotLabel(cellAt(analysis.matrix, r, analysis.cols.periodCol))) slots++;
    }
    var score = header.count * 10 + Math.min(slots, 12) + (analysis.kind === 'transposed' ? 5 : 0);
    return {
      score: score,
      isTimetable: header.count >= 2 && slots >= 1,
      reason: header.count >= 2 && slots >= 1 ? 'ok' : 'weak',
      headerRow: header.row,
      weekdayCount: header.count
    };
  }

  function listSheets(input, options) {
    var wb = readWorkbook(input);
    var names = wb.SheetNames || [];
    var out = names.map(function (name) {
      var info = sheetScore(wb.Sheets[name]);
      return {
        name: name,
        score: info.score,
        isTimetable: info.isTimetable,
        headerRow: info.headerRow,
        weekdayCount: info.weekdayCount,
        reason: info.reason
      };
    });
    out.sort(function (a, b) { return b.score - a.score; });
    return out;
  }

  function parseWorkbook(input, options) {
    options = options || {};
    var wb = readWorkbook(input);
    var names = wb.SheetNames || [];
    if (!names.length) return noTimetableResult('', [], [{ code: 'NO_TIMETABLE', message: '工作簿里没有工作表。' }]);

    var sheetName = options.sheetName && wb.Sheets[options.sheetName] ? options.sheetName : null;
    if (!sheetName) {
      var best = null;
      names.forEach(function (n) {
        var info = sheetScore(wb.Sheets[n]);
        if (!best || info.score > best.score) best = { name: n, score: info.score };
      });
      sheetName = best ? best.name : names[0];
    }
    var analysis = analyzeSheet(wb.Sheets[sheetName]);
    if (!analysis.ok) {
      var why = analysis.reason === 'no-header'
        ? '没有找到星期表头（星期/周一/English weekday/日期均可），也不像列表式课表。'
        : '没有找到节次列（如“0102节”“第1-2节”“第一大节”）。';
      return noTimetableResult(sheetName, names, [{ code: 'NO_TIMETABLE', message: why, detail: analysis.reason }]);
    }
    var result;
    if (analysis.kind === 'result') {
      result = analysis.result;
    } else {
      result = buildResult(analysis, { sheetName: sheetName, sheetNames: names });
      if (!result.courses.length && analysis.matrix) {
        var cb = tryGridCombos(analysis.matrix, analysis.merges, semesterYear(analysis.matrix));
        if (cb && cb.result.courses.length) result = cb.result;
      }
    }
    result.sheetName = sheetName;
    result.sheetNames = names;
    return result;
  }

  function defaultSlots(count) {
    var out = [];
    var n = Math.max(1, Math.min(20, +count || 12));
    for (var i = 0; i < n; i++) {
      out.push({ index: i, label: '第' + (i + 1) + '节', periods: [i + 1], start: null, end: null, raw: '' });
    }
    return out;
  }

  return {
    parseWorkbook: parseWorkbook,
    listSheets: listSheets,
    defaultSlots: defaultSlots,
    parseWeeks: parseWeeks,
    parsePeriods: parsePeriods,
    VERSION: VERSION
  };
});
