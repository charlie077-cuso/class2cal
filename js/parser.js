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

  function parseWeekdayCell(v) {
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
    var m = /(?:星期|周|礼拜)\s*([一二三四五六日天])/.exec(s);
    if (m) return DAY_WORDS[m[1]];
    if (/^[一二三四五六日天]$/.test(s)) return DAY_WORDS[s];
    return 0;
  }

  function cellAt(matrix, r, c) {
    return matrix[r] ? matrix[r][c] : '';
  }

  function detectHeader(matrix) {
    var best = null;
    var maxRow = Math.min(matrix.length, 25);
    for (var r = 0; r < maxRow; r++) {
      var map = {}, count = 0;
      var row = matrix[r] || [];
      for (var c = 0; c < row.length; c++) {
        var day = parseWeekdayCell(row[c]);
        if (day && !map[day]) { map[day] = c; count++; }
      }
      if (count >= 2 && (!best || count > best.count)) {
        best = { row: r, weekdayCols: map, count: count };
      }
    }
    return best;
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
    if (weekIdx > 0) {
      // 教师：周次行之前、且不是分组备注的最后一行
      for (var j = weekIdx - 1; j >= 1; j--) {
        if (/^[（(]/.test(lines[j]) || /分组/.test(lines[j])) continue;
        if (looksLikeTeacher(lines[j]) || (lines[j] && !looksLikeLocation(lines[j]))) { out.teacher = lines[j]; break; }
      }
      // 地点：周次行之后的第一行（或最后一行）
      for (var k = weekIdx + 1; k < lines.length; k++) {
        if (/^[（(]/.test(lines[k])) continue;
        out.location = lines[k];
        break;
      }
      if (!out.location && lines.length > weekIdx + 1) out.location = lines[lines.length - 1];
    } else {
      // 没有周次行：第二行当教师，最后一行若像地点则当地点
      if (lines.length >= 2 && !/^[（(]/.test(lines[1])) out.teacher = lines[1];
      if (lines.length >= 3 && looksLikeLocation(lines[lines.length - 1])) out.location = lines[lines.length - 1];
    }
    out.name = out.name.replace(/^[\s·•]+/, '').trim();
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
    meta.department = pick(/院系[：:]\s*([^\s]+)/) || pick(/系部[：:]\s*([^\s]+)/);
    return meta;
  }

  /* ---------------- 主解析 ---------------- */

  function noTimetableResult(sheetName, sheetNames, warnings) {
    return {
      sheetName: sheetName || '',
      sheetNames: sheetNames || [],
      meta: { title: '', school: '', semester: '', className: '', major: '', department: '', rawLines: [] },
      slots: [],
      courses: [],
      warnings: warnings || [{ code: 'NO_TIMETABLE', message: '没有识别出课表：找不到“星期一…星期日”表头或节次列。' }],
      detected: { headerRow: -1, weekdayCols: {}, periodCol: -1, slotsFound: 0, confidence: 0 },
      stats: { courseCount: 0, slotCount: 0 }
    };
  }

  function analyzeSheet(ws) {
    var sm = sheetMatrix(ws);
    var matrix = sm.matrix, merges = sm.merges;
    var header = detectHeader(matrix);
    if (!header) {
      return { ok: false, reason: 'no-header', matrix: matrix, merges: merges, header: null };
    }
    var cols = detectSlotColumn(matrix, header, merges);
    if (cols.periodCol < 0) {
      return { ok: false, reason: 'no-slot-column', matrix: matrix, merges: merges, header: header, cols: cols };
    }
    return { ok: true, matrix: matrix, merges: merges, header: header, cols: cols };
  }

  function buildResult(ctx, options) {
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

    var meta = extractMeta(matrix, header.row);
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
    if (!analysis.ok || !analysis.header) {
      return { score: 0, isTimetable: false, reason: analysis.reason || 'no-header', headerRow: analysis.header ? analysis.header.row : -1, weekdayCount: analysis.header ? analysis.header.count : 0 };
    }
    var header = analysis.header;
    var slots = 0;
    for (var r = header.row + 1; r < analysis.matrix.length && r <= header.row + 60; r++) {
      if (isSlotLabel(cellAt(analysis.matrix, r, analysis.cols.periodCol))) slots++;
    }
    var score = header.count * 10 + Math.min(slots, 12);
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
        ? '没有找到“星期一…星期日”表头。'
        : '没有找到节次列（如“0102节”“第1-2节”）。';
      return noTimetableResult(sheetName, names, [{ code: 'NO_TIMETABLE', message: why, detail: analysis.reason }]);
    }
    analysis.resultOptions = { sheetName: sheetName, sheetNames: names };
    return buildResult(analysis, { sheetName: sheetName, sheetNames: names });
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
