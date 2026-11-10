/* ============================================================
   课表转日历 · 交互逻辑
   依赖：vendor/xlsx.full.min.js → js/parser.js → js/ics.js
   ============================================================ */
(function () {
  'use strict';

  var TT = window.TT || {};
  var DAYS = ['', '周一', '周二', '周三', '周四', '周五', '周六', '周日'];
  var DAY_FULL = ['', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'];
  var COLORS = ['#4f6bed', '#0d9f9a', '#d9822b', '#a25bd6', '#3a8fd9', '#d95b7a', '#5f9e3f', '#b07d2b'];
  var STORAGE_KEY = 'class2cal.settings.v1';

  var $ = function (sel, el) { return (el || document).querySelector(sel); };
  var $$ = function (sel, el) { return Array.prototype.slice.call((el || document).querySelectorAll(sel)); };

  var state = {
    step: 1,
    file: null,
    fileName: '',
    buffer: null,
    parse: null,
    sheetName: null,
    sheetOptions: [],
    courses: [],
    slots: [],
    editedSlots: {},
    manual: false,
    gridVisible: false,
    editId: null,
    lastIcsUrl: null,
    options: {
      semesterStart: '',
      firstPeriodTime: '08:00',
      lessonMinutes: 45,
      breakMinutes: 10,
      afternoonStart: '14:00',
      timeZone: 'Asia/Shanghai',
      reminderMinutes: 10,
      mode: 'recurring',
      calendarName: '课程表'
    }
  };

  /* ---------------- 工具 ---------------- */

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  function ymd(date) {
    return date.getFullYear() + '-' + pad2(date.getMonth() + 1) + '-' + pad2(date.getDate());
  }

  function parseYmd(s) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
    if (!m) return null;
    var d = new Date(+m[1], +m[2] - 1, +m[3]);
    if (d.getFullYear() !== +m[1] || d.getMonth() !== +m[2] - 1 || d.getDate() !== +m[3]) return null;
    return d;
  }

  function toMinutes(t) {
    var m = /^(\d{1,2}):(\d{2})$/.exec(String(t || ''));
    if (!m) return null;
    return +m[1] * 60 + +m[2];
  }

  function fmtMinutes(min) {
    min = ((min % 1440) + 1440) % 1440;
    return pad2(Math.floor(min / 60)) + ':' + pad2(min % 60);
  }

  function mondayOfWeek(date) {
    var d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    var day = d.getDay() || 7;
    d.setDate(d.getDate() - (day - 1));
    return d;
  }

  function toast(msg, isError) {
    var el = $('#status');
    if (!el) return;
    el.textContent = msg;
    el.classList.toggle('error', !!isError);
    el.hidden = false;
    el.classList.add('show');
    clearTimeout(toast._timer);
    toast._timer = setTimeout(function () {
      el.classList.remove('show');
      setTimeout(function () { el.hidden = true; }, 250);
    }, isError ? 3600 : 2400);
  }

  function courseColor(name) {
    var h = 0;
    name = String(name || '');
    for (var i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0;
    return COLORS[h % COLORS.length];
  }

  /* ---------------- 周次文本解析/格式化 ---------------- */

  function parseWeeksInput(text) {
    var raw = String(text == null ? '' : text).trim();
    if (!raw) return { weeks: [], error: '请填写周次，例如 1-16' };

    var s = raw
      .replace(/[，、；;／\/]/g, ',')
      .replace(/[～~至]/g, '-')
      .replace(/第/g, '')
      .replace(/\s+/g, '');

    var parity = 0; // 0=全部 1=单周 2=双周
    if (/单周|单/.test(s)) parity = 1;
    else if (/双周|双/.test(s)) parity = 2;
    s = s.replace(/(单双周?|单周|双周|单|双)/g, '');

    // 允许 "1-16周" 这类残留的「周」
    s = s.replace(/周/g, '');

    if (!s && parity) {
      s = '1-30';
    }
    if (!s) return { weeks: [], error: '请填写周次，例如 1-16' };

    var tokens = s.split(',').filter(function (x) { return x; });
    var set = {};
    for (var i = 0; i < tokens.length; i++) {
      var m = /^(\d{1,2})(?:-(\d{1,2}))?$/.exec(tokens[i]);
      if (!m) return { weeks: [], error: '无法识别「' + tokens[i] + '」，请用 1-16 或 1,3,5-9 这样的写法' };
      var a = +m[1];
      var b = m[2] == null ? a : +m[2];
      if (a > b) { var t = a; a = b; b = t; }
      if (a < 1 || b > 60) return { weeks: [], error: '周次应在 1-60 之间' };
      for (var w = a; w <= b; w++) set[w] = true;
    }
    var weeks = Object.keys(set).map(Number).sort(function (a, b) { return a - b; });
    if (parity === 1) weeks = weeks.filter(function (w) { return w % 2 === 1; });
    if (parity === 2) weeks = weeks.filter(function (w) { return w % 2 === 0; });
    if (!weeks.length) return { weeks: [], error: '按单双周过滤后没有剩下周次，请检查' };
    return { weeks: weeks, error: '' };
  }

  function formatWeeks(weeks) {
    if (!weeks || !weeks.length) return '';
    var sorted = weeks.slice().sort(function (a, b) { return a - b; });
    var parts = [];
    var start = sorted[0], prev = sorted[0];
    for (var i = 1; i <= sorted.length; i++) {
      var cur = sorted[i];
      if (i < sorted.length && cur === prev + 1) { prev = cur; continue; }
      parts.push(start === prev ? String(start) : start + '-' + prev);
      start = prev = cur;
    }
    return parts.join(',');
  }

  function weeksDisplay(weeks) {
    if (!weeks || !weeks.length) return '周次待补充';
    var text = formatWeeks(weeks);
    var odd = weeks.every(function (x) { return x % 2 === 1; });
    var even = weeks.every(function (x) { return x % 2 === 0; });
    if (odd && weeks.length >= 3) return text + ' 单周';
    if (even && weeks.length >= 2) return text + ' 双周';
    return text + ' 周';
  }

  /* ---------------- 步骤切换 ---------------- */

  function setStep(n) {
    state.step = n;
    $$('.step').forEach(function (sec) {
      sec.classList.toggle('active', +sec.getAttribute('data-step') === n);
    });
    $$('.steps li').forEach(function (li) {
      var s = +li.getAttribute('data-step-dot');
      li.classList.toggle('active', s === n);
      li.classList.toggle('done', s < n);
    });

    var nav = $('#bottom-nav');
    var btnNext = $('#btn-next');
    var btnBack = $('#btn-back');
    if (n === 1) {
      nav.hidden = true;
    } else {
      nav.hidden = false;
      btnNext.hidden = n === 4;
      if (n === 2) { btnNext.textContent = '下一步：校对时间'; btnNext.setAttribute('data-testid', 'next-2'); btnBack.setAttribute('data-testid', 'back-2'); }
      if (n === 3) { btnNext.textContent = '下一步：生成日历'; btnNext.setAttribute('data-testid', 'next-3'); btnBack.setAttribute('data-testid', 'back-3'); }
      if (n === 4) { btnBack.setAttribute('data-testid', 'back-4'); }
      btnBack.textContent = n === 4 ? '返回修改时间' : '上一步';
    }
    $('#btn-restart').hidden = n === 1;
    window.scrollTo({ top: 0, behavior: 'smooth' });
    if (n === 2) renderCourses();
    if (n === 3) prepareCalibration();
    if (n === 4) prepareExport();
  }

  /* ---------------- 文件上传与解析 ---------------- */

  function setBusy(busy) {
    var box = $('#pick-btn');
    var strong = box && box.querySelector('strong');
    if (!strong) return;
    strong.textContent = busy ? '正在解析课表，请稍候…' : '点击选择课表文件';
    box.disabled = !!busy;
    $('#pick-btn').style.opacity = busy ? '.6' : '';
  }

  function showParseError(text) {
    $('#parse-summary').hidden = true;
    $('#parse-error').hidden = false;
    $('#parse-error-text').textContent = text || '这个文件的表格结构比较特殊，暂时无法自动识别。';
  }

  function handleFile(file) {
    if (!file) return;
    var name = file.name || '';
    if (!/\.(xls|xlsx|csv)$/i.test(name)) {
      toast('请选择 .xls、.xlsx 或 .csv 格式的课表文件', true);
      return;
    }
    if (file.size > 30 * 1024 * 1024) {
      toast('文件太大（超过 30MB），请检查是否选错文件', true);
      return;
    }
    if (!TT.parseWorkbook) {
      showParseError('解析模块没有加载成功，请刷新页面后重试。');
      return;
    }

    $('#parse-error').hidden = true;
    setBusy(true);
    state.fileName = name;

    file.arrayBuffer().then(function (buf) {
      state.buffer = buf;
      var sheets = [];
      try {
        sheets = TT.listSheets ? TT.listSheets(buf) : [];
      } catch (e) { sheets = []; }
      state.sheetOptions = sheets || [];
      var best = sheets.filter(function (s) { return s.isTimetable; })[0] || sheets[0];
      var opts = best ? { sheetName: best.name } : {};
      var result;
      try {
        result = TT.parseWorkbook(buf, opts);
      } catch (e) {
        throw new Error('解析失败：' + (e && e.message ? e.message : e));
      }
      if (!result || !result.courses || !result.courses.length) {
        state.parse = result || null;
        state.slots = (result && result.slots) || [];
        var msg = '没有从文件里找到课程。可能原因：\n· 这不是课表文件\n· 课表在多个工作表中，或格式特殊';
        showParseError(msg);
        setBusy(false);
        return;
      }
      applyParse(result, best ? best.name : result.sheetName);
      setBusy(false);
    }).catch(function (e) {
      setBusy(false);
      showParseError('读取文件失败：' + (e && e.message ? e.message : e));
    });
  }

  function applyParse(result, sheetName) {
    state.parse = result;
    state.sheetName = sheetName || result.sheetName || '';
    state.manual = false;
    state.slots = (result.slots || []).map(function (s, i) {
      return {
        index: i,
        label: s.label || ('第' + (i + 1) + '节'),
        periods: s.periods || [],
        start: s.start || null,
        end: s.end || null,
        _fileStart: s.start || null,
        _fileEnd: s.end || null,
        raw: s.raw || ''
      };
    });
    if (!state.slots.length) state.slots = TT.defaultSlots ? TT.defaultSlots(12) : [];
    state.editedSlots = {};

    state.courses = (result.courses || []).map(function (c, i) {
      return {
        id: 'c' + (i + 1) + '_' + Math.random().toString(36).slice(2, 7),
        name: c.name || '',
        teacher: c.teacher || '',
        location: c.location || '',
        weekday: c.weekday || 1,
        slotIndex: typeof c.slotIndex === 'number' ? c.slotIndex : 0,
        slotEndIndex: typeof c.slotEndIndex === 'number' ? c.slotEndIndex : (typeof c.slotIndex === 'number' ? c.slotIndex : 0),
        periodLabel: c.periodLabel || '',
        weeks: Array.isArray(c.weeks) ? c.weeks.slice() : [],
        weeksText: c.weeksText || (Array.isArray(c.weeks) ? formatWeeks(c.weeks) + '周' : ''),
        note: c.note || '',
        raw: c.raw || ''
      };
    });

    // 用文件中的第一条时间作为默认「第 1 节开始时间」
    var firstWithTime = null;
    for (var i = 0; i < state.slots.length; i++) {
      if (state.slots[i].start) { firstWithTime = state.slots[i]; break; }
    }
    if (firstWithTime) state.options.firstPeriodTime = firstWithTime.start;

    // 默认日历名称带上学期
    var sem = result.meta && result.meta.semester;
    if (sem) state.options.calendarName = '课程表 ' + sem;

    renderSummary();
    saveSettings();
  }

  function renderSummary() {
    var r = state.parse;
    if (!r) return;
    var meta = r.meta || {};
    $('#parse-summary').hidden = false;
    $('#parse-title').textContent = meta.title || '课表已解析';
    var parts = [];
    if (state.sheetName) parts.push('工作表：' + state.sheetName);
    if (state.fileName) parts.push(state.fileName);
    $('#parse-subtitle').textContent = parts.join(' · ');

    var badge = $('#parse-badge');
    var warnCount = (r.warnings || []).length;
    badge.textContent = warnCount ? '解析完成，有 ' + warnCount + ' 条提示' : '解析成功';
    badge.className = 'badge ' + (warnCount ? 'badge-warn' : 'badge-ok');

    var pairs = [];
    if (meta.school) pairs.push(['学校', meta.school]);
    if (meta.semester) pairs.push(['学期', meta.semester]);
    if (meta.className) pairs.push(['班级', meta.className]);
    if (meta.major) pairs.push(['专业', meta.major]);
    if (meta.department) pairs.push(['院系', meta.department]);
    var metaHtml = pairs.map(function (p) {
      return '<div><dt>' + esc(p[0]) + '</dt><dd>' + esc(p[1]) + '</dd></div>';
    }).join('');
    $('#parse-meta').innerHTML = metaHtml;
    $('#parse-meta').hidden = !pairs.length;

    $('#course-count').textContent = state.courses.length;
    $('#slot-count').textContent = state.slots.length;
    var maxWeek = 0;
    state.courses.forEach(function (c) {
      (c.weeks || []).forEach(function (w) { if (w > maxWeek) maxWeek = w; });
    });
    $('#week-count').textContent = maxWeek || '—';

    // 工作表选择
    var pickerRow = $('#sheet-picker-row');
    var select = $('#sheet-select');
    if (state.sheetOptions && state.sheetOptions.length > 1) {
      pickerRow.hidden = false;
      select.innerHTML = state.sheetOptions.map(function (s) {
        return '<option value="' + esc(s.name) + '"' + (s.name === state.sheetName ? ' selected' : '') + '>' +
          esc(s.name) + (s.isTimetable ? '' : '（可能不是课表）') + '</option>';
      }).join('');
    } else {
      pickerRow.hidden = true;
    }

    // 警告
    var warnBox = $('#warnings');
    var warnings = (r.warnings || []).slice(0, 6);
    if (warnings.length) {
      warnBox.hidden = false;
      warnBox.innerHTML = warnings.map(function (w) {
        return '<div class="warn-item"><b>提示</b><span>' + esc(w.message || w.code || '') + '</span></div>';
      }).join('');
    } else {
      warnBox.hidden = true;
    }
  }

  function switchSheet(sheetName) {
    if (!state.buffer || !TT.parseWorkbook) return;
    setBusy(true);
    setTimeout(function () {
      try {
        var result = TT.parseWorkbook(state.buffer, { sheetName: sheetName });
        if (!result || !result.courses || !result.courses.length) {
          showParseError('工作表「' + sheetName + '」里没有找到课程，换一个工作表试试。');
        } else {
          $('#parse-error').hidden = true;
          applyParse(result, sheetName);
          toast('已切换到工作表：' + sheetName);
        }
      } catch (e) {
        showParseError('解析工作表失败：' + (e && e.message ? e.message : e));
      }
      setBusy(false);
    }, 30);
  }

  function enterManualMode() {
    state.manual = true;
    state.parse = null;
    state.sheetName = '';
    state.sheetOptions = [];
    state.courses = [];
    state.slots = TT.defaultSlots ? TT.defaultSlots(12) : [];
    state.editedSlots = {};
    $('#parse-error').hidden = true;
    $('#parse-summary').hidden = true;
    $('#review-notice').hidden = false;
    $('#review-notice').className = 'notice notice-info';
    $('#review-notice').textContent = '当前是手动模式：请逐门添加课程。下面的时段表默认 12 节课，可在下一步校对具体时间。';
    setStep(2);
  }

  /* ---------------- 步骤 2：课程核对 ---------------- */

  function renderCourses() {
    var wrap = $('#courses-preview');
    var empty = $('#courses-empty');
    $('#review-count').textContent = state.courses.length + ' 门课程';

    if (!state.courses.length) {
      wrap.innerHTML = '';
      empty.hidden = false;
      $('#grid-wrap').hidden = true;
      return;
    }
    empty.hidden = true;

    var missingWeeks = state.courses.filter(function (c) { return !(c.weeks && c.weeks.length); });
    var notice = $('#review-notice');
    if (!state.manual && missingWeeks.length) {
      notice.hidden = false;
      notice.className = 'notice notice-warn';
      notice.textContent = '有 ' + missingWeeks.length + ' 门课没有识别出周次，请点击课程卡片补充，否则这些课不会被导出。';
    } else if (!state.manual) {
      notice.hidden = true;
    }

    var html = '';
    for (var d = 1; d <= 7; d++) {
      var list = state.courses.filter(function (c) { return c.weekday === d; })
        .sort(function (a, b) { return (a.slotIndex || 0) - (b.slotIndex || 0); });
      if (!list.length) continue;
      html += '<section class="day-group"><h3>' + DAY_FULL[d] + '</h3>';
      list.forEach(function (c) {
        var slotLabel = c.periodLabel || (state.slots[c.slotIndex] && state.slots[c.slotIndex].label) || '';
        var time = '';
        var s = state.slots[c.slotIndex];
        var e = state.slots[c.slotEndIndex || c.slotIndex];
        if (s && s.start) time = s.start + (e && e.end ? '-' + e.end : '');
        var meta = [];
        if (c.teacher) meta.push('<span>教师：' + esc(c.teacher) + '</span>');
        if (time) meta.push('<span>' + esc(time) + '</span>');
        meta.push('<span>周次：' + esc(weeksDisplay(c.weeks)) + '</span>');
        if (c.location) meta.push('<span class="loc">' + esc(c.location) + '</span>');
        var warn = !(c.weeks && c.weeks.length) ? '<span class="tag-warn">缺周次</span>' : '';
        html += '<article class="course" data-id="' + esc(c.id) + '" style="--course-color:' + courseColor(c.name) + '">' +
          '<div class="course-body" data-action="edit">' +
            '<div class="course-title"><h4>' + esc(c.name || '（未命名课程）') + '</h4>' +
              (slotLabel ? '<span class="slot-tag">' + esc(slotLabel) + '</span>' : '') + warn + '</div>' +
            '<div class="course-meta">' + meta.join('') + '</div>' +
            (c.note ? '<div class="course-note">' + esc(c.note) + '</div>' : '') +
          '</div>' +
          '<div class="course-actions">' +
            '<button type="button" class="icon-btn" data-action="edit" aria-label="编辑">' +
              '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>' +
            '</button>' +
            '<button type="button" class="icon-btn del" data-action="del" aria-label="删除">' +
              '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14"/></svg>' +
            '</button>' +
          '</div>' +
        '</article>';
      });
      html += '</section>';
    }
    wrap.innerHTML = html;
    if (state.gridVisible) renderGrid(); else $('#grid-wrap').hidden = true;
  }

  function renderGrid() {
    var wrap = $('#grid-wrap');
    var table = $('#timetable-grid');
    if (!state.courses.length) { wrap.hidden = true; return; }
    wrap.hidden = false;

    var html = '<thead><tr><th>节次</th>';
    for (var d = 1; d <= 7; d++) html += '<th>' + DAYS[d] + '</th>';
    html += '</tr></thead><tbody>';
    state.slots.forEach(function (slot, i) {
      html += '<tr><td class="slot-label">' + esc(slot.label) + '</td>';
      for (var day = 1; day <= 7; day++) {
        var list = state.courses.filter(function (c) { return c.weekday === day && c.slotIndex === i; });
        html += '<td>' + list.map(function (c) {
          var span = (c.slotEndIndex && c.slotEndIndex > c.slotIndex)
            ? '（' + ((state.slots[c.slotIndex] || {}).label || '') + '–' + ((state.slots[c.slotEndIndex] || {}).label || '') + '）'
            : '';
          return '<div class="tt-chip" style="background:' + courseColor(c.name) + '1a;color:var(--ink-strong)">' +
            '<b>' + esc(c.name) + '</b><span>' + esc(c.location || '') + span + '</span></div>';
        }).join('') + '</td>';
      }
      html += '</tr>';
    });
    html += '</tbody>';
    table.innerHTML = html;
  }

  /* ---------------- 课程编辑对话框 ---------------- */

  function fillSlotSelects() {
    var startSel = $('#edit-slot-start');
    var endSel = $('#edit-slot-end');
    var opts = state.slots.map(function (s, i) {
      var label = s.label || ('第' + (i + 1) + '节');
      if (s.start) label += '（' + s.start + (s.end ? '-' + s.end : '') + '）';
      return '<option value="' + i + '">' + esc(label) + '</option>';
    }).join('');
    startSel.innerHTML = opts;
    endSel.innerHTML = opts;
  }

  function openEditDialog(course) {
    fillSlotSelects();
    state.editId = course ? course.id : null;
    $('#edit-title').textContent = course ? '编辑课程' : '添加课程';
    $('#edit-name').value = course ? course.name : '';
    $('#edit-weekday').value = course ? course.weekday : 1;
    $('#edit-slot-start').value = course ? course.slotIndex : 0;
    $('#edit-slot-end').value = course ? (course.slotEndIndex == null ? course.slotIndex : course.slotEndIndex) : 0;
    $('#edit-teacher').value = course ? course.teacher : '';
    $('#edit-location').value = course ? course.location : '';
    $('#edit-note').value = course ? course.note : '';
    $('#edit-weeks').value = course ? (course.weeks && course.weeks.length ? formatWeeks(course.weeks) : '') : '1-16';
    $('#edit-error').hidden = true;
    $('#edit-delete').hidden = !course;
    var dlg = $('#edit-dialog');
    if (!dlg.open) dlg.showModal();
    setTimeout(function () { $('#edit-name').focus(); }, 50);
  }

  function saveEditedCourse() {
    var name = $('#edit-name').value.trim();
    var weeksRes = parseWeeksInput($('#edit-weeks').value);
    var err = $('#edit-error');
    if (!name) { err.textContent = '请填写课程名称'; err.hidden = false; return; }
    if (weeksRes.error) { err.textContent = weeksRes.error; err.hidden = false; return; }
    var startIdx = +$('#edit-slot-start').value || 0;
    var endIdx = +$('#edit-slot-end').value || 0;
    if (endIdx < startIdx) endIdx = startIdx;
    var data = {
      name: name,
      teacher: $('#edit-teacher').value.trim(),
      location: $('#edit-location').value.trim(),
      weekday: +$('#edit-weekday').value || 1,
      slotIndex: startIdx,
      slotEndIndex: endIdx,
      periodLabel: (state.slots[startIdx] && state.slots[startIdx].label) || '',
      weeks: weeksRes.weeks,
      weeksText: $('#edit-weeks').value.trim(),
      note: $('#edit-note').value.trim()
    };
    if (state.editId) {
      for (var i = 0; i < state.courses.length; i++) {
        if (state.courses[i].id === state.editId) {
          Object.assign(state.courses[i], data);
          break;
        }
      }
    } else {
      data.id = 'c' + Date.now() + '_' + Math.random().toString(36).slice(2, 6);
      data.raw = '';
      state.courses.push(data);
    }
    $('#edit-dialog').close();
    renderCourses();
    toast(state.editId ? '课程已更新' : '课程已添加');
  }

  function deleteCourse(id) {
    var c = null;
    state.courses = state.courses.filter(function (x) { if (x.id === id) c = x; return x.id !== id; });
    if (c) toast('已删除「' + (c.name || '课程') + '」');
    renderCourses();
  }

  /* ---------------- 步骤 3：校对时间 ---------------- */

  function updatePresetUI() {
    var sel = $('#schedule-preset');
    if (!sel) return;
    var key = (+state.options.lessonMinutes || 45) + '-' + (+state.options.breakMinutes || 10);
    var match = false;
    for (var i = 0; i < sel.options.length; i++) {
      if (sel.options[i].value === key) { sel.value = key; match = true; }
    }
    if (!match) sel.value = 'custom';
    $('#custom-time-fields').hidden = sel.value !== 'custom';
  }

  function syncCalibrationInputs() {
    var o = state.options;
    $('#semester-start').value = o.semesterStart || '';
    $('#first-period-time').value = o.firstPeriodTime || '08:00';
    $('#afternoon-start').value = o.afternoonStart || '14:00';
    $('#lesson-minutes').value = String(o.lessonMinutes || 45);
    $('#break-minutes').value = String(o.breakMinutes || 10);
    updatePresetUI();
  }

  function slotPeriodText(slot) {
    var periods = slot.periods || [];
    if (!periods.length) return '';
    return '第 ' + formatPeriods(periods) + ' 节';
  }

  function renderPreview() {
    var list = $('#slot-preview');
    if (!list) return;
    list.innerHTML = state.slots.map(function (slot, i) {
      var time = slot.start && slot.end
        ? esc(slot.start) + ' – ' + esc(slot.end)
        : '待填写';
      return '<li data-preview-slot="' + i + '">' +
        '<span class="slot-preview-label">' + esc(slot.label || ('第 ' + (i + 1) + ' 节')) + '</span>' +
        '<span class="slot-preview-periods">' + esc(slotPeriodText(slot)) + '</span>' +
        '<span class="slot-preview-time" data-testid="preview-time-' + i + '">' + time + '</span>' +
      '</li>';
    }).join('');
  }

  function renderSlotTable() {
    var body = $('#slot-times');
    var rows = state.slots.map(function (slot, i) {
      var periods = slot.periods || [];
      var periodText = periods.length ? '第 ' + formatPeriods(periods) + ' 节' : '';
      var detected = '';
      if (slot._fileStart) {
        detected = '<span class="detected-time">' + esc(slot._fileStart) + (slot._fileEnd ? '-' + esc(slot._fileEnd) : '') + '</span>';
      } else if (slot.start) {
        detected = '<span class="detected-time">' + esc(slot.start) + (slot.end ? '-' + esc(slot.end) : '') + '</span>';
      } else {
        detected = '<span class="no-detect">—</span>';
      }
      return '<tr data-slot="' + i + '">' +
        '<td class="slot-label"><b>' + esc(slot.label) + '</b>' + (periodText ? '<small>' + esc(periodText) + '</small>' : '') + '</td>' +
        '<td><input type="time" data-slot-start="' + i + '" data-testid="slot-time-' + i + '" value="' + esc(slot.start || '') + '"></td>' +
        '<td><input type="time" data-slot-end="' + i + '" data-testid="slot-time-' + i + '-end" value="' + esc(slot.end || '') + '"></td>' +
        '<td class="col-detect">' + detected + '</td>' +
      '</tr>';
    }).join('');
    body.innerHTML = rows;
    renderPreview();

    $('#slot-source-hint').textContent = state.slots.some(function (s) { return s._fileStart; })
      ? '文件里带有时间，已优先填入；可展开高级逐节核对。'
      : '已按上面的设置自动推算，可展开高级逐节调整。';
  }

  function formatPeriods(periods) {
    if (!periods.length) return '';
    var sorted = periods.slice().sort(function (a, b) { return a - b; });
    if (sorted.length === 1) return String(sorted[0]);
    var consecutive = true;
    for (var i = 1; i < sorted.length; i++) if (sorted[i] !== sorted[i - 1] + 1) consecutive = false;
    return consecutive ? sorted[0] + '-' + sorted[sorted.length - 1] : sorted.join(',');
  }

  function computeSlotTimes(preferFile) {
    var lesson = +state.options.lessonMinutes || 45;
    var gap = +state.options.breakMinutes || 10;
    var afternoon = toMinutes(state.options.afternoonStart);
    var cursor = toMinutes(state.options.firstPeriodTime);
    if (cursor == null) cursor = 8 * 60;
    var prevLast = 0;

    state.slots.forEach(function (slot, i) {
      var periods = slot.periods || [];
      var count = Math.max(1, periods.length || 1);
      var firstPeriod = periods.length ? periods[0] : (i * 2 + 1);

      // 上午 → 下午：从用户填写的「下午第 1 节开始」继续；没填则午休 2 小时兜底
      if (firstPeriod >= 5 && prevLast > 0 && prevLast <= 4) {
        if (afternoon != null) cursor = Math.max(cursor, afternoon);
        else cursor += 120;
      }

      var startMin = cursor;
      var endMin = startMin + count * lesson + (count - 1) * gap;

      if (preferFile && slot._fileStart) {
        var fs = toMinutes(slot._fileStart);
        if (fs != null) {
          startMin = fs;
          var fe = slot._fileEnd ? toMinutes(slot._fileEnd) : null;
          endMin = fe != null ? fe : startMin + count * lesson + (count - 1) * gap;
        }
      }

      cursor = endMin + gap;
      prevLast = periods.length ? periods[periods.length - 1] : firstPeriod;

      if (!state.editedSlots[i]) {
        slot.start = fmtMinutes(startMin);
        slot.end = fmtMinutes(endMin);
      }
    });
  }

  function prepareCalibration() {
    // 学期开始日期必须由用户填写：不预填，避免导出到错误的学期
    syncCalibrationInputs();
    // 首次进入：优先使用文件时间；若全部无文件时间则按答案推算
    computeSlotTimes(!state._calibrated);
    state._calibrated = true;
    renderSlotTable();
  }

  function validateCalibration() {
    var dateStr = $('#semester-start').value;
    var d = parseYmd(dateStr);
    if (!d) {
      $('#semester-start-error').textContent = '请选择学期第 1 周的周一日期';
      $('#semester-start-error').hidden = false;
      toast('请先选择学期第 1 周的周一日期', true);
      return false;
    }
    if (d.getDay() !== 1) {
      var monday = mondayOfWeek(d);
      $('#semester-start').value = ymd(monday);
      state.options.semesterStart = ymd(monday);
      toast('已自动调整为该周的周一：' + ymd(monday));
    }
    $('#semester-start-error').hidden = true;

    var bad = null;
    state.slots.forEach(function (slot, i) {
      if (bad != null) return;
      var s = toMinutes(slot.start), e = toMinutes(slot.end);
      if (s == null || e == null) bad = { i: i, why: '时间没有填完整' };
      else if (e <= s) bad = { i: i, why: '结束时间需要晚于开始时间' };
    });
    if (bad) {
      var advanced = $('#advanced-slots');
      if (advanced) advanced.open = true;
      toast('时段「' + (state.slots[bad.i].label || ('第' + (bad.i + 1) + '节')) + '」' + bad.why, true);
      var row = $('[data-slot="' + bad.i + '"]');
      if (row) row.scrollIntoView({ behavior: 'smooth', block: 'center' });
      return false;
    }
    return true;
  }

  /* ---------------- 步骤 4：导出 ---------------- */

  function buildSettings() {
    return {
      semesterStart: state.options.semesterStart,
      timeZone: state.options.timeZone,
      calendarName: state.options.calendarName || '课程表',
      reminderMinutes: +state.options.reminderMinutes || 0,
      mode: state.options.mode,
      now: new Date()
    };
  }

  function previewIcs() {
    if (!TT.buildICS) return null;
    try {
      return TT.buildICS({ courses: state.courses, slots: state.slots, settings: buildSettings() });
    } catch (e) {
      return { error: e && e.message ? e.message : String(e) };
    }
  }

  function prepareExport() {
    var res = previewIcs();
    var summary = $('#export-summary');
    if (!res || res.error) {
      summary.innerHTML = '<span class="bad">生成失败：' + esc(res && res.error ? res.error : 'ICS 模块未加载') + '</span>';
      return;
    }
    var warnHtml = '';
    var warns = (res.warnings || []).slice(0, 3);
    if (warns.length) {
      warnHtml = '<div class="notice notice-warn" style="text-align:left;margin-top:10px">' +
        warns.map(function (w) { return esc(w.message || w.code || ''); }).join('<br>') +
        (res.warnings.length > 3 ? '<br>…' : '') + '</div>';
    }
    summary.innerHTML = '将生成 <b>' + res.eventCount + '</b> 个日历事件<br>' +
      '共 ' + state.courses.length + ' 门课程 · ' + state.slots.length + ' 个时段 · ' +
      esc(state.options.calendarName || '课程表') + warnHtml;
    $('#export-hint').textContent = state.options.mode === 'recurring'
      ? '每周重复模式：同一门课在整个学期里是一条重复事件，日历更简洁。'
      : '单独事件模式：每次课各占一条，适合导入到不支持重复事件的旧日历。';
  }

  function downloadIcs() {
    var res = previewIcs();
    if (!res || res.error) { toast('生成失败：' + (res && res.error ? res.error : '未知错误'), true); return; }
    if (!res.eventCount) { toast('没有可导出的事件，请先核对课程周次与时间', true); return; }

    var blob = new Blob([res.ics], { type: 'text/calendar;charset=utf-8' });
    if (state.lastIcsUrl) { try { URL.revokeObjectURL(state.lastIcsUrl); } catch (e) {} }
    state.lastIcsUrl = URL.createObjectURL(blob);
    var filename = ('课表-' + (state.options.semesterStart || '') + '.ics').replace(/[\\/:*?"<>|]/g, '-');

    var link = $('#download-ics');
    link.href = state.lastIcsUrl;
    link.download = filename;
    link.hidden = false;

    var a = document.createElement('a');
    a.href = state.lastIcsUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();

    saveSettings();
    toast('已生成 ' + res.eventCount + ' 个事件，开始下载');
    $('#export-hint').textContent = '如果手机没有自动打开日历：到通知栏或「文件」里点开下载的 .ics 文件，选择日历导入。';
  }

  function icsFileName() {
    return ('课表-' + (state.options.semesterStart || '') + '.ics').replace(/[\\/:*?"<>|]/g, '-');
  }

  /* 安卓 / 移动端：把 .ics 作为文件分享到其他 App（微信、邮箱、网盘、日历 App 等） */
  function shareIcs() {
    var res = previewIcs();
    if (!res || res.error) { toast('生成失败：' + (res && res.error ? res.error : '未知错误'), true); return; }
    if (!res.eventCount) { toast('没有可导出的事件，请先核对课程周次与时间', true); return; }

    var file;
    try {
      file = new File([res.ics], icsFileName(), { type: 'text/calendar' });
    } catch (e) {
      toast('当前浏览器不支持分享文件，请用下载按钮', true);
      return;
    }
    if (!(navigator.canShare && navigator.canShare({ files: [file] }))) {
      toast('当前浏览器不支持分享文件，请用下载按钮', true);
      return;
    }

    navigator.share({
      files: [file],
      title: state.options.calendarName || '课程表'
    }).then(function () {
      saveSettings();
      toast('已打开系统分享，选择日历 / 微信 / 邮箱即可发送');
    }).catch(function (err) {
      if (err && err.name === 'AbortError') return;
      toast('分享失败，请改用下载按钮', true);
    });
  }

  /* 安卓 Chrome 的「安装到主屏幕」 */
  var deferredInstallPrompt = null;
  function setupInstall() {
    if (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) return;
    window.addEventListener('beforeinstallprompt', function (e) {
      e.preventDefault();
      deferredInstallPrompt = e;
      var card = $('#install-card');
      if (card) card.hidden = false;
    });
    var btn = $('#btn-install');
    if (!btn) return;
    btn.addEventListener('click', function () {
      if (!deferredInstallPrompt) {
        toast('请用 Chrome 右上角菜单 →「添加到主屏幕 / 安装应用」');
        return;
      }
      deferredInstallPrompt.prompt();
      deferredInstallPrompt = null;
      $('#install-card').hidden = true;
    });
  }

  /* ---------------- 设置持久化 ---------------- */

  function saveSettings() {
    try {
      var o = Object.assign({}, state.options);
      o.calendarName = state.options.calendarName;
      localStorage.setItem(STORAGE_KEY, JSON.stringify(o));
    } catch (e) {}
  }

  function loadSettings() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      var o = JSON.parse(raw);
      if (!o || typeof o !== 'object') return;
      Object.keys(state.options).forEach(function (k) {
        if (o[k] !== undefined && o[k] !== null && o[k] !== '') state.options[k] = o[k];
      });
    } catch (e) {}
  }

  function resetAll() {
    state.file = null;
    state.buffer = null;
    state.fileName = '';
    state.parse = null;
    state.sheetName = '';
    state.sheetOptions = [];
    state.courses = [];
    state.slots = [];
    state.editedSlots = {};
    state.manual = false;
    state.gridVisible = false;
    state.editId = null;
    state._calibrated = false;
    if (state.lastIcsUrl) { try { URL.revokeObjectURL(state.lastIcsUrl); } catch (e) {} }
    state.lastIcsUrl = null;
    $('#file-input').value = '';
    $('#parse-summary').hidden = true;
    $('#parse-error').hidden = true;
    $('#download-ics').hidden = true;
    $('#grid-wrap').hidden = true;
    $('#review-notice').hidden = true;
    $('#btn-toggle-grid').textContent = '周视图';
    setStep(1);
  }

  /* ---------------- 事件绑定 ---------------- */

  function bind() {
    // 上传
    $('#pick-btn').addEventListener('click', function () { $('#file-input').click(); });
    $('#file-input').addEventListener('change', function (e) {
      var f = e.target.files && e.target.files[0];
      e.target.value = ''; // 允许再次选择同一个文件
      if (f) handleFile(f);
    });
    var drop = $('#drop-zone');
    drop.addEventListener('click', function (e) {
      // 整个拖拽区都可点击选文件（按钮本身已处理时不重复触发）
      if (e.target.closest('button')) return;
      $('#file-input').click();
    });
    ['dragenter', 'dragover'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); drop.classList.add('dragover'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); drop.classList.remove('dragover'); });
    });
    drop.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handleFile(f);
    });
    window.addEventListener('dragover', function (e) { e.preventDefault(); });
    window.addEventListener('drop', function (e) { e.preventDefault(); });

    $('#btn-change-file').addEventListener('click', function () { $('#file-input').click(); });
    $('#btn-retry').addEventListener('click', function () { $('#file-input').click(); });
    $('#btn-manual').addEventListener('click', enterManualMode);
    $('#btn-to-review').addEventListener('click', function () {
      if (!state.courses.length) { toast('没有可核对的课程', true); return; }
      setStep(2);
    });
    $('#sheet-select').addEventListener('change', function (e) { switchSheet(e.target.value); });

    // 核对课程
    $('#courses-preview').addEventListener('click', function (e) {
      var btn = e.target.closest('[data-action]');
      if (!btn) return;
      var card = e.target.closest('.course');
      if (!card) return;
      var id = card.getAttribute('data-id');
      var course = state.courses.filter(function (c) { return c.id === id; })[0];
      if (!course) return;
      var action = btn.getAttribute('data-action');
      if (action === 'edit') openEditDialog(course);
      if (action === 'del') {
        if (window.confirm('确定删除「' + (course.name || '这门课') + '」吗？')) deleteCourse(id);
      }
    });
    $('#btn-add-course').addEventListener('click', function () { openEditDialog(null); });
    $('#btn-add-course-2').addEventListener('click', function () { openEditDialog(null); });
    $('#btn-toggle-grid').addEventListener('click', function () {
      state.gridVisible = !state.gridVisible;
      this.textContent = state.gridVisible ? '列表视图' : '周视图';
      if (state.gridVisible) renderGrid(); else $('#grid-wrap').hidden = true;
    });

    // 编辑对话框
    $('#edit-form').addEventListener('submit', function (e) { e.preventDefault(); saveEditedCourse(); });
    $('#edit-cancel').addEventListener('click', function () { $('#edit-dialog').close(); });
    $('#edit-delete').addEventListener('click', function () {
      if (!state.editId) return;
      var course = state.courses.filter(function (c) { return c.id === state.editId; })[0];
      if (course && window.confirm('确定删除「' + (course.name || '这门课') + '」吗？')) {
        $('#edit-dialog').close();
        deleteCourse(state.editId);
      }
    });

    // 校对时间
    $('#semester-start').addEventListener('change', function () {
      state.options.semesterStart = this.value;
      var d = parseYmd(this.value);
      $('#semester-start-error').hidden = true;
      if (d && d.getDay() !== 1) {
        var monday = mondayOfWeek(d);
        this.value = ymd(monday);
        state.options.semesterStart = ymd(monday);
        toast('已自动调整为该周的周一：' + ymd(monday));
      }
      saveSettings();
    });
    ['first-period-time', 'afternoon-start'].forEach(function (id) {
      $('#' + id).addEventListener('change', function () {
        if (id === 'first-period-time') state.options.firstPeriodTime = this.value;
        else state.options.afternoonStart = this.value;
        state.editedSlots = {};
        computeSlotTimes(false);
        renderSlotTable();
        saveSettings();
      });
    });
    $('#schedule-preset').addEventListener('change', function () {
      var v = this.value;
      if (v === 'custom') {
        $('#custom-time-fields').hidden = false;
        return;
      }
      var p = v.split('-');
      state.options.lessonMinutes = +p[0];
      state.options.breakMinutes = +p[1];
      $('#lesson-minutes').value = p[0];
      $('#break-minutes').value = p[1];
      $('#custom-time-fields').hidden = true;
      state.editedSlots = {};
      computeSlotTimes(false);
      renderSlotTable();
      saveSettings();
    });
    ['lesson-minutes', 'break-minutes'].forEach(function (id) {
      $('#' + id).addEventListener('change', function () {
        if (id === 'lesson-minutes') state.options.lessonMinutes = +this.value;
        else state.options.breakMinutes = +this.value;
        updatePresetUI();
        state.editedSlots = {};
        computeSlotTimes(false);
        renderSlotTable();
        saveSettings();
      });
    });
    $('#btn-reset-slots').addEventListener('click', function () {
      state.editedSlots = {};
      computeSlotTimes(false);
      renderSlotTable();
      toast('已按上方设置重新推算所有时段');
    });
    $('#slot-times').addEventListener('change', function (e) {
      var t = e.target;
      var i = t.getAttribute('data-slot-start');
      if (i == null) i = t.getAttribute('data-slot-end');
      if (i == null) return;
      i = +i;
      state.editedSlots[i] = true;
      if (t.hasAttribute('data-slot-start')) state.slots[i].start = t.value;
      else state.slots[i].end = t.value;
      renderPreview();
    });

    // 导出设置
    $('#opt-calendar-name').addEventListener('input', function () { state.options.calendarName = this.value; saveSettings(); });
    $('#opt-timezone').addEventListener('change', function () { state.options.timeZone = this.value; saveSettings(); prepareExport(); });
    $('#opt-reminder').addEventListener('change', function () { state.options.reminderMinutes = +this.value; saveSettings(); });
    $('#opt-mode').addEventListener('change', function () { state.options.mode = this.value; saveSettings(); prepareExport(); });
    $('#btn-export').addEventListener('click', downloadIcs);
    $('#btn-share').addEventListener('click', shareIcs);

    // 导航
    $('#btn-next').addEventListener('click', function () {
      if (state.step === 2) {
        if (!state.courses.length) { toast('请先添加至少一门课程', true); return; }
        setStep(3);
      } else if (state.step === 3) {
        if (!validateCalibration()) return;
        setStep(4);
      }
    });
    $('#btn-back').addEventListener('click', function () {
      if (state.step > 1) setStep(state.step - 1);
    });
    $('#btn-restart').addEventListener('click', function () {
      if (window.confirm('重新开始会清空当前课表，但会保留你的时间设置。继续吗？')) resetAll();
    });

    // 快捷键：Esc 关闭对话框由原生处理
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') {
        var dlg = $('#edit-dialog');
        if (dlg.open) dlg.close();
      }
    });
  }

  /* ---------------- 初始化 ---------------- */

  function init() {
    loadSettings();
    syncCalibrationInputs();
    bind();
    setStep(1);

    // 安装到主屏幕（安卓 Chrome）
    setupInstall();

    // 分享 .ics（安卓 Chrome / 部分浏览器支持文件分享）
    try {
      if (navigator.canShare && navigator.canShare({ files: [new File(['x'], 'a.ics', { type: 'text/calendar' })] })) {
        $('#btn-share').hidden = false;
      }
    } catch (e) {}

    // PWA：仅在线部署时注册，避免本地开发缓存干扰
    if ('serviceWorker' in navigator && location.protocol === 'https:' &&
        location.hostname !== 'localhost' && location.hostname !== '127.0.0.1') {
      window.addEventListener('load', function () {
        navigator.serviceWorker.register('sw.js').catch(function () {});
      });
    }

    // 调试接口（E2E/排查用）
    window.__TT_APP__ = {
      state: state,
      setStep: setStep,
      parseWeeksInput: parseWeeksInput,
      formatWeeks: formatWeeks
    };
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
