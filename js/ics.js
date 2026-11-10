/* ============================================================
 * ICS 生成模块（课程 → RFC 5545 日历）
 * 浏览器：window.TT.buildICS
 * Node：  module.exports = { buildICS }
 * 无第三方依赖
 * ============================================================ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.TT = root.TT || {};
    var api = factory();
    for (var k in api) if (Object.prototype.hasOwnProperty.call(api, k)) root.TT[k] = api[k];
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var VERSION = '1.0.0';
  var DEFAULT_TZ = 'Asia/Shanghai';
  var ZONES = {
    'Asia/Shanghai': { offset: '+0800', name: 'CST' },
    'Asia/Hong_Kong': { offset: '+0800', name: 'HKT' },
    'Asia/Taipei': { offset: '+0800', name: 'CST' },
    'Asia/Singapore': { offset: '+0800', name: 'SGT' },
    'Asia/Tokyo': { offset: '+0900', name: 'JST' },
    'UTC': { offset: '+0000', name: 'UTC' }
  };
  var WEEKDAYS = ['', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA', 'SU'];
  var DAY_NAMES = ['', '周一', '周二', '周三', '周四', '周五', '周六', '周日'];

  /* ---------------- 小工具 ---------------- */

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  function trim(v) { return v == null ? '' : String(v).replace(/\u00a0/g, ' ').trim(); }

  function escapeText(v) {
    return String(v == null ? '' : v)
      .replace(/\\/g, '\\\\')
      .replace(/;/g, '\\;')
      .replace(/,/g, '\\,')
      .replace(/\r\n?/g, '\\n')
      .replace(/\n/g, '\\n');
  }

  function parseYmd(v) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trim(v));
    if (!m) return null;
    var y = +m[1], mo = +m[2], d = +m[3];
    var date = new Date(y, mo - 1, d);
    if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) return null;
    return date;
  }

  function addDays(date, days) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);
  }

  function formatDate(date) {
    return date.getFullYear() + pad2(date.getMonth() + 1) + pad2(date.getDate());
  }

  function formatDateTimeLocal(date, timeStr) {
    var m = /^(\d{1,2}):(\d{2})$/.exec(trim(timeStr));
    if (!m) return null;
    return formatDate(date) + 'T' + pad2(+m[1]) + m[2] + '00';
  }

  function formatUtcStamp(date) {
    return date.getUTCFullYear() + pad2(date.getUTCMonth() + 1) + pad2(date.getUTCDate()) +
      'T' + pad2(date.getUTCHours()) + pad2(date.getUTCMinutes()) + pad2(date.getUTCSeconds()) + 'Z';
  }

  function timeToMinutes(v) {
    var m = /^(\d{1,2}):(\d{2})$/.exec(trim(v));
    if (!m) return null;
    var h = +m[1], mm = +m[2];
    if (h > 23 || mm > 59) return null;
    return h * 60 + mm;
  }

  function normalizeWeeks(weeks) {
    if (!Array.isArray(weeks)) return [];
    var seen = {}, out = [];
    weeks.forEach(function (w) {
      w = +w;
      if (Number.isFinite(w) && w > 0 && w <= 60 && !seen[w]) { seen[w] = 1; out.push(w); }
    });
    return out.sort(function (a, b) { return a - b; });
  }

  function compressWeeks(weeks) {
    if (!weeks.length) return '';
    var parts = [];
    var start = weeks[0], prev = weeks[0];
    for (var i = 1; i <= weeks.length; i++) {
      var cur = weeks[i];
      if (i < weeks.length && cur === prev + 1) { prev = cur; continue; }
      parts.push(start === prev ? String(start) : start + '-' + prev);
      start = prev = cur;
    }
    return parts.join(',');
  }

  function hash(str) {
    var h = 2166136261;
    for (var i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = (h * 16777619) >>> 0;
    }
    return ('0000000' + h.toString(16)).slice(-8);
  }

  // 按 UTF-8 字节数折行，续行以单个空格开头；不切断多字节字符
  function foldLine(line) {
    var out = [];
    var cur = '';
    var bytes = 0;
    var limit = 75;
    var chars = Array.from(line);
    for (var i = 0; i < chars.length; i++) {
      var ch = chars[i];
      var b = utf8Len(ch);
      if (bytes + b > limit && cur) {
        out.push(cur);
        cur = ' ' + ch;
        bytes = 1 + b;
      } else {
        cur += ch;
        bytes += b;
      }
    }
    out.push(cur);
    return out.join('\r\n');
  }

  function utf8Len(ch) {
    var code = ch.codePointAt(0);
    if (code < 0x80) return 1;
    if (code < 0x800) return 2;
    if (code < 0x10000) return 3;
    return 4;
  }

  function vtimezone(tz) {
    var zone = ZONES[tz] || ZONES[DEFAULT_TZ];
    return [
      'BEGIN:VTIMEZONE',
      'TZID:' + tz,
      'BEGIN:STANDARD',
      'DTSTART:19700101T000000',
      'TZOFFSETFROM:' + zone.offset,
      'TZOFFSETTO:' + zone.offset,
      'TZNAME:' + zone.name,
      'END:STANDARD',
      'END:VTIMEZONE'
    ].join('\r\n');
  }

  /* ---------------- 主函数 ---------------- */

  function buildICS(input) {
    input = input || {};
    var courses = Array.isArray(input.courses) ? input.courses : [];
    var slots = Array.isArray(input.slots) ? input.slots : [];
    var settings = input.settings || {};
    var errors = [];
    var warnings = [];

    var tz = trim(settings.timeZone) || DEFAULT_TZ;
    if (!ZONES[tz]) {
      warnings.push({ code: 'UNKNOWN_TIMEZONE', message: '不支持的时区「' + tz + '」，已回退到 ' + DEFAULT_TZ });
      tz = DEFAULT_TZ;
    }

    var semesterStart = parseYmd(settings.semesterStart);
    if (!semesterStart) {
      errors.push({ code: 'INVALID_SEMESTER_START', message: '缺少有效的学期第 1 周周一日期（YYYY-MM-DD）' });
    } else if (semesterStart.getDay() !== 1) {
      warnings.push({ code: 'SEMESTER_START_NOT_MONDAY', message: '学期开始日期不是周一，已按所选日期计算，请确认日期是否正确。' });
    }

    var calendarName = trim(settings.calendarName) || '课程表';
    var reminder = +settings.reminderMinutes || 0;
    if (reminder < 0) reminder = 0;
    var mode = settings.mode === 'single' ? 'single' : 'recurring';
    var now = settings.now instanceof Date ? settings.now : new Date();

    /* 停课日期（全局）：这些日期上的课程不生成事件 */
    var skipSet = {};
    if (Array.isArray(settings.skipDates)) {
      settings.skipDates.forEach(function (v) {
        var d = parseYmd(v);
        if (d) skipSet[formatDate(d)] = 1;
        else warnings.push({ code: 'INVALID_SKIP_DATE', message: '忽略无效的停课日期「' + v + '」' });
      });
    }
    var dtstamp = formatUtcStamp(now);

    var lines = [];
    lines.push('BEGIN:VCALENDAR');
    lines.push('VERSION:2.0');
    lines.push('PRODID:-//ClassTable2ICS//CN');
    lines.push('CALSCALE:GREGORIAN');
    lines.push('METHOD:PUBLISH');
    lines.push(foldLine('X-WR-CALNAME:' + escapeText(calendarName)));
    lines.push('X-WR-TIMEZONE:' + tz);
    lines.push(vtimezone(tz));

    if (!semesterStart) {
      lines.push('END:VCALENDAR');
      return { ics: lines.join('\r\n') + '\r\n', eventCount: 0, errors: errors, warnings: warnings };
    }

    var eventCount = 0;

    function eventLines(opts) {
      var out = [];
      out.push('BEGIN:VEVENT');
      out.push('UID:' + opts.uid);
      out.push('DTSTAMP:' + dtstamp);
      out.push('DTSTART;TZID=' + tz + ':' + opts.start);
      out.push('DTEND;TZID=' + tz + ':' + opts.end);
      if (opts.rrule) out.push(opts.rrule);
      if (opts.exdates && opts.exdates.length) {
        out.push(foldLine('EXDATE;TZID=' + tz + ':' + opts.exdates.join(',')));
      }
      out.push(foldLine('SUMMARY:' + escapeText(opts.summary)));
      if (opts.location) out.push(foldLine('LOCATION:' + escapeText(opts.location)));
      out.push(foldLine('DESCRIPTION:' + escapeText(opts.description)));
      if (reminder > 0) {
        out.push('BEGIN:VALARM');
        out.push('ACTION:DISPLAY');
        out.push(foldLine('DESCRIPTION:' + escapeText(opts.summary)));
        out.push('TRIGGER:-PT' + reminder + 'M');
        out.push('END:VALARM');
      }
      out.push('END:VEVENT');
      return out;
    }

    courses.forEach(function (course, index) {
      if (!course || typeof course !== 'object') return;
      var name = trim(course.name) || '课程';
      var weekday = +course.weekday;
      if (!(weekday >= 1 && weekday <= 7)) {
        warnings.push({ code: 'INVALID_WEEKDAY', message: '跳过「' + name + '」：星期无效' });
        return;
      }
      var weeks = normalizeWeeks(course.weeks);
      if (!weeks.length) {
        warnings.push({ code: 'NO_WEEKS', message: '跳过「' + name + '」：没有周次' });
        return;
      }
      var slotIndex = +course.slotIndex;
      var slotEndIndex = course.slotEndIndex == null ? slotIndex : +course.slotEndIndex;
      var startSlot = slots[slotIndex];
      var endSlot = slots[slotEndIndex] || startSlot;
      if (!startSlot || !endSlot) {
        warnings.push({ code: 'INVALID_SLOT', message: '跳过「' + name + '」：时段不存在' });
        return;
      }
      var startTime = trim(startSlot.start);
      var endTime = trim(endSlot.end) || trim(startSlot.end);
      if (timeToMinutes(startTime) == null || timeToMinutes(endTime) == null) {
        warnings.push({ code: 'MISSING_TIME', message: '跳过「' + name + '」：' + (DAY_NAMES[weekday]) + ' ' + trim(startSlot.label || '') + ' 缺少上课时间，请回到「校对时间」补充。' });
        return;
      }
      if (timeToMinutes(endTime) <= timeToMinutes(startTime)) {
        warnings.push({ code: 'INVALID_TIME_RANGE', message: '跳过「' + name + '」：结束时间不晚于开始时间' });
        return;
      }

      var teacher = trim(course.teacher);
      var description = '教师：' + teacher + '\n周次：' + compressWeeks(weeks) + '周';
      var location = trim(course.location);
      var key = [name, teacher, location, weekday, slotIndex, slotEndIndex, weeks.join(',')].join('|');

      if (mode === 'single') {
        weeks.forEach(function (w) {
          var date = addDays(semesterStart, (w - 1) * 7 + (weekday - 1));
          if (skipSet[formatDate(date)]) return;
          var start = formatDateTimeLocal(date, startTime);
          var end = formatDateTimeLocal(date, endTime);
          if (!start || !end) return;
          lines = lines.concat(eventLines({
            uid: 'c' + index + '-' + w + '-' + hash(key + '|' + w) + '@class2cal',
            start: start,
            end: end,
            summary: name,
            location: location,
            description: description
          }));
          eventCount++;
        });
        return;
      }

      var first = weeks[0], last = weeks[weeks.length - 1];
      var firstDate = addDays(semesterStart, (first - 1) * 7 + (weekday - 1));
      var startLocal = formatDateTimeLocal(firstDate, startTime);
      var endLocal = formatDateTimeLocal(firstDate, endTime);
      if (!startLocal || !endLocal) {
        warnings.push({ code: 'INVALID_TIME_RANGE', message: '跳过「' + name + '」：时间格式错误' });
        return;
      }
      if (weeks.length === 1 && skipSet[formatDate(firstDate)]) {
        return;
      }
      var hasWeek = {};
      weeks.forEach(function (w) { hasWeek[w] = 1; });
      var exdates = [];
      for (var w = first; w <= last; w++) {
        var d = addDays(semesterStart, (w - 1) * 7 + (weekday - 1));
        if (!hasWeek[w] || skipSet[formatDate(d)]) {
          var stamp = formatDateTimeLocal(d, startTime);
          if (stamp) exdates.push(stamp);
        }
      }
      var rrule = weeks.length > 1
        ? 'RRULE:FREQ=WEEKLY;COUNT=' + (last - first + 1) + ';BYDAY=' + WEEKDAYS[weekday] + ';WKST=MO'
        : '';
      lines = lines.concat(eventLines({
        uid: 'c' + index + '-' + hash(key) + '@class2cal',
        start: startLocal,
        end: endLocal,
        rrule: rrule,
        exdates: exdates,
        summary: name,
        location: location,
        description: description
      }));
      eventCount++;
    });

    /* 补课 / 临时调课：额外单次事件 */
    var extraEvents = Array.isArray(settings.extraEvents) ? settings.extraEvents : [];
    extraEvents.forEach(function (ev, i) {
      if (!ev || typeof ev !== 'object') return;
      var name = trim(ev.name) || '补课';
      var date = parseYmd(ev.date);
      var startTime = trim(ev.start);
      var endTime = trim(ev.end);
      if (!date || timeToMinutes(startTime) == null || timeToMinutes(endTime) == null || timeToMinutes(endTime) <= timeToMinutes(startTime)) {
        warnings.push({ code: 'INVALID_EXTRA_EVENT', message: '忽略补课「' + name + '」：日期或时间无效' });
        return;
      }
      var start = formatDateTimeLocal(date, startTime);
      var end = formatDateTimeLocal(date, endTime);
      if (!start || !end) {
        warnings.push({ code: 'INVALID_EXTRA_EVENT', message: '忽略补课「' + name + '」：日期或时间无效' });
        return;
      }
      var location = trim(ev.location);
      var key = 'x|' + ev.date + '|' + startTime + '|' + endTime + '|' + name + '|' + location;
      lines = lines.concat(eventLines({
        uid: 'x' + i + '-' + hash(key) + '@class2cal',
        start: start,
        end: end,
        summary: name,
        location: location,
        description: '补课 / 临时调课'
      }));
      eventCount++;
    });

    lines.push('END:VCALENDAR');
    var ics = lines.join('\r\n') + '\r\n';

    return {
      ics: ics,
      eventCount: eventCount,
      errors: errors,
      warnings: warnings
    };
  }

  return { buildICS: buildICS, VERSION: VERSION };
});
