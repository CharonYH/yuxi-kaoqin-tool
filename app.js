(() => {
  'use strict';

  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];
  const state = {
    workbook: null, report: null, selectedEmployeeIndex: 0, selectedDay: 1,
    filter: 'all', query: '', scope: 'all', sort: 'original'
  };
  const { WEEKDAYS, FORMAT_ERROR_MESSAGE, clean } = YuxiAttendanceParser;

  const els = {
    uploadView: $('#uploadView'), dashboard: $('#dashboard'), fileInput: $('#fileInput'), dropZone: $('#dropZone'),
    backButton: $('#backButton'), exportButton: $('#exportButton'), fileName: $('#fileName'), periodText: $('#periodText'),
    employeeCount: $('#employeeCount'), punchCount: $('#punchCount'), recordedDays: $('#recordedDays'), emptyDays: $('#emptyDays'),
    coverageChart: $('#coverageChart'), attendanceMatrix: $('#attendanceMatrix'), employeePanel: $('#employeePanel'),
    searchInput: $('#searchInput'), sortSelect: $('#sortSelect'), resultCount: $('#resultCount'), toast: $('#toast')
  };

  function getWeekday(day) {
    return new Date(Date.UTC(state.report.year, state.report.month - 1, day)).getUTCDay();
  }

  function isWeekend(day) {
    const weekday = getWeekday(day);
    return weekday === 0 || weekday === 6;
  }

  function statusFor(times) {
    if (!times.length) return 'empty';
    if (times.length === 1) return 'single';
    if (times.length === 2) return 'pair';
    return 'multiple';
  }

  function formatNumber(value) {
    return new Intl.NumberFormat('zh-CN').format(value);
  }

  function escapeHtml(value) {
    return clean(value).replace(/[&<>'"]/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' })[character]);
  }

  function showToast(message, isError = false) {
    els.toast.textContent = message;
    els.toast.classList.toggle('is-error', isError);
    els.toast.classList.add('is-visible');
    window.clearTimeout(showToast.timer);
    showToast.timer = window.setTimeout(() => els.toast.classList.remove('is-visible'), 3200);
  }

  async function loadFile(file) {
    if (!file) return;
    if (!/\.xlsx?$/i.test(file.name)) { showToast(FORMAT_ERROR_MESSAGE, true); return; }
    els.dropZone.classList.add('is-loading');
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
      state.workbook = workbook;
      state.report = YuxiAttendanceParser.parseReport(workbook, file.name, XLSX);
      state.selectedEmployeeIndex = 0;
      state.selectedDay = 1;
      state.filter = 'all';
      state.query = '';
      state.scope = 'all';
      state.sort = 'original';
      renderDashboard();
      els.uploadView.hidden = true;
      els.dashboard.hidden = false;
      window.scrollTo({ top: 0, behavior: 'smooth' });
      showToast(`已读取 ${state.report.employees.length} 名员工的考勤记录`);
    } catch (error) {
      console.error(error);
      showToast(FORMAT_ERROR_MESSAGE, true);
    } finally {
      els.dropZone.classList.remove('is-loading');
      els.fileInput.value = '';
    }
  }

  function renderDashboard() {
    const report = state.report;
    els.fileName.textContent = report.fileName;
    els.periodText.textContent = `${report.period} · ${report.sheetName}`;
    els.employeeCount.textContent = formatNumber(report.employees.length);
    els.punchCount.textContent = formatNumber(report.totals.punches);
    els.recordedDays.textContent = formatNumber(report.totals.recorded);
    els.emptyDays.textContent = formatNumber(report.totals.empty);
    els.searchInput.value = '';
    els.sortSelect.value = 'original';
    $$('.filter-chip').forEach(button => button.classList.toggle('is-active', button.dataset.filter === 'all'));
    $$('.scope-chip').forEach(button => button.classList.toggle('is-active', button.dataset.scope === 'all'));
    renderCoverage();
    renderMatrix();
    renderEmployee();
  }

  function renderCoverage() {
    const report = state.report;
    els.coverageChart.style.gridTemplateColumns = `repeat(${report.days.length}, minmax(14px, 1fr))`;
    els.coverageChart.style.minWidth = `${Math.max(640, report.days.length * 23)}px`;
    const counts = report.days.map(day => report.employees.filter(employee => employee.days.find(item => item.day === day)?.times.length).length);
    const max = Math.max(...counts, 1);
    els.coverageChart.innerHTML = report.days.map((day, index) => {
      const count = counts[index];
      const weekend = isWeekend(day);
      return `<div class="coverage-bar-wrap ${weekend ? 'is-weekend' : ''}">
        <span class="bar-tooltip">${day} 日 · ${count}/${report.employees.length} 人</span>
        <span class="coverage-bar" style="height:${Math.max(3, (count / max) * 88)}px"></span>
        <small>${day}</small>
      </div>`;
    }).join('');
  }

  function scopedDays() {
    return state.report.days.filter(day => {
      if (state.scope === 'workday') return !isWeekend(day);
      if (state.scope === 'weekend') return isWeekend(day);
      return true;
    });
  }

  function filteredEmployees() {
    const query = state.query.toLowerCase();
    const days = scopedDays();
    const result = state.report.employees.map((employee, originalIndex) => ({ employee, originalIndex })).filter(({ employee }) => {
      const matchesQuery = !query || `${employee.name} ${employee.id} ${employee.department}`.toLowerCase().includes(query);
      const scopedEntries = employee.days.filter(entry => days.includes(entry.day));
      const matchesStatus = state.filter === 'all'
        || (state.filter === 'empty' && scopedEntries.some(entry => entry.times.length === 0))
        || (state.filter === 'single' && scopedEntries.some(entry => entry.times.length === 1))
        || (state.filter === 'multiple' && scopedEntries.some(entry => statusFor(entry.times) === 'multiple'));
      return matchesQuery && matchesStatus;
    });
    if (state.sort === 'name') result.sort((a, b) => a.employee.name.localeCompare(b.employee.name, 'zh-CN'));
    return result;
  }

  function renderMatrix() {
    const report = state.report;
    const visibleDays = scopedDays();
    const people = filteredEmployees();
    if (people.length && !people.some(({ originalIndex }) => originalIndex === state.selectedEmployeeIndex)) {
      state.selectedEmployeeIndex = people[0].originalIndex;
    }
    if (!visibleDays.includes(state.selectedDay)) state.selectedDay = visibleDays[0] || 1;
    els.resultCount.textContent = `${people.length} 名员工 · ${visibleDays.length} 天`;
    const headers = `<div class="matrix-cell matrix-corner">员工 / 日期</div>${visibleDays.map(day => {
      const weekday = getWeekday(day);
      return `<div class="matrix-cell matrix-day ${isWeekend(day) ? 'is-weekend' : ''}"><span>${String(day).padStart(2, '0')}</span><small>周${WEEKDAYS[weekday]}</small></div>`;
    }).join('')}`;

    const rows = people.length ? people.map(({ employee, originalIndex }) => {
      const name = `<button class="matrix-cell matrix-name ${originalIndex === state.selectedEmployeeIndex ? 'is-selected' : ''}" type="button" data-employee="${originalIndex}"><span>${escapeHtml(employee.name)}</span><small>${escapeHtml(employee.id)}</small></button>`;
      const cells = visibleDays.map(day => {
        const entry = employee.days.find(item => item.day === day) || { times: [] };
        const status = statusFor(entry.times);
        const label = `${employee.name}，${report.month}月${day}日，${entry.times.length ? entry.times.join('、') : '无记录'}`;
        return `<div class="matrix-cell"><button type="button" class="day-button status-${status} ${isWeekend(day) ? 'is-weekend' : ''} ${originalIndex === state.selectedEmployeeIndex && day === state.selectedDay ? 'is-selected' : ''}" data-employee="${originalIndex}" data-day="${day}" data-count="${entry.times.length}" aria-label="${escapeHtml(label)}" title="${escapeHtml(label)}"></button></div>`;
      }).join('');
      return name + cells;
    }).join('') : '<div class="matrix-no-results">没有符合当前条件的员工</div>';

    els.attendanceMatrix.style.gridTemplateColumns = `154px repeat(${visibleDays.length}, 40px)`;
    els.attendanceMatrix.style.minWidth = `${154 + visibleDays.length * 40 + 56}px`;
    els.attendanceMatrix.innerHTML = headers + rows;
    $$('[data-employee]', els.attendanceMatrix).forEach(button => button.addEventListener('click', () => {
      state.selectedEmployeeIndex = Number(button.dataset.employee);
      if (button.dataset.day) state.selectedDay = Number(button.dataset.day);
      renderMatrix();
      renderEmployee();
    }));
  }

  function employeeStats(employee) {
    return employee.days.reduce((acc, entry) => {
      const status = statusFor(entry.times);
      if (status === 'empty') acc.zeroDays += 1;
      if (status === 'single') acc.oneDay += 1;
      if (status === 'pair') acc.twoDays += 1;
      if (status === 'multiple') acc.multipleDays += 1;
      return acc;
    }, { zeroDays: 0, oneDay: 0, twoDays: 0, multipleDays: 0 });
  }

  function renderEmployee() {
    const report = state.report;
    const visiblePeople = filteredEmployees();
    if (!visiblePeople.length) {
      els.employeePanel.innerHTML = '<div class="employee-empty"><span class="empty-glyph" aria-hidden="true"></span><h2>没有匹配结果</h2><p>请调整姓名、工号、记录状态或日期范围。</p></div>';
      return;
    }
    const employee = report.employees[state.selectedEmployeeIndex];
    if (!employee) return;
    const stats = employeeStats(employee);
    const firstWeekday = getWeekday(1);
    const blankDays = Array.from({ length: firstWeekday }, () => '<div class="calendar-day calendar-day--blank"></div>').join('');
    const calendarDays = report.days.map(day => {
      const entry = employee.days.find(item => item.day === day) || { times: [] };
      const status = statusFor(entry.times);
      return `<button type="button" class="calendar-day status-${status} ${isWeekend(day) ? 'is-weekend' : ''} ${day === state.selectedDay ? 'is-active' : ''}" data-calendar-day="${day}" aria-label="${report.month}月${day}日，${entry.times.length ? entry.times.join('、') : '无记录'}">
        <b>${day}</b><span>${entry.times.length ? escapeHtml(entry.times.join(' · ')) : '—'}</span>
      </button>`;
    }).join('');
    const selected = employee.days.find(item => item.day === state.selectedDay) || { times: [] };
    els.employeePanel.innerHTML = `
      <div class="employee-head">
        <div class="employee-identity"><span class="employee-avatar">${escapeHtml(employee.name.slice(-1))}</span><div><h2>${escapeHtml(employee.name)}</h2><p>NO. ${escapeHtml(employee.id)} · ${escapeHtml(employee.department)}</p></div></div>
        <div class="employee-nav">
          <button class="icon-button" type="button" data-nav="prev" aria-label="上一名员工"><svg viewBox="0 0 24 24"><path d="m15 18-6-6 6-6"/></svg></button>
          <button class="icon-button" type="button" data-nav="next" aria-label="下一名员工"><svg viewBox="0 0 24 24"><path d="m9 18 6-6-6-6"/></svg></button>
        </div>
      </div>
      <div class="employee-stats">
        <div class="mini-stat mini-stat--zero"><strong>${stats.zeroDays}</strong><span>0 次天数</span></div>
        <div class="mini-stat mini-stat--one"><strong>${stats.oneDay}</strong><span>1 次天数</span></div>
        <div class="mini-stat mini-stat--two"><strong>${stats.twoDays}</strong><span>2 次天数</span></div>
        <div class="mini-stat mini-stat--multiple"><strong>${stats.multipleDays}</strong><span>3 次及以上</span></div>
      </div>
      <div class="calendar-labels">${WEEKDAYS.map(day => `<span>${day}</span>`).join('')}</div>
      <div class="calendar-grid">${blankDays}${calendarDays}</div>
      <div class="selected-day-detail">
        <p>${report.year} 年 ${report.month} 月 ${state.selectedDay} 日 · 周${WEEKDAYS[getWeekday(state.selectedDay)]} · ${selected.times.length} 次打卡</p>
        <strong class="${selected.times.length ? '' : 'no-punch'}">${selected.times.length ? selected.times.map(escapeHtml).join(' / ') : '无打卡记录'}</strong>
      </div>`;

    $$('[data-calendar-day]', els.employeePanel).forEach(button => button.addEventListener('click', () => {
      state.selectedDay = Number(button.dataset.calendarDay); renderMatrix(); renderEmployee();
    }));
    $$('[data-nav]', els.employeePanel).forEach(button => button.addEventListener('click', () => {
      const direction = button.dataset.nav === 'next' ? 1 : -1;
      state.selectedEmployeeIndex = (state.selectedEmployeeIndex + direction + report.employees.length) % report.employees.length;
      renderMatrix(); renderEmployee();
    }));
  }

  function exportCsv() {
    const report = state.report;
    const rows = [['工号', '姓名', '部门', '日期', '星期', '打卡记录', '打卡次数']];
    report.employees.forEach(employee => employee.days.forEach(entry => {
      rows.push([
        employee.id, employee.name, employee.department,
        `${report.year}-${String(report.month).padStart(2, '0')}-${String(entry.day).padStart(2, '0')}`,
        `周${WEEKDAYS[getWeekday(entry.day)]}`, entry.times.join(' / '), entry.times.length
      ]);
    }));
    const csv = '\ufeff' + rows.map(row => row.map(value => `"${String(value).replace(/"/g, '""')}"`).join(',')).join('\r\n');
    const link = document.createElement('a');
    link.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    link.download = `${report.fileName.replace(/\.xlsx?$/i, '')}_考勤记录明细.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
    showToast('明细 CSV 已生成');
  }

  els.fileInput.addEventListener('change', event => loadFile(event.target.files[0]));
  ['dragenter', 'dragover'].forEach(type => els.dropZone.addEventListener(type, event => { event.preventDefault(); els.dropZone.classList.add('is-dragging'); }));
  ['dragleave', 'drop'].forEach(type => els.dropZone.addEventListener(type, event => { event.preventDefault(); els.dropZone.classList.remove('is-dragging'); }));
  els.dropZone.addEventListener('drop', event => loadFile(event.dataTransfer.files[0]));
  els.backButton.addEventListener('click', () => { els.dashboard.hidden = true; els.uploadView.hidden = false; state.report = null; });
  els.exportButton.addEventListener('click', exportCsv);
  els.searchInput.addEventListener('input', event => {
    state.query = event.target.value.trim();
    renderMatrix();
    renderEmployee();
  });
  els.sortSelect.addEventListener('change', event => {
    state.sort = event.target.value;
    renderMatrix();
    renderEmployee();
  });
  $$('.scope-chip').forEach(button => button.addEventListener('click', () => {
    state.scope = button.dataset.scope;
    $$('.scope-chip').forEach(item => item.classList.toggle('is-active', item === button));
    renderMatrix();
    renderEmployee();
  }));
  $$('.filter-chip').forEach(button => button.addEventListener('click', () => {
    state.filter = button.dataset.filter;
    $$('.filter-chip').forEach(item => item.classList.toggle('is-active', item === button));
    renderMatrix();
    renderEmployee();
  }));
})();
