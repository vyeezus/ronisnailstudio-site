// --- Roni's Nail Studio — Custom Booking Calendar (reschedule + new booking) ---

function bootBookingPage() {
    const params = new URLSearchParams(window.location.search);
    let serviceNames = params.get('names');
    let serviceIds = params.get('services');
    const priceStr = params.get('price');
    let timeStr = params.get('time');
    const isReschedule = params.get('reschedule') === '1';
    const rescheduleEventId = params.get('eventId');
    const rescheduleToken = params.get('token');

    const SCRIPT_PROXY = `${window.location.origin}/api/booking`;
    const SCRIPT_DIRECT =
        'https://script.google.com/macros/s/AKfycbzdT_rV3dR7Th4VHeLE3uJcyTPr4bI-6uy-_Im6xz-nZ0rGPToj85zy7Is7LmpNVS0Wwg/exec';

    /** Google Apps Script cold starts + mobile networks often exceed 8s; keep under ~30s browser limits. */
    const JSONP_TIMEOUT_MS = 26000;

    function safeDecodeURIComponent(str) {
        if (str == null || str === '') return '';
        const s = String(str);
        try {
            return decodeURIComponent(s.replace(/\+/g, ' '));
        } catch (e) {
            return s;
        }
    }

    const DEFAULT_WORK_HOURS = {
        1: { start: 11, end: 18 },
        2: { start: 11, end: 18 },
        3: { start: 9, end: 16 },
        5: { start: 9, end: 16 },
    };
    let workHours = { ...DEFAULT_WORK_HOURS };

    function dayHours(dayOfWeek) {
        return workHours[dayOfWeek] || workHours[String(dayOfWeek)];
    }

    function mergeWorkHoursFromPayload(data) {
        if (!data || typeof data !== 'object' || Array.isArray(data)) return;
        const next = {};
        for (const k of Object.keys(data)) {
            const d = parseInt(k, 10);
            if (isNaN(d) || d < 0 || d > 6) continue;
            const h = data[k];
            const st = typeof h.start === 'number' ? h.start : parseFloat(h.start);
            const en = typeof h.end === 'number' ? h.end : parseFloat(h.end);
            if (!isFinite(st) || !isFinite(en)) continue;
            const start = Math.floor(st);
            const end = Math.floor(en);
            if (start < 0 || end > 24 || start >= end) continue;
            next[d] = { start, end };
        }
        if (Object.keys(next).length > 0) workHours = next;
    }

    function loadWorkHoursJsonp(baseUrl) {
        return new Promise((resolve, reject) => {
            const callbackName = 'workHours_' + Date.now();
            const timeoutId = setTimeout(() => {
                cleanup();
                reject(new Error('work hours jsonp timeout'));
            }, JSONP_TIMEOUT_MS);

            function cleanup() {
                clearTimeout(timeoutId);
                try {
                    delete window[callbackName];
                } catch (e) { /* ignore */ }
                const tag = document.getElementById('jsonp-work-hours');
                if (tag) tag.remove();
            }

            window[callbackName] = function (payload) {
                mergeWorkHoursFromPayload(payload);
                cleanup();
                resolve();
            };

            const qs = new URLSearchParams({
                callback: callbackName,
                action: 'work_hours',
            });
            const script = document.createElement('script');
            script.id = 'jsonp-work-hours';
            script.src = `${baseUrl}?${qs.toString()}`;
            script.onerror = () => {
                cleanup();
                reject(new Error('work hours jsonp load error'));
            };
            document.body.appendChild(script);
        });
    }

    async function loadJsonpTwice_(loader) {
        let lastErr;
        for (let attempt = 0; attempt < 2; attempt++) {
            try {
                return await loader();
            } catch (e) {
                lastErr = e;
            }
        }
        throw lastErr;
    }

    async function loadWorkHours() {
        try {
            await loadJsonpTwice_(() => loadWorkHoursJsonp(SCRIPT_PROXY));
        } catch (e) {
            console.warn('Proxy /api/booking failed; loading work hours from Google directly.', e);
            try {
                await loadJsonpTwice_(() => loadWorkHoursJsonp(SCRIPT_DIRECT));
            } catch (e2) {
                console.error('Work hours JSONP failed entirely.', e2);
            }
        }
    }

    let serviceMinutes = 60;
    if (timeStr) {
        const decodedTime = decodeURIComponent(timeStr).toLowerCase();
        const hMatch = decodedTime.match(/(\d+)\s*h/);
        const mMatch = decodedTime.match(/(\d+)\s*m/);
        let total = 0;
        if (hMatch) total += parseInt(hMatch[1], 10) * 60;
        if (mMatch) total += parseInt(mMatch[1], 10);
        if (total > 0) serviceMinutes = total;
    }

    const bookingFlow = document.getElementById('booking-flow');
    const selectionsHeading = document.querySelector('.booking-page .form-section h2');
    const backLink = document.querySelector('.booking-page .back-link');
    const tagsContainer = document.getElementById('selected-services');
    const successBlock = document.getElementById('booking-success');
    const successTitle = successBlock ? successBlock.querySelector('h3') : null;
    const successMessage = document.getElementById('success-message');

    let currentYear;
    let currentMonth;
    let selectedDate = null;
    let selectedSlot = null;
    let busyTimes = [];

    const today = new Date();
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const minBookableStart = new Date(todayStart);
    minBookableStart.setDate(minBookableStart.getDate() + 2);
    /** Last bookable calendar day: same date one month out (e.g. Apr 3 → May 3). */
    const maxBookableEnd = new Date(todayStart);
    maxBookableEnd.setMonth(maxBookableEnd.getMonth() + 1);
    const maxBookableYmd = `${maxBookableEnd.getFullYear()}-${String(maxBookableEnd.getMonth() + 1).padStart(2, '0')}-${String(maxBookableEnd.getDate()).padStart(2, '0')}`;
    currentYear = today.getFullYear();
    currentMonth = today.getMonth();

    function stripTimeLocal(d) {
        return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
    function startOfSundayWeek(d) {
        const x = stripTimeLocal(d);
        x.setDate(x.getDate() - x.getDay());
        return x;
    }
    function parseYmdToLocalDate(ymd) {
        const p = String(ymd).split('-').map(Number);
        if (p.length !== 3 || p.some((n) => !Number.isFinite(n))) return stripTimeLocal(todayStart);
        return new Date(p[0], p[1] - 1, p[2]);
    }
    const minWeekSunday = startOfSundayWeek(minBookableStart);
    const maxWeekSunday = startOfSundayWeek(parseYmdToLocalDate(maxBookableYmd));
    let weekViewSunday = new Date(minWeekSunday.getTime());
    let calViewMode = 'month';

    /** Calendar month index for comparisons (year * 12 + month 0–11). */
    function monthKey(year, month0) {
        return year * 12 + month0;
    }

    /** Earliest month in the picker; latest month is the one that contains maxBookableYmd. */
    const bookingWindowMinKey = monthKey(today.getFullYear(), today.getMonth());
    const bookingWindowMaxKey = monthKey(maxBookableEnd.getFullYear(), maxBookableEnd.getMonth());

    function addCalendarMonths(year, month0, delta) {
        const d = new Date(year, month0 + delta, 1);
        return { year: d.getFullYear(), month: d.getMonth() };
    }

    function updateCalNavState() {
        if (!prevBtn || !nextBtn) return;
        if (calViewMode === 'week') {
            const atMin = weekViewSunday.getTime() <= minWeekSunday.getTime();
            const atMax = weekViewSunday.getTime() >= maxWeekSunday.getTime();
            prevBtn.disabled = atMin;
            nextBtn.disabled = atMax;
            prevBtn.setAttribute('aria-disabled', atMin ? 'true' : 'false');
            nextBtn.setAttribute('aria-disabled', atMax ? 'true' : 'false');
        } else {
            const vk = monthKey(currentYear, currentMonth);
            const atMin = vk <= bookingWindowMinKey;
            const atMax = vk >= bookingWindowMaxKey;
            prevBtn.disabled = atMin;
            nextBtn.disabled = atMax;
            prevBtn.setAttribute('aria-disabled', atMin ? 'true' : 'false');
            nextBtn.setAttribute('aria-disabled', atMax ? 'true' : 'false');
        }
    }

    const calGrid = document.getElementById('cal-grid');
    const calPaneMonth = document.getElementById('cal-pane-month');
    const calPaneWeek = document.getElementById('cal-pane-week');
    const calWeekGrid = document.getElementById('cal-week-grid');
    const btnCalViewMonth = document.getElementById('cal-view-month');
    const btnCalViewWeek = document.getElementById('cal-view-week');
    const monthLabel = document.getElementById('cal-month-label');
    const prevBtn = document.getElementById('cal-prev');
    const nextBtn = document.getElementById('cal-next');
    const slotsSection = document.getElementById('slots-section');
    const slotsDateLabel = document.getElementById('slots-date-label');
    const slotsContainer = document.getElementById('slots-container');
    const customerSection = document.getElementById('customer-section');
    const confirmBtn = document.getElementById('confirm-btn');
    const bookError = document.getElementById('book-error');
    const nameInput = document.getElementById('cust-name');
    const phoneInput = document.getElementById('cust-phone');
    const emailInput = document.getElementById('cust-email');
    const notesInput = document.getElementById('cust-notes');

    const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];
    const DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    function loadCalendarJsonp(baseUrl, ignoreEventId) {
        return new Promise((resolve, reject) => {
            const callbackName = 'processCalendarData_' + Date.now();
            const timeoutId = setTimeout(() => {
                cleanup();
                reject(new Error('jsonp timeout'));
            }, JSONP_TIMEOUT_MS);

            function cleanup() {
                clearTimeout(timeoutId);
                try {
                    delete window[callbackName];
                } catch (e) { /* ignore */ }
                const tag = document.getElementById('jsonp-script');
                if (tag) tag.remove();
            }

            window[callbackName] = function (data) {
                cleanup();
                busyTimes = Array.isArray(data) ? data : [];
                resolve();
            };

            const qs = new URLSearchParams();
            qs.set('callback', callbackName);
            if (ignoreEventId) qs.set('ignoreEventId', ignoreEventId);

            const script = document.createElement('script');
            script.id = 'jsonp-script';
            script.src = `${baseUrl}?${qs.toString()}`;
            script.onerror = () => {
                cleanup();
                reject(new Error('jsonp load error'));
            };
            document.body.appendChild(script);
        });
    }

    function loadRescheduleMetaJsonp(baseUrl, eventId, token) {
        return new Promise((resolve, reject) => {
            const callbackName = 'rescheduleMeta_' + Date.now();
            const timeoutId = setTimeout(() => {
                cleanup();
                reject(new Error('reschedule meta timeout'));
            }, JSONP_TIMEOUT_MS);

            function cleanup() {
                clearTimeout(timeoutId);
                try {
                    delete window[callbackName];
                } catch (e) { /* ignore */ }
                const tag = document.getElementById('jsonp-reschedule-meta');
                if (tag) tag.remove();
            }

            window[callbackName] = function (payload) {
                cleanup();
                resolve(payload);
            };

            const qs = new URLSearchParams({
                callback: callbackName,
                action: 'reschedule_meta',
                eventId: eventId,
                token: token,
            });
            const script = document.createElement('script');
            script.id = 'jsonp-reschedule-meta';
            script.src = `${baseUrl}?${qs.toString()}`;
            script.onerror = () => {
                cleanup();
                reject(new Error('reschedule meta load error'));
            };
            document.body.appendChild(script);
        });
    }

    async function fetchBusyTimes(ignoreEventId) {
        try {
            await loadJsonpTwice_(() => loadCalendarJsonp(SCRIPT_PROXY, ignoreEventId));
        } catch (e) {
            console.warn('Proxy /api/booking failed; loading calendar from Google directly.', e);
            try {
                await loadJsonpTwice_(() => loadCalendarJsonp(SCRIPT_DIRECT, ignoreEventId));
            } catch (e2) {
                console.error('Calendar JSONP failed entirely.', e2);
                busyTimes = [];
            }
        }
    }

    const GAP_THRESHOLD = 60;
    /** Start times offered on the hour grid (15-minute steps: :00, :15, :30, :45). */
    const SLOT_START_MINUTES = [0, 15, 30, 45];

    /** Local wall time — avoids new Date('yyyy-mm-ddTHH:mm:ss') UTC vs local bugs (Safari / email vs calendar skew). */
    function localWallDateTime(dateStr, hour, minute) {
        const p = dateStr.split('-').map(Number);
        if (p.length !== 3 || p.some((n) => !Number.isFinite(n))) return new Date(NaN);
        return new Date(p[0], p[1] - 1, p[2], hour, minute, 0);
    }

    function isOptimalSlot(slotStart, slotEnd, dayOfWeek, dateStr) {
        const hours = dayHours(dayOfWeek);
        if (!hours) return false;
        const workStart = localWallDateTime(dateStr, hours.start, 0);
        const workEnd = localWallDateTime(dateStr, hours.end, 0);

        let prevEnd = workStart;
        busyTimes.forEach(busy => {
            const bEnd = new Date(busy.end);
            if (bEnd <= slotStart && bEnd > prevEnd) {
                prevEnd = bEnd;
            }
        });

        let nextStart = workEnd;
        busyTimes.forEach(busy => {
            const bStart = new Date(busy.start);
            if (bStart >= slotEnd && bStart < nextStart) {
                nextStart = bStart;
            }
        });

        const gapBefore = (slotStart - prevEnd) / 60000;
        const gapAfter = (nextStart - slotEnd) / 60000;

        if (gapBefore > 30 && gapBefore < GAP_THRESHOLD) return false;
        if (gapAfter > 30 && gapAfter < GAP_THRESHOLD) return false;

        return true;
    }

    function hasAnyAvailableSlot(dateStr, dayOfWeek) {
        const p = dateStr.split('-').map(Number);
        const slotDayStart = new Date(p[0], p[1] - 1, p[2]);
        if (slotDayStart < minBookableStart) return false;
        if (dateStr > maxBookableYmd) return false;
        const hours = dayHours(dayOfWeek);
        if (!hours) return false;
        for (let h = hours.start; h < hours.end; h++) {
            for (const m of SLOT_START_MINUTES) {
                const slotStart = localWallDateTime(dateStr, h, m);
                const slotEnd = new Date(slotStart.getTime() + serviceMinutes * 60000);

                const isConflict = busyTimes.some(busy => {
                    const bStart = new Date(busy.start);
                    const bEnd = new Date(busy.end);
                    return (slotStart < bEnd && slotEnd > bStart);
                });

                const closingDateTime = localWallDateTime(dateStr, hours.end, 0);
                const isOverClosing = slotEnd > closingDateTime;

                if (!isConflict && !isOverClosing && !isSlotStartInPastForToday(dateStr, slotStart)) {
                    return true;
                }
            }
        }
        return false;
    }

    function mergeBusyIntervals(intervals) {
        if (!intervals.length) return [];
        const sorted = intervals.slice().sort((a, b) => a.start - b.start);
        const out = [{ start: sorted[0].start, end: sorted[0].end }];
        for (let i = 1; i < sorted.length; i++) {
            const cur = sorted[i];
            const last = out[out.length - 1];
            if (cur.start <= last.end) last.end = Math.max(last.end, cur.end);
            else out.push({ start: cur.start, end: cur.end });
        }
        return out;
    }

    function busySegmentsForDay(dateStr) {
        const day0 = localWallDateTime(dateStr, 0, 0);
        const day1 = new Date(day0.getTime() + 86400000);
        const out = [];
        busyTimes.forEach(b => {
            const bs = new Date(b.start);
            const be = new Date(b.end);
            const s = Math.max(bs.getTime(), day0.getTime());
            const e = Math.min(be.getTime(), day1.getTime());
            if (e > s) out.push({ start: s, end: e });
        });
        return mergeBusyIntervals(out);
    }

    function weekTimelineHourBounds(weekSunday) {
        let minH = 24;
        let maxH = 0;
        for (let i = 0; i < 7; i++) {
            const d = new Date(weekSunday);
            d.setDate(d.getDate() + i);
            const dateStr = formatDate(d);
            const wh = dayHours(d.getDay());
            if (wh) {
                minH = Math.min(minH, wh.start);
                maxH = Math.max(maxH, wh.end);
            }
            busySegmentsForDay(dateStr).forEach(seg => {
                const s = new Date(seg.start);
                const e = new Date(seg.end);
                minH = Math.min(minH, s.getHours());
                const endH = e.getHours() + (e.getMinutes() > 0 || e.getSeconds() > 0 ? 1 : 0);
                maxH = Math.max(maxH, endH);
            });
        }
        if (minH >= maxH) {
            minH = 8;
            maxH = 19;
        }
        minH = Math.max(0, minH - 1);
        maxH = Math.min(24, maxH + 1);
        return { startHour: minH, endHour: maxH };
    }

    function blockLayoutMs(dateStr, segStartMs, segEndMs, startHour, endHour, totalPx) {
        const t0 = localWallDateTime(dateStr, startHour, 0).getTime();
        const t1 = localWallDateTime(dateStr, endHour, 0).getTime();
        const span = t1 - t0;
        if (span <= 0) return null;
        const s = Math.max(segStartMs, t0);
        const e = Math.min(segEndMs, t1);
        if (e <= s) return null;
        const top = ((s - t0) / span) * totalPx;
        const h = ((e - s) / span) * totalPx;
        return { top, height: Math.max(3, h) };
    }

    function formatHourLabel12(hour) {
        const h12 = hour % 12 || 12;
        const ap = hour < 12 ? 'AM' : 'PM';
        return `${h12} ${ap}`;
    }

    function formatWeekRangeLabel() {
        const end = new Date(weekViewSunday);
        end.setDate(end.getDate() + 6);
        const y = weekViewSunday.getFullYear();
        const o = { month: 'short', day: 'numeric' };
        if (end.getFullYear() !== y) {
            return `${weekViewSunday.toLocaleDateString('en-US', { ...o, year: 'numeric' })} – ${end.toLocaleDateString('en-US', { ...o, year: 'numeric' })}`;
        }
        return `${weekViewSunday.toLocaleDateString('en-US', o)} – ${end.toLocaleDateString('en-US', o)}, ${y}`;
    }

    function setCalView(mode) {
        calViewMode = mode === 'week' ? 'week' : 'month';
        if (btnCalViewMonth) btnCalViewMonth.classList.toggle('is-active', calViewMode === 'month');
        if (btnCalViewWeek) btnCalViewWeek.classList.toggle('is-active', calViewMode === 'week');
        if (calPaneMonth) calPaneMonth.hidden = calViewMode !== 'month';
        if (calPaneWeek) calPaneWeek.hidden = calViewMode !== 'week';
        if (calViewMode === 'week') {
            if (selectedDate) {
                weekViewSunday = startOfSundayWeek(parseYmdToLocalDate(selectedDate));
                if (weekViewSunday.getTime() < minWeekSunday.getTime()) {
                    weekViewSunday = new Date(minWeekSunday.getTime());
                }
                if (weekViewSunday.getTime() > maxWeekSunday.getTime()) {
                    weekViewSunday = new Date(maxWeekSunday.getTime());
                }
            } else {
                weekViewSunday = new Date(minWeekSunday.getTime());
            }
        }
        refreshCalendar();
    }

    function renderMonthCalendar() {
        if (!calGrid) return;
        calGrid.innerHTML = '';
        monthLabel.textContent = `${MONTHS[currentMonth]} ${currentYear}`;
        DAYS.forEach(d => {
            const lbl = document.createElement('div');
            lbl.className = 'cal-day-label';
            lbl.textContent = d;
            calGrid.appendChild(lbl);
        });

        const firstDay = new Date(currentYear, currentMonth, 1).getDay();
        const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();

        for (let i = 0; i < firstDay; i++) {
            const empty = document.createElement('div');
            empty.className = 'cal-day empty';
            calGrid.appendChild(empty);
        }

        for (let day = 1; day <= daysInMonth; day++) {
            const btn = document.createElement('button');
            btn.type = 'button';
            const dateObj = new Date(currentYear, currentMonth, day);
            const dayOfWeek = dateObj.getDay();
            const dateStr = formatDate(dateObj);

            btn.className = 'cal-day';
            btn.textContent = String(day);

            if (dateObj < todayStart) {
                btn.classList.add('past');
            } else if (dateObj < minBookableStart) {
                btn.classList.add('unavailable');
            } else if (dateStr > maxBookableYmd) {
                btn.classList.add('unavailable');
            } else if (!dayHours(dayOfWeek)) {
                btn.classList.add('unavailable');
            } else {
                if (!hasAnyAvailableSlot(dateStr, dayOfWeek)) {
                    btn.classList.add('unavailable');
                } else {
                    btn.classList.add('available');
                    if (selectedDate === dateStr) btn.classList.add('selected');
                    btn.addEventListener('click', () => selectDate(dateStr, btn, dayOfWeek));
                }
            }
            if (dateStr === formatDate(today) && !btn.classList.contains('past')) {
                btn.classList.add('today');
            }
            calGrid.appendChild(btn);
        }
        updateCalNavState();
    }

    function renderWeekCalendar() {
        if (!calWeekGrid) return;
        calWeekGrid.innerHTML = '';
        monthLabel.textContent = formatWeekRangeLabel();

        const { startHour, endHour } = weekTimelineHourBounds(weekViewSunday);
        const PX_PER_H = 52;
        const bodyHeight = (endHour - startHour) * PX_PER_H;
        const todayStr = formatDate(today);

        const wrap = document.createElement('div');
        wrap.className = 'cal-week-timeline-wrap';

        const scroll = document.createElement('div');
        scroll.className = 'cal-week-timeline-scroll';

        const headerRow = document.createElement('div');
        headerRow.className = 'cal-week-timeline-header-row';
        const corner = document.createElement('div');
        corner.className = 'cal-week-timeline-corner';
        corner.setAttribute('aria-hidden', 'true');
        headerRow.appendChild(corner);

        for (let i = 0; i < 7; i++) {
            const d = new Date(weekViewSunday);
            d.setDate(d.getDate() + i);
            const dateStr = formatDate(d);
            const dow = d.getDay();
            const hcell = document.createElement('div');
            hcell.className = 'cal-week-head-cell';
            if (dow === 0 || dow === 6) hcell.classList.add('is-weekend');
            if (dateStr === todayStr) hcell.classList.add('is-today-head');
            const dowSpan = document.createElement('span');
            dowSpan.className = 'cal-week-head-dow';
            dowSpan.textContent = DAYS[dow];
            const domSpan = document.createElement('span');
            domSpan.className = 'cal-week-head-dom';
            domSpan.textContent = String(d.getDate());
            hcell.appendChild(dowSpan);
            hcell.appendChild(domSpan);
            headerRow.appendChild(hcell);
        }
        scroll.appendChild(headerRow);

        const body = document.createElement('div');
        body.className = 'cal-week-timeline-body';

        const timeGutter = document.createElement('div');
        timeGutter.className = 'cal-week-time-gutter';
        timeGutter.style.height = `${bodyHeight}px`;
        for (let h = startHour; h < endHour; h++) {
            const lab = document.createElement('div');
            lab.className = 'cal-week-time-label';
            lab.style.top = `${(h - startHour) * PX_PER_H}px`;
            lab.textContent = formatHourLabel12(h);
            timeGutter.appendChild(lab);
        }
        body.appendChild(timeGutter);

        const colsWrap = document.createElement('div');
        colsWrap.className = 'cal-week-cols-wrap';
        colsWrap.style.height = `${bodyHeight}px`;

        function appendDimLayer(layer, lay) {
            const dim = document.createElement('div');
            dim.className = 'cal-week-dim';
            dim.style.top = `${lay.top}px`;
            dim.style.height = `${lay.height}px`;
            layer.appendChild(dim);
        }

        for (let i = 0; i < 7; i++) {
            const d = new Date(weekViewSunday);
            d.setDate(d.getDate() + i);
            const dateStr = formatDate(d);
            const dayOfWeek = d.getDay();
            const wh = dayHours(dayOfWeek);

            const col = document.createElement('div');
            col.className = 'cal-week-col';
            col.dataset.calDate = dateStr;
            if (selectedDate === dateStr) col.classList.add('is-selected');

            const hit = document.createElement('button');
            hit.type = 'button';
            hit.className = 'cal-week-col-hit';
            hit.style.minHeight = `${bodyHeight}px`;

            const colInner = document.createElement('div');
            colInner.className = 'cal-week-col-inner';
            colInner.style.height = `${bodyHeight}px`;

            const gridBg = document.createElement('div');
            gridBg.className = 'cal-week-col-grid';
            gridBg.style.height = `${bodyHeight}px`;
            for (let h = startHour; h < endHour; h++) {
                const line = document.createElement('div');
                line.className = 'cal-week-hour-line';
                line.style.top = `${(h - startHour) * PX_PER_H}px`;
                gridBg.appendChild(line);
            }

            const layer = document.createElement('div');
            layer.className = 'cal-week-busy-layer';
            layer.style.height = `${bodyHeight}px`;

            if (wh) {
                const t0 = localWallDateTime(dateStr, startHour, 0).getTime();
                const t1 = localWallDateTime(dateStr, endHour, 0).getTime();
                const tOpen = localWallDateTime(dateStr, wh.start, 0).getTime();
                const tClose = localWallDateTime(dateStr, wh.end, 0).getTime();
                if (tOpen > t0) {
                    const lay = blockLayoutMs(dateStr, t0, tOpen, startHour, endHour, bodyHeight);
                    if (lay) appendDimLayer(layer, lay);
                }
                if (tClose < t1) {
                    const lay = blockLayoutMs(dateStr, tClose, t1, startHour, endHour, bodyHeight);
                    if (lay) appendDimLayer(layer, lay);
                }
            } else {
                const dimFull = document.createElement('div');
                dimFull.className = 'cal-week-dim cal-week-dim-full';
                dimFull.style.height = `${bodyHeight}px`;
                layer.appendChild(dimFull);
            }

            busySegmentsForDay(dateStr).forEach(seg => {
                const lay = blockLayoutMs(dateStr, seg.start, seg.end, startHour, endHour, bodyHeight);
                if (!lay) return;
                const blk = document.createElement('div');
                blk.className = 'cal-week-busy-block';
                blk.style.top = `${lay.top}px`;
                blk.style.height = `${lay.height}px`;
                blk.title = 'Booked / unavailable';
                layer.appendChild(blk);
            });

            const selectable =
                stripTimeLocal(d) >= todayStart &&
                stripTimeLocal(d) >= minBookableStart &&
                dateStr <= maxBookableYmd &&
                wh &&
                hasAnyAvailableSlot(dateStr, dayOfWeek);

            if (!selectable) {
                hit.disabled = true;
                hit.classList.add('is-disabled');
                col.classList.add('is-disabled');
            }
            hit.setAttribute(
                'aria-label',
                selectable ? `Select ${dateStr} to see open times` : `${dateStr} is not available to book`
            );

            colInner.appendChild(gridBg);
            colInner.appendChild(layer);
            hit.appendChild(colInner);
            hit.addEventListener('click', () => {
                if (!hit.disabled) selectDate(dateStr, null, dayOfWeek);
            });

            col.appendChild(hit);
            colsWrap.appendChild(col);
        }

        let nowCol = -1;
        for (let i = 0; i < 7; i++) {
            const d = new Date(weekViewSunday);
            d.setDate(d.getDate() + i);
            if (formatDate(d) === todayStr) {
                nowCol = i;
                break;
            }
        }
        if (nowCol >= 0) {
            const t0 = localWallDateTime(todayStr, startHour, 0).getTime();
            const t1 = localWallDateTime(todayStr, endHour, 0).getTime();
            const nowMs = Date.now();
            if (nowMs >= t0 && nowMs <= t1) {
                const top = ((nowMs - t0) / (t1 - t0)) * bodyHeight;
                const line = document.createElement('div');
                line.className = 'cal-week-now-line';
                line.style.top = `${top}px`;
                line.style.left = `${(100 / 7) * nowCol}%`;
                line.style.width = `${100 / 7}%`;
                colsWrap.appendChild(line);
            }
        }

        body.appendChild(colsWrap);
        scroll.appendChild(body);
        wrap.appendChild(scroll);
        calWeekGrid.appendChild(wrap);

        updateCalNavState();
    }

    function refreshCalendar() {
        if (calViewMode === 'week') renderWeekCalendar();
        else renderMonthCalendar();
    }

    function formatDate(d) {
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }

    function isDateStrToday(dateStr) {
        return dateStr === formatDate(new Date());
    }

    /** Today only: start time is already at or before now (local) — not bookable. */
    function isSlotStartInPastForToday(dateStr, slotStart) {
        if (!isDateStrToday(dateStr)) return false;
        return slotStart.getTime() <= Date.now();
    }

    if (btnCalViewMonth) {
        btnCalViewMonth.addEventListener('click', () => setCalView('month'));
    }
    if (btnCalViewWeek) {
        btnCalViewWeek.addEventListener('click', () => setCalView('week'));
    }

    prevBtn.addEventListener('click', () => {
        if (calViewMode === 'week') {
            const n = new Date(weekViewSunday);
            n.setDate(n.getDate() - 7);
            if (n.getTime() < minWeekSunday.getTime()) return;
            weekViewSunday = n;
            renderWeekCalendar();
        } else {
            const prev = addCalendarMonths(currentYear, currentMonth, -1);
            if (monthKey(prev.year, prev.month) < bookingWindowMinKey) return;
            currentYear = prev.year;
            currentMonth = prev.month;
            renderMonthCalendar();
        }
    });

    nextBtn.addEventListener('click', () => {
        if (calViewMode === 'week') {
            const n = new Date(weekViewSunday);
            n.setDate(n.getDate() + 7);
            if (n.getTime() > maxWeekSunday.getTime()) return;
            weekViewSunday = n;
            renderWeekCalendar();
        } else {
            const nxt = addCalendarMonths(currentYear, currentMonth, 1);
            if (monthKey(nxt.year, nxt.month) > bookingWindowMaxKey) return;
            currentYear = nxt.year;
            currentMonth = nxt.month;
            renderMonthCalendar();
        }
    });

    function selectDate(dateStr, btnEl, dayOfWeek) {
        const ae = document.activeElement;
        if (ae && typeof ae.blur === 'function') ae.blur();

        selectedDate = dateStr;
        selectedSlot = null;
        customerSection.style.display = 'none';
        confirmBtn.disabled = true;
        document.querySelectorAll('.cal-day.selected').forEach(el => el.classList.remove('selected'));
        if (calWeekGrid) {
            calWeekGrid.querySelectorAll('.cal-week-col.is-selected').forEach(el => el.classList.remove('is-selected'));
            const wcol = calWeekGrid.querySelector(`.cal-week-col[data-cal-date="${dateStr}"]`);
            if (wcol) wcol.classList.add('is-selected');
        }
        if (btnEl) btnEl.classList.add('selected');

        slotsSection.style.display = 'block';
        const dateObj = localWallDateTime(dateStr, 12, 0);
        slotsDateLabel.textContent = dateObj.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });

        requestAnimationFrame(() => {
            renderTimeSlots(dateStr, dayOfWeek);
            setTimeout(() => {
                slotsSection.scrollIntoView({ behavior: 'auto', block: 'start' });
            }, 50);
        });
    }

    function renderTimeSlots(dateStr, dayOfWeek) {
        slotsContainer.innerHTML = '';
        const grid = document.createElement('div');
        grid.className = 'slots-grid';

        if (dateStr > maxBookableYmd) {
            slotsContainer.innerHTML = '';
            return;
        }

        const hours = dayHours(dayOfWeek);
        if (!hours) {
            slotsContainer.innerHTML = '';
            return;
        }
        for (let h = hours.start; h < hours.end; h++) {
            for (const m of SLOT_START_MINUTES) {
                const timeStr12 = format12h(h, m);
                const slotStart = localWallDateTime(dateStr, h, m);
                const slotEnd = new Date(slotStart.getTime() + serviceMinutes * 60000);

                const isConflict = busyTimes.some(busy => {
                    const bStart = new Date(busy.start);
                    const bEnd = new Date(busy.end);
                    return (slotStart < bEnd && slotEnd > bStart);
                });

                const closingDateTime = localWallDateTime(dateStr, hours.end, 0);
                const isOverClosing = slotEnd > closingDateTime;
                const isPastToday = isSlotStartInPastForToday(dateStr, slotStart);
                const isOptimal = isOptimalSlot(slotStart, slotEnd, dayOfWeek, dateStr);

                const pill = document.createElement('button');
                pill.className = 'slot-pill';
                pill.textContent = timeStr12;

                if (isConflict || isOverClosing || isPastToday) {
                    pill.classList.add('busy');
                    pill.style.textDecoration = 'line-through';
                    pill.style.opacity = '0.35';
                    pill.disabled = true;
                } else {
                    if (!isOptimal) {
                        pill.classList.add('slot-pill-suboptimal');
                        pill.style.opacity = '0.72';
                    }
                    pill.addEventListener('click', () => {
                        selectedSlot = timeStr12;
                        document.querySelectorAll('.slot-pill.selected').forEach(p => p.classList.remove('selected'));
                        pill.classList.add('selected');
                        showCustomerSection();
                    });
                }
                grid.appendChild(pill);
            }
        }
        slotsContainer.appendChild(grid);
    }

    function format12h(h, m) {
        const ampm = h >= 12 ? 'PM' : 'AM';
        let h12 = h % 12;
        if (h12 === 0) h12 = 12;
        return `${h12}:${String(m).padStart(2, '0')} ${ampm}`;
    }

    function showCustomerSection() {
        document.getElementById('slot-preview').innerHTML = `Selected: <strong>${slotsDateLabel.textContent} @ ${selectedSlot}</strong>`;
        customerSection.style.display = 'block';
        updateConfirmBtn();
    }

    function updateConfirmBtn() {
        if (confirmBtn.classList.contains('is-submitting')) return;
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        const phoneRegex = /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/im;

        const isNameValid = nameInput.value.trim().length > 0;
        const isEmailValid = emailRegex.test(emailInput.value.trim());
        const phoneVal = phoneInput.value.trim();
        const isPhoneValid = isReschedule
            ? phoneVal.replace(/\D/g, '').length >= 10
            : phoneRegex.test(phoneVal);

        confirmBtn.disabled = !isNameValid || !isEmailValid || !isPhoneValid || !selectedSlot;
    }

    nameInput.addEventListener('input', updateConfirmBtn);
    phoneInput.addEventListener('input', updateConfirmBtn);
    emailInput.addEventListener('input', updateConfirmBtn);

    let bookingSubmitInFlight = false;

    confirmBtn.addEventListener('click', async () => {
        if (bookingSubmitInFlight) return;
        if (!selectedSlot || !nameInput.value.trim()) return;

        bookingSubmitInFlight = true;
        const originalBtnText = confirmBtn.innerHTML;
        confirmBtn.classList.add('is-submitting');
        confirmBtn.setAttribute('aria-busy', 'true');
        confirmBtn.innerHTML = '<span class="spinner" aria-hidden="true"></span>';
        bookError.innerHTML = '';

        const bookingData = isReschedule
            ? {
                reschedule: true,
                eventId: rescheduleEventId,
                token: rescheduleToken,
                date: selectedDate,
                time: selectedSlot,
            }
            : {
                clientName: nameInput.value.trim(),
                phone: phoneInput.value.trim(),
                email: emailInput.value.trim(),
                service: serviceNames,
                date: selectedDate,
                time: selectedSlot,
                durationMinutes: serviceMinutes,
                clientNotes: notesInput && notesInput.value ? notesInput.value.trim() : '',
            };

        try {
            const res = await fetch(SCRIPT_PROXY, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(bookingData),
            });
            const data = await res.json().catch(() => ({}));
            if (!res.ok || data.status !== 'success') {
                const err = new Error(data.message || data.status || 'failed');
                err.apiMessage = data.message;
                throw err;
            }

            document.getElementById('booking-flow').style.display = 'none';
            document.querySelector('.booking-page .form-section:first-of-type').style.display = 'none';
            if (successBlock) successBlock.style.display = 'block';
            if (isReschedule && successTitle) {
                successTitle.textContent = 'Appointment rescheduled';
                if (successMessage) {
                    successMessage.textContent = 'Check your email for the updated date and time.';
                    successMessage.style.display = 'block';
                }
            }
            const ae = document.activeElement;
            if (ae && typeof ae.blur === 'function') ae.blur();
            requestAnimationFrame(() => {
                if (successBlock) {
                    successBlock.scrollIntoView({ block: 'start', behavior: 'instant' });
                }
            });
        } catch (err) {
            const apiMsg = err && err.apiMessage;
            const msg =
                apiMsg === 'date_outside_booking_window'
                    ? 'That date is outside the booking window. You can book from two days ahead through one month from today—refresh and choose an open date.'
                    : apiMsg === 'date_too_soon'
                      ? 'Appointments must be booked at least two days in advance (not today or tomorrow). Please pick a later date.'
                      : apiMsg === 'invalid_day_or_time'
                        ? 'That day or time isn’t available. Refresh the page and choose an open slot from the calendar.'
                        : 'Submission failed. Please try again.';
            bookError.innerHTML = `<div class="booking-error">${msg}</div>`;
            confirmBtn.classList.remove('is-submitting');
            confirmBtn.removeAttribute('aria-busy');
            confirmBtn.innerHTML = originalBtnText;
            bookingSubmitInFlight = false;
            updateConfirmBtn();
        }
    });

    function populateServiceTags(namesBlob) {
        tagsContainer.innerHTML = '';
        if (!namesBlob) return;
        String(namesBlob).split(',').forEach(name => {
            const t = name.trim();
            if (!t) return;
            const tag = document.createElement('span');
            tag.className = 'service-tag';
            tag.textContent = safeDecodeURIComponent(t);
            tagsContainer.appendChild(tag);
        });
    }

    async function initBooking() {
        let ignoreEventId = '';

        if (isReschedule) {
            if (!rescheduleEventId || !rescheduleToken) {
                bookingFlow.innerHTML =
                    '<p style="font-style:italic;color:var(--text-muted);text-align:center;padding:2rem;">This reschedule link is invalid or incomplete. Use the link from your reminder email, or <a href="index.html">book from the site</a>.</p>';
                return;
            }
            if (selectionsHeading) selectionsHeading.textContent = 'Reschedule — your services';
            if (backLink) {
                backLink.textContent = 'Back to home';
                backLink.href = 'index.html';
            }
            let meta;
            try {
                meta = await loadJsonpTwice_(() => loadRescheduleMetaJsonp(SCRIPT_PROXY, rescheduleEventId, rescheduleToken));
            } catch (e1) {
                try {
                    meta = await loadJsonpTwice_(() => loadRescheduleMetaJsonp(SCRIPT_DIRECT, rescheduleEventId, rescheduleToken));
                } catch (e2) {
                    meta = null;
                }
            }
            if (!meta || !meta.ok) {
                bookingFlow.innerHTML =
                    '<p style="font-style:italic;color:var(--text-muted);text-align:center;padding:2rem;">We could not load this appointment. It may already be cancelled or the link expired. <a href="index.html">Return home</a>.</p>';
                return;
            }
            serviceIds = 'reschedule';
            serviceNames = meta.service;
            serviceMinutes = Math.max(30, Number(meta.durationMinutes) || 60);
            nameInput.value = meta.clientName || '';
            phoneInput.value = meta.phone || '';
            emailInput.value = meta.email || '';
            nameInput.readOnly = true;
            phoneInput.readOnly = true;
            emailInput.readOnly = true;
            nameInput.style.opacity = '0.85';
            phoneInput.style.opacity = '0.85';
            emailInput.style.opacity = '0.85';
            confirmBtn.textContent = 'Confirm new time';
            const notesFieldEl = document.getElementById('client-notes-field');
            if (notesFieldEl) notesFieldEl.style.display = 'none';
            // Do not strip this event from busy times: until they confirm, it still occupies the calendar.
            // Ignoring it made overlapping slots (e.g. 3:00 for a 75m appt after 2:30) look falsely open.
        } else if (!serviceIds) {
            bookingFlow.innerHTML =
                '<p style="font-style:italic;color:var(--text-muted); text-align:center; padding: 2rem;">No services selected. <a href="index.html">Go back</a> and choose your services first.</p>';
            return;
        }

        const loadingEl = document.getElementById('calendar-loading');

        try {
            populateServiceTags(serviceNames);
            monthLabel.textContent = `${MONTHS[currentMonth]} ${currentYear}`;

            if (priceStr && timeStr && !isReschedule) {
                const summaryInfo = document.createElement('div');
                summaryInfo.className = 'summary-info-lite';
                const p = safeDecodeURIComponent(priceStr);
                const t = safeDecodeURIComponent(timeStr);
                summaryInfo.innerHTML = `
            <p>Estimated Total: <strong>${p.replace(/</g, '&lt;')}</strong></p>
            <p>Estimated Time: <strong>${t.replace(/</g, '&lt;')}</strong></p>
        `;
                tagsContainer.after(summaryInfo);
            }

            await Promise.all([loadWorkHours(), fetchBusyTimes(ignoreEventId)]);
        } finally {
            refreshCalendar();
            if (loadingEl) {
                loadingEl.hidden = true;
                loadingEl.style.display = 'none';
                loadingEl.setAttribute('aria-hidden', 'true');
            }
        }
    }

    initBooking().catch(() => {
        bookingFlow.innerHTML =
            '<p style="text-align:center;padding:2rem;color:var(--text-muted);">Something went wrong loading the calendar. Please refresh or try again later.</p>';
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bootBookingPage);
} else {
    bootBookingPage();
}
