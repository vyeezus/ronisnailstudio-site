// --- Roni's Nail Studio — Custom Booking Calendar (reschedule + new booking) ---

document.addEventListener('DOMContentLoaded', () => {
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

    const WORK_HOURS = {
        1: { start: 11, end: 18 },
        2: { start: 11, end: 18 },
        3: { start: 9, end: 16 },
        5: { start: 9, end: 16 },
    };

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
    currentYear = today.getFullYear();
    currentMonth = today.getMonth();

    const calGrid = document.getElementById('cal-grid');
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

    const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];
    const DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    function loadCalendarJsonp(baseUrl, ignoreEventId) {
        return new Promise((resolve, reject) => {
            const callbackName = 'processCalendarData_' + Date.now();
            const timeoutId = setTimeout(() => {
                cleanup();
                reject(new Error('jsonp timeout'));
            }, 8000);

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
                renderCalendar();
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
            }, 8000);

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
            await loadCalendarJsonp(SCRIPT_PROXY, ignoreEventId);
        } catch (e) {
            console.warn('Proxy /api/booking failed; loading calendar from Google directly.', e);
            try {
                await loadCalendarJsonp(SCRIPT_DIRECT, ignoreEventId);
            } catch (e2) {
                console.error('Calendar JSONP failed entirely.', e2);
                busyTimes = [];
                renderCalendar();
            }
        }
    }

    const GAP_THRESHOLD = 60;

    function isOptimalSlot(slotStart, slotEnd, dayOfWeek, dateStr) {
        const hours = WORK_HOURS[dayOfWeek];
        const workStart = new Date(`${dateStr}T${String(hours.start).padStart(2, '0')}:00:00`);
        const workEnd = new Date(`${dateStr}T${String(hours.end).padStart(2, '0')}:00:00`);

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
        const hours = WORK_HOURS[dayOfWeek];
        for (let h = hours.start; h < hours.end; h++) {
            for (const m of [0, 30]) {
                const slotStart = new Date(`${dateStr}T${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:00`);
                const slotEnd = new Date(slotStart.getTime() + serviceMinutes * 60000);

                const isConflict = busyTimes.some(busy => {
                    const bStart = new Date(busy.start);
                    const bEnd = new Date(busy.end);
                    return (slotStart < bEnd && slotEnd > bStart);
                });

                const closingDateTime = new Date(`${dateStr}T${String(hours.end).padStart(2, '0')}:00:00`);
                const isOverClosing = slotEnd > closingDateTime;

                if (!isConflict && !isOverClosing) {
                    if (isOptimalSlot(slotStart, slotEnd, dayOfWeek, dateStr)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    function renderCalendar() {
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
            const dateObj = new Date(currentYear, currentMonth, day);
            const dayOfWeek = dateObj.getDay();
            const dateStr = formatDate(dateObj);

            btn.className = 'cal-day';
            btn.textContent = day;

            if (dateObj < new Date(today.getFullYear(), today.getMonth(), today.getDate())) {
                btn.classList.add('past');
            } else if (!WORK_HOURS[dayOfWeek]) {
                btn.classList.add('unavailable');
            } else {
                if (!hasAnyAvailableSlot(dateStr, dayOfWeek)) {
                    btn.classList.add('unavailable');
                } else {
                    btn.classList.add('available');
                    if (dateStr === formatDate(today)) btn.classList.add('today');
                    if (selectedDate === dateStr) btn.classList.add('selected');
                    btn.addEventListener('click', () => selectDate(dateStr, btn, dayOfWeek));
                }
            }
            calGrid.appendChild(btn);
        }
    }

    function formatDate(d) {
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }

    prevBtn.addEventListener('click', () => {
        currentMonth--;
        if (currentMonth < 0) { currentMonth = 11; currentYear--; }
        renderCalendar();
    });

    nextBtn.addEventListener('click', () => {
        currentMonth++;
        if (currentMonth > 11) { currentMonth = 0; currentYear++; }
        renderCalendar();
    });

    function selectDate(dateStr, btnEl, dayOfWeek) {
        selectedDate = dateStr;
        selectedSlot = null;
        customerSection.style.display = 'none';
        confirmBtn.disabled = true;

        document.querySelectorAll('.cal-day.selected').forEach(d => d.classList.remove('selected'));
        if (btnEl) btnEl.classList.add('selected');

        slotsSection.style.display = 'block';
        const dateObj = new Date(dateStr + 'T12:00:00');
        slotsDateLabel.textContent = dateObj.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });

        renderTimeSlots(dateStr, dayOfWeek);
    }

    function renderTimeSlots(dateStr, dayOfWeek) {
        slotsContainer.innerHTML = '';
        const grid = document.createElement('div');
        grid.className = 'slots-grid';

        const hours = WORK_HOURS[dayOfWeek];
        for (let h = hours.start; h < hours.end; h++) {
            for (const m of [0, 30]) {
                const timeStr12 = format12h(h, m);
                const slotStart = new Date(`${dateStr}T${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:00`);
                const slotEnd = new Date(slotStart.getTime() + serviceMinutes * 60000);

                const isConflict = busyTimes.some(busy => {
                    const bStart = new Date(busy.start);
                    const bEnd = new Date(busy.end);
                    return (slotStart < bEnd && slotEnd > bStart);
                });

                const closingDateTime = new Date(`${dateStr}T${String(hours.end).padStart(2, '0')}:00:00`);
                const isOverClosing = slotEnd > closingDateTime;
                const isOptimal = isOptimalSlot(slotStart, slotEnd, dayOfWeek, dateStr);

                const pill = document.createElement('button');
                pill.className = 'slot-pill';
                pill.textContent = timeStr12;

                if (isConflict || isOverClosing || !isOptimal) {
                    pill.classList.add('busy');
                    pill.style.textDecoration = 'line-through';
                    pill.style.opacity = '0.35';
                    pill.disabled = true;
                } else {
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
        customerSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        updateConfirmBtn();
    }

    function updateConfirmBtn() {
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

    confirmBtn.addEventListener('click', async () => {
        if (!selectedSlot || !nameInput.value.trim()) return;

        confirmBtn.disabled = true;
        const originalBtnText = confirmBtn.innerHTML;
        confirmBtn.innerHTML = '<div class="spinner" style="margin: 0 auto;"></div>';
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
            };

        try {
            const res = await fetch(SCRIPT_PROXY, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(bookingData),
            });
            const data = await res.json().catch(() => ({}));
            if (!res.ok || data.status !== 'success') {
                throw new Error(data.status || 'failed');
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
        } catch (err) {
            bookError.innerHTML = '<div class="booking-error">Submission failed. Please try again.</div>';
            confirmBtn.innerHTML = originalBtnText;
            confirmBtn.disabled = false;
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
            tag.textContent = decodeURIComponent(t);
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
            monthLabel.textContent = 'Loading…';

            let meta;
            try {
                meta = await loadRescheduleMetaJsonp(SCRIPT_PROXY, rescheduleEventId, rescheduleToken);
            } catch (e1) {
                try {
                    meta = await loadRescheduleMetaJsonp(SCRIPT_DIRECT, rescheduleEventId, rescheduleToken);
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
            // Do not strip this event from busy times: until they confirm, it still occupies the calendar.
            // Ignoring it made overlapping slots (e.g. 3:00 for a 75m appt after 2:30) look falsely open.
        } else if (!serviceIds) {
            bookingFlow.innerHTML =
                '<p style="font-style:italic;color:var(--text-muted); text-align:center; padding: 2rem;">No services selected. <a href="index.html">Go back</a> and choose your services first.</p>';
            return;
        }

        populateServiceTags(serviceNames);

        if (priceStr && timeStr && !isReschedule) {
            const summaryInfo = document.createElement('div');
            summaryInfo.className = 'summary-info-lite';
            summaryInfo.innerHTML = `
            <p>Estimated Total: <strong>${decodeURIComponent(priceStr)}</strong></p>
            <p>Estimated Time: <strong>${decodeURIComponent(timeStr)}</strong></p>
        `;
            tagsContainer.after(summaryInfo);
        }

        await fetchBusyTimes(ignoreEventId);
    }

    initBooking().catch(() => {
        bookingFlow.innerHTML =
            '<p style="text-align:center;padding:2rem;color:var(--text-muted);">Something went wrong loading the calendar. Please refresh or try again later.</p>';
    });
});
