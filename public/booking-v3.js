// --- Roni's Nail Studio — Custom Booking Calendar (V2 - FIX - Final Sync) ---

document.addEventListener('DOMContentLoaded', () => {
    // Parse URL params
    const params = new URLSearchParams(window.location.search);
    const serviceNames = params.get('names');
    const serviceIds = params.get('services');
    const priceStr = params.get('price');
    const timeStr = params.get('time');
    /** Netlify proxy (needs `public/_redirects` + functions). Falls back to Google for calendar JSONP if /api/booking 404s. */
    const SCRIPT_PROXY = `${window.location.origin}/api/booking`;
    const SCRIPT_DIRECT =
        'https://script.google.com/macros/s/AKfycbzdT_rV3dR7Th4VHeLE3uJcyTPr4bI-6uy-_Im6xz-nZ0rGPToj85zy7Is7LmpNVS0Wwg/exec';
    
    // Roni's Working Hours
    const WORK_HOURS = {
        1: { start: 11, end: 18 }, // Monday: 11am-6pm
        2: { start: 11, end: 18 }, // Tuesday: 11am-6pm
        3: { start: 9, end: 16 }, // Wednesday: 9am-4pm
        5: { start: 9, end: 16 }  // Friday: 9am-4pm
    };

    // Calculate Service Duration in Minutes - Much more aggressive regex
    let serviceMinutes = 60;
    console.log('Original Time String:', timeStr);
    if (timeStr) {
        const decodedTime = decodeURIComponent(timeStr).toLowerCase();
        console.log('Decoded Time:', decodedTime);

        // This looks for any number followed by 'h', then any number followed by 'm'
        const hMatch = decodedTime.match(/(\d+)\s*h/);
        const mMatch = decodedTime.match(/(\d+)\s*m/);

        let total = 0;
        if (hMatch) total += parseInt(hMatch[1]) * 60;
        if (mMatch) total += parseInt(mMatch[1]);
        if (total > 0) serviceMinutes = total;
    }
    console.log('Calculated Service Duration (min):', serviceMinutes);

    if (!serviceIds) {
        document.getElementById('booking-flow').innerHTML =
            '<p style="font-style:italic;color:var(--text-muted); text-align:center; padding: 2rem;">No services selected. <a href="index.html">Go back</a> and choose your services first.</p>';
        return;
    }

    const tagsContainer = document.getElementById('selected-services');
    if (serviceNames) {
        serviceNames.split(',').forEach(name => {
            const tag = document.createElement('span');
            tag.className = 'service-tag';
            tag.textContent = decodeURIComponent(name.trim());
            tagsContainer.appendChild(tag);
        });
    }

    if (priceStr && timeStr) {
        const summaryInfo = document.createElement('div');
        summaryInfo.className = 'summary-info-lite';
        summaryInfo.innerHTML = `
            <p>Estimated Total: <strong>${decodeURIComponent(priceStr)}</strong></p>
            <p>Estimated Time: <strong>${decodeURIComponent(timeStr)}</strong></p>
        `;
        tagsContainer.after(summaryInfo);
    }

    // State
    let currentYear, currentMonth;
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

    const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];
    const DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    function loadCalendarJsonp(baseUrl) {
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
                console.log('RECEIVED BUSY TIMES (JSONP) from', baseUrl, data);
                busyTimes = Array.isArray(data) ? data : [];
                renderCalendar();
                resolve();
            };

            const script = document.createElement('script');
            script.id = 'jsonp-script';
            script.src = `${baseUrl}?callback=${callbackName}`;
            script.onerror = () => {
                cleanup();
                reject(new Error('jsonp load error'));
            };
            document.body.appendChild(script);
        });
    }

    async function fetchBusyTimes() {
        console.log('--- FETCHING BUSY TIMES ---');
        try {
            await loadCalendarJsonp(SCRIPT_PROXY);
        } catch (e) {
            console.warn('Proxy /api/booking failed; loading calendar from Google directly.', e);
            try {
                await loadCalendarJsonp(SCRIPT_DIRECT);
            } catch (e2) {
                console.error('Calendar JSONP failed entirely.', e2);
                busyTimes = [];
                renderCalendar();
            }
        }
    }

    const GAP_THRESHOLD = 60; // 60 minutes buffer

    function isOptimalSlot(slotStart, slotEnd, dayOfWeek, dateStr) {
        const hours = WORK_HOURS[dayOfWeek];
        const workStart = new Date(`${dateStr}T${String(hours.start).padStart(2, '0')}:00:00`);
        const workEnd = new Date(`${dateStr}T${String(hours.end).padStart(2, '0')}:00:00`);

        let prevEnd = workStart;
        busyTimes.forEach(busy => {
            const bStart = new Date(busy.start);
            const bEnd = new Date(busy.end);
            // Must be an event that ends today, before or exactly when our slot starts
            if (bEnd <= slotStart && bEnd > prevEnd) {
                prevEnd = bEnd;
            }
        });

        let nextStart = workEnd;
        busyTimes.forEach(busy => {
            const bStart = new Date(busy.start);
            // Must be an event that starts today, after or exactly when our slot ends
            if (bStart >= slotEnd && bStart < nextStart) {
                nextStart = bStart;
            }
        });

        const gapBefore = (slotStart - prevEnd) / 60000;
        const gapAfter = (nextStart - slotEnd) / 60000;

        // "i think 15m and 30m gaps are ok but anything above that may be too much"
        // Meaning exactly 45 mins is our un-fillable "Swiss Cheese" gap. 
        // So we strictly forbid gaps > 30m and < 60m. 
        if (gapBefore > 30 && gapBefore < GAP_THRESHOLD) return false;
        if (gapAfter > 30 && gapAfter < GAP_THRESHOLD) return false;

        return true;
    }

    function hasAnyAvailableSlot(dateStr, dayOfWeek) {
        const hours = WORK_HOURS[dayOfWeek];
        for (let h = hours.start; h < hours.end; h++) {
            for (let m of [0, 30]) {
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
                        return true; // Found at least one working slot!
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
                // If there are zero viable slots, cross out the whole day
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
            for (let m of [0, 30]) {
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

    const phoneInput = document.getElementById('cust-phone');
    const emailInput = document.getElementById('cust-email');

    nameInput.addEventListener('input', updateConfirmBtn);
    phoneInput.addEventListener('input', updateConfirmBtn);
    emailInput.addEventListener('input', updateConfirmBtn);

    function updateConfirmBtn() {
        // Simple but highly effective Regex for standard emails
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        // Strong regex to accept various standard phone formats like (XXX) XXX-XXXX or XXXXXXXXXX
        const phoneRegex = /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/im;

        const isNameValid = nameInput.value.trim().length > 0;
        const isEmailValid = emailRegex.test(emailInput.value.trim());
        const isPhoneValid = phoneRegex.test(phoneInput.value.trim());

        confirmBtn.disabled = !isNameValid || !isEmailValid || !isPhoneValid || !selectedSlot;
    }

    confirmBtn.addEventListener('click', async () => {
        if (!selectedSlot || !nameInput.value.trim()) return;

        // Display sending state for better UX
        confirmBtn.disabled = true;
        const originalBtnText = confirmBtn.innerHTML;
        confirmBtn.innerHTML = '<div class="spinner" style="margin: 0 auto;"></div>';
        bookError.innerHTML = '';

        const bookingData = {
            clientName: nameInput.value.trim(),
            phone: document.getElementById('cust-phone').value.trim(),
            email: document.getElementById('cust-email').value.trim(),
            service: serviceNames,
            date: selectedDate,
            time: selectedSlot
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
            document.querySelector('.form-section:first-of-type').style.display = 'none';
            document.getElementById('booking-success').style.display = 'block';

        } catch (err) {
            bookError.innerHTML = `<div class="booking-error">Submission failed. Please try again.</div>`;
            confirmBtn.innerHTML = originalBtnText;
            confirmBtn.disabled = false;
        }
    });

    fetchBusyTimes();
});
