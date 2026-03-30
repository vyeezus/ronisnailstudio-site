// --- SQUARE BOOKING INTEGRATION SETTINGS ---
const SQUARE_BASE_URL = 'https://book.squareup.com/appointments/j6m5t42kcalyv4/location/LZHF8SWPV4XF9/services';

// Square Service Map — 1:1 with live Square menu
const SQUARE_SERVICE_MAP = {
    // Gel Overlay (Structure Gel)
    'sg-new': 'ET7ZTD5JVABTPXX5MGNKEYY3',       // Overlay (New Set) - $80
    'sg-refill': 'FQQ37Z4R2XSSXW5Y2JIPGWVH',     // Refill - $75
    'sg-removal-new': 'KZ7ALKL2DCYEAWL3BD6DFAKO', // Removal + New Set - $90

    // Gel Extensions (Gel X)
    'gx-short': 'ACY3LO4YLNVHCZJERWGG4R5E',           // Gel X Short - $75
    'gx-medium': 'NPHFBTB3UAZEEMACAX5KKKEU',          // Gel X Medium - $80
    'gx-long': 'VP7ZSRXCRGQGCX7QXRPTLHOW',            // Gel X Long - $90
    'gx-removal-short': 'Q3J63Y3FKJP5RPZ55TSURO6M',   // Removal + Gel X Short - $85
    'gx-removal-medium': 'B2GELF7B25XZVBXG5IZE6F2L',  // Removal + Gel X Medium - $90
    'gx-removal-long': 'C3LDH3JNIKUUAQN3IOIQK5YH',    // Removal + Gel X Long - $100

    // Design Tiers (ADD-ON)
    'tier1': 'QVWJPPNRPTT2BWKDUHXUYBEG',  // Tier 1: Simple Art - $20-40
    'tier2': 'DWYX3BHATZ6WDHKUNQC23PPP',   // Tier 2: Moderate Art - $40-60
    'tier3': 'BJM3MWA2GQCZS6HPQFU63XJX',  // Tier 3: Advanced Art - $60-80
    'tier4': 'NLBDWZQDGENWRCFVHZGPU3HD'    // Tier 4: Extreme Art - $80-100+
};
// ------------------------------------------

document.addEventListener('DOMContentLoaded', () => {
    // Booking Form Selectors
    const baseServices = document.querySelectorAll('input[name="base-service"]');
    const designTiers = document.querySelectorAll('input[name="design-tier"]');
    
    // Summary Selectors
    const summaryItems = document.getElementById('summary-items');
    const totalPriceEl = document.getElementById('total-price');
    const totalTimeEl = document.getElementById('total-time');

    // Policy Selectors
    const policyModal = document.getElementById('policy-modal');
    const openPolicyLink = document.getElementById('open-policy-link');
    const closeModalBtn = document.getElementById('close-modal');
    const modalAgreeBtn = document.getElementById('modal-agree-btn');
    const policyCheckbox = document.getElementById('policy-agree-checkbox');
    const bookBtn = document.getElementById('book-btn');

    /* --- CALCULATOR LOGIC --- */
    function formatTime(minutes) {
        const h = Math.floor(minutes / 60);
        const m = minutes % 60;
        if (h > 0 && m > 0) return `${h}h ${m}m`;
        if (h > 0) return `${h}h`;
        return `${m}m`;
    }

    function calculateTotal() {
        summaryItems.innerHTML = '';
        let minPrice = 0;
        let maxPrice = 0;
        let totalTime = 0;

        // 1. Base Service
        const selectedBase = document.querySelector('input[name="base-service"]:checked');
        if (selectedBase) {
            const price = parseInt(selectedBase.dataset.price);
            const time = parseInt(selectedBase.dataset.time);
            minPrice += price;
            maxPrice += price;
            totalTime += time;

            const name = selectedBase.nextElementSibling.querySelector('h3').innerText;
            addSummaryItem(name, `$${price}`);
        }

        // 2. Design Tier
        const selectedTier = document.querySelector('input[name="design-tier"]:checked');
        if (selectedTier && selectedTier.value !== 'none') {
            const pMin = parseInt(selectedTier.dataset.priceMin);
            const pMax = parseInt(selectedTier.dataset.priceMax);
            const time = parseInt(selectedTier.dataset.time);
            
            minPrice += pMin;
            maxPrice += pMax;
            totalTime += time;

            const name = selectedTier.nextElementSibling.querySelector('h3').innerText.trim();
            addSummaryItem(name, `+$${pMin} - $${pMax}`);
        }

        // Update DOM
        if (minPrice === maxPrice) {
            totalPriceEl.innerText = `$${minPrice}`;
        } else {
            totalPriceEl.innerText = `$${minPrice} - $${maxPrice}`;
        }
        
        totalTimeEl.innerText = formatTime(totalTime);
    }

    function addSummaryItem(name, priceStr) {
        const li = document.createElement('li');
        li.className = 'summary-item';
        li.innerHTML = `
            <span class="item-name">${name}</span>
            <span class="item-price">${priceStr}</span>
        `;
        summaryItems.appendChild(li);
    }

    /* --- POLICY LOGIC --- */
    function updateBookButton() {
        bookBtn.disabled = !policyCheckbox.checked;
    }

    policyCheckbox.addEventListener('change', updateBookButton);

    bookBtn.addEventListener('click', () => {
        let squareServices = [];
        let serviceNames = [];

        // Grab Base Service Square ID + Name
        const selectedBase = document.querySelector('input[name="base-service"]:checked');
        if (selectedBase && SQUARE_SERVICE_MAP[selectedBase.value]) {
            squareServices.push(SQUARE_SERVICE_MAP[selectedBase.value]);
            serviceNames.push(selectedBase.nextElementSibling.querySelector('h3').innerText);
        }

        // Grab Design Tier Square ID + Name
        const selectedTier = document.querySelector('input[name="design-tier"]:checked');
        if (selectedTier && selectedTier.value !== 'none' && SQUARE_SERVICE_MAP[selectedTier.value]) {
            squareServices.push(SQUARE_SERVICE_MAP[selectedTier.value]);
            serviceNames.push(selectedTier.nextElementSibling.querySelector('h3').innerText.trim());
        }

        // Grab Price/Time for the next page
        const totalPrice = totalPriceEl.innerText;
        const totalTime = totalTimeEl.innerText;

        const bookingUrl = `booking.html?services=${encodeURIComponent(squareServices.join(','))}&names=${encodeURIComponent(serviceNames.join(','))}&price=${encodeURIComponent(totalPrice)}&time=${encodeURIComponent(totalTime)}`;

        bookBtn.classList.add('is-navigating');
        bookBtn.setAttribute('aria-busy', 'true');
        bookBtn.innerHTML = '<span class="spinner" aria-hidden="true"></span>';

        const go = () => {
            window.location.href = bookingUrl;
        };
        requestAnimationFrame(() => {
            requestAnimationFrame(() => {
                setTimeout(go, 64);
            });
        });
    });

    function openModal(e) {
        if (e) e.preventDefault();
        policyModal.classList.add('active');
    }

    function closeModal() {
        policyModal.classList.remove('active');
    }

    openPolicyLink.addEventListener('click', openModal);
    closeModalBtn.addEventListener('click', closeModal);
    
    // Agreeing from the modal
    modalAgreeBtn.addEventListener('click', () => {
        policyCheckbox.checked = true;
        updateBookButton();
        closeModal();
    });

    // Close modal if clicked outside content
    policyModal.addEventListener('click', (e) => {
        if (e.target === policyModal) {
            closeModal();
        }
    });

    // Form Event listeners
    baseServices.forEach(el => el.addEventListener('change', calculateTotal));
    designTiers.forEach(el => el.addEventListener('change', calculateTotal));

    // Initial calculation setup
    calculateTotal();
    updateBookButton();
});
