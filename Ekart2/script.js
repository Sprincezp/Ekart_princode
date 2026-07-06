
        // Create Map for O(1) lookup
        const pincodeMap = new Map();
        pincodeData.forEach(item => pincodeMap.set(item.pincode, item));

        // DOM Elements
        const pincodeInput = document.getElementById('pincodeInput');
        const searchBtn = document.getElementById('searchBtn');
        const btnText = document.getElementById('btnText');
        const errorMsg = document.getElementById('errorMsg');
        const errorText = document.getElementById('errorText');
        const resultSection = document.getElementById('resultSection');
        const themeBtn = document.getElementById('themeBtn');

        // Theme Toggle
        function initTheme() {
            const saved = localStorage.getItem('theme');
            if (saved === 'light') {
                document.documentElement.setAttribute('data-theme', 'light');
                themeBtn.querySelector('.icon-sun').style.display = 'none';
                themeBtn.querySelector('.icon-moon').style.display = 'block';
            }
        }
        initTheme();

        themeBtn.addEventListener('click', () => {
            const current = document.documentElement.getAttribute('data-theme');
            const newTheme = current === 'light' ? 'dark' : 'light';
            document.documentElement.setAttribute('data-theme', newTheme);
            localStorage.setItem('theme', newTheme);
            
            const sunIcon = themeBtn.querySelector('.icon-sun');
            const moonIcon = themeBtn.querySelector('.icon-moon');
            if (newTheme === 'light') {
                sunIcon.style.display = 'none';
                moonIcon.style.display = 'block';
            } else {
                sunIcon.style.display = 'block';
                moonIcon.style.display = 'none';
            }
        });

        // Input Handler - Numbers Only
        pincodeInput.addEventListener('input', (e) => {
            e.target.value = e.target.value.replace(/\D/g, '');
            hideError();
            if (e.target.value.length === 6) searchPincode();
        });

        // Enter Key
        pincodeInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') searchPincode();
        });

        // Search Button
        searchBtn.addEventListener('click', searchPincode);

        // Search Function
        async function searchPincode() {
            const pincode = pincodeInput.value.trim();

            if (!pincode) {
                showError('Please enter a pincode');
                return;
            }

            if (pincode.length !== 6) {
                showError('Pincode must be exactly 6 digits');
                return;
            }

            // Loading State
            setLoading(true);
            hideError();
            await new Promise(r => setTimeout(r, 400));
            setLoading(false);

            const result = pincodeMap.get(pincode);
            if (result) {
                showResult(result);
            } else {
                showNotFound(pincode);
            }
        }

        // Helper Functions
        function showError(msg) {
            pincodeInput.classList.add('error');
            errorText.textContent = msg;
            errorMsg.classList.add('show');
            resultSection.innerHTML = '';
        }

        function hideError() {
            pincodeInput.classList.remove('error');
            errorMsg.classList.remove('show');
        }

        function setLoading(loading) {
            searchBtn.disabled = loading;
            btnText.innerHTML = loading ? '<div class="spinner"></div>' : 'Search';
        }

        function renderYesNo(val) {
            const isYes = val.toLowerCase() === 'yes';
            return `<span class="status-${isYes ? 'yes' : 'no'}">
                <svg viewBox="0 0 24 24">
                    ${isYes 
                        ? '<path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>' 
                        : '<path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>'}
                </svg>
                ${isYes ? 'Yes' : 'No'}
            </span>`;
        }

        // Show Result
        function showResult(data) {
            const airBranch = data.airBranch === '-' ? 'N/A' : data.airBranch;
            const typeClass = data.pincodeType === 'STD' ? 'service-std' : 'service-oda';
            const typeText = data.pincodeType === 'STD' ? 'Serviceable' : 'ODA (Out of Delivery Area)';
            const statusClass = data.status === 'Active' ? 'status-active' : 'status-inactive';

            resultSection.innerHTML = `
                <div class="result-card">
                    <div class="result-header">
                        <div class="result-pincode">
                            <span class="pincode-badge">${data.pincode}</span>
                            <span class="result-branch">${data.branch}</span>
                        </div>
                        <span class="status-badge ${statusClass}">${data.status}</span>
                    </div>
                    <div class="result-body">
                        <div class="info-grid">
                            <div class="info-box">
                                <div class="info-label">Branch Name</div>
                                <div class="info-value">${data.branch}</div>
                            </div>
                            <div class="info-box">
                                <div class="info-label">Air Branch</div>
                                <div class="info-value">${airBranch}</div>
                            </div>
                            <div class="info-box">
                                <div class="info-label">Service Type</div>
                                <div class="info-value">
                                    <span class="service-type ${typeClass}">${typeText}</span>
                                </div>
                            </div>
                            <div class="info-box">
                                <div class="info-label">Serviceable</div>
                                <div class="info-value">${renderYesNo(data.isServiceable)}</div>
                            </div>
                        </div>

                        <div class="service-section">
                            <div class="service-title">Road Services</div>
                            <div class="service-grid">
                                <div class="service-item">
                                    <span class="service-item-name">Pickup</span>
                                    ${renderYesNo(data.isRoadPickupAllowed)}
                                </div>
                                <div class="service-item">
                                    <span class="service-item-name">Delivery</span>
                                    ${renderYesNo(data.isRoadDeliveryAllowed)}
                                </div>
                            </div>
                        </div>

                        <div class="service-section">
                            <div class="service-title">Air Services</div>
                            <div class="service-grid">
                                <div class="service-item">
                                    <span class="service-item-name">Pickup</span>
                                    ${renderYesNo(data.isAirPickupAllowed)}
                                </div>
                                <div class="service-item">
                                    <span class="service-item-name">Delivery</span>
                                    ${renderYesNo(data.isAirDeliveryAllowed)}
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            `;
        }

        // Show Not Found
        function showNotFound(pincode) {
            resultSection.innerHTML = `
                <div class="not-found">
                    <div class="not-found-icon">
                        <svg viewBox="0 0 24 24">
                            <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
                        </svg>
                    </div>
                    <h3 class="not-found-title">Not Serviceable</h3>
                    <p class="not-found-text">We currently do not provide service to this location.</p>
                </div>
            `;
        }

        // Focus on load
        pincodeInput.focus();
