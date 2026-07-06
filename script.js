
    // Convert array to Map for instant searching
    const dataMap = new Map(pincodeData.map(item => [item.pincode, item]));

    const input = document.getElementById('pincodeInput');
    const resultContainer = document.getElementById('resultContainer');
    const errorMsg = document.getElementById('errorMsg');
    const loader = document.getElementById('loader');
    const btnText = document.getElementById('btnText');

    // Trigger search on Enter Key
    input.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') checkPincode();
    });

    function checkPincode() {
        const pin = input.value.trim();
        
        // Validation
        if (!/^\d{6}$/.test(pin)) {
            errorMsg.style.display = 'block';
            resultContainer.style.display = 'none';
            return;
        }
        
        errorMsg.style.display = 'none';
        showLoading(true);

        // Simulate a small delay for better UX (loading feel)
        setTimeout(() => {
            const data = dataMap.get(pin);
            renderResult(data);
            showLoading(false);
        }, 400);
    }

    function showLoading(isLoading) {
        if (isLoading) {
            loader.style.display = 'block';
            btnText.style.display = 'none';
        } else {
            loader.style.display = 'none';
            btnText.style.display = 'block';
        }
    }

    function renderResult(data) {
        resultContainer.style.display = 'block';
        
        if (!data) {
            resultContainer.innerHTML = `
                <div class="not-found animate__animated animate__shakeX">
                    <i class="fa-solid fa-circle-exclamation"></i> Not Serviceable
                    <p style="font-size: 0.8rem; color: #6b7280; font-weight: 400; margin-top: 5px;">
                        We currently do not provide service to this location.
                    </p>
                </div>
            `;
            return;
        }

        const typeLabel = data.pincodeType === 'STD' ? 'Serviceable' : 'ODA (Out of Delivery Area)';
        const typeColor = data.pincodeType === 'STD' ? 'var(--success)' : 'var(--warning)';

        resultContainer.innerHTML = `
            <div class="result-card">
                <div class="header-row">
                    <div>
                        <div class="label">Primary Branch</div>
                        <div class="branch-name">${data.branch}</div>
                    </div>
                    <span class="status-badge ${data.status === 'Active' ? 'status-active' : 'status-inactive'}">
                        ${data.status}
                    </span>
                </div>

                <div class="info-grid">
                    <div class="info-item">
                        <div class="label">Air Branch</div>
                        <div class="value">${data.airBranch === '-' ? 'N/A' : data.airBranch}</div>
                    </div>
                    <div class="info-item">
                        <div class="label">Pincode Type</div>
                        <div class="value" style="color: ${typeColor}">${typeLabel}</div>
                    </div>
                </div>

                <hr style="margin: 15px 0; border: none; border-top: 1px dashed #e5e7eb;">

                <div class="info-grid">
                    <div class="info-item">
                        <div class="label">Road Service</div>
                        <div class="service-tag ${data.isRoadDeliveryAllowed === 'Yes' ? 'tag-yes' : 'tag-no'}">
                            <i class="fa-solid ${data.isRoadDeliveryAllowed === 'Yes' ? 'fa-truck' : 'fa-truck-slash'}"></i>
                            Pickup: ${data.isRoadPickupAllowed} / Del: ${data.isRoadDeliveryAllowed}
                        </div>
                    </div>
                    <div class="info-item">
                        <div class="label">Air Service</div>
                        <div class="service-tag ${data.isAirDeliveryAllowed === 'Yes' ? 'tag-yes' : 'tag-no'}">
                            <i class="fa-solid ${data.isAirDeliveryAllowed === 'Yes' ? 'fa-plane' : 'fa-plane-slash'}"></i>
                            Pickup: ${data.isAirPickupAllowed} / Del: ${data.isAirDeliveryAllowed}
                        </div>
                    </div>
                </div>
            </div>
        `;
    }
