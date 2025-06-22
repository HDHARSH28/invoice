class EstimateGenerator {
    constructor() {
        this.estimates = JSON.parse(localStorage.getItem('estimates')) || [];
        this.currentEstimate = {};
        this.itemCounter = 0;
        
        this.checkLibraries();
        this.initializeEventListeners();
        this.generateEstimateNumber();
        this.setDefaultDates();
        this.addInitialItem();
        this.updateHistoryDisplay();
    }

    checkLibraries() {
        // Check if required libraries are loaded
        const libraries = {
            'XLSX': typeof XLSX !== 'undefined',
            'jsPDF': typeof window.jspdf !== 'undefined',
            'html2canvas': typeof html2canvas !== 'undefined'
        };
        
        console.log('Library status:', libraries);
        
        // Show warning if libraries are missing
        const missingLibraries = Object.entries(libraries)
            .filter(([name, loaded]) => !loaded)
            .map(([name]) => name);
            
        if (missingLibraries.length > 0) {
            console.warn('Missing libraries:', missingLibraries);
        }
    }

    initializeEventListeners() {
        // Tab switching
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => this.switchTab(e.target.dataset.tab));
        });

        // Form interactions
        document.getElementById('addItemBtn').addEventListener('click', () => this.addItem());
        document.getElementById('previewBtn').addEventListener('click', () => this.previewEstimate());
        document.getElementById('saveBtn').addEventListener('click', () => this.saveEstimate());
        document.getElementById('clearBtn').addEventListener('click', () => this.clearForm());

        // Modal interactions
        document.querySelector('.close').addEventListener('click', () => this.closeModal());
        document.querySelector('.close-success').addEventListener('click', () => this.closeSaveSuccessModal());
        document.getElementById('printBtn').addEventListener('click', () => this.printEstimate());
        document.getElementById('downloadBtn').addEventListener('click', () => this.downloadPDF());

        // Save success modal actions
        document.getElementById('printInvoiceBtn').addEventListener('click', () => this.printSavedEstimate());
        document.getElementById('createNewBtn').addEventListener('click', () => this.createNewEstimate());
        document.getElementById('viewHistoryBtn').addEventListener('click', () => this.viewHistory());

        // History interactions
        document.getElementById('searchInput').addEventListener('input', (e) => this.searchEstimates(e.target.value));
        document.getElementById('exportBtn').addEventListener('click', () => this.exportAllEstimates());

        // Tax rate change
        document.getElementById('taxRate').addEventListener('input', () => this.calculateTotals());
        
        // Discount rate change
        document.getElementById('discountAmountInput').addEventListener('input', () => this.calculateTotals());

        // Close modal when clicking outside
        window.addEventListener('click', (e) => {
            if (e.target.classList.contains('modal')) {
                this.closeModal();
                this.closeSaveSuccessModal();
            }
        });
    }

    switchTab(tabName) {
        // Update tab buttons
        document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        // Update tab content
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        document.getElementById(`${tabName}-tab`).classList.add('active');

        if (tabName === 'history') {
            this.updateHistoryDisplay();
        }
    }

    generateEstimateNumber() {
        // Find the highest estimate number
        let maxNumber = 0;
        this.estimates.forEach(estimate => {
            const number = parseInt(estimate.estimateNumber);
            if (!isNaN(number) && number > maxNumber) {
                maxNumber = number;
            }
        });
        
        // Get the next number
        const nextNumber = maxNumber + 1;
        
        // Determine the number of digits needed
        let digitCount = 4; // Start with 4 digits
        if (nextNumber > 9999) {
            digitCount = nextNumber.toString().length;
        }
        
        // Format the estimate number
        const estimateNumber = nextNumber.toString().padStart(digitCount, '0');
        
        document.getElementById('estimateNumber').value = estimateNumber;
    }

    setDefaultDates() {
        const today = new Date();
        document.getElementById('estimateDate').value = today.toISOString().split('T')[0];
    }

    addItem() {
        this.itemCounter++;
        const itemsContainer = document.getElementById('itemsList');
        
        const itemRow = document.createElement('div');
        itemRow.className = 'item-row';
        itemRow.dataset.itemId = this.itemCounter;
        
        itemRow.innerHTML = `
            <input type="text" placeholder="Item description" class="item-description" required>
            <input type="number" placeholder="1" class="item-quantity" min="1" value="1" required>
            <input type="number" placeholder="0.00" class="item-rate" min="0" step="0.01" required>
            <span class="item-amount">Rs. 0.00</span>
            <button type="button" class="remove-item" onclick="estimateGen.removeItem(${this.itemCounter})">Remove</button>
        `;
        
        itemsContainer.appendChild(itemRow);
        
        // Add event listeners for calculation
        const quantityInput = itemRow.querySelector('.item-quantity');
        const rateInput = itemRow.querySelector('.item-rate');
        
        quantityInput.addEventListener('input', () => this.calculateItemAmount(itemRow));
        rateInput.addEventListener('input', () => this.calculateItemAmount(itemRow));
        
        this.calculateTotals();
    }

    addInitialItem() {
        this.addItem();
    }

    removeItem(itemId) {
        const itemRow = document.querySelector(`[data-item-id="${itemId}"]`);
        if (itemRow) {
            itemRow.remove();
            this.calculateTotals();
        }
    }

    calculateItemAmount(itemRow) {
        const quantity = parseFloat(itemRow.querySelector('.item-quantity').value) || 0;
        const rate = parseFloat(itemRow.querySelector('.item-rate').value) || 0;
        const amount = quantity * rate;
        
        itemRow.querySelector('.item-amount').textContent = `Rs. ${amount.toFixed(2)}`;
        this.calculateTotals();
    }

    calculateTotals() {
        const itemRows = document.querySelectorAll('.item-row');
        let subtotal = 0;
        
        itemRows.forEach(row => {
            const quantity = parseFloat(row.querySelector('.item-quantity').value) || 0;
            const rate = parseFloat(row.querySelector('.item-rate').value) || 0;
            subtotal += quantity * rate;
        });
        
        const taxRate = parseFloat(document.getElementById('taxRate').value) || 0;
        const taxAmount = (subtotal * taxRate) / 100;
        
        const discountAmount = parseFloat(document.getElementById('discountAmountInput').value) || 0;
        
        const total = subtotal + taxAmount - discountAmount;
        
        document.getElementById('subtotal').textContent = `Rs. ${subtotal.toFixed(2)}`;
        document.getElementById('taxAmount').textContent = `Rs. ${taxAmount.toFixed(2)}`;
        document.getElementById('discountAmount').textContent = `Rs. ${discountAmount.toFixed(2)}`;
        document.getElementById('total').textContent = `Rs. ${total.toFixed(2)}`;
    }

    collectFormData() {
        const itemRows = document.querySelectorAll('.item-row');
        const items = [];
        
        itemRows.forEach(row => {
            const description = row.querySelector('.item-description').value;
            const quantity = parseFloat(row.querySelector('.item-quantity').value) || 0;
            const rate = parseFloat(row.querySelector('.item-rate').value) || 0;
            
            if (description && quantity > 0 && rate >= 0) {
                items.push({
                    description,
                    quantity,
                    rate,
                    amount: quantity * rate
                });
            }
        });
        
        const subtotal = items.reduce((sum, item) => sum + item.amount, 0);
        const taxRate = parseFloat(document.getElementById('taxRate').value) || 0;
        const taxAmount = (subtotal * taxRate) / 100;
        
        const discountAmount = parseFloat(document.getElementById('discountAmountInput').value) || 0;
        
        const total = subtotal + taxAmount - discountAmount;
        
        return {
            estimateNumber: document.getElementById('estimateNumber').value,
            estimateDate: document.getElementById('estimateDate').value,
            businessInfo: {
                name: document.getElementById('businessName').value,
                email: document.getElementById('businessEmail').value,
                phone: document.getElementById('businessPhone').value,
                address: document.getElementById('businessAddress').value
            },
            clientInfo: {
                name: document.getElementById('clientName').value,
                phone: document.getElementById('clientPhone').value,
                address: document.getElementById('clientAddress').value
            },
            items,
            subtotal,
            taxRate,
            taxAmount,
            discountAmount,
            total,
            createdAt: new Date().toISOString()
        };
    }

    validateForm(data) {
        const errors = [];
        
        if (!data.businessInfo.name) errors.push('Business name is required');
        if (!data.clientInfo.name) errors.push('Client name is required');
        if (!data.estimateDate) errors.push('Estimate date is required');
        if (data.items.length === 0) errors.push('At least one item is required');
        
        return errors;
    }

    previewEstimate() {
        const data = this.collectFormData();
        const errors = this.validateForm(data);
        
        if (errors.length > 0) {
            alert('Please fix the following errors:\n' + errors.join('\n'));
            return;
        }
        
        this.currentEstimate = data;
        this.generateEstimatePreview(data);
        document.getElementById('previewModal').style.display = 'block';
    }

    generateEstimatePreview(data) {
        const preview = document.getElementById('estimatePreview');
        const emptyRows = Math.max(0, 15 - data.items.length);
        
        preview.innerHTML = `
        <div class="print-template">
            <header class="print-header">
                <div class="header-left">
                    <h1 class="main-title">ESTIMATE</h1>
                    <div class="company-details">
                        <p><strong>${data.businessInfo.name}</strong></p>
                        <p>${data.businessInfo.address.replace(/\n/g, '<br>')}</p>
                        <p>${data.businessInfo.phone}</p>
                        <p>${data.businessInfo.email}</p>
                </div>
                </div>
                <div class="header-right">
                    <div class="company-logo">
                        <img src="logo.png" alt="New Patel Tiles & Sanitary Logo">
                    </div>
                    <div class="bill-to">
                        <h3>BILL TO</h3>
                        <p>${data.clientInfo.name}</p>
                        <p>${data.clientInfo.address ? data.clientInfo.address.replace(/\n/g, '<br>') : ''}</p>
                        <p>${data.clientInfo.phone || ''}</p>
                    </div>
                </div>
            </header>

            <div class="estimate-meta">
                <div class="meta-item">
                    <span>ESTIMATE NO</span>
                    <span>${data.estimateNumber}</span>
                </div>
                <div class="meta-item">
                    <span>DATE</span>
                    <span>${this.formatDate(data.estimateDate)}</span>
                </div>
            </div>
            
            <table class="items-table">
                <thead>
                    <tr>
                        <th class="qty-col">Qty</th>
                        <th class="desc-col">Description</th>
                        <th class="unit-col">Unit Cost</th>
                        <th class="amount-col">Amount</th>
                    </tr>
                </thead>
                <tbody>
                    ${data.items.map(item => `
                        <tr>
                            <td>${item.quantity}</td>
                            <td>${item.description}</td>
                            <td>Rs. ${item.rate.toFixed(2)}</td>
                            <td>Rs. ${item.amount.toFixed(2)}</td>
                        </tr>
                    `).join('')}
                    ${Array(emptyRows).fill(0).map(() => `
                        <tr>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            
            <footer class="print-footer">
                <div class="footer-left">
                    <div class="terms">
                        <h4>TERMS & CONDITIONS</h4>
                        <p>Your terms and conditions can be added here.</p>
                    </div>
                </div>
                <div class="footer-right">
                    <div class="totals">
                        <div class="total-row">
                            <span>SUB TOTAL</span>
                            <span>Rs. ${data.subtotal.toFixed(2)}</span>
                        </div>
                        ${data.discountAmount > 0 ? `
                            <div class="total-row">
                                <span>DISCOUNT</span>
                                <span>Rs. ${data.discountAmount.toFixed(2)}</span>
                            </div>
                        ` : ''}
                        <div class="total-row">
                            <span>TAX</span>
                            <span>Rs. ${data.taxAmount.toFixed(2)}</span>
                        </div>
                        <div class="total-row grand-total">
                            <span>TOTAL</span>
                            <span>Rs. ${data.total.toFixed(2)}</span>
                        </div>
                    </div>
                </div>
            </footer>
        </div>
    `;
    }

    async saveEstimate() {
        const data = this.collectFormData();
        const errors = this.validateForm(data);
        
        if (errors.length > 0) {
            alert('Please fix the following errors:\n' + errors.join('\n'));
            return;
        }
        
        if (this.isEditMode) {
            // Update existing estimate
            const index = this.estimates.findIndex(est => est.estimateNumber === this.editingEstimateNumber);
            if (index !== -1) {
                this.estimates[index] = data;
                localStorage.setItem('estimates', JSON.stringify(this.estimates));
                this.currentEstimate = data;
                
                // Auto-generate and download PDF file
                await this.generatePDFFile(data);
                
                // Update Excel structure with all estimates
                await this.updateExcelStructure();
                
                // Show success modal
                document.getElementById('saveSuccessModal').style.display = 'block';
                document.getElementById('savedEstimateNumber').textContent = data.estimateNumber;
                
                // Reset edit mode
                this.isEditMode = false;
                this.editingEstimateNumber = null;
                
                // Reset button text
                const saveBtn = document.getElementById('saveBtn');
                saveBtn.textContent = 'Save Estimate';
            }
        } else {
        // Check if estimate number already exists
        const existingEstimate = this.estimates.find(est => est.estimateNumber === data.estimateNumber);
        if (existingEstimate) {
            if (!confirm('An estimate with this number already exists. Do you want to update it?')) {
                return;
            }
            // Update existing estimate
            const index = this.estimates.findIndex(est => est.estimateNumber === data.estimateNumber);
            this.estimates[index] = data;
        } else {
            // Add new estimate
            this.estimates.push(data);
        }
        
        localStorage.setItem('estimates', JSON.stringify(this.estimates));
        this.currentEstimate = data;
        
            // Auto-generate and download PDF file
            await this.generatePDFFile(data);
            
            // Update Excel structure with all estimates
            await this.updateExcelStructure();
        
        // Show success modal
        document.getElementById('saveSuccessModal').style.display = 'block';
        document.getElementById('savedEstimateNumber').textContent = data.estimateNumber;
        }
    }

    async generatePDFFile(estimateData) {
        try {
            // Check if required libraries are loaded
            if (typeof window.jspdf === 'undefined') {
                throw new Error('jsPDF library not loaded');
            }
            
            if (typeof html2canvas === 'undefined') {
                throw new Error('html2canvas library not loaded');
            }
            
            // Show loading indicator
            const saveBtn = document.getElementById('saveBtn');
            const originalText = saveBtn.textContent;
            saveBtn.textContent = 'Generating PDF...';
            saveBtn.disabled = true;
            
            // Generate estimate preview HTML
            this.generateEstimatePreview(estimateData);
            const estimateElement = document.getElementById('estimatePreview');
            
            // Wait for DOM to update and ensure element is visible
            await new Promise(resolve => setTimeout(resolve, 300));
            
            // Make sure the preview modal is visible for html2canvas
            const modal = document.getElementById('previewModal');
            modal.style.display = 'block';
            modal.style.position = 'absolute';
            modal.style.left = '-9999px';
            
            try {
                // Convert HTML to canvas using html2canvas
                const canvas = await html2canvas(estimateElement, {
                    scale: 1.5, // Slightly lower scale for better compatibility
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: '#ffffff',
                    logging: false,
                    removeContainer: true
                });
                
                // Create PDF using jsPDF for A5 size
                const { jsPDF } = window.jspdf;
                const pdf = new jsPDF('p', 'mm', 'a5'); // A5 size
                
                const imgWidth = 148; // A5 width in mm
                const pageHeight = 210; // A5 height in mm
                const imgHeight = (canvas.height * imgWidth) / canvas.width;
                let heightLeft = imgHeight;
                let position = 0;
                
                // Add image to PDF
                const imgData = canvas.toDataURL('image/jpeg', 0.95);
                pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
                heightLeft -= pageHeight;
                
                // Add additional pages if content is longer than one page
                while (heightLeft >= 0) {
                    position = heightLeft - imgHeight;
                    pdf.addPage();
                    pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
                    heightLeft -= pageHeight;
                }
            
            // Generate filename
                const filename = `Estimate_${estimateData.estimateNumber}_${estimateData.estimateDate}.pdf`;
                
                // Download the PDF
                pdf.save(filename);
                
            } finally {
                // Hide the modal again
                modal.style.display = 'none';
                modal.style.position = 'fixed';
                modal.style.left = 'auto';
            }
            
            // Reset button
            saveBtn.textContent = originalText;
            saveBtn.disabled = false;
            
        } catch (error) {
            console.error('Error generating PDF file:', error);
            
            // Reset button
            const saveBtn = document.getElementById('saveBtn');
            saveBtn.textContent = 'Save Estimate';
            saveBtn.disabled = false;
            
            // Try fallback method
            console.log('Trying fallback PDF generation method...');
            try {
                await this.generatePDFFileFallback(estimateData);
                return; // Success with fallback
            } catch (fallbackError) {
                console.error('Fallback PDF generation also failed:', fallbackError);
            }
            
            // Show specific error message
            let errorMessage = 'There was an error generating the PDF file, but your estimate has been saved successfully.';
            
            if (error.message.includes('jsPDF')) {
                errorMessage += ' Please refresh the page and try again.';
            } else if (error.message.includes('html2canvas')) {
                errorMessage += ' Please try using the "Download PDF" button in the preview modal.';
            } else {
                errorMessage += ' Please try using the "Download PDF" button in the preview modal.';
            }
            
            alert(errorMessage);
        }
    }

    async generatePDFFileFallback(estimateData) {
        try {
            this.generateEstimatePreview(estimateData);
            
            const printWindow = window.open('', '_blank');
            const estimateContent = document.getElementById('estimatePreview').outerHTML;
            
            printWindow.document.write(`
                <!DOCTYPE html>
                <html>
                <head>
                    <title>Estimate ${estimateData.estimateNumber}</title>
                    <link rel="stylesheet" href="styles.css">
                    <style>
                       body { background: #fff !important; }
                       .print-template {
                           border: none !important;
                           box-shadow: none !important;
                       }
                    </style>
                </head>
                <body>
                    ${estimateContent}
                </body>
                </html>
            `);
            
            printWindow.document.close();
            
            setTimeout(() => {
                printWindow.focus();
                printWindow.print();
                printWindow.close();
            }, 500);
            
        } catch (error) {
            console.error('Error in fallback PDF generation:', error);
            alert('PDF generation failed. Please use your browser\'s "Print to PDF" option from the preview modal.');
        }
    }

    printSavedEstimate() {
        this.generateEstimatePreview(this.currentEstimate);
        this.closeSaveSuccessModal();
        
        const printWindow = window.open('', '_blank');
        const estimateContent = document.getElementById('estimatePreview').outerHTML;
        
        printWindow.document.write(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Estimate ${this.currentEstimate.estimateNumber}</title>
                <link rel="stylesheet" href="styles.css">
                <style>
                   /* Override styles for clean printing */
                   body { background: #fff !important; }
                   .print-template {
                       border: none !important;
                       box-shadow: none !important;
                       margin: 0;
                       padding: 0;
                   }
                </style>
            </head>
            <body>
                ${estimateContent}
            </body>
            </html>
        `);
        
        printWindow.document.close();

        setTimeout(() => {
        printWindow.focus();
        printWindow.print();
        printWindow.close();
        }, 500);
    }

    createNewEstimate() {
        this.closeSaveSuccessModal();
        this.clearForm();
        this.generateEstimateNumber();
        this.switchTab('create');
    }

    viewHistory() {
        this.closeSaveSuccessModal();
        this.switchTab('history');
    }

    clearForm() {
        // Clear all form fields
        document.querySelectorAll('input, textarea').forEach(field => {
            if (field.id !== 'estimateNumber') {
                field.value = '';
            }
        });
        
        // Clear items
        document.getElementById('itemsList').innerHTML = '';
        this.itemCounter = 0;
        
        // Reset dates and add initial item
        this.setDefaultDates();
        this.addInitialItem();
        this.calculateTotals();
        
        // Reset edit mode
        this.isEditMode = false;
        this.editingEstimateNumber = null;
        
        // Reset button text
        const saveBtn = document.getElementById('saveBtn');
        saveBtn.textContent = 'Save Estimate';
    }

    closeModal() {
        document.getElementById('previewModal').style.display = 'none';
    }

    closeSaveSuccessModal() {
        document.getElementById('saveSuccessModal').style.display = 'none';
    }

    printEstimate() {
        this.generateEstimatePreview(this.currentEstimate);
        const printWindow = window.open('', '_blank');
        const estimateContent = document.getElementById('estimatePreview').outerHTML;
        
        printWindow.document.write(`
             <!DOCTYPE html>
             <html>
             <head>
                <title>Estimate</title>
                <link rel="stylesheet" href="styles.css">
                <style>
                   body { background: #fff !important; }
                   .print-template {
                       border: none !important;
                       box-shadow: none !important;
                   }
                </style>
             </head>
             <body>
                ${estimateContent}
             </body>
             </html>
        `);
        
        printWindow.document.close();

        setTimeout(() => {
            printWindow.focus();
            printWindow.print();
            printWindow.close();
        }, 500);
    }

    async downloadPDF() {
        // Generate PDF for the current preview estimate
        if (this.currentEstimate && Object.keys(this.currentEstimate).length > 0) {
            try {
                await this.generatePDFFile(this.currentEstimate);
            } catch (error) {
                console.error('Primary PDF generation failed, trying fallback:', error);
                try {
                    await this.generatePDFFileFallback(this.currentEstimate);
                } catch (fallbackError) {
                    console.error('Fallback PDF generation also failed:', fallbackError);
                    alert('PDF generation failed. Please use your browser\'s "Print to PDF" option.');
                }
            }
        } else {
            alert('No estimate data available for PDF download.');
        }
    }

    updateHistoryDisplay() {
        this.displayEstimates(this.estimates);
        this.updateStats();
    }

    displayEstimates(estimates) {
        const historyList = document.getElementById('historyList');
        
        if (estimates.length === 0) {
            historyList.innerHTML = '<p style="text-align: center; color: #4a5568; padding: 40px;">No estimates found.</p>';
            return;
        }
        
        historyList.innerHTML = estimates.map(estimate => `
            <div class="history-item">
                <div class="history-item-header">
                    <span class="estimate-number">${estimate.estimateNumber}</span>
                    <span class="estimate-total">Rs. ${estimate.total.toFixed(2)}</span>
                </div>
                <div class="history-item-details">
                    <div><strong>Client:</strong> ${estimate.clientInfo.name}</div>
                    <div><strong>Date:</strong> ${this.formatDate(estimate.estimateDate)}</div>
                    <div><strong>Items:</strong> ${estimate.items.length}</div>
                </div>
                <div class="history-item-actions">
                    <button class="btn btn-primary btn-small" onclick="estimateGen.viewEstimate('${estimate.estimateNumber}')">View</button>
                    <button class="btn btn-secondary btn-small" onclick="estimateGen.editEstimate('${estimate.estimateNumber}')">Edit</button>
                    <button class="btn btn-secondary btn-small" onclick="estimateGen.duplicateEstimate('${estimate.estimateNumber}')">Duplicate</button>
                    <button class="btn btn-outline btn-small" onclick="estimateGen.deleteEstimate('${estimate.estimateNumber}')">Delete</button>
                </div>
            </div>
        `).join('');
    }

    formatDate(dateStr) {
        const date = new Date(dateStr);
        return date.toLocaleDateString();
    }

    viewEstimate(estimateNumber) {
        const estimate = this.estimates.find(est => est.estimateNumber === estimateNumber);
        if (estimate) {
            this.generateEstimatePreview(estimate);
            document.getElementById('previewModal').style.display = 'block';
        }
    }

    updateStats() {
        const totalEstimates = this.estimates.length;
        const totalRevenue = this.estimates.reduce((sum, est) => sum + est.total, 0);
        
        // Calculate this month revenue
        const now = new Date();
        const thisMonth = now.getMonth();
        const thisYear = now.getFullYear();
        
        const thisMonthRevenue = this.estimates.reduce((sum, est) => {
            const estimateDate = new Date(est.estimateDate);
            if (estimateDate.getMonth() === thisMonth && estimateDate.getFullYear() === thisYear) {
                return sum + est.total;
            }
            return sum;
        }, 0);
        
        // Calculate this year revenue
        const thisYearRevenue = this.estimates.reduce((sum, est) => {
            const estimateDate = new Date(est.estimateDate);
            if (estimateDate.getFullYear() === thisYear) {
                return sum + est.total;
            }
            return sum;
        }, 0);
        
        document.getElementById('totalEstimates').textContent = totalEstimates;
        document.getElementById('totalRevenue').textContent = `Rs. ${totalRevenue.toFixed(2)}`;
        document.getElementById('thisMonthRevenue').textContent = `Rs. ${thisMonthRevenue.toFixed(2)}`;
        document.getElementById('thisYearRevenue').textContent = `Rs. ${thisYearRevenue.toFixed(2)}`;
    }

    searchEstimates(query) {
        if (!query.trim()) {
            this.displayEstimates(this.estimates);
            return;
        }
        
        const filtered = this.estimates.filter(estimate => 
            estimate.estimateNumber.toLowerCase().includes(query.toLowerCase()) ||
            estimate.clientInfo.name.toLowerCase().includes(query.toLowerCase()) ||
            estimate.businessInfo.name.toLowerCase().includes(query.toLowerCase())
        );
        
        this.displayEstimates(filtered);
    }

    duplicateEstimate(estimateNumber) {
        const estimate = this.estimates.find(est => est.estimateNumber === estimateNumber);
        if (estimate) {
            // Switch to create tab
            this.switchTab('create');
            
            // Fill form with estimate data
            document.getElementById('businessName').value = estimate.businessInfo.name;
            document.getElementById('businessEmail').value = estimate.businessInfo.email;
            document.getElementById('businessPhone').value = estimate.businessInfo.phone;
            document.getElementById('businessAddress').value = estimate.businessInfo.address;
            
            document.getElementById('clientName').value = estimate.clientInfo.name;
            document.getElementById('clientPhone').value = estimate.clientInfo.phone;
            document.getElementById('clientAddress').value = estimate.clientInfo.address;
            
            document.getElementById('taxRate').value = estimate.taxRate;
            document.getElementById('discountAmountInput').value = estimate.discountAmount || 0;
            
            // Clear existing items and add estimate items
            document.getElementById('itemsList').innerHTML = '';
            this.itemCounter = 0;
            
            estimate.items.forEach(item => {
                this.addItem();
                const lastItem = document.querySelector('.item-row:last-child');
                lastItem.querySelector('.item-description').value = item.description;
                lastItem.querySelector('.item-quantity').value = item.quantity;
                lastItem.querySelector('.item-rate').value = item.rate;
                this.calculateItemAmount(lastItem);
            });
            
            // Generate new estimate number
            this.generateEstimateNumber();
            this.setDefaultDates();
        }
    }

    editEstimate(estimateNumber) {
        const estimate = this.estimates.find(est => est.estimateNumber === estimateNumber);
        if (estimate) {
            // Switch to create tab
            this.switchTab('create');
            
            // Set edit mode
            this.isEditMode = true;
            this.editingEstimateNumber = estimateNumber;
            
            // Fill form with estimate data
            document.getElementById('businessName').value = estimate.businessInfo.name;
            document.getElementById('businessEmail').value = estimate.businessInfo.email;
            document.getElementById('businessPhone').value = estimate.businessInfo.phone;
            document.getElementById('businessAddress').value = estimate.businessInfo.address;
            
            document.getElementById('clientName').value = estimate.clientInfo.name;
            document.getElementById('clientPhone').value = estimate.clientInfo.phone;
            document.getElementById('clientAddress').value = estimate.clientInfo.address;
            
            document.getElementById('estimateDate').value = estimate.estimateDate;
            document.getElementById('taxRate').value = estimate.taxRate;
            document.getElementById('discountAmountInput').value = estimate.discountAmount || 0;
            
            // Clear existing items and add estimate items
            document.getElementById('itemsList').innerHTML = '';
            this.itemCounter = 0;
            
            estimate.items.forEach(item => {
                this.addItem();
                const lastItem = document.querySelector('.item-row:last-child');
                lastItem.querySelector('.item-description').value = item.description;
                lastItem.querySelector('.item-quantity').value = item.quantity;
                lastItem.querySelector('.item-rate').value = item.rate;
                this.calculateItemAmount(lastItem);
            });
            
            // Update button text to indicate edit mode
            const saveBtn = document.getElementById('saveBtn');
            saveBtn.textContent = 'Update Estimate';
            
            // Scroll to top of form
            document.querySelector('.estimate-form').scrollIntoView({ behavior: 'smooth' });
        }
    }

    deleteEstimate(estimateNumber) {
        if (confirm('Are you sure you want to delete this estimate? This action cannot be undone.')) {
            this.estimates = this.estimates.filter(est => est.estimateNumber !== estimateNumber);
            localStorage.setItem('estimates', JSON.stringify(this.estimates));
            this.updateHistoryDisplay();
            
            // Update Excel structure after deletion
            this.updateExcelStructure();
        }
    }

    exportAllEstimates() {
        if (this.estimates.length === 0) {
            alert('No estimates to export.');
            return;
        }
        
        this.updateExcelStructure();
    }

    async updateExcelStructure() {
        try {
            // Create a new workbook
            const wb = XLSX.utils.book_new();
            
            // Create summary sheet with all estimates
            const summaryData = [
                ['Estimate Number', 'Date', 'Business Name', 'Client Name', 'Client Phone', 'Subtotal', 'Tax Rate', 'Tax Amount', 'Discount Amount', 'Total', 'Items Count', 'Created At']
            ];
            
            this.estimates.forEach(estimate => {
                summaryData.push([
                    estimate.estimateNumber,
                    estimate.estimateDate,
                    estimate.businessInfo.name,
                    estimate.clientInfo.name,
                    estimate.clientInfo.phone || '',
                    estimate.subtotal.toFixed(2),
                    estimate.taxRate,
                    estimate.taxAmount.toFixed(2),
                    estimate.discountAmount || 0,
                    estimate.total.toFixed(2),
                    estimate.items.length,
                    this.formatDate(estimate.createdAt)
                ]);
            });
            
            const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(wb, summaryWs, 'All Estimates Summary');
            
            // Create detailed items sheet with all items from all estimates
            const itemsData = [
                ['Estimate Number', 'Item Description', 'Quantity', 'Rate', 'Total Amount', 'Estimate Date', 'Client Name', 'Business Name']
            ];
            
            this.estimates.forEach(estimate => {
                estimate.items.forEach(item => {
                    itemsData.push([
                        estimate.estimateNumber,
                        item.description,
                        item.quantity,
                        item.rate.toFixed(2),
                        item.amount.toFixed(2),
                        estimate.estimateDate,
                        estimate.clientInfo.name,
                        estimate.businessInfo.name
                    ]);
                });
            });
            
            const itemsWs = XLSX.utils.aoa_to_sheet(itemsData);
            XLSX.utils.book_append_sheet(wb, itemsWs, 'All Items Detail');
            
            // Create business information sheet
            const businessData = [
                ['Business Name', 'Email', 'Phone', 'Address', 'Total Estimates', 'Total Revenue']
            ];
            
            // Group by business
            const businessMap = new Map();
            this.estimates.forEach(estimate => {
                const businessKey = estimate.businessInfo.name;
                if (!businessMap.has(businessKey)) {
                    businessMap.set(businessKey, {
                        ...estimate.businessInfo,
                        totalEstimates: 0,
                        totalRevenue: 0
                    });
                }
                const business = businessMap.get(businessKey);
                business.totalEstimates++;
                business.totalRevenue += estimate.total;
            });
            
            businessMap.forEach(business => {
                businessData.push([
                    business.name,
                    business.email,
                    business.phone,
                    business.address,
                    business.totalEstimates,
                    business.totalRevenue.toFixed(2)
                ]);
            });
            
            const businessWs = XLSX.utils.aoa_to_sheet(businessData);
            XLSX.utils.book_append_sheet(wb, businessWs, 'Business Information');
            
            // Create client information sheet
            const clientData = [
                ['Client Name', 'Phone', 'Address', 'Total Estimates', 'Total Amount']
            ];
            
            // Group by client
            const clientMap = new Map();
            this.estimates.forEach(estimate => {
                const clientKey = estimate.clientInfo.name;
                if (!clientMap.has(clientKey)) {
                    clientMap.set(clientKey, {
                        ...estimate.clientInfo,
                        totalEstimates: 0,
                        totalAmount: 0
                    });
                }
                const client = clientMap.get(clientKey);
                client.totalEstimates++;
                client.totalAmount += estimate.total;
            });
            
            clientMap.forEach(client => {
                clientData.push([
                    client.name,
                    client.phone || '',
                    client.address || '',
                    client.totalEstimates,
                    client.totalAmount.toFixed(2)
                ]);
            });
            
            const clientWs = XLSX.utils.aoa_to_sheet(clientData);
            XLSX.utils.book_append_sheet(wb, clientWs, 'Client Information');
            
            // Generate filename with timestamp
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0];
            const filename = `Estimate_Database_${timestamp}.xlsx`;
            
            // Write and download the file
            XLSX.writeFile(wb, filename);
            
        } catch (error) {
            console.error('Error updating Excel structure:', error);
            // Don't show alert as this is a background process
        }
    }

    formatDate(dateString) {
        const date = new Date(dateString);
        return date.toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        });
    }
}

// Initialize the application
const estimateGen = new EstimateGenerator();
