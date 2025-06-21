class InvoiceGenerator {
    constructor() {
        this.invoices = JSON.parse(localStorage.getItem('invoices')) || [];
        this.currentInvoice = {};
        this.itemCounter = 0;
        
        this.initializeEventListeners();
        this.generateInvoiceNumber();
        this.setDefaultDates();
        this.addInitialItem();
        this.updateHistoryDisplay();
    }

    initializeEventListeners() {
        // Tab switching
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => this.switchTab(e.target.dataset.tab));
        });

        // Form interactions
        document.getElementById('addItemBtn').addEventListener('click', () => this.addItem());
        document.getElementById('previewBtn').addEventListener('click', () => this.previewInvoice());
        document.getElementById('saveBtn').addEventListener('click', () => this.saveInvoice());
        document.getElementById('clearBtn').addEventListener('click', () => this.clearForm());

        // Modal interactions
        document.querySelector('.close').addEventListener('click', () => this.closeModal());
        document.querySelector('.close-success').addEventListener('click', () => this.closeSaveSuccessModal());
        document.getElementById('printBtn').addEventListener('click', () => this.printInvoice());
        document.getElementById('downloadBtn').addEventListener('click', () => this.downloadPDF());

        // Save success modal actions
        document.getElementById('printInvoiceBtn').addEventListener('click', () => this.printSavedInvoice());
        document.getElementById('createNewBtn').addEventListener('click', () => this.createNewInvoice());
        document.getElementById('viewHistoryBtn').addEventListener('click', () => this.viewHistory());

        // History interactions
        document.getElementById('searchInput').addEventListener('input', (e) => this.searchInvoices(e.target.value));
        document.getElementById('exportBtn').addEventListener('click', () => this.exportAllInvoices());

        // Tax rate change
        document.getElementById('taxRate').addEventListener('input', () => this.calculateTotals());

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

    generateInvoiceNumber() {
        const lastInvoice = this.invoices[this.invoices.length - 1];
        const lastNumber = lastInvoice ? parseInt(lastInvoice.invoiceNumber.replace('INV-', '')) : 0;
        const newNumber = (lastNumber + 1).toString().padStart(4, '0');
        document.getElementById('invoiceNumber').value = `INV-${newNumber}`;
    }

    setDefaultDates() {
        const today = new Date();
        const dueDate = new Date(today);
        dueDate.setDate(today.getDate() + 30);

        document.getElementById('invoiceDate').value = today.toISOString().split('T')[0];
        document.getElementById('dueDate').value = dueDate.toISOString().split('T')[0];
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
            <span class="item-amount">$0.00</span>
            <button type="button" class="remove-item" onclick="invoiceGen.removeItem(${this.itemCounter})">Remove</button>
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
        
        itemRow.querySelector('.item-amount').textContent = `$${amount.toFixed(2)}`;
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
        const total = subtotal + taxAmount;
        
        document.getElementById('subtotal').textContent = `$${subtotal.toFixed(2)}`;
        document.getElementById('taxAmount').textContent = `$${taxAmount.toFixed(2)}`;
        document.getElementById('total').textContent = `$${total.toFixed(2)}`;
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
        const total = subtotal + taxAmount;
        
        return {
            invoiceNumber: document.getElementById('invoiceNumber').value,
            invoiceDate: document.getElementById('invoiceDate').value,
            dueDate: document.getElementById('dueDate').value,
            businessInfo: {
                name: document.getElementById('businessName').value,
                email: document.getElementById('businessEmail').value,
                phone: document.getElementById('businessPhone').value,
                address: document.getElementById('businessAddress').value
            },
            clientInfo: {
                name: document.getElementById('clientName').value,
                email: document.getElementById('clientEmail').value,
                phone: document.getElementById('clientPhone').value,
                address: document.getElementById('clientAddress').value
            },
            items,
            subtotal,
            taxRate,
            taxAmount,
            total,
            createdAt: new Date().toISOString()
        };
    }

    validateForm(data) {
        const errors = [];
        
        if (!data.businessInfo.name) errors.push('Business name is required');
        if (!data.clientInfo.name) errors.push('Client name is required');
        if (!data.invoiceDate) errors.push('Invoice date is required');
        if (!data.dueDate) errors.push('Due date is required');
        if (data.items.length === 0) errors.push('At least one item is required');
        
        return errors;
    }

    previewInvoice() {
        const data = this.collectFormData();
        const errors = this.validateForm(data);
        
        if (errors.length > 0) {
            alert('Please fix the following errors:\n' + errors.join('\n'));
            return;
        }
        
        this.currentInvoice = data;
        this.generateInvoicePreview(data);
        document.getElementById('previewModal').style.display = 'block';
    }

    generateInvoicePreview(data) {
        const preview = document.getElementById('invoicePreview');
        
        preview.innerHTML = `
            <div class="invoice-header">
                <div>
                    <h1 class="invoice-title">INVOICE</h1>
                </div>
                <div class="invoice-info">
                    <p><strong>Invoice #:</strong> ${data.invoiceNumber}</p>
                    <p><strong>Date:</strong> ${this.formatDate(data.invoiceDate)}</p>
                    <p><strong>Due Date:</strong> ${this.formatDate(data.dueDate)}</p>
                </div>
            </div>
            
            <div class="invoice-parties">
                <div class="party-info">
                    <h3>From:</h3>
                    <p><strong>${data.businessInfo.name}</strong></p>
                    ${data.businessInfo.email ? `<p>${data.businessInfo.email}</p>` : ''}
                    ${data.businessInfo.phone ? `<p>${data.businessInfo.phone}</p>` : ''}
                    ${data.businessInfo.address ? `<p>${data.businessInfo.address.replace(/\n/g, '<br>')}</p>` : ''}
                </div>
                <div class="party-info">
                    <h3>To:</h3>
                    <p><strong>${data.clientInfo.name}</strong></p>
                    ${data.clientInfo.email ? `<p>${data.clientInfo.email}</p>` : ''}
                    ${data.clientInfo.phone ? `<p>${data.clientInfo.phone}</p>` : ''}
                    ${data.clientInfo.address ? `<p>${data.clientInfo.address.replace(/\n/g, '<br>')}</p>` : ''}
                </div>
            </div>
            
            <table class="invoice-table">
                <thead>
                    <tr>
                        <th>Description</th>
                        <th class="text-right">Quantity</th>
                        <th class="text-right">Rate</th>
                        <th class="text-right">Amount</th>
                    </tr>
                </thead>
                <tbody>
                    ${data.items.map(item => `
                        <tr>
                            <td>${item.description}</td>
                            <td class="text-right">${item.quantity}</td>
                            <td class="text-right">$${item.rate.toFixed(2)}</td>
                            <td class="text-right">$${item.amount.toFixed(2)}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            
            <div class="invoice-totals">
                <div class="total-row">
                    <span>Subtotal:</span>
                    <span>$${data.subtotal.toFixed(2)}</span>
                </div>
                ${data.taxRate > 0 ? `
                    <div class="total-row">
                        <span>Tax (${data.taxRate}%):</span>
                        <span>$${data.taxAmount.toFixed(2)}</span>
                    </div>
                ` : ''}
                <div class="total-row final">
                    <span>Total:</span>
                    <span>$${data.total.toFixed(2)}</span>
                </div>
            </div>
        `;
    }

    async saveInvoice() {
        const data = this.collectFormData();
        const errors = this.validateForm(data);
        
        if (errors.length > 0) {
            alert('Please fix the following errors:\n' + errors.join('\n'));
            return;
        }
        
        // Check if invoice number already exists
        const existingInvoice = this.invoices.find(inv => inv.invoiceNumber === data.invoiceNumber);
        if (existingInvoice) {
            if (!confirm('An invoice with this number already exists. Do you want to update it?')) {
                return;
            }
            // Update existing invoice
            const index = this.invoices.findIndex(inv => inv.invoiceNumber === data.invoiceNumber);
            this.invoices[index] = data;
        } else {
            // Add new invoice
            this.invoices.push(data);
        }
        
        localStorage.setItem('invoices', JSON.stringify(this.invoices));
        this.currentInvoice = data;
        
        // Auto-generate and download Excel file
        await this.generateExcelFile(data);
        
        // Show success modal
        document.getElementById('saveSuccessModal').style.display = 'block';
    }

    async generateExcelFile(invoiceData) {
        try {
            // Create a new workbook
            const wb = XLSX.utils.book_new();
            
            // Create invoice summary sheet
            const summaryData = [
                ['Invoice Details'],
                ['Invoice Number', invoiceData.invoiceNumber],
                ['Invoice Date', this.formatDate(invoiceData.invoiceDate)],
                ['Due Date', this.formatDate(invoiceData.dueDate)],
                [''],
                ['Business Information'],
                ['Business Name', invoiceData.businessInfo.name],
                ['Email', invoiceData.businessInfo.email],
                ['Phone', invoiceData.businessInfo.phone],
                ['Address', invoiceData.businessInfo.address],
                [''],
                ['Client Information'],
                ['Client Name', invoiceData.clientInfo.name],
                ['Email', invoiceData.clientInfo.email],
                ['Phone', invoiceData.clientInfo.phone],
                ['Address', invoiceData.clientInfo.address],
                [''],
                ['Invoice Summary'],
                ['Subtotal', `$${invoiceData.subtotal.toFixed(2)}`],
                ['Tax Rate', `${invoiceData.taxRate}%`],
                ['Tax Amount', `$${invoiceData.taxAmount.toFixed(2)}`],
                ['Total Amount', `$${invoiceData.total.toFixed(2)}`]
            ];
            
            const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(wb, summaryWs, 'Invoice Summary');
            
            // Create items detail sheet
            const itemsData = [
                ['Item Description', 'Quantity', 'Rate', 'Total Amount']
            ];
            
            invoiceData.items.forEach(item => {
                itemsData.push([
                    item.description,
                    item.quantity,
                    `$${item.rate.toFixed(2)}`,
                    `$${item.amount.toFixed(2)}`
                ]);
            });
            
            // Add totals row
            itemsData.push(['', '', 'Subtotal:', `$${invoiceData.subtotal.toFixed(2)}`]);
            if (invoiceData.taxRate > 0) {
                itemsData.push(['', '', `Tax (${invoiceData.taxRate}%):`, `$${invoiceData.taxAmount.toFixed(2)}`]);
            }
            itemsData.push(['', '', 'Total:', `$${invoiceData.total.toFixed(2)}`]);
            
            const itemsWs = XLSX.utils.aoa_to_sheet(itemsData);
            XLSX.utils.book_append_sheet(wb, itemsWs, 'Items Detail');
            
            // Generate filename
            const filename = `Invoice_${invoiceData.invoiceNumber}_${invoiceData.invoiceDate}.xlsx`;
            
            // Write and download the file
            XLSX.writeFile(wb, filename);
            
        } catch (error) {
            console.error('Error generating Excel file:', error);
            alert('There was an error generating the Excel file, but your invoice has been saved successfully.');
        }
    }

    printSavedInvoice() {
        this.generateInvoicePreview(this.currentInvoice);
        this.closeSaveSuccessModal();
        
        // Create a new window for printing
        const printWindow = window.open('', '_blank');
        const invoiceContent = document.getElementById('invoicePreview').innerHTML;
        
        printWindow.document.write(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Invoice ${this.currentInvoice.invoiceNumber}</title>
                <style>
                    body { font-family: Arial, sans-serif; margin: 20px; }
                    .invoice-header { display: flex; justify-content: space-between; margin-bottom: 30px; border-bottom: 2px solid #333; padding-bottom: 20px; }
                    .invoice-title { font-size: 2rem; color: #333; margin: 0; }
                    .invoice-info { text-align: right; }
                    .invoice-parties { display: flex; justify-content: space-between; margin-bottom: 30px; }
                    .party-info h3 { color: #333; margin-bottom: 10px; }
                    .invoice-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
                    .invoice-table th, .invoice-table td { padding: 10px; border-bottom: 1px solid #ddd; text-align: left; }
                    .invoice-table th { background: #f5f5f5; font-weight: bold; }
                    .text-right { text-align: right; }
                    .invoice-totals { margin-left: auto; width: 300px; }
                    .total-row { display: flex; justify-content: space-between; padding: 5px 0; border-bottom: 1px solid #ddd; }
                    .total-row.final { font-weight: bold; font-size: 1.2rem; border-bottom: 2px solid #333; }
                </style>
            </head>
            <body>
                ${invoiceContent}
            </body>
            </html>
        `);
        
        printWindow.document.close();
        printWindow.focus();
        printWindow.print();
        printWindow.close();
    }

    createNewInvoice() {
        this.closeSaveSuccessModal();
        this.clearForm();
        this.generateInvoiceNumber();
        this.switchTab('create');
    }

    viewHistory() {
        this.closeSaveSuccessModal();
        this.switchTab('history');
    }

    clearForm() {
        // Clear all form fields
        document.querySelectorAll('input, textarea').forEach(field => {
            if (field.id !== 'invoiceNumber') {
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
    }

    closeModal() {
        document.getElementById('previewModal').style.display = 'none';
    }

    closeSaveSuccessModal() {
        document.getElementById('saveSuccessModal').style.display = 'none';
    }

    printInvoice() {
        window.print();
    }

    downloadPDF() {
        // For a simple implementation, we'll use the browser's print to PDF functionality
        alert('Use your browser\'s "Print to PDF" option to save the invoice as PDF.');
        window.print();
    }

    updateHistoryDisplay() {
        this.displayInvoices(this.invoices);
        this.updateStats();
    }

    displayInvoices(invoices) {
        const historyList = document.getElementById('historyList');
        
        if (invoices.length === 0) {
            historyList.innerHTML = '<p style="text-align: center; color: #4a5568; padding: 40px;">No invoices found.</p>';
            return;
        }
        
        historyList.innerHTML = invoices.map(invoice => `
            <div class="history-item">
                <div class="history-item-header">
                    <span class="invoice-number">${invoice.invoiceNumber}</span>
                    <span class="invoice-total">$${invoice.total.toFixed(2)}</span>
                </div>
                <div class="history-item-details">
                    <div><strong>Client:</strong> ${invoice.clientInfo.name}</div>
                    <div><strong>Date:</strong> ${this.formatDate(invoice.invoiceDate)}</div>
                    <div><strong>Due:</strong> ${this.formatDate(invoice.dueDate)}</div>
                    <div><strong>Items:</strong> ${invoice.items.length}</div>
                </div>
                <div class="history-item-actions">
                    <button class="btn btn-primary btn-small" onclick="invoiceGen.viewInvoice('${invoice.invoiceNumber}')">View</button>
                    <button class="btn btn-secondary btn-small" onclick="invoiceGen.duplicateInvoice('${invoice.invoiceNumber}')">Duplicate</button>
                    <button class="btn btn-outline btn-small" onclick="invoiceGen.deleteInvoice('${invoice.invoiceNumber}')">Delete</button>
                </div>
            </div>
        `).join('');
    }

    updateStats() {
        const totalInvoices = this.invoices.length;
        const totalRevenue = this.invoices.reduce((sum, inv) => sum + inv.total, 0);
        const avgInvoice = totalInvoices > 0 ? totalRevenue / totalInvoices : 0;
        
        document.getElementById('totalInvoices').textContent = totalInvoices;
        document.getElementById('totalRevenue').textContent = `$${totalRevenue.toFixed(2)}`;
        document.getElementById('avgInvoice').textContent = `$${avgInvoice.toFixed(2)}`;
    }

    searchInvoices(query) {
        if (!query.trim()) {
            this.displayInvoices(this.invoices);
            return;
        }
        
        const filtered = this.invoices.filter(invoice => 
            invoice.invoiceNumber.toLowerCase().includes(query.toLowerCase()) ||
            invoice.clientInfo.name.toLowerCase().includes(query.toLowerCase()) ||
            invoice.businessInfo.name.toLowerCase().includes(query.toLowerCase())
        );
        
        this.displayInvoices(filtered);
    }

    viewInvoice(invoiceNumber) {
        const invoice = this.invoices.find(inv => inv.invoiceNumber === invoiceNumber);
        if (invoice) {
            this.generateInvoicePreview(invoice);
            document.getElementById('previewModal').style.display = 'block';
        }
    }

    duplicateInvoice(invoiceNumber) {
        const invoice = this.invoices.find(inv => inv.invoiceNumber === invoiceNumber);
        if (invoice) {
            // Switch to create tab
            this.switchTab('create');
            
            // Fill form with invoice data
            document.getElementById('businessName').value = invoice.businessInfo.name;
            document.getElementById('businessEmail').value = invoice.businessInfo.email;
            document.getElementById('businessPhone').value = invoice.businessInfo.phone;
            document.getElementById('businessAddress').value = invoice.businessInfo.address;
            
            document.getElementById('clientName').value = invoice.clientInfo.name;
            document.getElementById('clientEmail').value = invoice.clientInfo.email;
            document.getElementById('clientPhone').value = invoice.clientInfo.phone;
            document.getElementById('clientAddress').value = invoice.clientInfo.address;
            
            document.getElementById('taxRate').value = invoice.taxRate;
            
            // Clear existing items and add invoice items
            document.getElementById('itemsList').innerHTML = '';
            this.itemCounter = 0;
            
            invoice.items.forEach(item => {
                this.addItem();
                const lastItem = document.querySelector('.item-row:last-child');
                lastItem.querySelector('.item-description').value = item.description;
                lastItem.querySelector('.item-quantity').value = item.quantity;
                lastItem.querySelector('.item-rate').value = item.rate;
                this.calculateItemAmount(lastItem);
            });
            
            // Generate new invoice number
            this.generateInvoiceNumber();
            this.setDefaultDates();
        }
    }

    deleteInvoice(invoiceNumber) {
        if (confirm('Are you sure you want to delete this invoice? This action cannot be undone.')) {
            this.invoices = this.invoices.filter(inv => inv.invoiceNumber !== invoiceNumber);
            localStorage.setItem('invoices', JSON.stringify(this.invoices));
            this.updateHistoryDisplay();
        }
    }

    exportAllInvoices() {
        if (this.invoices.length === 0) {
            alert('No invoices to export.');
            return;
        }
        
        this.generateAllInvoicesExcel();
    }

    async generateAllInvoicesExcel() {
        try {
            const wb = XLSX.utils.book_new();
            
            // Create summary sheet
            const summaryData = [
                ['Invoice Number', 'Date', 'Due Date', 'Business Name', 'Client Name', 'Client Email', 'Subtotal', 'Tax Rate', 'Tax Amount', 'Total', 'Items Count']
            ];
            
            this.invoices.forEach(invoice => {
                summaryData.push([
                    invoice.invoiceNumber,
                    invoice.invoiceDate,
                    invoice.dueDate,
                    invoice.businessInfo.name,
                    invoice.clientInfo.name,
                    invoice.clientInfo.email,
                    invoice.subtotal.toFixed(2),
                    invoice.taxRate,
                    invoice.taxAmount.toFixed(2),
                    invoice.total.toFixed(2),
                    invoice.items.length
                ]);
            });
            
            const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(wb, summaryWs, 'All Invoices Summary');
            
            // Create detailed items sheet
            const itemsData = [
                ['Invoice Number', 'Item Description', 'Quantity', 'Rate', 'Total Amount', 'Invoice Date', 'Client Name']
            ];
            
            this.invoices.forEach(invoice => {
                invoice.items.forEach(item => {
                    itemsData.push([
                        invoice.invoiceNumber,
                        item.description,
                        item.quantity,
                        item.rate.toFixed(2),
                        item.amount.toFixed(2),
                        invoice.invoiceDate,
                        invoice.clientInfo.name
                    ]);
                });
            });
            
            const itemsWs = XLSX.utils.aoa_to_sheet(itemsData);
            XLSX.utils.book_append_sheet(wb, itemsWs, 'All Items Detail');
            
            // Generate filename
            const filename = `All_Invoices_Export_${new Date().toISOString().split('T')[0]}.xlsx`;
            
            // Write and download the file
            XLSX.writeFile(wb, filename);
            
        } catch (error) {
            console.error('Error generating Excel export:', error);
            alert('There was an error generating the Excel export file.');
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
const invoiceGen = new InvoiceGenerator();