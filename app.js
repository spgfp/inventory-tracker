document.addEventListener('DOMContentLoaded', function() {
    // Initialize the app
    let currentDate = new Date();
    let currentWeek = getWeekNumber(currentDate);
    let currentYear = currentDate.getFullYear();
    
    // Default inventory items - you can customize this list
    const defaultItems = [
        'Apple Juice', 'Bacon Bits', 'Bagels', 'Beef Sticks', 'Brownies', 'Caramel Apple Pops', 'Cereal', 'Cheese',
        'Chips', 'Chocolate Chips', 'Cholula', 'Churros', 'Cinnamon Rolls', 'Cream Cheese', 'Donuts (Powdered)',
        'Eggs', 'Famous Amos', 'French Fries', 'Fun Dip', 'Goldfish', 'Gummy Bears', 'Granulated Garlic',
        'Hot Chocolate', 'Ice Drinks', 'Icing', 'Jimmy Deans', 'Keystone Detergent', 'Keystone Sanitizer',
        'Lemonade', 'Macarons', 'Milk', 'Mints', 'Nutella', 'Oil', 'Orange Juice', 'Pandas', 'Pan Quillon Sheets',
        'Pepper', 'Pocky Sticks', 'Pretzels', 'Pure Leaf Tea', 'Red Vines', 'Rice Krispies', 'Scones',
        'Sour Patch', 'Sprinkles', 'Starbucks Frappuccinos', 'Sugar', 'Sugar Cookies', 'Tissue Pick Ups',
        'Tortilla (8")', 'Water'
    ];
    
    // Load saved data from localStorage
    let inventoryData = JSON.parse(localStorage.getItem('inventoryData')) || {};
    let masterItemList = JSON.parse(localStorage.getItem('masterItemList')) || [...defaultItems];
    
    // DOM Elements
    const currentWeekElement = document.getElementById('currentWeek');
    const prevWeekBtn = document.getElementById('prevWeek');
    const nextWeekBtn = document.getElementById('nextWeek');
    const inventoryTable = document.getElementById('inventoryTable');
    const addNewItemBtn = document.getElementById('addNewItem');
    const exportExcelBtn = document.getElementById('exportExcel');
    const emailExportBtn = document.getElementById('emailExport');
    const printListBtn = document.getElementById('printList');
    const emailModal = new bootstrap.Modal(document.getElementById('emailModal'));
    const addItemModal = new bootstrap.Modal(document.getElementById('addItemModal'));
    const sendEmailBtn = document.getElementById('sendEmail');
    const emailInput = document.getElementById('emailInput');
    const fromNameInput = document.getElementById('fromNameInput');
    const saveNewItemBtn = document.getElementById('saveNewItem');
    const newItemNameInput = document.getElementById('newItemName');
    const resetDataBtn = document.getElementById('resetData');
    
    // EmailJS Configuration - Replace with your own credentials
    const EMAILJS_CONFIG = {
        publicKey: 'YOUR_EMAILJS_PUBLIC_KEY', // Replace with your EmailJS public key
        serviceId: 'YOUR_EMAILJS_SERVICE_ID', // Replace with your EmailJS service ID
        templateId: 'YOUR_EMAILJS_TEMPLATE_ID' // Replace with your EmailJS template ID
    };
    
    // Gmail API Configuration
    const GMAIL_CONFIG = {
        clientId: 'YOUR_GOOGLE_CLIENT_ID', // Replace with your Google OAuth client ID
        apiKey: 'YOUR_GOOGLE_API_KEY', // Replace with your Google API key
        scope: 'https://www.googleapis.com/auth/gmail.send'
    };
    
    let isGmailSignedIn = false;
    
    // Initialize EmailJS
    if (EMAILJS_CONFIG.publicKey !== 'YOUR_EMAILJS_PUBLIC_KEY') {
        emailjs.init(EMAILJS_CONFIG.publicKey);
    }
    
    // Initialize the app
    updateWeekDisplay();
    loadWeekData();
    
    // Event Listeners
    prevWeekBtn.addEventListener('click', goToPreviousWeek);
    nextWeekBtn.addEventListener('click', goToNextWeek);
    addNewItemBtn.addEventListener('click', () => addItemModal.show());
    saveNewItemBtn.addEventListener('click', addNewItem);
    newItemNameInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') addNewItem();
    });
    exportExcelBtn.addEventListener('click', exportToExcel);
    emailExportBtn.addEventListener('click', () => emailModal.show());
    printListBtn.addEventListener('click', printInventoryList);
    sendEmailBtn.addEventListener('click', sendEmail);
    resetDataBtn.addEventListener('click', resetData);
    
    // Functions
    function getWeekNumber(date) {
        const firstDayOfYear = new Date(date.getFullYear(), 0, 1);
        const pastDaysOfYear = (date - firstDayOfYear) / 86400000;
        return Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
    }
    
    function updateWeekDisplay() {
        const weekDates = getWeekDateRange(currentYear, currentWeek);
        currentWeekElement.textContent = `Week of ${weekDates.start} - ${weekDates.end}`;
    }
    
    function getWeekDateRange(year, weekNumber) {
        // Get the first day of the year
        const firstDayOfYear = new Date(year, 0, 1);
        
        // Calculate the first Sunday of the year
        const firstSunday = new Date(firstDayOfYear);
        const dayOfWeek = firstDayOfYear.getDay();
        if (dayOfWeek !== 0) {
            firstSunday.setDate(firstDayOfYear.getDate() - dayOfWeek);
        }
        
        // Calculate the start of the target week (Sunday)
        const weekStart = new Date(firstSunday);
        weekStart.setDate(firstSunday.getDate() + (weekNumber - 1) * 7);
        
        // Calculate the end of the target week (Saturday)
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekStart.getDate() + 6);
        
        // Format dates
        const formatDate = (date) => {
            const month = date.getMonth() + 1;
            const day = date.getDate();
            return `${month}/${day}`;
        };
        
        return {
            start: formatDate(weekStart),
            end: formatDate(weekEnd)
        };
    }
    
    function goToPreviousWeek() {
        if (currentWeek === 1) {
            currentWeek = 52;
            currentYear--;
        } else {
            currentWeek--;
        }
        updateWeekDisplay();
        loadWeekData();
    }
    
    function goToNextWeek() {
        if (currentWeek === 52) {
            currentWeek = 1;
            currentYear++;
        } else {
            currentWeek++;
        }
        updateWeekDisplay();
        loadWeekData();
    }
    
    function addNewItem() {
        const itemName = newItemNameInput.value.trim();
        if (itemName && !masterItemList.includes(itemName)) {
            masterItemList.push(itemName);
            masterItemList.sort();
            saveMasterList();
            loadWeekData();
            addItemModal.hide();
            newItemNameInput.value = '';
        } else if (masterItemList.includes(itemName)) {
            alert('Item already exists in the list!');
        }
    }
    
    function deleteItem(itemName) {
        if (confirm(`Are you sure you want to delete "${itemName}" from the master list? This will remove it from all weeks.`)) {
            masterItemList = masterItemList.filter(item => item !== itemName);
            
            // Remove from all weeks' data
            Object.keys(inventoryData).forEach(weekKey => {
                if (inventoryData[weekKey][itemName]) {
                    delete inventoryData[weekKey][itemName];
                }
            });
            
            saveMasterList();
            saveData();
            loadWeekData();
        }
    }
    
    function resetData() {
        if (confirm('Are you sure you want to reset all data? This will:\n\n• Replace all items with the new updated list\n• Clear all saved quantities\n• Reset to default settings\n\nThis action cannot be undone.')) {
            // Clear localStorage
            localStorage.removeItem('inventoryData');
            localStorage.removeItem('masterItemList');
            
            // Reset to default items
            masterItemList = [...defaultItems];
            inventoryData = {};
            
            // Save and reload
            saveMasterList();
            saveData();
            loadWeekData();
            
            alert('Data has been reset successfully! Your new inventory items are now loaded.');
        }
    }
    
    function printInventoryList() {
        // Set print date attribute for CSS
        const now = new Date();
        const printDate = now.toLocaleDateString() + ' ' + now.toLocaleTimeString();
        document.body.setAttribute('data-print-date', printDate);
        
        // Create a print-specific version of the table
        const printWindow = window.open('', '_blank');
        
        // Get current week data
        const weekKey = `${currentYear}-W${currentWeek.toString().padStart(2, '0')}`;
        
        // Build HTML for print
        let printHTML = `
<!DOCTYPE html>
<html>
<head>
    <title>Inventory List - Week ${currentWeek}, ${currentYear}</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background: white;
            color: black;
        }
        .header {
            text-align: center;
            margin-bottom: 30px;
        }
        .title {
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 10px;
        }
        .subtitle {
            font-size: 16px;
            color: #666;
            margin-bottom: 5px;
        }
        .print-date {
            font-size: 12px;
            color: #999;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th {
            background: #f5f5f5;
            border: 1px solid #ccc;
            padding: 10px;
            text-align: center;
            font-weight: bold;
            font-size: 14px;
        }
        td {
            border: 1px solid #ccc;
            padding: 8px 10px;
            vertical-align: middle;
            font-size: 12px;
        }
        .item-name {
            font-weight: normal;
        }
        .quantity {
            text-align: center;
            font-weight: bold;
            font-size: 14px;
        }
        .empty {
            color: #999;
            font-style: italic;
        }
        @media print {
            body {
                margin: 0;
                padding: 15px;
            }
            .no-print {
                display: none;
            }
        }
    </style>
</head>
<body>
    <div class="header">
        <div class="title">Weekly Inventory Tracker</div>
        <div class="subtitle">Week ${currentWeek} - ${currentYear}</div>
        <div class="print-date">Printed on: ${printDate}</div>
    </div>
    
    <table>
        <thead>
            <tr>
                <th style="width: 60%; text-align: left;">Item</th>
                <th style="width: 20%;">Have</th>
                <th style="width: 20%;">Need</th>
            </tr>
        </thead>
        <tbody>
`;
        
        // Add each item to the table
        masterItemList.forEach(itemName => {
            const itemData = inventoryData[weekKey] && inventoryData[weekKey][itemName] 
                ? inventoryData[weekKey][itemName] 
                : { have: '', need: '' };
            
            const haveValue = itemData.have || '';
            const needValue = itemData.need || '';
            
            printHTML += `
            <tr>
                <td class="item-name">${itemName}</td>
                <td class="quantity">${haveValue || '<span class="empty">-</span>'}</td>
                <td class="quantity">${needValue || '<span class="empty">-</span>'}</td>
            </tr>
            `;
        });
        
        printHTML += `
        </tbody>
    </table>
    
    <div class="no-print" style="margin-top: 30px; text-align: center;">
        <button onclick="window.print()" style="padding: 10px 20px; font-size: 16px; cursor: pointer; background: #007bff; color: white; border: none; border-radius: 5px; margin-right: 10px;">Print</button>
        <button onclick="window.close()" style="padding: 10px 20px; font-size: 16px; cursor: pointer; background: #6c757d; color: white; border: none; border-radius: 5px;">Close</button>
    </div>
</body>
</html>
`;
        
        // Write HTML to print window
        printWindow.document.write(printHTML);
        printWindow.document.close();
        
        // Focus the print window
        printWindow.focus();
    }
    
    function editItemName(oldName, newName) {
        if (newName.trim() && newName !== oldName && !masterItemList.includes(newName)) {
            const index = masterItemList.indexOf(oldName);
            if (index !== -1) {
                masterItemList[index] = newName;
                masterItemList.sort();
                
                // Update all weeks' data
                Object.keys(inventoryData).forEach(weekKey => {
                    if (inventoryData[weekKey][oldName]) {
                        inventoryData[weekKey][newName] = inventoryData[weekKey][oldName];
                        delete inventoryData[weekKey][oldName];
                    }
                });
                
                saveMasterList();
                saveData();
                loadWeekData();
            }
        }
    }
    
    function updateQuantity(itemName, type, value) {
        const weekKey = `${currentYear}-W${currentWeek.toString().padStart(2, '0')}`;
        if (!inventoryData[weekKey]) {
            inventoryData[weekKey] = {};
        }
        if (!inventoryData[weekKey][itemName]) {
            inventoryData[weekKey][itemName] = { have: '', need: '' };
        }
        
        inventoryData[weekKey][itemName][type] = value.trim();
        saveData();
    }
    
    function renderTable() {
        inventoryTable.innerHTML = '';
        const weekKey = `${currentYear}-W${currentWeek.toString().padStart(2, '0')}`;
        
        masterItemList.forEach(itemName => {
            const itemData = inventoryData[weekKey] && inventoryData[weekKey][itemName] 
                ? inventoryData[weekKey][itemName] 
                : { have: 0, need: 0 };
            
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <span class="item-name" data-item="${itemName}">${itemName}</span>
                </td>
                <td class="text-center">
                    <input type="text" class="quantity-input" value="${itemData.have || ''}" 
                           data-item="${itemName}" data-type="have" placeholder="-">
                </td>
                <td class="text-center">
                    <input type="text" class="quantity-input" value="${itemData.need || ''}" 
                           data-item="${itemName}" data-type="need" placeholder="-">
                </td>
                <td class="text-center">
                    <button class="btn btn-sm btn-outline-primary btn-edit" data-action="edit" data-item="${itemName}">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn btn-sm btn-outline-danger btn-edit" data-action="delete" data-item="${itemName}">
                        <i class="fas fa-trash"></i>
                    </button>
                </td>
            `;
            
            inventoryTable.appendChild(row);
            
            // Add event listeners
            const haveInput = row.querySelector('input[data-type="have"]');
            const needInput = row.querySelector('input[data-type="need"]');
            const editBtn = row.querySelector('button[data-action="edit"]');
            const deleteBtn = row.querySelector('button[data-action="delete"]');
            const itemNameSpan = row.querySelector('.item-name');
            
            haveInput.addEventListener('change', (e) => {
                updateQuantity(itemName, 'have', e.target.value);
            });
            
            needInput.addEventListener('change', (e) => {
                updateQuantity(itemName, 'need', e.target.value);
            });
            
            editBtn.addEventListener('click', () => {
                editItemNameInline(itemNameSpan, itemName);
            });
            
            deleteBtn.addEventListener('click', () => {
                deleteItem(itemName);
            });
        });
    }
    
    function editItemNameInline(span, currentName) {
        const input = document.createElement('input');
        input.type = 'text';
        input.value = currentName;
        input.className = 'item-name editing';
        
        span.parentNode.replaceChild(input, span);
        input.focus();
        input.select();
        
        function saveEdit() {
            const newName = input.value.trim();
            editItemName(currentName, newName);
        }
        
        function cancelEdit() {
            input.parentNode.replaceChild(span, input);
        }
        
        input.addEventListener('blur', saveEdit);
        input.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                saveEdit();
            } else if (e.key === 'Escape') {
                cancelEdit();
            }
        });
    }
    
    function loadWeekData() {
        renderTable();
    }
    
    function saveData() {
        localStorage.setItem('inventoryData', JSON.stringify(inventoryData));
    }
    
    function saveMasterList() {
        localStorage.setItem('masterItemList', JSON.stringify(masterItemList));
    }
    
    function exportToExcel() {
        const weekKey = `${currentYear}-W${currentWeek.toString().padStart(2, '0')}`;
        const weekDates = getWeekDateRange(currentYear, currentWeek);
        
        // Prepare data for export
        const exportData = [
            ['Item Name', 'Have', 'Need']
        ];
        
        masterItemList.forEach(itemName => {
            const itemData = inventoryData[weekKey] && inventoryData[weekKey][itemName] 
                ? inventoryData[weekKey][itemName] 
                : { have: '', need: '' };
            
            exportData.push([itemName, itemData.have, itemData.need]);
        });
        
        // Create worksheet
        const ws = XLSX.utils.aoa_to_sheet(exportData);
        
        // Auto-size column widths to prevent text cutoff
        const colWidths = [
            { wch: 25 }, // Item Name column - wider for long item names
            { wch: 8 },  // Have column
            { wch: 8 }   // Need column
        ];
        ws['!cols'] = colWidths;
        
        // Create workbook
        const wb = XLSX.utils.book_new();
        const sheetName = `${weekDates.start.replace('/', '-')}_to_${weekDates.end.replace('/', '-')}`;
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        
        // Generate Excel file and trigger download
        const fileName = `Inventory_${weekDates.start.replace('/', '-')}_to_${weekDates.end.replace('/', '-')}.xlsx`;
        XLSX.writeFile(wb, fileName);
    }
    
    function generateExcelBuffer() {
        const weekKey = `${currentYear}-W${currentWeek.toString().padStart(2, '0')}`;
        
        // Prepare data for export
        const exportData = [
            ['Item Name', 'Have', 'Need']
        ];
        
        masterItemList.forEach(itemName => {
            const itemData = inventoryData[weekKey] && inventoryData[weekKey][itemName] 
                ? inventoryData[weekKey][itemName] 
                : { have: '', need: '' };
            
            exportData.push([itemName, itemData.have, itemData.need]);
        });
        
        // Create worksheet
        const ws = XLSX.utils.aoa_to_sheet(exportData);
        
        // Auto-size column widths to prevent text cutoff
        const colWidths = [
            { wch: 25 }, // Item Name column - wider for long item names
            { wch: 8 },  // Have column
            { wch: 8 }   // Need column
        ];
        ws['!cols'] = colWidths;
        
        // Create workbook
        const wb = XLSX.utils.book_new();
        const weekDates = getWeekDateRange(currentYear, currentWeek);
        const sheetName = `${weekDates.start.replace('/', '-')}_to_${weekDates.end.replace('/', '-')}`;
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        
        // Generate buffer
        return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    }
    
    function arrayBufferToBase64(buffer) {
        let binary = '';
        const bytes = new Uint8Array(buffer);
        const len = bytes.byteLength;
        for (let i = 0; i < len; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return btoa(binary);
    }
    
    function sendEmailWithFile() {
        const email = emailInput.value.trim();
        const fromName = fromNameInput.value.trim() || 'Inventory Tracker';
        const weekDates = getWeekDateRange(currentYear, currentWeek);
        
        // Generate Excel file
        const excelBuffer = generateExcelBuffer();
        const blob = new Blob([excelBuffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        // Create download URL
        const url = URL.createObjectURL(blob);
        const fileName = `Inventory_${weekDates.start.replace('/', '-')}_to_${weekDates.end.replace('/', '-')}.xlsx`;
        
        // Create temporary download link
        const downloadLink = document.createElement('a');
        downloadLink.href = url;
        downloadLink.download = fileName;
        downloadLink.style.display = 'none';
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        
        // Clean up the URL
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        
        // Create email content
        const subject = `Weekly Inventory - ${weekDates.start} to ${weekDates.end}`;
        const body = `Hello,

Please find the weekly inventory for ${weekDates.start} - ${weekDates.end}.

I've downloaded the Excel file to attach to this email. The file contains all inventory items with their current "Have" and "Need" quantities.

Best regards,
${fromName}`;
        
        // Create mailto link
        const mailtoLink = `mailto:${email}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
        
        // Try to open email client - works better on tablets
        const emailWindow = window.open(mailtoLink, '_blank');
        
        // Fallback for tablets - also try direct navigation
        setTimeout(() => {
            if (!emailWindow || emailWindow.closed) {
                window.location.href = mailtoLink;
            }
        }, 100);
        
        // Close modal
        emailModal.hide();
        emailInput.value = '';
        fromNameInput.value = '';
        
        // Detect if we're on a mobile/tablet device
        const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
        
        if (isMobile) {
            alert(`Excel file downloaded to your device! Your email app should open automatically. Look for the downloaded file "${fileName}" in your Downloads folder to attach it.`);
        } else {
            alert(`Excel file downloaded! Your email client should open with a pre-filled email. Please attach the downloaded file: ${fileName}`);
        }
    }
    
    function sendEmailWithText() {
        const email = emailInput.value.trim();
        const fromName = fromNameInput.value.trim() || 'Inventory Tracker';
        
        // Check if EmailJS is configured
        if (EMAILJS_CONFIG.publicKey === 'YOUR_PUBLIC_KEY') {
            alert('Email service is not configured. Please see the README for setup instructions.');
            return;
        }
        
        // Prepare email data
        const weekKey = `${currentYear}-W${currentWeek.toString().padStart(2, '0')}`;
        const weekDates = getWeekDateRange(currentYear, currentWeek);
        
        // Create formatted inventory table for email
        let inventoryTable = 'ITEM NAME                | HAVE | NEED\n';
        inventoryTable += '--------------------------|------|------\n';
        
        let hasData = false;
        masterItemList.forEach(itemName => {
            const itemData = inventoryData[weekKey] && inventoryData[weekKey][itemName] 
                ? inventoryData[weekKey][itemName] 
                : { have: '', need: '' };
            
            if (itemData.have || itemData.need) {
                hasData = true;
                const paddedName = itemName.padEnd(24);
                const paddedHave = (itemData.have || '').toString().padStart(4);
                const paddedNeed = (itemData.need || '').toString().padStart(4);
                inventoryTable += `${paddedName} | ${paddedHave} | ${paddedNeed}\n`;
            }
        });
        
        if (!hasData) {
            inventoryTable = 'No inventory data for this week.';
        }
        
        // Email template parameters
        const templateParams = {
            to_email: email,
            from_name: fromName,
            week_dates: `${weekDates.start} - ${weekDates.end}`,
            inventory_table: inventoryTable,
            subject: `Weekly Inventory - ${weekDates.start} to ${weekDates.end}`,
            excel_note: `You can also download the Excel file directly from the app using the "Export to Excel" button.`
        };
        
        // Show loading state
        const sendBtn = document.getElementById('sendEmail');
        const originalText = sendBtn.textContent;
        sendBtn.textContent = 'Sending...';
        sendBtn.disabled = true;
        
        // Send email using EmailJS
        emailjs.send(EMAILJS_CONFIG.serviceId, EMAILJS_CONFIG.templateId, templateParams)
            .then(function(response) {
                alert(`Inventory list for ${weekDates.start} - ${weekDates.end} has been sent to ${email}!`);
                emailModal.hide();
                emailInput.value = '';
                fromNameInput.value = '';
            })
            .catch(function(error) {
                console.error('Email sending failed:', error);
                alert('Failed to send email. Please check your configuration and try again.');
            })
            .finally(function() {
                // Restore button state
                sendBtn.textContent = originalText;
                sendBtn.disabled = false;
            });
    }
    
    async function uploadToFileIO(blob, fileName) {
        const formData = new FormData();
        formData.append('file', blob, fileName);
        
        try {
            const response = await fetch('https://file.io/', {
                method: 'POST',
                body: formData
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const result = await response.json();
            console.log('File.io response:', result);
            
            if (result.success && result.link) {
                return result.link;
            } else {
                throw new Error(result.message || 'Upload failed - no download link received');
            }
        } catch (error) {
            console.error('File.io upload failed:', error);
            throw new Error(`File upload failed: ${error.message}`);
        }
    }
    
    function convertExcelToCSV(weekKey) {
        let csvContent = 'Item Name,Have,Need\n';
        
        masterItemList.forEach(itemName => {
            const itemData = inventoryData[weekKey] && inventoryData[weekKey][itemName] 
                ? inventoryData[weekKey][itemName] 
                : { have: '', need: '' };
            
            csvContent += `"${itemName}","${itemData.have}","${itemData.need}"\n`;
        });
        
        return csvContent;
    }
    
    function sendEmailWithLink() {
        const email = emailInput.value.trim();
        const fromName = fromNameInput.value.trim() || 'Inventory Tracker';
        
        // Check if EmailJS is configured
        if (EMAILJS_CONFIG.publicKey === 'YOUR_PUBLIC_KEY') {
            alert('Email service is not configured. Please see the README for setup instructions.');
            return;
        }
        
        const weekDates = getWeekDateRange(currentYear, currentWeek);
        const weekKey = `${currentYear}-W${currentWeek.toString().padStart(2, '0')}`;
        
        // Show loading state
        const sendBtn = document.getElementById('sendEmail');
        const originalText = sendBtn.textContent;
        sendBtn.textContent = 'Preparing data...';
        sendBtn.disabled = true;
        
        try {
            // Try Excel upload first, fallback to CSV in email body
            const fileName = `Inventory_${weekDates.start.replace('/', '-')}_to_${weekDates.end.replace('/', '-')}.xlsx`;
            const excelBuffer = generateExcelBuffer();
            const blob = new Blob([excelBuffer], { 
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
            });
            
            sendBtn.textContent = 'Uploading file...';
            
            // Upload to file.io and send email
            uploadToFileIO(blob, fileName)
                .then(function(downloadUrl) {
                    sendBtn.textContent = 'Sending email...';
                    
                    // Create email with download link
                    const emailBody = `Hello,

Please find your weekly inventory for ${weekDates.start} - ${weekDates.end}.

Click the link below to download the Excel file:
${downloadUrl}

The Excel file contains all inventory items with their current "Have" and "Need" quantities.

Note: This download link will expire after 14 days or after the first download.

Best regards,
${fromName}`;
                    
                    // Email template parameters
                    const templateParams = {
                        to_email: email,
                        from_name: fromName,
                        week_dates: `${weekDates.start} - ${weekDates.end}`,
                        subject: `Weekly Inventory - ${weekDates.start} to ${weekDates.end}`,
                        email_body: emailBody
                    };
                    
                    // Send email using EmailJS
                    return emailjs.send(EMAILJS_CONFIG.serviceId, EMAILJS_CONFIG.templateId, templateParams);
                })
                .then(function(response) {
                    alert(`Email with download link sent to ${email}! The recipient can click the link to download the Excel file.`);
                    emailModal.hide();
                    emailInput.value = '';
                    fromNameInput.value = '';
                })
                .catch(function(error) {
                    console.error('Upload or email failed, trying CSV fallback:', error);
                    
                    // Fallback: Send CSV data in email body
                    sendBtn.textContent = 'Sending CSV data...';
                    
                    const csvData = convertExcelToCSV(weekKey);
                    const emailBody = `Hello,

Here's your weekly inventory for ${weekDates.start} - ${weekDates.end}.

File upload failed, so here's your inventory data in CSV format (you can copy and paste this into Excel):

${csvData}

Best regards,
${fromName}`;
                    
                    const templateParams = {
                        to_email: email,
                        from_name: fromName,
                        week_dates: `${weekDates.start} - ${weekDates.end}`,
                        subject: `Weekly Inventory - ${weekDates.start} to ${weekDates.end}`,
                        email_body: emailBody
                    };
                    
                    return emailjs.send(EMAILJS_CONFIG.serviceId, EMAILJS_CONFIG.templateId, templateParams);
                })
                .then(function(response) {
                    alert(`Email sent to ${email}! (Note: File upload failed, so CSV data was sent in the email body instead)`);
                    emailModal.hide();
                    emailInput.value = '';
                    fromNameInput.value = '';
                })
                .catch(function(finalError) {
                    console.error('All email methods failed:', finalError);
                    alert('Failed to send email. Please check your EmailJS configuration or try the "Send with Excel attachment" option instead.');
                })
                .finally(function() {
                    // Restore button state
                    sendBtn.textContent = originalText;
                    sendBtn.disabled = false;
                });
                
        } catch (error) {
            console.error('Unexpected error:', error);
            alert('An unexpected error occurred. Please try again.');
            sendBtn.textContent = originalText;
            sendBtn.disabled = false;
        }
    }
    
    // Gmail API functions
    function initializeGapi() {
        if (GMAIL_CONFIG.clientId === 'YOUR_GOOGLE_CLIENT_ID') {
            return Promise.reject('Gmail API not configured');
        }
        
        return new Promise((resolve, reject) => {
            gapi.load('auth2:client', () => {
                gapi.client.init({
                    apiKey: GMAIL_CONFIG.apiKey,
                    clientId: GMAIL_CONFIG.clientId,
                    scope: GMAIL_CONFIG.scope,
                    discoveryDocs: ['https://www.googleapis.com/discovery/v1/apis/gmail/v1/rest']
                }).then(() => {
                    const authInstance = gapi.auth2.getAuthInstance();
                    isGmailSignedIn = authInstance.isSignedIn.get();
                    resolve();
                }).catch(reject);
            });
        });
    }
    
    function signInToGmail() {
        return gapi.auth2.getAuthInstance().signIn();
    }
    
    function createEmailMessage(to, subject, body, attachment) {
        const boundary = '-------314159265358979323846';
        const delimiter = "\r\n--" + boundary + "\r\n";
        const close_delim = "\r\n--" + boundary + "--";
        
        let message = delimiter +
            'Content-Type: text/plain; charset="UTF-8"\r\n' +
            'MIME-Version: 1.0\r\n' +
            'Content-Transfer-Encoding: 7bit\r\n\r\n' +
            body + delimiter;
        
        if (attachment) {
            message += 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n' +
                'MIME-Version: 1.0\r\n' +
                'Content-Transfer-Encoding: base64\r\n' +
                'Content-Disposition: attachment; filename="' + attachment.name + '"\r\n\r\n' +
                attachment.data + delimiter;
        }
        
        message += close_delim;
        
        const email = [
            'Content-Type: multipart/mixed; boundary="' + boundary + '"',
            'MIME-Version: 1.0',
            'to: ' + to,
            'subject: ' + subject,
            '',
            message
        ].join('\r\n');
        
        return btoa(unescape(encodeURIComponent(email))).replace(/\+/g, '-').replace(/\//g, '_');
    }
    
    function sendGmailWithAttachment() {
        const email = emailInput.value.trim();
        const fromName = fromNameInput.value.trim() || 'Inventory Tracker';
        const weekDates = getWeekDateRange(currentYear, currentWeek);
        
        const sendBtn = document.getElementById('sendEmail');
        const originalText = sendBtn.textContent;
        sendBtn.textContent = 'Signing in to Gmail...';
        sendBtn.disabled = true;
        
        initializeGapi()
            .then(() => {
                if (!isGmailSignedIn) {
                    return signInToGmail();
                }
            })
            .then(() => {
                sendBtn.textContent = 'Creating email...';
                
                // Generate Excel file
                const excelBuffer = generateExcelBuffer();
                const base64Excel = arrayBufferToBase64(excelBuffer);
                const fileName = `Inventory_${weekDates.start.replace('/', '-')}_to_${weekDates.end.replace('/', '-')}.xlsx`;
                
                const subject = `Weekly Inventory - ${weekDates.start} to ${weekDates.end}`;
                const body = `Hello,

Please find attached your weekly inventory for ${weekDates.start} - ${weekDates.end}.

The Excel file contains all inventory items with their current "Have" and "Need" quantities, plus an "Order From" column for your notes.

Best regards,
${fromName}`;
                
                const attachment = {
                    name: fileName,
                    data: base64Excel
                };
                
                const encodedEmail = createEmailMessage(email, subject, body, attachment);
                
                sendBtn.textContent = 'Sending email...';
                
                return gapi.client.gmail.users.messages.send({
                    userId: 'me',
                    resource: {
                        raw: encodedEmail
                    }
                });
            })
            .then(() => {
                alert(`Email with Excel attachment sent successfully to ${email}!`);
                emailModal.hide();
                emailInput.value = '';
                fromNameInput.value = '';
            })
            .catch((error) => {
                console.error('Gmail sending failed:', error);
                if (error === 'Gmail API not configured') {
                    alert('Gmail API is not configured. Please see the README for setup instructions or use the manual attachment method.');
                } else {
                    alert('Failed to send email via Gmail. Please try the manual attachment method instead.');
                }
            })
            .finally(() => {
                sendBtn.textContent = originalText;
                sendBtn.disabled = false;
            });
    }
    
    function sendEmail() {
        const email = emailInput.value.trim();
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        
        if (!emailRegex.test(email)) {
            alert('Please enter a valid email address');
            return;
        }
        
        // Only use the manual file attachment method
        sendEmailWithFile();
    }
    
    // Initialize with current week's data
    loadWeekData();
});
