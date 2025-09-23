# Weekly Inventory Tracker

A web-based application for tracking inventory items on a weekly basis with predefined items and editable quantities. Features week-to-week navigation, Excel export, and email functionality.

## Features

- **Predefined Item List**: Starts with 25 common grocery/household items
- **Weekly Date Ranges**: Shows actual Sunday-Saturday date ranges instead of week numbers
- **Editable Quantities**: Set "Have" and "Need" amounts for each item per week
- **Item Management**: Add new items, edit names, or delete items from the master list
- **Data Persistence**: All data saved in browser's localStorage
- **Export Options**:
  - Export to Excel (.xlsx) format with date-based filenames
  - Email export functionality with real email sending capability

## How to Use

1. **Set Quantities**:
   - Click in the "Have" or "Need" columns and enter numbers
   - Changes are saved automatically

2. **Navigate Weeks**:
   - Use arrow buttons to move between weeks
   - Date ranges show as "Week of MM/DD - MM/DD"

3. **Manage Items**:
   - Click "Add Item" to include custom inventory items
   - Click the edit (pencil) icon to rename items
   - Click the trash icon to remove items

4. **Export Data**:
   - **Excel**: Downloads as `Inventory_MM-DD_to_MM-DD.xlsx`
   - **Email**: Send inventory list via email (requires setup)

## Email Options

The app offers two email methods:

### Option 1: Email with Excel Attachment
- **No setup required** - Works immediately
- Downloads Excel file and opens your email client
- Pre-fills recipient, subject, and message
- You manually attach the downloaded Excel file
- Works with any email client (Outlook, Gmail, etc.)

### Option 2: Automatic Email with Download Link (Recommended)
- **Fully automatic** - No manual attachment needed
- Uploads Excel file to temporary cloud storage (file.io)
- Sends email with direct download link
- Recipient clicks link to download Excel file
- Requires EmailJS setup but sends automatically

### Option 3: Text-Only Email (EmailJS)
For text-only emails, configure EmailJS:

1. **Create EmailJS Account**:
   - Go to [https://www.emailjs.com/](https://www.emailjs.com/)
   - Sign up for a free account

2. **Set up Email Service**:
   - Add an email service (Gmail, Outlook, etc.)
   - Note your Service ID

3. **Create Email Template**:
   - Create a new template with these variables:
     - `{{to_email}}` - Recipient email
     - `{{from_name}}` - Sender name
     - `{{week_dates}}` - Week date range
     - `{{subject}}` - Email subject
     - `{{inventory_table}}` - Formatted inventory data
     - `{{excel_note}}` - Note about Excel export
   - Note your Template ID

4. **Get Public Key**:
   - Go to Account > API Keys
   - Copy your Public Key

5. **Configure the App**:
   - Open `app.js`
   - Replace the EmailJS configuration:
   ```javascript
   const EMAILJS_CONFIG = {
       publicKey: 'your_public_key_here',
       serviceId: 'your_service_id_here', 
       templateId: 'your_template_id_here'
   };
   ```

### Sample EmailJS Template
```
Subject: {{subject}}

Hello,

Here's the weekly inventory for {{week_dates}}:

{{inventory_table}}

{{excel_note}}

Best regards,
{{from_name}}
```

## Technical Details

- Built with HTML5, CSS3, and vanilla JavaScript
- Uses Bootstrap 5 for responsive design
- SheetJS for Excel export functionality
- EmailJS for client-side email sending
- Data stored in browser's localStorage

## Running the Application

Simply open `index.html` in any modern web browser. No server installation required.

## Browser Compatibility

Works in all modern browsers:
- Google Chrome (latest)
- Mozilla Firefox (latest)
- Microsoft Edge (latest)
- Safari (latest)

## File Structure

```
InventoryTracker/
├── index.html          # Main application file
├── app.js             # JavaScript functionality
├── styles.css         # Custom styling
├── server.py          # Optional Python server
└── README.md          # This file
```
