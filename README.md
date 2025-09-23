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
   - **Email**: Send inventory list via email 
