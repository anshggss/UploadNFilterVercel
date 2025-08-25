# Excel Data Sorter Backend

This backend receives an Excel file, filters it according to your rules, and returns the filtered file.

## Setup

1. Install dependencies:
   ```
   npm install
   ```
2. Start the server:
   ```
   npm start
   ```

The server will run on port 5000 by default.

## API

### POST /api/filter
- Accepts: Multipart/form-data with a file field named `file` (the raw Excel file)
- Returns: Filtered Excel file as an attachment

## Filtering Logic
- The filtering logic is a placeholder. Update `filterExcel.js` to implement your rules. 