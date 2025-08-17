# FastAPI CRUD with Excel Export

A FastAPI application with CRUD operations and Excel export functionality.

## Features

- ✅ Create, Read, Update, Delete (CRUD) operations
- ✅ Dark theme UI
- ✅ Edit items with modal
- ✅ Download Excel template
- ✅ Export all data to Excel
- ✅ Upload Excel files to import data
- ✅ Select all checkbox and group delete functionality

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
uvicorn main:app --reload
```

3. Open your browser and go to: `http://localhost:8000`

## Excel Functionality

### Download Excel Template
- Click the green "📋 Download Excel Template" button
- Downloads an Excel file with headers: Name, Description
- Includes sample data row for reference
- Use this template to prepare data for import

### Download All Data
- Click the blue "📥 Download All Data" button
- Downloads an Excel file with all current items from the database
- Includes headers and all data rows
- Useful for backup and analysis

### Upload Excel Data
- Use the purple "📤 Upload Excel File" button in the Import section
- Select an Excel (.xlsx) file with your data
- Data will be imported as new items (IDs are auto-generated)
- Supports the same format as the template (Name, Description)

## API Endpoints

- `GET /` - Main page with CRUD interface
- `GET /health` - Health check endpoint
- `GET /favicon.ico` - Favicon endpoint (prevents 404 errors)
- `GET /robots.txt` - Robots.txt endpoint (prevents 404 errors)
- `POST /items/` - Create new item
- `GET /items/` - Get all items
- `PUT /items/{item_id}` - Update item
- `DELETE /items/{item_id}` - Delete item
- `DELETE /items/group` - Delete multiple items (requires JSON body: `{"item_ids": [1, 2, 3]}`)
- `GET /download-excel-template` - Download Excel template
- `GET /download-excel-data` - Download all data as Excel
- `POST /upload-excel` - Upload Excel file to import data

## File Structure

```
fastapi/
├── main.py              # FastAPI application with all endpoints
├── models.py            # SQLAlchemy models
├── schemas.py           # Pydantic schemas
├── database.py          # Database configuration
├── templates/
│   └── index.html      # HTML template with dark theme
├── requirements.txt     # Python dependencies
└── README.md           # This file
```

## How It Works

1. **Template Download**: Creates a new Excel workbook with headers and sample data
2. **Data Export**: Fetches all items from database and creates Excel file
3. **Styling**: Applies professional styling with blue headers and auto-adjusted column widths
4. **File Response**: Returns Excel files as downloadable attachments

## Notes

- Excel files are created using the `openpyxl` library
- Files are temporarily stored and automatically cleaned up
- Column widths are automatically adjusted based on content
- Professional styling with blue headers and proper formatting
