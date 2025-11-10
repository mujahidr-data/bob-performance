# HiBob Field Updater - Google Apps Script

Comprehensive automation tool for managing employee data in HiBob (Bob) HR platform via Google Sheets.

## Overview

The HiBob Field Updater is a Google Apps Script tool that allows you to:
- Pull employee data, field metadata, and list values from HiBob API
- Update single fields for multiple employees in bulk
- Update history tables (Salary, Work History, Variable Pay)
- Validate data before uploading to prevent errors
- Handle rate limiting and batch processing automatically

## Features

### Field Uploader
- **Search & Select Fields**: Search through 38,000+ available fields with real-time filtering
- **Bulk Updates**: Update a single field for multiple employees in one operation
- **Data Validation**: Built-in validation before uploading changes
- **Batch Processing**: Handles large datasets (40-1000+ rows) automatically

### History Tables
- **Salary/Payroll History**: Update employee salary history
- **Work History**: Update job title, department, site changes
- **Variable Pay**: Update bonus and commission history

### Employee Management
- **Filtered Pulls**: Pull employees by status (Active/Inactive), site, and location
- **CIQ ID Mapping**: Automatic mapping between CIQ IDs and Bob IDs
- **Terminated Employees**: Support for including terminated employees in updates

## Files

- **bob-fieldupdater.gs**: Main Apps Script file with all functionality
- **FieldSelector.html**: HTML template for the field selection dialog

## Setup

### 1. Install clasp (if not already installed)

```bash
npm install -g @google/clasp
```

### 2. Login to clasp

```bash
clasp login
```

### 3. Clone/Push to Apps Script

The project is already configured with `.clasp.json`. To push updates:

```bash
clasp push
```

### 4. Configure Credentials

1. Open the Apps Script project
2. Run `setBobServiceUser(id, token)` function
3. Enter your HiBob service user ID and token
4. Test connection with `testBobConnection()`

## Usage

### Initial Setup (One-Time)

1. **Set Credentials**: Run `setBobServiceUser()` in Apps Script editor
2. **Pull Fields**: Bob menu → SETUP → 1. Pull Fields
3. **Pull Lists**: Bob menu → SETUP → 2. Pull Lists
4. **Pull Employees**: Bob menu → SETUP → 3. Pull Employees

### Standard Workflow

1. **Select Field**: Bob menu → SETUP → 5. Select Field to Update
2. **Enter Data**: Paste CIQ IDs and new values in Uploader sheet (starting row 8)
3. **Validate**: Bob menu → VALIDATE → Validate Field Upload Data
4. **Upload**: 
   - Quick Upload (<40 rows): Bob menu → UPLOAD → Quick Upload
   - Batch Upload (40-1000+ rows): Bob menu → UPLOAD → Batch Upload

### History Tables

1. **Setup**: Bob menu → SETUP → 6. History Tables → Setup History Uploader
2. **Select Table Type**: Choose Salary/Payroll, Work History, or Variable Pay
3. **Generate Columns**: Bob menu → SETUP → 6. History Tables → Generate Columns for Table
4. **Enter Data**: Paste data starting at row 14
5. **Validate & Upload**: Same as regular field updates

## Menu Structure

```
Bob
├── >> View Documentation
├── >> Test API Connection
├── SETUP
│   ├── 1. Pull Fields
│   ├── 2. Pull Lists
│   ├── 3. Pull Employees
│   ├── 4. Setup Field Uploader
│   ├── 5. Select Field to Update
│   └── 6. History Tables
├── VALIDATE
│   ├── Validate Field Upload Data
│   └── Validate History Upload Data
├── UPLOAD
│   ├── Quick Upload (<40 rows)
│   ├── Batch Upload (40-1000+ rows)
│   └── Retry Failed Rows Only
├── MONITORING
│   └── Check Batch Status
├── CONTROL
│   └── Stop Batch Upload
└── CLEANUP
    └── Clear All Upload Data
```

## API Rate Limits

- **PUT /v1/people/{id}**: 10 requests per minute (auto-delayed 6 seconds between requests)
- **Batch Processing**: 45 rows every 5 minutes for large datasets
- **Quick Upload**: Best for <40 rows (~4 minutes max)

## Documentation

Run `createDocumentationSheet()` or use Bob menu → >> View Documentation to create a comprehensive guide sheet in your Google Spreadsheet.

## Requirements

- Google Apps Script access
- HiBob API credentials (Service User ID and Token)
- Appropriate permissions for HiBob API endpoints

## Version

**Version 2.0** - Last Updated: 2025-10-29

## Support

For issues or questions:
1. Check the documentation sheet created by `createDocumentationSheet()`
2. Review Apps Script execution logs
3. Check the troubleshooting section in the documentation

