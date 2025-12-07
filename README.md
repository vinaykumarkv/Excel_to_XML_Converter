# Excel to XML Converter - User Manual

## Table of Contents
1. [Overview](#overview)
2. [Installation](#installation)
3. [Getting Started](#getting-started)
4. [Creating XML Mappings](#creating-xml-mappings)
5. [Generating XML Files](#generating-xml-files)
6. [Digital Signature Features](#digital-signature-features)
7. [Template Management](#template-management)
8. [Troubleshooting](#troubleshooting)

---

## Overview

**Excel to XML Converter** is a desktop application that converts Excel spreadsheet data into structured XML files with digital signature capabilities. It's designed for creating XML-based reports.

### Key Features
- ✅ Convert Excel data to XML format
- ✅ Flexible field mapping with three types: Single Fields, Repeated Blocks, and Nested Blocks
- ✅ Digital signature support (DSA-SHA1)
- ✅ Template system to save and reuse mappings
- ✅ Handle merged cells automatically
- ✅ XML signature verification

---

## Installation

### System Requirements
- **Operating System**: Windows 10 or later
- **Memory**: 4 GB RAM minimum
- **Disk Space**: 100 MB free space

### Installation Steps

1. **Download** the installer file: `ExcelToXMLConverter_Setup.exe`

2. **Run the installer** by double-clicking the file

3. **Follow the installation wizard**:
   - Accept the license agreement
   - Choose installation location (default recommended)
   - Select whether to create a desktop shortcut
   - Click "Install"

4. **Launch the application** from:
   - Desktop shortcut (if created)
   - Start Menu → Excel to XML Converter

---

## Getting Started

### Application Interface

When you open the application, you'll see:

```
┌─────────────────────────────────────────────────┐
│     Excel to Signed XML Generator               │
├─────────────────────────────────────────────────┤
│  [No Excel file selected]    [Browse Excel]     │
├─────────────────────────────────────────────────┤
│                                                 │
│  (Mapping Configuration Area)                   │
│                                                 │
├─────────────────────────────────────────────────┤
│  [+Single] [+Repeated] [+Nested] [Save Template]│
│  [Load Template] [Sign XML] [Verify XML]        │
│                              [GENERATE XML] ➜  │
└─────────────────────────────────────────────────┘
```

---

## Creating XML Mappings

### Step 1: Select Your Excel File

1. Click **"Browse Excel"** button
2. Navigate to your Excel file (.xlsx or .xls)
3. Select the file and click "Open"
4. The filename will appear at the top of the window

### Step 2: Add Field Mappings

There are three types of mappings you can create:

---

#### Type 1: Single Field

**Use for**: Individual data points that appear once in your XML

**Example**: Document title, date, company name

**How to add**:
1. Click **"+ Single Field"** button
2. Fill in the fields:
   - **Tag name**: The XML element name (e.g., "DocumentTitle")
   - **Row**: Excel row number (e.g., "5")
   - **Col**: Excel column number (e.g., "2" for column B)
   - **Fixed value**: Optional - use this instead of Excel cell if you want a constant value

**Example**:
```
Tag name: CompanyName
Row: 3
Col: 2
Fixed value: (leave empty to use cell B3)
```

This creates:
```xml
<CompanyName>Acme Corporation</CompanyName>
```

---

#### Type 2: Repeated Block

**Use for**: Data that repeats across multiple rows (like a table)

**Example**: Bat analysis data, test results, product listings

**How to add**:
1. Click **"+ Repeated Block"** button
2. Configure the block:
   - **Block name**: Parent element name (e.g., "BatAnalysis")
   - **Rows - from**: Starting row number (e.g., "2")
   - **Rows - to**: Ending row number (leave blank for auto-detect)
3. Click **"+ Add Field"** to add columns:
   - **Field**: XML element name (e.g., "BatNumber")
   - **Col**: Column number (e.g., "1" for column A)
   - **Offset**: Row offset (usually "0", use "1" if data is one row below)

**Example**:
```
Block name: BatAnalysis
Rows: 2 to (blank)

Fields:
  - Field: BatNumber, Col: 1, Offset: 0
  - Field: TestResult, Col: 2, Offset: 0
  - Field: Purity, Col: 3, Offset: 0
```

This reads rows 2, 3, 4... until an empty row and creates:
```xml
<BatAnalysis>
  <BatNumber>B001</BatNumber>
  <TestResult>Pass</TestResult>
  <Purity>99.5</Purity>
</BatAnalysis>
<BatAnalysis>
  <BatNumber>B002</BatNumber>
  <TestResult>Pass</TestResult>
  <Purity>99.8</Purity>
</BatAnalysis>
```

---

#### Type 3: Nested Block

**Use for**: Grouped information under one parent element

**Example**: Signature information, address blocks, metadata

**How to add**:
1. Click **"+ Nested Block"** button
2. Configure:
   - **Block name**: Parent element name (e.g., "Signature")
3. Click **"+ Add Sub-tag"** for each child element:
   - **Tag**: XML element name (e.g., "SignedBy")
   - **Fixed value**: Constant text OR leave empty to use cell
   - **Row**: Excel row (if not using fixed value)
   - **Col**: Excel column (if not using fixed value)

**Example**:
```
Block name: Signature

Sub-tags:
  - Tag: SignedBy, Fixed value: John Doe
  - Tag: Date, Row: 10, Col: 5
  - Tag: Department, Fixed value: Quality Control
```

This creates:
```xml
<Signature>
  <SignedBy>John Doe</SignedBy>
  <Date>2024-12-07</Date>
  <Department>Quality Control</Department>
</Signature>
```

---

### Removing Mappings

Click the **"X"** button on any mapping row to delete it.

---

## Generating XML Files

### Step-by-Step Process

1. **Select Excel file** (if not already selected)

2. **Configure all mappings** (Single Fields, Repeated Blocks, Nested Blocks)

3. Click the **"GENERATE XML"** button (large orange button on the right)

4. **Choose save location**:
   - A file dialog will open
   - Enter your filename (e.g., "report.xml")
   - Click "Save"

5. **Success!** You'll see a confirmation message with the file path

### Understanding Row and Column Numbers

- **Rows** are numbered: 1, 2, 3, 4...
- **Columns** are numbered: 1 (A), 2 (B), 3 (C), 4 (D)...

**Example**:
```
     A    B    C    D
1  | ID | Name | Age |
2  | 1  | John | 25  |
3  | 2  | Jane | 30  |
```

Cell "John" is at **Row: 2, Col: 2**

---

## Digital Signature Features

### Why Use Digital Signatures?

Digital signatures:
- ✅ Prove the XML hasn't been tampered with
- ✅ Verify the document's authenticity
- ✅ Meet compliance requirements for documents

---

### Signing an XML File

1. **Generate your XML file first** (see above)

2. Click the **"Sign XML"** button

3. **Select the XML file** you want to sign

4. The application will:
   - Automatically generate encryption keys (first time only)
   - Sign the document with DSA-SHA1
   - Save as `filename_signed.xml`

5. **Success!** A confirmation dialog shows where the signed file was saved

**Note**: Keys are generated automatically and reused during the same session. You don't need to generate keys manually.

---

### Verifying a Signed XML File

1. Click the **"Verify XML"** button

2. **Select the signed XML file** (usually ending in `_signed.xml`)

3. Results:
   - ✅ **"Valid"**: Signature is authentic and document hasn't been modified
   - ❌ **"Invalid"**: Signature verification failed or document was altered

**Important**: Verification only works for files signed by the same application instance. To verify files across different computers, you'll need to share keys (advanced feature).

---

## Template Management

Templates save your mapping configuration so you don't have to recreate it every time.

### Saving a Template

1. **Configure all your mappings** (Single, Repeated, Nested blocks)

2. Click **"Save Template"** button

3. **Choose location and name**:
   - File dialog opens
   - Name your template (e.g., "monthly_report_template.json")
   - Click "Save"

4. **Confirmation** appears when saved successfully

### Loading a Template

1. Click **"Load Template"** button

2. **Select your template file** (.json)

3. All mappings are restored instantly

4. You can now:
   - Select a different Excel file
   - Click "GENERATE XML"

**Tip**: Create templates for recurring reports to save time!

---

## Common Workflows

### Workflow 1: One-Time XML Generation

```
1. Browse Excel → Select file
2. Add Single Field(s)
3. Add Repeated Block(s) if needed
4. Add Nested Block(s) if needed
5. Generate XML → Choose save location
6. Done!
```

---

### Workflow 2: Recurring Reports (Using Templates)

```
First Time:
1. Browse Excel → Select sample file
2. Configure all mappings
3. Save Template → Name it
4. Generate XML

Every Time After:
1. Load Template → Select saved template
2. Browse Excel → Select new data file
3. Generate XML
4. Done!
```

---

### Workflow 3: Signed XML for Compliance

```
1. Browse Excel → Select file
2. Load Template (or configure mappings)
3. Generate XML → Save as "report.xml"
4. Sign XML → Select "report.xml"
5. Signed file saved as "report_signed.xml"
6. (Optional) Verify XML → Select signed file
7. Done!
```

---

## Troubleshooting

### Issue: "Select Excel file first!"

**Solution**: Click "Browse Excel" and select an Excel file before generating XML.

---

### Issue: Generated XML is empty or missing data

**Causes**:
- Wrong row/column numbers
- Excel file has data in different location

**Solutions**:
1. Open your Excel file and verify row/column positions
2. Check that row numbers match (Excel's row 1 = 1, not 0)
3. Check that column numbers are correct (A=1, B=2, C=3...)
4. For repeated blocks, ensure the starting row is correct

---

### Issue: Repeated block only shows one item

**Solutions**:
1. Check "Rows - to" field (leave blank for auto-detect)
2. Ensure Excel has data in subsequent rows
3. Verify the first column isn't empty (stops at first empty cell in column 1)

---

### Issue: Merged cells showing wrong data

The application handles merged cells automatically. It reads from the top-left cell of any merged range.

---

### Issue: "Signing Failed" error

**Solutions**:
1. Ensure the XML file exists and isn't corrupted
2. Close any other programs using the XML file
3. Check you have write permissions to the folder
4. Try saving to a different location (e.g., Desktop)

---

### Issue: Template won't load

**Causes**:
- Template file corrupted
- Template created with different version

**Solutions**:
1. Try creating mappings manually
2. Save a new template
3. Contact support if issue persists

---

### Issue: Application won't start

**Solutions**:
1. Restart your computer
2. Reinstall the application
3. Check Windows Event Viewer for error details
4. Ensure you have administrator rights

---

## Tips & Best Practices

### 💡 Tip 1: Test with Small Data First
Before processing large Excel files, test with a small sample to verify mappings are correct.

### 💡 Tip 2: Use Fixed Values for Constants
If certain XML elements always have the same value (like company name), use "Fixed value" instead of row/col.

### 💡 Tip 3: Name Templates Descriptively
Use names like "monthly_report_v1.json" or "bat_analysis_2024.json" for easy identification.

### 💡 Tip 4: Keep Excel Structure Consistent
For templates to work across multiple files, keep the same row/column structure in all Excel files.

### 💡 Tip 5: Verify After Signing
Always verify signed files to ensure the signature was applied correctly.

### 💡 Tip 6: Back Up Templates
Save template files to a secure location. They're much faster than recreating mappings!

---

## File Formats

### Supported Excel Formats
- `.xlsx` (Excel 2007 and later)
- `.xls` (Excel 97-2003)

### Output Format
- `.xml` (UTF-8 encoded)

### Template Format
- `.json` (JSON configuration file)

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl + O` | Browse Excel file |
| `Ctrl + S` | Save template |
| `Ctrl + L` | Load template |
| `Ctrl + G` | Generate XML |

---

## Support & Contact

### Need Help?

**Email**: vinaykumar.kv@outlook.com  
**Website**: https://vinaykumarkv.github.io   


### Reporting Bugs

Please include:
- Windows version
- Application version
- Steps to reproduce the issue
- Screenshots if applicable
- Sample Excel file (if possible)

---

## Version History

### Version 1.0 (Current)
- Initial release
- Excel to XML conversion
- Three mapping types (Single, Repeated, Nested)
- DSA-SHA1 digital signatures
- Template save/load functionality
- Merged cell support

---

## License

This software is provided "as is" without warranty of any kind. See LICENSE for full terms.

---

**Thank you for using Excel to XML Converter!**


---

*Last Updated: December 2025*
