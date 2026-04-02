# BlazeMeter CSV to Excel Converter

A Python program that converts BlazeMeter test result CSV files to a formatted Performance Test Result Template in Excel format.

## 📋 Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Output Format](#output-format)
- [Pass/Fail Criteria](#passfail-criteria)
- [Examples](#examples)
- [Troubleshooting](#troubleshooting)

---

## 🎯 Overview

This tool automates the conversion of BlazeMeter CSV test results into a standardized Excel report format. It extracts key performance metrics, applies pass/fail logic based on configurable thresholds, and generates detailed analysis of test results.

---

## ✨ Features

- ✅ **Automated Data Extraction** - Extracts summary metrics and processes individual transactions
- ✅ **Test Type Support** - API (500ms threshold) and UI (2000ms threshold) tests
- ✅ **Comprehensive Reporting** - Calculates pass/fail counts and generates analysis
- ✅ **Professional Formatting** - Color-coded results with proper styling
- ✅ **Smart File Naming** - Includes test type and timestamp: `filename-API-converted-YYYYMMDD-HHMMSS.xlsx`

---

## 📦 Prerequisites

### Required Software
- **Python 3.7+**
- **pip** (Python package installer)

### Required Python Packages
```bash
pandas>=2.0.0
openpyxl>=3.1.0
```

---

## 🚀 Installation

### 1. Install Required Packages

```bash
pip3 install pandas openpyxl
```

**Alternative (if you get permission errors):**
```bash
pip3 install --user pandas openpyxl
```

### 2. Verify Installation

```bash
pip3 list | grep -E "pandas|openpyxl"
```

---

## 💻 Usage

### Basic Syntax

```bash
python3 blazemeter_to_excel.py <input_csv_file> <test_type> [output_xlsx_file]
```

### Parameters

| Parameter | Required | Description | Valid Values |
|-----------|----------|-------------|--------------|
| `input_csv_file` | ✅ Yes | Path to BlazeMeter CSV file | Any valid CSV file path |
| `test_type` | ✅ Yes | Type of performance test | `API` or `UI` |
| `output_xlsx_file` | ❌ No | Custom output filename | Any valid Excel filename |

### Test Types

- **API**: Uses 500ms as P95% threshold
- **UI**: Uses 2000ms as P95% threshold

---

## 📊 Output Format

### Generated Excel File Structure

| Row | Field | Description |
|-----|-------|-------------|
| 3 | Average Throughput (Kbytes/second) | From ALL row: Avg. Bandwidth |
| 4 | Total Hits | From ALL row: # Samples |
| 5 | Average Hits per Second (TPS) | From ALL row: Avg. Hits/s |
| 6 | 95% Response Time (ms) | From ALL row: Avg/P95% |
| 7 | Errors | From ALL row: Error Percentage |
| 10 | Failed count of Transaction | Count of failed transactions |
| 11 | Passed count of Transaction | Count of passed transactions |
| 12 | Result | Overall Pass/Fail (color-coded) |
| 13 | Analysis | Detailed explanation of results |
| 16+ | Transaction Details | All transactions sorted alphabetically |

### Transaction Columns

| Column | Field | Format |
|--------|-------|--------|
| A | Transaction Name | Text |
| B | Count | Integer |
| C | Avg. | Integer (ms) |
| D | P95% | Integer (ms) |
| E | Error% | Decimal |
| F | Result | Pass/Fail (color-coded) |

---

## ✅ Pass/Fail Criteria

### Individual Transaction Pass/Fail

**API Tests:**
- ✅ **Pass**: P95% ≤ 500ms (Blue)
- ❌ **Fail**: P95% > 500ms (Red)

**UI Tests:**
- ✅ **Pass**: P95% ≤ 2000ms (Blue)
- ❌ **Fail**: P95% > 2000ms (Red)

### Overall Result (Row 12)

The overall result is **Fail** if ANY of these conditions are true:

1. ❌ **Error Rate > 1%** (from ALL row)
2. ❌ **P95% exceeds threshold:**
   - API: P95% > 500ms
   - UI: P95% > 2000ms
3. ❌ **Any failed transactions exist**

Otherwise, the result is ✅ **Pass**.

### Analysis Text (Row 13)

**Example Fail Messages:**
- `"Error rate is high (1.77%). There are 24 failed transaction(s)."`
- `"Overall P95% response time is over SLA (662 ms > 500 ms)."`
- `"There are 5 failed transaction(s). Overall P95% response time is over SLA (2500 ms > 2000 ms)."`

**Example Pass Message:**
- `"All metrics are within acceptable thresholds."`

---

## 📝 Examples

### Example 1: Convert with API Test Type

```bash
python3 blazemeter_to_excel.py data.csv API
```

**Output:** `data-API-converted-20260401-162512.xlsx`

### Example 2: Convert with UI Test Type

```bash
python3 blazemeter_to_excel.py data.csv UI
```

**Output:** `data-UI-converted-20260401-162512.xlsx`

### Example 3: Specify Custom Output Name

```bash
python3 blazemeter_to_excel.py data.csv API my_performance_report.xlsx
```

**Output:** `my_performance_report.xlsx`

### Example 4: Convert File in Different Directory

```bash
python3 blazemeter_to_excel.py ./Python/R4R-1x-data.csv API
```

**Output:** `./Python/R4R-1x-data-API-converted-20260401-162512.xlsx`

---

## 🖥️ Expected Console Output

```
Reading CSV file: data.csv
  ✓ Extracted summary data from 'ALL' row
  ✓ Sorted transaction data by 'Element Label' (ascending)
  ✓ Using template: /path/to/Performance-Test-result-Template.xlsx
  ✓ Unmerged all cells for data update
  ✓ Updated summary data in template
  ✓ Added 94 transactions to report
  ✓ Failed transactions: 24
  ✓ Passed transactions: 70
  ✓ Overall result: Fail
  ✓ Analysis: Error rate is high (1.77%). There are 24 failed transaction(s).
  ✓ Successfully created: data-API-converted-20260401-162512.xlsx

  Summary:
    - Total transactions: 94
    - Total hits: 106,991
    - Average throughput: 2980.9 KB/s
    - Error rate: 1.77%
    - Output file size: 13,974 bytes
```

---

## 🔧 Troubleshooting

### Issue: "ModuleNotFoundError: No module named 'pandas'"

**Solution:**
```bash
pip3 install pandas openpyxl
```

### Issue: "FileNotFoundError: Input file not found"

**Solution:** 
- Check that the CSV file path is correct
- Use absolute path or navigate to the correct directory
- Verify the file exists: `ls -la data.csv`

### Issue: "No 'ALL' row found in the CSV file"

**Solution:** 
- Ensure your CSV file is a valid BlazeMeter export
- Check that the CSV contains a row with "ALL" in the "Element Label" column

### Issue: "Invalid test type"

**Solution:** 
- Use either `API` or `UI` (case-insensitive)
- Example: `python3 blazemeter_to_excel.py data.csv API`

### Issue: Permission denied

**Solution:**
```bash
chmod +x blazemeter_to_excel.py
```

---

## 📋 Input File Requirements

### Required CSV Columns

Your BlazeMeter CSV file must contain these columns:

- `Element Label`
- `# Samples`
- `Avg. Response Time (ms)`
- `95% line (ms)`
- `Error Percentage`
- `Avg. Hits/s`
- `Avg. Bandwidth (KBytes/s)`

### Required Data

- Must include a row with `Element Label` = "ALL"
- All numeric columns should contain valid numbers

---

## 💡 Tips and Best Practices

1. **Organize Your Files**
   - Keep CSV files in a dedicated directory
   - Use descriptive filenames for your test results

2. **Version Control**
   - The timestamp in the output filename helps track different test runs
   - Keep original CSV files for reference

3. **Batch Processing**
   - You can create a shell script to process multiple files:
     ```bash
     for file in *.csv; do
         python3 blazemeter_to_excel.py "$file" API
     done
     ```

4. **Review Results**
   - Always review the Analysis row (Row 13) for detailed insights
   - Check individual failed transactions for patterns

5. **Share Reports**
   - The generated Excel files are self-contained and easy to share
   - Include the test type in your report naming for clarity

---

## 📁 Project Structure

```
Scripts/
├── blazemeter_to_excel.py          # Main program
├── README.md                        # This file
├── data.csv                         # Sample input file
└── Python/
    └── Performance-Test-result-Template.xlsx  # Excel template
```

---

## 🤝 Contributing

If you'd like to contribute to this project:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📄 License

This project is for internal use by the Performance Testing Team.

---

## 📞 Support

For questions or issues:
- Check the [Troubleshooting](#troubleshooting) section
- Review the console output for specific error messages
- Verify all prerequisites are installed correctly

---

## 📚 Version History

### Version 1.0 (April 2026)
- ✅ Initial release
- ✅ Support for API and UI test types
- ✅ Automated pass/fail logic
- ✅ Analysis text generation
- ✅ Template-based formatting
- ✅ Color-coded results
- ✅ Integer formatting for Avg. and P95% values

---

**Last Updated:** April 2, 2026  
**Author:** Performance Testing Team  
**Program:** `blazemeter_to_excel.py`
