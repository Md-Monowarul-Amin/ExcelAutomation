# Accounting Analysis Automation with OpenPyXL

This project automates the process of reading financial data from a file, performing **Vertical Analysis** and **Ratio Analysis**, and generating corresponding graphs using the Python library `openpyxl`.

## Features

- **File Input**: Reads financial data from a file and populates an Excel sheet.
- **Vertical Analysis**: Calculates percentages of individual line items relative to a base figure in the financial statements.
- **Ratio Analysis**: Computes common financial ratios such as profitability, liquidity, and solvency ratios.
- **Graph Generation**: Automatically generates charts and graphs to visualize the analysis results.

## Technologies Used

- **Python**
- **OpenPyXL**: For reading, writing, and modifying Excel files, and generating graphs.

## Key Functionalities

### Vertical Analysis
Vertical analysis expresses each item in a financial statement as a percentage of a base amount. For example:
- In the income statement, each line item is calculated as a percentage of total revenue.
- In the balance sheet, each asset, liability, or equity item is calculated as a percentage of total assets.

### Ratio Analysis
The script calculates key financial ratios, such as:
- **Profitability Ratios**: Net profit margin, gross profit margin.
- **Liquidity Ratios**: Current ratio, quick ratio.
- **Solvency Ratios**: Debt-to-equity ratio.

### Graphs and Visualizations
The program uses `openpyxl` to embed graphs in the Excel sheet for:
- Income statement breakdowns.
- Comparative balance sheet analysis.
- Ratio trends.



---

### Author
Md. Monowarul Amin
