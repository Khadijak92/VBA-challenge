# Stock Market Analysis VBA Script

## Introduction

This VBA script is a comprehensive tool for analyzing stock market data across multiple worksheets in an Excel workbook. The script loops through all worksheets, calculates yearly changes, percentage changes, and total stock volume for each stock ticker, and generates a comprehensive summary table. It also highlights positive and negative changes in the "Percentage Change" column and identifies stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

## Features

- **Data Processing:**
  - Calculates yearly changes, percentage changes, and total stock volume.
  - Populates a summary table with Ticker, Yearly Change, Percentage Change, and Total Stock Volume.

- **Conditional Formatting:**
  - Applies conditional formatting to highlight positive and negative percentage changes.

- **Maximum Values:**
  - Identifies stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

## Getting Started

### Prerequisites

Ensure that you have an Excel workbook with data organized in columns for Ticker, Opening Price, Closing Price, and Volume.
Test the script on a backup or sample data before applying it to your actual dataset.

### Installation

1. Open the Excel workbook where you want to perform the stock market analysis.

2. Press `ALT` + `F11` to open the VBA editor.

3. Insert a new module (Right-click on any item in the project explorer, choose `Insert` -> `Module`).

4. Copy and paste the provided VBA script into the module.

5. Close the VBA editor.

6. Run the script by pressing `ALT` + `F8`, selecting `StockMarket`, and clicking `Run`.

## Usage

Run the script to automate the stock market analysis across all worksheets in the Excel workbook. The results will be displayed in a summary table, and cells in the "Percentage Change" column will be color-coded for better visualization.

## Customization
- Modify the script to accommodate variations in your data structure.

## Contributing

Contributions are welcome! If you have any suggestions, bug reports, or improvements, please open an issue or submit a pull request following the guidelines in the [Contributing](CONTRIBUTING.md) file.

## License



## Acknowledgments

- The script structure and initial implementation were inspired by .


