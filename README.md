# Google Sheets Clone

A web application that mimics the UI and functionality of Google Sheets, with a focus on mathematical functions, data quality functions, and key interactions.

## Features

- Spreadsheet interface with Google Sheets-like UI
- Cell formatting (bold, italic, font size, color)
- Row and column resizing
- Mathematical functions: SUM, AVERAGE, MAX, MIN, COUNT
- Data quality functions: TRIM, UPPER, LOWER, REMOVE_DUPLICATES, FIND_AND_REPLACE
- Cell dependency handling
- Drag selection

## Data Structures

### Cell Data Model

The application uses a normalized data structure for efficient updates and rendering:

```javascript
{
  cells: {
    "A1": { value: "10", formula: "10", formatted: "10", styles: {} },
    "B1": { value: "20", formula: "20", formatted: "20", styles: {} },
    "C1": { value: "30", formula: "=SUM(A1:B1)", formatted: "30", styles: {} }
  },
  columns: { "A": { width: 100 }, "B": { width: 100 } },
  rows: { "1": { height: 25 }, "2": { height: 25 } }
}