# Google Sheets Clone

A web application that mimics the UI and functionality of Google Sheets, focusing on mathematical functions, data quality functions, and key interactions. This project was developed as part of Assignment 1 to create a functional spreadsheet application with core Google Sheets features.

## Table of Contents
1. [Project Overview](#project-overview)
2. [Features](#features)
3. [Installation and Setup](#installation-and-setup)
4. [Tech Stack and Architecture](#tech-stack-and-architecture)
5. [Data Structures](#data-structures)
6. [Functionality Guide](#functionality-guide)
7. [Testing](#testing)
8. [Bonus Features Implemented](#bonus-features-implemented)

## Project Overview

This application strives to recreate the Google Sheets experience with key spreadsheet functionalities. It includes a familiar interface with toolbars, formula editing, cell manipulation, and various functions for data analysis and manipulation.

## Features

### User Interface
- Complete Google Sheets-like UI with toolbar, formula bar, and cell grid
- Draggable cell selections and content
- Multiple sheets support with tabs
- Context menus for quick actions

### Cell Operations
- Cell selection (single, range, multiple ranges)
- Drag to select ranges
- Copy, cut, and paste functionality
- Cell formatting (bold, italic, underline, colors, font sizes, font families)

### Row and Column Management
- Add and delete rows/columns
- Resize rows and columns
- Header navigation

### Mathematical Functions
- `SUM`: Calculate the sum of a range of cells
- `AVERAGE`: Calculate the average of a range of cells
- `MAX`: Find the maximum value in a range
- `MIN`: Find the minimum value in a range
- `COUNT`: Count numeric cells in a range
- `COUNTIF`: Count cells that meet specific criteria

### Data Quality Functions
- `TRIM`: Remove leading and trailing whitespace
- `UPPER`: Convert text to uppercase
- `LOWER`: Convert text to lowercase
- `CONCATENATE`: Join text from multiple cells
- Remove duplicates from a range
- Find and replace text across cells

### Logical Functions
- `IF`: Conditional logic with true/false outcomes

### Advanced Features
- Cell dependency tracking and automatic recalculation
- Absolute and relative cell references ($A$1, A1)
- Multiple number formats (percentage, currency)
- Cell comments
- Chart creation and visualization
- Undo/redo functionality

## Installation and Setup

### Prerequisites
- Node.js (v14.0.0 or higher)
- npm (v6.0.0 or higher)

### Installation Steps

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/google-sheets-clone.git
   cd google-sheets-clone
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Start the development server:
   ```bash
   npm start
   ```

4. Open your browser and navigate to:
   ```
   http://localhost:3000
   ```

### Building for Production
```bash
npm run build
```

## Tech Stack and Architecture

### Tech Stack
- **React**: Frontend framework for building the user interface
- **JavaScript (ES6+)**: Core programming language
- **CSS3**: Styling and layout
- **Papaparse**: CSV parsing library
- **SheetJS (XLSX)**: Excel file handling
- **Recharts**: Chart visualization

### Why This Tech Stack?
- **React**: Provides component-based architecture ideal for a complex UI with many reusable elements like cells, rows, and toolbars. React's state management capabilities help handle the complex data flow in a spreadsheet application.
- **Pure JavaScript**: Used for formula evaluation and cell dependency management to maintain control over the complex logic without unnecessary abstractions.
- **Papaparse & SheetJS**: These libraries provide robust file format support without requiring backend services.
- **Recharts**: Offers flexible chart creation with React integration for data visualization.

## Data Structures

The application uses carefully designed data structures for efficient updates and rendering:

### Cell Data Model
```javascript
{
  cells: {
    "A1": { 
      value: "10",       // The evaluated value
      formula: "10",     // The raw input (with = for formulas)
      formatted: "10",   // The display value after formatting
    },
    "B1": { 
      value: "20", 
      formula: "20", 
      formatted: "20"
    },
    "C1": { 
      value: "30", 
      formula: "=SUM(A1:B1)", 
      formatted: "30"
    }
  }
}
```

### Cell Styles Model
```javascript
{
  cellStyles: {
    "A1": {
      fontWeight: "bold",
      color: "#FF0000",
      backgroundColor: "#F5F5F5",
      textAlign: "center",
      fontSize: "12px",
      fontFamily: "Arial",
      textDecoration: "underline",
      fontStyle: "italic",
      numberFormat: "percentage"
    }
  }
}
```

### Dependencies Tracking
```javascript
{
  dependencies: {
    "A1": ["C1", "D1"],  // Cells that depend on A1
    "B1": ["C1"]         // Cells that depend on B1
  }
}
```

### Sheet Data Structure
```javascript
{
  sheetData: {
    "Sheet1": {
      cells: { /* cell data */ },
      cellStyles: { /* style data */ },
      dependencies: { /* dependency data */ }
    },
    "Sheet2": {
      cells: { /* cell data */ },
      cellStyles: { /* style data */ },
      dependencies: { /* dependency data */ }
    }
  }
}
```

### Grid Structure
```javascript
{
  visibleColumns: ["A", "B", "C", "D", "E", "F", "G", "H"],
  visibleRows: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
}
```

### Charts Data Structure
```javascript
{
  charts: [
    {
      id: "chart-123456",
      type: "bar",         // bar, line, pie, area
      title: "Sales Data",
      data: [...],         // Processed data for the chart
      xAxisKey: "month",   // X-axis data field
      yAxisKeys: ["sales", "revenue"], // Y-axis data fields
      stacked: false,      // Whether the chart is stacked
      position: {
        top: 100,
        left: 100,
        width: 500,
        height: 300
      }
    }
  ]
}
```

### Why This Data Structure?
1. **Normalized State**: By using object structures with cell IDs as keys, we achieve O(1) lookup time for any cell.
2. **Separation of Concerns**: Keeping cell data, styles, and dependencies separate allows targeted updates without re-rendering the entire grid.
3. **Formula Evaluation**: The separation of `value`, `formula`, and `formatted` properties enables clean formula processing and display.
4. **Efficient Updates**: When a cell changes, we can quickly identify dependent cells through the dependency graph.
5. **Multiple Sheet Support**: The sheet data structure allows for multiple sheets while maintaining the same interface for each sheet.

## Functionality Guide

### Basic Navigation
- Use arrow keys to navigate between cells
- Click on a cell to select it
- Click and drag to select a range of cells
- Use Tab to move right, Enter to move down

### Entering Data
- Click on a cell and start typing
- Press Enter to confirm and move down
- Press Tab to confirm and move right
- Press Escape to cancel editing

### Using Formulas
- Start with an equal sign (=)
- Reference cells by their column and row (A1, B2)
- Use functions like `=SUM(A1:A10)` or `=AVERAGE(B1:B5)`
- Combine functions with operators: `=SUM(A1:A5)/COUNT(A1:A5)`

### File Operations
- Use File menu to create, open, and save spreadsheets
- Import CSV and Excel files
- Export to CSV

### Formatting
- Use the toolbar to apply bold, italic, and underline
- Change text color and background color
- Adjust text alignment and font size

### Working with Rows and Columns
- Right-click for context menu options
- Insert or delete rows and columns
- Use the "+" at the end of rows/columns to add more

### Creating Charts
- Select data range
- Click Insert > Chart
- Choose chart type and customize as needed
- Drag to position on the sheet

### Using Functions
- Mathematical: `SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNT`
- Text: `TRIM`, `UPPER`, `LOWER`, `CONCATENATE`
- Logical: `IF`, `COUNTIF`

### Keyboard Shortcuts
- Ctrl+B: Bold
- Ctrl+I: Italic
- Ctrl+U: Underline
- Ctrl+Z: Undo
- Ctrl+Y: Redo
- Ctrl+C: Copy
- Ctrl+X: Cut
- Ctrl+V: Paste
- F2: Edit cell

## Testing

The application includes comprehensive testing for all functions:

1. **Mathematical Functions**:
   - Test SUM with positive, negative, and mixed values
   - Test AVERAGE with different data sets
   - Verify MAX and MIN across ranges
   - Test COUNT with mixed data types

2. **Data Quality Functions**:
   - Verify TRIM removes expected whitespace
   - Test UPPER and LOWER with mixed case text
   - Test CONCATENATE with various inputs
   - Verify FIND_AND_REPLACE with different patterns

3. **Cell Dependencies**:
   - Test formula recalculation when referenced cells change
   - Verify formula error handling

4. **UI Testing**:
   - Test cell selection and range selection
   - Verify drag behavior for selections
   - Test copy/paste functionality
   - Verify formatting application

## Bonus Features Implemented

1. **Advanced Cell References**:
   - Support for absolute references ($A$1)
   - Mixed references ($A1, A$1)

2. **File Operations**:
   - Save/load spreadsheets (.json format)
   - Import CSV and Excel files
   - Export to CSV

3. **Data Visualization**:
   - Line charts
   - Bar charts
   - Pie charts
   - Area charts

4. **Enhanced UI Features**:
   - Multiple sheets support
   - Cell comments
   - Conditional formatting
   - Custom number formats

5. **Additional Functions**:
   - IF for conditional logic
   - COUNTIF for conditional counting
