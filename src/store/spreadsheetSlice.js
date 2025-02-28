// src/store/spreadsheetSlice.js
import { createSlice } from '@reduxjs/toolkit';
import { 
  evaluateFormula,
  updateDependentCells
} from '../utils/formulaEvaluator';
import { 
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEIGHT,
  NUM_INITIAL_COLS,
  NUM_INITIAL_ROWS
} from '../utils/constants';

// Helper to generate initial cells
const generateInitialCells = () => {
  const cells = {};
  for (let row = 1; row <= NUM_INITIAL_ROWS; row++) {
    for (let col = 0; col < NUM_INITIAL_COLS; col++) {
      const colLetter = String.fromCharCode(65 + col); // A, B, C, ...
      const cellId = `${colLetter}${row}`;
      cells[cellId] = {
        value: '',
        formula: '',
        formatted: '',
        styles: {}
      };
    }
  }
  return cells;
};

// Helper to generate initial columns
const generateInitialColumns = () => {
  const columns = {};
  for (let col = 0; col < NUM_INITIAL_COLS; col++) {
    const colLetter = String.fromCharCode(65 + col);
    columns[colLetter] = { width: DEFAULT_COLUMN_WIDTH };
  }
  return columns;
};

// Helper to generate initial rows
const generateInitialRows = () => {
  const rows = {};
  for (let row = 1; row <= NUM_INITIAL_ROWS; row++) {
    rows[row] = { height: DEFAULT_ROW_HEIGHT };
  }
  return rows;
};

// Create the initial state
const initialState = {
  cells: generateInitialCells(),
  columns: generateInitialColumns(),
  rows: generateInitialRows(),
  activeCell: null,
  selectedRange: null,
  cellDependencies: {}, // Track dependencies between cells
  isEditing: false,
  showFormulaBar: true,
};

const spreadsheetSlice = createSlice({
  name: 'spreadsheet',
  initialState,
  reducers: {
    // Set active cell
    setActiveCell: (state, action) => {
      state.activeCell = action.payload;
      if (!action.payload.shiftKey) {
        state.selectedRange = null;
      }
    },
    
    // Set the cell value and formula
    setCellContent: (state, action) => {
      const { cellId, content } = action.payload;
      
      // If content starts with '=', it's a formula
      const isFormula = typeof content === 'string' && content.startsWith('=');
      
      // Update the cell
      state.cells[cellId] = {
        ...state.cells[cellId],
        formula: content,
        value: isFormula ? evaluateFormula(content, cellId, state.cells) : content
      };
      
      // Format the displayed value
      state.cells[cellId].formatted = state.cells[cellId].value;
      
      // Update dependent cells
      if (state.cellDependencies[cellId]) {
        updateDependentCells(cellId, state.cells, state.cellDependencies);
      }
    },
    
    // Set cell styles
    setCellStyles: (state, action) => {
      const { cellId, styles } = action.payload;
      state.cells[cellId] = {
        ...state.cells[cellId],
        styles: {
          ...state.cells[cellId].styles,
          ...styles
        }
      };
    },
    
    // Set selected range
    setSelectedRange: (state, action) => {
      state.selectedRange = action.payload;
    },
    
    // Toggle edit mode
    setEditMode: (state, action) => {
      state.isEditing = action.payload;
    },
    
    // Add a new row
    addRow: (state, action) => {
      const { afterRow } = action.payload;
      const newRowNum = afterRow + 1;
      
      // Shift all rows down
      Object.keys(state.rows)
        .filter(row => parseInt(row) > afterRow)
        .sort((a, b) => parseInt(b) - parseInt(a))
        .forEach(rowNum => {
          state.rows[parseInt(rowNum) + 1] = state.rows[rowNum];
          
          // Shift cells down
          Object.keys(state.columns).forEach(col => {
            const oldCellId = `${col}${rowNum}`;
            const newCellId = `${col}${parseInt(rowNum) + 1}`;
            state.cells[newCellId] = state.cells[oldCellId];
          });
        });
      
      // Add new row
      state.rows[newRowNum] = { height: DEFAULT_ROW_HEIGHT };
      
      // Add empty cells for the new row
      Object.keys(state.columns).forEach(col => {
        const cellId = `${col}${newRowNum}`;
        state.cells[cellId] = {
          value: '',
          formula: '',
          formatted: '',
          styles: {}
        };
      });
      
      // Update the highest row number
      state.rows[NUM_INITIAL_ROWS + 1] = { height: DEFAULT_ROW_HEIGHT };
    },
    
    // Add a new column
    addColumn: (state, action) => {
      const { afterCol } = action.payload;
      const afterColCode = afterCol.charCodeAt(0);
      const newColCode = afterColCode + 1;
      const newCol = String.fromCharCode(newColCode);
      
      // Shift all columns to the right
      Object.keys(state.columns)
        .filter(col => col.charCodeAt(0) > afterColCode)
        .sort((a, b) => b.charCodeAt(0) - a.charCodeAt(0))
        .forEach(col => {
          const newColLetter = String.fromCharCode(col.charCodeAt(0) + 1);
          state.columns[newColLetter] = state.columns[col];
          
          // Shift cells right
          Object.keys(state.rows).forEach(row => {
            const oldCellId = `${col}${row}`;
            const newCellId = `${newColLetter}${row}`;
            state.cells[newCellId] = state.cells[oldCellId];
          });
        });
      
      // Add new column
      state.columns[newCol] = { width: DEFAULT_COLUMN_WIDTH };
      
      // Add empty cells for the new column
      Object.keys(state.rows).forEach(row => {
        const cellId = `${newCol}${row}`;
        state.cells[cellId] = {
          value: '',
          formula: '',
          formatted: '',
          styles: {}
        };
      });
      
      // Update the highest column
      const highestColLetter = String.fromCharCode(65 + NUM_INITIAL_COLS);
      state.columns[highestColLetter] = { width: DEFAULT_COLUMN_WIDTH };
    },
    
    // Delete a row
    deleteRow: (state, action) => {
      const { rowNum } = action.payload;
      const maxRow = Math.max(...Object.keys(state.rows).map(r => parseInt(r)));
      
      // Delete the row
      delete state.rows[rowNum];
      
      // Shift all rows up
      for (let r = rowNum; r < maxRow; r++) {
        state.rows[r] = state.rows[r + 1];
        
        // Shift cells up
        Object.keys(state.columns).forEach(col => {
          const cellId = `${col}${r}`;
          const nextCellId = `${col}${r + 1}`;
          state.cells[cellId] = state.cells[nextCellId] || {
            value: '',
            formula: '',
            formatted: '',
            styles: {}
          };
        });
      }
      
      // Delete the last row
      delete state.rows[maxRow];
      
      // Clean up cells in the deleted row
      Object.keys(state.columns).forEach(col => {
        const cellId = `${col}${maxRow}`;
        delete state.cells[cellId];
      });
    },
    
    // Delete a column
    deleteColumn: (state, action) => {
      const { col } = action.payload;
      const colCode = col.charCodeAt(0);
      const maxColCode = Math.max(...Object.keys(state.columns).map(c => c.charCodeAt(0)));
      
      // Delete the column
      delete state.columns[col];
      
      // Shift all columns left
      for (let c = colCode; c < maxColCode; c++) {
        const currentCol = String.fromCharCode(c);
        const nextCol = String.fromCharCode(c + 1);
        state.columns[currentCol] = state.columns[nextCol];
        
        // Shift cells left
        Object.keys(state.rows).forEach(row => {
          const cellId = `${currentCol}${row}`;
          const nextCellId = `${nextCol}${row}`;
          state.cells[cellId] = state.cells[nextCellId] || {
            value: '',
            formula: '',
            formatted: '',
            styles: {}
          };
        });
      }
      
      // Delete the last column
      delete state.columns[String.fromCharCode(maxColCode)];
      
      // Clean up cells in the deleted column
      Object.keys(state.rows).forEach(row => {
        const cellId = `${String.fromCharCode(maxColCode)}${row}`;
        delete state.cells[cellId];
      });
    },
    
    // Resize a row
    resizeRow: (state, action) => {
      const { rowNum, height } = action.payload;
      state.rows[rowNum].height = height;
    },
    
    // Resize a column
    resizeColumn: (state, action) => {
      const { col, width } = action.payload;
      state.columns[col].width = width;
    },
  }
});

export const { 
  setActiveCell, 
  setCellContent, 
  setCellStyles, 
  setSelectedRange, 
  setEditMode,
  addRow, 
  addColumn, 
  deleteRow, 
  deleteColumn,
  resizeRow,
  resizeColumn
} = spreadsheetSlice.actions;

export default spreadsheetSlice.reducer;