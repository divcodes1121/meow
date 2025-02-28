/* eslint-disable no-restricted-globals */
import React, { useState, useEffect, useRef } from 'react';
import './App.css';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';

import ChartDialog from './ChartDialog';
import { 
  LineChartComponent, 
  BarChartComponent, 
  PieChartComponent, 
  AreaChartComponent 
} from './ChartComponents';

function App() {
  // ======== STATE VARIABLES ========
  // Menu state
  const [activeMenu, setActiveMenu] = useState(null);

  // Add to your existing state variables section
const [sheetData, setSheetData] = useState({
  Sheet1: {
    cells: {},
    cellStyles: {},
    dependencies: {}
  }
});


// Chart state
const [charts, setCharts] = useState([]);
const [showChartDialog, setShowChartDialog] = useState(false);
const [activeChart, setActiveChart] = useState(null);

  // Cell data state
  const [cells, setCells] = useState({});
  const [cellStyles, setCellStyles] = useState({});
  const [dependencies, setDependencies] = useState({});
  const [activeCell, setActiveCell] = useState(null);
  const [selectedRange, setSelectedRange] = useState(null);
  const [inputValue, setInputValue] = useState("");
  const [editingCell, setEditingCell] = useState(null);
  const cellInputRef = useRef(null);
  
  // Grid structure state
  const [visibleColumns, setVisibleColumns] = useState(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']);
  const [visibleRows, setVisibleRows] = useState(Array.from({ length: 25 }, (_, i) => i + 1));
  // Removed unused variables
  const gridRef = useRef(null);
  
  // Sheets and document state
  const [sheets, setSheets] = useState(['Sheet1']);
  const [activeSheet, setActiveSheet] = useState('Sheet1');
  const [fileName, setFileName] = useState("Untitled spreadsheet");
  
  // Formatting state
  const [fontSize, setFontSize] = useState(10);
  const [fontFamily, setFontFamily] = useState('Arial');
  
  // Color picker state
  const [showColorPicker, setShowColorPicker] = useState(false);
  const [colorPickerType, setColorPickerType] = useState('text');
  const [colorPickerPosition, setColorPickerPosition] = useState({ top: 0, left: 0 });
  const colorPickerRef = useRef(null);
  
  // Context menu state
  const [showContextMenu, setShowContextMenu] = useState(false);
  const [contextMenuPosition, setContextMenuPosition] = useState({ top: 0, left: 0 });
  
  // Cell comments state
  const [cellComments, setCellComments] = useState({});
  const [showCommentPopup, setShowCommentPopup] = useState(false);
  const [commentCell, setCommentCell] = useState(null);
  const [commentText, setCommentText] = useState("");
  
  // Conditional formatting state
  const [conditionalFormats, setConditionalFormats] = useState([]);
  const [showFormatRuleDialog, setShowFormatRuleDialog] = useState(false);
  const [formatRuleData, setFormatRuleData] = useState({
    range: null,
    type: 'greaterThan',
    value: '',
    backgroundColor: '#b7e1cd', // light green default
    textColor: '#000000',
  });
  
  // Data validation state
  const [dataValidations, setDataValidations] = useState({});
  const [showValidationDialog, setShowValidationDialog] = useState(false);
  const [validationData, setValidationData] = useState({
    range: null,
    type: 'list',   // list, number, date, text
    criteria: 'between', // between, notBetween, equal, notEqual, greaterThan, lessThan, etc.
    value1: '',
    value2: '',     // for between/notBetween
    listValues: '',  // comma-separated for list type
    errorMessage: 'The value you entered does not match the data validation criteria.',
    showDropdown: true, // for list type
  });

  // Rules manager state
  const [showRulesManager, setShowRulesManager] = useState(false);
  
  // Undo/Redo state
  const [history, setHistory] = useState([]);
  const [historyIndex, setHistoryIndex] = useState(-1);
  
  // Dragging state
  const isDragging = useRef(false);
  const dragStartCell = useRef(null);
  
  // Clipboard state
  const [clipboard, setClipboard] = useState(null);

  // Colors palette for color picker
  const colorPalette = [
    // First row - standard colors
    ['#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff'],
    // Second row - red palette
    ['#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff'],
    // Third row - pastel palette
    ['#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc'],
    // Fourth row - light palette
    ['#dd7e6b', '#ea9999', '#f9cb9c', '#ffe599', '#b6d7a8', '#a2c4c9', '#a4c2f4', '#9fc5e8', '#b4a7d6', '#d5a6bd'],
    // Fifth row - medium palette
    ['#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0'],
    // Sixth row - dark palette
    ['#a61c00', '#cc0000', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#3d85c6', '#674ea7', '#a64d79'],
    // Seventh row - very dark palette
    ['#85200c', '#990000', '#b45f06', '#bf9000', '#38761d', '#134f5c', '#1155cc', '#0b5394', '#351c75', '#741b47'],
  ];

  // ======== INITIALIZATION ========
  // Initialize cells with empty values
  useEffect(() => {
    const initialCells = {};
    for (let row = 1; row <= 100; row++) {
      for (let col = 0; col < 26; col++) {
        const colLetter = String.fromCharCode(65 + col);
        const cellId = `${colLetter}${row}`;
        initialCells[cellId] = { value: "", formula: "", formatted: "" };
      }
    }
    setCells(initialCells);
    
    // Add initial state to history
    addToHistory({
      cells: initialCells,
      cellStyles: {},
      mergedCells: {}
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Add event listener for menu clicks
  useEffect(() => {
    document.addEventListener('click', handleClickOutsideMenu);
    
    return () => {
      document.removeEventListener('click', handleClickOutsideMenu);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeMenu]);

  // ======== UNDO/REDO FUNCTIONALITY ========
  // Add state changes to history for undo/redo
  const addToHistory = (newState) => {
    // Remove any future states if we're not at the end of history
    const newHistory = history.slice(0, historyIndex + 1);
    // Add new state to history
    newHistory.push(newState);
    // Update history and historyIndex
    setHistory(newHistory);
    setHistoryIndex(newHistory.length - 1);
  };

  // Undo function
  const handleUndo = () => {
    if (historyIndex > 0) {
      const newIndex = historyIndex - 1;
      const prevState = history[newIndex];
      
      // Restore state
      setCells(prevState.cells);
      setCellStyles(prevState.cellStyles);
      
      setHistoryIndex(newIndex);
    }
  };

  // Redo function
  const handleRedo = () => {
    if (historyIndex < history.length - 1) {
      const newIndex = historyIndex + 1;
      const nextState = history[newIndex];
      
      // Restore state
      setCells(nextState.cells);
      setCellStyles(nextState.cellStyles);
      
      setHistoryIndex(newIndex);
    }
  };

  // ======== EVENT HANDLERS ========
  // Menu functions
  const handleMenuClick = (menuName) => {
    if (activeMenu === menuName) {
      setActiveMenu(null); // Close menu if clicking on already open menu
    } else {
      setActiveMenu(menuName); // Open clicked menu
    }
  };

  // Close menu when clicking outside
  const handleClickOutsideMenu = (e) => {
    if (activeMenu && !e.target.closest('.menu-item')) {
      setActiveMenu(null);
    }
  };

  // Handle cell mouse down
  const handleCellMouseDown = (cellId, e) => {
    // Right click - show context menu
    if (e.button === 2) {
      e.preventDefault();
      setContextMenuPosition({ top: e.clientY, left: e.clientX });
      setShowContextMenu(true);
      setActiveCell(cellId);
      setInputValue(cells[cellId]?.formula || "");
      return;
    }
    
    if (e.detail === 2) {
      // Double click - start editing directly in the cell
      setEditingCell(cellId);
      setActiveCell(cellId);
      
      // Get the current cell value
      const currentValue = cells[cellId]?.formula || "";
      setInputValue(currentValue);
      
      // Focus the input but don't manually set cursor position - let browser handle it
      setTimeout(() => {
        if (cellInputRef.current) {
          cellInputRef.current.focus();
        }
      }, 10);
    } else {
      // Single click - select cell or start multi-cell selection
      setActiveCell(cellId);
      setInputValue(cells[cellId]?.formula || "");
      
      // Start dragging
      isDragging.current = true;
      dragStartCell.current = cellId;
      
      // If Shift key is pressed, extend selection from previous active cell
      if (e.shiftKey && selectedRange) {
        // Use the previous active cell as the start of the range
        setSelectedRange({ 
          start: selectedRange.start, 
          end: cellId 
        });
      } else {
        // Normal single-cell selection
        setSelectedRange({ start: cellId, end: cellId });
      }
      
      // Update formula bar cell reference
      const formulaCellRef = document.getElementById('formula-cell-ref');
      if (formulaCellRef) formulaCellRef.value = cellId;
    }
  };

  // Handle cell mouse over during drag
  const handleCellMouseOver = (cellId) => {
    if (isDragging.current && dragStartCell.current) {
      // Update selected range
      setSelectedRange({ 
        start: dragStartCell.current, 
        end: cellId 
      });
    }
  };

  // Handle mouse up to end dragging
  const handleMouseUp = (e) => {
    // Check if dragging was active
    if (isDragging.current) {
      // Stop dragging
      isDragging.current = false;
      dragStartCell.current = null;
    }
  };

  // Ensure mouse up event is added in useEffect
  useEffect(() => {
    document.addEventListener('mouseup', handleMouseUp);
    document.addEventListener('contextmenu', handleContextMenu);
    
    // Close context menu when clicking outside
    const handleClickOutside = (e) => {
      if (showContextMenu) {
        setShowContextMenu(false);
      }
    };
    document.addEventListener('click', handleClickOutside);
    
    return () => {
      document.removeEventListener('mouseup', handleMouseUp);
      document.removeEventListener('contextmenu', handleContextMenu);
      document.removeEventListener('click', handleClickOutside);
    };
  }, [showContextMenu]); // Only re-run when showContextMenu changes

  // Prevent default context menu
  const handleContextMenu = (e) => {
    e.preventDefault();
  };

  // Handle formula input change
  const handleFormulaChange = (e) => {
    setInputValue(e.target.value);
  };

  // Handle cell input change (for direct editing)
  const handleCellInputChange = (e) => {
    const newValue = e.target.value;
    setInputValue(newValue);
    
    // Apply changes immediately as user types
    if (editingCell) {
      const newCells = { ...cells };
      
      // Update cell with current input value
      newCells[editingCell] = { 
        value: newValue, 
        formula: newValue, 
        formatted: newValue.startsWith('=') ? newValue : formatValue(newValue, cellStyles[editingCell]) 
      };
      
      setCells(newCells);
      
      // If this is a formula cell, we'll fully evaluate it on finish
      if (!newValue.startsWith('=')) {
        // For non-formulas, update dependent cells immediately
        updateDependentCells(editingCell, newCells, dependencies);
      }
    }
  };

  // Handle finishing cell editing
  const finishCellEditing = () => {
    if (editingCell) {
      // Save current state before applying formula
      addToHistory({
        cells: {...cells},
        cellStyles: {...cellStyles}
      });
      
      applyFormula();
      setEditingCell(null);
    }
  };

  // Handle cell input key down
  const handleCellInputKeyDown = (e) => {
    // Don't interfere with normal text editing keys
    if (e.key === 'Backspace' || e.key === 'Delete') {
      // Let the default behavior handle character deletion
      return;
    }

    if (e.key === 'Enter') {
      finishCellEditing();
      
      // Move to the cell below
      const currentCol = editingCell.match(/[A-Z]+/)[0];
      const currentRow = parseInt(editingCell.match(/[0-9]+/)[0]);
      const currentRowIdx = visibleRows.indexOf(currentRow);
      
      if (currentRowIdx < visibleRows.length - 1) {
        const newRow = visibleRows[currentRowIdx + 1];
        setActiveCell(`${currentCol}${newRow}`);
      }
    } else if (e.key === 'Escape') {
      setEditingCell(null);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      finishCellEditing();
      
      // Move to the next cell
      const currentCol = editingCell.match(/[A-Z]+/)[0];
      const currentRow = parseInt(editingCell.match(/[0-9]+/)[0]);
      const currentColIdx = visibleColumns.indexOf(currentCol);
      
      if (currentColIdx < visibleColumns.length - 1) {
        const newCol = visibleColumns[currentColIdx + 1];
        setActiveCell(`${newCol}${currentRow}`);
      }
    }
  };

  // Handle keyboard shortcuts
  const handleKeyboardShortcut = (e) => {
    // Early return if in edit mode with specific keys
    if (editingCell && (
      e.key === 'Backspace' || 
      e.key === 'Delete' || 
      (e.ctrlKey && (e.key === 'v' || e.key === 'c' || e.key === 'x'))
    )) {
      return; // Let the browser handle these operations in edit mode
    }
    
    // Prevent default and stop propagation for all menu shortcuts
    const handleMenuShortcut = (action) => {
      e.preventDefault();
      e.stopPropagation(); // This is key to preventing event bubbling
      action();
    };

    // Bold: Ctrl+B
    if (e.ctrlKey && e.key === 'b') {
      handleMenuShortcut(() => applyFormatting('bold'));
      return;
    }
    
    // Italic: Ctrl+I
    if (e.ctrlKey && e.key === 'i') {
      handleMenuShortcut(() => applyFormatting('italic'));
      return;
    }
    
    // Underline: Ctrl+U
    if (e.ctrlKey && e.key === 'u') {
      handleMenuShortcut(() => applyFormatting('underline'));
      return;
    }
    
    // Undo: Ctrl+Z
    if (e.ctrlKey && e.key === 'z' && !e.shiftKey) {
      handleMenuShortcut(handleUndo);
      return;
    }
    
    // Redo: Ctrl+Y or Ctrl+Shift+Z
    if ((e.ctrlKey && e.key === 'y') || (e.ctrlKey && e.shiftKey && e.key === 'z')) {
      handleMenuShortcut(handleRedo);
      return;
    }
    
    // Copy: Ctrl+C
    if (e.ctrlKey && e.key === 'c') {
      handleMenuShortcut(handleCopy);
      return;
    }
    
    // Cut: Ctrl+X
    if (e.ctrlKey && e.key === 'x') {
      handleMenuShortcut(handleCut);
      return;
    }
    
    // Paste: Ctrl+V
    if (e.ctrlKey && e.key === 'v') {
      handleMenuShortcut(handlePaste);
      return;
    }
    
    // Print: Ctrl+P
    if (e.ctrlKey && e.key === 'p') {
      handleMenuShortcut(() => window.print());
      return;
    }
  };
  
  // Modified handleGridKeyDown function to properly handle typing directly into cells
  const handleGridKeyDown = (e) => {
    // Skip handling keyboard shortcuts if in edit mode - this is crucial for deletion to work properly
    if (editingCell) {
      return; // Let default behavior handle all keys when editing
    }
    
    // Handle keyboard shortcuts for non-editing mode
    handleKeyboardShortcut(e);
    
    // Check if we can start typing
    const isPrintableChar = 
      e.key.length === 1 && 
      !e.ctrlKey && 
      !e.altKey && 
      !e.metaKey;
    
    if (isPrintableChar && activeCell && !editingCell) {
      // Start editing mode with the pressed key
      setEditingCell(activeCell);
      setInputValue(e.key);
      e.preventDefault();
      
      // Defer focus to ensure React has updated state
      setTimeout(() => {
        if (cellInputRef.current) {
          cellInputRef.current.focus();
          // Move cursor to the end
          cellInputRef.current.selectionStart = cellInputRef.current.value.length;
          cellInputRef.current.selectionEnd = cellInputRef.current.value.length;
        }
      }, 0);
      
      return;
    }
    
    // If no active cell, do nothing
    if (!activeCell) return;
    
    const currentCol = activeCell.match(/[A-Z]+/)[0];
    const currentRow = parseInt(activeCell.match(/[0-9]+/)[0]);
    const currentColIdx = visibleColumns.indexOf(currentCol);
    const currentRowIdx = visibleRows.indexOf(currentRow);
    
    switch (e.key) {
      case 'ArrowUp':
        if (currentRowIdx > 0) {
          const newRow = visibleRows[currentRowIdx - 1];
          setActiveCell(`${currentCol}${newRow}`);
        }
        break;
        
      case 'ArrowDown':
        if (currentRowIdx < visibleRows.length - 1) {
          const newRow = visibleRows[currentRowIdx + 1];
          setActiveCell(`${currentCol}${newRow}`);
        }
        break;
        
      case 'ArrowLeft':
        if (currentColIdx > 0) {
          const newCol = visibleColumns[currentColIdx - 1];
          setActiveCell(`${newCol}${currentRow}`);
        }
        break;
        
      case 'ArrowRight':
        if (currentColIdx < visibleColumns.length - 1) {
          const newCol = visibleColumns[currentColIdx + 1];
          setActiveCell(`${newCol}${currentRow}`);
        }
        break;
        
      case 'Enter':
        if (currentRowIdx < visibleRows.length - 1) {
          const newRow = visibleRows[currentRowIdx + 1];
          setActiveCell(`${currentCol}${newRow}`);
        }
        break;
        
      case 'Tab':
        e.preventDefault();
        if (currentColIdx < visibleColumns.length - 1) {
          const newCol = visibleColumns[currentColIdx + 1];
          setActiveCell(`${newCol}${currentRow}`);
        } else if (currentRowIdx < visibleRows.length - 1) {
          const newRow = visibleRows[currentRowIdx + 1];
          setActiveCell(`${visibleColumns[0]}${newRow}`);
        }
        break;
        
      case 'Delete':
      case 'Backspace':
        // Clear selected cells
        clearCells();
        break;
        
      case 'F2':
        // Start editing directly, similar to Excel/Google Sheets
        setEditingCell(activeCell);
        setInputValue(cells[activeCell]?.formula || "");
        
        // Focus the edit input after rendering
        setTimeout(() => {
          if (cellInputRef.current) {
            cellInputRef.current.focus();
            cellInputRef.current.select();
          }
        }, 0);
        break;
        
      default:
        break;
    }
  };

  // Add this at the end of the EVENT HANDLERS section
// Sheet switching function
const switchSheet = (sheetName) => {
  // Save current sheet state
  const currentSheetData = { ...sheetData };
  currentSheetData[activeSheet] = {
    cells: { ...cells },
    cellStyles: { ...cellStyles },
    dependencies: { ...dependencies }
  };
  
  // Update sheet data state
  setSheetData(currentSheetData);
  
  // Set active sheet
  setActiveSheet(sheetName);
  
  // Load the selected sheet's data
  setCells(currentSheetData[sheetName]?.cells || {});
  setCellStyles(currentSheetData[sheetName]?.cellStyles || {});
  setDependencies(currentSheetData[sheetName]?.dependencies || {});
};

// Chart-related functions
// Open chart dialog
const openChartDialog = () => {
  setShowChartDialog(true);
};

// Close chart dialog and optionally add the new chart
const closeChartDialog = (chartData) => {
  setShowChartDialog(false);
  
  if (chartData) {
    // Add the new chart to state
    const newChart = {
      id: `chart-${Date.now()}`,
      ...chartData,
      position: {
        top: 100,
        left: 100,
        width: 500,
        height: 300
      }
    };
    
    setCharts([...charts, newChart]);
    setActiveChart(newChart.id);
    
    // Save current state to history
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles},
      charts: [...charts, newChart]
    });
  }
};

// Delete a chart
const deleteChart = (chartId) => {
  // Save current state to history before deleting
  addToHistory({
    cells: {...cells},
    cellStyles: {...cellStyles},
    charts: [...charts]
  });
  
  const updatedCharts = charts.filter(chart => chart.id !== chartId);
  setCharts(updatedCharts);
  
  if (activeChart === chartId) {
    setActiveChart(null);
  }
};

// Update chart position or size
const updateChartPosition = (chartId, position) => {
  const updatedCharts = charts.map(chart => {
    if (chart.id === chartId) {
      return { ...chart, position: { ...chart.position, ...position } };
    }
    return chart;
  });
  
  setCharts(updatedCharts);
};

// Handle chart dragging
const startChartDrag = (chartId, e) => {
  const chartEl = document.getElementById(chartId);
  const initialX = e.clientX;
  const initialY = e.clientY;
  const chartRect = chartEl.getBoundingClientRect();
  
  const handleMouseMove = (moveEvent) => {
    const deltaX = moveEvent.clientX - initialX;
    const deltaY = moveEvent.clientY - initialY;
    
    updateChartPosition(chartId, {
      top: chartRect.top + deltaY,
      left: chartRect.left + deltaX
    });
  };
  
  const handleMouseUp = () => {
    document.removeEventListener('mousemove', handleMouseMove);
    document.removeEventListener('mouseup', handleMouseUp);
    
    // Save position in history after drag ends
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles},
      charts: [...charts]
    });
  };
  
  document.addEventListener('mousemove', handleMouseMove);
  document.addEventListener('mouseup', handleMouseUp);
  
  // Set as active chart
  setActiveChart(chartId);
  e.stopPropagation();
};

// Render chart based on type
const renderChart = (chart) => {
  const { type, data, xAxisKey, yAxisKeys, title, stacked } = chart;
  
  switch (type) {
    case 'line':
      return (
        <LineChartComponent 
          data={data} 
          xKey={xAxisKey} 
          yKeys={yAxisKeys} 
          title={title} 
        />
      );
    case 'bar':
      return (
        <BarChartComponent 
          data={data} 
          xKey={xAxisKey} 
          yKeys={yAxisKeys} 
          title={title} 
          stacked={stacked}
        />
      );
    case 'pie':
      return (
        <PieChartComponent 
          data={data} 
          nameKey="name" 
          valueKey="value" 
          title={title} 
        />
      );
    case 'area':
      return (
        <AreaChartComponent 
          data={data} 
          xKey={xAxisKey} 
          yKeys={yAxisKeys} 
          title={title} 
          stacked={stacked}
        />
      );
    default:
      return null;
  }
};

  // ======== FORMULA HANDLING ========
  // Apply formula when pressing Enter or leaving the input
  const applyFormula = () => {
    if (!activeCell) return;
    
    const formula = inputValue;
    const newCells = { ...cells };
    const newDependencies = { ...dependencies };
    
    // Validate cell value if not a formula
    if (!formula.startsWith('=')) {
      const validationResult = validateCellValue(activeCell, formula);
      if (!validationResult.valid) {
        const proceed = confirm(`${validationResult.message}\n\nDo you want to continue anyway?`);
        if (!proceed) return; // Don't apply the invalid value
      }
    }
    
    // Clear old dependencies for this cell
    Object.keys(newDependencies).forEach(dep => {
      if (newDependencies[dep]?.includes(activeCell)) {
        newDependencies[dep] = newDependencies[dep].filter(cell => cell !== activeCell);
      }
    });
    
    if (formula.startsWith('=')) {
      // Extract cell references and update dependencies
      const cellRefs = extractCellReferences(formula);
      cellRefs.forEach(ref => {
        if (!newDependencies[ref]) {
          newDependencies[ref] = [];
        }
        if (!newDependencies[ref].includes(activeCell)) {
          newDependencies[ref].push(activeCell);
        }
      });
      
      // Evaluate formula
      try {
        let value = evaluateFormula(formula, newCells);
        newCells[activeCell] = { 
          value: value, 
          formula: formula, 
          formatted: formatValue(value, cellStyles[activeCell]) 
        };
      } catch (error) {
        newCells[activeCell] = { 
          value: "#ERROR", 
          formula: formula, 
          formatted: "#ERROR" 
        };
      }
    } else {
      // Regular value
      newCells[activeCell] = { 
        value: formula, 
        formula: formula, 
        formatted: formatValue(formula, cellStyles[activeCell])
      };
    }
    
    setCells(newCells);
    setDependencies(newDependencies);
    
    // Update dependent cells
    updateDependentCells(activeCell, newCells, newDependencies);
  };

  // Extract cell references from a formula
  const extractCellReferences = (formula) => {
    // Match both regular references (A1) and absolute references ($A$1)
    const cellPattern = /\$?[A-Z]+\$?[0-9]+/g;
    return formula.match(cellPattern) || [];
  };

  // Update cells that depend on the changed cell
  const updateDependentCells = (cellId, cellsData, deps) => {
    if (!deps[cellId]) return;
    
    const dependentCells = deps[cellId];
    dependentCells.forEach(depCell => {
      const formula = cellsData[depCell]?.formula;
      if (formula && formula.startsWith('=')) {
        try {
          const value = evaluateFormula(formula, cellsData);
          cellsData[depCell].value = value;
          cellsData[depCell].formatted = formatValue(value, cellStyles[depCell]);
        } catch (error) {
          cellsData[depCell].value = "#ERROR";
          cellsData[depCell].formatted = "#ERROR";
        }
        
        // Recursively update cells that depend on this cell
        updateDependentCells(depCell, cellsData, deps);
      }
    });
    
    setCells({...cellsData});
  };

  // Split formula parameters respecting nested parentheses and commas in strings
  const splitParameters = (content) => {
    const result = [];
    let currentParam = '';
    let parenCount = 0;
    let inQuote = false;
    
    for (let i = 0; i < content.length; i++) {
      const char = content[i];
      
      if (char === ',' && parenCount === 0 && !inQuote) {
        result.push(currentParam.trim());
        currentParam = '';
        continue;
      }
      
      if (char === '(') parenCount++;
      if (char === ')') parenCount--;
      if (char === '"') inQuote = !inQuote;
      
      currentParam += char;
    }
    
    if (currentParam) {
      result.push(currentParam.trim());
    }
    
    return result;
  };

  // Evaluate formula
  const evaluateFormula = (formula, cellsData) => {
    // Handle mathematical functions
    if (formula.startsWith('=SUM(') && formula.endsWith(')')) {
      const range = formula.substring(5, formula.length - 1);
      return evaluateMathFunction(range, cellsData, 'SUM');
    }
    
    if (formula.startsWith('=AVERAGE(') && formula.endsWith(')')) {
      const range = formula.substring(9, formula.length - 1);
      return evaluateMathFunction(range, cellsData, 'AVERAGE');
    }
    
    if (formula.startsWith('=MAX(') && formula.endsWith(')')) {
      const range = formula.substring(5, formula.length - 1);
      return evaluateMathFunction(range, cellsData, 'MAX');
    }
    
    if (formula.startsWith('=MIN(') && formula.endsWith(')')) {
      const range = formula.substring(5, formula.length - 1);
      return evaluateMathFunction(range, cellsData, 'MIN');
    }
    
    if (formula.startsWith('=COUNT(') && formula.endsWith(')')) {
      const range = formula.substring(7, formula.length - 1);
      return evaluateMathFunction(range, cellsData, 'COUNT');
    }
    
    if (formula.startsWith('=COUNTIF(') && formula.endsWith(')')) {
      const content = formula.substring(9, formula.length - 1);
      const [range, criteria] = content.split(',').map(part => part.trim());
      return evaluateCountIfFunction(range, criteria, cellsData);
    }
    
    if (formula.startsWith('=IF(') && formula.endsWith(')')) {
      const content = formula.substring(4, formula.length - 1);
      const parts = splitParameters(content);
      if (parts.length === 3) {
        const [condition, trueValue, falseValue] = parts;
        return evaluateIfFunction(condition, trueValue, falseValue, cellsData);
      }
      return "#ERROR: IF requires 3 arguments";
    }
    
    // Data quality functions
    if (formula.startsWith('=TRIM(') && formula.endsWith(')')) {
      const param = formula.substring(6, formula.length - 1);
      return evaluateDataQualityFunction(param, cellsData, 'TRIM');
    }
    
    if (formula.startsWith('=UPPER(') && formula.endsWith(')')) {
      const param = formula.substring(7, formula.length - 1);
      return evaluateDataQualityFunction(param, cellsData, 'UPPER');
    }
    
    if (formula.startsWith('=LOWER(') && formula.endsWith(')')) {
      const param = formula.substring(7, formula.length - 1);
      return evaluateDataQualityFunction(param, cellsData, 'LOWER');
    }
    
    if (formula.startsWith('=CONCATENATE(') && formula.endsWith(')')) {
      const content = formula.substring(12, formula.length - 1);
      const parts = splitParameters(content);
      return evaluateConcatenateFunction(parts, cellsData);
    }
    
    // Cell reference (handle both relative and absolute references)
    if (formula.match(/^=\$?[A-Z]+\$?[0-9]+$/)) {
      const cellRef = formula.substring(1).replace(/\$/g, '');
      return cellsData[cellRef]?.value || "";
    }
    
    // Basic arithmetic with cell references
    if (formula.startsWith('=')) {
      let expression = formula.substring(1);
      
      // Replace cell references with their values
      const cellRefs = extractCellReferences(formula);
      cellRefs.forEach(ref => {
        // Remove $ signs for absolute references
        const pureRef = ref.replace(/\$/g, '');
        const cellValue = cellsData[pureRef]?.value || "0";
        // Escape to handle string values properly
        const valueToInsert = isNaN(parseFloat(cellValue)) ? 
          `"${cellValue}"` : cellValue;
        
        // Create a regex that handles $ signs as special characters
        const escapedRef = ref.replace(/\$/g, '\\$');
        const refRegex = new RegExp(escapedRef, 'g');
        expression = expression.replace(refRegex, valueToInsert);
      });
      
      // CAUTION: Using eval for formula evaluation - for educational purposes only
      try {
        // eslint-disable-next-line no-eval
        return eval(expression).toString();
      } catch (error) {
        console.error("Formula evaluation error:", error);
        return "#ERROR";
      }
    }
    
    return formula.substring(1);
  };

  // Evaluate CONCATENATE function
  const evaluateConcatenateFunction = (params, cellsData) => {
    try {
      return params.map(param => {
        param = param.trim();
        
        // Cell reference
        if (param.match(/^[A-Z]+[0-9]+$/)) {
          return cellsData[param]?.value || "";
        }
        
        // String literal
        if (param.startsWith('"') && param.endsWith('"')) {
          return param.substring(1, param.length - 1);
        }
        
        return param;
      }).join('');
    } catch (error) {
      return "#ERROR";
    }
  };

  // Evaluate IF function
  const evaluateIfFunction = (condition, trueValue, falseValue, cellsData) => {
    try {
      // Process condition
      let processedCondition = condition;
      const cellRefs = extractCellReferences(condition);
      
      cellRefs.forEach(ref => {
        const pureRef = ref.replace(/\$/g, '');
        const cellValue = cellsData[pureRef]?.value || "";
        const valueToInsert = isNaN(parseFloat(cellValue)) ? 
          `"${cellValue}"` : cellValue;
        
        const escapedRef = ref.replace(/\$/g, '\\$');
        const refRegex = new RegExp(escapedRef, 'g');
        processedCondition = processedCondition.replace(refRegex, valueToInsert);
      });
      
      // Evaluate condition
      // eslint-disable-next-line no-eval
      const conditionResult = eval(processedCondition);
      
      // Process true/false values for cell references
      const processValue = (value) => {
        if (value.startsWith('=')) {
          return evaluateFormula(value, cellsData);
        }
        
        const cellRef = extractCellReferences(value)[0];
        if (cellRef) {
          const pureRef = cellRef.replace(/\$/g, '');
          return cellsData[pureRef]?.value || "";
        }
        
        // Remove quotes if string literal
        if (value.startsWith('"') && value.endsWith('"')) {
          return value.substring(1, value.length - 1);
        }
        
        return value;
      };
      
      return conditionResult ? processValue(trueValue) : processValue(falseValue);
    } catch (error) {
      console.error("IF function error:", error);
      return "#ERROR";
    }
  };

  // Evaluate COUNTIF function
  const evaluateCountIfFunction = (range, criteria, cellsData) => {
    try {
      const cells = range.includes(':') ? 
        getCellsInRange(...range.split(':').map(r => r.replace(/\$/g, ''))) : 
        [range.replace(/\$/g, '')];
      
      // Process criteria
      let comparison = '===';
      let criteriaValue = criteria;
      
      if (criteria.startsWith('"') && criteria.endsWith('"')) {
        criteriaValue = criteria.substring(1, criteria.length - 1);
      } else if (criteria.startsWith('>=') || criteria.startsWith('<=') || 
                 criteria.startsWith('<>')) {
        comparison = criteria.substring(0, 2);
        criteriaValue = criteria.substring(2);
      } else if (criteria.startsWith('>') || criteria.startsWith('<') || 
                 criteria.startsWith('=')) {
        comparison = criteria.substring(0, 1);
        criteriaValue = criteria.substring(1);
      }
      
      if (comparison === '=') comparison = '===';
      if (comparison === '<>') comparison = '!==';
      
      // Convert to number if possible
      if (!isNaN(parseFloat(criteriaValue))) {
        criteriaValue = parseFloat(criteriaValue);
      }
      
      // Count matching cells
      let count = 0;
      cells.forEach(cellId => {
        const cellValue = cellsData[cellId]?.value || "";
        const numericValue = !isNaN(parseFloat(cellValue)) ? parseFloat(cellValue) : cellValue;
        
        // Evaluate the comparison
        let matches = false;
        switch (comparison) {
          case '===': matches = numericValue === criteriaValue; break;
          case '>': matches = numericValue > criteriaValue; break;
          case '<': matches = numericValue < criteriaValue; break;
          case '>=': matches = numericValue >= criteriaValue; break;
          case '<=': matches = numericValue <= criteriaValue; break;
          case '!==': matches = numericValue !== criteriaValue; break;
          default: matches = false;
        }
        
        if (matches) count++;
      });
      
      return count.toString();
    } catch (error) {
      return "#ERROR";
    }
  };

  // Evaluate mathematical functions
  const evaluateMathFunction = (range, cellsData, func) => {
    let cellsList = [];
    
    if (range.includes(':')) {
      // Range notation (e.g., A1:B3)
      const [start, end] = range.split(':');
      cellsList = getCellsInRange(start.replace(/\$/g, ''), end.replace(/\$/g, ''));
    } else if (range.includes(',')) {
      // List notation (e.g., A1,B1,C1)
      cellsList = range.split(',').map(cell => cell.trim().replace(/\$/g, ''));
    } else {
      // Single cell
      cellsList = [range.replace(/\$/g, '')];
    }
    
    // Extract numeric values
    const values = cellsList.map(cellId => {
      const value = parseFloat(cellsData[cellId]?.value || 0);
      return isNaN(value) ? 0 : value;
    }).filter(val => !isNaN(val));
    
    // Calculate result based on function
    switch (func) {
      case 'SUM':
        return values.reduce((sum, val) => sum + val, 0).toString();
        
      case 'AVERAGE':
        return values.length > 0 ? 
          (values.reduce((sum, val) => sum + val, 0) / values.length).toString() : 
          "0";
          
      case 'MAX':
        return values.length > 0 ? 
          Math.max(...values).toString() : 
          "0";
          
      case 'MIN':
        return values.length > 0 ? 
          Math.min(...values).toString() : 
          "0";
          
      case 'COUNT':
        return values.length.toString();
        
      default:
        return "#ERROR";
    }
  };

  // Evaluate data quality functions
  const evaluateDataQualityFunction = (param, cellsData, func) => {
    let value = "";
    
    // Check if param is a cell reference
    if (param.match(/^\$?[A-Z]+\$?[0-9]+$/)) {
      const cellRef = param.replace(/\$/g, '');
      value = cellsData[cellRef]?.value || "";
    } else if (param.startsWith('"') && param.endsWith('"')) {
      // String literal
      value = param.substring(1, param.length - 1);
    } else {
      value = param;
    }
    
    // Apply function
    switch (func) {
      case 'TRIM':
        return value.trim();
        
      case 'UPPER':
        return value.toUpperCase();
        
      case 'LOWER':
        return value.toLowerCase();
        
      default:
        return "#ERROR";
    }
  };

  // Format value based on cell styles
  const formatValue = (value, style) => {
    if (!style) return value;
    
    // Implement number formatting, date formatting, etc.
    if (style.numberFormat === 'percentage' && !isNaN(parseFloat(value))) {
      return (parseFloat(value) * 100) + '%';
    }
    
    if (style.numberFormat === 'currency' && !isNaN(parseFloat(value))) {
      return '$' + parseFloat(value).toFixed(2);
    }
    
    if (style.numberFormat === 'date' && !isNaN(Date.parse(value))) {
      return new Date(value).toLocaleDateString();
    }
    
    return value;
  };

  // Validate cell value against data validation rules
  const validateCellValue = (cellId, value) => {
    // Default valid result
    return { valid: true, message: "" };
  };

  // ======== CELL RANGE UTILITIES ========
  // Get all cells in a range (e.g., A1:C3)
  const getCellsInRange = (start, end) => {
    // Extract column and row components
    const startCol = start.match(/[A-Z]+/)[0];
    const startRow = parseInt(start.match(/[0-9]+/)[0]);
    const endCol = end.match(/[A-Z]+/)[0];
    const endRow = parseInt(end.match(/[0-9]+/)[0]);
    
    // Convert column letters to indices
    const startColIdx = columnToIndex(startCol);
    const endColIdx = columnToIndex(endCol);
    
    // Create list of cells in range
    const cellsList = [];
    
    for (let row = Math.min(startRow, endRow); row <= Math.max(startRow, endRow); row++) {
      for (let colIdx = Math.min(startColIdx, endColIdx); colIdx <= Math.max(startColIdx, endColIdx); colIdx++) {
        const colLetter = indexToColumn(colIdx);
        cellsList.push(`${colLetter}${row}`);
      }
    }
    
    return cellsList;
  };

  // Convert column letter to index (A->0, B->1, ...)
  const columnToIndex = (col) => {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
      result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result - 1;
  };

  // Convert index to column letter (0->A, 1->B, ...)
  const indexToColumn = (idx) => {
    let result = '';
    idx = idx + 1; // 1-based
    
    while (idx > 0) {
      const remainder = (idx - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      idx = Math.floor((idx - remainder) / 26);
    }
    
    return result;
  };

  // Check if a cell is in the selected range
  const isCellInRange = (cellId) => {
    if (!selectedRange || !cellId) return false;
    
    return isCellInCustomRange(cellId, selectedRange.start, selectedRange.end);
  };

  // Check if a cell is in a custom range
  const isCellInCustomRange = (cellId, start, end) => {
    if (!cellId || !start || !end) return false;
    
    // Extract components
    const cellCol = cellId.match(/[A-Z]+/)[0];
    const cellRow = parseInt(cellId.match(/[0-9]+/)[0]);
    
    const startCol = start.match(/[A-Z]+/)[0];
    const startRow = parseInt(start.match(/[0-9]+/)[0]);
    const endCol = end.match(/[A-Z]+/)[0];
    const endRow = parseInt(end.match(/[0-9]+/)[0]);
    
    // Convert to indices for comparison
    const cellColIdx = columnToIndex(cellCol);
    const startColIdx = columnToIndex(startCol);
    const endColIdx = columnToIndex(endCol);
    
    // Check if cell is within range bounds
    return (
      cellColIdx >= Math.min(startColIdx, endColIdx) &&
      cellColIdx <= Math.max(startColIdx, endColIdx) &&
      cellRow >= Math.min(startRow, endRow) &&
      cellRow <= Math.max(startRow, endRow)
    );
  };

  // ======== CELL FORMATTING ========
  // Toggle color picker
  const toggleColorPicker = (type, element) => {
    if (showColorPicker && colorPickerType === type) {
      setShowColorPicker(false);
    } else {
      setColorPickerType(type);
      setShowColorPicker(true);
      
      // Calculate position based on the button element
      const rect = element.getBoundingClientRect();
      setColorPickerPosition({
        top: rect.bottom + 5,
        left: rect.left
      });
    }
  };
  
  // Apply color
  const applyColor = (color) => {
    // Save current state before applying color
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    if (colorPickerType === 'text') {
      applyFormatting('textColor', color);
    } else {
      applyFormatting('backgroundColor', color);
    }
    setShowColorPicker(false);
  };
  
  // Close color picker when clicking outside
  useEffect(() => {
    const handleClickOutside = (e) => {
      // Use capture phase to ensure this runs before other event handlers
      if (colorPickerRef.current && !colorPickerRef.current.contains(e.target)) {
        setShowColorPicker(false);
      }
    };
    
    // Add event listener with capture option to prevent multiple triggers
    document.addEventListener('mousedown', handleClickOutside, { capture: true });
    
    // Cleanup listener on component unmount
    return () => {
      document.removeEventListener('mousedown', handleClickOutside, { capture: true });
    };
  }, []); // Empty dependency array ensures this runs only once on mount

  // Apply formatting
  const applyFormatting = (styleType, value) => {
    if (!activeCell && !selectedRange) return;
    
    // Save current state before applying formatting
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    const newCellStyles = { ...cellStyles };
    const cellsToFormat = [];
    
    // Determine which cells to format
    if (selectedRange) {
      cellsToFormat.push(...getCellsInRange(selectedRange.start, selectedRange.end));
    } else if (activeCell) {
      cellsToFormat.push(activeCell);
    }
    
    // Apply style to each cell
    cellsToFormat.forEach(cellId => {
      if (!newCellStyles[cellId]) {
        newCellStyles[cellId] = {};
      }
      
      switch (styleType) {
        case 'bold':
          newCellStyles[cellId].fontWeight = 
            newCellStyles[cellId].fontWeight === 'bold' ? 'normal' : 'bold';
          break;
          
        case 'italic':
          newCellStyles[cellId].fontStyle = 
            newCellStyles[cellId].fontStyle === 'italic' ? 'normal' : 'italic';
          break;
          
        case 'underline':
          newCellStyles[cellId].textDecoration = 
            newCellStyles[cellId].textDecoration === 'underline' ? 'none' : 'underline';
          break;
          
        case 'fontSize':
          newCellStyles[cellId].fontSize = value + 'px';
          break;
          
        case 'fontFamily':
          newCellStyles[cellId].fontFamily = value;
          break;
          
        case 'textAlign':
          newCellStyles[cellId].textAlign = value;
          break;
          
        case 'backgroundColor':
          newCellStyles[cellId].backgroundColor = value;
          break;
          
        case 'textColor':
          newCellStyles[cellId].color = value;
          break;
          
        case 'numberFormat':
          newCellStyles[cellId].numberFormat = value;
          
          // Also update the formatted value
          if (cells[cellId]) {
            const newCells = { ...cells };
            newCells[cellId].formatted = formatValue(newCells[cellId].value, {
              ...newCellStyles[cellId],
              numberFormat: value
            });
            setCells(newCells);
          }
          break;
          
        default:
          break;
      }
    });
    
    setCellStyles(newCellStyles);
  };

  // Handle copy
  const handleCopy = () => {
    if (!selectedRange && !activeCell) return;
    
    const copyRange = selectedRange || { start: activeCell, end: activeCell };
    const copyCells = {};
    
    getCellsInRange(copyRange.start, copyRange.end).forEach(cellId => {
      copyCells[cellId] = { ...cells[cellId] };
    });
    
    setClipboard({
      startCol: copyRange.start.match(/[A-Z]+/)[0],
      startRow: parseInt(copyRange.start.match(/[0-9]+/)[0]),
      endCol: copyRange.end.match(/[A-Z]+/)[0],
      endRow: parseInt(copyRange.end.match(/[0-9]+/)[0]),
      cells: copyCells
    });
  };

  // Handle paste
  const handlePaste = () => {
    if (!clipboard || !activeCell) return;
    
    // Save current state before pasting
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    const targetCol = activeCell.match(/[A-Z]+/)[0];
    const targetRow = parseInt(activeCell.match(/[0-9]+/)[0]);
    
    const colOffset = columnToIndex(targetCol) - columnToIndex(clipboard.startCol);
    const rowOffset = targetRow - clipboard.startRow;
    
    const newCells = { ...cells };
    
    Object.entries(clipboard.cells).forEach(([cellId, cellData]) => {
      const col = cellId.match(/[A-Z]+/)[0];
      const row = parseInt(cellId.match(/[0-9]+/)[0]);
      
      const newColIdx = columnToIndex(col) + colOffset;
      const newRowIdx = row + rowOffset;
      
      if (newColIdx >= 0 && newRowIdx > 0) {
        const newCol = indexToColumn(newColIdx);
        const newCellId = `${newCol}${newRowIdx}`;
        
        // Copy cell data
        newCells[newCellId] = { ...cellData };
        
        // Update formulas to reflect new position
        if (cellData.formula && cellData.formula.startsWith('=')) {
          const updatedFormula = updateFormulaCellReferences(
            cellData.formula, 
            colOffset, 
            rowOffset
          );
          
          newCells[newCellId].formula = updatedFormula;
          
          // Re-evaluate the formula with the new references
          try {
            const value = evaluateFormula(updatedFormula, newCells);
            newCells[newCellId].value = value;
            newCells[newCellId].formatted = formatValue(value, cellStyles[newCellId]);
          } catch (error) {
            newCells[newCellId].value = "#ERROR";
            newCells[newCellId].formatted = "#ERROR";
          }
        }
      }
    });
    
    setCells(newCells);
  };

  // Update formula cell references when copying/pasting
  const updateFormulaCellReferences = (formula, colOffset, rowOffset) => {
    if (!formula.startsWith('=')) return formula;
    
    // Get all cell references
    const cellRefs = extractCellReferences(formula);
    
    let updatedFormula = formula;
    
    cellRefs.forEach(ref => {
      // Skip absolute references (those with $ sign)
      if (ref.includes('$')) {
        return;
      }
      
      const col = ref.match(/[A-Z]+/)[0];
      const row = parseInt(ref.match(/[0-9]+/)[0]);
      
      const newColIdx = columnToIndex(col) + colOffset;
      const newRowIdx = row + rowOffset;
      
      if (newColIdx >= 0 && newRowIdx > 0) {
        const newCol = indexToColumn(newColIdx);
        const newRef = `${newCol}${newRowIdx}`;
        
        // Replace the old reference with the new one
        updatedFormula = updatedFormula.replace(new RegExp(ref, 'g'), newRef);
      }
    });
    
    return updatedFormula;
  };

  // Handle cut
  const handleCut = () => {
    if (!selectedRange && !activeCell) return;
    
    // Save current state before cutting
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    // Copy cells first
    handleCopy();
    
    // Then clear the selected cells
    const cellsToClear = selectedRange 
      ? getCellsInRange(selectedRange.start, selectedRange.end) 
      : [activeCell];
    
    const newCells = { ...cells };
    cellsToClear.forEach(cellId => {
      newCells[cellId] = { value: "", formula: "", formatted: "" };
    });
    
    setCells(newCells);
  };

  // Clear selected cells
  const clearCells = () => {
    if (!selectedRange && !activeCell) return;
    
    // Save current state before clearing cells
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    const cellsToClear = selectedRange 
      ? getCellsInRange(selectedRange.start, selectedRange.end) 
      : [activeCell];
    
    const newCells = { ...cells };
    cellsToClear.forEach(cellId => {
      newCells[cellId] = { value: "", formula: "", formatted: "" };
    });
    
    setCells(newCells);
  };

  // Handle context menu items
  const handleContextMenuItem = (action) => {
    setShowContextMenu(false);
    
    switch (action) {
      case 'cut':
        handleCut();
        break;
      case 'copy':
        handleCopy();
        break;
      case 'paste':
        handlePaste();
        break;
      case 'delete':
        clearCells();
        break;
      case 'insertRowAbove':
        if (activeCell) {
          const activeRow = parseInt(activeCell.match(/[0-9]+/)[0]);
          insertRowAt(activeRow - 1); // Insert before current row
        }
        break;
      case 'insertRowBelow':
        if (activeCell) {
          const activeRow = parseInt(activeCell.match(/[0-9]+/)[0]);
          insertRowAt(activeRow); // Insert after current row
        }
        break;
      case 'insertColumnLeft':
        if (activeCell) {
          const activeCol = activeCell.match(/[A-Z]+/)[0];
          const activeColIdx = visibleColumns.indexOf(activeCol);
          insertColumnAt(activeColIdx); // Insert before current column
        }
        break;
      case 'insertColumnRight':
        if (activeCell) {
          const activeCol = activeCell.match(/[A-Z]+/)[0];
          const activeColIdx = visibleColumns.indexOf(activeCol);
          insertColumnAt(activeColIdx + 1); // Insert after current column
        }
        break;
      case 'deleteRow':
        deleteRow();
        break;
      case 'deleteColumn':
        deleteColumn();
        break;
      default:
        break;
    }
  };

  // ======== ROW AND COLUMN OPERATIONS ========
  // Add a row at the end
  const addRow = () => {
    // Get the max row number
    const maxRow = Math.max(...visibleRows);
    
    // Add a new row at the end
    const newRows = [...visibleRows, maxRow + 1];
    
    // Save current state before adding row
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    setVisibleRows(newRows);
  };

  // Insert a row at a specific position
  const insertRowAt = (index) => {
    // Save current state before inserting
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    // Find the row number at the specified index or create a new one
    const targetRowIndex = Math.min(index, visibleRows.length);
    
    // Find the next consecutive row number
    const newRowNumber = Math.max(...visibleRows) + 1;
    
    // Create updated rows array with the new row inserted
    const updatedRows = [...visibleRows];
    updatedRows.splice(targetRowIndex, 0, newRowNumber);
    
    // Update the visible rows
    setVisibleRows(updatedRows);
  };

  // Add a column at the end
  const addColumn = () => {
    // Get the last column
    const lastCol = visibleColumns[visibleColumns.length - 1];
    
    // Calculate the next column letter
    const nextColIdx = columnToIndex(lastCol) + 1;
    const nextCol = indexToColumn(nextColIdx);
    
    // Save current state before adding column
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    // Add the new column
    setVisibleColumns([...visibleColumns, nextCol]);
  };

  // Insert a column at a specific position
  const insertColumnAt = (index) => {
    // Save current state before inserting
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    // Calculate the new column letter
    let newColIndex;
    if (index < visibleColumns.length) {
      // Insert between existing columns
      newColIndex = columnToIndex(visibleColumns[index]);
    } else {
      // Add at the end
      newColIndex = columnToIndex(visibleColumns[visibleColumns.length - 1]) + 1;
    }
    
    // Get the next available column letter after the last one
    const lastColIndex = columnToIndex(visibleColumns[visibleColumns.length - 1]);
    const newColLetter = indexToColumn(lastColIndex + 1);
    
    // Add the new column to visibleColumns at the correct position
    const updatedColumns = [...visibleColumns];
    updatedColumns.splice(index, 0, newColLetter);
    setVisibleColumns(updatedColumns);
  };

  // Delete a row
  const deleteRow = (rowNumber) => {
    if (visibleRows.length <= 1) {
      alert("Cannot delete the only row.");
      return;
    }
    
    // If no row number is provided, use activeCell's row
    if (rowNumber === undefined && activeCell) {
      rowNumber = parseInt(activeCell.match(/[0-9]+/)[0]);
    } else if (rowNumber === undefined) {
      return; // No row to delete
    }
    
    // Save current state before deleting row
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    setVisibleRows(visibleRows.filter(row => row !== rowNumber));
    
    // If the active cell is in this row, clear it
    if (activeCell && parseInt(activeCell.match(/[0-9]+/)[0]) === rowNumber) {
      setActiveCell(null);
    }
  };

  // Delete a column
  const deleteColumn = (colLetter) => {
    if (visibleColumns.length <= 1) {
      alert("Cannot delete the only column.");
      return;
    }
    
    // If no column letter is provided, use activeCell's column
    if (colLetter === undefined && activeCell) {
      colLetter = activeCell.match(/[A-Z]+/)[0];
    } else if (colLetter === undefined) {
      return; // No column to delete
    }
    
    // Save current state before deleting column
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    setVisibleColumns(visibleColumns.filter(col => col !== colLetter));
    
    // If the active cell is in this column, clear it
    if (activeCell && activeCell.match(/[A-Z]+/)[0] === colLetter) {
      setActiveCell(null);
    }
  };
  
  // ======== SAVE/LOAD FUNCTIONALITY ========
  // Save the spreadsheet
  const saveSpreadsheet = () => {
    const data = {
      cells,
      cellStyles,
      dependencies,
      visibleColumns,
      visibleRows,
      sheets,
      activeSheet,
      fileName,
      charts // Add this line
    };
    
    // Create a downloadable JSON file
    const blob = new Blob([JSON.stringify(data)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = `${fileName || 'spreadsheet'}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    
    alert("Spreadsheet saved successfully!");
  };

 // Replace the entire loadSpreadsheet function
const loadSpreadsheet = (e) => {
  const file = e.target.files[0];
  
  if (!file) return;
  
  // Extract file extension
  const fileName = file.name;
  const fileExt = fileName.slice(fileName.lastIndexOf('.')).toLowerCase();
  
  // Set temporary filename (will be overwritten for JSON files)
  setFileName(fileName.slice(0, fileName.lastIndexOf('.')) || "Untitled spreadsheet");
  
  // Create FileReader instance
  const reader = new FileReader();
  
  // Save current state before loading
  addToHistory({
    cells: {...cells},
    cellStyles: {...cellStyles}
  });
  
  if (fileExt === '.json') {
    // Handle JSON files as before
    reader.onload = (event) => {
      try {
        const data = JSON.parse(event.target.result);
        
        // Update state with loaded data
        setCells(data.cells || {});
        setCellStyles(data.cellStyles || {});
        setDependencies(data.dependencies || {});
        setVisibleColumns(data.visibleColumns || ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']);
        setVisibleRows(data.visibleRows || Array.from({ length: 25 }, (_, i) => i + 1));
        setSheets(data.sheets || ['Sheet1']);
        setActiveSheet(data.activeSheet || 'Sheet1');
        setCharts(data.charts || []);
        setFileName(data.fileName || fileName.slice(0, fileName.lastIndexOf('.')));
        
        alert("Spreadsheet loaded successfully!");
      } catch (error) {
        alert("Error loading spreadsheet: " + error.message);
      }
    };
    reader.readAsText(file);
  } 
  else if (fileExt === '.csv') {
    // Handle CSV files
    reader.onload = async (event) => {
      try {
        const csvData = event.target.result;
        await parseCSV(csvData);
        alert("CSV loaded successfully!");
      } catch (error) {
        alert("Error loading CSV: " + error.message);
      }
    };
    reader.readAsText(file);
  } 
  else if (fileExt === '.xlsx' || fileExt === '.xls') {
    // Handle Excel files - need to use ArrayBuffer for Excel files
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        
        parseExcel(workbook);
        alert("Excel file loaded successfully!");
      } catch (error) {
        alert("Error loading Excel file: " + error.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }
  else {
    alert("Unsupported file format. Please use .json, .csv, .xlsx, or .xls files.");
  }
};

// After your existing loadSpreadsheet function under the "SAVE/LOAD FUNCTIONALITY" section
// Add this new function
const parseCSV = (csvText) => {
  return new Promise((resolve, reject) => {
    Papa.parse(csvText, {
      header: false,
      dynamicTyping: true, // Automatically convert numbers and booleans
      skipEmptyLines: true,
      complete: (results) => {
        try {
          const rows = results.data;
          const newCells = { ...cells };
          
          // Process data into cells format
          rows.forEach((row, rowIdx) => {
            row.forEach((cellValue, colIdx) => {
              const colLetter = String.fromCharCode(65 + colIdx);
              const cellId = `${colLetter}${rowIdx + 1}`;
              
              // Convert to string
              cellValue = cellValue !== null ? String(cellValue) : "";
              
              newCells[cellId] = {
                value: cellValue,
                formula: cellValue,
                formatted: cellValue
              };
            });
          });
          
          // Calculate the visible columns needed
          const maxColumns = Math.max(...rows.map(row => row.length), 10);
          const requiredCols = Array.from({ length: maxColumns }, (_, i) => 
            String.fromCharCode(65 + i));
          
          // Update state with the CSV data
          setCells(newCells);
          setVisibleColumns(requiredCols);
          setVisibleRows(Array.from({ length: Math.max(rows.length, 25) }, (_, i) => i + 1));
          setActiveSheet('Sheet1');
          setSheets(['Sheet1']);
          
          resolve({ rows: rows.length, cols: maxColumns });
        } catch (error) {
          reject(error);
        }
      },
      error: (error) => {
        reject(error);
      }
    });
  });
};

// Add this new function after parseCSV
const parseExcel = (workbook) => {
  // Create storage for all sheet data
  const newSheetData = {};
  const workbookSheets = workbook.SheetNames.map(name => name);
  
  // Process each sheet in the workbook
  workbookSheets.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    
    const sheetCells = {};
    
    // Process cells in this sheet
    jsonData.forEach((row, rowIdx) => {
      row.forEach((cellValue, colIdx) => {
        const colLetter = String.fromCharCode(65 + colIdx);
        const cellId = `${colLetter}${rowIdx + 1}`;
        
        // Convert to string
        cellValue = cellValue !== null ? String(cellValue) : "";
        
        sheetCells[cellId] = {
          value: cellValue,
          formula: cellValue,
          formatted: cellValue
        };
      });
    });
    
    // Store this sheet's data
    newSheetData[sheetName] = {
      cells: sheetCells,
      cellStyles: {},
      dependencies: {}
    };
  });
  
  // Set the sheets array
  setSheets(workbookSheets);
  
  // Set active sheet to the first one
  const firstSheet = workbookSheets[0] || 'Sheet1';
  setActiveSheet(firstSheet);
  
  // Update sheetData state with all sheet information
  setSheetData(newSheetData);
  
  // Set the active sheet's data as current
  setCells(newSheetData[firstSheet].cells || {});
  setCellStyles(newSheetData[firstSheet].cellStyles || {});
  setDependencies(newSheetData[firstSheet].dependencies || {});
  
  // Calculate visible columns and rows based on the active sheet
  const activeSheetData = newSheetData[firstSheet].cells;
  const cellIds = Object.keys(activeSheetData);
  
  // If there are cells
  if (cellIds.length > 0) {
    // Get column letters and row numbers
    const columns = cellIds.map(id => id.match(/[A-Z]+/)[0]);
    const rows = cellIds.map(id => parseInt(id.match(/[0-9]+/)[0]));
    
    // Get max column index
    const maxColIndex = Math.max(...columns.map(col => columnToIndex(col)));
    // Get max row
    const maxRow = Math.max(...rows);
    
    // Set visible columns and rows
    setVisibleColumns(Array.from({ length: Math.max(maxColIndex + 1, 10) }, 
      (_, i) => indexToColumn(i)));
    setVisibleRows(Array.from({ length: Math.max(maxRow, 25) }, (_, i) => i + 1));
  } else {
    // Default visibility if no cells
    setVisibleColumns(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']);
    setVisibleRows(Array.from({ length: 25 }, (_, i) => i + 1));
  }
};

  // ======== DATA ANALYSIS ========
  // Export to CSV
  const exportToCsv = () => {
    // Determine the range of data
    const usedRows = [];
    const usedCols = [];
    
    Object.keys(cells).forEach(cellId => {
      if (cells[cellId]?.value) {
        const col = cellId.match(/[A-Z]+/)[0];
        const row = parseInt(cellId.match(/[0-9]+/)[0]);
        
        if (!usedCols.includes(col)) usedCols.push(col);
        if (!usedRows.includes(row)) usedRows.push(row);
      }
    });
    
    if (usedRows.length === 0 || usedCols.length === 0) {
      alert("No data to export");
      return;
    }
    
    // Sort rows and columns
    usedRows.sort((a, b) => a - b);
    usedCols.sort((a, b) => {
      return columnToIndex(a) - columnToIndex(b);
    });
    
    // Generate CSV content
    let csv = "";
    
    // Add header row if requested
    const includeHeaders = true;
    if (includeHeaders) {
      csv += "," + usedCols.join(",") + "\n";
    }
    
    // Add data rows
    usedRows.forEach(row => {
      let rowData = includeHeaders ? row + "," : "";
      
      usedCols.forEach((col, index) => {
        const cellId = `${col}${row}`;
        let cellValue = cells[cellId]?.value || "";
        
        // Escape commas and quotes
        if (cellValue.includes(",") || cellValue.includes('"')) {
          cellValue = `"${cellValue.replace(/"/g, '""')}"`;
        }
        
        rowData += cellValue;
        if (index < usedCols.length - 1) rowData += ",";
      });
      
      csv += rowData + "\n";
    });
    
    // Create a download link
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `${fileName}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Find and replace
  const handleFindReplace = () => {
    const findValue = prompt("Find:");
    if (findValue === null) return;
    
    const replaceValue = prompt("Replace with:");
    if (replaceValue === null) return;
    
    // Save current state before find/replace
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    const newCells = { ...cells };
    let replacementCount = 0;
    
    // If range is selected, search only in that range
    if (selectedRange) {
      const cellsInRange = getCellsInRange(selectedRange.start, selectedRange.end);
      
      cellsInRange.forEach(cellId => {
        if (newCells[cellId]?.value.includes(findValue)) {
          const oldValue = newCells[cellId].value;
          const newValue = oldValue.replaceAll(findValue, replaceValue);
          
          newCells[cellId] = {
            ...newCells[cellId],
            value: newValue,
            formula: newCells[cellId].formula.startsWith('=') ? 
              newCells[cellId].formula : newValue,
            formatted: formatValue(newValue, cellStyles[cellId])
          };
          
          replacementCount += (oldValue.match(new RegExp(findValue, 'g')) || []).length;
        }
      });
    } else {
      // Search entire sheet
      Object.keys(newCells).forEach(cellId => {
        if (newCells[cellId]?.value.includes(findValue)) {
          const oldValue = newCells[cellId].value;
          const newValue = oldValue.replaceAll(findValue, replaceValue);
          
          newCells[cellId] = {
            ...newCells[cellId],
            value: newValue,
            formula: newCells[cellId].formula.startsWith('=') ? 
              newCells[cellId].formula : newValue,
            formatted: formatValue(newValue, cellStyles[cellId])
          };
          
          replacementCount += (oldValue.match(new RegExp(findValue, 'g')) || []).length;
        }
      });
    }
    
    setCells(newCells);
    alert(`Replaced ${replacementCount} occurrences.`);
  };

  // Remove duplicates
  const handleRemoveDuplicates = () => {
    if (!selectedRange) {
      alert("Please select a range to remove duplicates");
      return;
    }
    
    const { start, end } = selectedRange;
    const startCol = start.match(/[A-Z]+/)[0];
    const endCol = end.match(/[A-Z]+/)[0];
    
    if (startCol !== endCol) {
      alert("Please select a single column to remove duplicates");
      return;
    }
    
    // Save current state before removing duplicates
    addToHistory({
      cells: {...cells},
      cellStyles: {...cellStyles}
    });
    
    const cellsInRange = getCellsInRange(start, end);
    const values = cellsInRange.map(cellId => cells[cellId]?.value);
    const uniqueValues = [...new Set(values)];
    
    const newCells = { ...cells };
    
    // Update cells with unique values
    for (let i = 0; i < cellsInRange.length; i++) {
      if (i < uniqueValues.length) {
        newCells[cellsInRange[i]] = {
          value: uniqueValues[i],
          formula: uniqueValues[i],
          formatted: formatValue(uniqueValues[i], cellStyles[cellsInRange[i]])
        };
      } else {
        newCells[cellsInRange[i]] = {
          value: "",
          formula: "",
          formatted: ""
        };
      }
    }
    
    setCells(newCells);
    alert(`Removed ${values.length - uniqueValues.length} duplicate rows.`);
  };

  // Add cell comment
  const addCellComment = () => {
    if (!activeCell) return;
    
    setCommentCell(activeCell);
    setCommentText(cellComments[activeCell] || "");
    setShowCommentPopup(true);
  };

  // Save cell comment
  const saveCellComment = () => {
    if (!commentCell) return;
    
    const newCellComments = { ...cellComments };
    
    if (commentText.trim() === "") {
      delete newCellComments[commentCell];
    } else {
      newCellComments[commentCell] = commentText;
    }
    
    setCellComments(newCellComments);
    setShowCommentPopup(false);
    setCommentCell(null);
    setCommentText("");
  };

  // Create menu actions object
  const menuActions = {
    newSpreadsheet: () => {
      // Reset spreadsheet to initial state
      const initialCells = {};
      for (let row = 1; row <= 100; row++) {
        for (let col = 0; col < 26; col++) {
          const colLetter = String.fromCharCode(65 + col);
          const cellId = `${colLetter}${row}`;
          initialCells[cellId] = { value: "", formula: "", formatted: "" };
        }
      }
      setCells(initialCells);
      setFileName("Untitled spreadsheet");
      setActiveSheet('Sheet1');
      setSheets(['Sheet1']);
      
      // Clear styles and other states
      setCellStyles({});
      setCellComments({});
      setDependencies({});
      
      // Add initial state to history
      addToHistory({
        cells: initialCells,
        cellStyles: {}
      });
    },
    
    openSpreadsheet: () => {
      // Trigger file input
      document.getElementById('fileInput').click();
    },
    
    saveSpreadsheet: () => {
      saveSpreadsheet();
    },
    
    exportToCsv: () => {
      exportToCsv();
    },
    
    printSpreadsheet: () => {
      window.print();
    },
    
    undo: () => {
      handleUndo();
    },
    
    redo: () => {
      handleRedo();
    },
    
    cut: () => {
      handleCut();
    },
    
    copy: () => {
      handleCopy();
    },
    
    paste: () => {
      handlePaste();
    },
    
    findReplace: () => {
      handleFindReplace();
    },
    
    insertRow: () => {
      if (activeCell) {
        const activeRow = parseInt(activeCell.match(/[0-9]+/)[0]);
        insertRowAt(activeRow);
      } else {
        addRow();
      }
    },
    
    insertColumn: () => {
      if (activeCell) {
        const activeCol = activeCell.match(/[A-Z]+/)[0];
        const activeColIdx = visibleColumns.indexOf(activeCol);
        insertColumnAt(activeColIdx);
      } else {
        addColumn();
      }
    },
    
    toggleBold: () => {
      applyFormatting('bold');
    },
    
    toggleItalic: () => {
      applyFormatting('italic');
    },
    
    toggleUnderline: () => {
      applyFormatting('underline');
    },
    
    dataValidation: () => {
      // Open validation dialog
      alert("Data validation feature");
    },
    
    removeDuplicates: () => {
      handleRemoveDuplicates();
    },
    
    showHelp: () => {
      alert(
        "Keyboard Shortcuts:\n\n" +
        "Ctrl+B: Bold\n" +
        "Ctrl+I: Italic\n" +
        "Ctrl+U: Underline\n" +
        "Ctrl+Z: Undo\n" +
        "Ctrl+Y: Redo\n" +
        "Ctrl+C: Copy\n" +
        "Ctrl+X: Cut\n" +
        "Ctrl+V: Paste\n" +
        "Enter: Move down\n" +
        "Tab: Move right\n" +
        "F2: Edit cell\n" +
        "Delete: Clear cell"
      );
    },
    
    showAbout: () => {
      alert("Google Sheets Clone\nCreated as a React application\n© 2023");
    },

    // Additional actions
    zoomIn: () => alert("Zoom In feature"),
    zoomOut: () => alert("Zoom Out feature"),
    toggleFullScreen: () => {
      if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen();
      } else {
        if (document.exitFullscreen) {
          document.exitFullscreen();
        }
      }
    },
    showFormulas: () => alert("Show Formulas feature"),
    insertFunction: () => alert("Insert Function feature"),
    numberFormat: () => alert("Number Format feature"),
    sortData: () => alert("Sort Data feature"),
    filterData: () => alert("Filter Data feature"),
    checkSpelling: () => alert("Check Spelling feature"),
    wordCount: () => alert("Word Count feature"),
    protectSheet: () => alert("Protect Sheet feature"),
    addCellComment: () => addCellComment()
  };

  // Define menu items with their actions
  const menuItems = {
    'File': [
      { label: 'New', action: menuActions.newSpreadsheet, shortcut: 'Ctrl+N' },
      { label: 'Open...', action: menuActions.openSpreadsheet, shortcut: 'Ctrl+O' },
      { label: 'Save', action: menuActions.saveSpreadsheet, shortcut: 'Ctrl+S' },
      { label: 'Export to CSV', action: menuActions.exportToCsv, shortcut: 'Ctrl+E' },
      { label: 'Print', action: menuActions.printSpreadsheet, shortcut: 'Ctrl+P' },
    ],
    'Edit': [
      { label: 'Undo', action: menuActions.undo, shortcut: 'Ctrl+Z' },
      { label: 'Redo', action: menuActions.redo, shortcut: 'Ctrl+Y' },
      { label: 'Cut', action: menuActions.cut, shortcut: 'Ctrl+X' },
      { label: 'Copy', action: menuActions.copy, shortcut: 'Ctrl+C' },
      { label: 'Paste', action: menuActions.paste, shortcut: 'Ctrl+V' },
      { label: 'Find and Replace', action: menuActions.findReplace, shortcut: 'Ctrl+H' }
    ],
    'View': [
      { label: 'Zoom In', action: menuActions.zoomIn, shortcut: 'Ctrl++' },
      { label: 'Zoom Out', action: menuActions.zoomOut, shortcut: 'Ctrl+-' },
      { label: 'Full Screen', action: menuActions.toggleFullScreen, shortcut: 'F11' },
      { label: 'Show Formulas', action: menuActions.showFormulas, shortcut: 'Ctrl+`' }
    ],
    'Insert': [
      { label: 'Row', action: menuActions.insertRow, shortcut: 'Alt+I R' },
      { label: 'Column', action: menuActions.insertColumn, shortcut: 'Alt+I C' },
      { label: 'Chart', action: openChartDialog, shortcut: 'Alt+I H' },
      { label: 'Comment', action: menuActions.addCellComment },
      { label: 'Function', action: menuActions.insertFunction, shortcut: 'Shift+F3' }
    ],
    'Format': [
      { label: 'Bold', action: () => applyFormatting('bold'), shortcut: 'Ctrl+B' },
      { label: 'Italic', action: () => applyFormatting('italic'), shortcut: 'Ctrl+I' },
      { label: 'Underline', action: () => applyFormatting('underline'), shortcut: 'Ctrl+U' },
      { label: 'Number Format', action: menuActions.numberFormat }
    ],
    'Data': [
      { label: 'Sort Sheet', action: menuActions.sortData },
      { label: 'Filter', action: menuActions.filterData },
      { label: 'Remove Duplicates', action: handleRemoveDuplicates }
    ],
    'Tools': [
      { label: 'Spell Check', action: menuActions.checkSpelling },
      { label: 'Word Count', action: menuActions.wordCount },
      { label: 'Protect Sheet', action: menuActions.protectSheet }
    ],
    'Help': [
      { label: 'Keyboard Shortcuts', action: menuActions.showHelp, shortcut: 'F1' },
      { label: 'About', action: menuActions.showAbout }
    ]
  };

  // ======== COMPONENT RENDER ========
  return (
    <div 
      style={{ 
        fontFamily: 'Arial, sans-serif', 
        fontSize: '12px',
        height: '100vh',
        display: 'flex',
        flexDirection: 'column',
        overflow: 'hidden'
      }}
      tabIndex={0}
      onKeyDown={handleKeyboardShortcut}
    >
      {/* Menu Bar */}
      <div style={{
        display: 'flex',
        padding: '8px 16px',
        borderBottom: '1px solid #ccc',
        backgroundColor: '#f8f9fa'
      }}>
        <div style={{ 
          display: 'flex', 
          alignItems: 'center',
          marginRight: '20px'
        }}>
          <span style={{ 
            width: '24px', 
            height: '24px', 
            backgroundColor: '#0f9d58', 
            borderRadius: '4px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            color: 'white',
            marginRight: '8px',
            fontSize: '16px'
          }}>
            <span role="img" aria-label="Menu">≡</span>
          </span>
          <input 
            type="text" 
            value={fileName}
            onChange={(e) => setFileName(e.target.value)}
            style={{
              border: 'none',
              outline: 'none',
              fontSize: '18px',
              fontWeight: 'normal',
              width: '250px',
              backgroundColor: 'transparent'
            }}
          />
          <span style={{ 
            color: '#5f6368', 
            marginLeft: '8px',
            cursor: 'pointer' 
          }}>
            <span role="img" aria-label="Star">★</span>
          </span>
        </div>
        
        {/* Menu Items */}
        <div style={{ display: 'flex' }}>
          {Object.keys(menuItems).map(menuName => (
            <div 
              key={menuName}
              className="menu-item"
              style={{ 
                padding: '8px 12px', 
                cursor: 'pointer',
                borderRadius: '4px',
                backgroundColor: activeMenu === menuName ? '#f1f3f4' : 'transparent',
                transition: 'background-color 0.2s ease',
                position: 'relative'
              }}
              onMouseOver={(e) => {
                if (activeMenu && activeMenu !== menuName) {
                  setActiveMenu(menuName);
                } else if (!activeMenu) {
                  e.currentTarget.style.backgroundColor = '#f1f3f4';
                }
              }}
              onMouseOut={(e) => {
                if (!activeMenu) {
                  e.currentTarget.style.backgroundColor = 'transparent';
                }
              }}
              onClick={() => handleMenuClick(menuName)}
            >
              {menuName}
              
              {/* Dropdown Menu */}
              {activeMenu === menuName && (
                <div 
                  style={{
                    position: 'absolute',
                    top: '100%',
                    left: '0',
                    backgroundColor: 'white',
                    border: '1px solid #ddd',
                    borderRadius: '4px',
                    boxShadow: '0 2px 10px rgba(0,0,0,0.2)',
                    zIndex: 1000,
                    minWidth: '180px',
                    marginTop: '4px'
                  }}
                >
                  {menuItems[menuName].map((item, index) => (
                    <div 
                      key={index}
                      style={{
                        padding: '8px 16px',
                        display: 'flex',
                        justifyContent: 'space-between',
                        cursor: 'pointer',
                        transition: 'background-color 0.2s ease',
                        whiteSpace: 'nowrap'
                      }}
                      onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
                      onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
                      onClick={(e) => {
                        e.stopPropagation(); // Prevent the click from bubbling to parent
                        setActiveMenu(null); // Close menu after action
                        item.action(); // Execute the action
                      }}
                    >
                      <span>{item.label}</span>
                      {item.shortcut && (
                        <span style={{ 
                          marginLeft: '24px', 
                          fontSize: '11px', 
                          color: '#5f6368',
                          fontFamily: 'monospace'
                        }}>
                          {item.shortcut}
                        </span>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>

        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center' }}>
          <div style={{ 
            padding: '8px', 
            borderRadius: '50%',
            marginRight: '8px',
            cursor: 'pointer',
            transition: 'background-color 0.2s ease'
          }}
          onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
          onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
          >
            <span role="img" aria-label="Comments">💬</span>
          </div>
          <button 
            style={{ 
              backgroundColor: '#c2e7ff',
              color: '#001d35',
              border: 'none',
              borderRadius: '24px',
              padding: '8px 24px',
              fontWeight: '500',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              transition: 'background-color 0.2s ease'
            }}
            onClick={menuActions.saveSpreadsheet}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#b3deff'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = '#c2e7ff'}
          >
            <span role="img" aria-label="Lock" style={{ marginRight: '8px' }}>🔒</span>
            Share
          </button>
        </div>
      </div>
      
      {/* Toolbar */}

      <div style={{
        display: 'flex',
        padding: '4px 16px',
        borderBottom: '1px solid #ccc',
        backgroundColor: '#f8f9fa',
        alignItems: 'center'
      }}>
        <div style={{ display: 'flex', alignItems: 'center', marginRight: '16px' }}>
          <button 
            title="Undo (Ctrl+Z)"
            style={{ 
              padding: '6px',
              backgroundColor: 'transparent',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px',
              transition: 'background-color 0.2s ease',
              opacity: historyIndex > 0 ? 1 : 0.5
            }}
            onClick={handleUndo}
            disabled={historyIndex <= 0}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = historyIndex > 0 ? '#f1f3f4' : 'transparent'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
          >
            <span role="img" aria-label="Undo">↩️</span>
          </button>
          <button 
            title="Redo (Ctrl+Y)"
            style={{ 
              padding: '6px',
              backgroundColor: 'transparent',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px',
              transition: 'background-color 0.2s ease',
              opacity: historyIndex < history.length - 1 ? 1 : 0.5
            }}
            onClick={handleRedo}
            disabled={historyIndex >= history.length - 1}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = historyIndex < history.length - 1 ? '#f1f3f4' : 'transparent'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
          >
            <span role="img" aria-label="Redo">↪️</span>
          </button>
          <button 
            title="Print (Ctrl+P)"
            style={{ 
              padding: '6px',
              backgroundColor: 'transparent',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px',
              transition: 'background-color 0.2s ease'
            }}
            onClick={() => window.print()}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
          >
            <span role="img" aria-label="Print">🖨️</span>
          </button>
          {/* Add this button to your toolbar, alongside other toolbar buttons */}
<button 
  title="Insert Chart"
  style={{ 
    padding: '6px 10px',
    backgroundColor: 'transparent',
    border: '1px solid #ddd',
    borderRadius: '4px',
    cursor: 'pointer',
    marginRight: '8px',
    transition: 'background-color 0.2s ease'
  }}
  onClick={openChartDialog}
  onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
  onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
>
  <span role="img" aria-label="Chart">📊</span>
</button>
          
        </div>
        
        <select 
          style={{ 
            padding: '4px 8px',
            border: '1px solid #ddd',
            borderRadius: '4px',
            marginRight: '8px',
            cursor: 'pointer'
          }}
          value={fontSize}
          onChange={(e) => {
            const newSize = parseInt(e.target.value);
            setFontSize(newSize);
            applyFormatting('fontSize', newSize);
          }}
        >
          <option value="8">8</option>
          <option value="9">9</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
          <option value="14">14</option>
          <option value="18">18</option>
          <option value="24">24</option>
          <option value="36">36</option>
        </select>
        
        {/* Font family dropdown */}
        <select 
          style={{ 
            padding: '4px 8px',
            border: '1px solid #ddd',
            borderRadius: '4px',
            marginRight: '16px',
            width: '120px',
            cursor: 'pointer'
          }}
          value={fontFamily}
          onChange={(e) => {
            const newFont = e.target.value;
            setFontFamily(newFont);
            applyFormatting('fontFamily', newFont);
          }}
        >
          <option value="Arial">Arial</option>
          <option value="Calibri">Calibri</option>
          <option value="Cambria">Cambria</option>
          <option value="Georgia">Georgia</option>
          <option value="Helvetica">Helvetica</option>
          <option value="Times New Roman">Times New Roman</option>
          <option value="Verdana">Verdana</option>
          <option value="Roboto">Roboto</option>
        </select>
        
        <button 
          title="Bold (Ctrl+B)"
          style={{ 
            padding: '6px 10px',
            backgroundColor: cellStyles[activeCell]?.fontWeight === 'bold' ? '#e8f0fe' : 'transparent',
            border: '1px solid #ddd',
            borderRadius: '4px',
            cursor: 'pointer',
            marginRight: '4px',
            fontWeight: 'bold',
            color: cellStyles[activeCell]?.fontWeight === 'bold' ? '#1a73e8' : 'inherit',
            transition: 'all 0.2s ease'
          }}
          onClick={() => applyFormatting('bold')}
          onMouseOver={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.fontWeight === 'bold' ? '#d3e0fd' : '#f1f3f4'}
          onMouseOut={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.fontWeight === 'bold' ? '#e8f0fe' : 'transparent'}
        >B</button>
        <button
          title="Italic (Ctrl+I)"
          style={{ 
            padding: '6px 10px',
            backgroundColor: cellStyles[activeCell]?.fontStyle === 'italic' ? '#e8f0fe' : 'transparent',
            border: '1px solid #ddd',
            borderRadius: '4px',
            cursor: 'pointer',
            marginRight: '4px',
            fontStyle: 'italic',
            color: cellStyles[activeCell]?.fontStyle === 'italic' ? '#1a73e8' : 'inherit',
            transition: 'all 0.2s ease'
          }}
          onClick={() => applyFormatting('italic')}
          onMouseOver={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.fontStyle === 'italic' ? '#d3e0fd' : '#f1f3f4'}
          onMouseOut={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.fontStyle === 'italic' ? '#e8f0fe' : 'transparent'}
        >I</button>
        <button
          title="Underline (Ctrl+U)"
          style={{ 
            padding: '6px 10px',
            backgroundColor: cellStyles[activeCell]?.textDecoration === 'underline' ? '#e8f0fe' : 'transparent',
            border: '1px solid #ddd',
            borderRadius: '4px',
            cursor: 'pointer',
            marginRight: '16px',
            textDecoration: 'underline',
            color: cellStyles[activeCell]?.textDecoration === 'underline' ? '#1a73e8' : 'inherit',
            transition: 'all 0.2s ease'
          }}
          onClick={() => applyFormatting('underline')}
          onMouseOver={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.textDecoration === 'underline' ? '#d3e0fd' : '#f1f3f4'}
          onMouseOut={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.textDecoration === 'underline' ? '#e8f0fe' : 'transparent'}
        >U</button>
        
        <div style={{ display: 'flex', marginRight: '16px' }}>
          {[
            { align: 'left', symbol: '⟨' }, 
            { align: 'center', symbol: '≡' }, 
            { align: 'right', symbol: '⟩' }
          ].map((item, index) => (
            <button 
              key={index}
              style={{ 
                padding: '6px 10px',
                backgroundColor: cellStyles[activeCell]?.textAlign === item.align ? '#e8f0fe' : 'transparent',
                border: '1px solid #ddd',
                borderLeft: index > 0 ? 'none' : '1px solid #ddd',
                borderTopLeftRadius: index === 0 ? '4px' : '0',
                borderBottomLeftRadius: index === 0 ? '4px' : '0',
                borderTopRightRadius: index === 2 ? '4px' : '0',
                borderBottomRightRadius: index === 2 ? '4px' : '0',
                cursor: 'pointer',
                color: cellStyles[activeCell]?.textAlign === item.align ? '#1a73e8' : 'inherit',
                transition: 'all 0.2s ease'
              }}
              onClick={() => applyFormatting('textAlign', item.align)}
              onMouseOver={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.textAlign === item.align ? '#d3e0fd' : '#f1f3f4'}
              onMouseOut={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.textAlign === item.align ? '#e8f0fe' : 'transparent'}
            >
              {item.symbol}
            </button>
          ))}
        </div>
        
        <div style={{ display: 'flex', marginRight: '16px' }}>
          <button 
            style={{ 
              padding: '6px 10px',
              backgroundColor: 'transparent',
              border: '1px solid #ddd',
              borderRadius: '4px',
              cursor: 'pointer',
              marginRight: '4px',
              transition: 'background-color 0.2s ease',
              position: 'relative'
            }}
            onClick={(e) => toggleColorPicker('text', e.currentTarget)}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
          >
            <span style={{ 
              display: 'inline-block', 
              width: '12px', 
              height: '12px', 
              backgroundColor: cellStyles[activeCell]?.color || '#000', 
              marginRight: '4px',
              border: '1px solid #ddd'
            }}></span>
            A
          </button>
          <button 
            style={{ 
              padding: '6px 10px',
              backgroundColor: 'transparent',
              border: '1px solid #ddd',
              borderRadius: '4px',
              cursor: 'pointer',
              transition: 'background-color 0.2s ease',
              position: 'relative'
            }}
            onClick={(e) => toggleColorPicker('background', e.currentTarget)}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
          >
            <span style={{ 
              display: 'inline-block', 
              width: '12px', 
              height: '12px', 
              backgroundColor: cellStyles[activeCell]?.backgroundColor || '#fff', 
              marginRight: '4px',
              border: '1px solid #ddd'
            }}></span>
            <span role="img" aria-label="Color Palette">🎨</span>
          </button>
        </div>
        
        <div style={{ display: 'flex', marginRight: '16px' }}>
          <button 
            style={{ 
              padding: '6px 10px',
              backgroundColor: cellStyles[activeCell]?.numberFormat === 'percentage' ? '#e8f0fe' : 'transparent',
              border: '1px solid #ddd',
              borderTopLeftRadius: '4px',
              borderBottomLeftRadius: '4px',
              cursor: 'pointer',
              color: cellStyles[activeCell]?.numberFormat === 'percentage' ? '#1a73e8' : 'inherit',
              transition: 'all 0.2s ease'
            }}
            onClick={() => applyFormatting('numberFormat', 'percentage')}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.numberFormat === 'percentage' ? '#d3e0fd' : '#f1f3f4'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.numberFormat === 'percentage' ? '#e8f0fe' : 'transparent'}
          >%</button>
          <button 
            style={{ 
              padding: '6px 10px',
              backgroundColor: cellStyles[activeCell]?.numberFormat === 'currency' ? '#e8f0fe' : 'transparent',
              border: '1px solid #ddd',
              borderLeft: 'none',
              borderTopRightRadius: '4px',
              borderBottomRightRadius: '4px',
              cursor: 'pointer',
              color: cellStyles[activeCell]?.numberFormat === 'currency' ? '#1a73e8' : 'inherit',
              transition: 'all 0.2s ease'
            }}
            onClick={() => applyFormatting('numberFormat', 'currency')}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.numberFormat === 'currency' ? '#d3e0fd' : '#f1f3f4'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = cellStyles[activeCell]?.numberFormat === 'currency' ? '#e8f0fe' : 'transparent'}
          >$</button>
        </div>
        
        {/* Action buttons */}
        <div style={{ marginLeft: 'auto', display: 'flex' }}>
          <button 
            title="Find and Replace"
            style={{ 
              padding: '6px 10px',
              backgroundColor: 'transparent',
              border: '1px solid #ddd',
              borderRadius: '4px',
              cursor: 'pointer',
              marginRight: '8px',
              transition: 'background-color 0.2s ease'
            }}
            onClick={handleFindReplace}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
          >
            <span role="img" aria-label="Find and Replace">🔍</span>
          </button>
          <button 
            title="Remove Duplicates"
            style={{ 
              padding: '6px 10px',
              backgroundColor: 'transparent',
              border: '1px solid #ddd',
              borderRadius: '4px',
              cursor: 'pointer',
              marginRight: '8px',
              transition: 'background-color 0.2s ease'
            }}
            onClick={handleRemoveDuplicates}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
          >
            <span role="img" aria-label="Remove Duplicates">🧹</span>
          </button>
        </div>
      </div>

      {/* Formula Bar (continued) */}
      <div style={{
        display: 'flex',
        padding: '4px 8px',
        borderBottom: '1px solid #ccc',
        backgroundColor: 'white',
        alignItems: 'center'
      }}>
        <div style={{ 
          width: '30px',
          height: '30px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: '#f1f3f4',
          borderRadius: '4px',
          marginRight: '8px',
          fontSize: '16px'
        }}>
          fx
        </div>
        <input 
          id="formula-cell-ref"
          value={activeCell || ''}
          readOnly
          style={{
            width: '60px',
            padding: '6px',
            border: '1px solid #ddd',
            borderRadius: '4px',
            marginRight: '8px',
            textAlign: 'center',
            backgroundColor: '#f8f9fa'
          }}
        />
        <input 
          type="text" 
          value={inputValue}
          onChange={handleFormulaChange}
          onBlur={applyFormula}
          onKeyDown={(e) => e.key === 'Enter' && applyFormula()}
          style={{
            flex: 1,
            padding: '6px',
            border: '1px solid #ddd',
            borderRadius: '4px'
          }}
        />
      </div>
      
      {/* Spreadsheet Grid */}
      <div 
        ref={gridRef}
        tabIndex={0}
        onKeyDown={handleGridKeyDown}
        style={{
          flex: 1,
          overflow: 'auto',
          position: 'relative',
          outline: 'none'
        }}
      >
        <div style={{
          display: 'grid',
          gridTemplateColumns: `40px repeat(${visibleColumns.length}, 100px)`,
          gridTemplateRows: `30px repeat(${visibleRows.length}, 25px)`,
          backgroundColor: '#f8f9fa',
          color: '#5f6368'
        }}>
          {/* Top-left corner */}
          <div style={{ 
            gridRow: '1',
            gridColumn: '1',
            borderRight: '1px solid #ddd',
            borderBottom: '1px solid #ddd',
            backgroundColor: '#f8f9fa',
            position: 'sticky',
            top: 0,
            left: 0,
            zIndex: 3
          }}></div>
          
          {/* Column headers */}
          {visibleColumns.map((col, idx) => (
            <div 
              key={col}
              style={{ 
                gridRow: '1',
                gridColumn: `${idx + 2}`,
                borderRight: '1px solid #ddd',
                borderBottom: '1px solid #ddd',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontWeight: '500',
                position: 'sticky',
                top: 0,
                zIndex: 2,
                backgroundColor: '#f8f9fa'
              }}
            >
              {col}
            </div>
          ))}
          
          {/* Row headers */}
          {visibleRows.map((row, idx) => (
            <div 
              key={row}
              style={{ 
                gridRow: `${idx + 2}`,
                gridColumn: '1',
                borderRight: '1px solid #ddd',
                borderBottom: '1px solid #ddd',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontWeight: '500',
                position: 'sticky',
                left: 0,
                zIndex: 2,
                backgroundColor: '#f8f9fa'
              }}
            >
              {row}
            </div>
          ))}
          
          {/* Cells */}
          {visibleRows.map((row, rowIdx) => (
            visibleColumns.map((col, colIdx) => {
              const cellId = `${col}${row}`;
              
              const cell = cells[cellId] || { value: "", formula: "", formatted: "" };
              const isActive = activeCell === cellId;
              const isInSelectedRange = isCellInRange(cellId);
              const isEditing = editingCell === cellId;
              const hasComment = cellComments[cellId];
              
              // Determine background color with priorities
              let cellBackgroundColor = 
                isActive ? '#e8f0fe' : 
                isInSelectedRange ? '#f1f3f4' : 
                cellStyles[cellId]?.backgroundColor || 'white';

              // Determine text color
              let cellTextColor = cellStyles[cellId]?.color || 'black';
              
              return (
                <div 
                  key={cellId}
                  style={{ 
                    gridRow: `${rowIdx + 2}`,
                    gridColumn: `${colIdx + 2}`,
                    borderRight: '1px solid #ddd',
                    borderBottom: '1px solid #ddd',
                    padding: isEditing ? '0' : '4px',
                    cursor: 'cell',
                    backgroundColor: cellBackgroundColor,
                    color: cellTextColor,
                    outline: isActive ? '2px solid #1a73e8' : 'none',
                    position: 'relative',
                    overflow: 'hidden',
                    whiteSpace: 'nowrap',
                    textOverflow: 'ellipsis',
                    fontFamily: cellStyles[cellId]?.fontFamily || fontFamily,
                    fontSize: cellStyles[cellId]?.fontSize || `${fontSize}px`,
                    fontWeight: cellStyles[cellId]?.fontWeight || 'normal',
                    fontStyle: cellStyles[cellId]?.fontStyle || 'normal',
                    textDecoration: cellStyles[cellId]?.textDecoration || 'none',
                    textAlign: cellStyles[cellId]?.textAlign || 'left',
                    zIndex: isActive || isEditing ? 5 : 1
                  }}
                  onMouseDown={(e) => handleCellMouseDown(cellId, e)}
                  onMouseOver={() => handleCellMouseOver(cellId)}
                  onContextMenu={(e) => {
                    e.preventDefault();
                    setContextMenuPosition({ top: e.clientY, left: e.clientX });
                    setShowContextMenu(true);
                    setActiveCell(cellId);
                  }}
                >
                  {isEditing ? (
                    <input
                      ref={cellInputRef}
                      type="text"
                      value={inputValue}
                      onChange={handleCellInputChange}
                      onBlur={finishCellEditing}
                      onKeyDown={handleCellInputKeyDown}
                      onPaste={(e) => {
                        // Allow default paste behavior and update state after paste
                        setTimeout(() => {
                          if (cellInputRef.current) {
                            setInputValue(cellInputRef.current.value);
                          }
                        }, 0);
                      }}
                      style={{
                        width: '100%',
                        height: '100%',
                        padding: '4px',
                        border: 'none',
                        outline: 'none',
                        fontFamily: 'inherit',
                        fontSize: 'inherit',
                        fontWeight: 'inherit',
                        fontStyle: 'inherit',
                        textAlign: cellStyles[cellId]?.textAlign || 'left',
                        color: cellStyles[cellId]?.color || 'inherit',
                        backgroundColor: 'white',
                      }}
                      autoFocus
                    />
                  ) : (
                    <>
                      <div style={{ height: '100%', display: 'flex', alignItems: 'center' }}>
                        {cell.formatted || cell.value}
                      </div>
                      {hasComment && (
                        <div
                          style={{
                            position: 'absolute',
                            top: 0,
                            right: 0,
                            width: 0,
                            height: 0,
                            borderLeft: '6px solid transparent',
                            borderRight: '6px solid #f44336',
                            borderBottom: '6px solid transparent',
                            transform: 'rotate(45deg)',
                            cursor: 'pointer'
                          }}
                          onClick={() => {
                            setCommentCell(cellId);
                            setCommentText(cellComments[cellId]);
                            setShowCommentPopup(true);
                          }}
                          title={cellComments[cellId]}
                        />
                      )}
                    </>
                  )}
                </div>
              );
            })
          ))}
        </div>
      </div>
      
      {/* Sheet Tabs */}
      <div style={{
        display: 'flex',
        borderTop: '1px solid #ccc',
        backgroundColor: '#f8f9fa',
        padding: '4px 16px'
      }}>
        <div style={{ display: 'flex', alignItems: 'center' }}>
          {sheets.map(sheet => (
            <div 
              key={sheet}
              style={{
                padding: '8px 16px',
                borderRight: '1px solid #ddd',
                backgroundColor: activeSheet === sheet ? 'white' : '#f8f9fa',
                borderTop: activeSheet === sheet ? '2px solid #1a73e8' : 'none',
                cursor: 'pointer',
                transition: 'background-color 0.2s ease',
                position: 'relative'
              }}
              onClick={() => setActiveSheet(sheet)}
              onMouseOver={(e) => {
                if (activeSheet !== sheet) {
                  e.currentTarget.style.backgroundColor = '#f1f3f4';
                }
              }}
              onMouseOut={(e) => {
                if (activeSheet !== sheet) {
                  e.currentTarget.style.backgroundColor = '#f8f9fa';
                }
              }}
            >
              {sheet}
            </div>
          ))}
          <div 
            style={{
              padding: '8px',
              cursor: 'pointer',
              color: '#5f6368',
              fontSize: '16px',
              transition: 'background-color 0.2s ease',
              borderRadius: '50%'
            }}
            onClick={() => {
              // Add a new sheet
              const newSheetNumber = sheets.length + 1;
              const newSheetName = `Sheet${newSheetNumber}`;
              setSheets([...sheets, newSheetName]);
              setActiveSheet(newSheetName);
            }}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
            title="Add sheet"
          >
            +
          </div>
        </div>
        
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center' }}>
          <input 
            type="file" 
            id="fileInput" 
            style={{ display: 'none' }}
            onChange={loadSpreadsheet}
            accept=".json,.csv,.xlsx,.xls"
          />
          <button 
            style={{
              padding: '6px 10px',
              backgroundColor: 'transparent',
              border: '1px solid #ddd',
              borderRadius: '4px',
              cursor: 'pointer',
              marginRight: '8px',
              transition: 'background-color 0.2s ease'
            }}
            onClick={() => document.getElementById('fileInput').click()}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
            title="Open spreadsheet"
          >
            <span role="img" aria-label="Open">📂 Open</span>
          </button>
          <button 
            style={{
              padding: '6px 10px',
              backgroundColor: 'transparent',
              border: '1px solid #ddd',
              borderRadius: '4px',
              cursor: 'pointer',
              marginRight: '8px',
              transition: 'background-color 0.2s ease'
            }}
            onClick={saveSpreadsheet}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
            title="Save spreadsheet"
          >
            <span role="img" aria-label="Save">💾 Save</span>
          </button>
          <button 
            style={{
              padding: '6px 10px',
              backgroundColor: 'transparent',
              border: '1px solid #ddd',
              borderRadius: '4px',
              cursor: 'pointer',
              transition: 'background-color 0.2s ease'
            }}
            onClick={exportToCsv}
            onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f3f4'}
            onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
            title="Export to CSV"
          >
            <span role="img" aria-label="Export">📊 Export</span>
          </button>
        </div>
      </div>
      
      {/* Context Menu */}
      {showContextMenu && (
        <div
          style={{
            position: 'absolute',
            top: contextMenuPosition.top,
            left: contextMenuPosition.left,
            backgroundColor: 'white',
            border: '1px solid #ddd',
            borderRadius: '4px',
            padding: '4px 0',
            boxShadow: '0 2px 10px rgba(0,0,0,0.2)',
            zIndex: 1000,
            minWidth: '180px'
          }}
        >
          {[
            { action: 'cut', label: 'Cut', shortcut: 'Ctrl+X' },
            { action: 'copy', label: 'Copy', shortcut: 'Ctrl+C' },
            { action: 'paste', label: 'Paste', shortcut: 'Ctrl+V' },
            { isDivider: true },
            { action: 'delete', label: 'Delete', shortcut: 'Delete' },
            { isDivider: true },
            { action: 'insertRowAbove', label: 'Insert row above', shortcut: 'Alt+I R' },
            { action: 'insertRowBelow', label: 'Insert row below', shortcut: '' },
            { action: 'insertColumnLeft', label: 'Insert column left', shortcut: 'Alt+I C' },
            { action: 'insertColumnRight', label: 'Insert column right', shortcut: '' },
            { isDivider: true },
            { action: 'deleteRow', label: 'Delete row', shortcut: '' },
            { action: 'deleteColumn', label: 'Delete column', shortcut: '' },
            { isDivider: true }
          ].map((item, index) => 
            item.isDivider ? (
              <div key={`divider-${index}`} style={{
                height: '1px',
                backgroundColor: '#ddd',
                margin: '4px 0'
              }}></div>
            ) : (
              <div 
                key={item.action}
                style={{
                  padding: '6px 24px 6px 16px',
                  cursor: 'pointer',
                  display: 'flex',
                  justifyContent: 'space-between',
                  transition: 'background-color 0.2s ease',
                  color: (item.action === 'paste' && !clipboard) ? '#aaa' : 'inherit'
                }}
                onClick={() => {
                  if (item.action === 'paste' && !clipboard) return; // Disable paste if nothing in clipboard
                  handleContextMenuItem(item.action);
                }}
                onMouseOver={(e) => {
                  if (item.action === 'paste' && !clipboard) return; // Don't highlight disabled items
                  e.currentTarget.style.backgroundColor = '#f1f3f4';
                }}
                onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
              >
                <span>{item.label}</span>
                {item.shortcut && (
                  <span style={{ color: '#5f6368', marginLeft: '16px', fontSize: '11px' }}>{item.shortcut}</span>
                )}
              </div>
            )
          )}
        </div>
      )}
      
      {/* Comment Popup */}
      {showCommentPopup && (
        <div
          style={{
            position: 'fixed',
            top: '50%',
            left: '50%',
            transform: 'translate(-50%, -50%)',
            backgroundColor: 'white',
            border: '1px solid #ddd',
            borderRadius: '8px',
            padding: '16px',
            boxShadow: '0 4px 16px rgba(0,0,0,0.2)',
            zIndex: 1100,
            width: '300px'
          }}
        >
          <div style={{ marginBottom: '12px', fontWeight: 'bold' }}>Add Comment</div>
          <textarea
            value={commentText}
            onChange={(e) => setCommentText(e.target.value)}
            style={{
              width: '100%',
              height: '100px',
              padding: '8px',
              border: '1px solid #ddd',
              borderRadius: '4px',
              resize: 'none',
              marginBottom: '16px'
            }}
            placeholder="Enter your comment here..."
          />
          <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
            <button
              style={{
                padding: '8px 16px',
                backgroundColor: 'transparent',
                border: '1px solid #ddd',
                borderRadius: '4px',
                marginRight: '8px',
                cursor: 'pointer'
              }}
              onClick={() => setShowCommentPopup(false)}
            >
              Cancel
            </button>
            <button
              style={{
                padding: '8px 16px',
                backgroundColor: '#1a73e8',
                color: 'white',
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer'
              }}
              onClick={saveCellComment}
            >
              Save
            </button>
          </div>
        </div>
      )}
      
      {/* Color Picker Popup */}
      {showColorPicker && (
        <div 
          ref={colorPickerRef}
          style={{
            position: 'absolute',
            top: colorPickerPosition.top,
            left: colorPickerPosition.left,
            backgroundColor: 'white',
            border: '1px solid #ddd',
            borderRadius: '4px',
            padding: '8px',
            boxShadow: '0 2px 10px rgba(0,0,0,0.2)',
            zIndex: 1000
          }}
        >
          <div style={{ marginBottom: '8px', fontWeight: '500' }}>
            {colorPickerType === 'text' ? 'Text Color' : 'Background Color'}
          </div>
          <div>
            {colorPalette.map((row, rowIdx) => (
              <div key={rowIdx} style={{ display: 'flex', marginBottom: '4px' }}>
                {row.map((color, colIdx) => (
                  <div 
                    key={colIdx}
                    style={{
                      width: '20px',
                      height: '20px',
                      backgroundColor: color,
                      margin: '2px',
                      border: '1px solid #ddd',
                      cursor: 'pointer',
                      transition: 'border 0.2s ease'
                    }}
                    onClick={() => applyColor(color)}
                    onMouseOver={(e) => e.currentTarget.style.border = '1px solid #1a73e8'}
                    onMouseOut={(e) => e.currentTarget.style.border = '1px solid #ddd'}
                  />
                ))}
              </div>
            ))}
          </div>
        </div>
      )}
      
      {/* Keyboard shortcuts help */}
      <div
        style={{
          position: 'absolute',
          bottom: '16px',
          right: '16px',
          backgroundColor: 'white',
          border: '1px solid #ddd',
          borderRadius: '50%',
          width: '32px',
          height: '32px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          cursor: 'pointer',
          boxShadow: '0 2px 5px rgba(0,0,0,0.1)',
          zIndex: 10
        }}
        onClick={() => menuActions.showHelp()}
        title="Keyboard shortcuts"
      >
        ?
      </div>
      /* 
This code snippet completes the chart rendering in App.js.
Add this right before the closing return statement (before the final closing div).
*/

{/* Charts Container */}
{charts.length > 0 && (
  <div className="charts-container">
    {charts.map(chart => (
      <div
        id={chart.id}
        key={chart.id}
        className={`chart-wrapper ${activeChart === chart.id ? 'active' : ''}`}
        style={{
          top: chart.position.top,
          left: chart.position.left,
          width: chart.position.width,
          height: chart.position.height,
          zIndex: activeChart === chart.id ? 50 : 10
        }}
        onClick={() => setActiveChart(chart.id)}
        onMouseDown={(e) => startChartDrag(chart.id, e)}
      >
        <div className="chart-header">
          <h3 className="chart-title">{chart.title || 'Chart'}</h3>
          <div className="chart-controls">
            <button
              className="chart-control-button"
              onClick={(e) => {
                e.stopPropagation();
                deleteChart(chart.id);
              }}
              title="Delete chart"
            >
              ×
            </button>
          </div>
        </div>
        {renderChart(chart)}
      </div>
    ))}
  </div>
)}

{/* Chart Dialog */}
{showChartDialog && (
  <ChartDialog
    onClose={() => setShowChartDialog(false)}
    onSave={closeChartDialog}
    cells={cells}
  />
)}
    </div>
  );
}

export default App;