// src/components/Cell/Cell.js
import React, { useState, useEffect, useRef } from 'react';
import { useDispatch, useSelector } from 'react-redux';
import { 
  setActiveCell, 
  setCellContent, 
  setCellStyles, 
  setEditMode 
} from '../../store/spreadsheetSlice';
import './Cell.scss';

const Cell = ({ cellId }) => {
  const dispatch = useDispatch();
  const cellRef = useRef(null);
  const inputRef = useRef(null);
  
  // Get cell data from Redux store
  const cell = useSelector(state => state.spreadsheet.cells[cellId]);
  const activeCell = useSelector(state => state.spreadsheet.activeCell);
  const isEditing = useSelector(state => state.spreadsheet.isEditing);
  const selectedRange = useSelector(state => state.spreadsheet.selectedRange);
  
  // Local state for editing
  const [inputValue, setInputValue] = useState('');
  
  // Check if this cell is the active cell
  const isActive = activeCell === cellId;
  
  // Check if this cell is in the selected range
  const isInRange = selectedRange ? isInSelectedRange(cellId, selectedRange) : false;
  
  // When the active cell changes and this cell becomes active, focus it
  useEffect(() => {
    if (isActive && isEditing && inputRef.current) {
      inputRef.current.focus();
    }
  }, [isActive, isEditing]);
  
  // When cell data changes, update the input value if this is the active cell
  useEffect(() => {
    if (isActive && cell) {
      setInputValue(cell.formula || '');
    }
  }, [isActive, cell]);
  
  // Handle cell click
  const handleCellClick = (e) => {
    if (!isActive) {
      dispatch(setActiveCell(cellId));
      
      // If double click, enter edit mode
      if (e.detail === 2) {
        dispatch(setEditMode(true));
        setInputValue(cell.formula || '');
      } else {
        dispatch(setEditMode(false));
      }
    }
  };
  
  // Handle key presses
  const handleKeyDown = (e) => {
    if (e.key === 'Enter') {
      // Save changes and exit edit mode
      dispatch(setCellContent({ cellId, content: inputValue }));
      dispatch(setEditMode(false));
    } else if (e.key === 'Escape') {
      // Cancel edit mode
      dispatch(setEditMode(false));
      setInputValue(cell.formula || '');
    }
  };
  
  // Render the cell
  if (isActive && isEditing) {
    return (
      <div
        className={`cell active ${isInRange ? 'in-range' : ''}`}
        ref={cellRef}
        style={cell.styles || {}}
      >
        <input
          ref={inputRef}
          type="text"
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyDown={handleKeyDown}
          onBlur={() => {
            dispatch(setCellContent({ cellId, content: inputValue }));
            dispatch(setEditMode(false));
          }}
        />
      </div>
    );
  }
  
  return (
    <div
      className={`cell ${isActive ? 'active' : ''} ${isInRange ? 'in-range' : ''}`}
      onClick={handleCellClick}
      ref={cellRef}
      style={cell.styles || {}}
    >
      {cell.formatted}
    </div>
  );
};

// Helper to check if a cell is in the selected range
const isInSelectedRange = (cellId, range) => {
  if (!range || !range.start || !range.end) return false;
  
  // Parse cell IDs
  const [colLetter, rowStr] = cellId.match(/([A-Z]+)([0-9]+)/).slice(1);
  const row = parseInt(rowStr);
  const col = colLetter.charCodeAt(0);
  
  // Parse range
  const [startColLetter, startRowStr] = range.start.match(/([A-Z]+)([0-9]+)/).slice(1);
  const [endColLetter, endRowStr] = range.end.match(/([A-Z]+)([0-9]+)/).slice(1);
  
  const startRow = parseInt(startRowStr);
  const startCol = startColLetter.charCodeAt(0);
  const endRow = parseInt(endRowStr);
  const endCol = endColLetter.charCodeAt(0);
  
  // Check if cell is in range
  return (
    row >= Math.min(startRow, endRow) &&
    row <= Math.max(startRow, endRow) &&
    col >= Math.min(startCol, endCol) &&
    col <= Math.max(startCol, endCol)
  );
};

export default Cell;