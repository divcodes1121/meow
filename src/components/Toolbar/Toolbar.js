// src/components/Toolbar/Toolbar.js
import React from 'react';
import { useSelector, useDispatch } from 'react-redux';
import { 
  setCellStyles, 
  addRow, 
  addColumn, 
  deleteRow, 
  deleteColumn 
} from '../../store/spreadsheetSlice';
import './Toolbar.scss';

const Toolbar = () => {
  const dispatch = useDispatch();
  
  // Get active cell and its styles from Redux store
  const activeCell = useSelector(state => state.spreadsheet.activeCell);
  const cells = useSelector(state => state.spreadsheet.cells);
  
  // Get active cell's row and column
  const getActiveCellRowCol = () => {
    if (!activeCell) return { row: null, col: null };
    
    const match = activeCell.match(/([A-Z]+)([0-9]+)/);
    if (match) {
      return { col: match[1], row: parseInt(match[2]) };
    }
    
    return { row: null, col: null };
  };
  
  const { row, col } = getActiveCellRowCol();
  
  // Get active cell's styles
  const getActiveStyles = () => {
    if (!activeCell || !cells[activeCell]) return {};
    return cells[activeCell].styles || {};
  };
  
  const activeStyles = getActiveStyles();
  
  // Handle styling changes
  const handleStyleChange = (style, value) => {
    if (!activeCell) return;
    
    dispatch(setCellStyles({
      cellId: activeCell,
      styles: { [style]: value }
    }));
  };
  
  // Handle adding a row
  const handleAddRow = () => {
    if (row) {
      dispatch(addRow({ afterRow: row }));
    }
  };
  
  // Handle adding a column
  const handleAddColumn = () => {
    if (col) {
      dispatch(addColumn({ afterCol: col }));
    }
  };
  
  // Handle deleting a row
  const handleDeleteRow = () => {
    if (row) {
      dispatch(deleteRow({ rowNum: row }));
    }
  };
  
  // Handle deleting a column
  const handleDeleteColumn = () => {
    if (col) {
      dispatch(deleteColumn({ col }));
    }
  };
  
  return (
    <div className="toolbar">
      <div className="toolbar-section">
        <button
          className={`toolbar-button ${activeStyles.fontWeight === 'bold' ? 'active' : ''}`}
          onClick={() => handleStyleChange('fontWeight', activeStyles.fontWeight === 'bold' ? 'normal' : 'bold')}
          title="Bold"
        >
          B
        </button>
        
        <button
          className={`toolbar-button ${activeStyles.fontStyle === 'italic' ? 'active' : ''}`}
          onClick={() => handleStyleChange('fontStyle', activeStyles.fontStyle === 'italic' ? 'normal' : 'italic')}
          title="Italic"
        >
          I
        </button>
        
        <select
          className="toolbar-select"
          value={activeStyles.fontSize || '14px'}
          onChange={(e) => handleStyleChange('fontSize', e.target.value)}
          title="Font Size"
        >
          <option value="10px">10</option>
          <option value="12px">12</option>
          <option value="14px">14</option>
          <option value="16px">16</option>
          <option value="18px">18</option>
          <option value="20px">20</option>
          <option value="24px">24</option>
        </select>
        
        <input
          type="color"
          className="toolbar-color"
          value={activeStyles.color || '#000000'}
          onChange={(e) => handleStyleChange('color', e.target.value)}
          title="Text Color"
        />
      </div>
      
      <div className="toolbar-section">
        <button
          className="toolbar-button"
          onClick={handleAddRow}
          title="Add Row"
        >
          + Row
        </button>
        
        <button
          className="toolbar-button"
          onClick={handleAddColumn}
          title="Add Column"
        >
          + Column
        </button>
        
        <button
          className="toolbar-button"
          onClick={handleDeleteRow}
          title="Delete Row"
        >
          - Row
        </button>
        
        <button
          className="toolbar-button"
          onClick={handleDeleteColumn}
          title="Delete Column"
        >
          - Column
        </button>
      </div>
    </div>
  );
};

export default Toolbar;