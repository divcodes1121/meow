// src/components/FormulaBar/FormulaBar.js
import React, { useState, useEffect } from 'react';
import { useSelector, useDispatch } from 'react-redux';
import { setCellContent, setEditMode } from '../../store/spreadsheetSlice';
import './FormulaBar.scss';

const FormulaBar = () => {
  const dispatch = useDispatch();
  
  // Get active cell and its formula from Redux store
  const activeCell = useSelector(state => state.spreadsheet.activeCell);
  const cells = useSelector(state => state.spreadsheet.cells);
  const isEditing = useSelector(state => state.spreadsheet.isEditing);
  
  // Local state for the formula input
  const [formula, setFormula] = useState('');
  
  // Update the formula input when the active cell changes
  useEffect(() => {
    if (activeCell && cells[activeCell]) {
      setFormula(cells[activeCell].formula || '');
    } else {
      setFormula('');
    }
  }, [activeCell, cells]);
  
  // Handle formula change
  const handleFormulaChange = (e) => {
    setFormula(e.target.value);
  };
  
  // Handle formula submission
  const handleFormulaSubmit = () => {
    if (activeCell) {
      dispatch(setCellContent({ cellId: activeCell, content: formula }));
    }
  };
  
  // Handle key press events
  const handleKeyDown = (e) => {
    if (e.key === 'Enter') {
      handleFormulaSubmit();
      dispatch(setEditMode(false));
    } else if (e.key === 'Escape') {
      if (activeCell && cells[activeCell]) {
        setFormula(cells[activeCell].formula || '');
      }
      dispatch(setEditMode(false));
    }
  };
  
  return (
    <div className="formula-bar">
      <div className="cell-reference">
        {activeCell || ''}
      </div>
      <div className="formula-input-container">
        <div className="formula-prefix">=</div>
        <input
          type="text"
          className="formula-input"
          value={formula}
          onChange={handleFormulaChange}
          onKeyDown={handleKeyDown}
          onBlur={handleFormulaSubmit}
          onFocus={() => dispatch(setEditMode(true))}
          placeholder="Enter formula or value"
        />
      </div>
    </div>
  );
};

export default FormulaBar;