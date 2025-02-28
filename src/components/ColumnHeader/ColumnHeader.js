// src/components/ColumnHeader/ColumnHeader.js
import React, { useRef } from 'react';
import { useSelector, useDispatch } from 'react-redux';
import { resizeColumn } from '../../store/spreadsheetSlice';
import './ColumnHeader.scss';

const ColumnHeader = ({ col }) => {
  const dispatch = useDispatch();
  const headerRef = useRef(null);
  const resizeRef = useRef(null);
  
  // Get column width from Redux store
  const columnWidth = useSelector(state => state.spreadsheet.columns[col]?.width || 100);
  
  // Handle column resize
  const handleResizeMouseDown = (e) => {
    e.preventDefault();
    
    const startX = e.clientX;
    const startWidth = columnWidth;
    
    const handleMouseMove = (moveEvent) => {
      const delta = moveEvent.clientX - startX;
      const newWidth = Math.max(50, startWidth + delta);
      dispatch(resizeColumn({ col, width: newWidth }));
    };
    
    const handleMouseUp = () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
    
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };
  
  return (
    <div
      className="column-header"
      ref={headerRef}
      style={{ width: columnWidth }}
    >
      {col}
      <div
        className="resize-handle"
        ref={resizeRef}
        onMouseDown={handleResizeMouseDown}
      />
    </div>
  );
};

export default ColumnHeader;