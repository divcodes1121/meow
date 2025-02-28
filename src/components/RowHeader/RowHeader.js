// src/components/RowHeader/RowHeader.js
import React, { useRef } from 'react';
import { useSelector, useDispatch } from 'react-redux';
import { resizeRow } from '../../store/spreadsheetSlice';
import './RowHeader.scss';

const RowHeader = ({ row }) => {
  const dispatch = useDispatch();
  const headerRef = useRef(null);
  const resizeRef = useRef(null);
  
  // Get row height from Redux store
  const rowHeight = useSelector(state => state.spreadsheet.rows[row]?.height || 25);
  
  // Handle row resize
  const handleResizeMouseDown = (e) => {
    e.preventDefault();
    
    const startY = e.clientY;
    const startHeight = rowHeight;
    
    const handleMouseMove = (moveEvent) => {
      const delta = moveEvent.clientY - startY;
      const newHeight = Math.max(25, startHeight + delta);
      dispatch(resizeRow({ rowNum: row, height: newHeight }));
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
      className="row-header"
      ref={headerRef}
      style={{ height: rowHeight }}
    >
      {row}
      <div
        className="resize-handle"
        ref={resizeRef}
        onMouseDown={handleResizeMouseDown}
      />
    </div>
  );
};

export default RowHeader;