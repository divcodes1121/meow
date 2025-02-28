import React from 'react';
import { useSelector } from 'react-redux';

const Grid = () => {
  // Get data from Redux store
  const cells = useSelector(state => state.spreadsheet.cells || {});
  const columns = useSelector(state => state.spreadsheet.columns || {});
  const rows = useSelector(state => state.spreadsheet.rows || {});

  const columnKeys = Object.keys(columns);
  const rowKeys = Object.keys(rows);

  return (
    <div className="grid-container">
      <div className="corner-header"></div>
      <div className="column-headers">
        {columnKeys.map(col => (
          <div 
            key={col}
            className="column-header" 
            style={{ width: columns[col].width }}
          >
            {col}
          </div>
        ))}
      </div>
      <div className="grid-body">
        <div className="row-headers">
          {rowKeys.map(row => (
            <div 
              key={row}
              className="row-header" 
              style={{ height: rows[row].height }}
            >
              {row}
            </div>
          ))}
        </div>
        <div className="cells-container">
          {rowKeys.map(row => (
            <div 
              key={row}
              className="row" 
              style={{ height: rows[row].height }}
            >
              {columnKeys.map(col => {
                const cellId = `${col}${row}`;
                const cell = cells[cellId] || { value: '', formatted: '', styles: {} };
                return (
                  <div 
                    key={cellId}
                    className="cell-wrapper" 
                    style={{ width: columns[col].width }}
                    data-cell-id={cellId}
                  >
                    <div 
                      className="cell"
                      style={cell.styles || {}}
                    >
                      {cell.formatted || ''}
                    </div>
                  </div>
                );
              })}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

export default Grid;