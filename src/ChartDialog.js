import React, { useState, useEffect } from 'react';
import './App.css';

const ChartDialog = ({ onClose, onSave, cells }) => {
  // State for chart properties
  const [chartType, setChartType] = useState('line');
  const [chartTitle, setChartTitle] = useState('');
  const [dataRange, setDataRange] = useState('');
  const [headerRow, setHeaderRow] = useState(true);
  const [stacked, setStacked] = useState(false);
  const [xAxisColumn, setXAxisColumn] = useState('');
  const [yAxisColumns, setYAxisColumns] = useState([]);
  
  // Preview data
  const [previewData, setPreviewData] = useState([]);
  const [headers, setHeaders] = useState([]);
  
  // Handle data range input change
  const handleRangeChange = (e) => {
    setDataRange(e.target.value);
  };
  
  // Extract range into cell references
  const parseRange = (range) => {
    if (!range.includes(':')) return [];
    
    const [start, end] = range.split(':');
    
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
      const rowData = [];
      for (let colIdx = Math.min(startColIdx, endColIdx); colIdx <= Math.max(startColIdx, endColIdx); colIdx++) {
        const colLetter = indexToColumn(colIdx);
        const cellId = `${colLetter}${row}`;
        rowData.push(cellId);
      }
      cellsList.push(rowData);
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
  
  // Preview the data when range changes
  useEffect(() => {
    if (!dataRange) {
      setPreviewData([]);
      setHeaders([]);
      return;
    }
    
    const cellGrid = parseRange(dataRange);
    if (cellGrid.length === 0) return;
    
    // Extract headers if needed
    let dataStartIdx = 0;
    let extractedHeaders = [];
    
    if (headerRow && cellGrid.length > 0) {
      extractedHeaders = cellGrid[0].map(cellId => cells[cellId]?.value || '');
      dataStartIdx = 1;
    } else {
      // Generate default headers (Column A, Column B, etc.)
      extractedHeaders = cellGrid[0].map((_, idx) => `Column ${indexToColumn(idx)}`);
    }
    
    setHeaders(extractedHeaders);
    
    // Extract data
    const extractedData = [];
    for (let i = dataStartIdx; i < cellGrid.length; i++) {
      const rowObj = {};
      
      cellGrid[i].forEach((cellId, colIdx) => {
        const header = extractedHeaders[colIdx] || `Column ${indexToColumn(colIdx)}`;
        const cellValue = cells[cellId]?.value || '';
        
        // Try to convert to number if possible
        rowObj[header] = !isNaN(parseFloat(cellValue)) ? parseFloat(cellValue) : cellValue;
      });
      
      extractedData.push(rowObj);
    }
    
    setPreviewData(extractedData);
    
    // Set default X and Y axes based on available data
    if (extractedHeaders.length > 0 && extractedData.length > 0) {
      if (!xAxisColumn) {
        setXAxisColumn(extractedHeaders[0]);
      }
      
      // Set Y-axis columns (all except X-axis column)
      if (yAxisColumns.length === 0) {
        const potentialYColumns = extractedHeaders.filter(header => header !== xAxisColumn);
        setYAxisColumns(potentialYColumns.length > 0 ? [potentialYColumns[0]] : []);
      }
    }
  }, [dataRange, headerRow, cells]);
  
  // Prepare chart data and save
  const handleSave = () => {
    if (!xAxisColumn || yAxisColumns.length === 0 || previewData.length === 0) {
      alert('Please select valid data range and axes for the chart');
      return;
    }
    
    const chartData = {
      type: chartType,
      title: chartTitle,
      data: previewData,
      xAxisKey: xAxisColumn,
      yAxisKeys: yAxisColumns,
      stacked: stacked
    };
    
    onSave(chartData);
  };
  
  // Toggle Y-axis column selection
  const toggleYAxisColumn = (column) => {
    if (yAxisColumns.includes(column)) {
      setYAxisColumns(yAxisColumns.filter(col => col !== column));
    } else {
      setYAxisColumns([...yAxisColumns, column]);
    }
  };
  
  return (
    <div className="chart-dialog-overlay" onClick={(e) => {
      // Close the dialog when clicking on the overlay (outside the dialog)
      if (e.target.className === 'chart-dialog-overlay') {
        onClose();
      }
    }}>
      <div className="chart-dialog">
        <div className="chart-dialog-header">
          <h2>Create Chart</h2>
          <button className="close-button" onClick={onClose}>×</button>
        </div>
        
        <div className="chart-dialog-content">
          <div className="form-group">
            <label htmlFor="chart-title">Chart Title</label>
            <input
              id="chart-title"
              type="text"
              value={chartTitle}
              onChange={(e) => setChartTitle(e.target.value)}
              placeholder="Enter chart title"
            />
          </div>
          
          <div className="form-group">
            <label htmlFor="chart-type">Chart Type</label>
            <select
              id="chart-type"
              value={chartType}
              onChange={(e) => setChartType(e.target.value)}
            >
              <option value="line">Line Chart</option>
              <option value="bar">Bar Chart</option>
              <option value="pie">Pie Chart</option>
              <option value="area">Area Chart</option>
            </select>
          </div>
          
          <div className="form-group">
            <label htmlFor="data-range">Data Range (e.g., A1:C10)</label>
            <input
              id="data-range"
              type="text"
              value={dataRange}
              onChange={handleRangeChange}
              placeholder="Enter cell range (e.g., A1:C10)"
            />
          </div>
          
          <div className="form-group">
            <label>
              <input
                type="checkbox"
                checked={headerRow}
                onChange={(e) => setHeaderRow(e.target.checked)}
              />
              First row contains headers
            </label>
          </div>
          
          {(chartType === 'bar' || chartType === 'area') && (
            <div className="form-group">
              <label>
                <input
                  type="checkbox"
                  checked={stacked}
                  onChange={(e) => setStacked(e.target.checked)}
                />
                Stacked
              </label>
            </div>
          )}
          
          {headers.length > 0 && (
            <>
              <div className="form-group">
                <label htmlFor="x-axis">X-Axis</label>
                <select
                  id="x-axis"
                  value={xAxisColumn}
                  onChange={(e) => setXAxisColumn(e.target.value)}
                >
                  {headers.map((header, idx) => (
                    <option key={idx} value={header}>{header}</option>
                  ))}
                </select>
              </div>
              
              <div className="form-group">
                <label>Y-Axis (Select one or more)</label>
                <div className="checkbox-group">
                  {headers
                    .filter(header => header !== xAxisColumn)
                    .map((header, idx) => (
                      <label key={idx}>
                        <input
                          type="checkbox"
                          checked={yAxisColumns.includes(header)}
                          onChange={() => toggleYAxisColumn(header)}
                        />
                        {header}
                      </label>
                    ))}
                </div>
              </div>
            </>
          )}
          
          {previewData.length > 0 && (
            <div className="data-preview">
              <h3>Data Preview</h3>
              <div className="table-container">
                <table>
                  <thead>
                    <tr>
                      {headers.map((header, idx) => (
                        <th key={idx}>{header}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {previewData.slice(0, 5).map((row, rowIdx) => (
                      <tr key={rowIdx}>
                        {headers.map((header, colIdx) => (
                          <td key={colIdx}>{row[header]}</td>
                        ))}
                      </tr>
                    ))}
                    {previewData.length > 5 && (
                      <tr>
                        <td colSpan={headers.length} style={{ textAlign: 'center' }}>
                          ... {previewData.length - 5} more rows
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
        
        <div className="chart-dialog-footer">
          <button onClick={onClose}>Cancel</button>
          <button 
            onClick={handleSave}
            disabled={!xAxisColumn || yAxisColumns.length === 0 || previewData.length === 0}
            className="primary-button"
          >
            Create Chart
          </button>
        </div>
      </div>
    </div>
  );
};

export default ChartDialog;