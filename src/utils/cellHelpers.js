// src/utils/cellHelpers.js
/**
 * Convert column letter to index (0-based)
 * e.g., A -> 0, B -> 1, Z -> 25, AA -> 26
 */
export const colLetterToIndex = (colLetter) => {
    let result = 0;
    for (let i = 0; i < colLetter.length; i++) {
      result = result * 26 + (colLetter.charCodeAt(i) - 64);
    }
    return result - 1;
  };
  
  /**
   * Convert column index to letter
   * e.g., 0 -> A, 1 -> B, 25 -> Z, 26 -> AA
   */
  export const colIndexToLetter = (index) => {
    let temp = index + 1;
    let result = '';
    
    while (temp > 0) {
      const remainder = (temp - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      temp = Math.floor((temp - remainder) / 26);
    }
    
    return result;
  };
  
  /**
   * Parse a cell reference like "A1" into { col: "A", row: 1 }
   */
  export const parseCellReference = (cellRef) => {
    const match = cellRef.match(/([A-Z]+)([0-9]+)/);
    if (match) {
      return { col: match[1], row: parseInt(match[2]) };
    }
    return null;
  };
  
  /**
   * Get all cells in a range like "A1:B3"
   * Returns an array of cell IDs: ["A1", "A2", "A3", "B1", "B2", "B3"]
   */
  export const getCellsInRange = (range) => {
    const [start, end] = range.split(':');
    const startRef = parseCellReference(start);
    const endRef = parseCellReference(end);
    
    if (!startRef || !endRef) return [];
    
    const startColIndex = colLetterToIndex(startRef.col);
    const endColIndex = colLetterToIndex(endRef.col);
    const startRow = startRef.row;
    const endRow = endRef.row;
    
    const cells = [];
    
    for (let colIdx = startColIndex; colIdx <= endColIndex; colIdx++) {
      const col = colIndexToLetter(colIdx);
      for (let row = startRow; row <= endRow; row++) {
        cells.push(`${col}${row}`);
      }
    }
    
    return cells;
  };