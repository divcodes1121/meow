// src/utils/formulaEvaluator.js
import * as math from 'mathjs';
import { mathFunctions, dataQualityFunctions } from './mathFunctions';
import { getCellsInRange, parseCellReference } from './cellHelpers';

// Global object to track cell dependencies
let cellDependencyGraph = {};

/**
 * Parse and evaluate a formula expression
 */
export const evaluateFormula = (formula, cellId, cellsData) => {
  if (!formula || typeof formula !== 'string' || !formula.startsWith('=')) {
    return formula;
  }
  
  // Remove the '=' sign
  const expression = formula.substring(1);
  
  try {
    // Handle built-in functions
    if (expression.includes('(') && expression.includes(')')) {
      const functionMatch = expression.match(/([A-Z_]+)\((.*)\)/);
      
      if (functionMatch) {
        const [fullMatch, funcName, args] = functionMatch;
        
        // Track dependencies for this cell
        updateDependencies(cellId, expression, cellsData);
        
        // Mathematical functions
        if (funcName in mathFunctions) {
          // Handle range arguments
          if (args.includes(':')) {
            const cellValues = getCellValuesFromRange(args, cellsData);
            return mathFunctions[funcName](cellValues);
          } else {
            // Handle comma-separated arguments
            const argValues = args.split(',').map(arg => {
              // Check if argument is a cell reference
              const cellRef = arg.trim();
              if (cellRef.match(/^[A-Z]+[0-9]+$/)) {
                return cellsData[cellRef]?.value || 0;
              }
              return arg.trim();
            });
            
            return mathFunctions[funcName](argValues);
          }
        } 
        // Data quality functions
        else if (funcName in dataQualityFunctions) {
          // Most data quality functions operate on a single value
          const argValues = args.split(',').map(arg => {
            const cellRef = arg.trim();
            if (cellRef.match(/^[A-Z]+[0-9]+$/)) {
              return cellsData[cellRef]?.value || '';
            }
            // Remove quotes if string literal
            if (cellRef.startsWith('"') && cellRef.endsWith('"')) {
              return cellRef.substring(1, cellRef.length - 1);
            }
            return arg.trim();
          });
          
          return dataQualityFunctions[funcName](...argValues);
        }
      }
    }
    
    // Handle basic expressions with cell references
    let evaluatedExpression = expression;
    
    // Replace cell references with their values
    const cellRefs = expression.match(/[A-Z]+[0-9]+/g) || [];
    cellRefs.forEach(ref => {
      const value = cellsData[ref]?.value || 0;
      evaluatedExpression = evaluatedExpression.replace(ref, value);
    });
    
    // Update dependencies
    updateDependencies(cellId, expression, cellsData);
    
    // Evaluate the expression
    return math.evaluate(evaluatedExpression);
  } catch (error) {
    console.error('Formula evaluation error:', error);
    return `#ERROR: ${error.message}`;
  }
};

/**
 * Get cell values from a range expression like "A1:B3"
 */
const getCellValuesFromRange = (rangeExpression, cellsData) => {
  const range = rangeExpression.trim();
  
  if (range.includes(':')) {
    const cells = getCellsInRange(range);
    return cells.map(cellId => cellsData[cellId]?.value || 0);
  }
  
  // Single cell reference
  if (range.match(/^[A-Z]+[0-9]+$/)) {
    return [cellsData[range]?.value || 0];
  }
  
  // Comma-separated values or cell references
  return range.split(',').map(item => {
    const cellRef = item.trim();
    if (cellRef.match(/^[A-Z]+[0-9]+$/)) {
      return cellsData[cellRef]?.value || 0;
    }
    return cellRef;
  });
};

/**
 * Update the dependency graph when a formula is evaluated
 */
export const updateDependencies = (cellId, expression, cellsData) => {
  // Extract all cell references from the expression
  const cellRefs = expression.match(/[A-Z]+[0-9]+/g) || [];
  
  // Also handle ranges
  const ranges = expression.match(/[A-Z]+[0-9]+:[A-Z]+[0-9]+/g) || [];
  let allDependencies = [...cellRefs];
  
  // Add cells from ranges
  ranges.forEach(range => {
    const cellsInRange = getCellsInRange(range);
    allDependencies = [...allDependencies, ...cellsInRange];
  });
  
  // Remove duplicates
  allDependencies = [...new Set(allDependencies)];
  
  // For each dependency, update the reverse dependency graph
  allDependencies.forEach(depCellId => {
    if (!cellDependencyGraph[depCellId]) {
      cellDependencyGraph[depCellId] = [];
    }
    
    // Add cellId as a dependent of depCellId if not already there
    if (!cellDependencyGraph[depCellId].includes(cellId)) {
      cellDependencyGraph[depCellId].push(cellId);
    }
  });
  
  return cellDependencyGraph;
};

/**
 * Update all cells that depend on a given cell
 */
export const updateDependentCells = (cellId, cells, dependencies = cellDependencyGraph) => {
  const dependents = dependencies[cellId] || [];
  
  dependents.forEach(depCellId => {
    if (cells[depCellId] && cells[depCellId].formula) {
      const formula = cells[depCellId].formula;
      cells[depCellId].value = evaluateFormula(formula, depCellId, cells);
      cells[depCellId].formatted = cells[depCellId].value;
      
      // Recursively update any cells that depend on this one
      updateDependentCells(depCellId, cells, dependencies);
    }
  });
};

/**
 * Get the dependency graph
 */
export const getDependencyGraph = () => {
  return cellDependencyGraph;
};

/**
 * Reset the dependency graph (useful for testing)
 */
export const resetDependencyGraph = () => {
  cellDependencyGraph = {};
};