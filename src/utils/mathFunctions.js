// src/utils/mathFunctions.js
// Mathematical functions
export const mathFunctions = {
    SUM: (values) => {
      if (!values || !values.length) return 0;
      return values.reduce((sum, val) => {
        const num = parseFloat(val);
        return sum + (isNaN(num) ? 0 : num);
      }, 0);
    },
    
    AVERAGE: (values) => {
      if (!values || !values.length) return 0;
      const sum = mathFunctions.SUM(values);
      const count = values.filter(val => !isNaN(parseFloat(val))).length;
      return count > 0 ? sum / count : 0;
    },
    
    MAX: (values) => {
      if (!values || !values.length) return 0;
      const nums = values
        .map(val => parseFloat(val))
        .filter(num => !isNaN(num));
      return nums.length > 0 ? Math.max(...nums) : 0;
    },
    
    MIN: (values) => {
      if (!values || !values.length) return 0;
      const nums = values
        .map(val => parseFloat(val))
        .filter(num => !isNaN(num));
      return nums.length > 0 ? Math.min(...nums) : 0;
    },
    
    COUNT: (values) => {
      if (!values || !values.length) return 0;
      return values.filter(val => !isNaN(parseFloat(val))).length;
    }
  };
  
  // Data quality functions
  export const dataQualityFunctions = {
    TRIM: (value) => {
      if (typeof value !== 'string') return value;
      return value.trim();
    },
    
    UPPER: (value) => {
      if (typeof value !== 'string') return value;
      return value.toUpperCase();
    },
    
    LOWER: (value) => {
      if (typeof value !== 'string') return value;
      return value.toLowerCase();
    },
    
    // This needs to be implemented differently as it affects multiple cells
    REMOVE_DUPLICATES: (values, cellRefs) => {
      // This is more of a placeholder - the actual implementation will be in the spreadsheet logic
      return "REMOVE_DUPLICATES is handled at the spreadsheet level";
    },
    
    // This also needs to be implemented differently
    FIND_AND_REPLACE: (value, find, replace) => {
      if (typeof value !== 'string') return value;
      return value.replace(new RegExp(find, 'g'), replace);
    }
  };