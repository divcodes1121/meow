// src/store/index.js
import { configureStore } from '@reduxjs/toolkit';
import spreadsheetReducer from './spreadsheetSlice';

export const store = configureStore({
  reducer: {
    spreadsheet: spreadsheetReducer
  }
});

export default store;