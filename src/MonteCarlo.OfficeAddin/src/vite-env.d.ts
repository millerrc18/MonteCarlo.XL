/// <reference types="vite/client" />

declare module 'react-plotly.js';
declare module '*.css';

declare namespace CustomFunctions {
  enum ErrorCode {
    invalidValue = '#VALUE!',
  }

  class Error {
    constructor(code: ErrorCode, message?: string);
    code: ErrorCode;
    message: string;
  }

  function associate(id: string, implementation: (...args: any[]) => any): void;
}
