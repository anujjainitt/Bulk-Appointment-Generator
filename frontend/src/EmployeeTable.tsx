import React from 'react';

interface EmployeeTableProps {
  columns: string[];
  rows: any[][];
  colWidths: number[];
  setColWidths: (w: number[]) => void;
  page: number;
  setPage: (p: number) => void;
  rowsPerPage: number;
  setRowsPerPage: (n: number) => void;
  search: string;
  setSearch: (s: string) => void;
  selectedRows: number[];
  setSelectedRows: (rows: number[]) => void;
  handleSelectAll: (e: React.ChangeEvent<HTMLInputElement>) => void;
  handleSelectRow: (idx: number) => void;
}

export const EmployeeTable: React.FC<EmployeeTableProps> = () => {
  return null;
}; 