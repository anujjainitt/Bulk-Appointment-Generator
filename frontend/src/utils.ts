// Utility functions for the frontend
import * as XLSX from 'xlsx';
import { format, parseISO, isValid } from 'date-fns';

export function formatDateUI(date: any) {
  if (!date) return '';
  if (typeof date === 'number') {
    const parsed = XLSX.SSF ? XLSX.SSF.parse_date_code(date) : null;
    if (parsed) {
      const jsDate = new Date(parsed.y, parsed.m - 1, parsed.d);
      return format(jsDate, 'd MMMM yyyy');
    }
  }
  let d = typeof date === 'string' ? parseISO(date) : new Date(date);
  if (!isValid(d)) d = new Date(date);
  if (!isValid(d)) return date;
  return format(d, 'd MMMM yyyy');
} 