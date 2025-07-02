/// <reference types="react-scripts" />
// @ts-ignore: No type definitions for 'react-resizable'
import { useState, useRef, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'
import {
  Box,
  Button,
  Container,
  Typography,
  Paper,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  LinearProgress,
  Alert,
  Stack,
  IconButton,
  useMediaQuery,
  useTheme,
  Drawer,
  List,
  ListItem,
  ListItemButton,
  ListItemIcon,
  ListItemText,
  Divider,
  AppBar,
  Toolbar,
} from '@mui/material'
import CloudUploadIcon from '@mui/icons-material/CloudUpload'
import DownloadIcon from '@mui/icons-material/Download'
import InsertDriveFileIcon from '@mui/icons-material/InsertDriveFile'
import DescriptionIcon from '@mui/icons-material/Description'
import { format, parseISO, isValid } from 'date-fns'
import Tooltip from '@mui/material/Tooltip'
import './App.css'
import { ThemeProvider, createTheme } from '@mui/material/styles'
import { motion } from 'framer-motion'
import Pagination from '@mui/material/Pagination'
import Select from '@mui/material/Select'
import MenuItem from '@mui/material/MenuItem'
import TextField from '@mui/material/TextField'
import { ResizableBox } from 'react-resizable'
import 'react-resizable/css/styles.css'
import Checkbox from '@mui/material/Checkbox'

const REQUIRED_COLUMNS = [
  'Date of Joining',
  'Name',
  'Email',
  'Contact',
  'Designation',
  'Place of Joining',
  'Address',
  'HR Name',
  'HR Designation',
  'Effective Date',
]

// Helper to format date for UI
function formatDateUI(date: any) {
  if (!date) return '';
  // Handle Excel date serials (numbers)
  if (typeof date === 'number') {
    // Excel's day 0 is 1899-12-30
    const parsed = XLSX.SSF ? XLSX.SSF.parse_date_code(date) : null;
    if (parsed) {
      const jsDate = new Date(parsed.y, parsed.m - 1, parsed.d);
      return format(jsDate, 'd MMMM yyyy');
    }
  }
  // Try to parse as ISO or fallback to Date
  let d = typeof date === 'string' ? parseISO(date) : new Date(date);
  if (!isValid(d)) d = new Date(date);
  if (!isValid(d)) return date;
  return format(d, 'd MMMM yyyy');
}

const themeOptions = {
  palette: {
    primary: {
      main: '#1976d2', // Simple blue
      contrastText: '#fff',
    },
    secondary: {
      main: '#424242', // Neutral gray
      contrastText: '#fff',
    },
    background: {
      default: '#f5f6fa', // Light gray
      paper: '#fff',
      sidebar: '#f5f6fa',
    },
    text: {
      primary: '#212b36',
      secondary: '#637381',
    },
  },
  shape: {
    borderRadius: 8,
  },
  components: {
    MuiButton: {
      styleOverrides: {
        root: {
          textTransform: 'none',
          fontWeight: 600,
          letterSpacing: 1,
          boxShadow: 'none',
          paddingLeft: 20,
          paddingRight: 20,
          paddingTop: 8,
          paddingBottom: 8,
        } as any,
        containedPrimary: {
          background: '#1976d2',
          color: '#fff',
        } as any,
        containedSecondary: {
          background: '#424242',
          color: '#fff',
        } as any,
        outlinedPrimary: {
          borderColor: '#1976d2',
          color: '#1976d2',
          background: '#f5f6fa',
        } as any,
        outlinedSecondary: {
          borderColor: '#424242',
          color: '#424242',
          background: '#fff',
        } as any,
      },
    },
    MuiPaper: {
      styleOverrides: {
        root: {
          background: '#fff',
        } as any,
      },
    },
    MuiDrawer: {
      styleOverrides: {
        paper: {
          background: '#f5f6fa',
          color: '#212b36',
        } as any,
      },
    },
    MuiTableHead: {
      styleOverrides: {
        root: {
          background: '#f5f6fa',
        } as any,
      },
    },
    MuiTableCell: {
      styleOverrides: {
        head: {
          color: '#1976d2',
          fontWeight: 700,
        } as any,
      },
    },
  },
}
const modernTheme = createTheme(themeOptions as any)

function App() {
  const [excelFile, setExcelFile] = useState<File | null>(null)
  const [excelData, setExcelData] = useState<any[][]>([])
  const [loading, setLoading] = useState(false)
  const [downloaded, setDownloaded] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const theme = useTheme()
  const isMobile = useMediaQuery(theme.breakpoints.down('sm'))
  const [selectedPage, setSelectedPage] = useState('bulk')
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const [page, setPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(10);
  const [search, setSearch] = useState('');
  // Column widths state for resizable columns
  const defaultColWidth = 180;
  const [colWidths, setColWidths] = useState(() =>
    excelData[0]?.map(() => defaultColWidth) || []
  );
  // Update colWidths if excelData changes (e.g., new file upload)
  useEffect(() => {
    if (excelData[0]) {
      setColWidths(excelData[0].map(() => defaultColWidth));
    }
  }, [excelData]);
  // Handler for resizing columns
  const handleResize = (index: number, newWidth: number) => {
    setColWidths((widths) => {
      const updated = [...widths];
      updated[index] = Math.max(newWidth, 80); // min width
      return updated;
    });
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    setExcelFile(file || null)
    setDownloaded(false)
    setError(null)
    if (file) {
      setLoading(true)
      const reader = new FileReader()
      reader.onload = (evt) => {
        const bstr = evt.target?.result as string
        const wb = XLSX.read(bstr, { type: 'binary' })
        const wsname = wb.SheetNames[0]
        const ws = wb.Sheets[wsname]
        const data = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 })
        if (data.length === 0) {
          setExcelData([])
          setLoading(false)
          return
        }
        const header = data[0] as string[]
        const colIndexes = REQUIRED_COLUMNS.map(col => header.indexOf(col))
        const filteredData = [
          REQUIRED_COLUMNS,
          ...data.slice(1).map((row: any) => colIndexes.map(idx => (idx !== -1 ? row[idx] : '')))
        ]
        setExcelData(filteredData)
        setLoading(false)
      }
      reader.readAsBinaryString(file)
    }
    // Reset the input so the same file can be uploaded again
    if (fileInputRef.current) fileInputRef.current.value = ''
  }

  const handleDownloadSample = () => {
    const sampleData = [
      REQUIRED_COLUMNS,
      ['2024-07-01', 'John Doe', 'john@example.com', '1234567890', 'Developer', 'Bangalore', '123 Main St', 'Priya HR', 'HR Manager', '2024-07-15'],
      ['2024-07-15', 'Jane Smith', 'jane@example.com', '9876543210', 'Manager', 'Delhi', '456 Park Ave', 'Amit HR', 'Senior HR', '2024-08-01'],
    ]
    const ws = XLSX.utils.aoa_to_sheet(sampleData)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Sample')
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
    const blob = new Blob([wbout], { type: 'application/octet-stream' })
    saveAs(blob, 'sample_format.xlsx')
  }

  // Row selection state
  const [selectedRows, setSelectedRows] = useState<number[]>([]);
  // Handle select all
  const handleSelectAll = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.checked) {
      setSelectedRows(filteredRows.map((_, idx) => idx));
    } else {
      setSelectedRows([]);
    }
  };
  // Handle select one
  const handleSelectRow = (idx: number) => {
    setSelectedRows((prev) =>
      prev.includes(idx) ? prev.filter(i => i !== idx) : [...prev, idx]
    );
  };

  const handleGenerateDoc = async () => {
    if (!excelFile && selectedRows.length === 0) return;
    setLoading(true);
    setDownloaded(false);
    setError(null);
    try {
      let response;
      if (selectedRows.length > 0) {
        // Prepare selected data as JSON
        const selectedData = [excelData[0], ...selectedRows.map(idx => filteredRows[idx])];
        response = await fetch('http://localhost:5000/upload-excel', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ data: selectedData }),
        });
      } else {
        // Fallback to file upload (all rows)
        const formData = new FormData();
        if (excelFile) {
          formData.append('file', excelFile);
        }
        response = await fetch('http://localhost:5000/upload-excel', {
          method: 'POST',
          body: formData,
        });
      }
      if (!response.ok) throw new Error('Failed to generate document');
      const blob = await response.blob();
      const disposition = response.headers.get('Content-Disposition');
      let filename = 'appointment_letters.zip';
      if (disposition && disposition.indexOf('filename=') !== -1) {
        filename = disposition.split('filename=')[1].replace(/['"]/g, '');
      }
      saveAs(blob, filename);
      setDownloaded(true);
    } catch (err) {
      setError('Error generating document. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const handleChangePage = (event: React.ChangeEvent<unknown>, value: number) => {
    setPage(value);
  };

  const handleChangeRowsPerPage = (event: React.ChangeEvent<{ value: unknown }>) => {
    setRowsPerPage(Number(event.target.value));
    setPage(1);
  };

  // Find column indexes for search fields
  const nameIdx = excelData[0]?.indexOf('Name');
  const contactIdx = excelData[0]?.indexOf('Contact');
  const emailIdx = excelData[0]?.indexOf('Email');
  // Filtered data based on search
  const filteredRows = excelData.length > 1 ? excelData.slice(1).filter((row: any[]) => {
    if (!search) return true;
    const name = row[nameIdx]?.toString().toLowerCase() || '';
    const contact = row[contactIdx]?.toString().toLowerCase() || '';
    const email = row[emailIdx]?.toString().toLowerCase() || '';
    return (
      name.includes(search.toLowerCase()) ||
      contact.includes(search.toLowerCase()) ||
      email.includes(search.toLowerCase())
    );
  }) : [];

  return (
    <ThemeProvider theme={modernTheme}>
      <Box sx={{ display: 'flex', minHeight: '100vh', minWidth: 0, width: '100%', overflowX: 'hidden', overflowY: 'hidden', background: '#f5f6fa', transition: 'background 0.5s' }}>
        {/* Header */}
        <AppBar position="fixed" elevation={4} sx={{ zIndex: (theme) => theme.zIndex.drawer + 1, background: 'linear-gradient(90deg, #1976d2 60%, #f50057 100%)', color: '#fff', boxShadow: '0 4px 24px 0 rgba(25, 118, 210, 0.15)' }}>
          <Toolbar sx={{ justifyContent: 'space-between', px: 4 }}>
            <Typography variant="h5" component="div" sx={{ fontWeight: 900, letterSpacing: 2, color: '#fff' }}>
              InTimeTec
            </Typography>
            <Typography variant="subtitle1" sx={{ fontWeight: 500, color: '#fff', opacity: 0.85 }}>
              Bulk Appointment Generator
            </Typography>
          </Toolbar>
        </AppBar>
        {/* Side Navigation Bar */}
        <Drawer
          variant={isMobile ? 'temporary' : 'permanent'}
          open={true}
          sx={{
            width: 250,
            flexShrink: 0,
            '& .MuiDrawer-paper': {
              width: 250,
              boxSizing: 'border-box',
              borderRight: 'none',
              background: 'linear-gradient(135deg, #212b36 0%, #1976d2 100%)',
              color: '#fff',
              boxShadow: '2px 0 16px 0 rgba(25, 118, 210, 0.10)',
              height: '100vh',
            },
          }}
          ModalProps={{ keepMounted: true }}
        >
          <Box sx={{ height: 72, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 900, fontSize: 22, letterSpacing: 1, color: '#fff', background: 'rgba(255,255,255,0.05)' }}>
            Menu
          </Box>
          <Divider sx={{ borderColor: 'rgba(255,255,255,0.15)' }} />
          <List>
            <ListItem disablePadding>
              <ListItemButton selected={selectedPage === 'bulk'} onClick={() => setSelectedPage('bulk')} sx={{ color: '#fff', '&.Mui-selected': { background: 'rgba(255,255,255,0.10)' } }}>
                <ListItemIcon sx={{ color: '#fff' }}>
                  <DescriptionIcon />
                </ListItemIcon>
                <ListItemText primary="Bulk Appointment Letter Generator" />
              </ListItemButton>
            </ListItem>
          </List>
        </Drawer>
        {/* Main Content */}
        <Box component={motion.main}
          initial={{ opacity: 0, y: 40 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, ease: 'easeOut' }}
          sx={{
            flexGrow: 1,
            p: { xs: 1, sm: 4, md: 6 },
            ml: 0,
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'stretch',
            minWidth: 0,
            width: '100%',
            minHeight: '100vh',
            background: '#f5f6fa',
          }}
        >
          <Paper elevation={2} sx={{ p: { xs: 2, sm: 4 }, width: '100%', flexGrow: 1, mt: 10, borderRadius: 4, boxSizing: 'border-box', position: 'relative', background: '#fff', boxShadow: '0 2px 8px 0 rgba(33, 43, 54, 0.08)', display: 'flex', flexDirection: 'column' }} component={motion.div} initial={{ scale: 0.98, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} transition={{ delay: 0.2, duration: 0.5 }}>
            {loading && <LinearProgress sx={{ position: 'absolute', top: 0, left: 0, width: '100%', zIndex: 2 }} />}
            <Stack spacing={3} alignItems="center" sx={{ width: '100%' }}>
              <Typography variant="h4" fontWeight={900} gutterBottom align="center" color="primary">
                Bulk Appointment Letter Generator
              </Typography>
              <motion.div whileHover={{ scale: 1.05, boxShadow: '0 4px 20px rgba(25, 118, 210, 0.15)' }} whileTap={{ scale: 0.97 }}>
                <Button
                  variant="contained"
                  component="label"
                  startIcon={<CloudUploadIcon />}
                  sx={{ minWidth: 200, fontSize: 18, borderRadius: 3 }}
                >
                  Upload Excel File
                  <input
                    type="file"
                    hidden
                    accept=".xlsx,.xls"
                    onChange={handleFileChange}
                    ref={fileInputRef}
                  />
                </Button>
              </motion.div>
              <motion.div whileHover={{ scale: 1.05, boxShadow: '0 4px 20px rgba(25, 118, 210, 0.10)' }} whileTap={{ scale: 0.97 }}>
                <Button
                  variant="outlined"
                  startIcon={<DownloadIcon />}
                  onClick={handleDownloadSample}
                  sx={{ borderRadius: 3 }}
                >
                  Download Sample Format
                </Button>
              </motion.div>
              {excelData.length > 1 && (
                <Typography variant="subtitle1" fontWeight={600} gutterBottom sx={{ mt: 2, mb: 1, textAlign: 'right', width: '100%' }}>
                  Total Employees: {excelData.length - 1}
                </Typography>
              )}
              {excelData.length > 1 && (
                <Box width="100%" flex={1} component={motion.div} initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.3 }}>
                  {/* Search Bar */}
                  <Box mb={2} display="flex" justifyContent="flex-end">
                    <TextField
                      label="Search by Name, Contact, Email"
                      variant="outlined"
                      size="small"
                      value={search}
                      onChange={e => { setSearch(e.target.value); setPage(1); }}
                      sx={{ width: 320 }}
                    />
                  </Box>
                  <TableContainer component={Paper} sx={{ maxHeight: '60vh', minHeight: 200, boxShadow: 1, border: '1px solid #e3e6f0', width: '100%', borderRadius: 0, p: 0, overflowY: 'auto', overflowX: 'auto' }}>
                    <Table size="small" stickyHeader sx={{ width: 'max-content', minWidth: '100%' }}>
                      <colgroup>
                        {excelData[0]?.map((col: string, idx: number) => (
                          <col key={idx} style={{ width: colWidths[idx] || defaultColWidth }} />
                        ))}
                      </colgroup>
                      <TableHead>
                        <TableRow>
                          <TableCell padding="checkbox" align="center" sx={{ background: '#1976d2', borderRight: '1px solid #e3e6f0', width: 40 }}>
                            <Checkbox
                              indeterminate={selectedRows.length > 0 && selectedRows.length < filteredRows.length}
                              checked={filteredRows.length > 0 && selectedRows.length === filteredRows.length}
                              onChange={handleSelectAll}
                              inputProps={{ 'aria-label': 'select all rows' }}
                              sx={{
                                color: '#fff',
                                '&.Mui-checked': { color: '#fff' },
                                '& .MuiSvgIcon-root': {
                                  borderRadius: '4px',
                                  boxShadow: '0 1px 4px 0 rgba(25, 118, 210, 0.10)',
                                },
                                '&:hover': {
                                  backgroundColor: 'rgba(25, 118, 210, 0.15)',
                                  borderRadius: '4px',
                                },
                              }}
                              icon={<svg width="18" height="18" viewBox="0 0 18 18"><rect width="18" height="18" rx="4" fill="#fff" stroke="#1976d2" strokeWidth="2"/></svg>}
                              checkedIcon={<svg width="18" height="18" viewBox="0 0 18 18"><rect width="18" height="18" rx="4" fill="#fff" stroke="#1976d2" strokeWidth="2"/><path d="M5 9.5L8 12.5L13 7.5" stroke="#1976d2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>}
                              indeterminateIcon={<svg width="18" height="18" viewBox="0 0 18 18"><rect width="18" height="18" rx="4" fill="#fff" stroke="#1976d2" strokeWidth="2"/><rect x="4" y="8.25" width="10" height="1.5" rx="0.75" fill="#1976d2"/></svg>}
                            />
                          </TableCell>
                          {excelData[0].map((col: string, idx: number) => {
                            const isAddress = col === 'Address';
                            return (
                              <TableCell
                                key={idx}
                                sx={{
                                  color: '#fff',
                                  fontWeight: 700,
                                  fontSize: 15,
                                  borderRight: '1px solid #e3e6f0',
                                  background: '#1976d2',
                                  whiteSpace: isAddress ? 'nowrap' : 'normal',
                                  py: 0.25,
                                  px: 0,
                                  borderTopLeftRadius: 0,
                                  borderTopRightRadius: 0,
                                  overflow: isAddress ? 'hidden' : 'visible',
                                  textOverflow: isAddress ? 'ellipsis' : 'unset',
                                  wordBreak: 'break-word',
                                  position: 'relative',
                                  minWidth: 80,
                                  maxWidth: 600,
                                  textAlign: 'center',
                                }}
                                align="center"
                              >
                                <ResizableBox
                                  width={colWidths[idx] || defaultColWidth}
                                  height={16}
                                  axis="x"
                                  resizeHandles={['e']}
                                  minConstraints={[80, 16]}
                                  maxConstraints={[600, 16]}
                                  handle={<span style={{ position: 'absolute', right: 0, top: 0, height: '100%', width: 8, cursor: 'col-resize', zIndex: 2, background: 'rgba(25, 118, 210, 0.15)' }} />}
                                  onResizeStop={(e: unknown, data: { size: { width: number } }) => handleResize(idx, data.size.width)}
                                  draggableOpts={{ enableUserSelectHack: false }}
                                >
                                  <div style={{ width: '100%', height: '100%', padding: '0 4px', display: 'flex', alignItems: 'center', justifyContent: 'center', boxSizing: 'border-box', overflow: isAddress ? 'hidden' : 'visible', textOverflow: isAddress ? 'ellipsis' : 'unset', whiteSpace: isAddress ? 'nowrap' : 'normal', wordBreak: 'break-word', textAlign: 'center' }}>
                                    {col}
                                  </div>
                                </ResizableBox>
                              </TableCell>
                            );
                          })}
                        </TableRow>
                      </TableHead>
                      <TableBody>
                        {filteredRows
                          .slice((page - 1) * rowsPerPage, (page - 1) * rowsPerPage + rowsPerPage)
                          .map((row: any[], idx: number) => {
                            const globalIdx = (page - 1) * rowsPerPage + idx;
                            return (
                              <TableRow
                                key={idx}
                                sx={{
                                  backgroundColor: idx % 2 === 0 ? '#f8fafc' : '#e3e6f0',
                                  '&:hover': {
                                    backgroundColor: '#e0f7fa',
                                    transition: 'background 0.2s',
                                  },
                                  cursor: 'pointer',
                                }}
                              >
                                <TableCell padding="checkbox" align="center">
                                  <Checkbox
                                    checked={selectedRows.includes(globalIdx)}
                                    onChange={() => handleSelectRow(globalIdx)}
                                    inputProps={{ 'aria-label': 'select row' }}
                                    sx={{
                                      color: '#1976d2',
                                      '&.Mui-checked': { color: '#1976d2' },
                                      '& .MuiSvgIcon-root': {
                                        borderRadius: '4px',
                                        boxShadow: '0 1px 4px 0 rgba(25, 118, 210, 0.10)',
                                      },
                                      '&:hover': {
                                        backgroundColor: 'rgba(25, 118, 210, 0.08)',
                                        borderRadius: '4px',
                                      },
                                    }}
                                    icon={<svg width="18" height="18" viewBox="0 0 18 18"><rect width="18" height="18" rx="4" fill="#1976d2"/></svg>}
                                    checkedIcon={<svg width="18" height="18" viewBox="0 0 18 18"><rect width="18" height="18" rx="4" fill="#1976d2"/><path d="M5 9.5L8 12.5L13 7.5" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>}
                                    indeterminateIcon={<svg width="18" height="18" viewBox="0 0 18 18"><rect width="18" height="18" rx="4" fill="#1976d2"/><rect x="4" y="8.25" width="10" height="1.5" rx="0.75" fill="#fff"/></svg>}
                                  />
                                </TableCell>
                                {row.map((cell, cidx) => {
                                  const colName = excelData[0][cidx];
                                  const isDateCol = colName === 'Date of Joining' || colName === 'Effective Date';
                                  const isAddress = colName === 'Address';
                                  const displayValue = isDateCol ? formatDateUI(cell) : cell;
                                  return (
                                    <TableCell
                                      key={cidx}
                                      align="center"
                                      sx={{
                                        fontSize: 15,
                                        borderRight: '1px solid #f0f0f0',
                                        py: 0.25,
                                        px: 2,
                                        overflow: isAddress ? 'hidden' : 'visible',
                                        textOverflow: isAddress ? 'ellipsis' : 'unset',
                                        whiteSpace: isAddress ? 'nowrap' : 'normal',
                                        background: 'inherit',
                                        wordBreak: 'break-word',
                                        minWidth: colWidths[cidx] || defaultColWidth,
                                        maxWidth: 600,
                                        textAlign: 'center',
                                      }}
                                    >
                                      {isAddress ? (
                                        <Tooltip title={displayValue || ''} placement="top" arrow>
                                          <span style={{ display: 'inline-block', width: '100%', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', verticalAlign: 'bottom', wordBreak: 'break-word', textAlign: 'center' }}>{displayValue}</span>
                                        </Tooltip>
                                      ) : (
                                        <span style={{ display: 'inline-block', width: '100%', verticalAlign: 'bottom', wordBreak: 'break-word', textAlign: 'center' }}>{displayValue}</span>
                                      )}
                                    </TableCell>
                                  );
                                })}
                              </TableRow>
                            );
                          })}
                      </TableBody>
                    </Table>
                  </TableContainer>
                  {/* Pagination and Rows Per Page Controls */}
                  <Box display="flex" justifyContent="space-between" alignItems="center" mt={2}>
                    <Box display="flex" alignItems="center">
                      <Typography variant="body2" sx={{ mr: 1 }}>Rows per page:</Typography>
                      <Select value={rowsPerPage} onChange={(event, child) => handleChangeRowsPerPage(event as React.ChangeEvent<{ value: unknown }>)} size="small">
                        {[5, 10, 20, 50, 100].map((option) => (
                          <MenuItem key={option} value={option}>{option}</MenuItem>
                        ))}
                      </Select>
                    </Box>
                    <Pagination
                      count={Math.ceil(filteredRows.length / rowsPerPage)}
                      page={page}
                      onChange={handleChangePage}
                      color="primary"
                      shape="rounded"
                      size="medium"
                    />
                  </Box>
                </Box>
              )}
              <motion.div whileHover={{ scale: 1.05, boxShadow: '0 4px 20px rgba(245, 0, 87, 0.10)' }} whileTap={{ scale: 0.97 }}>
                <Button
                  variant="contained"
                  color="secondary"
                  startIcon={<DownloadIcon />}
                  onClick={handleGenerateDoc}
                  disabled={!excelFile || loading}
                  sx={{ borderRadius: 3, minWidth: 200, fontSize: 18 }}
                >
                  {loading ? 'Generating...' : 'Generate Appointment Letters'}
                </Button>
              </motion.div>
            </Stack>
          </Paper>
        </Box>
      </Box>
    </ThemeProvider>
  )
}

export default App
