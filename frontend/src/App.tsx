import { useState } from 'react'
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
      return format(jsDate, 'd MMM yyyy');
    }
  }
  // Try to parse as ISO or fallback to Date
  let d = typeof date === 'string' ? parseISO(date) : new Date(date);
  if (!isValid(d)) d = new Date(date);
  if (!isValid(d)) return date;
  return format(d, 'd MMM yyyy');
}

function App() {
  const [excelFile, setExcelFile] = useState<File | null>(null)
  const [excelData, setExcelData] = useState<any[][]>([])
  const [loading, setLoading] = useState(false)
  const [downloaded, setDownloaded] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const theme = useTheme()
  const isMobile = useMediaQuery(theme.breakpoints.down('sm'))
  const [selectedPage, setSelectedPage] = useState('bulk')

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    setExcelFile(file || null)
    setDownloaded(false)
    setError(null)
    if (file) {
      const reader = new FileReader()
      reader.onload = (evt) => {
        const bstr = evt.target?.result as string
        const wb = XLSX.read(bstr, { type: 'binary' })
        const wsname = wb.SheetNames[0]
        const ws = wb.Sheets[wsname]
        const data = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 })
        if (data.length === 0) {
          setExcelData([])
          return
        }
        const header = data[0] as string[]
        const colIndexes = REQUIRED_COLUMNS.map(col => header.indexOf(col))
        const filteredData = [
          REQUIRED_COLUMNS,
          ...data.slice(1).map((row: any) => colIndexes.map(idx => (idx !== -1 ? row[idx] : '')))
        ]
        setExcelData(filteredData)
      }
      reader.readAsBinaryString(file)
    }
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

  const handleGenerateDoc = async () => {
    if (!excelFile) return
    setLoading(true)
    setDownloaded(false)
    setError(null)
    const formData = new FormData()
    formData.append('file', excelFile)
    try {
      const response = await fetch('http://localhost:5000/upload-excel', {
        method: 'POST',
        body: formData,
      })
      if (!response.ok) throw new Error('Failed to generate document')
      const blob = await response.blob()
      const disposition = response.headers.get('Content-Disposition');
      let filename = 'appointment_letters.zip';
      if (disposition && disposition.indexOf('filename=') !== -1) {
        filename = disposition.split('filename=')[1].replace(/['"]/g, '');
      }
      saveAs(blob, filename)
      setDownloaded(true)
    } catch (err) {
      setError('Error generating document. Please try again.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <Box sx={{ display: 'flex', minHeight: '100vh' }}>
      {/* Header */}
      <AppBar position="fixed" color="default" elevation={1} sx={{ zIndex: (theme) => theme.zIndex.drawer + 1 }}>
        <Toolbar sx={{ justifyContent: 'flex-end' }}>
          <Typography variant="h6" component="div" sx={{ fontWeight: 700 }}>
            InTimeTec.
          </Typography>
        </Toolbar>
      </AppBar>
      {/* Side Navigation Bar */}
      <Drawer
        variant={isMobile ? 'temporary' : 'permanent'}
        open={true}
        sx={{
          width: 240,
          flexShrink: 0,
          '& .MuiDrawer-paper': {
            width: 240,
            boxSizing: 'border-box',
            borderRight: '1px solid #e0e0e0',
          },
        }}
        ModalProps={{ keepMounted: true }}
      >
        <Box sx={{ height: 64, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: 18 }}>
          Menu
        </Box>
        <Divider />
        <List>
          <ListItem disablePadding>
            <ListItemButton selected={selectedPage === 'bulk'} onClick={() => setSelectedPage('bulk')}>
              <ListItemIcon>
                <DescriptionIcon />
              </ListItemIcon>
              <ListItemText primary="Bulk Appointment Letter Generator" />
            </ListItemButton>
          </ListItem>
        </List>
      </Drawer>
      {/* Main Content */}
      <Box component="main" sx={{
        flexGrow: 1,
        p: { xs: 1, sm: 3, md: 5 },
        ml: 0,
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        minHeight: '100vh',
        background: '#fafbfc',
        overflow: 'hidden',
        pt: 8, // Add top padding to avoid header overlap
      }}>
        {selectedPage === 'bulk' && (
          <Paper elevation={3} sx={{
            p: { xs: 2, sm: 4 },
            mb: 4,
            width: '100%',
            mx: 0,
            boxShadow: 2,
            minHeight: 'calc(100vh - 100px)',
            display: 'flex',
            flexDirection: 'column',
          }}>
            <Typography variant="h4" fontWeight={700} gutterBottom sx={{ textAlign: 'left' }}>
              Bulk Appointment Letter Generator
            </Typography>
            <Typography variant="subtitle1" color="text.secondary" mb={2} sx={{ textAlign: 'left' }}>
              Upload your Excel, preview the data, and generate appointment letters in bulk. Supports HR and Effective Date fields.
            </Typography>
            <Stack direction={isMobile ? 'column' : 'row'} spacing={2} justifyContent="flex-start" alignItems={isMobile ? 'stretch' : 'center'} mb={2}>
              <Button
                variant="outlined"
                startIcon={<DownloadIcon />}
                onClick={handleDownloadSample}
                color="primary"
              >
                Download Sample Excel
              </Button>
              <Button
                variant="contained"
                component="label"
                startIcon={<CloudUploadIcon />}
                color="primary"
              >
                {excelFile ? 'Change Excel File' : 'Upload Excel File'}
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  hidden
                  onChange={handleFileChange}
                />
              </Button>
              {excelFile && (
                <Stack direction="row" alignItems="center" spacing={1}>
                  <InsertDriveFileIcon color="action" />
                  <Typography variant="body2" noWrap maxWidth={120} title={excelFile.name} sx={{ textAlign: 'left' }}>
                    {excelFile.name}
                  </Typography>
                </Stack>
              )}
            </Stack>
            {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}
            {loading && <LinearProgress sx={{ mb: 2 }} />}
            {excelData.length > 0 && (
              <Box mt={3}>
                <Typography variant="h6" gutterBottom sx={{ textAlign: 'left' }}>Data Preview</Typography>
                <TableContainer component={Paper} sx={{ maxHeight: 400, borderRadius: 2, boxShadow: 2 }}>
                  <Table size="small" stickyHeader aria-label="data preview table">
                    <TableHead>
                      <TableRow sx={{ backgroundColor: 'primary.main' }}>
                        {REQUIRED_COLUMNS.map((col) => (
                          <TableCell
                            key={col}
                            sx={{
                              backgroundColor: 'primary.main',
                              color: 'white',
                              fontWeight: 700,
                              fontSize: { xs: 12, sm: 14 },
                              borderRight: '1px solid #e0e0e0',
                            }}
                            align="center"
                          >
                            {col}
                          </TableCell>
                        ))}
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {excelData.slice(1).map((row, i) => (
                        <TableRow
                          key={i}
                          sx={{
                            backgroundColor: i % 2 === 0 ? 'background.paper' : 'grey.100',
                            '&:last-child td, &:last-child th': { border: 0 },
                          }}
                        >
                          {row.map((cell: any, j: number) => {
                            const colName = REQUIRED_COLUMNS[j];
                            const value = (colName === 'Date of Joining' || colName === 'Effective Date') ? formatDateUI(cell) : cell;
                            return (
                              <TableCell
                                key={j}
                                align="center"
                                sx={{
                                  fontSize: { xs: 12, sm: 14 },
                                  borderRight: '1px solid #f0f0f0',
                                  py: 1.2,
                                  px: { xs: 0.5, sm: 1 },
                                  maxWidth: 180,
                                  overflow: 'hidden',
                                  textOverflow: 'ellipsis',
                                  whiteSpace: 'nowrap',
                                }}
                              >
                                <Tooltip title={value || ''} placement="top" arrow>
                                  <span style={{ display: 'inline-block', width: '100%', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', verticalAlign: 'bottom' }}>{value}</span>
                                </Tooltip>
                              </TableCell>
                            );
                          })}
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </TableContainer>
              </Box>
            )}
            <Box textAlign="left" mt={4}>
              <Button
                variant="contained"
                color="success"
                size="large"
                onClick={handleGenerateDoc}
                disabled={!excelFile || loading}
                sx={{ minWidth: 220 }}
              >
                {loading ? 'Generating...' : 'Generate Appointment Letters'}
              </Button>
              {downloaded && !loading && (
                <Alert severity="success" sx={{ mt: 2 }}>
                  Appointment letters downloaded!
                </Alert>
              )}
            </Box>
          </Paper>
        )}
        <Box textAlign="center" color="text.secondary" fontSize={14} mt={2}>
          &copy; {new Date().getFullYear()} Bulk Appointment Letter Generator
        </Box>
      </Box>
    </Box>
  )
}

export default App
