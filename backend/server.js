const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const archiver = require('archiver');
const app = express();
const PORT = process.env.PORT || 5000;
const uploadExcelRouter = require('./routes/uploadExcel');

app.use(cors());
app.use(express.json());

// Serve frontend static files
app.use(express.static(path.join(__dirname, '../frontend/dist')));

// Catch-all: serve React app for any non-API route
app.get(/^((?!\/upload-excel).)*$/, (req, res) => {
  res.sendFile(path.join(__dirname, '../frontend/dist/index.html'));
});

const upload = multer({ dest: 'uploads/' });

app.use('/upload-excel', uploadExcelRouter);

app.post('/upload-excel', upload.single('file'), async (req, res) => {
  let tempFilePath = req.file ? req.file.path : null;
  let responseSent = false;
  try {
    let data;
    if (req.body && req.body.data) {
      // Handle JSON data (selected rows)
      let parsed = req.body.data;
      if (typeof parsed === 'string') parsed = JSON.parse(parsed);
      // Convert array of arrays to array of objects using header
      const [header, ...rows] = parsed;
      data = rows.map(row => Object.fromEntries(header.map((h, i) => [h, row[i]])));
    } else if (req.file) {
      // Handle file upload (all rows)
      const workbook = XLSX.readFile(req.file.path);
      const sheetName = workbook.SheetNames[0];
      data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    } else {
      throw new Error('No data provided');
    }
    if (!data.length) throw new Error('No data found in Excel');

    // Step 2: Create and stream the zip as docs are generated
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename=appointment_letters.zip');
    const archiverLib = require('archiver');
    const archive = archiverLib('zip');
    let archiveError = false;
    archive.pipe(res);

    // Handle archive errors
    archive.on('error', (err) => {
      archiveError = true;
      console.error('Archive error:', err);
      res.end();
      if (tempFilePath && fs.existsSync(tempFilePath)) {
        fs.unlinkSync(tempFilePath);
      }
    });

    // Log when archive is fully written
    archive.on('end', () => {
      console.log('Archive stream ended. Total bytes:', archive.pointer());
    });

    // Clean up after response is finished
    res.on('finish', () => {
      if (tempFilePath && fs.existsSync(tempFilePath)) {
        fs.unlinkSync(tempFilePath);
      }
      if (!archiveError) {
        console.log('Response finished successfully.');
      }
    });
    res.on('close', () => {
      if (tempFilePath && fs.existsSync(tempFilePath)) {
        fs.unlinkSync(tempFilePath);
      }
      if (!archiveError) {
        console.log('Response closed.');
      }
    });
    res.on('error', (err) => {
      console.error('Response error:', err);
    });

    // Step 1: Generate and append each Word document to the zip
    for (let idx = 0; idx < data.length; idx++) {
      const row = data[idx];
      // Pick template and folder based on designation
      let templateFile = 'EmploymentAgreementandAppointment.docx';
      let folderName = 'Other';
      let designation = (row['Designation'] || '').trim().toLowerCase();
      if (designation === 'jr. software engineer' || designation === 'junior software engineer') {
        templateFile = 'JuniorSoftwareEngineer-Appointment_Letter.docx';
        folderName = 'Junior Software Engineer';
      }
      const templatePath = path.join(__dirname, 'templates', templateFile);
      const content = fs.readFileSync(templatePath, 'binary');
      let dateOfJoining = row['Date of Joining'] || '';
      if (dateOfJoining) {
        let dateObj;
        if (typeof dateOfJoining === 'number') {
          // Excel date serial
          const parsed = XLSX.SSF.parse_date_code(dateOfJoining);
          if (parsed) {
            dateObj = new Date(parsed.y, parsed.m - 1, parsed.d);
          }
        } else {
          // Try parsing as string
          dateObj = new Date(dateOfJoining);
        }
        if (dateObj && !isNaN(dateObj.getTime())) {
          const day = String(dateObj.getDate()).padStart(2, '0');
          const month = dateObj.toLocaleString('default', { month: 'long' });
          const year = dateObj.getFullYear();
          dateOfJoining = `${day}-${month}-${year}`;
        }
      }
      // Format Effective Date
      let effectiveDate = row['Effective Date'] || '';
      if (effectiveDate) {
        let effDateObj;
        if (typeof effectiveDate === 'number') {
          const parsed = XLSX.SSF.parse_date_code(effectiveDate);
          if (parsed) {
            effDateObj = new Date(parsed.y, parsed.m - 1, parsed.d);
          }
        } else {
          effDateObj = new Date(effectiveDate);
        }
        if (effDateObj && !isNaN(effDateObj.getTime())) {
          const day = String(effDateObj.getDate()).padStart(2, '0');
          const month = effDateObj.toLocaleString('default', { month: 'long' });
          const year = effDateObj.getFullYear();
          effectiveDate = `${day}-${month}-${year}`;
        }
      }
      const zip = new PizZip(content);
      const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
      // Sanitize the name to remove characters not allowed in filenames
      let safeName = row['Name'] ? String(row['Name']).replace(/[^a-zA-Z0-9 \-_\.]/g, '').trim() : '';
      let fileName = safeName ? `${safeName}.docx` : 'Appointment.docx';
      // Add to the correct folder in the zip
      const zipPath = `${folderName}/${fileName}`;
      doc.setData({
        Name: row['Name'] || '',
        Email: row['Email'] || '',
        Contact: row['Contact'] || '',
        'Date of Joining': dateOfJoining,
        Designation: row['Designation'] || '',
        'Place of Joining': row['Place of Joining'] || '',
        Address: row['Address'] || '',
        'HR Name': row['HR Name'] || '',
        'HR Designation': row['HR Designation'] || '',
        'Effective Date': effectiveDate,
      });
      try {
        doc.render();
      } catch (error) {
        console.error('Docxtemplater error:', error);
        if (!responseSent) {
          responseSent = true;
          res.status(500).json({ success: false, error: 'Failed to generate document for ' + (row['Name'] || 'unknown') });
        }
        if (tempFilePath && fs.existsSync(tempFilePath)) {
          fs.unlinkSync(tempFilePath);
        }
        return;
      }
      const buf = doc.getZip().generate({ type: 'nodebuffer' });
      archive.append(buf, { name: zipPath });
    }

    archive.finalize().catch((err) => {
      archiveError = true;
      console.error('Archive finalize error:', err);
      res.end();
    });
  } catch (err) {
    console.error('General error:', err);
    if (!responseSent) {
      responseSent = true;
      res.status(500).json({ success: false, error: err.message });
    }
    if (tempFilePath && fs.existsSync(tempFilePath)) {
      fs.unlinkSync(tempFilePath);
    }
  }
});

app.listen(PORT, () => {
  console.log(`Backend running on port ${PORT}`);
}); 