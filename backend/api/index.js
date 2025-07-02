const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const archiver = require('archiver');

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }
  let tempFilePath = null;
  let responseSent = false;
  try {
    let data;
    if (req.body && req.body.data) {
      let parsed = req.body.data;
      if (typeof parsed === 'string') parsed = JSON.parse(parsed);
      const [header, ...rows] = parsed;
      data = rows.map(row => Object.fromEntries(header.map((h, i) => [h, row[i]])));
    } else {
      throw new Error('No data provided');
    }
    if (!data.length) throw new Error('No data found in Excel');
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename=appointment_letters.zip');
    const archive = archiver('zip');
    let archiveError = false;
    archive.pipe(res);
    archive.on('error', (err) => {
      archiveError = true;
      console.error('Archive error:', err);
      res.end();
    });
    archive.on('end', () => {
      console.log('Archive stream ended. Total bytes:', archive.pointer());
    });
    res.on('error', (err) => {
      console.error('Response error:', err);
    });
    for (let idx = 0; idx < data.length; idx++) {
      const row = data[idx];
      let templateFile = 'EmploymentAgreementandAppointment.docx';
      let folderName = 'Other';
      let designation = (row['Designation'] || '').trim().toLowerCase();
      if (designation === 'jr. software engineer' || designation === 'junior software engineer') {
        templateFile = 'Junior Software Engineer-Appointment_Letter.docx';
        folderName = 'Junior Software Engineer';
      }
      let safeName = row['Name'] ? String(row['Name']).replace(/[^a-zA-Z0-9 \-_\.]/g, '').trim() : '';
      let fileName = safeName ? `${safeName}.docx` : `Appointment.docx`;
      const zipPath = `${folderName}/${fileName}`;
      const templatePath = path.join(__dirname, '../templates', templateFile);
      const content = fs.readFileSync(templatePath, 'binary');
      let dateOfJoining = row['Date of Joining'] || '';
      if (dateOfJoining) {
        let dateObj;
        if (typeof dateOfJoining === 'number') {
          const parsed = XLSX.SSF.parse_date_code(dateOfJoining);
          if (parsed) {
            dateObj = new Date(parsed.y, parsed.m - 1, parsed.d);
          }
        } else {
          dateObj = new Date(dateOfJoining);
        }
        if (dateObj && !isNaN(dateObj.getTime())) {
          const day = String(dateObj.getDate()).padStart(2, '0');
          const month = dateObj.toLocaleString('default', { month: 'long' });
          const year = dateObj.getFullYear();
          dateOfJoining = `${day}-${month}-${year}`;
        }
      }
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
  }
}; 