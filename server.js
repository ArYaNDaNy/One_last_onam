const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_FILE = path.join(__dirname, 'contributions.xlsx');

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

const COLUMNS = [
    { header: 'Timestamp', key: 'timestamp', width: 22 },
    { header: 'name', key: 'name', width: 25 },
    { header: 'tranasaction id', key: 'transaction_id', width: 25 },
    { header: 'payasam count', key: 'payasam_count', width: 15 },
    { header: 'mullapoo count', key: 'mullapoo_count', width: 15 },
    { header: 'contribution for pookalam amnt', key: 'pookalam_amount', width: 30 },
    { header: 'total amt', key: 'total_amount', width: 15 }
];

// Function to get or create workbook and sheet
async function getOrCreateSheet() {
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    if (fs.existsSync(EXCEL_FILE)) {
        await workbook.xlsx.readFile(EXCEL_FILE);
        worksheet = workbook.getWorksheet('Contributions') || workbook.worksheets[0];
        worksheet.columns = COLUMNS;
    } else {
        worksheet = workbook.addWorksheet('Contributions');
        worksheet.columns = COLUMNS;

        // Format header row style
        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        worksheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD4AF37' } // Gold theme
        };
        await workbook.xlsx.writeFile(EXCEL_FILE);
    }
    return { workbook, worksheet };
}

// API Endpoint to receive contribution data
app.post('/api/submit', async (req, res) => {
    try {
        const { name, transactionId, payasamCount, mullapooCount, pookalamAmount, totalAmount } = req.body;

        if (!name || !transactionId) {
            return res.status(400).json({ success: false, message: 'Name and Transaction ID are required.' });
        }

        const { workbook, worksheet } = await getOrCreateSheet();

        const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });

        // Use array format to guarantee column index placement regardless of file reload state
        worksheet.addRow([
            timestamp,
            String(name).trim(),
            String(transactionId).trim(),
            Number(payasamCount || 0),
            Number(mullapooCount || 0),
            Number(pookalamAmount || 30),
            Number(totalAmount || 0)
        ]);

        await workbook.xlsx.writeFile(EXCEL_FILE);

        console.log(`[${timestamp}] Recorded contribution from: ${name} (Txn: ${transactionId}) - Total: ₹${totalAmount}`);
        return res.json({ success: true, message: 'Contribution recorded in Excel successfully!' });
    } catch (error) {
        if (error.code === 'EBUSY') {
            console.error('\n❌ ERROR: Cannot write to Excel file because it is open in another program (like Microsoft Excel). Please close the file and try again.\n');
            return res.status(500).json({ success: false, message: 'Excel file is currently open. Please close it and try again.' });
        }
        console.error('Error logging to Excel:', error);
        return res.status(500).json({ success: false, message: 'Failed to record contribution in Excel.' });
    }
});

// Download live Excel sheet endpoint
app.get('/api/download', (req, res) => {
    if (fs.existsSync(EXCEL_FILE)) {
        res.download(EXCEL_FILE, 'contributions.xlsx');
    } else {
        res.status(404).send('Excel file has not been created yet.');
    }
});

const server = app.listen(PORT, () => {
    console.log(`Pookalam Backend Server running at http://localhost:${PORT}`);
    console.log(`Excel sheet path: ${EXCEL_FILE}`);
});

server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
        console.error(`Port ${PORT} is already in use. Please close the process using port ${PORT} or run on another port e.g. PORT=3001 npm start`);
    } else {
        console.error('Server error:', err);
    }
});
