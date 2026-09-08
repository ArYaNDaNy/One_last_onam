const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_FILE = path.join(__dirname, 'contributions.xlsx');
const JSON_FILE = path.join(__dirname, 'contributions.json');
const CSV_FILE = path.join(__dirname, 'contributions.csv');

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(__dirname));

// Serve index.html at root
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

const COLUMNS = [
    { header: 'Timestamp', key: 'timestamp', width: 22 },
    { header: 'name', key: 'name', width: 25 },
    { header: 'tranasaction id', key: 'transaction_id', width: 25 },
    { header: 'payasam count', key: 'payasam_count', width: 15 },
    { header: 'mullapoo count', key: 'mullapoo_count', width: 15 },
    { header: 'contribution for pookalam amnt', key: 'pookalam_amount', width: 30 },
    { header: 'total amt', key: 'total_amount', width: 15 }
];

// Helper to append to JSON backup file
function appendToJsonBackup(entry) {
    try {
        let entries = [];
        if (fs.existsSync(JSON_FILE)) {
            const raw = fs.readFileSync(JSON_FILE, 'utf8').trim();
            if (raw) entries = JSON.parse(raw);
        }
        entries.push(entry);
        fs.writeFileSync(JSON_FILE, JSON.stringify(entries, null, 2), 'utf8');
    } catch (e) {
        console.error('Warning: Failed to write to JSON backup:', e.message);
    }
}

// Helper to append to CSV backup file
function appendToCsvBackup(entry) {
    try {
        if (!fs.existsSync(CSV_FILE)) {
            const header = 'Timestamp,Name,Transaction ID,Payasam Count,Mullapoo Count,Pookalam Amount,Total Amount\n';
            fs.writeFileSync(CSV_FILE, header, 'utf8');
        }
        const sanitize = (val) => `"${String(val ?? '').replace(/"/g, '""')}"`;
        const line = `${sanitize(entry.timestamp)},${sanitize(entry.name)},${sanitize(entry.transactionId)},${entry.payasamCount},${entry.mullapooCount},${entry.pookalamAmount},${entry.totalAmount}\n`;
        fs.appendFileSync(CSV_FILE, line, 'utf8');
    } catch (e) {
        console.error('Warning: Failed to write to CSV backup:', e.message);
    }
}

// Function to get or create workbook and sheet
async function getOrCreateSheet() {
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    if (fs.existsSync(EXCEL_FILE)) {
        try {
            await workbook.xlsx.readFile(EXCEL_FILE);
            worksheet = workbook.getWorksheet('Contributions') || workbook.worksheets[0];
            if (worksheet) {
                worksheet.columns = COLUMNS;
            }
        } catch (readErr) {
            console.warn('Could not read existing Excel file, creating fresh sheet:', readErr.message);
        }
    }

    if (!worksheet) {
        worksheet = workbook.addWorksheet('Contributions');
        worksheet.columns = COLUMNS;

        // Format header row style
        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        worksheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD4AF37' } // Gold theme
        };
    }

    return { workbook, worksheet };
}

// API Health Check
app.get('/api/health', (req, res) => {
    let count = 0;
    if (fs.existsSync(JSON_FILE)) {
        try { count = JSON.parse(fs.readFileSync(JSON_FILE, 'utf8')).length; } catch {}
    }
    res.json({ status: 'ok', serverTime: new Date().toISOString(), totalContributionsRecorded: count });
});

// View all contributions in JSON format
app.get('/api/contributions', (req, res) => {
    try {
        if (fs.existsSync(JSON_FILE)) {
            const raw = fs.readFileSync(JSON_FILE, 'utf8').trim();
            return res.json(raw ? JSON.parse(raw) : []);
        }
        return res.json([]);
    } catch (e) {
        return res.status(500).json({ error: e.message });
    }
});

// API Endpoint to receive contribution data
app.post('/api/submit', async (req, res) => {
    try {
        // Support all payload variations (camelCase, form names, etc.)
        const name = (req.body.name || req.body.userName || '').trim();
        const transactionId = (req.body.transactionId || req.body.txnId || req.body['txn-id'] || '').trim();
        const payasamCount = Number(req.body.payasamCount ?? req.body.payasamQty ?? 0);
        const mullapooCount = Number(req.body.mullapooCount ?? req.body.mullapooQty ?? 0);
        const pookalamAmount = Number(req.body.pookalamAmount ?? req.body.pookalamAmt ?? 0);
        const totalAmount = Number(req.body.totalAmount ?? req.body.grandTotal ?? 0);

        if (!name || !transactionId) {
            return res.status(400).json({ 
                success: false, 
                message: 'Name and Transaction ID are required fields.' 
            });
        }

        const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });

        const record = {
            timestamp,
            name,
            transactionId,
            payasamCount,
            mullapooCount,
            pookalamAmount,
            totalAmount
        };

        // 1. Always record in JSON and CSV backups immediately (these never lock)
        appendToJsonBackup(record);
        appendToCsvBackup(record);

        // 2. Record in Excel file
        let excelSaved = false;
        try {
            const { workbook, worksheet } = await getOrCreateSheet();

            worksheet.addRow([
                timestamp,
                name,
                transactionId,
                payasamCount,
                mullapooCount,
                pookalamAmount,
                totalAmount
            ]);

            await workbook.xlsx.writeFile(EXCEL_FILE);
            excelSaved = true;
        } catch (excelErr) {
            if (excelErr.code === 'EBUSY') {
                console.warn(`[${timestamp}] Notice: contributions.xlsx is currently open in Excel. Record saved to JSON & CSV backups.`);
            } else {
                console.error(`[${timestamp}] Error writing to Excel:`, excelErr.message);
            }
        }

        console.log(`[${timestamp}] ✓ Recorded contribution: ${name} (Txn: ${transactionId}) - Total: ₹${totalAmount} [Excel: ${excelSaved ? 'saved' : 'backed up in JSON/CSV'}]`);

        return res.json({ 
            success: true, 
            message: excelSaved 
                ? 'Contribution recorded in Excel successfully!' 
                : 'Contribution saved safely in backup logs (Excel file was open).' 
        });
    } catch (error) {
        console.error('Error processing contribution submission:', error);
        return res.status(500).json({ success: false, message: 'Server error processing contribution.' });
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

// Download live CSV sheet endpoint
app.get('/api/download-csv', (req, res) => {
    if (fs.existsSync(CSV_FILE)) {
        res.download(CSV_FILE, 'contributions.csv');
    } else {
        res.status(404).send('CSV file has not been created yet.');
    }
});

const server = app.listen(PORT, () => {
    console.log(`Pookalam Backend Server running at http://localhost:${PORT}`);
    console.log(`Excel sheet path: ${EXCEL_FILE}`);
    console.log(`JSON backup path: ${JSON_FILE}`);
});

server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
        console.error(`Port ${PORT} is already in use.`);
    } else {
        console.error('Server error:', err);
    }
});
