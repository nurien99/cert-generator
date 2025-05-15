// Import necessary modules
const express = require('express');
const { PDFDocument, rgb, StandardFonts, LineCapStyle, BlendMode } = require('pdf-lib');
const fs = require('fs').promises;
const fsSync = require('fs');
const path = require('path');
const multer = require('multer');
const csv = require('csv-parser');
const archiver = require('archiver');
const fontkit = require('fontkit');

// Initialize Express application
const app = express();
const port = process.env.PORT || 3000;

// --- Multer Configuration ---
const storage = multer.memoryStorage();
const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        if (file.mimetype === 'text/csv' || file.originalname.endsWith('.csv')) {
            cb(null, true);
        } else {
            cb(new Error('Only .csv files are allowed!'), false);
        }
    }
});

// --- Middleware ---
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// --- Helper function to estimate line breaks (simplified) ---
function estimateLines(text, font, fontSize, maxWidth) {
    const words = text.split(' ');
    if (!words.length || text.trim() === '') return [''];

    let currentLine = words[0];
    if (!currentLine && words.length > 1) currentLine = words[1] || '';
    else if (!currentLine) return [''];

    let lines = [];
    for (let i = 1; i < words.length; i++) {
        const word = words[i];
        if (word.trim() === '') continue;

        const testLine = `${currentLine} ${word}`;
        const testWidth = font.widthOfTextAtSize(testLine, fontSize);
        if (testWidth <= maxWidth) {
            currentLine = testLine;
        } else {
            lines.push(currentLine);
            currentLine = word;
        }
    }
    lines.push(currentLine);
    return lines;
}


// --- Reusable PDF Generation Function ---
async function generateCertificatePdfBytes(userName, trainingType, date, templatePath) {
    const existingPdfBytes = await fs.readFile(templatePath);
    const pdfDoc = await PDFDocument.load(existingPdfBytes);

    pdfDoc.registerFontkit(fontkit);

    const georgiaBoldFontBytes = await fs.readFile(path.join(__dirname, 'fonts', 'GEORGIAB.TTF'));
    const georgiaBoldFont = await pdfDoc.embedFont(georgiaBoldFontBytes);

    const pages = pdfDoc.getPages();
    const firstPage = pages[0];
    const { width, height } = firstPage.getSize();

    // --- Function to draw text with true centering for each line ---
    const drawCenteredWrappedText = (text, options) => {
        const { font, fontSize, color, maxWidth, lineHeight, initialY, page } = options;
        const lines = estimateLines(text, font, fontSize, maxWidth);
        const numberOfLines = lines.length;

        // Calculate the Y for the first line to achieve bottom alignment of the block
        const targetBottomY = initialY; // initialY is where the last line's baseline should be
        let currentY = targetBottomY + (numberOfLines - 1) * lineHeight;

        lines.forEach(line => {
            const lineWidth = font.widthOfTextAtSize(line, fontSize);
            const x = (page.getWidth() - lineWidth) / 2; // Center each line on the page
            page.drawText(line, {
                x: x,
                y: currentY,
                size: fontSize,
                font: font,
                color: color,
                lineHeight: lineHeight, // Though not strictly needed here as we draw line by line
            });
            currentY -= lineHeight; // Move Y down for the next line
        });
    };


    // 1. Draw Recipient's Name (True Centered, Bottom-Aligned Vertically)
    const userNameFontSize = 48;
    const userNameText = String(userName || '');
    const maxWidthForUserName = 11.45 * 72; // 824.4 points
    const userNameLineHeight = userNameFontSize * 1.2;
    const targetBottomYForUserName = 520; // Your Y for username's LAST line baseline

    drawCenteredWrappedText(userNameText, {
        font: georgiaBoldFont,
        fontSize: userNameFontSize,
        color: rgb(0, 0, 0),
        maxWidth: maxWidthForUserName,
        lineHeight: userNameLineHeight,
        initialY: targetBottomYForUserName, // This is where the last line's baseline will be
        page: firstPage
    });


    // 2. Draw Training Type (True Centered, Bottom-Aligned Vertically)
    const trainingTypeFontSize = 32;
    const trainingTypeText = String(trainingType || '');
    const maxWidthForTrainingType = 11.85 * 72; // 853.2 points
    const maxHeightForTrainingType = 0.83 * 72; // 59.76 points (for reference)
    const trainingTypeLineHeight = trainingTypeFontSize * 1.2;
    const targetBottomYForTrainingType = 370; // Your Y for training type's LAST line baseline

    // Estimate lines for Training Type for height warning (optional)
    const trainingTypeLinesArray = estimateLines(trainingTypeText, georgiaBoldFont, trainingTypeFontSize, maxWidthForTrainingType);
    const numberOfTrainingTypeLines = trainingTypeLinesArray.length;
    const totalTrainingTypeTextHeight = (numberOfTrainingTypeLines -1) * trainingTypeLineHeight + trainingTypeFontSize;
    if (totalTrainingTypeTextHeight > maxHeightForTrainingType) {
        console.warn(`Warning: Training Type text "${trainingTypeText}" might exceed defined height.`);
    }

    drawCenteredWrappedText(trainingTypeText, {
        font: georgiaBoldFont,
        fontSize: trainingTypeFontSize,
        color: rgb(0, 0, 0),
        maxWidth: maxWidthForTrainingType,
        lineHeight: trainingTypeLineHeight,
        initialY: targetBottomYForTrainingType, // This is where the last line's baseline will be
        page: firstPage
    });


    // 3. Draw Completion Date (Fixed Location)
    let formattedCompletionDate = 'N/A';
    if (date) {
        try {
            formattedCompletionDate = new Date(date).toLocaleDateString('en-US', {
                year: 'numeric', month: 'long', day: 'numeric',
            });
        } catch (e) {
            console.warn("Invalid date format for:", date);
            formattedCompletionDate = String(date || 'N/A');
        }
    }
    const dateFontSize = 28;
    firstPage.drawText(formattedCompletionDate, {
        x: 880,    // Your fixed X for date
        y: 330,    // Your fixed Y for date (from bottom-left)
        size: dateFontSize,
        font: georgiaBoldFont,
        color: rgb(0, 0, 0),
    });

    return await pdfDoc.save();
}

// --- HTTP Routes ---
// GET /: Serves the main HTML form.
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// POST /generate-certificate: Handles SINGLE certificate generation.
app.post('/generate-certificate', async (req, res) => {
    try {
        const { userName, trainingType, date } = req.body;
        if (!userName || !trainingType || !date) {
            return res.status(400).send('All fields (User Name, Training Type, Date) are required.');
        }
        const templatePath = path.join(__dirname, 'templates', 'iLab Certificate.pdf');
        const pdfBytes = await generateCertificatePdfBytes(userName, trainingType, date, templatePath);
        const sanitizedUserName = String(userName || 'user').replace(/[^a-z0-9]/gi, '_').toLowerCase();
        res.setHeader('Content-Disposition', `attachment; filename="certificate_${sanitizedUserName}.pdf"`);
        res.setHeader('Content-Type', 'application/pdf');
        res.send(Buffer.from(pdfBytes));
    } catch (error) {
        console.error('Error generating single PDF:', error);
        if (error.message.includes('No such file or directory') && error.message.includes('fonts')) {
             res.status(500).send('Error generating certificate: Custom font file (e.g., GEORGIAB.TTF) not found in the "fonts" folder. Please check the path and filename in `app.js`.');
        } else if (error.message.includes('fontkit instance was found')) {
             res.status(500).send(`Error generating certificate with custom font: ${error.message}. Ensure fontkit is installed (run 'npm install') and registered correctly in app.js (using 'pdfDoc.registerFontkit(fontkit)').`);
        } else {
             res.status(500).send('Error generating single certificate. Please check server logs.');
        }
    }
});

// POST /bulk-generate-certificates: Handles BULK certificate generation from a CSV file.
app.post('/bulk-generate-certificates', upload.single('bulkDataFile'), async (req, res) => {
    if (!req.file) { return res.status(400).send('No CSV file uploaded.'); }
    const records = [];
    const templatePath = path.join(__dirname, 'templates', 'iLab Certificate.pdf');
    const tempDir = path.join(__dirname, 'temp_bulk_output');

    try {
        await fs.mkdir(tempDir, { recursive: true });
        const csvStream = require('stream').Readable.from(req.file.buffer.toString('utf8'));
        csvStream.pipe(csv({ mapHeaders: ({ header }) => header.trim() }))
            .on('data', (data) => records.push(data))
            .on('end', async () => {
                if (records.length === 0) { return res.status(400).send('CSV file is empty or not properly formatted.'); }
                const zipFileName = `certificates_bulk_${Date.now()}.zip`;
                const zipFilePath = path.join(tempDir, zipFileName);
                const outputZipStream = fsSync.createWriteStream(zipFilePath);
                const archive = archiver('zip', { zlib: { level: 9 } });
                archive.on('warning', (err) => { if (err.code === 'ENOENT') { console.warn('Archiver warning:', err); } else { throw err; } });
                archive.on('error', (err) => { throw err; });
                archive.pipe(outputZipStream);
                for (let i = 0; i < records.length; i++) {
                    const { userName, trainingType, date } = records[i];
                    if (!userName || !trainingType || !date) {
                        console.warn(`Skipping record ${i + 1}: Missing data`);
                        archive.append(`Skipped record ${i + 1}: Missing data\n`, { name: `error_log_${i+1}.txt` });
                        continue;
                    }
                    try {
                        const pdfBytes = await generateCertificatePdfBytes(userName, trainingType, date, templatePath);
                        const sName = String(userName).replace(/[^a-z0-9]/gi, '_').toLowerCase();
                        archive.append(Buffer.from(pdfBytes), { name: `certificate_${sName}_${i+1}.pdf` });
                    } catch (pdfError) {
                        console.error(`Error PDF for ${userName}:`, pdfError);
                        archive.append(`Error for ${userName}: ${pdfError.message}\n`, { name: `error_log_${userName}_${i+1}.txt`});
                    }
                }
                await archive.finalize();
                outputZipStream.on('close', () => {
                    res.download(zipFilePath, zipFileName, async (dlErr) => {
                        if (dlErr) console.error('Zip DL Err:', dlErr);
                        try { await fs.unlink(zipFilePath); } catch (cupErr){ console.error('Zip Cleanup Err:', cupErr); }
                    });
                });
                outputZipStream.on('error', (zsErr) => { if(!res.headersSent) res.status(500).send('ZIP Stream Err.');});
            })
            .on('error', (pErr) => { if(!res.headersSent) res.status(500).send('CSV Parse Err.');});
    } catch (err) {
        console.error('Bulk Gen Err:', err);
        if (!res.headersSent) {
            if (err.message.includes('fonts')) res.status(500).send('Font file error.');
            else if (err.message.includes('fontkit')) res.status(500).send('Fontkit error.');
            else res.status(500).send('Bulk error.');
        }
    }
});

// Global error handler
app.use((err, req, res, next) => {
    if (err instanceof multer.MulterError) { return res.status(400).send(`Upload Err: ${err.message}`); }
    else if (err) {
        if (err.message === 'Only .csv files are allowed!') return res.status(400).send(err.message);
        console.error("Unhandled Err:", err);
        if (!res.headersSent) return res.status(500).send('Server Err.');
    }
    next();
});

// Start server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
    console.log('Template: templates/iLab Certificate.pdf');
    console.log('Font: fonts/GEORGIAB.TTF (ensure name matches in app.js)');
    console.log('CSV Headers: userName,trainingType,date');
});
