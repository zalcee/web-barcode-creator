// Required modules
const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const bwipjs = require('bwip-js');
const fs = require('fs');
const PDFDocument = require('pdfkit');
const path = require('path');

// Initialize Express app
const app = express();

// Configure Multer for file uploads, storing files in 'uploads/' directory
const upload = multer({ dest: 'uploads/' });

// Create 'barcodes' directory if it doesn't exist
if (!fs.existsSync('barcodes')) {
    fs.mkdirSync('barcodes');
}

// Serve the HTML form for file upload
app.get('/', (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Barcode PDF Generator</title>
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
            <script src="https://cdn.tailwindcss.com"></script>
            <style>
                body {
                    font-family: 'Inter', sans-serif;
                    background-color: #f0f2f5;
                    display: flex;
                    flex-direction: column;
                    justify-content: center;
                    align-items: center;
                    min-height: 100vh;
                    margin: 0;
                    color: #333;
                }
                .container {
                    background-color: #ffffff;
                    padding: 2.5rem;
                    border-radius: 0.75rem;
                    box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
                    text-align: center;
                    max-width: 80%;
                    width: 400px; /* Slightly wider container for upload area */
                    // margin-top: 1.5rem;
                }

                #info-container {
                    background-color: #e0f2fe;
                    color: #0c4a6e;
                    padding: 1.5rem 2rem;
                    border-radius: 0.75rem;
                    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
                    text-align: center;
                    max-width: 90%;
                    width: 500px;
                    margin-bottom: 1rem;
                    border: 1px solid #90cdf4;
                }

                #info-container p {
                    font-size: 1rem;
                    margin-bottom: 0.75rem;
                    line-height: 1.5;
                }

                #info-container a {
                    color: #1d4ed8;
                    font-weight: 600;
                    text-decoration: none;
                    transition: color 0.2s ease-in-out;
                }

                #info-container a:hover {
                    color: #1e3a8a;
                    text-decoration: underline;
                }

                h2 {
                    color: #333;
                    margin-bottom: 1.5rem;
                    font-size: 1.875rem;
                    font-weight: 700;
                }
                form {
                    display: flex;
                    flex-direction: column;
                    gap: 1.5rem;
                }

                /* Custom File Input and Drag-and-Drop Area */
                .file-upload-wrapper {
                    border: 2px dashed #cbd5e1; /* Dashed border for drop zone */
                    border-radius: 0.75rem;
                    padding: 2rem;
                    background-color: #f8fafc;
                    transition: all 0.2s ease-in-out;
                    display: flex;
                    flex-direction: column;
                    align-items: center;
                    justify-content: center;
                    cursor: pointer;
                    position: relative;
                }

                .file-upload-wrapper.drag-over {
                    border-color: #6366f1; /* Highlight on drag over */
                    background-color: #eef2ff; /* Lighter background on drag over */
                }

                .file-upload-wrapper input[type="file"] {
                    position: absolute;
                    width: 100%;
                    height: 100%;
                    top: 0;
                    left: 0;
                    opacity: 0;
                    cursor: pointer;
                }

                .file-upload-wrapper svg {
                    margin-bottom: 1rem;
                    color: #94a3b8; /* Icon color */
                }

                .file-upload-wrapper p {
                    font-size: 1rem;
                    color: #64748b;
                    margin-bottom: 0.5rem;
                }

                .file-upload-wrapper .file-name {
                    font-size: 0.9rem;
                    color: #333;
                    font-weight: 500;
                    margin-top: 0.75rem;
                    white-space: nowrap;
                    overflow: hidden;
                    text-overflow: ellipsis;
                    max-width: 90%;
                }

                .file-upload-wrapper .upload-button {
                    background-color: #4f46e5;
                    color: white;
                    padding: 0.65rem 1.25rem;
                    border-radius: 0.375rem;
                    font-weight: 500;
                    transition: background-color 0.2s ease-in-out;
                }

                .file-upload-wrapper .upload-button:hover {
                    background-color: #4338ca;
                }


                button[type="submit"] {
                    background-color: #10b981;
                    color: white;
                    padding: 0.85rem 1.75rem;
                    border: none;
                    border-radius: 0.5rem;
                    font-size: 1.125rem;
                    font-weight: 600;
                    cursor: pointer;
                    transition: background-color 0.2s ease-in-out, transform 0.1s ease-in-out, box-shadow 0.2s ease-in-out;
                    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
                }
                button[type="submit"]:hover {
                    background-color: #059669;
                    transform: translateY(-1px);
                    box-shadow: 0 6px 8px -1px rgba(0, 0, 0, 0.15), 0 3px 5px -1px rgba(0, 0, 0, 0.08);
                }
                button[type="submit"]:active {
                    transform: translateY(0);
                    box-shadow: none;
                }
            </style>
        </head>
        <body>
            <div id="info-container">
                <p>To generate your barcode PDF, first download the CSV template from the link below. Fill in your barcode data, then upload the CSV file here.</p>
                <a href="https://docs.google.com/spreadsheets/d/17tWHEbEDGNAB3X30sa2Ml1GXgM-17UHpwrWz4LIqdXI/edit?gid=0#gid=0" target="_blank" class="text-blue-700 hover:text-blue-900 font-semibold text-lg">
                    Download CSV Template Here
                </a>
            </div>
            <div class="container">
                <h2>Upload XLSX or CSV to Generate Barcode PDF</h2>
                <form action="/upload" method="post" enctype="multipart/form-data">
                    <div class="file-upload-wrapper" id="drop-zone">
                        <svg xmlns="http://www.w3.org/2000/svg" class="h-16 w-16" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="1">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                        </svg>
                        <p>Drag & drop your file here, or</p>
                        <span class="upload-button">Browse File</span>
                        <input type="file" name="file" id="file-input" accept=".xlsx, .csv" required />
                        <span class="file-name" id="file-name">No file chosen</span>
                    </div>

                    <!-- NEW: Add a way to select the template -->
                    <div style="margin-top: 1rem; text-align: left;">
                        <label for="templateSelect" style="display: block; margin-bottom: 0.5rem; color: #333; font-weight: 600;">Select Barcode Template:</label>
                        <select name="template" id="templateSelect" style="width: 100%; padding: 0.75rem; border: 1px solid #d1d5db; border-radius: 0.5rem; background-color: #f9fafb;">
                            <option value="template1">KCC (Default)</option>
                            <option value="template2">SM</option>
                            
                        </select>
                    </div>

                    <button type="submit">Generate PDF</button>
                </form>
            </div>

            <script>
                const dropZone = document.getElementById('drop-zone');
                const fileInput = document.getElementById('file-input');
                const fileNameSpan = document.getElementById('file-name');

                // Prevent default drag behaviors
                ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                    dropZone.addEventListener(eventName, preventDefaults, false);
                    document.body.addEventListener(eventName, preventDefaults, false); // Prevent drop on entire body
                });

                function preventDefaults (e) {
                    e.preventDefault();
                    e.stopPropagation();
                }

                // Highlight drop zone when dragging over
                ['dragenter', 'dragover'].forEach(eventName => {
                    dropZone.addEventListener(eventName, () => dropZone.classList.add('drag-over'), false);
                });

                ['dragleave', 'drop'].forEach(eventName => {
                    dropZone.addEventListener(eventName, () => dropZone.classList.remove('drag-over'), false);
                });

                // Handle dropped files
                dropZone.addEventListener('drop', handleDrop, false);

                function handleDrop(e) {
                    const dt = e.dataTransfer;
                    const files = dt.files;

                    fileInput.files = files; // Assign dropped files to the input
                    updateFileName();
                }

                // Update file name display when file is selected via click or drop
                fileInput.addEventListener('change', updateFileName);

                function updateFileName() {
                    if (fileInput.files.length > 0) {
                        fileNameSpan.textContent = fileInput.files[0].name;
                    } else {
                        fileNameSpan.textContent = 'No file chosen';
                    }
                }
            </script>
        </body>
        </html>
    `);
});

// Helper function for drawing Template 1 (your original layout)
function drawTemplate1(doc, item, x, y, labelWidth, labelHeight) {
    doc.lineWidth(0.05).rect(x, y, labelWidth, labelHeight).stroke(); // Border for label

    const KCC_BW_Logo_Path = 'KCC_Mall_logo_bw.PNG'; // Corrected case
    const KCC_Logo_Path = 'KCC_Mall_logo.png';       // Assuming this is correct case

    // Diagnostic logs for image paths
    console.log(`[Template 1] Checking for ${KCC_BW_Logo_Path}: ${fs.existsSync(KCC_BW_Logo_Path)}`);
    console.log(`[Template 1] Checking for ${KCC_Logo_Path}: ${fs.existsSync(KCC_Logo_Path)}`);

    // Draw KCC logo or text
    if (fs.existsSync(KCC_BW_Logo_Path)) {
        console.log(`[Template 1] Embedding ${KCC_BW_Logo_Path}`);
        try {
            doc.image(KCC_BW_Logo_Path, x + 18, y + 5, { width: 32, height: 18 });
        } catch (imageError) {
            console.error(`[Template 1] Error embedding ${KCC_BW_Logo_Path}:`, imageError);
        }
    } else if (fs.existsSync(KCC_Logo_Path)) { // Fallback to the other logo if BW is not found
        console.log(`[Template 1] Embedding fallback ${KCC_Logo_Path}`);
        try {
            doc.image(KCC_Logo_Path, x + 18, y + 5, { width: 32, height: 18 });
        } catch (imageError) {
            console.error(`[Template 1] Error embedding fallback ${KCC_Logo_Path}:`, imageError);
        }
    }
    else {
        console.log(`[Template 1] Neither logo found. Drawing 'KCC' text.`);
        doc.fontSize(6).text('KCC', x + 5, y + 5);
    }

    // Draw SUPP and SKU information
    doc.font('Helvetica').fontSize(6).text(`${item.supp}\n${item.sku}`, x + 30, y + 5, { width: labelWidth - 35, lineGap: 0, align: 'right' });

    // Draw STOCK
    doc.font('Helvetica').fontSize(7).text(item.stock, x + 2, y + 77, { width: labelWidth - 4, align: 'right' });

    // Draw "HAVAIANAS" (assuming this is a fixed brand name)
    doc.font('Helvetica-Bold').fontSize(7).text('HAVAIANAS', x + 2, y + 28, { width: labelWidth - 4, align: 'center' });

    // Draw barcode image
    doc.image(item.barcodePath, x + 5, y + 35, { width: labelWidth - 10, height: 23 });

    // Draw barcode number
    doc.font('Helvetica').fontSize(7).text(item.barcode, x + 2, y + 60, { width: labelWidth - 4, align: 'center' });

    // Draw description
    doc.font('Helvetica-Bold').fontSize(4.5).text(item.description, x , y + 69, { width: labelWidth , align: 'center', ellipsis: true });

    // Draw SRP (Suggested Retail Price) formatted as currency
    doc.font('Helvetica-Bold').fontSize(10).text(`P ${item.srp.toLocaleString(undefined, { minimumFractionDigits: 2 })}`, x + 2, y + 85, { width: labelWidth - 4, align: 'center' });
}

// Helper function for drawing Template 2 (Example: a smaller label with only barcode and SRP)
function drawTemplate2(doc, item, x, y, labelWidth, labelHeight) {

    const SM_LOGO = "SM-DEPT.png";
    if (fs.existsSync(SM_LOGO)) {
        console.log(`[Template 2] Embedding ${SM_LOGO}`);
        try {
            doc.image(SM_LOGO, x + 18, y - 25, { width: 70, align: 'center' });
        } catch (imageError) {
            console.error(`[Template 2] Error embedding ${SM_LOGO}:`, imageError);
        }
    } else if (fs.existsSync(KCC_Logo_Path)) { // Fallback to the other logo if BW is not found
        console.log(`[Template 1] Embedding fallback ${SM_LOGO}`);
        try {
            doc.image(SM_LOGO, x + 18, y - 25, { width: 70 });
        } catch (imageError) {
            console.error(`[Template 2] Error embedding fallback ${SM_LOGO}:`, imageError);
        }
    }
    else {
        console.log(`[Template 1] Neither logo found. Drawing 'KCC' text.`);
        doc.fontSize(6).text('SM DEPT', x + 5, y + 5);
    }



    doc.lineWidth(0.05).rect(x, y, labelWidth, labelHeight).stroke(); // Border for label

    doc.fontSize(7).text(item.subcla, x + 10, y + 20, {width: labelWidth -10, align: 'center'});
    doc.font('Helvetica-Bold').fontSize(5).text(item.description, x + 5, y + 65, { width: labelWidth - 10, align:'center' });

    doc.image(item.barcodePath, x + 10, y + 30, { width: labelWidth - 20, height: 20 });
    doc.fontSize(8).text(item.barcode, x + 5, y + 55, { width: labelWidth - 10, align: 'center' });
    doc.font('Helvetica-Bold').fontSize(12).text(`P ${item.srp.toLocaleString(undefined, { minimumFractionDigits: 2 })}`, x + 5, y + 85, { width: labelWidth - 10, align: 'center' });
}

// Helper function for drawing Template 3 (Example: Simple Product Info with larger text)
function drawTemplate3(doc, item, x, y, labelWidth, labelHeight) {
    doc.lineWidth(0.05).rect(x, y, labelWidth, labelHeight).stroke(); // Border for label

    doc.font('Helvetica-Bold').fontSize(10).text(item.description, x + 5, y + 5, { width: labelWidth - 10, align: 'center' });
    doc.font('Helvetica').fontSize(8).text(`SKU: ${item.sku}`, x + 5, y + 20, { width: labelWidth - 10, align: 'center' });

    doc.image(item.barcodePath, x + 15, y + 35, { width: labelWidth - 30, height: 25 });
    doc.font('Helvetica').fontSize(9).text(item.barcode, x + 5, y + 65, { width: labelWidth - 10, align: 'center' });

    doc.font('Helvetica-Bold').fontSize(14).text(`SRP: P ${item.srp.toLocaleString(undefined, { minimumFractionDigits: 2 })}`, x + 5, y + 80, { width: labelWidth - 10, align: 'center' });
}

// Helper function for drawing Template 4 (Example: Large Barcode Focus)
function drawTemplate4(doc, item, x, y, labelWidth, labelHeight) {
    doc.lineWidth(0.1).rect(x, y, labelWidth, labelHeight).stroke(); // Border for label

    doc.font('Helvetica-Bold').fontSize(8).text(item.description, x + 5, y + 5, { width: labelWidth - 10, align: 'center' });

    doc.image(item.barcodePath, x + 5, y + 20, { width: labelWidth - 10, height: 40 }); // Larger barcode
    doc.font('Helvetica').fontSize(10).text(item.barcode, x + 5, y + 65, { width: labelWidth - 10, align: 'center' });

    doc.font('Helvetica-Bold').fontSize(12).text(`P ${item.srp.toLocaleString(undefined, { minimumFractionDigits: 2 })}`, x + 5, y + 85, { width: labelWidth - 10, align: 'center' });
}

// Helper function for drawing Template 5 (Example: Compact with all info)
function drawTemplate5(doc, item, x, y, labelWidth, labelHeight) {
    doc.lineWidth(0.05).rect(x, y, labelWidth, labelHeight).stroke(); // Border for label

    doc.font('Helvetica-Bold').fontSize(6).text(item.description, x + 2, y + 2, { width: labelWidth - 4, align: 'center' });
    doc.font('Helvetica').fontSize(5).text(`SUPP: ${item.supp} SKU: ${item.sku}`, x + 2, y + 12, { width: labelWidth - 4, align: 'center' });

    doc.image(item.barcodePath, x + 5, y + 20, { width: labelWidth - 10, height: 20 });
    doc.font('Helvetica').fontSize(6).text(item.barcode, x + 2, y + 42, { width: labelWidth - 4, align: 'center' });

    doc.font('Helvetica-Bold').fontSize(8).text(`P ${item.srp.toLocaleString(undefined, { minimumFractionDigits: 2 })}`, x + 2, y + 50, { width: labelWidth - 4, align: 'center' });
    doc.font('Helvetica').fontSize(5).text(`STOCK: ${item.stock}`, x + 2, y + 58, { width: labelWidth - 4, align: 'center' });
}


// Handle file upload and PDF generation
app.post('/upload', upload.single('file'), async (req, res) => {

    const barcodesDir = 'barcodes/';
    fs.readdir(barcodesDir, (err, files) => {
        if (err) {
            console.error('Error reading barcodes directory:', err);
        } else {
            for (const file of files) {
                fs.unlink(path.join(barcodesDir, file), err => {
                    if (err) {
                        console.error('Error deleting file:', file, err);
                    }
                });
            }
            console.log('Cleared existing barcode files in "barcodes" directory.');
        }
    });

    let workbook;
    let data;

    // Determine file type based on extension
    const fileExtension = path.extname(req.file.originalname).toLowerCase();

    try {
        if (fileExtension === '.xlsx') {
            // Read XLSX file
            workbook = XLSX.readFile(req.file.path);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            data = XLSX.utils.sheet_to_json(sheet);
        } else if (fileExtension === '.csv') {
            // Read CSV file. XLSX library can also parse CSV.
            const csvContent = fs.readFileSync(req.file.path, 'utf8');
            workbook = XLSX.read(csvContent, { type: 'string' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            data = XLSX.utils.sheet_to_json(sheet);
        } else {
            // If file type is not supported, send an error response
            return res.status(400).send('Unsupported file type. Please upload an XLSX or CSV file.');
        }
    } catch (error) {
        console.error("Error reading or parsing file:", error);
        return res.status(500).send('Error processing your file. Please ensure it is a valid XLSX or CSV.');
    }

    const labels = [];

    // Process each row to generate barcode images and collect label data
    for (const row of data) {
        // Skip rows without a BARCODE
        if (!row.BARCODE) continue;

        const barcodePath = `barcodes/${row.BARCODE}.png`;

        // Generate barcode image using bwip-js
        await new Promise((resolve, reject) => {
            bwipjs.toBuffer({
                bcid: 'code128', // Barcode type (Code 128)
                text: String(row.BARCODE), // Barcode data
                scale: 3, // Scaling factor
                height: 10, // Height of the barcode bars
                includetext: false, // Do not include human-readable text below barcode (we'll add it with PDFKit)
            }, (err, png) => {
                if (err) {
                    console.error(`Error generating barcode for ${row.BARCODE}:`, err);
                    return reject(err);
                }
                // Save the generated PNG image
                fs.writeFileSync(barcodePath, png);

                // Clean and parse SRP (Suggested Retail Price)
                // Ensure SRP is handled gracefully if it's missing or not a number
                const srpValue = row.SRP !== undefined && row.SRP !== null ? String(row.SRP) : '0';
                const cleanSRP = parseFloat(srpValue.replace(/[^\d.]/g, '')) || 0;


                // Add label data to the array
                labels.push({
                    supp: row.SUPP,
                    stock: row.STOCK,
                    sku: row.SKU,
                    barcode: row.BARCODE,
                    srp: cleanSRP,
                    description: row.DESC,
                    subcla: row.SUBCLA,
                    barcodePath // Path to the generated barcode image
                });
                resolve();
            });
        });
    }

    // Delete the uploaded file after processing
    fs.unlinkSync(req.file.path);

    // Get the selected template from the form data
    const selectedTemplate = req.body.template;
    console.log(`Selected Template: ${selectedTemplate}`); // Log which template was selected

    // Initialize PDFDocument with A4 size and margins
    const doc = new PDFDocument({ size: 'A4', margin: 30 });
    const filename = `Barcode_${Date.now()}.pdf`;

    // Set response headers for PDF download
    res.setHeader('Content-disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-type', 'application/pdf');

    // Pipe the PDF document to the response stream
    doc.pipe(res);

    // --- Template-specific Layout Configuration ---
    let labelsPerRow;
    let labelWidth;
    let labelHeight;
    let paddingX;
    let paddingY;

    switch (selectedTemplate) {
        case 'template1':
            labelsPerRow = 5;
            labelWidth = 100;
            labelHeight = 100;
            paddingX = 0;
            paddingY = 0;
            break;
        case 'template2': // Smaller label example
            labelsPerRow = 5;
            labelWidth = 100;
            labelHeight = 100;
            paddingX = 0;
            paddingY = 0;
            break;
        case 'template3': // Simple Product Info
            labelsPerRow = 4;
            labelWidth = 120;
            labelHeight = 100;
            paddingX = 10;
            paddingY = 10;
            break;
        case 'template4': // Large Barcode Focus
            labelsPerRow = 3;
            labelWidth = 160;
            labelHeight = 120;
            paddingX = 5;
            paddingY = 5;
            break;
        case 'template5': // Compact Template
            labelsPerRow = 6;
            labelWidth = 80;
            labelHeight = 65; // Smaller height
            paddingX = 2;
            paddingY = 2;
            break;
        default: // Default to template1 if nothing is selected or an unknown value
            labelsPerRow = 5;
            labelWidth = 100;
            labelHeight = 100;
            paddingX = 0;
            paddingY = 0;
            console.warn('Unknown template selected. Defaulting to Template 1.');
            break;
    }

    const startX = 30; // Starting X coordinate for the first label on a page
    const startY = 30; // Starting Y coordinate for the first label on a page

    // Dynamically calculate labels per column based on page height and label dimensions
    const pageHeight = doc.page.height;
    // Calculate available vertical space, accounting for top and bottom margins
    const availableHeight = pageHeight - startY - (doc.page.margins.bottom || 30);
    let labelsPerColumn = Math.floor(availableHeight / (labelHeight + paddingY));

    // Fallback to 1 label per column if calculation results in 0 (e.g., due to large labelHeight)
    if (labelsPerColumn === 0) {
        console.warn("Calculated labelsPerColumn is 0. Adjusting to 1 to prevent errors. Review labelHeight, paddingY, and page margins.");
        labelsPerColumn = 1;
    }
    // --- End PDF Layout Configuration ---

    // Iterate through each label data to draw it on the PDF
    labels.forEach((item, index) => {
        // Calculate the position of the current label within the current page
        const positionInPage = index % (labelsPerRow * labelsPerColumn);

        // Add a new page if it's the start of a new page block (and not the very first label)
        if (positionInPage === 0 && index !== 0) {
            doc.addPage();
        }

        // Calculate column and row index on the current page
        const col = positionInPage % labelsPerRow;
        const row = Math.floor(positionInPage / labelsPerRow);

        // Calculate absolute X and Y coordinates for the current label on the page
        const x = startX + col * (labelWidth + paddingX);
        const y = startY + row * (labelHeight + paddingY);

        // Call the appropriate drawing function based on selectedTemplate
        switch (selectedTemplate) {
            case 'template1':
                drawTemplate1(doc, item, x, y, labelWidth, labelHeight);
                break;
            case 'template2':
                drawTemplate2(doc, item, x, y, labelWidth, labelHeight);
                break;
            case 'template3':
                drawTemplate3(doc, item, x, y, labelWidth, labelHeight);
                break;
            case 'template4':
                drawTemplate4(doc, item, x, y, labelWidth, labelHeight);
                break;
            case 'template5':
                drawTemplate5(doc, item, x, y, labelWidth, labelHeight);
                break;
            default:
                drawTemplate1(doc, item, x, y, labelWidth, labelHeight); // Fallback
                break;
        }
    });

    // Finalize the PDF document
    doc.end();
});

// Start the server on port 8310
app.listen(8310, () => console.log('Server running at http://localhost:8310/'));
