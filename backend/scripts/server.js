const express = require('express');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { exec } = require('child_process');

const app = express();
app.use(express.json());

const GENERATED_DIR = '/app/generated';
if (!fs.existsSync(GENERATED_DIR)) {
    fs.mkdirSync(GENERATED_DIR, { recursive: true });
}

app.post('/generate', (req, res) => {
    try {
        const data = req.body;
        if (!data.applicant_name) {
            return res.status(400).json({ error: 'Missing applicant_name' });
        }

        // 1. Read Template Word Document
        const templatePath = '/app/templates/template.docx';
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        // 2. Render document with client data
        doc.render(data);

        // 3. Save generated Word document
        const timestamp = Date.now();
        const baseFilename = `${data.applicant_name.replace(/[^a-zA-Z0-9]/g, '_')}_${timestamp}`;
        const docxPath = path.join(GENERATED_DIR, `${baseFilename}.docx`);
        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        fs.writeFileSync(docxPath, buf);

        // 3b. Format the Word document using the Python script
        const formatScriptPath = path.join(__dirname, 'format_doc.py');
        exec(`python3 ${formatScriptPath} ${docxPath}`, (formatErr, formatStdout, formatStderr) => {
            if (formatErr) {
                console.error('Formatting script failed:', formatStderr || formatErr);
                // Continue to PDF generation even if formatting failed, as a fallback
            } else {
                console.log('Document formatted successfully:', formatStdout.trim());
            }

            // 4. Convert Word document to PDF using LibreOffice directly (now that Bookman font is installed)
            exec(`libreoffice --headless --convert-to pdf --outdir ${GENERATED_DIR} ${docxPath}`, (err) => {
                if (err) {
                    console.error('LibreOffice conversion failed:', err);
                    return res.status(500).json({ error: 'Failed to convert to PDF' });
                }

                // Return success with download names and URLs
                res.json({
                    success: true,
                    filename: baseFilename,
                    docx: 'https://n8n.srv1594654.hstgr.cloud/files/' + baseFilename + '.docx',
                    pdf: 'https://n8n.srv1594654.hstgr.cloud/files/' + baseFilename + '.pdf'
                });
            });
        });

    } catch (error) {
        console.error('Generation failed:', error);
        res.status(500).json({ error: error.message });
    }
});

const PORT = 5000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Affidavit API server running on http://0.0.0.0:${PORT}`);
});
