#!/usr/bin/env node

/**
 * CAA Affidavit Generator — DOCX Document Generation Script
 *
 * Usage:
 *   node generate.js '<JSON_DATA>'
 *
 * Input:
 *   A JSON string containing all 27 form fields + advocate field.
 *
 * Output:
 *   Prints a JSON object to stdout with the generated filename:
 *   { "success": true, "filename": "SUSANTA_NATH_20260712" }
 *
 * Dependencies:
 *   - docxtemplater (DOCX template engine)
 *   - pizzip (ZIP file handling for DOCX)
 */

'use strict';

const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

// ──────────────────────────────────────────────────────────
// Configuration
// ──────────────────────────────────────────────────────────

/** Directory containing template.docx */
const TEMPLATES_DIR = process.env.TEMPLATES_DIR || path.resolve(__dirname, '..', 'templates');

/** Directory where generated files are saved */
const OUTPUT_DIR = process.env.OUTPUT_DIR || path.resolve(__dirname, '..', 'generated');

/** Template filename */
const TEMPLATE_FILE = 'template.docx';


// ──────────────────────────────────────────────────────────
// Required Fields
// ──────────────────────────────────────────────────────────

const REQUIRED_FIELDS = [
    'applicant_name', 'guardian_type', 'father_name',
    'india_village', 'india_po', 'india_ps',
    'india_district', 'india_pin', 'india_state',
    'bd_village', 'bd_po', 'bd_ps', 'bd_district',
    'entry_date', 'verification_date', 'advocate',
    'witness_name', 'witness_guardian_type', 'witness_father',
    'witness_age', 'witness_occupation',
    'witness_village', 'witness_po', 'witness_ps',
    'witness_district', 'witness_pin', 'witness_state'
];


// ──────────────────────────────────────────────────────────
// Helper Functions
// ──────────────────────────────────────────────────────────

/**
 * Generate a filename from the applicant name and current date.
 * Format: APPLICANT_NAME_YYYYMMDD
 * Example: SUSANTA_NATH_20260712
 *
 * @param {string} applicantName - The applicant's full name
 * @returns {string} The sanitized filename (without extension)
 */
function generateFilename(applicantName) {
    // Sanitize the name: uppercase, replace spaces with underscores, remove special chars
    const sanitizedName = applicantName
        .toUpperCase()
        .trim()
        .replace(/\s+/g, '_')
        .replace(/[^A-Z0-9_]/g, '');

    // Current date in YYYYMMDD format
    const now = new Date();
    const dateStr = now.getFullYear().toString() +
        String(now.getMonth() + 1).padStart(2, '0') +
        String(now.getDate()).padStart(2, '0');

    return `${sanitizedName}_${dateStr}`;
}

/**
 * Validate that all required fields are present and non-empty.
 *
 * @param {Object} data - The form data object
 * @returns {{ valid: boolean, missing: string[] }}
 */
function validateData(data) {
    const missing = [];

    for (const field of REQUIRED_FIELDS) {
        if (!data[field] || String(data[field]).trim() === '') {
            missing.push(field);
        }
    }

    return {
        valid: missing.length === 0,
        missing: missing
    };
}

/**
 * Output result as JSON to stdout and exit.
 *
 * @param {boolean} success
 * @param {Object} data - Additional data to include in the response
 * @param {number} exitCode - Process exit code (0 for success, 1 for error)
 */
function outputResult(success, data, exitCode = 0) {
    const result = { success, ...data };
    process.stdout.write(JSON.stringify(result));
    process.exit(exitCode);
}


// ──────────────────────────────────────────────────────────
// Main Function
// ──────────────────────────────────────────────────────────

/**
 * Read all data from standard input.
 * @returns {Promise<string>}
 */
function readStdin() {
    return new Promise((resolve, reject) => {
        let data = '';
        process.stdin.setEncoding('utf8');
        process.stdin.on('data', chunk => {
            data += chunk;
        });
        process.stdin.on('end', () => {
            resolve(data);
        });
        process.stdin.on('error', err => {
            reject(err);
        });
    });
}

async function main() {
    // 1. Parse input — supports CLI argument, --file flag, or stdin
    let jsonStr;

    if (process.argv[2] === '--file' && process.argv[3]) {
        // Read JSON from a file: node generate.js --file /path/to/data.json
        try {
            jsonStr = fs.readFileSync(process.argv[3], 'utf8').trim();
        } catch (err) {
            outputResult(false, { message: `Failed to read input file: ${err.message}` }, 1);
        }
    } else if (process.argv[2]) {
        // Direct CLI argument: node generate.js '{"field":"value"}'
        jsonStr = process.argv[2];
    } else {
        // Stdin: cat data.json | node generate.js
        try {
            jsonStr = await readStdin();
            jsonStr = jsonStr.trim();
        } catch (err) {
            outputResult(false, { message: `Failed to read stdin: ${err.message}` }, 1);
        }
    }

    if (!jsonStr || jsonStr === '') {
        outputResult(false, { message: 'No input provided via CLI argument, --file, or stdin' }, 1);
    }

    let data;
    try {
        data = JSON.parse(jsonStr);
    } catch (err) {
        outputResult(false, { message: `Invalid JSON input: ${err.message}` }, 1);
    }

    // 2. Validate required fields
    const validation = validateData(data);
    if (!validation.valid) {
        outputResult(false, {
            message: `Missing required fields: ${validation.missing.join(', ')}`
        }, 1);
    }

    // 3. Ensure output directory exists
    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    }

    // 4. Load the DOCX template
    const templatePath = path.join(TEMPLATES_DIR, TEMPLATE_FILE);

    if (!fs.existsSync(templatePath)) {
        outputResult(false, {
            message: `Template not found at: ${templatePath}`
        }, 1);
    }

    let templateContent;
    try {
        templateContent = fs.readFileSync(templatePath, 'binary');
    } catch (err) {
        outputResult(false, {
            message: `Failed to read template: ${err.message}`
        }, 1);
    }

    // 5. Create Docxtemplater instance
    let doc;
    try {
        const zip = new PizZip(templateContent);
        doc = new Docxtemplater(zip, {
            // Keep empty paragraphs intact (important for legal formatting)
            paragraphLoop: true,
            // Preserve line breaks
            linebreaks: true,
            // Don't throw on missing tags — fill with empty string instead
            nullGetter: function () {
                return '';
            }
        });
    } catch (err) {
        outputResult(false, {
            message: `Failed to parse template DOCX: ${err.message}`
        }, 1);
    }

    // 6. Set template data
    //    All placeholders in the template are {key} format.
    //    Docxtemplater replaces them with the corresponding values.
    try {
        doc.render(data);
    } catch (err) {
        // Docxtemplater error — likely a template syntax issue
        const errorMessages = err.properties && err.properties.errors
            ? err.properties.errors.map(e => e.message).join('; ')
            : err.message;

        outputResult(false, {
            message: `Template rendering failed: ${errorMessages}`
        }, 1);
    }

    // 7. Generate output buffer
    const outputBuffer = doc.getZip().generate({
        type: 'nodebuffer',
        // Use DEFLATE compression for smaller file size
        compression: 'DEFLATE'
    });

    // 8. Save the generated DOCX
    const filename = generateFilename(data.applicant_name);
    const outputPath = path.join(OUTPUT_DIR, `${filename}.docx`);

    try {
        fs.writeFileSync(outputPath, outputBuffer);
    } catch (err) {
        outputResult(false, {
            message: `Failed to save generated DOCX: ${err.message}`
        }, 1);
    }

    // 9. Output success with filename
    outputResult(true, {
        filename: filename,
        path: outputPath
    });
}


// ──────────────────────────────────────────────────────────
// Execute
// ──────────────────────────────────────────────────────────
main();
