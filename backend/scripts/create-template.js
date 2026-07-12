#!/usr/bin/env node

/**
 * create-template.js
 * 
 * One-time script that converts the original CAA DEMO.docx into a
 * Docxtemplater-ready template.docx by replacing all dynamic values
 * with {placeholder} tags.
 * 
 * Usage:
 *   1. npm install pizzip   (in this directory)
 *   2. node create-template.js
 * 
 * The script handles text that is split across multiple XML runs
 * (a common issue in Word documents due to revision tracking).
 */

'use strict';

const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');

// ──────────────────────────────────────────────────────────
// Configuration
// ──────────────────────────────────────────────────────────

/** Path to the original DOCX file */
const SOURCE_FILE = path.resolve('C:\\Users\\User\\Desktop\\CAA\\CAA DEMO.docx');

/** Path to save the generated template */
const OUTPUT_FILE = path.resolve(__dirname, '..', 'templates', 'template.docx');

/**
 * Ordered list of text replacements.
 * 
 * CRITICAL: Order matters!
 * - Names are replaced first (they're unique strings)
 * - Guardian type (S/O) is replaced AFTER names (uses replaced names as context)
 * - Addresses use prefix labels (VILL-, PO-, PS-) as context to avoid ambiguity
 * - "West Bengal" for witness is replaced LAST (after applicant's "State –West Bengal" is done)
 * 
 * Each entry: [searchText, replaceText]
 * 
 * NOTE on dashes: Word uses U+2013 (EN DASH) for "–". We try both
 * the en dash and regular hyphen-minus for compatibility.
 */
const REPLACEMENTS = [
    // ═══ ROUND 1: Unique names (no ambiguity) ═══
    // Replace longer/more specific names first
    ['DURGAPADA DEY', '{witness_father}'],
    ['SUBASH NATH', '{father_name}'],
    ['SUSANTA NATH', '{applicant_name}'],
    ['SADHAN DEY', '{witness_name}'],

    // ═══ ROUND 2: Guardian type ═══
    // After name replacement, text looks like:
    //   Applicant: "{applicant_name}, S/O-{father_name}"   (no space after dash)
    //   Witness:   "{witness_name}, S/O- {witness_father}"  (space after dash)
    // Using the replaced names as context ensures correct matching
    ['S/O-{father_name}', '{guardian_type}-{father_name}'],
    ['S/O- {witness_father}', '{witness_guardian_type}- {witness_father}'],

    // ═══ ROUND 3: Applicant India address ═══
    // These use the applicant's address format (no periods in labels)
    ['VILL-MADHYA HARIA', 'VILL-{india_village}'],
    ['PO-BANIPUR', 'PO-{india_po}'],
    ['PS-HABRA', 'PS-{india_ps}'],
    // "NORTH 24 PARGANAS" (all caps) = applicant only
    // "North 24 Parganas" (title case) = witness/header — different string
    ['NORTH 24 PARGANAS', '{india_district}'],
    // "PIN-743233" (no space) = applicant; "PIN- 743233" (space) = witness
    ['PIN-743233', 'PIN-{india_pin}'],

    // ═══ ROUND 4: Bangladesh address ═══
    // BD village and PO can have same name — use prefix to distinguish
    ['VILL-SASHI BHUSAN', 'VILL-{bd_village}'],
    ['PO- SASHI BHUSAN', 'PO- {bd_po}'],
    ['PS- CHARFATION', 'PS- {bd_ps}'],
    ['DIST-BHOLA', 'DIST-{bd_district}'],

    // ═══ ROUND 5: Witness address ═══
    // Witness uses different label format (with periods: P.O.-, P.S.-)
    ['VILL-K.N ROY ROAD', 'VILL-{witness_village}'],
    ['P.O.- BANIPUR', 'P.O.- {witness_po}'],
    ['P.S.-HABRA', 'P.S.-{witness_ps}'],
    // "PIN- 743233" (with space) = witness only
    ['PIN- 743233', 'PIN- {witness_pin}'],

    // ═══ ROUND 6: Witness fields ═══
    ['about 52', 'about {witness_age}'],

    // ═══ ROUND 7: Dates ═══
    ['15/12/2012', '{entry_date}'],
    ['10/07/2026', '{verification_date}'],

    // ═══ ROUND 8: Advocate signature ═══
    ['Debabrata Sarkar, Adv, F/1675/856/2020', '{advocate}'],
];

/**
 * Replacements that use en dash (–) or em dash (—) characters.
 * We try multiple dash variants for compatibility since Word
 * may use different Unicode dash characters.
 */
const DASH_REPLACEMENTS = [
    // Applicant state: "State –West Bengal" → "State –{india_state}"
    { prefix: 'State ', suffix: 'West Bengal', replace: 'State \u2013{india_state}', field: 'india_state' },
    // Witness district: "DIST. – North 24 Parganas" → "DIST. – {witness_district}"
    { prefix: 'DIST. ', suffix: ' North 24 Parganas', replace: 'DIST. \u2013 {witness_district}', field: 'witness_district' },
    // Witness occupation: "Occupation – BUSINESS" → "Occupation – {witness_occupation}"
    { prefix: 'Occupation ', suffix: ' BUSINESS', replace: 'Occupation \u2013 {witness_occupation}', field: 'witness_occupation' },
];

/**
 * Late-stage replacements that must happen after all others.
 * "West Bengal" in witness context (after applicant's "State –West Bengal" is already replaced)
 */
const LATE_REPLACEMENTS = [
    ['West Bengal', '{witness_state}'],
];


// ──────────────────────────────────────────────────────────
// Core: Cross-Run Text Replacement
// ──────────────────────────────────────────────────────────

/**
 * Replace ALL occurrences of searchText with replaceText in the XML,
 * handling text that is split across multiple <w:t> elements.
 *
 * Algorithm:
 * 1. Find all <w:t> elements and their text content
 * 2. Concatenate all text to form a "text stream"
 * 3. Find searchText in the text stream
 * 4. Map the found position back to the specific <w:t> elements
 * 5. Modify those elements: put replacement in the first, clear the rest
 * 6. Repeat until no more occurrences
 *
 * @param {string} xml - The raw document.xml content
 * @param {string} searchText - The text to find
 * @param {string} replaceText - The replacement text
 * @returns {string} Modified XML
 */
function replaceTextAcrossRuns(xml, searchText, replaceText) {
    let result = xml;
    let safety = 0;
    const maxIterations = 100; // prevent infinite loops

    while (safety++ < maxIterations) {
        const replaced = replaceSingleOccurrence(result, searchText, replaceText);
        if (replaced === result) break; // no more occurrences
        result = replaced;
    }

    return result;
}

/**
 * Replace the FIRST occurrence of searchText in the XML's text content.
 *
 * @param {string} xml - The raw XML
 * @param {string} searchText - Text to find
 * @param {string} replaceText - Replacement text
 * @returns {string} Modified XML, or original if not found
 */
function replaceSingleOccurrence(xml, searchText, replaceText) {
    // 1. Extract all <w:t> elements with positions
    const textNodeRegex = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
    const nodes = [];
    let match;

    while ((match = textNodeRegex.exec(xml)) !== null) {
        nodes.push({
            xmlStart: match.index,
            xmlEnd: match.index + match[0].length,
            fullMatch: match[0],
            text: match[1],
            textStart: 0, // computed below
            textEnd: 0    // computed below
        });
    }

    if (nodes.length === 0) return xml;

    // 2. Build text position mapping
    let textPos = 0;
    for (const node of nodes) {
        node.textStart = textPos;
        node.textEnd = textPos + node.text.length;
        textPos += node.text.length;
    }

    // 3. Concatenate all text and find searchText
    const fullText = nodes.map(n => n.text).join('');
    const searchIdx = fullText.indexOf(searchText);
    if (searchIdx === -1) return xml; // not found

    const searchEnd = searchIdx + searchText.length;

    // 4. Find which nodes overlap with the search range
    const affected = nodes.filter(n =>
        n.textStart < searchEnd && n.textEnd > searchIdx
    );

    if (affected.length === 0) return xml;

    // 5. Compute new text for each affected node
    const firstNode = affected[0];
    const lastNode = affected[affected.length - 1];

    // Text before the match in the first node
    const prefixLen = searchIdx - firstNode.textStart;
    const prefix = firstNode.text.substring(0, prefixLen);

    // Text after the match in the last node
    const suffixStart = searchEnd - lastNode.textStart;
    const suffix = lastNode.text.substring(suffixStart);

    // 6. Rebuild XML (process from last to first to preserve positions)
    let result = xml;

    for (let i = affected.length - 1; i >= 0; i--) {
        const node = affected[i];
        let newText;

        if (affected.length === 1) {
            // Only one node affected: prefix + replacement + suffix
            newText = prefix + replaceText + suffix;
        } else if (i === 0) {
            // First node: prefix + replacement
            newText = prefix + replaceText;
        } else if (i === affected.length - 1) {
            // Last node: suffix only
            newText = suffix;
        } else {
            // Middle nodes: empty
            newText = '';
        }

        // Rebuild the <w:t> element with new text
        const newElement = rebuildTextElement(node.fullMatch, newText);
        result = result.substring(0, node.xmlStart) + newElement + result.substring(node.xmlEnd);
    }

    return result;
}

/**
 * Rebuild a <w:t> element with new text content,
 * preserving the original tag attributes.
 * Adds xml:space="preserve" if the new text has leading/trailing whitespace.
 *
 * @param {string} originalElement - The original <w:t>...</w:t> string
 * @param {string} newText - The new text content
 * @returns {string} Rebuilt element
 */
function rebuildTextElement(originalElement, newText) {
    // Parse the opening tag and closing tag
    const tagMatch = originalElement.match(/^(<w:t(?:\s[^>]*)?>)([\s\S]*)(<\/w:t>)$/);
    if (!tagMatch) return originalElement;

    let openTag = tagMatch[1];
    const closeTag = tagMatch[3];

    // Add xml:space="preserve" if new text has leading/trailing whitespace
    const needsPreserve = /^\s/.test(newText) || /\s$/.test(newText);
    if (needsPreserve && !openTag.includes('xml:space="preserve"')) {
        openTag = '<w:t xml:space="preserve">';
    }

    return openTag + newText + closeTag;
}


// ──────────────────────────────────────────────────────────
// Dash-aware Replacement
// ──────────────────────────────────────────────────────────

/**
 * Handle replacements that involve dash characters (en dash, em dash, etc.)
 * Tries multiple dash variants to find the right one.
 *
 * @param {string} xml - The XML content
 * @param {object} dashRep - Dash replacement config
 * @returns {{ xml: string, found: boolean }}
 */
function replaceDashPattern(xml, dashRep) {
    // Try different dash characters
    const dashes = ['\u2013', '\u2014', '\u2012', '-', '\u2015'];

    for (const dash of dashes) {
        const searchText = dashRep.prefix + dash + dashRep.suffix;
        const before = xml;
        const after = replaceTextAcrossRuns(xml, searchText, dashRep.replace);
        if (after !== before) {
            return { xml: after, found: true, dash: dash };
        }
    }

    return { xml: xml, found: false, dash: null };
}


// ──────────────────────────────────────────────────────────
// Verification
// ──────────────────────────────────────────────────────────

/**
 * Extract all text from the XML and check that no original
 * dynamic values remain (they should all be replaced by placeholders).
 */
function verifyTemplate(xml) {
    // Extract all text
    const textRegex = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
    let fullText = '';
    let match;
    while ((match = textRegex.exec(xml)) !== null) {
        fullText += match[1];
    }

    // Check for values that should have been replaced
    const shouldNotExist = [
        'SUSANTA NATH', 'SUBASH NATH', 'SADHAN DEY', 'DURGAPADA DEY',
        'MADHYA HARIA', 'SASHI BHUSAN', 'CHARFATION', 'BHOLA',
        'K.N ROY ROAD', '15/12/2012', 'Debabrata Sarkar'
    ];

    const remaining = shouldNotExist.filter(v => fullText.includes(v));

    if (remaining.length > 0) {
        console.warn('\n⚠ WARNING: The following values were NOT replaced:');
        remaining.forEach(v => console.warn(`  - "${v}"`));
    }

    // Check that placeholders exist
    const expectedPlaceholders = [
        '{applicant_name}', '{father_name}', '{guardian_type}',
        '{india_village}', '{india_po}', '{india_ps}', '{india_district}',
        '{india_pin}', '{india_state}',
        '{bd_village}', '{bd_po}', '{bd_ps}', '{bd_district}',
        '{entry_date}', '{verification_date}', '{advocate}',
        '{witness_name}', '{witness_father}', '{witness_guardian_type}',
        '{witness_age}', '{witness_occupation}',
        '{witness_village}', '{witness_po}', '{witness_ps}',
        '{witness_district}', '{witness_pin}', '{witness_state}'
    ];

    const missingPlaceholders = expectedPlaceholders.filter(p => !fullText.includes(p));

    if (missingPlaceholders.length > 0) {
        console.warn('\n⚠ WARNING: The following placeholders are MISSING:');
        missingPlaceholders.forEach(p => console.warn(`  - ${p}`));
    }

    return {
        unreplaced: remaining,
        missing: missingPlaceholders,
        fullText: fullText
    };
}


// ──────────────────────────────────────────────────────────
// Main
// ──────────────────────────────────────────────────────────

function main() {
    console.log('╔══════════════════════════════════════════════╗');
    console.log('║   CAA Affidavit — Template Creator           ║');
    console.log('╚══════════════════════════════════════════════╝\n');

    // 1. Check source file
    if (!fs.existsSync(SOURCE_FILE)) {
        console.error(`✗ Source file not found: ${SOURCE_FILE}`);
        console.error('  Please ensure the original CAA DEMO.docx is at the specified path.');
        process.exit(1);
    }
    console.log(`Source: ${SOURCE_FILE}`);

    // 2. Read the DOCX
    const content = fs.readFileSync(SOURCE_FILE, 'binary');
    const zip = new PizZip(content);

    // 3. Get document.xml
    let docXml = zip.file('word/document.xml').asText();
    console.log(`Document XML loaded (${docXml.length} characters)\n`);

    // 4. Apply standard replacements
    console.log('── Standard Replacements ──');
    let totalReplaced = 0;

    for (const [search, replace] of REPLACEMENTS) {
        const before = docXml;
        docXml = replaceTextAcrossRuns(docXml, search, replace);
        if (docXml !== before) {
            console.log(`  ✓ "${search}" → "${replace}"`);
            totalReplaced++;
        } else {
            console.log(`  ✗ NOT FOUND: "${search}"`);
        }
    }

    // 5. Apply dash-aware replacements
    console.log('\n── Dash-Character Replacements ──');

    for (const dashRep of DASH_REPLACEMENTS) {
        const result = replaceDashPattern(docXml, dashRep);
        docXml = result.xml;
        if (result.found) {
            const charName = result.dash === '\u2013' ? 'en-dash' :
                             result.dash === '\u2014' ? 'em-dash' :
                             result.dash === '-' ? 'hyphen' : `U+${result.dash.charCodeAt(0).toString(16)}`;
            console.log(`  ✓ ${dashRep.field} (matched with ${charName})`);
            totalReplaced++;
        } else {
            console.log(`  ✗ NOT FOUND: ${dashRep.field} (tried all dash variants)`);
        }
    }

    // 6. Apply late-stage replacements
    console.log('\n── Late-Stage Replacements ──');

    for (const [search, replace] of LATE_REPLACEMENTS) {
        const before = docXml;
        docXml = replaceTextAcrossRuns(docXml, search, replace);
        if (docXml !== before) {
            console.log(`  ✓ "${search}" → "${replace}"`);
            totalReplaced++;
        } else {
            console.log(`  ✗ NOT FOUND: "${search}"`);
        }
    }

    console.log(`\nTotal replacements: ${totalReplaced}`);

    // 7. Verify the template
    console.log('\n── Verification ──');
    const verification = verifyTemplate(docXml);

    if (verification.unreplaced.length === 0 && verification.missing.length === 0) {
        console.log('  ✓ All dynamic values replaced');
        console.log('  ✓ All placeholders present');
    }

    // 8. Save the template
    zip.file('word/document.xml', docXml);
    const outputBuffer = zip.generate({
        type: 'nodebuffer',
        compression: 'DEFLATE'
    });

    const outputDir = path.dirname(OUTPUT_FILE);
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    fs.writeFileSync(OUTPUT_FILE, outputBuffer);
    console.log(`\n✓ Template saved to: ${OUTPUT_FILE}`);
    console.log('\nDone! You can now use this template with generate.js');
}

main();
