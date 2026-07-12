const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');

function inspectFile(filename) {
    const filePath = path.join(__dirname, '..', 'templates', filename);
    if (!fs.existsSync(filePath)) {
        console.log(`\nFile ${filename} not found.`);
        return;
    }

    const content = fs.readFileSync(filePath);
    const zip = new PizZip(content);
    const docXmlStr = zip.file('word/document.xml').asText();

    console.log(`\n================ ANALYZING ${filename.toUpperCase()} LAYOUT ================`);

    // Parse Section margins from w:pgMar
    const pgMarRegex = /<w:pgMar\s+([^>]+)\/>/g;
    let pgMarMatch;
    let sectionIndex = 0;
    while ((pgMarMatch = pgMarRegex.exec(docXmlStr)) !== null) {
        const attrsStr = pgMarMatch[1];
        const attrs = {};
        attrsStr.split(/\s+/).forEach(attr => {
            const parts = attr.split('=');
            if (parts.length === 2) {
                const key = parts[0].replace('w:', '');
                const val = parseInt(parts[1].replace(/"/g, ''), 10);
                attrs[key] = val;
            }
        });
        
        console.log(`\nSection ${sectionIndex} Margins (1 cm = 567 dxa):`);
        console.log(`  Top Margin:    ${attrs.top} dxa (${(attrs.top / 567).toFixed(2)} cm)`);
        console.log(`  Bottom Margin: ${attrs.bottom} dxa (${(attrs.bottom / 567).toFixed(2)} cm)`);
        console.log(`  Left Margin:   ${attrs.left} dxa (${(attrs.left / 567).toFixed(2)} cm)`);
        console.log(`  Right Margin:  ${attrs.right} dxa (${(attrs.right / 567).toFixed(2)} cm)`);
        sectionIndex++;
    }

    // Parse font size and line spacing properties of paragraphs
    const pRegex = /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g;
    let pMatch;
    let pIndex = 0;
    let nonElIndex = 0;

    console.log("\n--- Paragraph Analysis ---");
    while ((pMatch = pRegex.exec(docXmlStr)) !== null && pIndex < 100) {
        pIndex++;
        const pContent = pMatch[1];
        const tRegex = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
        let tMatch;
        let text = "";
        while ((tMatch = tRegex.exec(pContent)) !== null) {
            text += tMatch[1];
        }
        
        if (text.trim().length > 0) {
            nonElIndex++;
            console.log(`P #${pIndex}: "${text.trim().substring(0, 70)}"`);
            
            // Spacing
            const spacingRegex = /<w:spacing\s+([^>]+)\/>/;
            const spacingMatch = spacingRegex.exec(pContent);
            if (spacingMatch) {
                console.log(`  Spacing: ${spacingMatch[1]}`);
            }
            
            // Justification
            const jcRegex = /<w:jc\s+w:val="([^"]+)"\/>/;
            const jcMatch = jcRegex.exec(pContent);
            if (jcMatch) {
                console.log(`  Alignment: ${jcMatch[1]}`);
            }
            
            // Fonts
            const rFontsRegex = /<w:rFonts\s+([^>]+)\/>/g;
            let rFontsMatch;
            const fonts = new Set();
            while ((rFontsMatch = rFontsRegex.exec(pContent)) !== null) {
                fonts.add(rFontsMatch[1]);
            }
            if (fonts.size > 0) {
                console.log(`  Fonts: ${Array.from(fonts).join(' | ')}`);
            }
        }
    }
}

inspectFile('sample.docx');
inspectFile('template.docx');
