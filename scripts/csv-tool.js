import fs from 'fs';
import path from 'path';

const JSON_PATH = path.resolve('./src/data/mockOrg.json');
const CSV_PATH = path.resolve('./mockOrg.csv');

function getRecordsFromJson() {
    const content = fs.readFileSync(JSON_PATH, 'utf-8');
    const data = JSON.parse(content);
    return data;
}

function recordsToJson(records) {
    const root = records.find(r => r.parentId === null || r.parentId === '');
    const rootId = root ? root.id : (records[0] ? records[0].id : '');
    
    const obj = {
        rootId: rootId,
        people: records
    };
    
    return JSON.stringify(obj, null, 2);
}

function stringifyCSV(records) {
    if (records.length === 0) return '';
    const headers = Object.keys(records[0]);
    let csv = headers.join(',') + '\n';
    
    records.forEach(row => {
        const rowStr = headers.map(header => {
            let val = row[header];
            if (val === null || val === undefined) val = '';
            if (Array.isArray(val)) val = val.join('|');
            let v = String(val);
            if (v.includes(',') || v.includes('"') || v.includes('\n') || v.includes('\r')) {
                return `"${v.replace(/"/g, '""')}"`;
            }
            return v;
        }).join(',');
        csv += rowStr + '\n';
    });
    return csv;
}

function parseCSV(csvText) {
    const lines = [];
    let state = 0;
    let value = "";
    let row = [];
    
    for (let i = 0; i < csvText.length; i++) {
        let char = csvText[i];
        if (state === 0) {
            if (char === '"') {
                state = 1;
            } else if (char === ',') {
                row.push(value);
                value = "";
            } else if (char === '\n') {
                if (value.endsWith('\r')) value = value.slice(0, -1);
                row.push(value);
                lines.push(row);
                row = [];
                value = "";
            } else {
                value += char;
            }
        } else if (state === 1) {
            if (char === '"') {
                if (i + 1 < csvText.length && csvText[i + 1] === '"') {
                    value += '"';
                    i++;
                } else {
                    state = 0;
                }
            } else {
                value += char;
            }
        }
    }
    if (value !== "" || csvText.endsWith(',')) {
        if (value.endsWith('\r')) value = value.slice(0, -1);
        row.push(value);
    }
    if (row.length > 0) lines.push(row);

    const headers = lines[0];
    const records = [];
    for (let i = 1; i < lines.length; i++) {
        const currentLine = lines[i];
        if (currentLine.length === 0 || (currentLine.length === 1 && currentLine[0] === '')) continue; 
        
        const record = {};
        headers.forEach((header, index) => {
            let val = currentLine[index];
            if (val === undefined) val = '';
            if (header === 'parentId' && val === '') val = null;
            else if (header === 'projects') val = val === '' ? [] : val.split('|');
            record[header] = val;
        });
        records.push(record);
    }
    return records;
}

const command = process.argv[2];

if (command === 'to-csv') {
    try {
        console.log(`Reading from ${JSON_PATH}...`);
        const data = getRecordsFromJson();
        const csvText = stringifyCSV(data.people);
        fs.writeFileSync(CSV_PATH, csvText, 'utf-8');
        console.log(`Successfully converted to CSV and saved to ${CSV_PATH}`);
    } catch (e) {
        console.error("Error running to-csv:", e.message);
    }
} else if (command === 'from-csv') {
    try {
        console.log(`Reading from ${CSV_PATH}...`);
        if (!fs.existsSync(CSV_PATH)) {
            console.error(`Cannot find CSV file at ${CSV_PATH}. Make sure you run 'to-csv' first or create the file.`);
            process.exit(1);
        }
        const csvText = fs.readFileSync(CSV_PATH, 'utf-8');
        const records = parseCSV(csvText);
        const jsonCode = recordsToJson(records);
        fs.writeFileSync(JSON_PATH, jsonCode, 'utf-8');
        console.log(`Successfully converted to JSON and saved to ${JSON_PATH}`);
    } catch (e) {
        console.error("Error running from-csv:", e.message);
    }
} else {
    console.log(`
Usage:
  node scripts/csv-tool.js to-csv    - Converts src/data/mockOrg.json to mockOrg.csv in the root folder.
  node scripts/csv-tool.js from-csv  - Converts mockOrg.csv into src/data/mockOrg.json.
`);
}
