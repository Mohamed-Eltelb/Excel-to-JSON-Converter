// Import SheetJS library
self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');

self.onmessage = function(e) {
    try {
        const { buffer, filename, options } = e.data;
        const data = new Uint8Array(buffer);
        
        // Process Excel file
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Get all headers
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const headers = [];
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell = worksheet[XLSX.utils.encode_cell({r: 0, c: C})];
            if (cell && cell.v) headers.push(cell.v);
        }
        
        // Convert to JSON
        let json = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: null });
        
        // Process data
        const processedData = json.map(row => {
            const cleanRow = {};
            headers.forEach(header => {
                const value = row[header] !== undefined ? row[header] : null;
                const cleanKey = options.camelCase 
                    ? String(header).toLowerCase()
                        .replace(/[^\w\s]/g, '')
                        .replace(/\s+(.)/g, (_, ch) => ch.toUpperCase())
                        .trim()
                    : header;
                
                cleanRow[cleanKey] = cleanValue(value);
            });
            return cleanRow;
        });
        
        // Send result back to main thread
        self.postMessage({
            jsonData: processedData,
            filename: filename.replace(/\.[^/.]+$/, "") + '.json'
        });
        
    } catch (error) {
        self.postMessage({ error: error.message });
    }
};

// Value cleaning function
function cleanValue(value) {
    if (typeof value === 'number' && isNaN(value)) return null;
    if (value === null || value === undefined) return null;
    if (typeof value === 'string') {
        const normalized = value.normalize('NFKC');
        return normalized
            .replace(/\u00C2\u00AE/g, '®')
            .replace(/\u00E2\u0084\u00A2/g, '™')
            .replace(/[^\x00-\x7F®™©±µ]/g, "")
            .trim() || null;
    }
    return value;
}