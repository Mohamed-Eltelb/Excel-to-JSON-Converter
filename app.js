document.addEventListener('DOMContentLoaded', function() {
    // UI Elements
    const fileInput = document.getElementById('fileInput');
    const dropArea = document.getElementById('dropArea');
    const preview = document.getElementById('preview');
    const downloadBtn = document.getElementById('downloadBtn');
    const copyBtn = document.getElementById('copyBtn');
    const clearBtn = document.getElementById('clearBtn');
    const prettyPrintCheckbox = document.getElementById('prettyPrint');
    const optionsContainer = document.querySelector('.mb-3');
    
    // Create showNulls checkbox dynamically
    const showNullsDiv = document.createElement('div');
    showNullsDiv.className = 'form-check mb-2';
    showNullsDiv.innerHTML = `
        <input class="form-check-input" type="checkbox" id="showNulls" checked>
        <label class="form-check-label" for="showNulls">Show null/empty values</label>
    `;
    optionsContainer.appendChild(showNullsDiv);
    const showNullsCheckbox = document.getElementById('showNulls');
    
    const selectAllBtn = document.getElementById('selectAllBtn');
    const deselectAllBtn = document.getElementById('deselectAllBtn');
    const selectionControls = document.getElementById('selection-controls');
    
    let jsonData = null;
    let originalHeaders = [];
    let selectedColumns = new Set();

    // Initialize the app
    function init() {
        fileInput.addEventListener('change', handleFileSelect);
        downloadBtn.addEventListener('click', downloadJSON);
        copyBtn.addEventListener('click', copyToClipboard);
        clearBtn.addEventListener('click', clearAll);
        
        // Setup drag and drop
        setupDragAndDrop();
        
        // Event listeners for options
        prettyPrintCheckbox.addEventListener('change', displayPreview);
        showNullsCheckbox.addEventListener('change', displayPreview);
        
        selectAllBtn.addEventListener('click', selectAllColumns);
        deselectAllBtn.addEventListener('click', deselectAllColumns);
    }

    // Process Excel file
    function processExcel(data, filename) {
        try {
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Get all original headers
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            originalHeaders = [];
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell = worksheet[XLSX.utils.encode_cell({r: 0, c: C})];
                if (cell && cell.v) originalHeaders.push(cell.v);
            }
            
            // Convert to JSON
            let json = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: null });
            
            // Process data
            json = json.map(row => {
                const cleanRow = {};
                originalHeaders.forEach(header => {
                    const value = row[header] !== undefined ? row[header] : null;
                    const cleanKey = String(header).toLowerCase()
                            .replace(/[^\w\s]/g, '')
                            .replace(/\s+(.)/g, (_, ch) => ch.toUpperCase())
                            .trim();
                    
                    cleanRow[cleanKey] = cleanValue(value);
                });
                return cleanRow;
            });
            
            jsonData = json;
            
            // Initialize column selection (select all by default)
            selectedColumns = new Set();
            originalHeaders.forEach(header => {
                const cleanKey = String(header).toLowerCase()
                        .replace(/[^\w\s]/g, '')
                        .replace(/\s+(.)/g, (_, ch) => ch.toUpperCase())
                        .trim();
                selectedColumns.add(cleanKey);
            });
            
            createColumnCheckboxes();
            displayPreview();
            
            // Enable buttons
            downloadBtn.disabled = false;
            copyBtn.disabled = false;
            downloadBtn.setAttribute('download', `${filename.replace(/\.[^/.]+$/, "")}.json`);
            
        } catch (error) {
            showError(error);
        }
    }

    // Create column selection checkboxes
    function createColumnCheckboxes() {
        const container = document.getElementById('columnCheckboxes');
        container.innerHTML = '';
        
        if (!jsonData || !originalHeaders.length) return;
        
        // Get all available keys (using first row as reference)
        const availableKeys = Object.keys(jsonData[0]);
        
        availableKeys.forEach(key => {
            const checkboxId = `col-${key.replace(/\s+/g, '-')}`;
            const checkbox = document.createElement('div');
            checkbox.className = 'form-check';
            checkbox.innerHTML = `
                <input class="form-check-input column-checkbox" type="checkbox" 
                       id="${checkboxId}" value="${key}" 
                       ${selectedColumns.has(key) ? 'checked' : ''}>
                <label class="form-check-label" for="${checkboxId}">${key}</label>
            `;
            container.appendChild(checkbox);
            selectionControls.classList.remove('d-none');
            // Add event listener
            checkbox.querySelector('input').addEventListener('change', function() {
                if (this.checked) {
                    selectedColumns.add(this.value);
                } else {
                    // Ensure at least one column remains selected
                    if (selectedColumns.size > 1) {
                        selectedColumns.delete(this.value);
                    } else {
                        this.checked = true;
                        showToast('You must select at least one column');
                    }
                }
                displayPreview();
            });
        });
    }

    // Select all columns
    function selectAllColumns() {
        if (!jsonData || !jsonData.length) return;
        
        const checkboxes = document.querySelectorAll('.column-checkbox');
        checkboxes.forEach(checkbox => {
            checkbox.checked = true;
            selectedColumns.add(checkbox.value);
        });
        displayPreview();
    }

    // Deselect all columns (but keep one selected)
    function deselectAllColumns() {
        if (!jsonData || !jsonData.length || selectedColumns.size === 0) return;
        
        const checkboxes = document.querySelectorAll('.column-checkbox');
        let firstChecked = false;
        
        checkboxes.forEach(checkbox => {
            if (!firstChecked) {
                checkbox.checked = true;
                firstChecked = true;
            } else {
                checkbox.checked = false;
            }
        });
        
        // Update selected columns set
        selectedColumns = new Set();
        if (checkboxes.length > 0) {
            selectedColumns.add(checkboxes[0].value);
        }
        
        displayPreview();
        showToast('One column must remain selected');
    }

    // Display preview with selected columns
    function displayPreview() {
        if (!jsonData) {
            preview.textContent = "No file selected";
            return;
        }
        
        try {
            const space = prettyPrintCheckbox.checked ? 2 : 0;
            const showNulls = showNullsCheckbox.checked;
            
            const filteredData = jsonData.map(row => {
                const filteredRow = {};
                Array.from(selectedColumns).forEach(key => {
                    if (showNulls || row[key] !== null) {
                        filteredRow[key] = row[key];
                    }
                });
                return filteredRow;
            });
            
            preview.textContent = JSON.stringify(filteredData, null, space);
            preview.scrollTop = 0;
        } catch (error) {
            showError(error);
        }
    }

    // Download JSON file with selected columns
    function downloadJSON() {
        if (!jsonData) return;
        
        try {
            const space = prettyPrintCheckbox.checked ? 2 : 0;
            const showNulls = showNullsCheckbox.checked;
            
            const filteredData = jsonData.map(row => {
                const filteredRow = {};
                Array.from(selectedColumns).forEach(key => {
                    if (showNulls || row[key] !== null) {
                        filteredRow[key] = row[key];
                    }
                });
                return filteredRow;
            });
            
            const jsonStr = JSON.stringify(filteredData, null, space);
            const blob = new Blob(["\uFEFF" + jsonStr], { 
                type: 'application/json;charset=utf-8' 
            });
            
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = downloadBtn.getAttribute('download') || 'converted.json';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (error) {
            showError(error);
        }
    }

    // Copy to clipboard with selected columns
    function copyToClipboard() {
        if (!jsonData) return;
        
        try {
            const space = prettyPrintCheckbox.checked ? 2 : 0;
            const showNulls = showNullsCheckbox.checked;
            
            const filteredData = jsonData.map(row => {
                const filteredRow = {};
                Array.from(selectedColumns).forEach(key => {
                    if (showNulls || row[key] !== null) {
                        filteredRow[key] = row[key];
                    }
                });
                return filteredRow;
            });
            
            const jsonStr = JSON.stringify(filteredData, null, space);
            navigator.clipboard.writeText(jsonStr)
                .then(() => {
                    copyBtn.textContent = 'Copied!';
                    copyBtn.classList.add('btn-success');
                    setTimeout(() => {
                        copyBtn.textContent = 'Copy to Clipboard';
                        copyBtn.classList.remove('btn-success');
                    }, 2000);
                })
                .catch(err => {
                    showError(err);
                });
        } catch (error) {
            showError(error);
        }
    }

    // Helper functions
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

    function showError(error) {
        console.error(error);
        preview.textContent = `Error: ${error.message}`;
        preview.classList.add('text-danger');
    }

    function showToast(message) {
        const toast = document.createElement('div');
        toast.className = 'position-fixed bottom-0 end-0 p-3';
        toast.innerHTML = `
            <div class="toast show" role="alert">
                <div class="toast-body">
                    ${message}
                </div>
            </div>
        `;
        document.body.appendChild(toast);
        setTimeout(() => {
            toast.remove();
        }, 3000);
    }

    function clearAll() {
        jsonData = null;
        originalHeaders = [];
        selectedColumns = new Set();
        fileInput.value = '';
        document.getElementById('columnCheckboxes').innerHTML = '<p class="text-muted"> Select a file to see column options.</p>';
        preview.textContent = 'No file selected';
        preview.classList.remove('text-danger');
        downloadBtn.disabled = true;
        copyBtn.disabled = true;
        downloadBtn.removeAttribute('download');
        selectionControls.classList.add('d-none');
    }

    function handleFileSelect() {
        if (this.files.length) {
            handleFile(this.files[0]);
        }
    }

    function handleFile(file) {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                processExcel(data, file.name);
            } catch (error) {
                showError(error);
            }
        };
        
        reader.onerror = function() {
            showError(new Error('Failed to read file'));
        };
        
        reader.readAsArrayBuffer(file);
    }

    function reprocessData() {
        if (jsonData && fileInput.files.length) {
            const tempData = jsonData;
            jsonData = null;
            processExcel(new Uint8Array(), fileInput.files[0].name);
            jsonData = tempData;
            displayPreview();
        }
    }

    function setupDragAndDrop() {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, preventDefaults, false);
            document.body.addEventListener(eventName, preventDefaults, false);
        });
        
        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }
        
        function highlight() {
            dropArea.classList.add('highlight');
        }
        
        function unhighlight() {
            dropArea.classList.remove('highlight');
        }
        
        ['dragenter', 'dragover'].forEach(eventName => {
            dropArea.addEventListener(eventName, highlight, false);
        });
        
        ['dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, unhighlight, false);
        });
        
        dropArea.addEventListener('drop', function(e) {
            const dt = e.dataTransfer;
            const files = dt.files;
            if (files.length) {
                fileInput.files = files;
                handleFile(files[0]);
            }
        });
    }

    // Initialize the application
    init();
});