// Tabs
document.getElementById('tabCreator').onclick = () => {
    document.getElementById('creatorSection').classList.remove('hidden');
    document.getElementById('compilerSection').classList.add('hidden');
    document.getElementById('tabCreator').classList.add('active');
    document.getElementById('tabCompiler').classList.remove('active');
};

document.getElementById('tabCompiler').onclick = () => {
    document.getElementById('compilerSection').classList.remove('hidden');
    document.getElementById('creatorSection').classList.add('hidden');
    document.getElementById('tabCompiler').classList.add('active');
    document.getElementById('tabCreator').classList.remove('active');
};

// Template Creator
function addMetadata() {
    const container = document.getElementById('metadataContainer');
    const div = document.createElement('div');
    div.innerHTML = `
        <input type="text" class="metadata-key" placeholder="Key">
        <input type="text" class="metadata-value" placeholder="Value">
        <button onclick="this.parentElement.remove()" style="background:#ef4444;padding:5px 10px;">×</button>
    `;
    container.appendChild(div);
}

function addColumn() {
    const container = document.getElementById('columnsContainer');
    const div = document.createElement('div');
    div.innerHTML = `
        <input type="text" class="column-input" placeholder="Column name">
        <button onclick="this.parentElement.remove();updatePreview()" style="background:#ef4444;padding:5px 10px;">×</button>
    `;
    container.appendChild(div);
}

function addRow() {
    const container = document.getElementById('rowsContainer');
    const div = document.createElement('div');
    div.innerHTML = `
        <input type="text" class="row-input" placeholder="Row label">
        <button onclick="this.parentElement.remove();updatePreview()" style="background:#ef4444;padding:5px 10px;">×</button>
    `;
    container.appendChild(div);
}

function updatePreview() {
    const columns = Array.from(document.querySelectorAll('.column-input')).map(i => i.value).filter(v => v.trim());
    const rows = Array.from(document.querySelectorAll('.row-input')).map(i => i.value).filter(v => v.trim());
    const metadata = {};
    document.querySelectorAll('#metadataContainer > div').forEach(div => {
        const key = div.querySelector('.metadata-key')?.value;
        const value = div.querySelector('.metadata-value')?.value;
        if (key && value) metadata[key] = value;
    });

    let html = '';

    // Metadata preview
    if (Object.keys(metadata).length > 0) {
        html += '<div class="preview-info"><strong>Metadata:</strong><br>';
        for (let [k, v] of Object.entries(metadata)) {
            html += `${k}: ${v}<br>`;
        }
        html += '</div>';
    }

    // Table preview
    if (columns.length > 0 || rows.length > 0) {
        html += '<table><thead><tr>';
        html += '<th></th>'; // Corner cell
        columns.forEach(col => html += `<th>${col}</th>`);
        html += '</tr></thead><tbody>';

        rows.forEach(row => {
            html += `<tr><td><strong>${row}</strong></td>`;
            columns.forEach(() => html += '<td>[data]</td>');
            html += '</tr>';
        });

        html += '</tbody></table>';
    }

    document.getElementById('preview').innerHTML = html || '<p>Add columns and/or rows to see preview</p>';
}

// Update preview when inputs change
document.addEventListener('input', (e) => {
    if (e.target.matches('.column-input, .row-input, .metadata-key, .metadata-value')) {
        updatePreview();
    }
});

// Initial preview
updatePreview();

function createTemplate() {
    const name = document.getElementById('templateName').value || 'template';
    const columns = Array.from(document.querySelectorAll('.column-input')).map(i => i.value).filter(v => v.trim());
    const rows = Array.from(document.querySelectorAll('.row-input')).map(i => i.value).filter(v => v.trim());
    const metadata = {};
    document.querySelectorAll('#metadataContainer > div').forEach(div => {
        const key = div.querySelector('.metadata-key')?.value;
        const value = div.querySelector('.metadata-value')?.value;
        if (key && value) metadata[key] = value;
    });

    if (columns.length === 0 && rows.length === 0) {
        alert('Add at least one column or row');
        return;
    }

    // Create workbook
    const wb = XLSX.utils.book_new();

    // Add metadata sheet
    if (Object.keys(metadata).length > 0) {
        const metadataData = Object.entries(metadata).map(([k, v]) => [k, v]);
        metadataData.unshift(['Key', 'Value']);
        const wsMeta = XLSX.utils.aoa_to_sheet(metadataData);
        XLSX.utils.book_append_sheet(wb, wsMeta, 'Metadata');
    }

    // Create data sheet
    const data = [['']];
    data[0] = [''].concat(columns); // Header row

    rows.forEach(rowLabel => {
        const rowData = [rowLabel];
        columns.forEach(() => rowData.push(''));
        data.push(rowData);
    });

    const wsData = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, wsData, 'Data');

    // Download
    XLSX.writeFile(wb, `${name}.xlsx`);
}

// Compiler
let templateFile = null;
let reportFiles = [];
const compileBtn = document.getElementById('compileBtn');

document.getElementById('templateInput').onchange = (e) => setTemplate(e.target.files[0]);

function setTemplate(file) {
    if (!file || !file.name.endsWith('.xlsx')) {
        showValidationError('Template must be a .xlsx file');
        return;
    }
    templateFile = file;
    document.getElementById('templateInfo').innerHTML = `<b>${file.name}</b> (${size(file.size)})`;
    document.getElementById('templateInfo').classList.remove('hidden');
    clearValidationErrors();
    updateBtn();
}

function showValidationError(msg) {
    const el = document.getElementById('validationErrors');
    el.textContent = msg;
    el.classList.remove('hidden');
}

function clearValidationErrors() {
    document.getElementById('validationErrors').classList.add('hidden');
}

// Reports upload
function validateReport(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const wb = XLSX.read(e.target.result, { type: 'array' });
                resolve(wb.SheetNames.length > 0);
            } catch {
                resolve(false);
            }
        };
        reader.onerror = () => resolve(false);
        reader.readAsArrayBuffer(file);
    });
}

async function addReportFile() {
    const input = document.getElementById('reportsInput');
    const file = input.files[0];
    if (!file) return;

    if (!file.name.endsWith('.xlsx')) {
        showValidationError(`${file.name} is not a .xlsx file`);
        return;
    }

    if (reportFiles.find(f => f.file.name === file.name)) {
        showValidationError(`${file.name} already added`);
        return;
    }

    const isValid = await validateReport(file);
    reportFiles.push({ file, valid: isValid });
    renderReports();
    input.value = '';
    clearValidationErrors();
    updateBtn();
}

function renderReports() {
    const list = document.getElementById('reportsList');
    if (reportFiles.length === 0) {
        list.classList.add('hidden');
        return;
    }
    list.innerHTML = reportFiles.map((item, i) => `
        <div class="file-item ${item.valid ? 'valid' : 'invalid'}">
            <span>${item.file.name} (${size(item.file.size)})
                <span class="file-status ${item.valid ? 'valid' : 'invalid'}">
                    ${item.valid ? '✓ Valid' : '✗ Invalid'}
                </span>
            </span>
            <button class="file-remove" onclick="removeReport(${i})">×</button>
        </div>
    `).join('');
    list.classList.remove('hidden');
}

function removeReport(i) {
    reportFiles.splice(i, 1);
    renderReports();
    updateBtn();
}

// Compile
let compiledWorkbook = null;

compileBtn.onclick = async () => {
    compileBtn.disabled = true;
    document.getElementById('progressSection').classList.remove('hidden');
    document.getElementById('downloadSection').classList.add('hidden');

    const bar = document.getElementById('progressFill');
    const validReports = reportFiles.filter(r => r.valid);
    const total = validReports.length;

    // Read template
    const templateData = await readFile(templateFile);
    const templateWb = XLSX.read(templateData, { type: 'array' });
    const newWb = XLSX.utils.book_new();

    // Copy template sheets
    templateWb.SheetNames.forEach(name => {
        const ws = templateWb.Sheets[name];
        XLSX.utils.book_append_sheet(newWb, ws, name);
    });

    // Append data from each report
    for (let i = 0; i < total; i++) {
        const { file } = validReports[i];
        const data = await readFile(file);
        const wb = XLSX.read(data, { type: 'array' });

        // Add each sheet with report name prefix
        wb.SheetNames.forEach(name => {
            const ws = wb.Sheets[name];
            const newSheetName = `${file.name.replace('.xlsx', '')} - ${name}`;
            XLSX.utils.book_append_sheet(newWb, ws, newSheetName);
        });

        bar.style.width = ((i + 1) / total * 100) + '%';
    }

    compiledWorkbook = newWb;

    document.getElementById('progressSection').classList.add('hidden');
    document.getElementById('downloadSection').classList.remove('hidden');
    compileBtn.disabled = false;
};

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

document.getElementById('downloadBtn').onclick = () => {
    if (compiledWorkbook) {
        XLSX.writeFile(compiledWorkbook, 'compiled_reports.xlsx');
    }
};

function size(bytes) {
    const k = 1024;
    const units = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return (bytes / Math.pow(k, i)).toFixed(1) + ' ' + units[i];
}

function updateBtn() {
    const hasValidReports = reportFiles.some(r => r.valid);
    compileBtn.disabled = !templateFile || !hasValidReports;
}
