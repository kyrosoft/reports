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
function addField() {
    const container = document.getElementById('fieldsContainer');
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'field-input';
    input.placeholder = 'Field name';
    container.appendChild(input);
}

function createTemplate() {
    const name = document.getElementById('templateName').value || 'template';
    const fields = Array.from(document.querySelectorAll('.field-input'))
        .map(i => i.value)
        .filter(v => v.trim());

    if (fields.length === 0) {
        alert('Add at least one field');
        return;
    }

    // Placeholder - would use xlsx library to create actual Excel file
    alert(`Template "${name}" with fields:\n${fields.join('\n')}\n\n(Integration with xlsx library required)`);
}

// Compiler
let templateFile = null;
let reportFiles = [];
const compileBtn = document.getElementById('compileBtn');

// Template upload
const templateDrop = document.getElementById('templateDropZone');
templateDrop.onclick = () => document.getElementById('templateInput').click();
templateDrop.ondragover = (e) => { e.preventDefault(); templateDrop.classList.add('drag-over'); };
templateDrop.ondragleave = () => templateDrop.classList.remove('drag-over');
templateDrop.ondrop = (e) => {
    e.preventDefault();
    templateDrop.classList.remove('drag-over');
    setTemplate(e.dataTransfer.files[0]);
};

document.getElementById('templateInput').onchange = (e) => setTemplate(e.target.files[0]);

function setTemplate(file) {
    if (!file || !file.name.endsWith('.xlsx')) {
        alert('Only .xlsx files are supported');
        return;
    }
    templateFile = file;
    document.getElementById('templateInfo').innerHTML = `<b>${file.name}</b> (${size(file.size)})`;
    document.getElementById('templateInfo').classList.remove('hidden');
    updateBtn();
}

// Reports upload
const reportsDrop = document.getElementById('reportsDropZone');
reportsDrop.onclick = () => document.getElementById('reportsInput').click();
reportsDrop.ondragover = (e) => { e.preventDefault(); reportsDrop.classList.add('drag-over'); };
reportsDrop.ondragleave = () => reportsDrop.classList.remove('drag-over');
reportsDrop.ondrop = (e) => {
    e.preventDefault();
    reportsDrop.classList.remove('drag-over');
    addReports(e.dataTransfer.files);
};

document.getElementById('reportsInput').onchange = (e) => addReports(e.target.files);

function addReports(files) {
    for (let file of files) {
        if (!file.name.endsWith('.xlsx')) {
            alert(`${file.name} is not a .xlsx file`);
            continue;
        }
        if (!reportFiles.find(f => f.name === file.name)) {
            reportFiles.push(file);
        }
    }
    renderReports();
    updateBtn();
}

function renderReports() {
    const list = document.getElementById('reportsList');
    if (reportFiles.length === 0) {
        list.classList.add('hidden');
        return;
    }
    list.innerHTML = reportFiles.map((f, i) => `
        <div class="file-item">
            <span>${f.name} (${size(f.size)})</span>
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
compileBtn.onclick = async () => {
    compileBtn.disabled = true;
    document.getElementById('progressSection').classList.remove('hidden');
    document.getElementById('downloadSection').classList.add('hidden');

    let progress = 0;
    const bar = document.getElementById('progressFill');
    while (progress < 100) {
        progress += Math.random() * 20;
        if (progress > 100) progress = 100;
        bar.style.width = progress + '%';
        await new Promise(r => setTimeout(r, 200));
    }

    document.getElementById('progressSection').classList.add('hidden');
    document.getElementById('downloadSection').classList.remove('hidden');
    compileBtn.disabled = false;
};

// Download (placeholder)
document.getElementById('downloadBtn').onclick = () => {
    alert('Backend integration required for actual download');
};

function size(bytes) {
    const k = 1024;
    const units = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return (bytes / Math.pow(k, i)).toFixed(1) + ' ' + units[i];
}

function updateBtn() {
    compileBtn.disabled = !templateFile || reportFiles.length === 0;
}
