<script lang="ts">
	import { onMount } from 'svelte';
	import * as XLSX from 'xlsx-js-style';
	import JSZip from 'jszip';

	// Types
	interface Column {
		value?: string;
		type?: 'string' | 'numeric' | 'date';
	}

	interface TemplateFile {
		name: string;
		content: ArrayBuffer;
	}

	interface ReportFile {
		file: {
			name: string;
			content: ArrayBuffer;
		};
		valid: boolean;
		sheets: number;
		sheetNames: string[];
		workbook: any;
		validationError?: string;
		originalName?: string;
	}

	interface ValidationResult {
		message: string;
		isValid: boolean;
	}

	interface Toast {
		message: string;
		type: 'error' | 'success';
		id: number;
	}

	interface ExcelValidation {
		valid: boolean;
		sheets: number;
		sheetNames: string[];
		workbook: any;
	}

	interface TemplateInfo {
		sheetName: string;
		columns: any[];
		data: any[][];
		schema?: SchemaEntry[];
	}

	interface SchemaEntry {
		column: string;
		type: 'string' | 'numeric' | 'date';
	}

	const SCHEMA_SHEET = '_schema';
	const DATA_TYPES = ['string', 'numeric', 'date'] as const;
	type DataType = typeof DATA_TYPES[number];

	// Toast state and functions
	let toasts = $state<Toast[]>([]);
	let toastId = 0;

	function showToast(message: string, type: 'error' | 'success' = 'error') {
		const id = toastId++;
		const toast: Toast = { message, type, id };
		toasts = [...toasts, toast];
		setTimeout(() => {
			toasts = toasts.filter((t) => t.id !== id);
		}, 3000);
	}

	// Compile Report state
	let templateFile = $state<TemplateFile | null>(null);
	let reportFiles = $state<ReportFile[]>([]);
	let progress = $state(0);
	let compiling = $state(false);

	// Validate Report state
	let validateTemplateFile = $state<TemplateFile | null>(null);
	let validateReportFile = $state<any | null>(null);
	let validationResult = $state<ValidationResult | null>(null);

	// Template Report state
	let columns = $state<Column[]>([{ value: '', type: 'string' }]);

	// Utility functions
	function formatFileSize(bytes: number): string {
		const k = 1024;
		const units = ['B', 'KB', 'MB', 'GB'];
		const i = Math.floor(Math.log(bytes) / Math.log(k));
		return (bytes / Math.pow(k, i)).toFixed(1) + ' ' + units[i];
	}

	function validateExcelFile(arrayBuffer: ArrayBuffer): ExcelValidation {
		try {
			const wb = XLSX.read(arrayBuffer, { type: 'array' });
			return {
				valid: wb.SheetNames.length > 0,
				sheets: wb.SheetNames.length,
				sheetNames: wb.SheetNames,
				workbook: wb
			};
		} catch {
			return { valid: false, sheets: 0, sheetNames: [], workbook: null };
		}
	}

	function readSchemaFromWorkbook(wb: any): SchemaEntry[] {
		if (!wb.SheetNames.includes(SCHEMA_SHEET)) return [];
		const ws = wb.Sheets[SCHEMA_SHEET];
		const data: any = XLSX.utils.sheet_to_json(ws, { header: 1 });
		// Row 0 is header ["column","type"], rows 1+ are entries
		const schema: SchemaEntry[] = [];
		for (let i = 1; i < data.length; i++) {
			if (data[i][0] && data[i][1]) {
				schema.push({ column: String(data[i][0]), type: data[i][1] as DataType });
			}
		}
		return schema;
	}

	function getTemplateInfo(templateFile: TemplateFile): TemplateInfo | null {
		if (!templateFile) return null;

		const wb = XLSX.read(templateFile.content, { type: 'array' });
		const sheetName = wb.SheetNames.filter((s) => s !== SCHEMA_SHEET)[0];
		if (!sheetName) return null;
		const ws = wb.Sheets[sheetName];
		const data: any = XLSX.utils.sheet_to_json(ws, { header: 1 });

		if (data.length === 0) return null;

		const schema = readSchemaFromWorkbook(wb);

		return {
			sheetName,
			columns: data[0],
			data,
			schema
		};
	}

	async function extractXlsxFromZip(file: File): Promise<any[]> {
		const zip = new JSZip();
		const arrayBuffer = await file.arrayBuffer();
		const zipContent = await zip.loadAsync(arrayBuffer);
		const xlsxFiles: any[] = [];

		for (const [filename, zipEntry] of Object.entries(zipContent.files)) {
			if (!zipEntry.dir && filename.toLowerCase().endsWith('.xlsx')) {
				const content = await zipEntry.async('arraybuffer');
				xlsxFiles.push({ name: filename, content, originalFile: file });
			}
		}

		return xlsxFiles;
	}

	// ── Schema cell validation helper ───────────────────────────────────────────
	function validateCellType(value: any, type: DataType): boolean {
		if (value === null || value === undefined || value === '') return true; // allow blanks
		switch (type) {
			case 'numeric':
				return !isNaN(Number(value));
			case 'date': {
				// Accept JS Date objects (xlsx parses some dates), serial numbers, or parseable strings
				if (value instanceof Date) return !isNaN(value.getTime());
				if (typeof value === 'number') return true; // Excel serial date
				const d = new Date(String(value));
				return !isNaN(d.getTime());
			}
			case 'string':
			default:
				return true;
		}
	}

	interface SchemaViolation {
		row: number;
		column: string;
		expected: DataType;
		got: any;
	}

	function validateDataAgainstSchema(
		ws: any,
		schema: SchemaEntry[],
		headerRow: any[]
	): SchemaViolation[] {
		if (!schema || schema.length === 0) return [];

		const data: any = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
		const violations: SchemaViolation[] = [];

		// Build column-index map from header
		const colIndex: Record<string, number> = {};
		headerRow.forEach((h, i) => (colIndex[String(h).trim()] = i));

		for (let rowIdx = 1; rowIdx < data.length; rowIdx++) {
			for (const entry of schema) {
				const idx = colIndex[entry.column];
				if (idx === undefined) continue;
				const cellVal = data[rowIdx][idx];
				if (!validateCellType(cellVal, entry.type)) {
					violations.push({
						row: rowIdx + 1, // 1-based for display
						column: entry.column,
						expected: entry.type,
						got: cellVal
					});
				}
			}
		}

		return violations;
	}

	// ── Excel workbook creation ──────────────────────────────────────────────────
	function createExcelWorkbook(
		cols: Column[],
		sheetName: string = 'KyroReports',
		filename: string = 'Template - KyroReports.xlsx'
	): void {
		const wb = XLSX.utils.book_new();

		// ── Data sheet ──────────────────────────────────────────────────────────
		const colNames = cols.map((c) => c.value || '');
		const wsData: any = XLSX.utils.aoa_to_sheet([colNames]);

		const range = XLSX.utils.decode_range(wsData['!ref']);
		for (let col = range.s.c; col <= range.e.c; col++) {
			const cell = XLSX.utils.encode_cell({ r: 0, c: col });
			if (!wsData[cell]) continue;
			wsData[cell].s = {
				font: { bold: true },
				alignment: { horizontal: 'center', vertical: 'center' }
			};
		}
		wsData['!cols'] = colNames.map((c) => ({ wch: Math.max(String(c).length + 2, 10) }));
		XLSX.utils.book_append_sheet(wb, wsData, sheetName);

		// ── Schema sheet ────────────────────────────────────────────────────────
		const schemaRows: any[][] = [['column', 'type']];
		cols.forEach((c) => {
			if (c.value?.trim()) schemaRows.push([c.value.trim(), c.type || 'string']);
		});
		const schemaWs: any = XLSX.utils.aoa_to_sheet(schemaRows);

		// Style schema header
		['A1', 'B1'].forEach((addr) => {
			if (schemaWs[addr])
				schemaWs[addr].s = { font: { bold: true }, fill: { fgColor: { rgb: 'E2E8F0' } } };
		});
		schemaWs['!cols'] = [{ wch: 30 }, { wch: 12 }];
		XLSX.utils.book_append_sheet(wb, schemaWs, SCHEMA_SHEET);

		XLSX.writeFile(wb, filename);
	}

	// ── Compile reports ──────────────────────────────────────────────────────────
	function compileReports(
		templateFile: TemplateFile,
		reportFiles: any[],
		templateInfo: TemplateInfo,
		onProgress: (p: number) => void
	): { warnings: string[] } {
		const newWb = XLSX.utils.book_new();
		const headerRow = ['No.', ...templateInfo.columns];
		const allData = [headerRow];
		let rowNum = 1;
		const warnings: string[] = [];

		for (let i = 0; i < reportFiles.length; i++) {
			const report = reportFiles[i];
			const ws = report.workbook.Sheets[templateInfo.sheetName];
			const data: any = XLSX.utils.sheet_to_json(ws, { header: 1 });

			// Schema validation per file
			if (templateInfo.schema && templateInfo.schema.length > 0) {
				const violations = validateDataAgainstSchema(ws, templateInfo.schema, templateInfo.columns);
				violations.forEach((v) => {
					warnings.push(
						`${report.file.name} row ${v.row} col "${v.column}": expected ${v.expected}, got "${v.got}"`
					);
				});
			}

			for (let j = 1; j < data.length; j++) {
				allData.push([rowNum++, ...data[j]]);
			}

			onProgress(((i + 1) / reportFiles.length) * 100);
		}

		const compiledWs: any = XLSX.utils.aoa_to_sheet(allData);

		const range = XLSX.utils.decode_range(compiledWs['!ref']);
		for (let col = range.s.c; col <= range.e.c; col++) {
			const cell = XLSX.utils.encode_cell({ r: 0, c: col });
			if (!compiledWs[cell]) continue;
			compiledWs[cell].s = {
				font: { bold: true },
				alignment: { horizontal: 'center', vertical: 'center' }
			};
		}
		compiledWs['!cols'] = headerRow.map((c: any) => ({ wch: Math.max(String(c).length + 2, 10) }));
		XLSX.utils.book_append_sheet(newWb, compiledWs, templateInfo.sheetName);

		// Copy schema into compiled workbook too
		if (templateInfo.schema && templateInfo.schema.length > 0) {
			const schemaRows: any[][] = [['column', 'type']];
			templateInfo.schema.forEach((e) => schemaRows.push([e.column, e.type]));
			const schemaWs: any = XLSX.utils.aoa_to_sheet(schemaRows);
			['A1', 'B1'].forEach((addr) => {
				if (schemaWs[addr])
					schemaWs[addr].s = { font: { bold: true }, fill: { fgColor: { rgb: 'E2E8F0' } } };
			});
			schemaWs['!cols'] = [{ wch: 30 }, { wch: 12 }];
			XLSX.utils.book_append_sheet(newWb, schemaWs, SCHEMA_SHEET);
		}

		XLSX.writeFile(newWb, 'Compiled - KyroReports.xlsx');
		return { warnings };
	}

	// ── Compile Report handlers ──────────────────────────────────────────────────
	async function setTemplate(file: any) {
		if (!file || !file.name?.toLowerCase()?.endsWith('.xlsx')) {
			showToast('Template must be a .xlsx file', 'error');
			return;
		}
		try {
			const content = await file.arrayBuffer();
			const validation = validateExcelFile(content);
			if (!validation.valid) { showToast('Invalid Excel file', 'error'); return; }
			templateFile = { name: file.name, content };
			if (reportFiles.length > 0) validateReportsAgainstTemplate();
		} catch { showToast('Failed to read file', 'error'); }
	}

	async function setReportsZip(file: any) {
		if (!file) return;
		if (!file.name?.toLowerCase()?.endsWith('.zip')) {
			showToast(`${file.name} is not a .zip file`, 'error'); return;
		}
		reportFiles = [];
		try {
			const xlsxFiles = await extractXlsxFromZip(file);
			if (xlsxFiles.length === 0) {
				showToast(`No .xlsx files found in ${file.name}`, 'error'); return;
			}
			for (const xlsxFile of xlsxFiles) {
				const validation = validateExcelFile(xlsxFile.content);
				reportFiles = [...reportFiles, {
					file: xlsxFile, valid: validation.valid,
					sheets: validation.sheets, sheetNames: validation.sheetNames,
					workbook: validation.workbook, originalName: file.name
				}];
			}
		} catch { showToast(`Failed to read ${file.name}`, 'error'); return; }
		if (templateFile) validateReportsAgainstTemplate();
	}

	function removeReport(index: number) {
		reportFiles = reportFiles.filter((_, i) => i !== index);
	}

	function validateReportsAgainstTemplate() {
		if (!templateFile || reportFiles.length === 0) return;
		const templateInfo = getTemplateInfo(templateFile);
		if (!templateInfo) return;

		reportFiles = reportFiles.map((report) => {
			if (!report.sheetNames.includes(templateInfo.sheetName))
				return { ...report, valid: false, validationError: `Sheet "${templateInfo.sheetName}" not found` };

			const ws = report.workbook.Sheets[templateInfo.sheetName];
			const reportData: any = XLSX.utils.sheet_to_json(ws, { header: 1 });

			if (reportData.length === 0)
				return { ...report, valid: false, validationError: 'Sheet is empty' };

			const templateColsStr = templateInfo.columns.map((c: any) => String(c).trim()).join(',');
			const reportColsStr = reportData[0].map((c: any) => String(c).trim()).join(',');

			if (templateColsStr !== reportColsStr)
				return { ...report, valid: false, validationError: `Columns don't match template` };

			// Schema type check
			if (templateInfo.schema && templateInfo.schema.length > 0) {
				const violations = validateDataAgainstSchema(ws, templateInfo.schema, templateInfo.columns);
				if (violations.length > 0)
					return { ...report, valid: false, validationError: `Type mismatch in ${violations.length} cell(s)` };
			}

			return { ...report, valid: true, validationError: undefined };
		});
	}

	async function compileAndDownload() {
		const errors: string[] = [];
		if (!templateFile) errors.push('Please upload a template file');
		if (reportFiles.length === 0) errors.push('Please upload a reports ZIP file');

		if (errors.length === 0 && templateFile && reportFiles.length > 0) {
			const templateInfo = getTemplateInfo(templateFile);
			if (!templateInfo) {
				errors.push('Template file is empty or invalid');
			} else {
				for (const report of reportFiles) {
					if (!report.valid) { errors.push(`${report.file.name}: Invalid Excel file`); continue; }
					if (!report.sheetNames.includes(templateInfo.sheetName)) {
						errors.push(`${report.file.name}: Sheet "${templateInfo.sheetName}" not found. Available: ${report.sheetNames.join(', ')}`);
						continue;
					}
					const ws = report.workbook.Sheets[templateInfo.sheetName];
					const reportData: any = XLSX.utils.sheet_to_json(ws, { header: 1 });
					if (reportData.length === 0) { errors.push(`${report.file.name}: Sheet is empty`); continue; }

					const templateColsStr = templateInfo.columns.map((c: any) => String(c).trim()).join(',');
					const reportColsStr = reportData[0].map((c: any) => String(c).trim()).join(',');
					if (templateColsStr !== reportColsStr)
						errors.push(`${report.file.name}: Columns mismatch. Expected: "${templateColsStr}", Found: "${reportColsStr}"`);
				}
			}
		}

		if (errors.length > 0) { errors.forEach((e) => showToast(e, 'error')); return; }

		compiling = true; progress = 0;
		try {
			const templateInfo = getTemplateInfo(templateFile!);
			const { warnings } = compileReports(templateFile!, reportFiles, templateInfo!, (p) => (progress = p));
			if (warnings.length > 0) {
				warnings.slice(0, 5).forEach((w) => showToast(`⚠ ${w}`, 'error'));
				if (warnings.length > 5) showToast(`…and ${warnings.length - 5} more type warnings`, 'error');
			}
			showToast('Report compiled and downloaded successfully', 'success');
		} catch { showToast('Failed to compile reports', 'error'); }
		compiling = false; progress = 0;
	}

	// ── Validate Report handlers ─────────────────────────────────────────────────
	async function setValidateTemplate(file: any) {
		if (!file || !file.name?.toLowerCase()?.endsWith('.xlsx')) {
			showToast('Template must be a .xlsx file', 'error'); return;
		}
		try {
			const content = await file.arrayBuffer();
			const validation = validateExcelFile(content);
			if (!validation.valid) { showToast('Invalid Excel file', 'error'); return; }
			validateTemplateFile = { name: file.name, content };
			if (validateReportFile) validateSingleReport();
		} catch { showToast('Failed to read file', 'error'); }
	}

	async function setValidateReport(file: any) {
		if (!file || !file.name?.toLowerCase()?.endsWith('.xlsx')) {
			showToast('Report must be a .xlsx file', 'error'); return;
		}
		try {
			const content = await file.arrayBuffer();
			const validation = validateExcelFile(content);
			if (!validation.valid) { showToast('Invalid Excel file', 'error'); return; }
			validateReportFile = { name: file.name, content, validation };
			if (validateTemplateFile) validateSingleReport();
			else validationResult = { message: 'Please upload a template first', isValid: false };
		} catch { showToast('Failed to read file', 'error'); }
	}

	function validateSingleReport() {
		if (!validateTemplateFile) { showToast('Please upload a template file first', 'error'); return; }
		if (!validateReportFile) { showToast('Please upload a report file first', 'error'); return; }

		const templateInfo = getTemplateInfo(validateTemplateFile);
		if (!templateInfo) { validationResult = { message: 'Template file is empty or invalid', isValid: false }; return; }

		const validation = validateReportFile.validation;

		if (!validation.sheetNames.includes(templateInfo.sheetName)) {
			validationResult = {
				message: `Sheet "${templateInfo.sheetName}" not found. Available: ${validation.sheetNames.join(', ')}`,
				isValid: false
			};
			return;
		}

		const ws = validation.workbook.Sheets[templateInfo.sheetName];
		const reportData: any = XLSX.utils.sheet_to_json(ws, { header: 1 });

		if (reportData.length === 0) { validationResult = { message: 'Sheet is empty', isValid: false }; return; }

		const templateColsStr = templateInfo.columns.map((c: any) => String(c).trim()).join(',');
		const reportColsStr = reportData[0].map((c: any) => String(c).trim()).join(',');

		if (templateColsStr !== reportColsStr) {
			validationResult = {
				message: `Columns don't match template.<br>Expected: "${templateColsStr}"<br>Found: "${reportColsStr}"`,
				isValid: false
			};
			return;
		}

		// Schema type validation
		if (templateInfo.schema && templateInfo.schema.length > 0) {
			const violations = validateDataAgainstSchema(ws, templateInfo.schema, templateInfo.columns);
			if (violations.length > 0) {
				const lines = violations.slice(0, 8).map(
					(v) => `Row ${v.row}, "${v.column}": expected <b>${v.expected}</b>, got "<i>${v.got}</i>"`
				);
				if (violations.length > 8) lines.push(`…and ${violations.length - 8} more`);
				validationResult = {
					message: `<b>✗ Type violations (${violations.length})</b><br>${lines.join('<br>')}`,
					isValid: false
				};
				return;
			}

			// Show schema summary when valid
			const schemaSummary = templateInfo.schema
				.map((e) => `<span class="inline-block px-1 rounded text-[10px] mr-1 mb-0.5 bg-slate-100">${e.column}: <b>${e.type}</b></span>`)
				.join('');
			validationResult = {
				message: `<b>✓ Valid</b><br>Report matches template structure and all type constraints.<br><span class="text-[10px] text-gray-500">Schema: ${schemaSummary}</span>`,
				isValid: true
			};
			return;
		}

		validationResult = {
			message: '<b>✓ Valid</b><br>Report matches template structure',
			isValid: true
		};
	}

	// ── Template Report functions ────────────────────────────────────────────────
	function renderTemplate() {
		if (columns.length === 0) columns = [{ value: '', type: 'string' }];
	}

	function autoResizeInput(input: HTMLInputElement) {
		const textLength = input.value.length;
		const calculatedWidth = Math.max(80, Math.min(200, textLength * 8 + 40));
		input.style.width = calculatedWidth + 'px';
	}

	function moveColumnLeft(index: number) {
		if (index > 0) {
			const nc = [...columns];
			[nc[index], nc[index - 1]] = [nc[index - 1], nc[index]];
			columns = nc;
		}
	}

	function moveColumnRight(index: number) {
		if (index < columns.length - 1) {
			const nc = [...columns];
			[nc[index], nc[index + 1]] = [nc[index + 1], nc[index]];
			columns = nc;
		}
	}

	function addColumn() {
		columns = [...columns, { value: '', type: 'string' }];
		setTimeout(() => {
			const inputs = document.querySelectorAll<HTMLInputElement>('.column-input-field');
			if (inputs.length > 0) inputs[inputs.length - 1].focus();
		}, 0);
	}

	function removeColumn(index: number) {
		if (columns.length <= 1) { showToast('You must have at least one column', 'error'); return; }
		columns = columns.filter((_, i) => i !== index);
	}

	function createTemplate() {
		const errors: string[] = [];
		columns.forEach((col, index) => {
			if (!col.value || col.value.trim() === '') errors.push(`Column ${index + 1} is empty`);
		});
		if (errors.length > 0) { errors.forEach((e) => showToast(e, 'error')); return; }

		const validColumns = columns.filter((c) => c.value?.trim());
		if (validColumns.length === 0) { showToast('Please add at least one column', 'error'); return; }

		createExcelWorkbook(validColumns, 'KyroReports', 'Template - KyroReports.xlsx');
		showToast('Template downloaded successfully', 'success');
	}

	const TYPE_COLORS: Record<DataType, string> = {
		string: 'bg-blue-50 border-blue-200 text-blue-700',
		numeric: 'bg-amber-50 border-amber-200 text-amber-700',
		date: 'bg-purple-50 border-purple-200 text-purple-700'
	};

	const TYPE_ICONS: Record<DataType, string> = {
		string: 'Aa',
		numeric: '123',
		date: '📅'
	};

	onMount(() => { renderTemplate(); });
</script>

<svelte:head>
	<title>Kyro Reports</title>
	<link
		rel="icon"
		href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><rect fill='%2310b981' width='100' height='100' rx='15'/><path fill='white' d='M25 30h50v8H25zm0 15h50v8H25zm0 15h35v8H25z'/></svg>"
	/>
</svelte:head>

<div class="min-h-screen bg-gray-50">
	<!-- Toast Container -->
	<div class="fixed top-12 right-5 z-[1000] flex flex-col gap-2.5 pointer-events-none">
		{#each toasts as toast (toast.id)}
			<div
				class="rounded px-4 py-3 shadow-lg text-xs min-w-[250px] max-w-[400px] pointer-events-auto transition-opacity duration-300 {toast.type === 'success'
					? 'bg-emerald-50 border border-emerald-500 border-l-4 border-l-emerald-500 text-emerald-600'
					: 'bg-white border border-red-500 border-l-4 border-l-red-500 text-red-600'}"
			>
				{toast.message}
			</div>
		{/each}
	</div>

	<!-- Navbar -->
	<nav class="bg-slate-800 text-white py-3 px-4 flex items-center sticky top-0 z-[100]">
		<div class="font-bold text-lg">Kyro Reports</div>
	</nav>

	<div class="max-w-300 mx-auto px-3 pb-10">

		<!-- ── Compile Report ─────────────────────────────────────────────────── -->
		<section class="bg-white border border-gray-200 rounded-md mb-4 mt-4 overflow-hidden shadow-sm">
			<h2 class="bg-slate-800 text-white px-4 py-2.5 m-0 text-xs font-semibold tracking-wide">
				Compile Report
			</h2>
			<div class="px-4 pt-4 pb-4">
				<!-- Upload Template -->
				<div class="mb-3 w-full">
					<h3 class="text-xs my-1.5 mx-0">
						1. Upload Template
						<span class="text-gray-500 text-[10px] ml-1">(Template: .xlsx file)</span>
					</h3>
					<input
						type="file" id="templateInput" accept=".xlsx"
						class="w-full px-1.5 py-1 mb-1 border border-gray-300 rounded text-xs"
						onchange={(e) => setTemplate((e.target as any).files?.[0])}
					/>
					{#if templateFile}
						{@const info = getTemplateInfo(templateFile)}
						<div class="px-2 py-1.5 mt-1.5 bg-green-50 rounded-md text-xs">
							<div class="flex justify-between items-center">
								<b>{templateFile.name}</b>
								<span class="text-gray-400">{formatFileSize(templateFile.content.byteLength)}</span>
							</div>
							{#if info?.schema && info.schema.length > 0}
								<div class="mt-1 flex flex-wrap gap-1">
									{#each info.schema as entry}
										<span class="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded border text-[10px] {TYPE_COLORS[entry.type]}">
											<span class="font-mono font-bold">{TYPE_ICONS[entry.type]}</span>
											{entry.column}
										</span>
									{/each}
								</div>
							{/if}
						</div>
					{/if}
				</div>

				<!-- Upload Reports ZIP -->
				<div class="mb-3 w-full">
					<h3 class="text-xs my-1.5 mx-0">
						2. Upload Reports ZIP
						<span class="text-gray-500 text-[10px] ml-1">(Reports: .zip containing multiple .xlsx files)</span>
					</h3>
					<input
						type="file" id="reportsInput" accept=".zip"
						class="w-full px-1.5 py-1 mb-1 border border-gray-300 rounded text-xs"
						onchange={(e) => setReportsZip((e.target as any).files?.[0])}
					/>
					{#if reportFiles.length > 0}
						<div class="mt-1.5 space-y-1">
							{#each reportFiles as report, i}
								<div class="px-2 py-1.5 rounded-md flex justify-between items-center text-xs {report.valid ? 'bg-green-50 border-l-2 border-l-emerald-500' : 'bg-red-50 border-l-2 border-l-red-500'}">
									<span>
										{report.file.name} ({formatFileSize(report.file.content.byteLength)})
										<span class="text-[10px] ml-1.5 {report.valid ? 'text-emerald-500' : 'text-red-500'}">
											{report.valid ? '✓ Valid' : `✗ ${report.validationError || 'Invalid'}`}
										</span>
										{#if report.originalName}
											<small class="ml-1 text-gray-400">from {report.originalName}</small>
										{/if}
									</span>
									<button class="bg-transparent border-none text-red-500 text-xs px-1 py-0.5 m-0" onclick={() => removeReport(i)}>×</button>
								</div>
							{/each}
						</div>
					{/if}
				</div>

				<!-- Progress and Download -->
				<div class="mt-3 w-full">
					{#if compiling}
						<div class="w-full h-1.5 bg-gray-200 rounded-md mt-1.5 overflow-hidden">
							<div class="h-full bg-blue-600 transition-all duration-300" style="width: {progress}%"></div>
						</div>
					{/if}
					<button
						onclick={compileAndDownload} disabled={compiling}
						class="w-full px-4 py-2.5 text-xs font-semibold my-2 min-h-9 bg-emerald-500 hover:bg-emerald-600 disabled:bg-gray-300 disabled:cursor-not-allowed flex items-center justify-center leading-none"
					>
						DOWNLOAD COMPILED XLSX
					</button>
				</div>
			</div>
		</section>

		<!-- ── Validate Report ────────────────────────────────────────────────── -->
		<section class="bg-white border border-gray-200 rounded-md mb-4 overflow-hidden shadow-sm">
			<h2 class="bg-slate-800 text-white px-4 py-2.5 m-0 text-xs font-semibold tracking-wide">
				Validate Report
			</h2>
			<div class="px-4 pt-4 pb-4">
				<div class="mb-3 w-full">
					<h3 class="text-xs my-1.5 mx-0">
						1. Upload Template
						<span class="text-gray-500 text-[10px] ml-1">(Template: .xlsx file)</span>
					</h3>
					<input
						type="file" id="validateTemplateInput" accept=".xlsx"
						class="w-full px-1.5 py-1 mb-1 border border-gray-300 rounded text-xs"
						onchange={(e) => setValidateTemplate((e.target as any).files?.[0])}
					/>
					{#if validateTemplateFile}
						{@const info = getTemplateInfo(validateTemplateFile)}
						<div class="px-2 py-1.5 mt-1.5 bg-green-50 rounded-md text-xs">
							<div class="flex justify-between items-center">
								<b>{validateTemplateFile.name}</b>
								<span class="text-gray-400">{formatFileSize(validateTemplateFile.content.byteLength)}</span>
							</div>
							{#if info?.schema && info.schema.length > 0}
								<div class="mt-1 flex flex-wrap gap-1">
									{#each info.schema as entry}
										<span class="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded border text-[10px] {TYPE_COLORS[entry.type]}">
											<span class="font-mono font-bold">{TYPE_ICONS[entry.type]}</span>
											{entry.column}
										</span>
									{/each}
								</div>
							{/if}
						</div>
					{/if}
				</div>

				<div class="mb-3 w-full">
					<h3 class="text-xs my-1.5 mx-0">
						2. Upload Report <span class="text-gray-500 text-[10px] ml-1">(Report: .xlsx file)</span>
					</h3>
					<input
						type="file" id="validateReportInput" accept=".xlsx"
						class="w-full px-1.5 py-1 mb-1 border border-gray-300 rounded text-xs"
						onchange={(e) => setValidateReport((e.target as any).files?.[0])}
					/>
					{#if validationResult}
						<div class="px-2 py-1.5 mt-1.5 rounded-md flex justify-between items-center text-xs {validationResult.isValid ? 'bg-green-50 border-l-2 border-l-emerald-500' : 'bg-red-50 border-l-2 border-l-red-500'}">
							{@html validationResult.message}
						</div>
					{/if}
				</div>
			</div>
		</section>

		<!-- ── Template Report ────────────────────────────────────────────────── -->
		<section class="bg-white border border-gray-200 rounded-md mb-4 overflow-hidden shadow-sm">
			<h2 class="bg-slate-800 text-white px-4 py-2.5 m-0 text-xs font-semibold tracking-wide">
				Template Report
			</h2>
			<div class="px-4 pt-4 pb-4">
				<div class="w-full">
					<!-- Combined column inputs + type dropdowns, single scrollable container -->
					<div class="my-2 px-2 pt-2 pb-2 border border-gray-300 rounded-md bg-gray-50 overflow-x-auto overflow-y-visible">
						<div class="inline-flex flex-nowrap gap-0 items-stretch p-0 min-w-max">
							{#each columns as col, index}
								<div class="group relative inline-flex flex-col flex-shrink-0">
									<!-- Name input -->
									<input
										type="text"
										bind:value={col.value}
										placeholder="Column {index + 1}..."
										class="column-input-field min-w-[80px] max-w-[180px] w-auto h-7 px-2 border border-gray-300 border-r-0 border-b-0 text-[11px] bg-white m-0 focus:outline-none focus:border-blue-600 focus:z-10 focus:shadow-[0_0_0_1px_#2563eb] {index === 0 ? 'rounded-tl-sm' : ''} {index === columns.length - 1 ? 'rounded-tr-sm !border-r border-gray-300' : ''}"
										use:autoResizeInput
									/>
									<!-- Type dropdown -->
									<div class="relative">
										<select
											bind:value={col.type}
											class="h-6 w-full pl-1 pr-4 text-[10px] border border-gray-300 border-r-0 appearance-none cursor-pointer focus:outline-none focus:border-blue-600 focus:z-10
												{col.type === 'numeric' ? 'bg-amber-50 text-amber-700' : col.type === 'date' ? 'bg-purple-50 text-purple-700' : 'bg-blue-50 text-blue-700'}
												{index === 0 ? 'rounded-bl-sm' : ''}
												{index === columns.length - 1 ? 'rounded-br-sm !border-r border-gray-300' : ''}"
										>
											<option value="string">Aa string</option>
											<option value="numeric">123 numeric</option>
											<option value="date">📅 date</option>
										</select>
										<span class="pointer-events-none absolute right-1 top-1/2 -translate-y-1/2 text-gray-400 text-[8px]">▾</span>
									</div>
									<!-- Remove button -->
									<button
										onclick={() => removeColumn(index)}
										class="absolute -top-2 -right-px bg-red-500 text-white border-none rounded-full w-[13px] h-[13px] text-[9px] leading-none cursor-pointer flex items-center justify-center transition-all duration-150 opacity-0 group-hover:opacity-100 hover:bg-red-600 hover:scale-110 z-10 p-0 m-0 min-h-0"
										title="Remove column"
									>×</button>
									{#if index > 0}
										<button
											onclick={() => moveColumnLeft(index)}
											class="absolute top-1/2 -translate-y-1/2 -left-px bg-indigo-500 text-white border-none rounded-full w-[13px] h-[13px] text-[8px] leading-none cursor-pointer flex items-center justify-center transition-all duration-150 opacity-0 group-hover:opacity-100 hover:bg-indigo-600 hover:scale-110 z-10 p-0 m-0 min-h-0 font-bold"
											title="Move left"
										>←</button>
									{/if}
									{#if index < columns.length - 1}
										<button
											onclick={() => moveColumnRight(index)}
											class="absolute top-1/2 -translate-y-1/2 -right-px bg-indigo-500 text-white border-none rounded-full w-[13px] h-[13px] text-[8px] leading-none cursor-pointer flex items-center justify-center transition-all duration-150 opacity-0 group-hover:opacity-100 hover:bg-indigo-600 hover:scale-110 z-10 p-0 m-0 min-h-0 font-bold"
											title="Move right"
										>→</button>
									{/if}
								</div>
							{/each}
							<button
								onclick={addColumn}
								class="bg-emerald-500 text-white border border-emerald-600 rounded-md w-[22px] h-[22px] text-sm leading-none cursor-pointer inline-flex items-center justify-center flex-shrink-0 p-0 m-0 ml-1 min-h-0 transition-colors hover:bg-emerald-600 self-center"
								title="Add column"
							>+</button>
						</div>
					</div>

					<p class="text-gray-500 text-[10px] mt-1.5 mb-0">{columns.length} column(s)</p>
				</div>

				<button
					onclick={createTemplate}
					class="w-full px-4 py-2.5 text-xs font-semibold my-2 min-h-9 bg-emerald-500 hover:bg-emerald-600 flex items-center justify-center leading-none"
				>
					DOWNLOAD TEMPLATE XLSX
				</button>
			</div>
		</section>
	</div>
</div>