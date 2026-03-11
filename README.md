# Kyro Reports

A modern web application for creating Excel templates and compiling multiple reports into a single consolidated document. Built with SvelteKit and runs entirely in the browser – no backend server required.

## Features

- **Template Creator**: Define custom column names and generate formatted .xlsx templates with styled headers
- **Report Compiler**: Upload a template and multiple reports (via ZIP), validate them, and compile into a single Excel file
- **Validation**: Automatic validation ensures all reports match the template structure (sheet names and columns)
- **Progress Tracking**: Visual progress bar during compilation
- **Toast Notifications**: Real-time feedback for success and error states
- **Auto-numbering**: Compiled reports include sequential row numbers for easy reference
- **Responsive Design**: Clean UI built with Tailwind CSS

## Tech Stack

- **Framework**: SvelteKit with Svelte 5
- **Styling**: Tailwind CSS v4
- **Excel Processing**: `xlsx-js-style` – Excel file generation and manipulation with styling support
- **ZIP Handling**: `JSZip` – ZIP archive handling for bulk report uploads
- **TypeScript**: Fully typed codebase for better developer experience

## Project Structure

```
src/
├── lib/
│   ├── components/          # Svelte components
│   │   ├── CompileReport.svelte
│   │   ├── ValidateReport.svelte
│   │   ├── TemplateReport.svelte
│   │   ├── ToastContainer.svelte
│   │   └── Navbar.svelte
│   ├── utils/              # Utility functions
│   │   ├── excel.ts        # Excel operations
│   │   └── toast.ts        # Toast notifications
│   ├── types.ts            # TypeScript types
│   └── assets/             # Static assets
├── routes/
│   ├── +layout.svelte       # Root layout
│   ├── +layout.css          # Global styles
│   └── +page.svelte         # Main page
```

## Getting Started

### Prerequisites

- Node.js 18+
- npm or yarn

### Installation

```bash
# Clone the repository
git clone <repository-url>
cd kyroreports

# Install dependencies
npm install

# Run development server
npm run dev

# Build for production
npm run build

# Preview production build
npm run preview
```

The application will be available at `http://localhost:5173` (or another port if 5173 is in use).

## How It Works

### Template Creator

1. Enter column names using the interactive input fields
2. Add, remove, or reorder columns as needed
3. Click **DOWNLOAD TEMPLATE XLSX** to generate a formatted Excel file
4. Distribute the template to users for data collection

### Report Compiler

1. Upload the original template .xlsx file
2. Upload a ZIP archive containing multiple filled report files (.xlsx)
3. The app validates each report against the template:
   - Checks sheet names match
   - Verifies column headers are identical
4. Click **DOWNLOAD COMPILED XLSX** to merge all reports into one file with:
   - Original headers with "No." as the first column
   - Sequential row numbers for each data row
   - Styled header row (bold, centered)

### Validate Report

1. Upload a template .xlsx file
2. Upload a single report .xlsx file to validate
3. Get instant feedback on whether the report matches the template structure

## Browser Compatibility

Requires a modern browser with ES6+ support:
- Chrome/Edge 90+
- Firefox 88+
- Safari 14+

## Development

```bash
# Run type checking
npm run check

# Run type checking in watch mode
npm run check:watch
```

## License

See [LICENSE](LICENSE) file for details.
