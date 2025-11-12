# D365 Import Configuration Tool

A desktop application for configuring D365 import jobs with support for Heaters, Tanks, and Pumps.

## Features

- Configure multiple Heaters, Tanks, and Pumps per job
- Full CRUD operations for dropdown options
- Generate part numbers and descriptions automatically
- Export configurations to PDF
- Save/load job configurations
- Open saved PDF locations directly from the app

## Installation

### For End Users

Download the latest `.exe` installer from the releases section and run it.

### For Developers

1. **Install Node.js** (v16 or higher recommended)

2. **Install dependencies:**
   ```bash
   npm install
   ```

3. **Run the app in development mode:**
   ```bash
   npm start
   ```

4. **Build the Windows executable:**
   ```bash
   npm run build:win
   ```

   This will create an installer in the `dist` folder.

## Building the Executable

### Prerequisites

- Node.js (v16+)
- npm or yarn

### Build Steps

1. Clone the repository:
   ```bash
   git clone https://github.com/lcharleslaing/d365_import_app.git
   cd d365_import_app
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Build for Windows:

   **Option A: Portable Version (Recommended - No Admin Required)**
   ```bash
   npm run build:packager
   ```
   Creates a portable app folder in `build-output/D365 Import App-win32-x64/`. No code signing issues, no admin privileges needed.

   **Option B: Create ZIP Package**
   ```bash
   npm run package
   ```
   Builds the app and creates a ZIP file: `build-output/D365-Import-App-Portable.zip` (ready to distribute)

   **Option C: Installer Version (Requires Admin)**
   ```bash
   npm run build:win
   ```
   Creates an installer using electron-builder (may require admin privileges for code signing).

   **📖 For detailed build instructions, see [BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md)**

### Build Options

- `npm start` - Run the app in development mode
- `npm run build:packager` - Build portable app using electron-packager (recommended, no admin needed)
- `npm run package` - Build and create ZIP file automatically
- `npm run build:win` - Build Windows installer (electron-builder, may require admin)
- `npm run build:portable` - Build portable .exe (electron-builder, may require admin)
- `npm run clean` - Remove all build output folders

### ⚠️ Distribution Notes

**Windows SmartScreen Warning**: 
- Unsigned executables may show a "Windows protected your PC" warning
- Users can click "More info" → "Run anyway" to proceed
- This is normal for unsigned applications

**Recommended**: Use `npm run build:packager` or `npm run package` for easier distribution with no admin privileges required.

📖 **See [BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md) for detailed step-by-step build instructions.**

## Project Structure

```
.
├── index.html          # Main application file (HTML/CSS/JS)
├── main.js             # Electron main process
├── preload.js          # Electron preload script (IPC bridge)
├── package.json        # Node.js dependencies and build config
├── BUILD_INSTRUCTIONS.md  # Detailed build instructions
└── build-output/       # Build output (created after build)
    └── D365 Import App-win32-x64/  # Packaged app folder
```

## Usage

1. **Set Jobs Folder**: Click "Set Jobs Folder" to select the root directory where PDFs will be saved
2. **Create Job**: Fill in the main form (Job Number, Customer Name, etc.)
3. **Add Configurations**: Add Heaters, Tanks, and/or Pumps with their configurations
4. **Generate Results**: Click "Generate Results" to see generated part numbers
5. **Export PDF**: Click "Export to PDF" to save the configuration report
6. **Open PDF Location**: After saving a PDF, click "Open PDF Location" to open the file in Windows Explorer

## Development

The app uses:
- **Electron** - Desktop app framework
- **Tailwind CSS** - Styling
- **DaisyUI** - UI components
- **jsPDF** - PDF generation

## License

MIT

