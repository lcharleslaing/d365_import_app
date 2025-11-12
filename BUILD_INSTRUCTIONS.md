# D365 Import App - Build Instructions

This document provides detailed instructions for building and distributing the D365 Import App as a portable Windows executable.

## Table of Contents
- [Prerequisites](#prerequisites)
- [Initial Setup](#initial-setup)
- [Building the App](#building-the-app)
- [Testing the Build](#testing-the-build)
- [Creating a Distribution Package](#creating-a-distribution-package)
- [Distribution to Users](#distribution-to-users)
- [Troubleshooting](#troubleshooting)
- [Future Updates](#future-updates)

---

## Prerequisites

Before building, ensure you have:

1. **Node.js** (v16 or higher)
   - Download from: https://nodejs.org/
   - Verify installation: `node --version`

2. **npm** (comes with Node.js)
   - Verify installation: `npm --version`

3. **Git Bash** (for Windows, recommended)
   - Or use PowerShell/Command Prompt

---

## Initial Setup

### 1. Clone or Navigate to Project Directory

```bash
cd /c/KemcoDev/D365_Import_App
```

### 2. Install Dependencies

```bash
npm install
```

This installs:
- `electron` - The Electron framework
- `electron-packager` - Tool for packaging the app
- `cross-env` - Cross-platform environment variable setting
- Other required dependencies

**Note:** If you encounter permission errors, you may need to run Git Bash as Administrator (right-click → Run as administrator).

---

## Building the App

### Method 1: Using electron-packager (Recommended - No Admin Required)

This method avoids code signing issues and doesn't require administrator privileges.

#### Step 1: Build the App

```bash
npm run build:packager
```

**What this does:**
- Packages your app for Windows (x64 architecture)
- Creates an unpacked application in `build-output/D365 Import App-win32-x64/`
- Includes all necessary Electron runtime files
- Packages your code into an `app.asar` file for distribution

#### Step 2: Verify Build Output

After building, you should see:
```
build-output/
  └── D365 Import App-win32-x64/
      ├── D365 Import App.exe  ← Main executable
      ├── resources/
      │   └── app.asar  ← Your packaged app code
      ├── locales/
      └── [other Electron runtime files]
```

---

### Method 2: Create Distribution ZIP (One Command)

To build and create a ZIP file in one step:

```bash
npm run package
```

This will:
1. Build the app using `electron-packager`
2. Create a ZIP file: `build-output/D365-Import-App-Portable.zip`

**Output:**
- `build-output/D365-Import-App-Portable.zip` (ready to distribute)

---

## Testing the Build

### Before Distribution

1. **Navigate to the build folder:**
   ```bash
   cd build-output/D365\ Import\ App-win32-x64
   ```

2. **Run the executable:**
   - Double-click `D365 Import App.exe`
   - Or from command line: `./D365\ Import\ App.exe`

3. **Test all features:**
   - ✅ Create a new job
   - ✅ Add Heater/Tank/Pump configurations
   - ✅ Generate part numbers
   - ✅ Save/load jobs
   - ✅ Export PDF
   - ✅ All CRUD operations

4. **Test the ZIP file:**
   - Extract `D365-Import-App-Portable.zip` to a temporary location
   - Run the app from the extracted folder
   - Verify everything works

---

## Creating a Distribution Package

### Option 1: Manual ZIP Creation

1. Navigate to `build-output/`
2. Right-click `D365 Import App-win32-x64` folder
3. Select "Send to" → "Compressed (zipped) folder"
4. Rename to `D365-Import-App-Portable.zip`

### Option 2: Automated ZIP Creation

```bash
npm run package
```

This automatically creates the ZIP file after building.

---

## Distribution to Users

### What to Share

Share the ZIP file: `D365-Import-App-Portable.zip` (approximately 205MB)

### User Instructions

Provide these instructions to your users:

```
D365 Import App - Installation Instructions

1. EXTRACT the ZIP file
   - Right-click "D365-Import-App-Portable.zip"
   - Select "Extract All..."
   - Choose a location (e.g., Desktop or Documents)

2. OPEN the extracted folder
   - Navigate to: D365 Import App-win32-x64

3. RUN the application
   - Double-click "D365 Import App.exe"
   - The app will start immediately

4. FIRST RUN SECURITY WARNING
   - Windows may show: "Windows protected your PC"
   - Click "More info"
   - Click "Run anyway"
   - This is normal for unsigned applications

NO INSTALLATION REQUIRED - Just extract and run!

The app saves all data locally in:
- Documents/D365/D365_Jobs.json
- PDF exports go to your selected Jobs folder
```

### Distribution Methods

- **Email:** If your email allows large attachments (205MB)
- **Network Share:** Place on a shared drive
- **Cloud Storage:** Upload to OneDrive, Google Drive, Dropbox, etc.
- **USB Drive:** Copy ZIP file to USB and share

---

## Troubleshooting

### Build Issues

#### Issue: "electron-packager: command not found"
**Solution:**
```bash
npm install
```

#### Issue: "Permission denied" errors
**Solution:**
- Close any File Explorer windows with the project folder open
- Try running Git Bash as Administrator (if possible)
- Or use PowerShell/Command Prompt instead

#### Issue: Build output folder is locked
**Solution:**
```bash
npm run clean
npm run build:packager
```

### Runtime Issues

#### Issue: Windows SmartScreen Warning
**Solution:** This is expected for unsigned apps. Users should:
1. Click "More info"
2. Click "Run anyway"
3. The app is safe - it's your own code

#### Issue: App won't start
**Solution:**
- Ensure all files in `D365 Import App-win32-x64/` are present
- Don't delete any files from the folder
- Try running as Administrator
- Check Windows Event Viewer for errors

#### Issue: PDF export doesn't work
**Solution:**
- Ensure the app has write permissions to the selected folder
- Try selecting a different folder (e.g., Desktop or Documents)

---

## Future Updates

### When You Make Code Changes

1. **Update your code** in:
   - `index.html`
   - `main.js`
   - `preload.js`
   - Any other source files

2. **Rebuild the app:**
   ```bash
   npm run package
   ```

3. **Test the new build:**
   - Extract and test the new ZIP
   - Verify all changes work correctly

4. **Distribute the updated ZIP:**
   - Share the new `D365-Import-App-Portable.zip`
   - Users can replace their old folder with the new one
   - Their data (`D365_Jobs.json`) will be preserved

### Version Management

Consider adding version numbers to your ZIP files:
- `D365-Import-App-v1.0.zip`
- `D365-Import-App-v1.1.zip`
- etc.

Update the version in `package.json` before building:
```json
{
  "version": "1.0.0"
}
```

---

## Build Scripts Reference

Available npm scripts:

| Command | Description |
|---------|-------------|
| `npm start` | Run the app in development mode |
| `npm run build:packager` | Build the app using electron-packager |
| `npm run package` | Build and create ZIP file automatically |
| `npm run clean` | Remove all build output folders |
| `npm run build:portable` | Build using electron-builder (requires admin) |
| `npm run build:dir` | Build directory only (electron-builder) |

---

## File Structure

```
D365_Import_App/
├── index.html          # Main application UI
├── main.js             # Electron main process
├── preload.js          # Electron preload script
├── package.json        # Project configuration
├── BUILD_INSTRUCTIONS.md  # This file
├── README.md           # Project documentation
└── build-output/       # Build output directory
    ├── D365 Import App-win32-x64/  # Packaged app
    └── D365-Import-App-Portable.zip  # Distribution ZIP
```

---

## Notes

- **File Size:** The ZIP file is ~205MB because it includes the entire Electron runtime
- **Portability:** The app is completely portable - no registry entries or system files
- **Data Storage:** User data is stored in `Documents/D365/` - safe to delete/replace the app folder
- **Updates:** Users can simply replace the app folder with a new version - their data persists
- **Security:** The app is unsigned, so Windows will show a warning. This is normal and expected.

---

## Support

For issues or questions:
- Check the troubleshooting section above
- Review the main README.md
- Check Electron documentation: https://www.electronjs.org/

---

**Last Updated:** November 2024
**Build Method:** electron-packager (no code signing required)

