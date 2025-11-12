# Distribution Guide for D365 Import App

## Potential Issues & Solutions

### ⚠️ Windows SmartScreen Warning

**Issue**: Windows may show a warning: "Windows protected your PC" when users try to install/run the app.

**Why**: The executable is not code-signed (code signing certificates cost $200-400/year).

**Solutions**:

1. **For End Users** (Easiest):
   - Click "More info" on the warning
   - Click "Run anyway"
   - Windows will remember this choice

2. **For IT Departments**:
   - Add the app to Windows Defender exclusions
   - Or use Group Policy to allow the app
   - Or distribute via company software portal

3. **Portable Version** (Recommended for easier distribution):
   - Use the portable build: `npm run build:portable`
   - This creates a single `.exe` file (no installer)
   - Users can run it directly without installation
   - Less likely to trigger SmartScreen

### 🛡️ Antivirus Software

**Issue**: Some antivirus software may flag Electron apps as suspicious.

**Solutions**:
- Add the app folder to antivirus exclusions
- Most modern antivirus software will allow it after user confirmation
- The portable version is less likely to trigger false positives

### 🏢 Corporate IT Policies

**Issue**: Corporate IT may block unsigned executables.

**Solutions**:
1. **Portable Version**: Often bypasses installer restrictions
2. **IT Approval**: Have IT whitelist the application
3. **Code Signing** (Future): Consider purchasing a code signing certificate for professional distribution

## Build Options

### Option 1: Installer (Current Default)
```bash
npm run build:win
```
- Creates an `.exe` installer in `dist/`
- Users install it like any Windows program
- Creates Start Menu and Desktop shortcuts
- May trigger SmartScreen warning (user can bypass)

### Option 2: Portable Version (Recommended)
```bash
npm run build:portable
```
- Creates a single `.exe` file in `dist/`
- No installation needed - just run the `.exe`
- Users can put it anywhere (USB drive, network share, etc.)
- Less likely to trigger security warnings
- Easier for IT to approve (no system changes)

## Distribution Recommendations

### For Internal Company Use:

1. **Use Portable Version**:
   - Build: `npm run build:portable`
   - Distribute the single `.exe` file
   - Users can run it from any location
   - No admin rights needed

2. **Provide Instructions**:
   - Include a simple README explaining:
     - How to handle SmartScreen warning (click "More info" → "Run anyway")
     - That it's safe and from your company
     - How to add to antivirus exclusions if needed

3. **Network Share Distribution**:
   - Place the `.exe` on a network share
   - Users can run it directly from there
   - IT can easily update it by replacing the file

### For External Distribution:

1. **Consider Code Signing** (if budget allows):
   - Purchase a code signing certificate (~$200-400/year)
   - Sign the executable to eliminate warnings
   - More professional appearance

2. **Use Portable Version**:
   - Easier for users to try without installation
   - Less IT friction

## Testing Before Distribution

1. Test on a clean Windows machine (without the app installed)
2. Test with Windows Defender enabled
3. Test with common antivirus software (if possible)
4. Verify the "Open PDF Location" feature works
5. Test file saving/loading functionality

## User Instructions Template

Include this with your distribution:

```
D365 Import App - Installation Instructions

1. Download the D365_Import_App.exe file

2. If Windows shows a security warning:
   - Click "More info"
   - Click "Run anyway"
   - This is normal for unsigned applications

3. For Portable Version:
   - Simply double-click the .exe file to run
   - No installation needed
   - You can move it to any folder

4. For Installer Version:
   - Run the installer
   - Follow the installation wizard
   - Launch from Start Menu or Desktop shortcut

5. If your antivirus blocks it:
   - Add the app folder to your antivirus exclusions
   - Or allow it when prompted

Need help? Contact [your IT support]
```

## Summary

**Best Approach for Internal Use**:
- Use **portable version** (`npm run build:portable`)
- Distribute the single `.exe` file
- Provide simple instructions for handling Windows warnings
- Most users can run it without IT intervention

**The portable version is recommended** because:
- ✅ No installation required
- ✅ Less likely to trigger security warnings
- ✅ Easier for IT to approve (no system changes)
- ✅ Can run from USB drive or network share
- ✅ No admin rights needed

