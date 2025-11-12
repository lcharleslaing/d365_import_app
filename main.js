const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs').promises;
const fsSync = require('fs');

let mainWindow;
let lastPDFPath = null;

function createWindow() {
    const iconPath = path.join(__dirname, 'build', 'icon.ico');
    const windowOptions = {
        width: 1400,
        height: 900,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true,
            enableRemoteModule: false
        }
    };
    
    // Add icon if it exists
    try {
        if (fsSync.existsSync(iconPath)) {
            windowOptions.icon = iconPath;
        }
    } catch (e) {
        // Icon file doesn't exist, continue without it
    }
    
    mainWindow = new BrowserWindow(windowOptions);

    mainWindow.loadFile('index.html');

    // Open DevTools in development (comment out for production)
    // mainWindow.webContents.openDevTools();

    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

// IPC Handlers for file operations
ipcMain.handle('show-open-dialog', async (event, options) => {
    const result = await dialog.showOpenDialog(mainWindow, options);
    return result;
});

ipcMain.handle('show-save-dialog', async (event, options) => {
    const result = await dialog.showSaveDialog(mainWindow, options);
    return result;
});

ipcMain.handle('show-directory-dialog', async (event, options) => {
    const result = await dialog.showOpenDialog(mainWindow, {
        ...options,
        properties: ['openDirectory']
    });
    return result;
});

ipcMain.handle('read-file', async (event, filePath) => {
    try {
        const data = await fs.readFile(filePath, 'utf8');
        return { success: true, data };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('write-file', async (event, filePath, data) => {
    try {
        // Ensure directory exists
        const dir = path.dirname(filePath);
        await fs.mkdir(dir, { recursive: true });
        
        await fs.writeFile(filePath, data, 'utf8');
        return { success: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('open-folder', async (event, folderPath) => {
    try {
        await shell.openPath(folderPath);
        return { success: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('show-item-in-folder', async (event, filePath) => {
    try {
        shell.showItemInFolder(filePath);
        return { success: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('get-file-path', async (event, filePath) => {
    // Return the file path (for Electron, we'll get full paths)
    return filePath;
});

ipcMain.handle('save-pdf', async (event, pdfData, fileName, defaultPath) => {
    try {
        const result = await dialog.showSaveDialog(mainWindow, {
            title: 'Save PDF',
            defaultPath: defaultPath || fileName,
            filters: [
                { name: 'PDF Files', extensions: ['pdf'] }
            ]
        });

        if (!result.canceled && result.filePath) {
            // Convert base64 to buffer
            const buffer = Buffer.from(pdfData, 'base64');
            await fs.writeFile(result.filePath, buffer);
            
            // Store the path for "Open Location" button
            lastPDFPath = result.filePath;
            
            return { success: true, filePath: result.filePath };
        }
        return { success: false, canceled: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('get-last-pdf-path', async () => {
    return lastPDFPath;
});

ipcMain.handle('open-pdf-location', async () => {
    if (lastPDFPath) {
        try {
            shell.showItemInFolder(lastPDFPath);
            return { success: true };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }
    return { success: false, error: 'No PDF path available' };
});

