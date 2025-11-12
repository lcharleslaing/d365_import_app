const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods that allow the renderer process to use
// the ipcRenderer without exposing the entire object
contextBridge.exposeInMainWorld('electronAPI', {
    // File dialogs
    showOpenDialog: (options) => ipcRenderer.invoke('show-open-dialog', options),
    showSaveDialog: (options) => ipcRenderer.invoke('show-save-dialog', options),
    showDirectoryDialog: (options) => ipcRenderer.invoke('show-directory-dialog', options),
    
    // File operations
    readFile: (filePath) => ipcRenderer.invoke('read-file', filePath),
    writeFile: (filePath, data) => ipcRenderer.invoke('write-file', filePath, data),
    
    // Folder operations
    openFolder: (folderPath) => ipcRenderer.invoke('open-folder', folderPath),
    showItemInFolder: (filePath) => ipcRenderer.invoke('show-item-in-folder', filePath),
    getFilePath: (filePath) => ipcRenderer.invoke('get-file-path', filePath),
    
    // PDF operations
    savePDF: (pdfData, fileName, defaultPath) => ipcRenderer.invoke('save-pdf', pdfData, fileName, defaultPath),
    getLastPDFPath: () => ipcRenderer.invoke('get-last-pdf-path'),
    openPDFLocation: () => ipcRenderer.invoke('open-pdf-location'),
    
    // Platform info
    platform: process.platform
});

