import { contextBridge, ipcRenderer } from 'electron'
import { electronAPI } from '@electron-toolkit/preload'

// Custom APIs for renderer
const api = {
  getAppVersion: () => ipcRenderer.invoke('get-app-version'),
  selectFolder: () => ipcRenderer.invoke('dialog:selectFolder'),
  getAttachments: (path, type) => ipcRenderer.invoke('get-attachments', path, type),
  sendEmail: (payload) => ipcRenderer.invoke('send-email', payload),
  saveSettings: (config) => ipcRenderer.invoke('save-settings', config),
  loadSettings: () => ipcRenderer.invoke('load-settings'),
  testSmtpConnection: (smtpConfig) => ipcRenderer.invoke('test-smtp-connection', smtpConfig),
  saveManageData: (data) => ipcRenderer.invoke('save-manage-data', data),
  loadManageData: () => ipcRenderer.invoke('load-manage-data'),
  saveTemplates: (templates) => ipcRenderer.invoke('save-templates', templates),
  loadTemplates: () => ipcRenderer.invoke('load-templates'),
  saveSignatures: (sigs) => ipcRenderer.invoke('save-signatures', sigs),
  loadSignatures: () => ipcRenderer.invoke('load-signatures'),
  getEmailIds: (path, type) => ipcRenderer.invoke('get-email-ids', path, type),
  // License (product key)
  saveLicense: (data) => ipcRenderer.invoke('save-license', data),
  loadLicense: () => ipcRenderer.invoke('load-license'),
  // Generic file-based storage
  saveData: (fileName, data) => ipcRenderer.invoke('save-data', fileName, data),
  loadData: (fileName, fallback) => ipcRenderer.invoke('load-data', fileName, fallback)
}

// Use `contextBridge` APIs to expose Electron APIs to
// renderer only if context isolation is enabled, otherwise
// just add to the DOM global.
if (process.contextIsolated) {
  try {
    contextBridge.exposeInMainWorld('electron', electronAPI)
    contextBridge.exposeInMainWorld('api', api)
  } catch (error) {
    console.error(error)
  }
} else {
  window.electron = electronAPI
  window.api = api
}
