import { app, shell, BrowserWindow, ipcMain, dialog } from 'electron'
import { join } from 'path'
import fs from 'fs'
import * as XLSX from 'xlsx'
import { electronApp, optimizer, is } from '@electron-toolkit/utils'
import icon from '../Logo/RRM EMAIL APP.png?asset'
import { sendEmail, verifySmtp } from './mailer.js'

const userDataPath = app.getPath('userData')
const configPath = join(userDataPath, 'rrm-config.json')
const manageConfigPath = join(userDataPath, 'rrm-manage.json')
const templateConfigPath = join(userDataPath, 'rrm-templates.json')
const signatureConfigPath = join(userDataPath, 'rrm-signatures.json')
const licenseConfigPath = join(userDataPath, 'rrm-license.json')

// ── Generic helpers ─────────────────────────────────────────────────────────
function saveData(fileName, data) {
  const filePath = join(userDataPath, fileName)
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf-8')
}

function loadData(fileName, fallback = null) {
  const filePath = join(userDataPath, fileName)
  if (!fs.existsSync(filePath)) return fallback
  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf-8'))
  } catch {
    return fallback
  }
}

function createWindow() {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 900,
    minHeight: 700,
    show: false,
    resizable: true,
    autoHideMenuBar: true,
    icon,
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      sandbox: false
    }
  })

  mainWindow.on('ready-to-show', () => {
    mainWindow.show()
  })

  mainWindow.webContents.setWindowOpenHandler((details) => {
    shell.openExternal(details.url)
    return { action: 'deny' }
  })

  // HMR for renderer base on electron-vite cli.
  // Load the remote URL for development or the local html file for production.
  if (is.dev && process.env['ELECTRON_RENDERER_URL']) {
    mainWindow.loadURL(process.env['ELECTRON_RENDERER_URL'])
  } else {
    mainWindow.loadFile(join(__dirname, '../renderer/index.html'))
  }
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(() => {
  // Set app user model id for windows
  electronApp.setAppUserModelId('com.electron')

  // Default open or close DevTools by F12 in development
  // and ignore CommandOrControl + R in production.
  // see https://github.com/alex8088/electron-toolkit/tree/master/packages/utils
  app.on('browser-window-created', (_, window) => {
    optimizer.watchWindowShortcuts(window)
  })

  // IPC test
  ipcMain.on('ping', () => console.log('pong'))

  ipcMain.handle('dialog:selectFolder', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
      properties: ['openDirectory']
    })
    if (canceled) {
      return null
    }

    const rootDir = filePaths[0]
    try {
      const list = fs.readdirSync(rootDir, { withFileTypes: true })
      const subfolders = list
        .filter((dirent) => dirent.isDirectory())
        .map((dirent) => {
          const folderObj = { name: dirent.name, innerFolder: null }
          try {
            const innerPath = join(rootDir, dirent.name)
            const innerList = fs.readdirSync(innerPath, { withFileTypes: true })
            const innerDirs = innerList.filter((d) => d.isDirectory())

            // Find the specific folder that contains "Claim Tag No"
            const targetDir = innerDirs.find((d) => d.name.toLowerCase().includes('claim tag no'))
            if (targetDir) {
              folderObj.innerFolder = targetDir.name
            }
          } catch {
            // gracefully ignore if we can't read inner directories
          }
          return folderObj
        })
      return { rootDir, subfolders }
    } catch (error) {
      console.error('Failed to read directory:', error)
      return { rootDir, subfolders: [] }
    }
  })

  const getRecursiveFiles = (currentPath, emailType, returnStats = true) => {
    const filesList = []

    const scan = (dir) => {
      const dirName = dir.split(/[\\/]/).pop() || ''
      let prefix = ''
      // Only apply prefix for subfolders, not the main folder (root of the scan)
      if (dir !== currentPath) {
        const dirLower = dirName.toLowerCase()
        if (dirLower.includes('claim')) {
          prefix = 'Claim - '
        } else if (dirLower.includes('tagging')) {
          prefix = 'HC - '
        } else if (['Claim', 'Retagging'].includes(emailType) && dirLower.includes('new')) {
          prefix = 'RT - '
        } else if (['Claim', 'Retagging'].includes(emailType) && dirLower.includes('old')) {
          prefix = 'HC - '
        }
      }

      try {
        const entries = fs.readdirSync(dir, { withFileTypes: true })
        for (const entry of entries) {
          const fullPath = join(dir, entry.name)

          if (entry.isDirectory()) {
            // Folder exclusions
            const entryNameLower = entry.name.toLowerCase()
            if (emailType === 'Tagging') {
              if (entry.name.toUpperCase().startsWith('HC')) continue
              if (entryNameLower === 'rejected' || entryNameLower === 'rejection') continue
            }
            if ((emailType === 'Claim' || emailType === 'Retagging') && entryNameLower === 'extra') continue

            scan(fullPath)
          } else if (entry.isFile()) {
            const nameUpper = entry.name.toUpperCase()
            const ext = entry.name.split('.').pop()?.toLowerCase() || ''

            // Global exclusions
            if (['zip', 'rar'].includes(ext)) continue

            // File exclusions
            if (emailType === 'Tagging') {
              if (nameUpper.startsWith('DETAIL')) continue
              const allowedExts = ['pdf', 'jpg', 'jpeg', 'png', 'xls', 'xlsx']
              if (!allowedExts.includes(ext)) continue
            } else if (emailType === 'Retagging') {
              if (nameUpper.includes('DETAIL')) continue
            }

            let attachmentName = entry.name
            if (prefix && ['jpeg', 'jpg', 'png'].includes(ext)) {
              if (!attachmentName.startsWith(prefix)) {
                attachmentName = prefix + attachmentName
              }
            }

            if (returnStats) {
              const stat = fs.statSync(fullPath)
              filesList.push({ name: attachmentName, size: stat.size, ext })
            } else {
              filesList.push({ name: attachmentName, path: fullPath })
            }
          }
        }
      } catch {
        // ignore errors during recursive scan
      }
    }

    scan(currentPath)
    return filesList
  }

  ipcMain.handle('get-attachments', async (_, folderPath, emailType) => {
    return getRecursiveFiles(folderPath, emailType, true)
  })

  ipcMain.handle('send-email', async (_, payload) => {
    const { smtpConfig, to, cc, subject, body } = payload
    const { folderPath, emailType, html, excludedAttachments } = payload
    const allAttachments = getRecursiveFiles(folderPath, emailType, false)
    const filteredAttachments =
      excludedAttachments && excludedAttachments.length > 0
        ? allAttachments.filter((a) => !excludedAttachments.includes(a.name))
        : allAttachments
    return await sendEmail({
      smtpConfig,
      to,
      cc,
      subject,
      body,
      attachments: filteredAttachments,
      html
    })
  })

  ipcMain.handle('test-smtp-connection', async (_, smtpConfig) => {
    return await verifySmtp(smtpConfig)
  })

  ipcMain.handle('save-settings', async (_, config) => {
    try {
      fs.writeFileSync(configPath, JSON.stringify(config))
      return { success: true }
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  ipcMain.handle('load-settings', async () => {
    try {
      if (fs.existsSync(configPath)) {
        const data = fs.readFileSync(configPath, 'utf-8')
        return JSON.parse(data)
      }
      return null
    } catch {
      return null
    }
  })

  ipcMain.handle('save-manage-data', async (_, data) => {
    try {
      fs.writeFileSync(manageConfigPath, JSON.stringify(data))
      return { success: true }
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  ipcMain.handle('load-manage-data', async () => {
    try {
      if (fs.existsSync(manageConfigPath)) {
        const data = fs.readFileSync(manageConfigPath, 'utf-8')
        return JSON.parse(data)
      }
      return []
    } catch {
      return []
    }
  })

  ipcMain.handle('save-templates', async (_, templates) => {
    try {
      fs.writeFileSync(templateConfigPath, JSON.stringify(templates))
      return { success: true }
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  ipcMain.handle('load-templates', async () => {
    try {
      if (fs.existsSync(templateConfigPath)) {
        const data = fs.readFileSync(templateConfigPath, 'utf-8')
        return JSON.parse(data)
      }
      return { tagging: '', retagging: '', claim: '' }
    } catch {
      return { tagging: '', retagging: '', claim: '' }
    }
  })

  // New: Get app version
  ipcMain.handle('get-app-version', async () => {
    try {
      const pkgPath = join(__dirname, '../../package.json')
      const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf-8'))
      return { version: pkg.version }
    } catch (e) {
      console.error('Failed to read package.json version:', e)
      return { version: 'unknown' }
    }
  })

  // ── License (product key) ────────────────────────────────────────────────
  ipcMain.handle('save-license', async (_, licenseData) => {
    try {
      saveData('rrm-license.json', licenseData)
      return { success: true }
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  ipcMain.handle('load-license', async () => {
    try {
      if (fs.existsSync(licenseConfigPath)) {
        return JSON.parse(fs.readFileSync(licenseConfigPath, 'utf-8'))
      }
      return null
    } catch {
      return null
    }
  })

  // ── Generic save/load IPC ────────────────────────────────────────────────
  ipcMain.handle('save-data', async (_, fileName, data) => {
    try {
      saveData(fileName, data)
      return { success: true }
    } catch (e) {
      return { success: false, error: e.message }
    }
  })

  ipcMain.handle('load-data', async (_, fileName, fallback = null) => {
    return loadData(fileName, fallback)
  })

  ipcMain.handle('save-signatures', async (_, signatures) => {
    try {
      if (fs.existsSync(signatureConfigPath)) {
        fs.copyFileSync(signatureConfigPath, signatureConfigPath + '.bak')
      }
      fs.writeFileSync(signatureConfigPath, JSON.stringify(signatures))
      return { success: true }
    } catch (e) {
      console.error('Error saving signatures:', e)
      return { success: false, error: e.message }
    }
  })

  ipcMain.handle('load-signatures', async () => {
    try {
      if (fs.existsSync(signatureConfigPath)) {
        const data = fs.readFileSync(signatureConfigPath, 'utf-8')
        return JSON.parse(data)
      }
      return null
    } catch (e) {
      console.error('Error loading signatures:', e)
      return null
    }
  })

  ipcMain.handle('get-email-ids', async (_, folderPath, emailType) => {
    try {
      const files = fs.readdirSync(folderPath)
      const dataFile = files.find((f) =>
        f.toLowerCase().includes(`${emailType.toLowerCase()} data`)
      )
      if (!dataFile) return null

      const fullPath = join(folderPath, dataFile)
      const workbook = XLSX.readFile(fullPath)
      const sheet = workbook.Sheets[workbook.SheetNames[0]]

      let mapping = { insurance: '', bank: '', leader: '' }

      if (emailType === 'Tagging') {
        mapping = {
          insurance: sheet['AL2']?.v || '',
          bank: sheet['AM2']?.v || '',
          leader: sheet['AN2']?.v || ''
        }
      } else if (emailType === 'Retagging') {
        mapping = {
          insurance: sheet['Y2']?.v || '',
          bank: sheet['Z2']?.v || '',
          leader: sheet['AA2']?.v || ''
        }
      } else if (emailType === 'Claim') {
        mapping = {
          insurance: sheet['AI2']?.v || '',
          bank: sheet['AJ2']?.v || '',
          leader: sheet['AK2']?.v || '',
          claimDetails: {
            claimRegNo: sheet['AH2']?.v || '',
            nameOfBeneficiary: sheet['B2']?.v || '',
            phoneNo: sheet['C2']?.v || '',
            address: sheet['D2']?.v || '',
            village: sheet['E2']?.v || '',
            taluka: sheet['G2']?.v || '',
            district: sheet['H2']?.v || '',
            state: sheet['I2']?.v || '',
            pinCode: sheet['J2']?.v || '',
            dateOfDeath: sheet['K2']?.v || '',
            timeOfDeath: sheet['L2']?.v || '',
            tagNo: sheet['M2']?.v || '',
            animalSpecies: sheet['N2']?.v || '',
            sumInsured: sheet['V2']?.v || '',
            insuredNameFinancingBank: sheet['AA2']?.v || '',
            insuredAddressBankBranch: sheet['AB2']?.v || '',
            proposalLoanAccountNo: sheet['AC2']?.v || '',
            insuranceCompany: sheet['AD2']?.v || '',
            policyNumber: sheet['AE2']?.v || '',
            policyDate: sheet['AF2']?.v || '',
            surveyerName: sheet['AL2']?.v || ''
          }
        }
      }

      return mapping
    } catch (e) {
      console.error('Error fetching email IDs:', e)
      return null
    }
  })

  createWindow()

  app.on('activate', function () {
    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.
