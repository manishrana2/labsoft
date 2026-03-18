const { app, BrowserWindow, shell } = require('electron')
const path = require('node:path')

const DEFAULT_APP_URL = 'https://manishrana2-labsoft.onrender.com'
const appUrl = process.env.LABSOFT_APP_URL || DEFAULT_APP_URL

const createWindow = () => {
  const win = new BrowserWindow({
    width: 1366,
    height: 840,
    minWidth: 1024,
    minHeight: 700,
    autoHideMenuBar: true,
    title: 'Labsoft',
    webPreferences: {
      contextIsolation: true,
      sandbox: true
    }
  })

  win.loadURL(appUrl)

  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url)
    return { action: 'deny' }
  })
}

app.whenReady().then(() => {
  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})
