const { app, BrowserWindow } = require('electron');
const path = require('path');
function createWindow() {
  new BrowserWindow({ width:1400, height:900, icon: path.join(__dirname,'icon.ico'),
    webPreferences:{ nodeIntegration:false } })
    .loadURL(`file://${path.join(__dirname,'../build/index.html')}`);
}
app.whenReady().then(createWindow);
app.on('window-all-closed', ()=>{ if(process.platform!=='darwin') app.quit(); });