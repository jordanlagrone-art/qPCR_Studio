const {app,BrowserWindow,Menu,ipcMain} = require('electron');
const path=require('path');
const PythonBridge=require('./python-bridge.cjs');
const engine=new PythonBridge(app);
app.setName('qPCR studio');
function openWindow(){
  const win=new BrowserWindow({width:1440,height:980,minWidth:860,minHeight:650,title:'qPCR studio',backgroundColor:'#f7f8fa',show:false,webPreferences:{preload:path.join(__dirname,'preload.cjs'),contextIsolation:true,nodeIntegration:false,sandbox:true}});
  win.loadFile(path.join(__dirname,'src/index.html'));
  win.once('ready-to-show',()=>win.show());
  win.webContents.setWindowOpenHandler(()=>({action:'deny'}));
  win.webContents.on('will-navigate',event=>event.preventDefault());
  Menu.setApplicationMenu(Menu.buildFromTemplate([{label:'File',submenu:[{role:'quit'}]},{label:'Edit',submenu:[{role:'undo'},{role:'redo'},{type:'separator'},{role:'cut'},{role:'copy'},{role:'paste'},{role:'selectAll'}]},{label:'View',submenu:[{role:'resetZoom'},{role:'zoomIn'},{role:'zoomOut'},{role:'togglefullscreen'}]}]));
}
app.whenReady().then(openWindow);
ipcMain.handle('python:request',(event,action,payload)=>{
  if(event.senderFrame!==event.sender.mainFrame)throw new Error('Only the app window may use Python.');
  return engine.request(action,payload);
});
app.on('before-quit',()=>engine.stop());
app.on('window-all-closed',()=>{if(process.platform!=='darwin')app.quit();});
app.on('activate',()=>{if(!BrowserWindow.getAllWindows().length)openWindow();});
