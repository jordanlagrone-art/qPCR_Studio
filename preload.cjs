const {contextBridge,ipcRenderer}=require('electron');
contextBridge.exposeInMainWorld('qpcrPython',{
  request:(action,payload)=>ipcRenderer.invoke('python:request',action,payload)
});
