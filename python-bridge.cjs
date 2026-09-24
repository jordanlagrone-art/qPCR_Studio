const {spawn}=require('child_process');
const path=require('path');
const fs=require('fs');
const readline=require('readline');
class PythonBridge {
  constructor(app){this.app=app;this.pending=new Map();this.nextId=0;this.child=null;}
  start(){
    if(this.child)return;
    const packaged=path.join(this.app.isPackaged?process.resourcesPath:path.join(__dirname,'python-dist'),'qpcr-engine','qpcr-engine.exe');
    const executable=this.app.isPackaged?packaged:path.join(__dirname,'.venv','Scripts','python.exe');
    const args=this.app.isPackaged?[]:['-u',path.join(__dirname,'python','engine.py')];
    if(!fs.existsSync(executable))throw new Error('Python engine is missing. Reinstall the desktop app, or run the documented Python setup for development.');
    const child=spawn(executable,args,{stdio:['pipe','pipe','pipe'],windowsHide:true,env:{...process.env,PYTHONIOENCODING:'utf-8',MPLCONFIGDIR:path.join(this.app.getPath('userData'),'matplotlib')}});
    this.child=child;let stderr='';
    child.stderr.on('data',chunk=>{stderr=(stderr+chunk.toString()).slice(-3000);});
    const fail=error=>{if(this.child!==child)return;this.child=null;for(const [id,p] of this.pending){clearTimeout(p.timer);p.reject(error);this.pending.delete(id);}};
    child.on('error',error=>fail(new Error('Could not start Python: '+error.message)));
    child.on('exit',code=>fail(new Error('Python engine stopped ('+code+'). '+stderr)));
    child.stdin.on('error',()=>{});
    readline.createInterface({input:child.stdout}).on('line',line=>{
      try{const message=JSON.parse(line),p=this.pending.get(message.id);if(!p)return;this.pending.delete(message.id);clearTimeout(p.timer);if(message.error)p.reject(new Error(message.error));else p.resolve(message.result);}catch{fail(new Error('Invalid response from Python engine.'));child.kill();}
    });
  }
  request(action,payload){
    if(!['health','analyze','stats','render'].includes(action))return Promise.reject(new Error('Unsupported Python action.'));
    if(this.pending.size>=12)return Promise.reject(new Error('Python is busy. Please try again shortly.'));
    try{this.start();}catch(error){return Promise.reject(error);}
    return new Promise((resolve,reject)=>{const id=++this.nextId;const timer=setTimeout(()=>{this.pending.delete(id);reject(new Error('Python request timed out. Please retry.'));this.stop();},120000);this.pending.set(id,{resolve,reject,timer});this.child.stdin.write(JSON.stringify({id,action,payload})+'\n');});
  }
  stop(){if(this.child)this.child.kill();}
}
module.exports=PythonBridge;
