"""Exercise the frozen engine, including dynamically imported export backends."""
import base64
import json
import pathlib
import subprocess

exe=pathlib.Path(__file__).resolve().parents[1]/'python-dist/v1.1.0/qpcr-engine/qpcr-engine.exe'
payload={'rows':[dict(gene='GENE',condition='Control',trialId='t1',trial='Trial 1',value=1),dict(gene='GENE',condition='Treatment',trialId='t1',trial='Trial 1',value=2)],'genes':['GENE'],'controls':{'graphTitle':'Frozen engine verification'}}
requests=[{'id':i,'action':'render','payload':dict(payload,format=fmt,export=True)} for i,fmt in enumerate(['png','svg','pdf'])]
process=subprocess.run([str(exe)],input=''.join(json.dumps(r)+'\n' for r in requests),text=True,encoding='utf-8',capture_output=True,timeout=90,creationflags=subprocess.CREATE_NO_WINDOW)
assert process.returncode==0,process.stderr
responses=[json.loads(line) for line in process.stdout.splitlines()]
assert len(responses)==3,process.stderr
for response,fmt in zip(responses,['png','svg','pdf']):
    assert 'error' not in response,response
    data=base64.b64decode(response['result']['data'])
    assert len(data)>1000
    assert {'png':b'\x89PNG','svg':b'<svg','pdf':b'%PDF'}[fmt] in data
    print(f'Frozen {fmt.upper()} export passed ({len(data)} bytes)')

