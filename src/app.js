(() => {
  'use strict';
  const state = {
    trials: [], raw: [], sampleMap: {}, selectedGenes: new Set(), refGene:'', calibrator:'', base:2,
    results: [], overrides:{}, analysisRun:false
  };
  const palettes = {botanical:['#547b69','#bc865b','#879cc2','#b18eb3','#c8b768','#668f9c','#a77978','#7d8580'],ocean:['#397c96','#54a9b3','#827eb7','#d09a62','#7189aa','#7fa587','#b57c97','#88949e'],accessible:['#0072b2','#e69f00','#009e73','#cc79a7','#56b4e9','#d55e00','#8a7900','#555555']};
  let colors=[...palettes.botanical];
  const $ = s => document.querySelector(s), $$ = s => [...document.querySelectorAll(s)];
  const esc = s => String(s ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));
  const num = v => { if(v === null || v === undefined || String(v).trim() === '') return null; const x = Number(v); return Number.isFinite(x) ? x : null; };
  const mean = a => { const b=a.filter(Number.isFinite); return b.length?b.reduce((s,x)=>s+x,0)/b.length:null; };
  const sd = a => { const b=a.filter(Number.isFinite); if(b.length<2)return 0; const m=mean(b); return Math.sqrt(b.reduce((s,x)=>s+(x-m)**2,0)/(b.length-1)); };
  const sem = a => { const b=a.filter(Number.isFinite); return b.length?sd(b)/Math.sqrt(b.length):0; };
  const fmt = (v,d=3) => Number.isFinite(v) ? (Math.abs(v)>=1000|| (Math.abs(v)>0&&Math.abs(v)<0.001) ? v.toExponential(3) : v.toFixed(d)) : '—';
  const unique = a => [...new Set(a.filter(v=>v!==null&&v!==undefined&&String(v).trim()!==''))];
  const slug = s => String(s).replace(/[^a-zA-Z0-9_-]/g,'_');

  // Navigation
  $$('.nav button').forEach(b=>b.addEventListener('click',()=>showPage(b.dataset.page)));
  function showPage(name){
    $$('.nav button').forEach(b=>b.classList.toggle('active',b.dataset.page===name));
    $$('.page').forEach(p=>p.classList.toggle('active',p.id===`page-${name}`));
    if((name==='graphs'||name==='results') && state.raw.length && !state.analysisRun){ if(!runAnalysis()) return; renderResults(); renderGraphControls(); } if(name==='graphs') { renderPlotEditor(); drawChart(); }
  }

  // Import
  $('#chooseFiles').onclick=()=>$('#fileInput').click();
  $('#fileInput').onchange=e=>importFiles([...e.target.files]);
  const dz=$('#dropzone');
  ['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('drag')}));
  ['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('drag')}));
  dz.addEventListener('drop',e=>importFiles([...e.dataTransfer.files]));

  async function importFiles(files){
    for(const file of files){
      try{ await importOne(file); }catch(err){ alert(`Could not import ${file.name}: ${err.message}`); }
    }
    refreshAll();
  }
  async function importOne(file){
    const ext=file.name.split('.').pop().toLowerCase(); let rows=[]; let sheetName='CSV';
    if(ext==='csv'){
      rows=parseCSV(await file.text());
    } else {
      if(!window.XLSX) throw new Error('The Excel reader library did not load. Connect to the internet once or export the file as CSV.');
      const ab=await file.arrayBuffer(); const wb=XLSX.read(ab,{type:'array'});
      let best=null;
      for(const sn of wb.SheetNames){
        const arr=XLSX.utils.sheet_to_json(wb.Sheets[sn],{defval:null,raw:true});
        if(!arr.length) continue;
        const keys=Object.keys(arr[0]).map(k=>k.toLowerCase());
        const score=(keys.some(k=>/sample/.test(k))?1:0)+(keys.some(k=>/target|gene/.test(k))?1:0)+(keys.some(k=>/^cq$|^ct$|c[q|t]/.test(k))?1:0);
        if(!best||score>best.score) best={sn,arr,score};
      }
      if(!best || best.score<2) throw new Error('No sheet with recognizable Sample / Target / Cq columns was found.');
      rows=best.arr; sheetName=best.sn;
    }
    if(!rows.length) throw new Error('No data rows found.');
    const mapping=detectColumns(rows[0]);
    if(!mapping.sample||!mapping.target||!mapping.cq) throw new Error('Could not identify Sample, Target, and Cq/Ct columns.');
    const id='t'+(Date.now()+Math.random()).toString(36).replace('.','');
    const trial={id,name:`Trial ${state.trials.length+1}`,fileName:file.name,sheetName,rowCount:0};
    const norm=[];
    rows.forEach((r,i)=>{
      const cq=num(r[mapping.cq]); const sample=String(r[mapping.sample]??'').trim(); const target=String(r[mapping.target]??'').trim();
      if(!sample||!target||cq===null) return;
      const omit=mapping.omit ? truthy(r[mapping.omit]) : false;
      norm.push({id:`${id}_${i}`,trialId:id,trialName:trial.name,fileName:file.name,well:mapping.well?String(r[mapping.well]??''):String(i+1),sample,target,cq,omit,originalCq:cq});
      if(!(sample in state.sampleMap)) state.sampleMap[sample]=sample;
    });
    if(!norm.length) throw new Error('No valid numeric Ct/Cq rows found. Blank and undetermined values are ignored.'); trial.rowCount=norm.length; state.trials.push(trial); state.raw.push(...norm);
    syncTrialNames(); initializeSelections(); invalidate();
  }
  function truthy(v){ return v===true || v===1 || ['true','yes','y','omit','excluded'].includes(String(v??'').toLowerCase().trim()); }
  function detectColumns(row){
    const keys=Object.keys(row); const find=(regs)=>keys.find(k=>regs.some(r=>r.test(k.trim().toLowerCase())));
    return {sample:find([/^sample$/,/^sample name$/, /sample/]), target:find([/^target$/,/^target name$/, /^gene$/, /target/]), cq:find([/^cq$/,/^ct$/,/^c[q|t]$/, /quantification cycle/,/cycle threshold/]), omit:find([/^omit$/,/exclude/,/excluded/]), well:find([/^well$/,/well position/])};
  }
  function parseCSV(text){if(window.XLSX){const wb=XLSX.read(text,{type:'string',raw:true});return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});}
    const lines=text.replace(/^\uFEFF/,'').split(/\r?\n/).filter(x=>x.trim()!==''); if(!lines.length)return[];
    const split=line=>{let out=[],cur='',q=false;for(let i=0;i<line.length;i++){const c=line[i];if(c==='"'){if(q&&line[i+1]==='"'){cur+='"';i++;}else q=!q;}else if(c===','&&!q){out.push(cur);cur='';}else cur+=c;}out.push(cur);return out;};
    const h=split(lines[0]); return lines.slice(1).map(l=>{const a=split(l),o={};h.forEach((k,i)=>o[k]=a[i]??'');return o;});
  }

  function initializeSelections(){
    const genes=unique(state.raw.map(r=>r.target)).sort((a,b)=>a.localeCompare(b));
    if(!state.refGene){ const preferred=genes.find(g=>/^(tbp|actb|gapdh|b2m|rplp0|hprt1)$/i.test(g)); state.refGene=preferred||genes[0]||''; }
    if(!state.selectedGenes.size) genes.filter(g=>g!==state.refGene).forEach(g=>state.selectedGenes.add(g));
    const conds=unique(Object.values(state.sampleMap)); if(!state.calibrator) state.calibrator=conds[0]||'';
  }
  function syncTrialNames(){
    const lookup=Object.fromEntries(state.trials.map(t=>[t.id,t.name])); state.raw.forEach(r=>r.trialName=lookup[r.trialId]||r.trialName);
  }
  function refreshAll(){ renderTrials(); renderRaw(); renderSetup(); renderResults(); renderGraphControls(); }

  function renderTrials(){
    const el=$('#trialList'); if(!state.trials.length){el.innerHTML='<div class="empty">No trials imported yet.</div>';return;}
    el.innerHTML=state.trials.map(t=>`<div class="trial"><div class="meta"><b>${esc(t.fileName)}</b><span>${esc(t.sheetName)} · ${t.rowCount} usable rows</span></div><input data-trial-name="${t.id}" value="${esc(t.name)}"><button class="btn small danger" data-remove-trial="${t.id}">Remove</button></div>`).join('');
    $$('[data-trial-name]').forEach(inp=>inp.addEventListener('change',e=>{const t=state.trials.find(x=>x.id===e.target.dataset.trialName);if(t){t.name=e.target.value.trim()||t.name;syncTrialNames();renderRaw();invalidate();}}));
    $$('[data-remove-trial]').forEach(b=>b.onclick=()=>{const id=b.dataset.removeTrial;state.trials=state.trials.filter(t=>t.id!==id);state.raw=state.raw.filter(r=>r.trialId!==id);syncMappingsAfterDelete();invalidate();refreshAll();});
  }
  function syncMappingsAfterDelete(){ const samples=new Set(state.raw.map(r=>r.sample)); Object.keys(state.sampleMap).forEach(s=>{if(!samples.has(s))delete state.sampleMap[s]}); initializeSelections(); }

  function renderRaw(){
    const wrap=$('#rawTableWrap'); const q=$('#rawSearch').value.trim().toLowerCase();
    let rows=[...state.raw]; if(q)rows=rows.filter(r=>[r.trialName,r.sample,r.target,r.well].some(v=>String(v).toLowerCase().includes(q))); const sort=$('#rawSort').value; if(sort==='gene')rows.sort((a,b)=>a.target.localeCompare(b.target)||a.sample.localeCompare(b.sample)||a.trialName.localeCompare(b.trialName)); else if(sort==='sample')rows.sort((a,b)=>a.sample.localeCompare(b.sample)||a.target.localeCompare(b.target)||a.trialName.localeCompare(b.trialName));
    if(!rows.length){wrap.innerHTML='<div class="empty"><div class="big">▦</div>No matching raw rows.</div>';return;}
    const max=800, view=rows.slice(0,max);
    wrap.innerHTML=`<table><thead><tr><th>Trial</th><th>Well</th><th>Sample</th><th>Target</th><th>Cq / Ct</th><th>Use</th></tr></thead><tbody>${view.map(r=>`<tr><td>${esc(r.trialName)}</td><td>${esc(r.well)}</td><td><input class="editable" data-raw="sample" data-id="${r.id}" value="${esc(r.sample)}"></td><td><input class="editable" data-raw="target" data-id="${r.id}" value="${esc(r.target)}"></td><td><input class="editable" type="number" step="0.001" data-raw="cq" data-id="${r.id}" value="${r.cq}"></td><td><label class="switch"><input type="checkbox" data-raw="omit" data-id="${r.id}" ${r.omit?'':'checked'}>${r.omit?'<span class="status excl">Excluded</span>':'<span class="status ok">Included</span>'}</label></td></tr>`).join('')}</tbody></table>${rows.length>max?`<div class="footerNote" style="padding:10px">Showing first ${max.toLocaleString()} of ${rows.length.toLocaleString()} rows. Use search to narrow the editor.</div>`:''}`;
    $$('[data-raw]').forEach(el=>el.addEventListener('change',e=>{const r=state.raw.find(x=>x.id===e.target.dataset.id); if(!r)return; const f=e.target.dataset.raw;
      if(f==='cq'){const v=num(e.target.value); if(v!==null)r.cq=v;else {e.target.value=r.cq;toast('Enter a numeric Ct value, or exclude the well.');return;}} else if(f==='omit'){r.omit=!e.target.checked;} else {const old=r[f]; r[f]=e.target.value.trim(); if(!r[f]){r[f]=old;e.target.value=old;return;} if(f==='sample' && old!==r.sample){ if(!(r.sample in state.sampleMap))state.sampleMap[r.sample]=r.sample; }} invalidate(); initializeSelections(); renderSetup();
    }));
  }
  $('#rawSearch').addEventListener('input',renderRaw); $('#rawSort').addEventListener('change',renderRaw); $('#clearExcluded').onclick=()=>{state.raw.forEach(r=>r.omit=false);renderRaw();invalidate()};

  function renderSetup(){
    const genes=unique(state.raw.map(r=>r.target)).sort((a,b)=>a.localeCompare(b));
    const geneGrid=$('#geneGrid'); geneGrid.innerHTML=genes.length?genes.map(g=>`<label class="geneChoice"><input type="checkbox" data-gene="${esc(g)}" ${state.selectedGenes.has(g)?'checked':''}>${esc(g)}${g===state.refGene?' <span class="badge">ref</span>':''}</label>`).join(''):'<div class="empty">Import data first.</div>';
    $$('[data-gene]').forEach(c=>c.onchange=e=>{if(e.target.checked)state.selectedGenes.add(e.target.dataset.gene);else state.selectedGenes.delete(e.target.dataset.gene);invalidate()});
    $('#refGene').innerHTML=genes.length?genes.map(g=>`<option ${g===state.refGene?'selected':''}>${esc(g)}</option>`).join(''):'<option>Import data first</option>';
    const samples=unique(state.raw.map(r=>r.sample)).sort((a,b)=>a.localeCompare(b));
    samples.forEach(s=>{if(!(s in state.sampleMap))state.sampleMap[s]=s;});
    renderSampleMap(samples); renderCalibrator();
  }
  function renderSampleMap(samples){ const wrap=$('#sampleMapWrap'); if(!samples.length){wrap.innerHTML='<div class="empty">Import data first.</div>';return;} wrap.innerHTML=`<table><thead><tr><th>Raw sample name</th><th>Analysis condition</th><th>Rows</th></tr></thead><tbody>${samples.map(s=>`<tr><td>${esc(s)}</td><td><input class="editable" data-map="${esc(s)}" value="${esc(state.sampleMap[s]||s)}"></td><td>${state.raw.filter(r=>r.sample===s&&!r.omit).length}</td></tr>`).join('')}</tbody></table>`; $$('[data-map]').forEach(i=>i.onchange=e=>{state.sampleMap[e.target.dataset.map]=e.target.value.trim()||e.target.dataset.map;invalidate();renderCalibrator();}); }
  function renderCalibrator(){ const conds=unique(Object.values(state.sampleMap)); if(!conds.includes(state.calibrator))state.calibrator=conds[0]||''; $('#calibrator').innerHTML=conds.length?conds.map(c=>`<option ${c===state.calibrator?'selected':''}>${esc(c)}</option>`).join(''):'<option>Import data first</option>'; }
  $('#refGene').onchange=e=>{state.refGene=e.target.value;invalidate();renderSetup();}; $('#calibrator').onchange=e=>{state.calibrator=e.target.value;invalidate()};
  $('#efficiency').onchange=e=>{$('#customBaseField').classList.toggle('hidden',e.target.value!=='custom');state.base=e.target.value==='custom'?Number($('#customBase').value)||2:2;invalidate()}; $('#customBase').onchange=e=>{state.base=Math.max(1.01,Number(e.target.value)||2);invalidate()};
  $('#selectAllGenes').onclick=()=>{unique(state.raw.map(r=>r.target)).filter(g=>g!==state.refGene).forEach(g=>state.selectedGenes.add(g));invalidate();renderSetup()}; $('#clearGenes').onclick=()=>{state.selectedGenes.clear();invalidate();renderSetup()};

  $('#runAnalysis').onclick=()=>{ if(runAnalysis()){renderResults();renderGraphControls();showPage('results');} };
  function runAnalysis(){
    if(!state.raw.length){alert('Import qPCR data first.');return false;} if(!state.refGene){alert('Choose a reference gene.');return false;} if(!state.calibrator){alert('Choose a calibrator condition.');return false;} if(!state.selectedGenes.size){alert('Select at least one target gene.');return false;}
    const usable=state.raw.filter(r=>!r.omit && Number.isFinite(r.cq)); const res=[];
    for(const trial of state.trials){
      const trialRows=usable.filter(r=>r.trialId===trial.id); const conditions=unique(trialRows.map(r=>state.sampleMap[r.sample]||r.sample));
      const refByCond={}; conditions.forEach(cond=>{const rr=trialRows.filter(r=>(state.sampleMap[r.sample]||r.sample)===cond && eqGene(r.target,state.refGene));refByCond[cond]=mean(rr.map(r=>r.cq));});
      for(const gene of [...state.selectedGenes]){
        const calRows=trialRows.filter(r=>(state.sampleMap[r.sample]||r.sample)===state.calibrator && eqGene(r.target,gene));
        const calTarget=mean(calRows.map(r=>r.cq)), calRef=refByCond[state.calibrator]; const calDct=Number.isFinite(calTarget)&&Number.isFinite(calRef)?calTarget-calRef:null;
        for(const cond of conditions){
          const targetRows=trialRows.filter(r=>(state.sampleMap[r.sample]||r.sample)===cond && eqGene(r.target,gene)); if(!targetRows.length)continue;
          const targetMean=mean(targetRows.map(r=>r.cq)), refMean=refByCond[cond]; const dct=Number.isFinite(targetMean)&&Number.isFinite(refMean)?targetMean-refMean:null; const ddct=Number.isFinite(dct)&&Number.isFinite(calDct)?dct-calDct:null; const expression=Number.isFinite(ddct)?Math.pow(state.base,-ddct):null; const log2fc=Number.isFinite(expression)&&expression>0?Math.log2(expression):null;
          res.push({trialId:trial.id,trial:trial.name,gene,condition:cond,targetMean,referenceMean:refMean,dct,calibratorDct:calDct,ddct,expression,log2fc,nTarget:targetRows.length,nReference:trialRows.filter(r=>(state.sampleMap[r.sample]||r.sample)===cond && eqGene(r.target,state.refGene)).length});
        }
      }
    }
    state.results=res; state.analysisRun=true; $('#staleNotice').classList.add('hidden'); toast('Analysis complete · '+res.length+' trial-level values'); return true;
  }
  function eqGene(a,b){return String(a).trim().toLowerCase()===String(b).trim().toLowerCase();}

  function renderResults(){
    const trials=state.trials.length, genes=unique(state.results.map(r=>r.gene)), conds=unique(state.results.map(r=>r.condition));
    $('#kpis').innerHTML=`<div class="kpi"><div class="v">${trials}</div><div class="l">Trials / biological replicates</div></div><div class="kpi"><div class="v">${genes.length}</div><div class="l">Genes analyzed</div></div><div class="kpi"><div class="v">${conds.length}</div><div class="l">Conditions</div></div><div class="kpi"><div class="v">${state.results.length}</div><div class="l">Trial-level values</div></div>`;
    $('#resultGeneFilter').innerHTML='<option value="all">All genes</option>'+genes.map(g=>`<option>${esc(g)}</option>`).join('');
    renderResultsTable();
  }
  function renderResultsTable(){ const wrap=$('#resultsWrap'); if(!state.results.length){wrap.innerHTML='<div class="empty"><div class="big">∑</div>Run the analysis to populate calculations.</div>';return;} const f=$('#resultGeneFilter').value; const rows=f==='all'?state.results:state.results.filter(r=>r.gene===f); wrap.innerHTML=`<table><thead><tr><th>Trial</th><th>Gene</th><th>Condition</th><th>Mean target Ct</th><th>Mean ref Ct</th><th>ΔCt</th><th>Calibrator ΔCt</th><th>ΔΔCt</th><th>2^-ΔΔCt</th><th>log₂FC</th><th>Target reps</th><th>Ref reps</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${esc(r.trial)}</td><td><b>${esc(r.gene)}</b></td><td>${esc(r.condition)}</td><td>${fmt(r.targetMean)}</td><td>${fmt(r.referenceMean)}</td><td>${fmt(r.dct)}</td><td>${fmt(r.calibratorDct)}</td><td>${fmt(r.ddct)}</td><td>${fmt(r.expression)}</td><td>${fmt(r.log2fc)}</td><td>${r.nTarget}</td><td>${r.nReference}</td></tr>`).join('')}</tbody></table>`; }
  $('#resultGeneFilter').onchange=renderResultsTable;

  function graphGenes(){
    const all=unique(state.results.map(r=>r.gene));
    const checked=$$('[data-graph-gene]:checked').map(x=>x.value).filter(g=>all.includes(g));
    if($('#graphMode').value==='single') return checked.slice(0,1);
    return checked;
  }
  function renderGraphControls(){
    const genes=unique(state.results.map(r=>r.gene));
    const grid=$('#graphGeneGrid');
    const oldSelected=new Set(state.pendingGraphGenes||$$('[data-graph-gene]:checked').map(x=>x.value));if(genes.length)delete state.pendingGraphGenes;
    if(!oldSelected.size && genes.length) oldSelected.add(genes[0]);
    grid.innerHTML=genes.length?genes.map((g,i)=>`<label class="graphGeneChoice"><input type="checkbox" data-graph-gene value="${esc(g)}" ${oldSelected.has(g)||(!oldSelected.size&&i===0)?'checked':''}>${esc(g)}</label>`).join(''):'<div class="empty">Run analysis first.</div>';
    const mode=$('#graphMode').value;
    $('#graphGeneLabel').textContent=mode==='single'?'Gene to display':'Genes to display';
    const boxes=$$('[data-graph-gene]');
    if(mode==='single'){
      let first=boxes.find(x=>x.checked)||boxes[0];
      boxes.forEach(x=>x.checked=(x===first));
    }
    boxes.forEach(box=>box.addEventListener('change',e=>{
      if($('#graphMode').value==='single'){
        boxes.forEach(x=>x.checked=(x===e.target));
      } else if(!boxes.some(x=>x.checked)) e.target.checked=true;
      syncStatGene(); renderReplicateSummary(); renderPlotEditor(); drawChart();
    }));
    syncStatGene();
    const conds=unique(state.results.map(r=>r.condition));
    $('#statPair').innerHTML='<option value="auto">First two available conditions</option>'+pairOptions(conds).map(p=>`<option value="${esc(p.join('|||'))}">${esc(p[0])} vs ${esc(p[1])}</option>`).join('');
    updateMetricNote(); renderReplicateSummary(); renderPlotEditor(); drawChart();
  }
  function syncStatGene(){
    const genes=unique(state.results.map(r=>r.gene)); const cur=$('#statGene').value; const displayed=graphGenes();
    $('#statGene').innerHTML=genes.length?genes.map(g=>`<option ${g===cur?'selected':''}>${esc(g)}</option>`).join(''):'<option value="">Run analysis first</option>';
    if(displayed.length && !displayed.includes($('#statGene').value)) $('#statGene').value=displayed[0];
  }
  function renderReplicateSummary(){
    const genes=graphGenes(), metric=$('#graphMetric').value;
    const rows=state.results.filter(r=>genes.includes(r.gene) && graphRowVisible(r,metric) && Number.isFinite(valueFor(r,metric,true)));
    if(!rows.length){$('#replicateSummary').innerHTML='';return;}
    const conds=unique(rows.map(r=>r.condition));
    const counts=[];
    genes.forEach(g=>conds.forEach(c=>{
      const n=new Set(rows.filter(r=>r.gene===g&&r.condition===c).map(r=>r.trialId)).size;
      if(n) counts.push(n);
    }));
    const trialCount=new Set(rows.map(r=>r.trialId)).size;
    const minN=Math.min(...counts), maxN=Math.max(...counts);
    const detail=minN===maxN
      ? `<span class="repPill"><b>Biological replicates:</b> n=${minN}</span>`
      : `<span class="repPill"><b>Biological replicates:</b> n=${minN}–${maxN} across groups</span>`;
    const imported=trialCount!==maxN?`<span class="repPill">${trialCount} imported trial${trialCount===1?'':'s'} represented</span>`:'';
    $('#replicateSummary').innerHTML=detail+imported;
  }
  function pairOptions(arr){const out=[];for(let i=0;i<arr.length;i++)for(let j=i+1;j<arr.length;j++)out.push([arr[i],arr[j]]);return out;}
  $('#graphMode').addEventListener('change',()=>{ if($('#graphMode').value==='multi' && $('#graphType').value==='bar') $('#graphType').value='grouped'; if($('#graphMode').value==='single' && $('#graphType').value==='grouped') $('#graphType').value='bar'; renderGraphControls();});
  ['graphMetric','graphType','errorType','yScale'].forEach(id=>$('#'+id).addEventListener('change',()=>{updateMetricNote();renderReplicateSummary();renderPlotEditor();drawChart()}));
  $('#showThreshold').addEventListener('change',()=>{updateThresholdControls();drawChart();});
  $('#thresholdValue').addEventListener('input',drawChart);
  $('#thresholdLabel').addEventListener('input',drawChart);
  $('#statGene').addEventListener('change',()=>{$('#statsResult').textContent='Ready to run a statistical test for '+$('#statGene').value+'.';});
  $('#graphTitle').addEventListener('input',drawChart);
  updateThresholdControls();

  function updateThresholdControls(){
    const on=$('#showThreshold').checked;
    $('#thresholdValue').disabled=!on;
    $('#thresholdLabel').disabled=!on;
  }
  function thresholdConfig(){
    const v=num($('#thresholdValue').value);
    return {enabled:$('#showThreshold').checked && v!==null, value:v, label:$('#thresholdLabel').value.trim()};
  }
  function drawThreshold(ctx,x1,x2,yFn,threshold,yscale){
    if(!threshold.enabled || !Number.isFinite(threshold.value)) return;
    if(yscale==='log' && threshold.value<=0) return;
    const yy=yFn(threshold.value);
    if(!Number.isFinite(yy)) return;
    ctx.save();
    ctx.setLineDash([2,5]);
    ctx.strokeStyle='#344054';
    ctx.lineWidth=1.6;
    ctx.beginPath();ctx.moveTo(x1,yy);ctx.lineTo(x2,yy);ctx.stroke();
    ctx.setLineDash([]);
    if(threshold.label){
      ctx.font='10px system-ui';ctx.textAlign='right';ctx.textBaseline='bottom';
      const label=`${threshold.label} (y=${fmt(threshold.value,2)})`;
      const tw=ctx.measureText(label).width;
      ctx.fillStyle='rgba(255,255,255,.88)';ctx.fillRect(x2-tw-7,yy-15,tw+7,14);
      ctx.fillStyle='#344054';ctx.fillText(label,x2-2,yy-2);
    }
    ctx.restore();
  }

  function valueFor(r,metric,useOverride=true){ const key=overrideKey(r,metric); if(useOverride && state.overrides[key]!==undefined && state.overrides[key]!==null && Number.isFinite(Number(state.overrides[key]))) return Number(state.overrides[key]); return metric==='expression'?r.expression:metric==='log2fc'?r.log2fc:r.dct; }
  function graphRowVisible(r,metric){ return metric!=='log2fc' || r.condition!==state.calibrator; }
  function updateMetricNote(){
    const note=$('#metricNote'), metric=$('#graphMetric').value;
    if(metric==='log2fc'){
      note.classList.remove('hidden');
      note.innerHTML=`<b>${esc(state.calibrator||'Calibrator')} is the normalization baseline.</b> Its log₂FC is defined as 0 and is intentionally hidden from log₂FC graphs. Only condition(s) compared with the calibrator are plotted.`;
    } else { note.classList.add('hidden'); note.innerHTML=''; }
  }
  function overrideKey(r,metric){return `${r.trialId}::${r.gene}::${r.condition}::${metric}`;}
  function renderPlotEditor(){
    const wrap=$('#plotEditorWrap'), genes=graphGenes(), metric=$('#graphMetric').value; const rows=state.results.filter(r=>genes.includes(r.gene) && graphRowVisible(r,metric));
    if(!rows.length){wrap.innerHTML='<div class="empty">Run the analysis first.</div>';return;}
    wrap.innerHTML=`<table><thead><tr><th>Gene</th><th>Trial</th><th>Condition</th><th>Calculated</th><th>Manual override</th><th>Plotted value</th></tr></thead><tbody>${rows.map(r=>{const k=overrideKey(r,metric), base=valueFor(r,metric,false), ov=state.overrides[k];return `<tr><td><b>${esc(r.gene)}</b></td><td>${esc(r.trial)}</td><td>${esc(r.condition)}</td><td>${fmt(base,4)}</td><td><input type="number" step="any" data-override="${esc(k)}" value="${ov??''}" placeholder="leave blank"></td><td>${fmt(valueFor(r,metric,true),4)}</td></tr>`}).join('')}</tbody></table>`;
    $$('[data-override]').forEach(i=>i.oninput=e=>{const k=e.target.dataset.override;if(e.target.value==='')delete state.overrides[k];else state.overrides[k]=Number(e.target.value);drawChart();renderReplicateSummary();const tr=e.target.closest('tr');if(tr)tr.lastElementChild.textContent=e.target.value===''?tr.children[3].textContent:fmt(Number(e.target.value),4);});
  }
  $('#clearOverrides').onclick=()=>{const genes=new Set(graphGenes());Object.keys(state.overrides).forEach(k=>{if(genes.has(k.split('::')[1]))delete state.overrides[k]});renderPlotEditor();renderReplicateSummary();drawChart();};

  function drawChart(){
    const canvas=$('#chart'), wrap=canvas.parentElement;
    const genes=graphGenes(), metric=$('#graphMetric').value, type=$('#graphType').value, yscale=$('#yScale').value;
    let rows=state.results.filter(r=>genes.includes(r.gene) && graphRowVisible(r,metric) && Number.isFinite(valueFor(r,metric,true)) && (yscale!=='log'||valueFor(r,metric,true)>0));
    const conds=unique(rows.map(r=>r.condition));

    const parentWidth=Math.max(320,wrap.clientWidth||wrap.getBoundingClientRect().width||800);
    const desiredWidth=Math.min(820,parentWidth);
    const cssHeight=Math.max(360,wrap.clientHeight||460);
    const dpr=Math.max(2,window.devicePixelRatio||1);
    canvas.style.width=Math.ceil(desiredWidth)+'px';
    canvas.style.height=cssHeight+'px';
    canvas.width=Math.ceil(desiredWidth*dpr);
    canvas.height=Math.ceil(cssHeight*dpr);
    const ctx=canvas.getContext('2d'); ctx.scale(dpr,dpr);
    const W=desiredWidth,H=cssHeight;
    ctx.clearRect(0,0,W,H); ctx.fillStyle='#fff';ctx.fillRect(0,0,W,H);
    if(!rows.length){ctx.fillStyle='#7a8597';ctx.font='14px system-ui';ctx.textAlign='center';ctx.fillText(metric==='log2fc'?'No comparison conditions are available beyond the calibrator.':'Run analysis or choose genes with complete data.',W/2,H/2);return;}
    if(type==='grouped'){ drawGroupedGeneChart(ctx,W,H,rows,genes,conds,metric,yscale); return; }
    const categories=[]; genes.forEach(g=>conds.forEach(c=>{if(rows.some(r=>r.gene===g&&r.condition===c))categories.push({gene:g,condition:c,key:g+'|||'+c})}));
    const groups=categories.map(cat=>rows.filter(r=>r.gene===cat.gene&&r.condition===cat.condition).map(r=>({r,v:valueFor(r,metric,true)})).filter(x=>Number.isFinite(x.v)));
    const multi=genes.length>1; const axisLabels=categories.map(cat=>multi?cat.gene+' · '+cat.condition:cat.condition); const axis=compactAxis(ctx,axisLabels,W,categories.length); const m={l:axis.left,r:24,t:58,b:axis.bottom}; const plotW=W-m.l-m.r,plotH=H-m.t-m.b;
    let vals=groups.flat().map(x=>x.v); let transform=x=>x, inv=x=>x;if(yscale==='log'){vals=vals.filter(v=>v>0);transform=x=>Math.log10(x);inv=x=>10**x;if(!vals.length){ctx.fillStyle='#b42318';ctx.textAlign='center';ctx.fillText('Log scale requires positive values.',W/2,H/2);return;}}
    const threshold=thresholdConfig();
    const ends=groups.flatMap(g=>{const v=g.map(p=>p.v), a=mean(v), e=$('#errorType').value==='sem'?sem(v):$('#errorType').value==='sd'?sd(v):0;return [a-e,a+e]}).filter(v=>Number.isFinite(v)&&(yscale!=='log'||v>0)); const rangeVals=[...vals,...ends];
    if(threshold.enabled && Number.isFinite(threshold.value) && (yscale!=='log'||threshold.value>0)) rangeVals.push(threshold.value);
    const tv=rangeVals.filter(v=>yscale!=='log'||v>0).map(transform);let ymin=Math.min(...tv),ymax=Math.max(...tv);if(ymin===ymax){ymin-=1;ymax+=1;}const pad=(ymax-ymin)*.16;ymin-=pad;ymax+=pad;if(yscale==='linear'&&metric==='expression')ymin=Math.min(0,ymin);const y=v=>m.t+(ymax-transform(yscale==='log'?Math.max(v,10**ymin):v))/(ymax-ymin)*plotH;const xCenter=i=>m.l+(i+.5)*plotW/categories.length;
    ctx.strokeStyle=$('#showGrid').checked?'#e6eaf0':'transparent';ctx.lineWidth=1;ctx.font='11px system-ui';ctx.fillStyle='#657185';ctx.textAlign='right';ctx.textBaseline='middle';for(let i=0;i<=5;i++){const yy=m.t+i*plotH/5;ctx.beginPath();ctx.moveTo(m.l,yy);ctx.lineTo(W-m.r,yy);ctx.stroke();const t=ymax-(ymax-ymin)*i/5;ctx.fillText(yscale==='log'?fmt(inv(t),2):fmt(t,2),m.l-9,yy)}
    if(metric==='log2fc'&&yscale==='linear'&&ymin<=0&&ymax>=0){const zy=y(0);ctx.save();ctx.setLineDash([2,5]);ctx.strokeStyle='#98a2b3';ctx.lineWidth=1.3;ctx.beginPath();ctx.moveTo(m.l,zy);ctx.lineTo(W-m.r,zy);ctx.stroke();ctx.restore();}
    drawThreshold(ctx,m.l,W-m.r,y,threshold,yscale);
    ctx.strokeStyle='#8290a5';ctx.beginPath();ctx.moveTo(m.l,m.t);ctx.lineTo(m.l,H-m.b);ctx.lineTo(W-m.r,H-m.b);ctx.stroke();
    categories.forEach((cat,i)=>drawAngledLabel(ctx,axisLabels[i],xCenter(i),H-m.b+12,axis));
    const palette=colors;
    const trialIds=unique(rows.map(r=>r.trialId)); const trialIndex=new Map(trialIds.map((id,i)=>[id,i]));
    groups.forEach((g,i)=>{
      if(!g.length)return;const xc=xCenter(i),bw=Math.min(70,plotW/categories.length*.58),values=g.map(x=>x.v),avg=mean(values),err=$('#errorType').value==='sem'?sem(values):$('#errorType').value==='sd'?sd(values):0;const geneColor=palette[genes.indexOf(categories[i].gene)%palette.length];
      if(type==='bar'){ctx.fillStyle=hexAlpha(geneColor,Number($('#barOpacity').value));ctx.strokeStyle=geneColor;ctx.lineWidth=2;const yy=y(avg),base=y(yscale==='log'?Math.min(...vals):Math.max(0,ymin));ctx.fillRect(xc-bw/2,Math.min(yy,base),bw,Math.abs(base-yy));ctx.strokeRect(xc-bw/2,Math.min(yy,base),bw,Math.abs(base-yy));}
      if(type==='box')drawBox(ctx,values,xc,bw,y,geneColor);
      if(type!=='paired'&&type!=='box')g.forEach(p=>{const ti=trialIndex.get(p.r.trialId)||0;const denom=Math.max(1,trialIds.length-1);const offset=trialIds.length===1?0:((ti/denom)-.5)*Math.min(bw*.7,44);dot(ctx,xc+offset,y(p.v),palette[ti%palette.length]);});
      if(err&&type!=='box'&&type!=='points'){ctx.strokeStyle='#45536a';ctx.lineWidth=1.4;const top=y(avg+err),bot=y(avg-err);ctx.beginPath();ctx.moveTo(xc,top);ctx.lineTo(xc,bot);ctx.moveTo(xc-8,top);ctx.lineTo(xc+8,top);ctx.moveTo(xc-8,bot);ctx.lineTo(xc+8,bot);ctx.stroke();}
      if(type==='points'){ctx.strokeStyle=geneColor;ctx.lineWidth=2;ctx.beginPath();ctx.moveTo(xc-bw/3,y(avg));ctx.lineTo(xc+bw/3,y(avg));ctx.stroke();}
    });
    if(type==='paired'){
      genes.forEach((gene,gi)=>{const geneCats=categories.map((c,i)=>({c,i})).filter(x=>x.c.gene===gene);trialIds.forEach((tid,k)=>{const pts=[];geneCats.forEach(({c,i})=>{const r=rows.find(z=>z.trialId===tid&&z.gene===gene&&z.condition===c.condition);if(r){const v=valueFor(r,metric,true);if(Number.isFinite(v))pts.push({x:xCenter(i),y:y(v)})}});if(pts.length>1){ctx.strokeStyle=hexAlpha(palette[k%palette.length],.55);ctx.lineWidth=1.5;ctx.beginPath();pts.forEach((p,j)=>j?ctx.lineTo(p.x,p.y):ctx.moveTo(p.x,p.y));ctx.stroke();}pts.forEach(p=>dot(ctx,p.x,p.y,palette[k%palette.length]));});});
    }
    const title=$('#graphTitle').value||(genes.length===1?genes[0]:'qPCR gene comparison');ctx.fillStyle='#1d2939';ctx.font='700 16px system-ui';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(title,W/2,24);
    ctx.save();ctx.translate(18,m.t+plotH/2);ctx.rotate(-Math.PI/2);ctx.font='12px system-ui';ctx.fillStyle='#536074';ctx.fillText(metricLabel(metric),0,0);ctx.restore();
    // trial legend: colors identify biological trials, making overlapping values visibly distinct
    if(trialIds.length>1){ctx.font='10px system-ui';ctx.textAlign='right';ctx.textBaseline='middle';let yy=12;trialIds.slice(0,6).forEach((tid,i)=>{const tr=state.trials.find(t=>t.id===tid);const label=tr?tr.name:'Trial '+(i+1);ctx.fillStyle=palette[i%palette.length];ctx.beginPath();ctx.arc(W-115,yy,3,0,Math.PI*2);ctx.fill();ctx.fillStyle='#657185';ctx.textAlign='left';ctx.fillText(label,W-107,yy);yy+=13;});}
    const legend=$('#chartLegend'); if(legend){legend.innerHTML=trialIds.map((tid,i)=>{const tr=state.trials.find(t=>t.id===tid);return `<span><i style="display:inline-block;width:8px;height:8px;border-radius:50%;border:2px solid ${palette[i%palette.length]};background:#fff;margin-right:5px"></i>${esc(tr?tr.name:'Trial '+(i+1))}</span>`}).join('')+'<span>Each point = one imported trial / biological replicate</span>'+(metric==='log2fc'?`<span><b>Baseline:</b> ${esc(state.calibrator)} = 0 (hidden)</span>`:'')+(threshold.enabled?`<span><b>Dotted threshold:</b> y=${esc(fmt(threshold.value,2))}${threshold.label?' · '+esc(threshold.label):''}</span>`:''); }
  }
  function drawGroupedGeneChart(ctx,W,H,rows,genes,conds,metric,yscale){
    const conditionPalette=colors;
    const trialPalette=colors;
    const groups=[];
    genes.forEach(g=>conds.forEach(c=>{
      const pts=rows.filter(r=>r.gene===g&&r.condition===c).map(r=>({r,v:valueFor(r,metric,true)})).filter(x=>Number.isFinite(x.v));
      if(pts.length)groups.push({gene:g,condition:c,pts});
    }));
    let vals=groups.flatMap(g=>g.pts.map(p=>p.v));
    if(!vals.length)return;
    let transform=x=>x, inv=x=>x;
    if(yscale==='log'){
      vals=vals.filter(v=>v>0); transform=x=>Math.log10(x); inv=x=>10**x;
      if(!vals.length){ctx.fillStyle='#b42318';ctx.font='14px system-ui';ctx.textAlign='center';ctx.fillText('Log scale requires positive values.',W/2,H/2);return;}
    }
    const errorMode=$('#errorType').value;
    const statEnds=[];
    groups.forEach(g=>{
      const v=g.pts.map(p=>p.v), avg=mean(v), err=errorMode==='sem'?sem(v):errorMode==='sd'?sd(v):0;
      statEnds.push(avg);
      if(err){statEnds.push(avg+err); if(yscale!=='log'||avg-err>0)statEnds.push(avg-err);}
    });
    const threshold=thresholdConfig();
    const rangeVals=[...vals,...statEnds.filter(Number.isFinite)];
    if(threshold.enabled && Number.isFinite(threshold.value) && (yscale!=='log'||threshold.value>0)) rangeVals.push(threshold.value);
    let tv=rangeVals.filter(v=>yscale!=='log'||v>0).map(transform);
    let ymin=Math.min(...tv), ymax=Math.max(...tv);
    if(ymin===ymax){ymin-=1;ymax+=1;}
    let pad=(ymax-ymin)*.18; ymin-=pad; ymax+=pad;
    if(yscale==='linear'&&metric==='expression')ymin=Math.min(0,ymin);
    const axis=compactAxis(ctx,genes,W,genes.length); const m={l:axis.left,r:26,t:86,b:axis.bottom}, plotW=W-m.l-m.r, plotH=H-m.t-m.b;
    const y=v=>m.t+(ymax-transform(v))/(ymax-ymin)*plotH;
    const zeroBase=yscale==='log'?Math.min(...vals):Math.max(0,ymin);
    ctx.strokeStyle=$('#showGrid').checked?'#e6eaf0':'transparent';ctx.lineWidth=1;ctx.font='11px system-ui';ctx.fillStyle='#657185';ctx.textAlign='right';ctx.textBaseline='middle';
    for(let i=0;i<=5;i++){const yy=m.t+i*plotH/5;ctx.beginPath();ctx.moveTo(m.l,yy);ctx.lineTo(W-m.r,yy);ctx.stroke();const t=ymax-(ymax-ymin)*i/5;ctx.fillText(yscale==='log'?fmt(inv(t),2):fmt(t,2),m.l-9,yy);}
    if(metric==='log2fc'&&yscale==='linear'&&ymin<=0&&ymax>=0){const zy=y(0);ctx.save();ctx.setLineDash([2,5]);ctx.strokeStyle='#98a2b3';ctx.lineWidth=1.3;ctx.beginPath();ctx.moveTo(m.l,zy);ctx.lineTo(W-m.r,zy);ctx.stroke();ctx.restore();}
    drawThreshold(ctx,m.l,W-m.r,y,threshold,yscale);
    ctx.strokeStyle='#8290a5';ctx.beginPath();ctx.moveTo(m.l,m.t);ctx.lineTo(m.l,H-m.b);ctx.lineTo(W-m.r,H-m.b);ctx.stroke();
    const geneW=plotW/Math.max(1,genes.length);
    const activeConds=conds.filter(c=>groups.some(g=>g.condition===c));
    const clusterW=Math.min(geneW*.78,150);
    const gap=Math.min(4,clusterW/(Math.max(1,activeConds.length)*5));
    const barW=Math.min(42,(clusterW-gap*Math.max(0,activeConds.length-1))/Math.max(1,activeConds.length));
    const trialIds=unique(rows.map(r=>r.trialId));
    const trialIndex=new Map(trialIds.map((id,i)=>[id,i]));
    genes.forEach((gene,gi)=>{
      const gx=m.l+(gi+.5)*geneW;
      const actual=activeConds.filter(c=>groups.some(g=>g.gene===gene&&g.condition===c));
      const totalW=actual.length*barW+Math.max(0,actual.length-1)*gap;
      actual.forEach((cond,ci)=>{
        const g=groups.find(z=>z.gene===gene&&z.condition===cond); if(!g)return;
        const xc=gx-totalW/2+barW/2+ci*(barW+gap);
        const values=g.pts.map(p=>p.v), avg=mean(values), err=errorMode==='sem'?sem(values):errorMode==='sd'?sd(values):0;
        const cidx=activeConds.indexOf(cond), c=conditionPalette[cidx%conditionPalette.length];
        const yy=y(avg), base=y(zeroBase);
        ctx.fillStyle=hexAlpha(c,Number($('#barOpacity').value));ctx.strokeStyle=c;ctx.lineWidth=1.5;
        ctx.fillRect(xc-barW/2,Math.min(yy,base),barW,Math.abs(base-yy));ctx.strokeRect(xc-barW/2,Math.min(yy,base),barW,Math.abs(base-yy));
        if(err){
          const hi=avg+err, lo=avg-err;
          if(yscale!=='log'||lo>0){ctx.strokeStyle='#344054';ctx.lineWidth=1.3;const top=y(hi),bot=y(lo);ctx.beginPath();ctx.moveTo(xc,top);ctx.lineTo(xc,bot);ctx.moveTo(xc-5,top);ctx.lineTo(xc+5,top);ctx.moveTo(xc-5,bot);ctx.lineTo(xc+5,bot);ctx.stroke();}
        }
        const spread=Math.min(barW*.58,24);
        g.pts.forEach(p=>{
          const ti=trialIndex.get(p.r.trialId)||0;
          const denom=Math.max(1,trialIds.length-1);
          const offset=trialIds.length===1?0:((ti/denom)-.5)*spread;
          dot(ctx,xc+offset,y(p.v),trialPalette[ti%trialPalette.length]);
        });
      });
      drawAngledLabel(ctx,gene,gx,H-m.b+12,axis);
    });
    const title=$('#graphTitle').value||(genes.length===1?genes[0]:'qPCR gene comparison');ctx.fillStyle='#1d2939';ctx.font='700 16px system-ui';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(title,W/2,24);
    ctx.save();ctx.translate(18,m.t+plotH/2);ctx.rotate(-Math.PI/2);ctx.font='12px system-ui';ctx.fillStyle='#536074';ctx.fillText(metricLabel(metric),0,0);ctx.restore();
    // condition legend across top
    ctx.font='10px system-ui';ctx.textBaseline='middle';let lx=m.l,ly=53;
    activeConds.forEach((cond,i)=>{const c=conditionPalette[i%conditionPalette.length];ctx.fillStyle=hexAlpha(c,Number($('#barOpacity').value));ctx.fillRect(lx,ly-5,10,10);ctx.strokeStyle=c;ctx.strokeRect(lx,ly-5,10,10);ctx.fillStyle='#536074';ctx.textAlign='left';ctx.fillText(cond,lx+14,ly);lx+=18+ctx.measureText(cond).width+14;});
    // HTML legend keeps trial identity explicit
    const legend=$('#chartLegend');
    if(legend){legend.innerHTML=`<span><b>Bars:</b> conditions</span>${trialIds.map((tid,i)=>{const tr=state.trials.find(t=>t.id===tid);return `<span><i style="display:inline-block;width:8px;height:8px;border-radius:50%;border:2px solid ${trialPalette[i%trialPalette.length]};background:#fff;margin-right:5px"></i>${esc(tr?tr.name:'Trial '+(i+1))}</span>`}).join('')}<span>Each point = one imported trial / biological replicate</span>${metric==='log2fc'?`<span><b>Baseline:</b> ${esc(state.calibrator)} = 0 (hidden)</span>`:''}${threshold.enabled?`<span><b>Dotted threshold:</b> y=${esc(fmt(threshold.value,2))}${threshold.label?' · '+esc(threshold.label):''}</span>`:''}`;}
  }

  function hexAlpha(hex,a){const h=hex.replace('#','');const n=parseInt(h,16),r=(n>>16)&255,g=(n>>8)&255,b=n&255;return `rgba(${r},${g},${b},${a})`}
  function dot(ctx,x,y,c){ctx.beginPath();ctx.arc(x,y,Number($('#pointSize').value),0,Math.PI*2);ctx.fillStyle='#fff';ctx.fill();ctx.lineWidth=2.2;ctx.strokeStyle=c;ctx.stroke()}
  function quantile(a,q){const b=[...a].sort((x,y)=>x-y);if(!b.length)return null;const p=(b.length-1)*q,lo=Math.floor(p),hi=Math.ceil(p);return b[lo]+(b[hi]-b[lo])*(p-lo)}
  function drawBox(ctx,vals,x,bw,y,c){const q1=quantile(vals,.25),med=quantile(vals,.5),q3=quantile(vals,.75),lo=Math.min(...vals),hi=Math.max(...vals);ctx.fillStyle=hexAlpha(c,.16);ctx.strokeStyle=c;ctx.lineWidth=1.8;ctx.fillRect(x-bw/2,y(q3),bw,y(q1)-y(q3));ctx.strokeRect(x-bw/2,y(q3),bw,y(q1)-y(q3));ctx.beginPath();ctx.moveTo(x-bw/2,y(med));ctx.lineTo(x+bw/2,y(med));ctx.moveTo(x,y(hi));ctx.lineTo(x,y(q3));ctx.moveTo(x,y(q1));ctx.lineTo(x,y(lo));ctx.stroke();vals.forEach((v,j)=>dot(ctx,x+(j-(vals.length-1)/2)*Math.min(12,bw/(vals.length+2)),y(v),c));}
  function metricLabel(m){return m==='expression'?'Relative expression (2^-ΔΔCt)':m==='log2fc'?'log₂ fold change':'ΔCt';}
  window.addEventListener('resize',()=>{if($('#page-graphs').classList.contains('active'))drawChart()});

  // Statistics — t distribution / F distribution via regularized incomplete beta
  function logGamma(z){const c=[676.5203681218851,-1259.1392167224028,771.32342877765313,-176.6150291621406,12.507343278686905,-0.13857109526572012,9.9843695780195716e-6,1.5056327351493116e-7];if(z<.5)return Math.log(Math.PI)-Math.log(Math.sin(Math.PI*z))-logGamma(1-z);z-=1;let x=.99999999999980993;for(let i=0;i<c.length;i++)x+=c[i]/(z+i+1);const t=z+c.length-.5;return .5*Math.log(2*Math.PI)+(z+.5)*Math.log(t)-t+Math.log(x)}
  function betacf(a,b,x){const MAX=200,EPS=3e-10,FPMIN=1e-30;let qab=a+b,qap=a+1,qam=a-1,c=1,d=1-qab*x/qap;if(Math.abs(d)<FPMIN)d=FPMIN;d=1/d;let h=d;for(let m=1;m<=MAX;m++){let m2=2*m,aa=m*(b-m)*x/((qam+m2)*(a+m2));d=1+aa*d;if(Math.abs(d)<FPMIN)d=FPMIN;c=1+aa/c;if(Math.abs(c)<FPMIN)c=FPMIN;d=1/d;h*=d*c;aa=-(a+m)*(qab+m)*x/((a+m2)*(qap+m2));d=1+aa*d;if(Math.abs(d)<FPMIN)d=FPMIN;c=1+aa/c;if(Math.abs(c)<FPMIN)c=FPMIN;d=1/d;const del=d*c;h*=del;if(Math.abs(del-1)<EPS)break;}return h}
  function betai(a,b,x){if(x<=0)return 0;if(x>=1)return 1;const bt=Math.exp(logGamma(a+b)-logGamma(a)-logGamma(b)+a*Math.log(x)+b*Math.log(1-x));return x<(a+1)/(a+b+2)?bt*betacf(a,b,x)/a:1-bt*betacf(b,a,1-x)/b}
  function tPValue(t,df){if(t===Infinity)return 0;if(!Number.isFinite(t)||!Number.isFinite(df)||df<=0)return null;const x=df/(df+t*t);return Math.min(1,betai(df/2,.5,x));}
  function pairedT(a,b){const pairs=[];for(let i=0;i<Math.min(a.length,b.length);i++)if(Number.isFinite(a[i])&&Number.isFinite(b[i]))pairs.push(a[i]-b[i]);if(pairs.length<2)return null;const m=mean(pairs),s=sd(pairs),t=s===0?(m===0?0:Infinity):m/(s/Math.sqrt(pairs.length)),df=pairs.length-1;return {name:'Paired t-test',t,df,p:tPValue(Math.abs(t),df),n:pairs.length};}
  function welchT(a,b){a=a.filter(Number.isFinite);b=b.filter(Number.isFinite);if(a.length<2||b.length<2)return null;const m1=mean(a),m2=mean(b),v1=sd(a)**2,v2=sd(b)**2,se=Math.sqrt(v1/a.length+v2/b.length);const t=se===0?(m1===m2?0:Infinity):(m1-m2)/se;const df=(v1/a.length+v2/b.length)**2/((v1*v1)/(a.length*a.length*(a.length-1))+(v2*v2)/(b.length*b.length*(b.length-1)));return {name:"Welch's t-test",t,df,p:tPValue(Math.abs(t),df),n:`${a.length}, ${b.length}`};}
  function anova(groups){groups=groups.map(g=>g.filter(Number.isFinite)).filter(g=>g.length);const N=groups.reduce((s,g)=>s+g.length,0),k=groups.length;if(k<2||N<=k)return null;const all=groups.flat(),gm=mean(all);let ssb=0,ssw=0;groups.forEach(g=>{const m=mean(g);ssb+=g.length*(m-gm)**2;ssw+=g.reduce((s,x)=>s+(x-m)**2,0)});const df1=k-1,df2=N-k,F=(ssb/df1)/(ssw/df2);const x=(df1*F)/(df1*F+df2);const p=F===Infinity?0:1-betai(df1/2,df2/2,x);return {name:'One-way ANOVA',F,df1,df2,p,n:N};}
  $('#runStats').onclick=runStats;
  function runStats(){
    const gene=$('#statGene').value,metric=$('#graphMetric').value,use=$('#statsUseOverrides').checked, rows=state.results.filter(r=>r.gene===gene&&Number.isFinite(valueFor(r,metric,use))), conds=unique(rows.map(r=>r.condition));
    if(!gene){$('#statsResult').textContent='Choose a gene for statistical testing.';return;}
    if(conds.length<2){$('#statsResult').textContent='At least two conditions with calculated biological-replicate values are required.';return;}
    const test=$('#statTest').value;let out=null,detail='';
    if(test==='anova'){
      const groups=conds.map(c=>rows.filter(r=>r.condition===c).map(r=>valueFor(r,metric,use)));
      out=anova(groups);if(out)detail=`F(${out.df1}, ${out.df2}) = ${fmt(out.F,3)} · total n = ${out.n}`;
    } else {
      const pair=$('#statPair').value==='auto'?conds.slice(0,2):$('#statPair').value.split('|||');
      if(!pair[0]||!pair[1]||!conds.includes(pair[0])||!conds.includes(pair[1])){$('#statsResult').textContent='Choose two conditions that exist for this gene.';return;}
      const mapTrial=c=>{const m=new Map();rows.filter(r=>r.condition===c).forEach(r=>{const v=valueFor(r,metric,use);if(Number.isFinite(v))m.set(r.trialId,v)});return m};
      if(test==='paired'){
        const A=mapTrial(pair[0]),B=mapTrial(pair[1]),ids=[...A.keys()].filter(id=>B.has(id));
        out=pairedT(ids.map(id=>A.get(id)),ids.map(id=>B.get(id)));detail=out?`t(${out.df.toFixed(0)}) = ${Number.isFinite(out.t)?fmt(out.t,3):'∞'} · paired biological n = ${out.n}`:`Only ${ids.length} matched trial${ids.length===1?'':'s'} contain both conditions; paired t-test requires at least 2.`;
      } else {
        const a=rows.filter(r=>r.condition===pair[0]).map(r=>valueFor(r,metric,use)),b=rows.filter(r=>r.condition===pair[1]).map(r=>valueFor(r,metric,use));
        out=welchT(a,b);detail=out?`t(${fmt(out.df,1)}) = ${Number.isFinite(out.t)?fmt(out.t,3):'∞'} · biological n = ${a.length} vs ${b.length}`:`Available biological n = ${a.length} vs ${b.length}; Welch's t-test requires at least 2 per condition.`;
      }
    }
    if(!out||!Number.isFinite(out.p)){$('#statsResult').innerHTML=`<strong>${esc(gene)}</strong><br>${detail||'Not enough biological replicate values to run this test.'}<div class="footerNote">Check the n labels above the graph. Each independent imported trial should contribute one value per condition.</div>`;return;}
    const sig=out.p<.05;$('#statsResult').innerHTML=`<strong>${esc(gene)} · ${esc(out.name)}</strong><br>${detail}<br>p = <span class="${sig?'pSig':'pNS'}">${out.p<.0001?'&lt; 0.0001':out.p.toFixed(4)}</span>${sig?' · p < 0.05':''}<div class="footerNote">Statistical significance should be interpreted in the context of study design, biological replication, and planned comparisons.</div>`;
  }

  // Export/save
  $('#exportResults').onclick=()=>{if(!state.results.length){alert('Run analysis first.');return;}const h=['Trial','Gene','Condition','Mean Target Ct','Mean Reference Ct','dCt','Calibrator dCt','ddCt','Relative Expression','log2FC','Target Technical Replicates','Reference Technical Replicates'];const data=state.results.map(r=>[r.trial,r.gene,r.condition,r.targetMean,r.referenceMean,r.dct,r.calibratorDct,r.ddct,r.expression,r.log2fc,r.nTarget,r.nReference]);downloadText('qpcr_results.csv',[h,...data].map(row=>row.map(csvCell).join(',')).join('\n'),'text/csv');};
  function csvCell(v){let s=v==null?'':String(v); if(typeof v==='string' && /^[=+@-]/.test(s))s="'"+s;return /[",\n]/.test(s)?`"${s.replace(/"/g,'""')}"`:s}
  $('#saveProject').onclick=()=>{const payload={version:2,figure:getFigure(),trials:state.trials,raw:state.raw,sampleMap:state.sampleMap,selectedGenes:[...state.selectedGenes],refGene:state.refGene,calibrator:state.calibrator,base:state.base,overrides:state.overrides};downloadText('qpcr_project.json',JSON.stringify(payload,null,2),'application/json');};
  $('#loadProjectBtn').onclick=()=>$('#projectInput').click(); $('#projectInput').onchange=async e=>{try{if(!e.target.files.length)return; const p=JSON.parse(await e.target.files[0].text()); validateProject(p);state.trials=p.trials||[];state.raw=p.raw||[];state.sampleMap=p.sampleMap||{};state.selectedGenes=new Set(p.selectedGenes||[]);state.refGene=p.refGene||'';state.calibrator=p.calibrator||'';state.base=p.base||2;state.overrides=p.overrides||{}; restoreFigure(p.figure); $('#efficiency').value=state.base===2?'2':'custom'; $('#customBase').value=state.base; $('#customBaseField').classList.toggle('hidden',state.base===2);syncTrialNames();initializeSelections();invalidate();refreshAll();}catch(err){alert('Could not open project: '+err.message)}};
  $('#downloadPng').onclick=()=>{drawChart();const a=document.createElement('a');a.download=`${slug(graphGenes().join('_')||'qpcr')}_graph.png`;a.href=$('#chart').toDataURL('image/png');a.click();};
  function downloadText(name,text,type){const blob=new Blob([text],{type}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();toast('Exported '+name);setTimeout(()=>URL.revokeObjectURL(a.href),1000)}

  // Demo mirrors the workflow in the supplied workbook (Tbp reference, M1 0G vs M1 10G).
  $('#loadDemo').onclick=()=>{ if(state.raw.length&&!confirm('Add the demo as another trial?'))return; const id='demo'+Date.now().toString(36)+Math.random().toString(36).slice(2,7),trial={id,name:`Trial ${state.trials.length+1}`,fileName:'Demo_qPCR.xlsx',sheetName:'Raw Data',rowCount:18}; const demo=[
    ['M1 0G','Tbp',20.545],['M1 0G','Tbp',20.530],['M1 0G','Tbp',21.178],['M1 10G','Tbp',21.605],['M1 10G','Tbp',20.689],['M1 10G','Tbp',20.927],
    ['M1 0G','Fap',32.537],['M1 0G','Fap',32.128],['M1 0G','Fap',32.832],['M1 10G','Fap',32.496],['M1 10G','Fap',31.516],['M1 10G','Fap',32.178],
    ['M1 0G','Acta2',18.598],['M1 0G','Acta2',19.051],['M1 0G','Acta2',18.506],['M1 10G','Acta2',17.954],['M1 10G','Acta2',17.887],['M1 10G','Acta2',17.643]
  ]; state.trials.push(trial); demo.forEach((d,i)=>{state.raw.push({id:`${id}_${i}`,trialId:id,trialName:trial.name,fileName:trial.fileName,well:String(i+1),sample:d[0],target:d[1],cq:d[2],omit:false,originalCq:d[2]});state.sampleMap[d[0]]=d[0]}); state.refGene='Tbp';state.selectedGenes=new Set(['Fap','Acta2']);state.calibrator='M1 0G';invalidate();refreshAll();};

  function invalidate(){
    state.analysisRun=false;
    state.results=[];
    $('#statsResult').textContent='Data changed. Run a new statistical test after analysis.';
    $('#staleNotice').classList.toggle('hidden',!state.raw.length);
  }
  function toast(message){
    $('#toast').textContent=message; $('#toast').classList.add('visible');
    clearTimeout(window.toastTimer); window.toastTimer=setTimeout(()=>$('#toast').classList.remove('visible'),3500);
  }
  function renderColors(){
    $('#colorSwatches').innerHTML=colors.map((c,i)=>`<label title="Series color ${i+1}"><input type="color" aria-label="Series color ${i+1}" value="${c}" data-color="${i}"><span>${i+1}</span></label>`).join('');
    $$('[data-color]').forEach(el=>el.oninput=()=>{colors[Number(el.dataset.color)]=el.value;$('#palettePreset').value='custom';drawChart();});
  }
  const figureIds=['palettePreset','pointSize','barOpacity','showGrid','graphMode','graphMetric','graphType','errorType','yScale','graphTitle','showThreshold','thresholdValue','thresholdLabel'];
  function getFigure(){return {colors,genes:graphGenes(),controls:Object.fromEntries(figureIds.map(id=>[id,$('#'+id).type==='checkbox'?$('#'+id).checked:$('#'+id).value]))};}
  function restoreFigure(f){if(!f)return;if(Array.isArray(f.genes))state.pendingGraphGenes=f.genes;if(Array.isArray(f.colors)&&f.colors.length===8&&f.colors.every(c=>/^#[0-9a-f]{6}$/i.test(c)))colors=f.colors;for(const id of figureIds){if(f.controls&&id in f.controls){const el=$('#'+id);if(el.type==='checkbox')el.checked=!!f.controls[id];else el.value=f.controls[id];}}renderColors();}
  function validateProject(p){
    if(!p||!Array.isArray(p.raw)||!Array.isArray(p.trials)||!Array.isArray(p.selectedGenes))throw new Error('This is not a qPCR studio project.');
    if(!p.trials.every(t=>typeof t.id==='string'&&/^[\w-]+$/.test(t.id)&&typeof t.name==='string'))throw new Error('Invalid trial records.');
    if(!p.raw.every(r=>typeof r.id==='string'&&/^[\w-]+$/.test(r.id)&&typeof r.sample==='string'&&typeof r.target==='string'&&Number.isFinite(r.cq)&&p.trials.some(t=>t.id===r.trialId)))throw new Error('Invalid raw data records.');
    if(p.base!==undefined&&(!Number.isFinite(p.base)||p.base<=1))throw new Error('Invalid amplification base.');
  }
  $('#palettePreset').onchange=()=>{if(palettes[$('#palettePreset').value])colors=[...palettes[$('#palettePreset').value]];renderColors();drawChart();};
  ['pointSize','barOpacity','showGrid'].forEach(id=>$('#'+id).addEventListener('input',drawChart));
  $('#globalSave').onclick=()=>$('#saveProject').click();
  $('#exploreDemo').onclick=()=>$('#loadDemo').click();
  document.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key==='s'){e.preventDefault();$('#saveProject').click();}});
  const graphCard=$('#page-graphs .chartCard');
  graphCard.insertAdjacentHTML('afterbegin','<div class="figure-preview-heading"><span>FIGURE PREVIEW</span><span class="badge">2× resolution export</span></div>');
  const next=document.createElement('button');next.className='btn primary';next.textContent='Continue to analysis →';next.onclick=()=>showPage('setup');$('#page-import .topbar .actions').append(next);
  const graphNext=document.createElement('button');graphNext.className='btn primary';graphNext.textContent='Open figure studio →';graphNext.onclick=()=>showPage('graphs');$('#page-results .topbar .actions').append(graphNext);
  renderColors();
  document.querySelectorAll('.field').forEach(field=>{const label=field.querySelector('label'),input=field.querySelector('input,select');if(label&&input&&input.id)label.htmlFor=input.id;});
  function compactAxis(ctx,labels,width,count){
    const fontSize=Math.max(8,Math.min(11,(width-150)/Math.max(1,count)*0.65));
    ctx.font=fontSize+'px system-ui';
    const maxWidth=160;
    const longest=Math.min(maxWidth,Math.max(0,...labels.map(label=>ctx.measureText(label).width)));
    return {fontSize,maxWidth,left:Math.max(72,Math.ceil(longest/Math.SQRT2)+12),bottom:Math.max(92,Math.ceil(longest/Math.SQRT2)+32)};
  }
  function drawAngledLabel(ctx,label,x,y,axis){
    ctx.save();ctx.font=axis.fontSize+'px system-ui';
    let text=label;
    if(ctx.measureText(text).width>axis.maxWidth){while(text.length&&ctx.measureText(text+'…').width>axis.maxWidth)text=text.slice(0,-1);text+='…';}
    ctx.translate(x,y);ctx.rotate(-Math.PI/4);ctx.textAlign='right';ctx.textBaseline='middle';ctx.fillStyle='#445167';ctx.fillText(text,0,0);ctx.restore();
  }

  refreshAll();
})();
