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
  async function showPage(name){
    $$('.nav button').forEach(b=>b.classList.toggle('active',b.dataset.page===name));
    $$('.page').forEach(p=>p.classList.toggle('active',p.id===`page-${name}`));
    if((name==='graphs'||name==='results') && state.raw.length && !state.analysisRun){ if(!await runAnalysis()) return; renderResults(); renderGraphControls(); } if(name==='graphs') { renderPlotEditor(); drawChart(); }
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

  $('#runAnalysis').onclick=async()=>{ if(await runAnalysis()){renderResults();renderGraphControls();showPage('results');} };
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

  // Export/save
  $('#exportResults').onclick=()=>{if(!state.results.length){alert('Run analysis first.');return;}const h=['Trial','Gene','Condition','Mean Target Ct','Mean Reference Ct','dCt','Calibrator dCt','ddCt','Relative Expression','log2FC','Target Technical Replicates','Reference Technical Replicates'];const data=state.results.map(r=>[r.trial,r.gene,r.condition,r.targetMean,r.referenceMean,r.dct,r.calibratorDct,r.ddct,r.expression,r.log2fc,r.nTarget,r.nReference]);downloadText('qpcr_results.csv',[h,...data].map(row=>row.map(csvCell).join(',')).join('\n'),'text/csv');};
  function csvCell(v){let s=v==null?'':String(v); if(typeof v==='string' && /^[=+@-]/.test(s))s="'"+s;return /[",\n]/.test(s)?`"${s.replace(/"/g,'""')}"`:s}
  $('#saveProject').onclick=()=>{const payload={version:2,figure:getFigure(),trials:state.trials,raw:state.raw,sampleMap:state.sampleMap,selectedGenes:[...state.selectedGenes],refGene:state.refGene,calibrator:state.calibrator,base:state.base,overrides:state.overrides};downloadText('qpcr_project.json',JSON.stringify(payload,null,2),'application/json');};
  $('#loadProjectBtn').onclick=()=>$('#projectInput').click(); $('#projectInput').onchange=async e=>{try{if(!e.target.files.length)return; const p=JSON.parse(await e.target.files[0].text()); validateProject(p);state.trials=p.trials||[];state.raw=p.raw||[];state.sampleMap=p.sampleMap||{};state.selectedGenes=new Set(p.selectedGenes||[]);state.refGene=p.refGene||'';state.calibrator=p.calibrator||'';state.base=p.base||2;state.overrides=p.overrides||{}; restoreFigure(p.figure); $('#efficiency').value=state.base===2?'2':'custom'; $('#customBase').value=state.base; $('#customBaseField').classList.toggle('hidden',state.base===2);syncTrialNames();initializeSelections();invalidate();refreshAll();}catch(err){alert('Could not open project: '+err.message)}};
  function downloadText(name,text,type){const blob=new Blob([text],{type}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();toast('Exported '+name);setTimeout(()=>URL.revokeObjectURL(a.href),1000)}

  // Demo mirrors the workflow in the supplied workbook (Tbp reference, M1 0G vs M1 10G).
  $('#loadDemo').onclick=()=>{ if(state.raw.length&&!confirm('Add the demo as another trial?'))return; const id='demo'+Date.now().toString(36)+Math.random().toString(36).slice(2,7),trial={id,name:`Trial ${state.trials.length+1}`,fileName:'Demo_qPCR.xlsx',sheetName:'Raw Data',rowCount:18}; const demo=[
    ['M1 0G','Tbp',20.545],['M1 0G','Tbp',20.530],['M1 0G','Tbp',21.178],['M1 10G','Tbp',21.605],['M1 10G','Tbp',20.689],['M1 10G','Tbp',20.927],
    ['M1 0G','Fap',32.537],['M1 0G','Fap',32.128],['M1 0G','Fap',32.832],['M1 10G','Fap',32.496],['M1 10G','Fap',31.516],['M1 10G','Fap',32.178],
    ['M1 0G','Acta2',18.598],['M1 0G','Acta2',19.051],['M1 0G','Acta2',18.506],['M1 10G','Acta2',17.954],['M1 10G','Acta2',17.887],['M1 10G','Acta2',17.643]
  ]; state.trials.push(trial); demo.forEach((d,i)=>{state.raw.push({id:`${id}_${i}`,trialId:id,trialName:trial.name,fileName:trial.fileName,well:String(i+1),sample:d[0],target:d[1],cq:d[2],omit:false,originalCq:d[2]});state.sampleMap[d[0]]=d[0]}); state.refGene='Tbp';state.selectedGenes=new Set(['Fap','Acta2']);state.calibrator='M1 0G';invalidate();refreshAll();};

  function invalidate(){
    dataRevision++;previewRevision++;clearTimeout(previewTimer);
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

  let dataRevision=0, analysisJob=null, previewRevision=0, previewTimer=null, previewBusy=false, previewQueued=false;
  async function pythonRequest(action,payload){
    if(!window.qpcrPython)throw new Error('Open qPCR studio as a desktop app to use the Python engine.');
    return window.qpcrPython.request(action,payload);
  }
  async function runAnalysis(){
    if(!state.raw.length){toast('Import qPCR data first.');return false;}
    if(!state.refGene||!state.calibrator||!state.selectedGenes.size){toast('Choose target genes, a reference, and a calibrator.');return false;}
    if(analysisJob)return analysisJob;
    const revision=dataRevision;
    $('#runAnalysis').disabled=true;$('#runAnalysis').textContent='Analyzing…';
    const payload={raw:state.raw,trials:state.trials,sampleMap:state.sampleMap,selectedGenes:[...state.selectedGenes],refGene:state.refGene,calibrator:state.calibrator,base:state.base};
    analysisJob=(async()=>{
      try{
        const response=await pythonRequest('analyze',payload);
        if(revision!==dataRevision){toast('Data changed during analysis. Run analysis again.');return false;}
        state.results=response.results;state.analysisRun=true;$('#staleNotice').classList.add('hidden');toast('Python analysis complete · '+state.results.length+' trial-level values');return true;
      }catch(error){toast(error.message);return false;}
      finally{$('#runAnalysis').disabled=false;$('#runAnalysis').textContent='Run analysis';analysisJob=null;}
    })();
    return analysisJob;
  }
  function figurePayload(format='png',exporting=false){
    const metric=$('#graphMetric').value, genes=graphGenes();
    return {...getFigure(),genes,rows:state.results.filter(r=>genes.includes(r.gene)&&graphRowVisible(r,metric)).map(r=>({...r,value:valueFor(r,metric,true)})),format,export:exporting,width:Math.min(820,Math.max(320,$('#chart').parentElement.clientWidth||720))};
  }
  function drawChart(){
    previewRevision++;clearTimeout(previewTimer);
    if(!$('#page-graphs').classList.contains('active'))return;
    if(!state.analysisRun||!graphGenes().length){$('#chart').removeAttribute('src');$('#chart').hidden=true;$('#chartStatus').textContent='Run analysis and select a gene to create a figure.';return;}
    $('#chartStatus').textContent='Rendering figure…';$('#chart').setAttribute('aria-busy','true');
    previewTimer=setTimeout(renderPythonPreview,180);
  }
  async function renderPythonPreview(){
    if(previewBusy){previewQueued=true;return;}
    if(!state.analysisRun||!$('#page-graphs').classList.contains('active'))return;
    previewBusy=true;previewQueued=false;
    const revision=previewRevision;
    try{
      const payload=figurePayload(),response=await pythonRequest('render',payload);
      if(revision!==previewRevision)return;
      const chart=$('#chart');chart.src='data:'+response.mime+';base64,'+response.data;chart.hidden=false;
      chart.style.width=response.width+'px';chart.dataset.engine=response.engine;chart.dataset.labelAngle=response.labelAngle;
      $('#chartStatus').textContent=response.omitted?response.omitted+' nonpositive values omitted on log scale.':'Rendered with Matplotlib + Seaborn';
      $('#chartLegend').textContent='Each point = one biological trial. '+(Object.keys(state.overrides).length?'Manual plot overrides are saved with the project.':'');
    }catch(error){if(revision===previewRevision){$('#chart').hidden=true;$('#chartStatus').textContent='Could not render figure: '+error.message;}}
    finally{previewBusy=false;if(revision===previewRevision)$('#chart').setAttribute('aria-busy','false');if(previewQueued||revision!==previewRevision){previewQueued=false;clearTimeout(previewTimer);previewTimer=setTimeout(renderPythonPreview,120);}}
  }
  async function exportFigure(format){
    if(!state.analysisRun||!graphGenes().length){toast('Run analysis and select a gene before exporting.');return;}
    const revision=previewRevision;
    const buttons=[$('#downloadPng'),$('#downloadSvg'),$('#downloadPdf')];buttons.forEach(b=>b.disabled=true);
    try{
      const response=await pythonRequest('render',figurePayload(format,true));
      if(revision!==previewRevision){toast('Figure settings changed during export. Please export again.');return;}
      const a=document.createElement('a');a.download=slug(graphGenes().join('_')||'qpcr')+'_graph.'+format;a.href='data:'+response.mime+';base64,'+response.data;a.click();toast(format.toUpperCase()+' figure exported');
    }catch(error){toast('Export failed: '+error.message);}
    finally{buttons.forEach(b=>b.disabled=false);}
  }
  async function runStats(){
    if(!state.analysisRun){toast('Run analysis first.');return;}
    const gene=$('#statGene').value, metric=$('#graphMetric').value, use=$('#statsUseOverrides').checked;
    const rows=state.results.filter(r=>r.gene===gene).map(r=>({...r,value:valueFor(r,metric,use)}));
    const payload={rows,test:$('#statTest').value,pair:$('#statPair').value==='auto'?null:$('#statPair').value.split('|||'),metric,useOverrides:use};
    const revision=dataRevision,signature=JSON.stringify(payload);$('#runStats').disabled=true;$('#statsResult').textContent='Running SciPy statistical test…';
    try{
      const result=await pythonRequest('stats',payload);
      const current={rows:state.results.filter(r=>r.gene===$('#statGene').value).map(r=>({...r,value:valueFor(r,$('#graphMetric').value,$('#statsUseOverrides').checked)})),test:$('#statTest').value,pair:$('#statPair').value==='auto'?null:$('#statPair').value.split('|||'),metric:$('#graphMetric').value,useOverrides:$('#statsUseOverrides').checked};
      if(revision!==dataRevision||signature!==JSON.stringify(current)){$('#statsResult').textContent='Settings changed. Run the statistical test again.';return;}
      $('#statsResult').innerHTML=`<strong>${esc(gene)} · ${esc(result.name)}</strong><br>${esc(result.detail)}<br>p = <span class="${result.p<.05?'pSig':'pNS'}">${result.p<.0001?'&lt; 0.0001':result.p.toFixed(4)}</span><div class="footerNote">SciPy · ${esc(metric)} · ${use?'using plot overrides':'calculated values only'} · no multiple-comparison correction</div>`;
    }catch(error){$('#statsResult').textContent=error.message;}
    finally{$('#runStats').disabled=false;}
  }
  $('#runStats').onclick=runStats;
  ['statGene','statTest','statPair','statsUseOverrides','graphMetric'].forEach(id=>$('#'+id).addEventListener('change',()=>{$('#statsResult').textContent='Settings changed. Run a statistical test for this selection.';}));
  window.addEventListener('resize',drawChart);
  $('#downloadPng').onclick=()=>exportFigure('png');
  $('#downloadPng').insertAdjacentHTML('afterend','<button class="btn" id="downloadSvg">Export SVG</button><button class="btn" id="downloadPdf">Export PDF</button>');
  $('#downloadSvg').onclick=()=>exportFigure('svg');$('#downloadPdf').onclick=()=>exportFigure('pdf');
  pythonRequest('health',{}).then(()=>{$('#saveStatus').textContent='Python engine ready';}).catch(error=>{$('#saveStatus').textContent='Python unavailable';toast(error.message);});

  refreshAll();
})();
