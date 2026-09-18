'use strict';
const $=id=>document.getElementById(id);
const escapeHTML=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const state={ready:false,root:'',nodes:[],filter:'',descending:false,sourceChecked:new Set(),sourceHighlight:'',files:[],revision:-1,checked:new Set(),selected:new Set(),anchor:'',focusPane:'source',settings:{printer:'',color:true,duplex:'单面',nup:1,layout:'自动',office:'自动'},job:{state:'idle'},lastSummary:'',generation:0};
let mutation=Promise.resolve(),toastTimer,dialogResolve,polling=false;
const selectModels=new Map();
async function call(method,...args){if(!window.pywebview?.api)throw Error('请通过桌面程序入口启动，不要用浏览器直接打开界面文件。');return window.pywebview.api[method](...args)}
function toast(text){$('toast-text').textContent=text;$('toast').classList.add('visible');clearTimeout(toastTimer);toastTimer=setTimeout(()=>$('toast').classList.remove('visible'),2700)}
function showDialog({title,body,confirm=false,accept='知道了',symbol='info'}){closeSelects();if(dialogResolve)finishDialog(false);$('dialog-title').textContent=title;$('dialog-body').textContent=body;$('dialog-symbol').innerHTML=icon(symbol);$('dialog-cancel').classList.toggle('hidden',!confirm);$('dialog-accept').textContent=accept;$('app-dialog').showModal();requestAnimationFrame(()=>(confirm?$('dialog-cancel'):$('dialog-accept')).focus());return new Promise(resolve=>dialogResolve=resolve)}
function finishDialog(result){if($('app-dialog').open)$('app-dialog').close();const resolve=dialogResolve;dialogResolve=null;resolve?.(result)}
async function showError(error){console.error(error);return showDialog({title:'操作未完成',body:error?.message||String(error)})}
function closeSelects(){document.querySelectorAll('.custom-select.open').forEach(el=>{el.classList.remove('open','open-up');el.querySelector('.select-trigger')?.setAttribute('aria-expanded','false')})}
function setSelect(id,values,value,onchange,placeholder='请选择'){
 const el=$(id);el.classList.remove("open","open-up");selectModels.set(id,{values,value,onchange,placeholder});
 el.innerHTML=`<button class="select-trigger" type="button" aria-haspopup="listbox" aria-expanded="false" ${values.length?'':'disabled'} title="${escapeHTML(value||placeholder)}"><span>${escapeHTML(value||placeholder)}</span>${icon('chevron')}</button><div class="select-menu" role="listbox" aria-label="${escapeHTML(el.getAttribute('aria-label')||id)}">${values.map((item,i)=>`<button type="button" role="option" aria-selected="${item===value}" class="${item===value?'chosen':''}" data-choice="${i}"><span>${escapeHTML(item)}</span>${item===value?icon('check'):''}</button>`).join('')}</div>`;
 el.querySelector('.select-trigger').onclick=e=>{e.stopPropagation();const wasOpen=el.classList.contains('open');closeSelects();if(!wasOpen){el.classList.add('open');el.querySelector('.select-trigger').setAttribute('aria-expanded','true');if(el.getBoundingClientRect().bottom+Math.min(values.length*34+14,235)>innerHeight-20)el.classList.add('open-up')}};
 el.querySelector('.select-menu').onclick=e=>{const btn=e.target.closest('[data-choice]');if(!btn)return;const model=selectModels.get(id);const choice=model.values[Number(btn.dataset.choice)];setSelect(id,model.values,choice,model.onchange,model.placeholder);model.onchange(choice);el.querySelector('.select-trigger').focus();e.stopPropagation()};
 el.onkeydown=e=>{const buttons=[...el.querySelectorAll('[data-choice]')];if(e.key==='Escape'){closeSelects();el.querySelector('.select-trigger').focus();e.stopPropagation()}else if(e.key==='ArrowDown'||e.key==='ArrowUp'){e.preventDefault();if(!el.classList.contains('open'))el.querySelector('.select-trigger').click();const current=buttons.indexOf(document.activeElement);buttons[(current+(e.key==='ArrowDown'?1:-1)+buttons.length)%buttons.length]?.focus()}};
}
function buildSettings(){setSelect('duplex-select',['单面','长边装订','短边装订'],state.settings.duplex,v=>{state.settings.duplex=v;updateFooter()});setSelect('office-select',['自动','优先 WPS','优先 Excel'],state.settings.office,v=>state.settings.office=v);updateLayout();setSelect('printer-select',[],state.settings.printer,v=>state.settings.printer=v,'正在查找打印机…')}
function updateLayout(){const n=Number($('nup-input').value);state.settings.nup=$('nup-input').value;const opts=n<=1?['自动']:n===2?['自动','Vertical（竖排）','Horizontal（并排）']:n===4?['自动（行优先）','列优先']:['自动（行优先）'];if(!opts.includes(state.settings.layout))state.settings.layout=opts[0];setSelect('layout-select',opts,state.settings.layout,v=>state.settings.layout=v);updateFooter()}
function setPrinters(data){if(!data.names.includes(state.settings.printer))state.settings.printer=data.default;setSelect('printer-select',data.names,state.settings.printer,v=>state.settings.printer=v,data.error||'未发现打印机');$('printer-select').title=data.error||state.settings.printer||'选择打印机'}
function fileKind(file){const ext=file.extension;return ['.doc','.docx'].includes(ext)?'word':['.xls','.xlsx'].includes(ext)?'excel':['.ppt','.pptx'].includes(ext)?'powerpoint':['.png','.jpg','.jpeg','.webp','.bmp','.gif','.tif','.tiff'].includes(ext)?'image':ext==='.pdf'?'pdf':'other'}
function fileIcon(file){const kind=fileKind(file);return `<span class="file-icon ${kind}">${icon(kind==='image'?'image':'file')}</span>`}
function loadedFiles(nodes=state.nodes){return nodes.flatMap(node=>node.isDir?loadedFiles(node.children||[]):[node])}
function visibleNodes(nodes=state.nodes,depth=0){return nodes.flatMap(node=>[{node,depth},...(node.isDir&&node.expanded?visibleNodes(node.children||[],depth+1):[])])}
function findNode(path,nodes=state.nodes){for(const node of nodes){if(node.path===path)return node;const child=node.children&&findNode(path,node.children);if(child)return child}return null}
function expandFiltered(nodes){for(const node of nodes)if(node.isDir){node.expanded=true;expandFiltered(node.children||[])}}
function renderSource(){const flat=visibleNodes();$('source-tree').innerHTML=flat.map(({node,depth})=>`<div class="source-row ${node.isDir?'folder-row':''} ${!node.isDir&&!node.supported?'unsupported':''}" data-path="${escapeHTML(node.path)}" data-dir="${node.isDir}" style="--depth:${depth}" title="${escapeHTML(node.path)}" tabindex="0" role="treeitem" ${node.isDir?`aria-expanded="${!!node.expanded}"`:''}>${node.isDir?`<button class="folder-chevron ${node.expanded?'expanded':''}" aria-label="展开或折叠文件夹">${icon('chevron-right')}</button>${icon('folder')}`:`<button class="check-box" role="checkbox" aria-checked="false" aria-label="勾选 ${escapeHTML(node.name)}" ${node.supported?'':'disabled'}></button>${fileIcon(node)}`}<span class="node-name">${escapeHTML(node.name)}</span><span class="node-type">${node.isDir?'文件夹':escapeHTML(node.extension.slice(1).toUpperCase()||'文件')}</span></div>`).join('');
 const files=loadedFiles();const noMatch=!!state.root&&!!state.filter&&!files.length;const empty=!state.root||!state.nodes.length;
 $('source-empty').classList.toggle('hidden',!empty);$('source-empty-copy').textContent=noMatch?'没有匹配的文件，试试其他文件名或扩展名。':state.root?'文件夹为空，请选择其他文件夹。':'选择或拖入文件夹';$('source-empty-copy').style.whiteSpace='pre-line';$('source-empty-button').dataset.action=noMatch?'clear-filter':'folder';$('source-empty-button').innerHTML=icon(noMatch?'close':'folder-plus')+(noMatch?'清除筛选':'选择文件夹');$('source-count').textContent=state.root?`${files.filter(f=>f.supported).length} 个可打印文件`:'尚未选择文件夹';paintSource();}
function paintSource(){document.querySelectorAll('.source-row').forEach(el=>{el.classList.toggle('source-highlight',el.dataset.path===state.sourceHighlight);const box=el.querySelector('.check-box');if(box){const checked=state.sourceChecked.has(el.dataset.path);box.classList.toggle('checked',checked);box.setAttribute('aria-checked',String(checked))}});const files=loadedFiles().filter(f=>f.supported);const checked=files.filter(f=>state.sourceChecked.has(f.path)).length;$('source-selected').textContent=`已勾选 ${checked} 项`;$('select-all-source').textContent=checked===files.length&&files.length?'取消全选':'全选可打印文件'}
async function loadRoot(path,reset=false){if(!path)return;const generation=++state.generation;$('source-scroll').classList.add('source-busy');try{if(reset){const selected=await call('select_root',path);if(selected.error)throw Error(selected.error);state.sourceChecked.clear();state.filter='';$('filter-input').value=''}const result=await call('browse',path,state.filter,state.descending);if(generation!==state.generation)return;if(result.error&&!result.nodes.length)throw Error(result.error);state.root=result.root;state.nodes=result.nodes;if(state.filter)expandFiltered(state.nodes);$('folder-name').textContent=result.name;$('folder-path').textContent=result.root;$('folder-name').title=result.root;$('folder-path').title=result.root;renderSource();if(result.error)toast('部分文件夹无法访问，其余文件已显示')}finally{if(generation===state.generation)$('source-scroll').classList.remove('source-busy')}}
async function toggleFolder(path){const node=findNode(path);if(!node)return;const generation=state.generation;if(node.children===null){const result=await call('browse',path,'',state.descending);if(generation!==state.generation)return;if(result.error)toast(result.error);node.children=result.nodes}node.expanded=!node.expanded;renderSource()}
function queueSnapshot(result){if(!result||result.revision<state.revision)return;state.revision=result.revision;state.files=result.files;const paths=new Set(state.files.map(f=>f.path));state.checked=new Set([...state.checked].filter(p=>paths.has(p)));state.selected=new Set([...state.selected].filter(p=>paths.has(p)));renderQueue()}
function mutateQueue(method,...args){mutation=mutation.then(()=>call(method,...args)).then(result=>{queueSnapshot(result);return result}).catch(showError);return mutation}
function renderQueue(){$('queue-list').innerHTML=state.files.map((file,index)=>`<div class="queue-row" tabindex="-1" role="option" aria-selected="false" data-path="${escapeHTML(file.path)}" title="${escapeHTML(file.path)}"><button class="check-box" role="checkbox" aria-checked="false" aria-label="勾选 ${escapeHTML(file.name)}"></button><span class="queue-index">${String(index+1).padStart(2,'0')}</span>${fileIcon(file)}<div class="queue-file-content"><div class="queue-file-name">${escapeHTML(file.name)}</div></div><span class="file-type-badge">${escapeHTML(file.extension.slice(1).toUpperCase())}</span></div>`).join('');$('queue-count').textContent=state.files.length;$('queue-empty').classList.toggle('hidden',!!state.files.length);paintQueue();updateFooter()}
function paintQueue(){document.querySelectorAll('.queue-row').forEach(el=>{const checked=state.checked.has(el.dataset.path),selected=state.selected.has(el.dataset.path);el.classList.toggle('selected',selected);el.classList.toggle('checked',checked);el.setAttribute('aria-selected',String(selected));const box=el.querySelector('.check-box');box.classList.toggle('checked',checked);box.setAttribute('aria-checked',String(checked))});$('queue-selection').textContent=state.checked.size||state.selected.size?`已勾选 ${state.checked.size} 项 · 高亮 ${state.selected.size} 项`:'勾选与高亮可独立调整顺序'}
function selectQueue(path,e){state.focusPane='queue';const i=state.files.findIndex(f=>f.path===path),anchor=state.files.findIndex(f=>f.path===state.anchor);if(e.shiftKey&&anchor>=0){if(!e.ctrlKey&&!e.metaKey)state.selected.clear();state.files.slice(Math.min(i,anchor),Math.max(i,anchor)+1).forEach(f=>state.selected.add(f.path))}else if(e.ctrlKey||e.metaKey){state.selected.has(path)?state.selected.delete(path):state.selected.add(path);state.anchor=path}else{state.selected=new Set([path]);state.anchor=path}paintQueue();$('queue-scroll').focus({preventScroll:true})}
function updateFooter(){const s=state.settings;$('status-detail').textContent=`${state.files.length} 个文件待打印  ·  ${s.duplex}  ·  ${s.color?'彩色':'黑白'}${Number(s.nup)>1?'  ·  PDF '+s.nup+' 合 1':''}`}
function renderJob(job){state.job=job;const active=job.state==='running';$('print-button').disabled=active;$('cancel-print').disabled=!active;$('print-label').textContent=active?'正在打印…':'开始打印';$('status-dot').classList.toggle('busy',active);$('progress-track').classList.toggle('active',active);$('progress-fill').style.width=(job.total?Math.min(100,job.processed/job.total*100):0)+'%';if(active||job.id){$('status-text').textContent=job.status;$('status-text').title=job.status}if(job.id&&!active&&state.lastSummary!==job.id){state.lastSummary=job.id;const lines=[`已处理 ${job.processed} / ${job.total} 个文件。`,`成功提交 ${job.success} 个，失败 ${job.failed.length} 个${job.cancelled?'，未提交 '+job.cancelled+' 个':''}。`,'','结果表示已提交给打印程序；实际出纸请以打印机状态为准。'];if(job.failed.length)lines.push('',...job.failed.slice(0,10).map(f=>`${f.name}：${f.reason}`));showDialog({title:job.state==='cancelled'?'已停止后续打印':'本批任务已处理',body:lines.join('\n'),symbol:job.failed.length?'info':'check'});$('status-text').textContent='准备就绪'}}
async function addPaths(paths){if(!paths.length)return;const result=await mutateQueue('queue_add',paths);if(result)toast(`已添加 ${result.added} 个文件${result.skipped?'，跳过 '+result.skipped+' 个重复或不支持的文件':''}`)}
async function reorder(selection,direction){const paths=[...(selection==='checked'?state.checked:state.selected)];if(!paths.length){toast(selection==='checked'?'请先勾选要移动的文件':'请先高亮要移动的文件');return}await mutateQueue('queue_move',paths,direction);const row=[...document.querySelectorAll('.queue-row')].find(el=>el.dataset.path===paths[0]);row?.scrollIntoView({block:'nearest'})}
async function removeSelected(){const paths=[...(state.checked.size?state.checked:state.selected)];if(!paths.length)return toast('请先勾选或高亮要移除的文件');await mutateQueue('queue_remove',paths);toast(`已移出 ${paths.length} 项，原始文件未改动`)}
async function requestPrint(){if(state.job.state==='running')return;if(!state.files.length)return showDialog({title:'先放入几份文件吧',body:'双击左侧文件，或将文件、文件夹拖入右侧待打印队列。\n也可以点击“添加文件”直接选择。'});const nup=Number($('nup-input').value);if(!Number.isInteger(nup)||nup<1||nup>16){$('nup-input').focus();return showDialog({title:'检查一下拼版页数',body:'PDF 每张页数需要填写 1～16 的整数。'})}state.settings.nup=nup;if(!state.settings.printer)return showDialog({title:'还没有选择打印机',body:'请先选择可用的打印机。如果列表为空，可以点击旁边的刷新按钮，或检查 Windows 打印机设置。'});await mutation;const confirmed=await showDialog({title:'开始打印这一批文件？',body:`${state.files.length} 个文件将按队列顺序提交到：\n${state.settings.printer}\n\n${state.settings.duplex} · ${state.settings.color?'彩色':'黑白'} · PDF ${nup} 页 / 张\n\n提交后，已进入系统队列的任务需要在 Windows 中取消。`,confirm:true,accept:'开始打印',symbol:'printer'});if(!confirmed)return;const result=await call('start_print',{...state.settings});if(result.error)throw Error(result.error);renderJob(result.job)}
async function requestClose(){const result=await call('window_action','close');if(result?.busy){const ok=await showDialog({title:'停止任务并退出？',body:'将停止尚未提交的文件，并等待当前任务完成清理后退出。\n\n已经进入系统打印队列的任务不会自动撤回。',confirm:true,accept:'停止并退出'});if(ok){$('status-text').textContent='正在停止任务并恢复打印机设置…';await call('stop_and_close')}}}
const actions={
 folder:async()=>{const path=await call('choose_folder');if(path)await loadRoot(path,true)},
 files:async()=>{const result=await call('choose_files');if(result){queueSnapshot(result);toast(`已添加 ${result.added} 个文件${result.skipped?'，跳过 '+result.skipped+' 个':''}`)}},
 filter:async()=>{state.filter=$('filter-input').value.trim();if(state.root)await loadRoot(state.root);else toast('请先选择文件夹')},
 'clear-filter':async()=>{$('filter-input').value='';state.filter='';if(state.root)await loadRoot(state.root)},
 'refresh-source':async()=>{if(state.root)await loadRoot(state.root);else toast('请先选择文件夹')},
 sort:async()=>{state.descending=!state.descending;$('sort-icon').innerHTML=icon(state.descending?'arrow-down':'arrow-up');if(state.root)await loadRoot(state.root)},
 'refresh-printers':async()=>{setPrinters(await call('get_printers'));toast('打印机列表已刷新')},
 'select-all-source':()=>{const files=loadedFiles().filter(f=>f.supported);const clear=files.length&&files.every(f=>state.sourceChecked.has(f.path));files.forEach(f=>clear?state.sourceChecked.delete(f.path):state.sourceChecked.add(f.path));paintSource()},
 'add-checked':async()=>{const paths=loadedFiles().filter(f=>f.supported&&state.sourceChecked.has(f.path)).map(f=>f.path);if(!paths.length)return toast('请先勾选左侧文件');await addPaths(paths);paths.forEach(p=>state.sourceChecked.delete(p));paintSource()},
 'checked-up':()=>reorder('checked',-1),'checked-down':()=>reorder('checked',1),'selected-up':()=>reorder('selected',-1),'selected-down':()=>reorder('selected',1),remove:removeSelected,
 'clear-queue':async()=>{if(!state.files.length)return;if(await showDialog({title:'清空待打印队列？',body:'只移除队列中的条目，不会删除原始文件。\n已经开始的打印任务不受影响。',confirm:true,accept:'清空队列'})){await mutateQueue('queue_clear');toast('队列已清空，原始文件未改动')}},
 'nup-minus':()=>{let n=Number($('nup-input').value);$('nup-input').value=Number.isInteger(n)?Math.max(1,Math.min(16,n-1)):1;updateLayout()},
 'nup-plus':()=>{let n=Number($('nup-input').value);$('nup-input').value=Number.isInteger(n)?Math.max(1,Math.min(16,n+1)):1;updateLayout()},
 print:requestPrint,'cancel-print':async()=>{const r=await call('cancel_print');$('status-text').textContent=r.status;$('cancel-print').disabled=true;toast('已请求停止后续打印')},
 minimize:()=>call('window_action','minimize'),maximize:()=>call('window_action','maximize'),close:requestClose
};
document.addEventListener('click',e=>{if(!e.target.closest('.custom-select'))closeSelects();const button=e.target.closest('[data-action]');if(button&&!button.disabled){Promise.resolve(actions[button.dataset.action]?.()).catch(showError)}const color=e.target.closest('[data-color]');if(color){state.settings.color=color.dataset.color==='true';document.querySelectorAll('[data-color]').forEach(b=>{const active=(b.dataset.color==='true')===state.settings.color;b.classList.toggle('active',active);b.setAttribute('aria-pressed',active)});updateFooter()}});
$('filter-input').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();actions.filter().catch(showError)}if(e.key==='Escape'){e.preventDefault();actions['clear-filter']().catch(showError)}});
$('nup-input').addEventListener('input',updateLayout);
$('source-tree').addEventListener('click',e=>{const row=e.target.closest('.source-row');if(!row)return;state.focusPane='source';const node=findNode(row.dataset.path);if(e.target.closest('.check-box')&&node.supported){state.sourceChecked.has(node.path)?state.sourceChecked.delete(node.path):state.sourceChecked.add(node.path);paintSource();return}if(e.target.closest('.folder-chevron')){toggleFolder(node.path).catch(showError);return}state.sourceHighlight=node.path;paintSource();row.focus({preventScroll:true})});
$('source-tree').addEventListener('dblclick',e=>{const row=e.target.closest('.source-row');if(!row||e.target.closest('button'))return;const node=findNode(row.dataset.path);if(node.isDir)toggleFolder(node.path).catch(showError);else if(node.supported)addPaths([node.path]).catch(showError);else toast('暂不支持打印这种文件类型')});
$('source-tree').addEventListener('keydown',e=>{const row=e.target.closest('.source-row');if(!row)return;const node=findNode(row.dataset.path);if(e.key===' '&&!node.isDir&&node.supported){e.preventDefault();state.sourceChecked.has(node.path)?state.sourceChecked.delete(node.path):state.sourceChecked.add(node.path);paintSource()}if(e.key==='Enter'){e.preventDefault();(node.isDir?toggleFolder(node.path):node.supported?addPaths([node.path]):Promise.resolve()).catch(showError)}if(e.key==='ArrowDown'||e.key==='ArrowUp'){e.preventDefault();const rows=[...document.querySelectorAll('.source-row')];rows[rows.indexOf(row)+(e.key==='ArrowDown'?1:-1)]?.focus()}});
$('queue-list').addEventListener('click',e=>{const row=e.target.closest('.queue-row');if(!row)return;state.focusPane='queue';const path=row.dataset.path;if(e.target.closest('.check-box')){state.checked.has(path)?state.checked.delete(path):state.checked.add(path);paintQueue()}else selectQueue(path,e)});
$('queue-list').addEventListener('dblclick',e=>{const row=e.target.closest('.queue-row');if(row&&!e.target.closest('button'))mutateQueue('queue_remove',[row.dataset.path])});
$('queue-scroll').addEventListener('focus',()=>state.focusPane='queue');
document.addEventListener('keydown',e=>{if($('app-dialog').open)return;const editable=e.target.matches('input,textarea');if(e.ctrlKey&&e.key.toLowerCase()==='o'){e.preventDefault();actions.folder().catch(showError);return}if(e.ctrlKey&&e.key==='Enter'){e.preventDefault();requestPrint().catch(showError);return}if(editable||state.focusPane!=='queue')return;if(e.key==='Delete'){e.preventDefault();removeSelected().catch(showError)}if(e.ctrlKey&&e.key.toLowerCase()==='a'){e.preventDefault();state.selected=new Set(state.files.map(f=>f.path));paintQueue()}if(e.altKey&&(e.key==='ArrowUp'||e.key==='ArrowDown')){e.preventDefault();reorder('selected',e.key==='ArrowUp'?-1:1).catch(showError)}if(e.key===' '&&e.target===$('queue-scroll')){e.preventDefault();state.selected.forEach(p=>state.checked.has(p)?state.checked.delete(p):state.checked.add(p));paintQueue()}});
$('dialog-close').onclick=() => finishDialog(false);$('dialog-cancel').onclick=()=>finishDialog(false);$('dialog-accept').onclick=()=>finishDialog(true);$('app-dialog').addEventListener('cancel',e=>{e.preventDefault();finishDialog(false)});
for(const id of ['source-drop','queue-drop']){const el=$(id);let dragDepth=0;el.addEventListener('dragenter',e=>{e.preventDefault();dragDepth++;el.classList.add('drop-active')});el.addEventListener('dragover',e=>{e.preventDefault();e.dataTransfer.dropEffect='copy'});el.addEventListener('dragleave',()=>{dragDepth=Math.max(0,dragDepth-1);if(!dragDepth)el.classList.remove('drop-active')});el.addEventListener('drop',e=>{e.preventDefault();dragDepth=0;el.classList.remove('drop-active')})}
document.addEventListener('dragover',e=>e.preventDefault());document.addEventListener('drop',e=>e.preventDefault());
window.desktop={
 requestClose:()=>requestClose().catch(showError),
 onNativeDrop:async(target,result)=>{document.querySelectorAll('.drop-active').forEach(el=>el.classList.remove('drop-active'));if(result.error){toast(result.error);return}if(target==='source')await loadRoot(result.root,true);else{queueSnapshot(result);toast(`已添加 ${result.added} 个文件${result.skipped?'，跳过 '+result.skipped+' 个':''}`)}},
 reportError:message=>showError(Error(message))
};
async function initialize() {
  if (state.ready) return;
  state.ready = true;
  try {
    const boot = await call('bootstrap');
    queueSnapshot(boot.queue);
    setPrinters(boot.printers);
    if (boot.root) await loadRoot(boot.root);

    // The icon is decorative. A missing/broken image must not block printing setup.
    const appIcon = document.getElementById('app-icon');
    if (appIcon && typeof boot.app_icon === 'string' && boot.app_icon) {
      appIcon.onload = () => appIcon.classList.remove('hidden');
      appIcon.onerror = () => appIcon.classList.add('hidden');
      appIcon.src = boot.app_icon;
    }

    $('status-text').textContent = '准备就绪';
    renderJob(boot.job);
    setInterval(async () => {
      if (polling) return;
      polling = true;
      try {
        const job = await call('poll_job');
        if (job.id) renderJob(job);
      } catch (error) {
        console.error(error);
      } finally {
        polling = false;
      }
    }, 450);
  } catch (error) {
    state.ready = false;
    $('status-text').textContent = '初始化未完成';
    await showError(error);
  }
}
buildSettings();renderSource();renderQueue();
window.addEventListener('pywebviewready',initialize);if(window.pywebview?.api)initialize();






