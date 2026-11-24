/* KalibriOffice 1.2: editing tools shared by the three desktop editors. */
function toolDialog(title,body,onSave,label='Применить'){
 const d=document.createElement('dialog');d.className='toolsDialog';
 d.innerHTML=`<h2>${esc(title)}</h2>${body}<div class="dialogActions"><button data-cancel>Отмена</button><button class="primary" data-save>${esc(label)}</button></div>`;
 document.body.append(d);d.onclose=()=>d.remove();d.querySelector('[data-cancel]').onclick=()=>d.close();
 d.querySelector('[data-save]').onclick=()=>{try{if(onSave(d)!==false)d.close()}catch(e){toast(e.message)}};d.showModal();return d;
}
function findReplace(){
 toolDialog('Поиск и замена',`<label>Найти<input id="findText" autofocus></label><label>Заменить на<input id="replaceText"></label><p class="hint">Замена всех совпадений с учётом регистра. Оформление абзацев сохраняется.</p>`,d=>{
 const from=d.querySelector('#findText').value,to=d.querySelector('#replaceText').value;if(!from){toast('Введи текст для поиска');return false}
 const walker=document.createTreeWalker($('editor'),NodeFilter.SHOW_TEXT);const nodes=[];while(walker.nextNode())nodes.push(walker.currentNode);let count=0;
 for(const n of nodes){const parts=n.nodeValue.split(from);count+=parts.length-1;if(parts.length>1)n.nodeValue=parts.join(to)}
 dirty();toast('Заменено совпадений: '+count);
 },'Заменить все');
}
function findText(){
 toolDialog('Найти в документе','<label>Текст<input id="searchText" autofocus></label>',d=>{
 const text=d.querySelector('#searchText').value;if(!text)return false;
 const walker=document.createTreeWalker($('editor'),NodeFilter.SHOW_TEXT);while(walker.nextNode()){const n=walker.currentNode,pos=n.nodeValue.toLocaleLowerCase('ru').indexOf(text.toLocaleLowerCase('ru'));if(pos>=0){const r=document.createRange();r.setStart(n,pos);r.setEnd(n,pos+text.length);const sel=window.getSelection();sel.removeAllRanges();sel.addRange(r);n.parentElement.scrollIntoView({block:'center',behavior:'smooth'});return true}}
 toast('Совпадений нет');return false;
 },'Найти');
}
function parseGridRange(text){
 const m=String(text).trim().toUpperCase().match(/^([A-Z])(\d{1,3})(?::([A-Z])(\d{1,3}))?$/);if(!m)throw Error('Укажи диапазон, например A1:D20');
 const c1=m[1].charCodeAt(0),r1=+m[2],c2=(m[3]||m[1]).charCodeAt(0),r2=+(m[4]||m[2]);
 if(r1<1||r2>100||r1>r2||c1>c2)throw Error('Диапазон должен быть в пределах A1:Z100');return {c1,r1,c2,r2};
}
function formattedCell(address){
 const value=calc(address),format=documents.sheet.formats?.[address];if(value===''||String(value).startsWith('#')||!format||!Number.isFinite(Number(value)))return value;
 const n=Number(value);return new Intl.NumberFormat('ru-RU',format==='percent'?{style:'percent',maximumFractionDigits:2}:format==='currency'?{style:'currency',currency:'RUB'}:{minimumFractionDigits:2,maximumFractionDigits:2}).format(n);
}
function formatRange(){
 toolDialog('Формат ячеек',`<label>Диапазон<input id="formatRange" value="${esc(selected)}"></label><label>Формат<select id="numberFormat"><option value="general">Общий</option><option value="number">Число (2 знака)</option><option value="percent">Проценты</option><option value="currency">Денежный (₽)</option></select></label>`,d=>{
 const range=parseGridRange(d.querySelector('#formatRange').value),f=d.querySelector('#numberFormat').value;documents.sheet.formats??={};
 for(let r=range.r1;r<=range.r2;r++)for(let c=range.c1;c<=range.c2;c++){const a=String.fromCharCode(c)+r;if(f==='general')delete documents.sheet.formats[a];else documents.sheet.formats[a]=f}
 refreshCells();dirty();
 });
}
function sortRange(){
 toolDialog('Сортировка строк',`<label>Диапазон<input id="sortRange" value="A1:D20"></label><label>Столбец сортировки<input id="sortColumn" value="${selected[0]}" maxlength="1"></label><label>Порядок<select id="sortOrder"><option value="asc">По возрастанию</option><option value="desc">По убыванию</option></select></label><label><input type="checkbox" id="sortHeader" checked> Первая строка — заголовок</label><p class="hint">Перемещаются целые строки внутри выбранного диапазона. Диапазоны с формулами не сортируются.</p>`,d=>{
 const range=parseGridRange(d.querySelector('#sortRange').value),col=d.querySelector('#sortColumn').value.toUpperCase().charCodeAt(0),first=range.r1+(d.querySelector('#sortHeader').checked?1:0);
 if(!Number.isFinite(col)||col<range.c1||col>range.c2)throw Error('Столбец должен входить в выбранный диапазон');
 const rows=[];for(let r=first;r<=range.r2;r++){const data={},formats={};for(let c=range.c1;c<=range.c2;c++){const a=String.fromCharCode(c)+r,v=documents.sheet.cells[a]||'';if(v.startsWith('='))throw Error('В диапазоне есть формулы. Выбери диапазон с обычными значениями.');data[c]=v;formats[c]=documents.sheet.formats?.[a]}rows.push({data,formats})}
 const direction=d.querySelector('#sortOrder').value==='asc'?1:-1;
 rows.sort((a,b)=>{const x=a.data[col],y=b.data[col];if(x===''||y==='')return x===''?(y===''?0:1):-1;return direction*(Number.isFinite(Number(x))&&Number.isFinite(Number(y))?Number(x)-Number(y):x.localeCompare(y,'ru',{numeric:true}))});
 documents.sheet.formats??={};rows.forEach((row,i)=>{for(let c=range.c1;c<=range.c2;c++){const a=String.fromCharCode(c)+(first+i);documents.sheet.cells[a]=row.data[c];if(row.formats[c])documents.sheet.formats[a]=row.formats[c];else delete documents.sheet.formats[a]}});refreshCells();dirty();toast('Строки отсортированы');
 });
}
function moveSlide(direction){capture();const list=documents.slides.slides,next=idx+direction;if(next<0||next>=list.length)return;[list[idx],list[next]]=[list[next],list[idx]];idx=next;drawSlides();dirty();ribbon()}
function speakerNotes(){capture();toolDialog('Заметки докладчика',`<label>Текст заметок<textarea id="speakerNotes" rows="8">${esc(documents.slides.slides[idx].notes||'')}</textarea></label><p class="hint">Заметки не показываются на слайде. Они сохраняются в проекте и экспортируются в PPTX.</p>`,d=>{documents.slides.slides[idx].notes=d.querySelector('#speakerNotes').value;dirty()})}
const originalRibbon=ribbon;
ribbon=function(){originalRibbon();if(!['Главная','Вставка'].includes(tab))return;let content='';
 if(kind==='write')content=group(btn('Найти','findText()')+btn('Заменить','findReplace()'),'Редактирование');
 if(kind==='sheet')content=group(btn('Сортировка','sortRange()')+btn('Формат ячеек','formatRange()'),'Обработка данных');
 if(kind==='slides')content=group(btn('Выше','moveSlide(-1)')+btn('Ниже','moveSlide(1)')+btn('Заметки','speakerNotes()'),'Порядок и заметки');
 $('ribbon').insertAdjacentHTML('beforeend',content);document.querySelectorAll('#ribbon button').forEach(b=>b.onmousedown=e=>e.preventDefault());
};
refreshCells=function(){document.querySelectorAll('[data-a]').forEach(el=>el.textContent=formattedCell(el.dataset.a))};
const originalDraw=draw;draw=function(){originalDraw();if(kind==='sheet')refreshCells()};
document.addEventListener('keydown',e=>{if(kind==='write'&&e.ctrlKey&&['f','h'].includes(e.key.toLowerCase())){e.preventDefault();e.stopImmediatePropagation();e.key.toLowerCase()==='f'?findText():findReplace()}});
draw();
