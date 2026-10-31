'use strict';
const input=document.querySelector('#file-input');
const dropzone=document.querySelector('#dropzone');
const error=document.querySelector('#error');
const reader=document.querySelector('#lector');
const fileName=document.querySelector('#file-name');
const fileKind=document.querySelector('#file-kind');
const paper=document.querySelector('#paper');
const textContainer=document.querySelector('#text-container');
const pdfContainer=document.querySelector('#pdf-container');
const pdfViewer=document.querySelector('#pdf-viewer');
const readerNote=document.querySelector('#reader-note');
const zoomLabel=document.querySelector('#zoom-label');
const download=document.querySelector('#download');
let currentUrl=null,zoom=100,openRequest=0;
const MAX_SIZE=50*1024*1024;
document.querySelectorAll('[data-open]').forEach(button=>button.addEventListener('click',()=>input.click()));
input.addEventListener('change',()=>{if(input.files[0])openFile(input.files[0]);input.value=''});
['dragenter','dragover'].forEach(type=>dropzone.addEventListener(type,event=>{event.preventDefault();dropzone.classList.add('dragging')}));
['dragleave','drop'].forEach(type=>dropzone.addEventListener(type,event=>{event.preventDefault();dropzone.classList.remove('dragging')}));
dropzone.addEventListener('drop',event=>{if(event.dataTransfer.files[0])openFile(event.dataTransfer.files[0])});
window.addEventListener('dragover',event=>event.preventDefault());window.addEventListener('drop',event=>event.preventDefault());
function showError(message){error.textContent=message;error.hidden=false}
function renderMarkdown(source){paper.replaceChildren();let list=null;for(const line of source.split(/\r?\n/)){let match;if((match=line.match(/^(#{1,3})\s+(.+)$/))){list=null;const heading=document.createElement('h'+match[1].length);heading.textContent=match[2];paper.append(heading)}else if((match=line.match(/^\s*[-*]\s+(.+)$/))){if(!list){list=document.createElement('ul');paper.append(list)}const item=document.createElement('li');item.textContent=match[1];list.append(item)}else{list=null;const paragraph=document.createElement('p');paragraph.textContent=line||'\u00a0';paper.append(paragraph)}}}
async function extractDocx(file){const zip=await JSZip.loadAsync(file);const entry=zip.file('word/document.xml');if(!entry)throw new Error('Este archivo no contiene un documento DOCX válido.');if(entry._data&&entry._data.uncompressedSize>15*1024*1024)throw new Error('El contenido del documento es demasiado grande para mostrarlo.');const xml=await entry.async('string');const doc=new DOMParser().parseFromString(xml,'application/xml');if(doc.querySelector('parsererror'))throw new Error('No se pudo leer el contenido de este DOCX.');const namespace='http://schemas.openxmlformats.org/wordprocessingml/2006/main';const paragraphs=Array.from(doc.getElementsByTagNameNS(namespace,'p'));return paragraphs.map(p=>Array.from(p.getElementsByTagNameNS(namespace,'t'),node=>node.textContent).join('')).join('\n')}
async function openFile(file){const request=++openRequest;error.hidden=true;const ext=file.name.split('.').pop().toLowerCase();if(!['pdf','docx','txt','md','markdown'].includes(ext)){showError('Formato no compatible. Selecciona un archivo PDF, DOCX, TXT o Markdown.');return}if(file.size>MAX_SIZE){showError('El archivo supera el límite de 50 MB.');return}let contents='';try{if(ext==='docx')contents=await extractDocx(file);else if(ext!=='pdf')contents=await file.text()}catch(e){if(request===openRequest)showError(e.message||'No se pudo abrir el archivo. Comprueba que no esté dañado o protegido.');return}if(request!==openRequest)return;if(currentUrl)URL.revokeObjectURL(currentUrl);currentUrl=URL.createObjectURL(file);fileName.textContent=file.name;fileKind.textContent=ext.toUpperCase()+' · '+(file.size/1024/1024).toFixed(2)+' MB';download.href=currentUrl;download.download=file.name;zoom=100;setZoom();const isPdf=ext==='pdf';pdfContainer.hidden=!isPdf;textContainer.hidden=isPdf;document.querySelector('#zoom-in').hidden=isPdf;document.querySelector('#zoom-out').hidden=isPdf;zoomLabel.hidden=isPdf;if(isPdf){pdfViewer.src=currentUrl;readerNote.textContent='Usa los controles del visor de tu navegador para buscar, ampliar o imprimir el PDF.'}else{pdfViewer.removeAttribute('src');if(ext==='md'||ext==='markdown')renderMarkdown(contents);else{paper.replaceChildren();paper.textContent=contents}readerNote.textContent=ext==='docx'?'Se muestra el texto extraído del DOCX. El diseño, las imágenes y algunos elementos pueden diferir del original.':'El archivo se lee localmente en tu navegador.'}reader.hidden=false;reader.scrollIntoView({behavior:matchMedia('(prefers-reduced-motion: reduce)').matches?'auto':'smooth',block:'start'})}
function setZoom(){zoomLabel.textContent=zoom+'%';paper.style.fontSize=(16*zoom/100)+'px'}
document.querySelector('#zoom-out').addEventListener('click',()=>{zoom=Math.max(70,zoom-10);setZoom()});document.querySelector('#zoom-in').addEventListener('click',()=>{zoom=Math.min(160,zoom+10);setZoom()});
