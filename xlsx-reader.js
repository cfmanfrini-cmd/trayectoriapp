/* XLSX reader for first worksheet, browser-native, offline. No macros or formulas executed. */
async function readXlsxFirstSheet(file){
 const bytes=new Uint8Array(await file.arrayBuffer());const dv=new DataView(bytes.buffer);let end=-1;
 for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--){if(dv.getUint32(i,true)===0x06054b50){end=i;break;}}
 if(end<0)throw Error('El archivo no parece ser un Excel .xlsx válido.');
 const count=dv.getUint16(end+10,true),central=dv.getUint32(end+16,true);let at=central;const files=new Map();
 for(let i=0;i<count;i++){
  if(dv.getUint32(at,true)!==0x02014b50)throw Error('Directorio ZIP no válido.');
  const method=dv.getUint16(at+10,true),size=dv.getUint32(at+20,true),len=dv.getUint16(at+28,true),extra=dv.getUint16(at+30,true),comment=dv.getUint16(at+32,true),offset=dv.getUint32(at+42,true);
  const name=new TextDecoder().decode(bytes.slice(at+46,at+46+len));files.set(name,{method,size,offset});at+=46+len+extra+comment;
 }
 async function read(name){const f=files.get(name);if(!f)return null;let p=f.offset;if(dv.getUint32(p,true)!==0x04034b50)throw Error('Entrada ZIP dañada.');p+=30+dv.getUint16(p+26,true)+dv.getUint16(p+28,true);let b=bytes.slice(p,p+f.size);
  if(f.method===8){if(!('DecompressionStream' in window))throw Error('Tu navegador no permite leer XLSX sin conexión. Probá con Chrome actualizado o exportá a CSV.');const stream=new Blob([b]).stream().pipeThrough(new DecompressionStream('deflate-raw'));b=new Uint8Array(await new Response(stream).arrayBuffer());}
  else if(f.method!==0)throw Error('Método de compresión de Excel no compatible.');return new TextDecoder().decode(b);
 }
 const xml=txt=>new DOMParser().parseFromString(txt,'application/xml');const tag=(node,name)=>Array.from(node.getElementsByTagName('*')).filter(n=>n.localName===name);
 const sharedText=await read('xl/sharedStrings.xml');const shared=sharedText?tag(xml(sharedText),'si').map(si=>tag(si,'t').map(t=>t.textContent).join('')):[];
 const workbookText=await read('xl/workbook.xml');const relText=await read('xl/_rels/workbook.xml.rels');let sheet='xl/worksheets/sheet1.xml';
 if(workbookText&&relText){const wb=xml(workbookText),rel=xml(relText);const first=tag(wb,'sheet')[0];const rid=first?.getAttribute('r:id');const target=tag(rel,'Relationship').find(x=>x.getAttribute('Id')===rid)?.getAttribute('Target');if(target){sheet=target.startsWith('/')?target.slice(1):target.startsWith('xl/')?target:'xl/'+target.replace(/^\.\//,'');}}
 const sheetText=await read(sheet);if(!sheetText)throw Error('No se encontró la primera hoja de Excel.');
 const doc=xml(sheetText),result=[];for(const row of tag(doc,'row')){const cells=[];for(const cell of Array.from(row.children).filter(n=>n.localName==='c')){const address=cell.getAttribute('r')||'A1',col=address.match(/^[A-Z]+/)?.[0]||'A';let idx=0;for(const letter of col)idx=idx*26+letter.charCodeAt(0)-64;idx--;
  const v=Array.from(cell.children).find(n=>n.localName==='v')?.textContent||'';const t=cell.getAttribute('t');let value=t==='s'?shared[Number(v)]||'':t==='inlineStr'?tag(cell,'t').map(n=>n.textContent).join(''):v;cells[idx]=value;
 }if(cells.some(x=>String(x||'').trim()))result.push(Array.from({length:cells.length},(_,i)=>String(cells[i]||'').trim()));}
 return result;
}
