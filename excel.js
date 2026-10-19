/* Leitura e gravação .xlsx em JavaScript puro: sem bibliotecas externas. */
(()=>{'use strict';
const titles=['Projeto','Descrição do Projeto','Data Programacao','Status Programacao','Tipo Intervencao','Numero PowerON','Equipamentos','Regiao','CHI','Horario Inicio','Horario Fim','Contratada','Solicitar viabilidade Gerador','Observação_1','Observação_2','Decisão Comitê'];
const fields=['projeto','descricao','data','status','tipoIntervencao','poweron','equipamento','regiao','chi','inicio','fim','contratada','solicitacaoGerador','observacao1','observacao2','decisaoComite'];
const norm=v=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'');
const expected=titles.map(norm);
const enc=new TextEncoder(),dec=new TextDecoder();
const clean=v=>String(v??'').replace(/[\r\n]+/g,' ').trim();
const asDate=v=>{if(!v&&v!==0)return '';if(typeof v==='number'||/^4\d{4}(\.\d+)?$/.test(String(v).trim()))return new Date(Date.UTC(1899,11,30)+Math.floor(+v)*86400000).toISOString().slice(0,10);let s=String(v).trim();if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let a=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);return a?`${a[3]}-${a[2].padStart(2,'0')}-${a[1].padStart(2,'0')}`:''};
const asTime=v=>{if(v==null||v==='')return '';let s=String(v).trim();if(/^\d{1,2}:\d{2}/.test(s))return s.slice(0,5).padStart(5,'0');if(/^\d+(\.\d+)?$/.test(s)){let m=Math.round((+s%1)*1440)%1440;return String(Math.floor(m/60)).padStart(2,'0')+':'+String(m%60).padStart(2,'0')}return s};
function normalizeRows(grid){
 const line=grid.findIndex(r=>r.some(c=>norm(c)==='projeto')&&r.some(c=>norm(c)==='dataprogramacao'));
 if(line<0)throw Error('Planilha sem as colunas Projeto e Data Programacao.');
 const names=grid[line].map(norm),indexes=expected.map(x=>names.indexOf(x));
 if(indexes[0]<0||indexes[2]<0||indexes[8]<0)throw Error('Colunas obrigatórias não encontradas: Projeto, Data Programacao e CHI.');
 const result=[];
 for(let i=line+1;i<grid.length;i++){
  let row=grid[i],data={};fields.forEach((key,k)=>data[key]=indexes[k]<0?'':row[indexes[k]]??'');
  data.projeto=clean(data.projeto);if(!data.projeto)continue;
  data.data=asDate(data.data);data.inicio=asTime(data.inicio);data.fim=asTime(data.fim);
  if(!data.data)throw Error('Data inválida no projeto '+data.projeto+'.');
  for(const k of fields.filter(k=>!['chi','data','inicio','fim'].includes(k)))data[k]=clean(data[k]);
  let chi=String(data.chi??'').replace(/\./g,'').replace(',','.');data.chi=Number(chi)||0;
  data.rowExcel=i+2;
  data.comite=/^sim\s*-?\s*comit[eê]$/i.test(data.decisaoComite)||(/^sim$/i.test(data.decisaoComite));
  data.gerador=/viabilizar.{0,35}gerador/i.test(data.solicitacaoGerador);
  let a=data.inicio.split(':').map(Number),b=data.fim.split(':').map(Number);data.duracaoHoras=a.length===2&&b.length===2?((b[0]*60+b[1]-a[0]*60-a[1]+1440)%1440)/60:0;
  result.push(data);
 }
 if(!result.length)throw Error('Nenhum registro válido na planilha.');return result;
}
function csvParse(text){let s=text.replace(/^\ufeff/,''),line=s.slice(0,s.search(/\r?\n/)),sep=(line.match(/;/g)||[]).length>=(line.match(/,/g)||[]).length?';':',';let quoted=false,cell='',row=[],all=[];for(let i=0;i<s.length;i++){const c=s[i];if(quoted){if(c==='"'&&s[i+1]==='"'){cell+='"';i++}else if(c==='"')quoted=false;else cell+=c;}else if(c==='"')quoted=true;else if(c===sep){row.push(cell);cell='';}else if(c==='\r'||c==='\n'){if(c==='\r'&&s[i+1]==='\n')i++;row.push(cell);all.push(row);row=[];cell='';}else cell+=c;}if(row.length||cell){row.push(cell);all.push(row);}return normalizeRows(all);}
async function unzip(buffer){const data=new DataView(buffer);let e=-1;for(let i=data.byteLength-22;i>=Math.max(0,data.byteLength-65557);i--)if(data.getUint32(i,true)===0x06054b50){e=i;break}if(e<0)throw Error('Arquivo ZIP/XLSX inválido.');const length=data.getUint16(e+10,true);let offset=data.getUint32(e+16,true),files={};for(let i=0;i<length;i++){
 if(data.getUint32(offset,true)!==0x02014b50)throw Error('Diretório XLSX inválido.');
 const method=data.getUint16(offset+10,true),size=data.getUint32(offset+20,true),nameLength=data.getUint16(offset+28,true),extra=data.getUint16(offset+30,true),comment=data.getUint16(offset+32,true),local=data.getUint32(offset+42,true),path=dec.decode(new Uint8Array(buffer,offset+46,nameLength));offset+=46+nameLength+extra+comment;
 if(path.endsWith('/'))continue;
 const start=local+30+data.getUint16(local+26,true)+data.getUint16(local+28,true),raw=new Uint8Array(buffer,start,size);
 if(method===0)files[path]=raw.slice();else if(method===8){if(typeof DecompressionStream==='undefined')throw Error('Importar XLSX requer Chrome ou Edge atualizado.');const stream=new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));files[path]=new Uint8Array(await new Response(stream).arrayBuffer())}else throw Error('Compactação XLSX não suportada: '+method);
 }return files;}
const xml=s=>new DOMParser().parseFromString(s,'application/xml');
const tags=(o,n)=>Array.from(o.getElementsByTagNameNS('*',n));
const indexCol=name=>{let n=0;for(let c of name)n=n*26+c.charCodeAt(0)-64;return n-1};
function sheetRows(text,shared){let all=[],rx=/<row\b[^>]*>[\s\S]*?<\/row>/g,m;
 while((m=rx.exec(text))){let raw=m[0];if(!/<(?:v|is)(?:[\s>])/.test(raw))continue;
 let doc=xml('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+raw+'</worksheet>');if(tags(doc,'parsererror').length)continue;
 let r=tags(doc,'row')[0],a=[];for(let c of tags(r,'c')){let ref=(c.getAttribute('r')||'').match(/^[A-Z]+/),idx=ref?indexCol(ref[0]):a.length,type=c.getAttribute('t'),v=tags(c,'v')[0];let value=v?v.textContent:'';if(type==='s')value=shared[+value]??'';else if(type==='inlineStr')value=tags(c,'t').map(x=>x.textContent).join('');if(idx<90)a[idx]=value;}
 if(a.some(v=>v!==''&&v!=null))all.push(a);
 }return all;}
async function readXlsx(file){const contents=await unzip(await file.arrayBuffer()),workbook=contents['xl/workbook.xml'];if(!workbook)throw Error('Não é um XLSX válido.');let wb=xml(dec.decode(workbook)),rels=xml(dec.decode(contents['xl/_rels/workbook.xml.rels'])),paths=new Map(tags(rels,'Relationship').map(x=>[x.getAttribute('Id'),x.getAttribute('Target')])),ss=[];
 if(contents['xl/sharedStrings.xml']){let doc=xml(dec.decode(contents['xl/sharedStrings.xml']));ss=tags(doc,'si').map(x=>tags(x,'t').map(y=>y.textContent).join(''))}
 let sheets=tags(wb,'sheet').sort((a,b)=>Number(b.getAttribute('name')==='Base_Sigeo')-Number(a.getAttribute('name')==='Base_Sigeo'));let errors=[];
 for(let sh of sheets){let id=sh.getAttribute('r:id')||sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'),path=paths.get(id);if(!path)continue;path=path.startsWith('/')?path.slice(1):path.startsWith('xl/')?path:'xl/'+path;path=path.replace(/\/\.\//g,'/');if(!contents[path])continue;try{return normalizeRows(sheetRows(dec.decode(contents[path]),ss))}catch(e){errors.push(sh.getAttribute('name')+': '+e.message)}}throw Error('Nenhuma aba com dados compatíveis. '+errors.join(' | '));}
const quote=v=>{let s=String(v??'').replace(/[\r\n]+/g,' ');return /[;"\r\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s};
function csvText(rows){return '\ufeff'+[titles,...rows.map(r=>fields.map(k=>k==='data'?String(r[k]||'').replace(/^(\d{4})-(\d{2})-(\d{2})$/,'$3/$2/$1'):r[k]))].map(r=>r.map(quote).join(';')).join('\r\n');}
function crc32(b){if(!crc32.tab){crc32.tab=Array.from({length:256},(_,n)=>{let c=n;for(let i=0;i<8;i++)c=c&1?0xedb88320^(c>>>1):c>>>1;return c>>>0})}let c=0xffffffff;for(let v of b)c=crc32.tab[(c^v)&255]^(c>>>8);return(c^0xffffffff)>>>0}
function zipWrite(entries){let chunks=[],directory=[],offset=0;for(const [path,text] of entries){const name=enc.encode(path),data=enc.encode(text),crc=crc32(data);let h=new Uint8Array(30+name.length),a=new DataView(h.buffer);a.setUint32(0,0x04034b50,true);a.setUint16(4,20,true);a.setUint32(14,crc,true);a.setUint32(18,data.length,true);a.setUint32(22,data.length,true);a.setUint16(26,name.length,true);h.set(name,30);chunks.push(h,data);let d=new Uint8Array(46+name.length),dv=new DataView(d.buffer);dv.setUint32(0,0x02014b50,true);dv.setUint16(4,20,true);dv.setUint16(6,20,true);dv.setUint32(16,crc,true);dv.setUint32(20,data.length,true);dv.setUint32(24,data.length,true);dv.setUint16(28,name.length,true);dv.setUint32(42,offset,true);d.set(name,46);directory.push(d);offset+=h.length+data.length}let size=directory.reduce((s,r)=>s+r.length,0);chunks.push(...directory);let e=new Uint8Array(22),v=new DataView(e.buffer);v.setUint32(0,0x06054b50,true);v.setUint16(8,directory.length,true);v.setUint16(10,directory.length,true);v.setUint32(12,size,true);v.setUint32(16,offset,true);chunks.push(e);return new Blob(chunks,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});}
const escape=s=>String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
function excelXml(rows){let matrix=[titles,...rows.map(r=>fields.map(k=>r[k]??''))];const letters=i=>{let s='';for(i++;i;i=Math.floor((i-1)/26))s=String.fromCharCode(65+(i-1)%26)+s;return s};let xml=matrix.map((row,i)=>'<row r="'+(i+1)+'">'+row.map((v,j)=>{if(v==null||v==='')return '';let cell=letters(j)+(i+1);if(i>0&&j===8)return '<c r="'+cell+'"><v>'+Number(v||0)+'</v></c>';return '<c r="'+cell+'" t="inlineStr"><is><t>'+escape(v)+'</t></is></c>'}).join('')+'</row>').join('');return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>'+xml+'</sheetData><autoFilter ref="A1:P'+matrix.length+'"/></worksheet>'}
function toXlsx(rows){let ct='<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>';let r='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';let wb='<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Base_Sigeo" sheetId="1" r:id="rId1"/></sheets></workbook>';let wr='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';return zipWrite([['[Content_Types].xml',ct],['_rels/.rels',r],['xl/workbook.xml',wb],['xl/_rels/workbook.xml.rels',wr],['xl/worksheets/sheet1.xml',excelXml(rows)]])}
async function parse(file){return /\.csv$/i.test(file.name)?csvParse(await file.text()):/\.xlsx$|\.xlsm$/i.test(file.name)?await readXlsx(file):Promise.reject(Error('Formato não suportado: use .xlsx ou .csv'))}
window.PIEExcel={parse,toXlsx,csvText,titles,fields};
})();
