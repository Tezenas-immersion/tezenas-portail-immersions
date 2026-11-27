// Local exports: no CDN, no transmission of candidate data to third parties.
const enc=new TextEncoder();
const crcTable=Uint32Array.from({length:256},(_,n)=>{for(let i=0;i<8;i++)n=(n>>>1)^((n&1)?0xedb88320:0);return n>>>0;});
function crc32(b){let c=0xffffffff;for(const x of b)c=(c>>>8)^crcTable[(c^x)&255];return (c^0xffffffff)>>>0;}
export function zip(files){
 const parts=[],central=[];let offset=0,size=0;
 for(const f of files){
  const name=enc.encode(f.name),data=typeof f.data==="string"?enc.encode(f.data):new Uint8Array(f.data);
  const h=new Uint8Array(30),v=new DataView(h.buffer),c=new Uint8Array(46),w=new DataView(c.buffer),crc=crc32(data);
  v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint16(6,2048,true);v.setUint16(12,33,true);v.setUint32(14,crc,true);v.setUint32(18,data.length,true);v.setUint32(22,data.length,true);v.setUint16(26,name.length,true);
  w.setUint32(0,0x02014b50,true);w.setUint16(4,20,true);w.setUint16(6,20,true);w.setUint16(8,2048,true);w.setUint16(14,33,true);w.setUint32(16,crc,true);w.setUint32(20,data.length,true);w.setUint32(24,data.length,true);w.setUint16(28,name.length,true);w.setUint32(42,offset,true);
  parts.push(h,name,data);central.push(c,name);offset+=h.length+name.length+data.length;size+=c.length+name.length;
 }
 const e=new Uint8Array(22),v=new DataView(e.buffer);v.setUint32(0,0x06054b50,true);v.setUint16(8,files.length,true);v.setUint16(10,files.length,true);v.setUint32(12,size,true);v.setUint32(16,offset,true);
 return new Blob([...parts,...central,e],{type:"application/zip"});
}
const xml=x=>String(x??"").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&apos;"}[c]));
function col(i){let s="";for(i++;i;i=Math.floor((i-1)/26))s=String.fromCharCode(65+(i-1)%26)+s;return s;}
export function xlsx(headers,rows){
 const sheet=[headers,...rows].map((row,i)=>'<row r="'+(i+1)+'">'+row.map((v,j)=>'<c r="'+col(j)+(i+1)+'" t="inlineStr"'+(!i?' s="1"':'')+'><is><t xml:space="preserve">'+xml(v)+'</t></is></c>').join("")+'</row>').join("");
 const ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main", rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
 const files=[
 {name:"[Content_Types].xml",data:'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'},
 {name:"_rels/.rels",data:'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+rel+'/officeDocument" Target="xl/workbook.xml"/></Relationships>'},
 {name:"xl/workbook.xml",data:'<workbook xmlns="'+ns+'" xmlns:r="'+rel+'"><sheets><sheet name="Inscriptions" sheetId="1" r:id="rId1"/></sheets></workbook>'},
 {name:"xl/_rels/workbook.xml.rels",data:'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+rel+'/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="'+rel+'/styles" Target="styles.xml"/></Relationships>'},
 {name:"xl/styles.xml",data:'<styleSheet xmlns="'+ns+'"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF102C6B"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'},
 {name:"xl/worksheets/sheet1.xml",data:'<worksheet xmlns="'+ns+'"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>'+headers.map((h,i)=>'<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+Math.min(48,Math.max(18,h.length+3))+'" customWidth="1"/>').join("")+'</cols><sheetData>'+sheet+'</sheetData><autoFilter ref="A1:'+col(headers.length-1)+(rows.length+1)+'"/></worksheet>'}
 ];
 return new Blob([zip(files)],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}
export function saveFile(blob,name){const u=URL.createObjectURL(blob),a=document.createElement("a");a.href=u;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),30000);}
