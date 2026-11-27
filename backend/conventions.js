// V0.9: verified, uncompressed DOCX recipe template. No third-party library.
const enc=new TextEncoder(), dec=new TextDecoder("utf-8",{fatal:true});
const HASH="f96be35b0bfe54a5e1d44bcae6a1d82729266275ec6b5e7c98a8cdf06c3edf64";
const keys=["Etablissement","AdresseEtablissement","TelephoneEtablissement","CiviliteChef","ChefEtablissement","Civilite","Nom","Prenom","DateNaissance","Terminale","Telephone","Email","DateImmersion","IntituleFormation","DateRedaction"];
const xmlEscape=v=>String(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,"").replace(/[\r\n\t]+/g," ").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&apos;"}[c]));
const table=Uint32Array.from({length:256},(_,n)=>{for(let i=0;i<8;i++)n=(n>>>1)^((n&1)?0xedb88320:0);return n>>>0;});
function crc32(data){let c=0xffffffff;for(const byte of data)c=(c>>>8)^table[(c^byte)&255];return (c^0xffffffff)>>>0;}
const buffer=n=>{const a=new Uint8Array(n);return [a,new DataView(a.buffer)];};
function concat(parts){const out=new Uint8Array(parts.reduce((n,p)=>n+p.length,0));let at=0;for(const p of parts){out.set(p,at);at+=p.length;}return out;}
export async function fillStoredDocx(input,values){
 const bytes=new Uint8Array(input);
 const hash=Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256",bytes)),b=>b.toString(16).padStart(2,"0")).join("");
 if(hash!==HASH)throw Error("template_integrity_"+bytes.length+"_"+hash+"_"+(bytes[0]===80&&bytes[1]===75?"zip":"nonzip"));
 const v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),entries=[];
 let p=0,filled=0;
 while(p+30<=bytes.length&&v.getUint32(p,true)===0x04034b50){
   const flags=v.getUint16(p+6,true),method=v.getUint16(p+8,true),crc=v.getUint32(p+14,true),size=v.getUint32(p+18,true),plain=v.getUint32(p+22,true),nl=v.getUint16(p+26,true),xl=v.getUint16(p+28,true);
   if(flags!==0||method!==0||plain!==size||p+30+nl+xl+size>bytes.length)throw Error("template_zip");
   const name=bytes.slice(p+30,p+30+nl),start=p+30+nl+xl;let data=bytes.slice(start,start+size),checksum=crc;
   if(dec.decode(name)==="word/document.xml"){
     const original=dec.decode(data),seen=new Set();
     const xml=original.replace(/\[\[([A-Za-z]+)\]\]/g,(_,k)=>{
       if(!keys.includes(k)||seen.has(k)||typeof values[k]!=="string"||!values[k].trim())throw Error("template_fields");
       seen.add(k);return xmlEscape(values[k]);
     });
     if(seen.size!==keys.length)throw Error("template_fields");
     data=enc.encode(xml);checksum=crc32(data);filled++;
   }
   entries.push({name,data,crc:checksum});p=start+size;
 }
 if(filled!==1||entries.length!==19||v.getUint32(p,true)!==0x02014b50)throw Error("template_zip");
 const local=[],central=[];let offset=0,centralSize=0;
 for(const e of entries){
   const [h,d]=buffer(30);
   d.setUint32(0,0x04034b50,true);d.setUint16(4,20,true);d.setUint16(12,33,true);d.setUint32(14,e.crc,true);d.setUint32(18,e.data.length,true);d.setUint32(22,e.data.length,true);d.setUint16(26,e.name.length,true);
   local.push(h,e.name,e.data);
   const [c,cv]=buffer(46);
   cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);cv.setUint16(14,33,true);cv.setUint32(16,e.crc,true);cv.setUint32(20,e.data.length,true);cv.setUint32(24,e.data.length,true);cv.setUint16(28,e.name.length,true);cv.setUint32(42,offset,true);
   central.push(c,e.name);centralSize+=46+e.name.length;offset+=30+e.name.length+e.data.length;
 }
 const [end,ev]=buffer(22);ev.setUint32(0,0x06054b50,true);ev.setUint16(8,entries.length,true);ev.setUint16(10,entries.length,true);ev.setUint32(12,centralSize,true);ev.setUint32(16,offset,true);
 return concat([...local,...central,end]);
}
function base(env){return "https://graph.microsoft.com/v1.0/sites/"+encodeURIComponent(env.SP_SITE_ID)+"/drive";}
const headers=token=>({Authorization:"Bearer "+token});
async function readBytes(url,token,stage){
 let r;
 try{r=await fetch(url,{headers:{...headers(token),"Accept":"application/octet-stream","Cache-Control":"no-cache"},redirect:"manual"});}catch{throw Error(stage+"_graph_fetch");}
 for(let n=0;[301,302,303,307,308].includes(r.status);n++){
   if(n>=5)throw Error(stage+"_redirect_limit");
   const loc=r.headers.get("Location");let target;
   try{target=new URL(loc,url);}catch{throw Error(stage+"_redirect");}
   if(!loc||target.protocol!=="https:"||target.username||target.password)throw Error(stage+"_redirect");
   // Graph supplies a preauthenticated URL. Never forward its bearer token.
   try{r=await fetch(target.toString(),{headers:{"Accept":"application/octet-stream","Cache-Control":"no-cache"},redirect:"manual"});}catch{throw Error(stage+"_redirect_fetch");}
   url=target.toString();
 }
 if(!r.ok)throw Error(stage+"_http_"+r.status);
 let bytes;try{bytes=await r.arrayBuffer();}catch{throw Error(stage+"_body");}
 if(bytes.byteLength===0||bytes.byteLength>10000000)throw Error(stage+"_size");
 return bytes;
}
const idOK=id=>typeof id==="string"&&/^TEST-V08-[a-f0-9-]{36}$/.test(id);
const name=(id,ext)=>"RECETTE-"+id+"."+ext;
const path=(env,folder,filename)=>base(env)+"/root:/"+folder+"/"+encodeURIComponent(filename)+":/content";
async function upload(env,token,filename,bytes,contentType){
 let r;try{r=await fetch(path(env,"Conventions",filename),{method:"PUT",headers:{...headers(token),"Content-Type":contentType},body:bytes,redirect:"manual"});}catch{throw Error("upload_"+(filename.endsWith(".pdf")?"pdf":"docx")+"_fetch");}
 if(!r.ok)throw Error("upload_http_"+r.status);
 let item;try{item=await r.json();}catch{throw Error("upload_json");}if(typeof item.id!=="string"||!item.id)throw Error("upload_id");return item;
}
const formatDate=d=>{if(typeof d!=="string"||!/^\d{4}-\d{2}-\d{2}$/.test(d))throw Error("date_invalid");return d.split("-").reverse().join("/");};
export async function generateRecipe(env,g,row){
 if(!idOK(row.ReservationID)||row.Voie!=="Voie scolaire"||row.Statut!=="Test"||row.VersionAcceptation!=="RECETTE-V08")throw Error("recipe_only");
 const values=Object.fromEntries(keys.filter(k=>!["DateNaissance","DateImmersion","DateRedaction"].includes(k)).map(k=>[k,row[k]]));
 values.DateNaissance=formatDate(row.DateNaissance);values.DateImmersion=formatDate(row.DateImmersion);
 values.DateRedaction=new Intl.DateTimeFormat("fr-FR",{timeZone:"Europe/Paris",day:"numeric",month:"long",year:"numeric"}).format(new Date());
 const template=await readBytes(path(env,"Modeles","Convention-Initial-Modele-Recette.docx"),g.token,"template");
 let docx;try{docx=await fillStoredDocx(template,values);}catch(e){if(/^template_(integrity_\d+_[a-f0-9]{64}_(zip|nonzip)|zip|fields)$/.test(e?.message||""))throw e;throw Error("template_fill_exception");}
 const doc=await upload(env,g.token,name(row.ReservationID,"docx"),docx,"application/vnd.openxmlformats-officedocument.wordprocessingml.document");
 const pdf=await readBytes(base(env)+"/items/"+encodeURIComponent(doc.id)+"/content?format=pdf",g.token,"conversion");
 if(new TextDecoder().decode(new Uint8Array(pdf,0,Math.min(5,pdf.byteLength)))!=="%PDF-")throw Error("conversion_not_pdf");
 const saved=await upload(env,g.token,name(row.ReservationID,"pdf"),pdf,"application/pdf");
 return {id:saved.id,filename:name(row.ReservationID,"pdf")};
}
export async function downloadRecipe(env,g,row){
 if(!idOK(row.ReservationID)||row.VersionAcceptation!=="RECETTE-V08"||!row.ConventionID||row.StatutConvention!=="Recette generee")throw Error("not_generated");
 const pdf=await readBytes(base(env)+"/items/"+encodeURIComponent(row.ConventionID)+"/content",g.token,"download");
 if(new TextDecoder().decode(new Uint8Array(pdf,0,Math.min(5,pdf.byteLength)))!=="%PDF-")throw Error("download_not_pdf");
 return {pdf,filename:name(row.ReservationID,"pdf")};
}
