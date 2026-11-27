const TENANT="a7f315c0-79f6-4c05-92b6-c9396003e258";
const API_ID="281e44b3-c333-4048-96e8-9765b2e6e30f";
const SPA_ID="1cbae7de-d9fc-4659-a7da-c0623c96e789";
const issuer="https://login.microsoftonline.com/"+TENANT+"/v2.0";
function decode(s){if(!/^[A-Za-z0-9_-]+$/.test(s))throw Error("encoding");return Uint8Array.from(atob(s.replace(/-/g,"+").replace(/_/g,"/").padEnd(Math.ceil(s.length/4)*4,"=")),c=>c.charCodeAt(0));}
function object(s){return JSON.parse(new TextDecoder().decode(decode(s)));}
let keysCache=null, keysUntil=0;
async function keys(force=false){
  if(!force&&keysCache&&Date.now()<keysUntil)return keysCache;
  const response=await fetch("https://login.microsoftonline.com/"+TENANT+"/discovery/v2.0/keys",{redirect:"error"});
  if(!response.ok)throw Error("jwks_http_"+response.status);
  const data=await response.json();if(!Array.isArray(data.keys))throw Error("keys");
  keysCache=data.keys;keysUntil=Date.now()+3600000;return keysCache;
}
async function authorize(request){
  const deny=reason=>{console.warn("admin_auth_rejected",reason);return 401;};
  const value=request.headers.get("Authorization")||"";
  if(!value.startsWith("Bearer ")||value.length>20000)return deny("bearer");
  let stage="decode";
  try{
    const parts=value.slice(7).split(".");if(parts.length!==3)return deny("format");
    const h=object(parts[0]), p=object(parts[1]);
    if(h.alg!=="RS256"||typeof h.kid!=="string")return deny("algorithm");
    const now=Math.floor(Date.now()/1000);
    if(p.ver!=="2.0")return deny("token_version");
    if(p.iss!==issuer)return deny("issuer");
    if(p.tid!==TENANT)return deny("tenant");
    if(p.aud!==API_ID)return deny("audience");
    if(p.azp!==SPA_ID)return deny("client");
    if(typeof p.oid!=="string")return deny("identity");
    if(!Number.isFinite(p.exp)||!Number.isFinite(p.nbf)||p.exp<=now||p.nbf>now+60)return deny("validity");
    const match=k=>k.kid===h.kid&&k.kty==="RSA"&&(!k.use||k.use==="sig");
    stage="jwks_fetch";
    let jwk=(await keys()).find(match);if(!jwk)jwk=(await keys(true)).find(match);if(!jwk)return deny("signing_key");
    stage="jwks_issuer";
    if(jwk.issuer&&jwk.issuer.replace("{tenantid}",TENANT)!==issuer)return deny("key_issuer");
    stage="key_import";
    const key=await crypto.subtle.importKey("jwk",{kty:"RSA",n:jwk.n,e:jwk.e,alg:"RS256",ext:true},{name:"RSASSA-PKCS1-v1_5",hash:"SHA-256"},false,["verify"]);
    stage="signature_verify";
    if(!await crypto.subtle.verify("RSASSA-PKCS1-v1_5",key,decode(parts[2]),new TextEncoder().encode(parts[0]+"."+parts[1])))return deny("signature");
    if(typeof p.scp!=="string"||!p.scp.split(" ").includes("access_as_user")||!Array.isArray(p.roles)||!p.roles.includes("Immersions.Admin"))return 403;
    return 200;
  }catch(e){
    const type=["TypeError","DataError","OperationError","NotSupportedError","InvalidAccessError","SyntaxError","Error"].includes(e?.name)?e.name:"UnknownError";
    const code=/^jwks_http_[0-9]{3}$/.test(e?.message||"")?e.message:"";
    return deny("verification_exception_"+stage+"_"+type+(code?"_"+code:""));
  }
}
export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = request.headers.get("Origin");
    const headers = new Headers({"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store","X-Content-Type-Options":"nosniff","Vary":"Origin"});
    if(origin === "https://tezenas-immersion.github.io") {
      headers.set("Access-Control-Allow-Origin",origin);
      headers.set("Access-Control-Allow-Methods","GET, OPTIONS");
      headers.set("Access-Control-Allow-Headers","Authorization");
    }
    const json=(data,status=200)=>new Response(JSON.stringify(data),{status,headers});
    if(url.pathname === "/admin/session") {
      if(request.method==="OPTIONS")return new Response(null,{status:204,headers});
      if(request.method!=="GET")return json({error:"method_not_allowed"},405);
      const status=await authorize(request);
      if(status!==200)return json({error:status===403?"forbidden":"unauthorized"},status);
      return json({status:"ok",role:"Immersions.Admin",mode:"auth-test"});
    }
    if(url.pathname === "/diagnostic/graph") {
      if(request.method !== "POST") return json({error:"method_not_allowed"},405);
      const supplied=request.headers.get("X-Diagnostic-Key") || "";
      if(!env.DIAGNOSTIC_KEY || supplied.length !== env.DIAGNOSTIC_KEY.length || !crypto.subtle.timingSafeEqual(new TextEncoder().encode(supplied),new TextEncoder().encode(env.DIAGNOSTIC_KEY))) return json({error:"unauthorized"},401);
      let step="configuration";
      try {
        if(!env.MS_TENANT_ID || !env.MS_CLIENT_ID || !env.MS_CLIENT_SECRET || !env.SP_SITE_ID) return json({status:"error",step},503);
        step="entra";
        const tokenResponse=await fetch("https://login.microsoftonline.com/"+encodeURIComponent(env.MS_TENANT_ID)+"/oauth2/v2.0/token",{method:"POST",body:new URLSearchParams({client_id:env.MS_CLIENT_ID,client_secret:env.MS_CLIENT_SECRET,scope:"https://graph.microsoft.com/.default",grant_type:"client_credentials"})});
        if(!tokenResponse.ok) return json({status:"error",step,http:tokenResponse.status},502);
        const token=await tokenResponse.json();
        if(!token.access_token) return json({status:"error",step},502);
        const graph="https://graph.microsoft.com/v1.0/sites/"+encodeURIComponent(env.SP_SITE_ID);
        const options={headers:{Authorization:"Bearer "+token.access_token}};
        step="sharepoint";
        const site=await fetch(graph+"?$select=id",options);
        if(!site.ok) return json({status:"error",step,http:site.status},502);
        await site.json();
        step="lists";
        const lists=await fetch(graph+"/lists?$select=id&$top=1",options);
        if(!lists.ok) return json({status:"error",step,http:lists.status},502);
        await lists.json();
        if(url.searchParams.get("init") === "1") {
          // Explicit, protected bootstrap. Run once, without concurrent calls.
          const ids={Parametres:"f3201118-d480-419d-8536-46fbaaac7577",AnneesScolaires:"12005aa7-5c1c-4d57-9b45-da37876baf46"};
          const now=new Date();
          const parts=new Intl.DateTimeFormat("fr-FR",{timeZone:"Europe/Paris",year:"numeric",month:"numeric"}).formatToParts(now);
          const year=Number(parts.find(p=>p.type==="year").value)-(Number(parts.find(p=>p.type==="month").value)<9?1:0);
          const yearId=year+"-"+(year+1);
          const seeds=[
            ["Parametres","Cle",[
              {Title:"TitreAccueil",Cle:"TitreAccueil",Valeur:"Découvrez votre future formation",Description:"Titre public provisoire"},
              {Title:"TexteAccueil",Cle:"TexteAccueil",Valeur:"Découvrez nos formations lors d'une journée d'immersion.",Description:"Présentation provisoire"},
              {Title:"CapaciteDefaut",Cle:"CapaciteDefaut",Valeur:"2",Description:"Places par journée et formation (hypothèse à confirmer)"},
              {Title:"DelaiMinimumJours",Cle:"DelaiMinimumJours",Valeur:"7",Description:"Délai minimum de réservation"},
              {Title:"ReservationsActives",Cle:"ReservationsActives",Valeur:"false",Description:"Réservations réelles désactivées"},
              {Title:"AnneeAutomatique",Cle:"AnneeAutomatique",Valeur:"true",Description:"Basculement au 1er septembre"}
            ]],
            ["AnneesScolaires","AnneeID",[{Title:yearId,AnneeID:yearId,DateDebut:year+"-09-01",DateFin:(year+1)+"-08-31",Actif:true}]]
          ];
          const results=[];
          for(const [name,key,rows] of seeds){
            const base=graph+"/lists/"+ids[name]+"/items";
            step="init_read_"+name;
            let next=base+"?$expand=fields($select="+key+")&$top=100";
            const keys=new Set();
            while(next){
              if(new URL(next).origin!=="https://graph.microsoft.com")throw Error("invalid pagination");
              const response=await fetch(next,options);
              if(!response.ok)return json({status:"error",step,http:response.status,results},502);
              const data=await response.json();
              for(const item of data.value||[])keys.add(item.fields?.[key]);
              next=data["@odata.nextLink"];
            }
            const result={list:name,created:0,existing:0};results.push(result);
            for(const fields of rows){
              if(keys.has(fields[key])){result.existing++;continue;}
              step="init_write_"+name;
              const response=await fetch(base,{method:"POST",headers:{...options.headers,"Content-Type":"application/json"},body:JSON.stringify({fields})});
              if(!response.ok)return json({status:"error",step,http:response.status,results},502);
              await response.json();keys.add(fields[key]);result.created++;
            }
          }
          return json({status:"ok",results,reservations:false});
        }
        if(url.searchParams.get("schema") === "1") {
          const expected={Parametres:["Cle","Valeur","Description"],AnneesScolaires:["AnneeID","DateDebut","DateFin","Actif"],Formations:["FormationID","Intitule","Voie","Niveau","Description","Actif","Ordre"],Disponibilites:["DisponibiliteID","FormationID","AnneeID","DateImmersion","HeureDebut","HeureFin","Capacite","Ouvert"],Reservations:["ReservationID","DisponibiliteID","FormationID","AnneeID","IntituleFormation","Voie","DateImmersion","Nom","Prenom","Email","Telephone","Etablissement","Statut","ConventionID","StatutConvention","StatutMail","DateReservation"],JournalActivite:["Action","Auteur","Cible","DateAction","Resultat"]};
          const found={};
          let next=graph+"/lists?$select=id,displayName&$top=100";
          while(next) {
            step="schema_discovery";
            const page=await fetch(next,options);
            if(!page.ok) return json({status:"error",step,http:page.status},502);
            const data=await page.json();
            for(const list of data.value||[]) if(Object.hasOwn(expected,list.displayName)) found[list.displayName]=list.id;
            next=data["@odata.nextLink"];
            if(next && new URL(next).origin!=="https://graph.microsoft.com") throw Error("invalid pagination");
          }
          const results=await Promise.all(Object.entries(expected).map(async ([name,fields])=>{
            if(!found[name])return {name,ok:false,missing:fields};
            const names=new Set();
            let next=graph+"/lists/"+encodeURIComponent(found[name])+"/columns?$select=name&$top=100";
            while(next){
              const page=await fetch(next,options);
              if(!page.ok)return {name,ok:false,http:page.status};
              const data=await page.json();
              for(const col of data.value||[])names.add(col.name);
              next=data["@odata.nextLink"];
              if(next && new URL(next).origin!=="https://graph.microsoft.com")throw Error("invalid pagination");
            }
            const missing=fields.filter(f=>!names.has(f));
            return {name,ok:missing.length===0,expected:fields.length,missing};
          }));
          return json({status:results.every(r=>r.ok)?"ok":"error",schema:results},results.every(r=>r.ok)?200:502);
        }
        if(url.searchParams.get("write") !== "1") return json({status:"ok",entra:true,sharepoint:true,lists:true,write_test:false});
        step="write";
        const name="diagnostic-"+crypto.randomUUID()+".txt";
        const drive=graph+"/drive";
        const file=await fetch(drive+"/root:/Modeles/"+name+":/content",{method:"PUT",headers:{...options.headers,"Content-Type":"text/plain"},body:"Test technique temporaire, sans donnees personnelles."});
        if(!file.ok) return json({status:"error",step,http:file.status},502);
        const item=await file.json();
        if(!item.id) return json({status:"error",step:"cleanup",write_test:true},502);
        step="cleanup";
        const deleted=await fetch(drive+"/items/"+encodeURIComponent(item.id),{method:"DELETE",headers:options.headers});
        if(!deleted.ok) return json({status:"error",step,http:deleted.status,write_test:true,cleanup:false},502);
        return json({status:"ok",entra:true,sharepoint:true,lists:true,write_test:true,cleanup:true});
      } catch {return json({status:"error",step},502);}
    }
    if(request.method === "OPTIONS") return new Response(null,{status:204,headers});
    if(request.method !== "GET") return json({error:"method_not_allowed"},405);
    if(url.pathname === "/" || url.pathname === "/health") return json({service:"tezenas-immersions-api",version:"0.2.2",status:"ok",mode:"demo",reservations:false});
    return json({error:"not_found"},404);
  }
};
