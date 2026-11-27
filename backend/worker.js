const CLASS_SEED=[["CG1","CG 1","BTS Comptabilité et Gestion","2026-12-10 2027-01-07 2027-01-14 2027-01-21 2027-01-28 2027-02-04 2027-02-11 2027-02-18 2027-02-25 2027-03-04 2027-03-11"],["SAM1","SAM 1","BTS AMRH (ex-SAM)","2026-12-11 2026-12-18 2027-01-08 2027-01-15 2027-01-22 2027-01-29 2027-02-05 2027-02-12 2027-02-19 2027-02-26 2027-03-05"],["MCO1","MCO 1","BTS MCO","2026-12-15 2027-01-05 2027-01-12 2027-02-02 2027-02-09 2027-02-16 2027-02-23 2027-03-02 2027-03-09"],["NDRC1","NDRC 1","BTS NDRC","2027-01-06 2027-01-13 2027-01-20 2027-02-03 2027-02-10 2027-02-17 2027-02-24 2027-03-03 2027-03-10"],["TOA1","TOA 1","BTS Tourisme","2026-12-18 2027-01-15 2027-01-22 2027-01-29 2027-02-05 2027-02-12 2027-02-19 2027-02-26 2027-03-05"],["TOB1","TOB 1","BTS Tourisme","2026-12-18 2027-01-15 2027-01-22 2027-01-29 2027-02-05 2027-02-12 2027-02-19 2027-02-26 2027-03-05"]];
const LIST_IDS={Formations:"76ea6394-cee3-405d-b67c-92fcfbb0c718",Disponibilites:"b92a4618-af9f-4e19-aae9-d68f8b0730f9",Reservations:"9da020ad-1265-43bd-b611-6e8c209908a4"};
async function businessGraph(env){
 const response=await fetch("https://login.microsoftonline.com/"+encodeURIComponent(env.MS_TENANT_ID)+"/oauth2/v2.0/token",{method:"POST",body:new URLSearchParams({client_id:env.MS_CLIENT_ID,client_secret:env.MS_CLIENT_SECRET,scope:"https://graph.microsoft.com/.default",grant_type:"client_credentials"})});
 if(!response.ok)throw Error("graph_token");
 const token=(await response.json()).access_token;if(!token)throw Error("graph_token");
 const base="https://graph.microsoft.com/v1.0/sites/"+encodeURIComponent(env.SP_SITE_ID)+"/lists/";
 const headers={Authorization:"Bearer "+token,"Content-Type":"application/json"};
 return {token,async read(list,fields){let next=base+LIST_IDS[list]+"/items?$expand=fields($select="+fields.join(",")+")&$top=100";const out=[];while(next){if(new URL(next).origin!=="https://graph.microsoft.com")throw Error("pagination");const r=await fetch(next,{headers});if(!r.ok)throw Error("graph_read");const j=await r.json();for(const item of j.value||[])out.push({...item.fields,_itemId:item.id,_etag:item.eTag});next=j["@odata.nextLink"]; }return out;},async add(list,fields){const r=await fetch(base+LIST_IDS[list]+"/items",{method:"POST",headers,body:JSON.stringify({fields})});if(!r.ok)throw Error("graph_write");await r.json();},async remove(list,item){if(!item._etag||!/^\d+$/.test(item._itemId))throw Error("missing_version");const r=await fetch(base+LIST_IDS[list]+"/items/"+item._itemId,{method:"DELETE",headers:{...headers,"If-Match":item._etag}});if(r.status===412)throw Error("conflict");if(!r.ok)throw Error("graph_delete");},async patch(list,item,fields){if(!item._etag||!/^\d+$/.test(item._itemId))throw Error("missing_version");const r=await fetch(base+LIST_IDS[list]+"/items/"+item._itemId+"/fields",{method:"PATCH",headers:{...headers,"If-Match":item._etag},body:JSON.stringify(fields)});if(r.status===412)throw Error("conflict");if(!r.ok)throw Error("graph_write");}};
}
const classFields=["FormationID","Intitule","Voie","Niveau","Description","Actif","Ordre"];
const slotFields=["DisponibiliteID","FormationID","AnneeID","DateImmersion","Capacite","Ouvert"];
function cutoff(){const d=new Date(new Intl.DateTimeFormat("en-CA",{timeZone:"Europe/Paris",year:"numeric",month:"2-digit",day:"2-digit"}).format(new Date())+"T00:00:00Z");d.setUTCDate(d.getUTCDate()+7);return d.toISOString().slice(0,10);}

const bookingFields=["ReservationID","DisponibiliteID","FormationID","AnneeID","IntituleFormation","Voie","DateImmersion","Nom","Prenom","Email","Telephone","Etablissement","Statut","ConventionID","StatutConvention","StatutMail","DateReservation","Civilite","DateNaissance","VilleResidence","CodePostalResidence","Terminale","SpecialitesBac","AdresseEtablissement","EmailEtablissement","TelephoneEtablissement","CiviliteChef","ChefEtablissement","Acceptation","DateAcceptation","VersionAcceptation","PlaceUnique","DemandeUnique"];
const cancelled=r=>["annule","annulee","annulé","annulée","cancelled"].includes(String(r.Statut||"").toLowerCase());
const isRecipe=r=>r.DisponibiliteID==="RECETTE-V06";
function validateCandidate(b){
 if(!b||typeof b!=="object"||Array.isArray(b))return null;
 const lengths={Nom:100,Prenom:100,Email:254,Telephone:30,Etablissement:200,Civilite:20,DateNaissance:10,VilleResidence:120,CodePostalResidence:5,Terminale:30,SpecialitesBac:200,AdresseEtablissement:400,EmailEtablissement:254,TelephoneEtablissement:30,CiviliteChef:20,ChefEtablissement:200};
 const out={};
 for(const [k,max] of Object.entries(lengths)){if(typeof b[k]!=="string"||!b[k].trim()||b[k].trim().length>max)return null;out[k]=b[k].trim();}
 if(!["Madame","Monsieur"].includes(out.Civilite)||!["Madame","Monsieur"].includes(out.CiviliteChef)||!["Professionnelle","Technologique","Générale"].includes(out.Terminale))return null;
 if(!/^\d{5}$/.test(out.CodePostalResidence)||![out.Email,out.EmailEtablissement].every(e=>/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e))||![out.Telephone,out.TelephoneEtablissement].every(t=>/^[+0-9().\s-]{6,30}$/.test(t)&&t.replace(/\D/g,"").length>=6))return null;
 const d=new Date(out.DateNaissance+"T12:00:00Z");if(!/^\d{4}-\d{2}-\d{2}$/.test(out.DateNaissance)||!Number.isFinite(d.getTime())||d.toISOString().slice(0,10)!==out.DateNaissance||out.DateNaissance<"1900-01-01"||d>new Date())return null;
 if(b.Acceptation!==true)return null;return {...out,Acceptation:true};
}
function remaining(slot,rows){return Math.max(0,Number(slot.Capacite)-rows.filter(r=>r.DisponibiliteID===slot.DisponibiliteID&&!cancelled(r)).length);}

const TENANT="a7f315c0-79f6-4c05-92b6-c9396003e258";
const API_ID="281e44b3-c333-4048-96e8-9765b2e6e30f";
const SPA_ID="1cbae7de-d9fc-4659-a7da-c0623c96e789";
const issuer="https://login.microsoftonline.com/"+TENANT+"/v2.0";
function decode(s){if(!/^[A-Za-z0-9_-]+$/.test(s))throw Error("encoding");return Uint8Array.from(atob(s.replace(/-/g,"+").replace(/_/g,"/").padEnd(Math.ceil(s.length/4)*4,"=")),c=>c.charCodeAt(0));}
function object(s){return JSON.parse(new TextDecoder().decode(decode(s)));}
let keysCache=null, keysUntil=0;
async function keys(force=false){
  if(!force&&keysCache&&Date.now()<keysUntil)return keysCache;
  const response=await fetch("https://login.microsoftonline.com/"+TENANT+"/discovery/v2.0/keys",{redirect:"manual"});
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
const handler = {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = request.headers.get("Origin");
    const headers = new Headers({"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store","X-Content-Type-Options":"nosniff","Vary":"Origin"});
    if(origin === "https://tezenas-immersion.github.io") {
      headers.set("Access-Control-Allow-Origin",origin);
      headers.set("Access-Control-Allow-Methods","GET, POST, OPTIONS");
      headers.set("Access-Control-Allow-Headers","Authorization, Content-Type");
    }
    const json=(data,status=200)=>new Response(JSON.stringify(data),{status,headers});
    if(url.pathname === "/admin/mail-test"){
      if(request.method==="OPTIONS")return new Response(null,{status:204,headers});
      if(request.method!=="POST")return json({error:"method_not_allowed"},405);
      const auth=await authorize(request);if(auth!==200)return json({error:auth===403?"forbidden":"unauthorized"},auth);
      try{
        const raw=await request.text();if(raw.length>1000)return json({error:"invalid_input"},400);
        let b;try{b=JSON.parse(raw);}catch{return json({error:"invalid_input"},400);}
        const to=typeof b?.email==="string"?b.email.trim().toLowerCase():"";
        if(to.length>254||(!/^[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@(groupetezenas\.org|tezenas\.org)$/.test(to)&&to!=="ludovic.lorito@trustteam.fr"))return json({error:"test_address_not_allowed"},400);
        const tokenResponse=await fetch("https://login.microsoftonline.com/"+encodeURIComponent(env.MS_TENANT_ID)+"/oauth2/v2.0/token",{method:"POST",body:new URLSearchParams({client_id:env.MS_CLIENT_ID,client_secret:env.MS_CLIENT_SECRET,scope:"https://graph.microsoft.com/.default",grant_type:"client_credentials"})});
        if(!tokenResponse.ok)return json({error:"graph_token",graphHttp:tokenResponse.status},502);
        const token=(await tokenResponse.json()).access_token;if(!token)return json({error:"graph_token"},502);
        const sender="Tezenas-immersion@tezenas.org";
        const result=await fetch("https://graph.microsoft.com/v1.0/users/"+encodeURIComponent(sender)+"/sendMail",{method:"POST",headers:{Authorization:"Bearer "+token,"Content-Type":"application/json"},body:JSON.stringify({message:{subject:"[RECETTE] Immers’Tez — confirmation de test",body:{contentType:"Text",content:"Bonjour,\n\nCeci est un test technique du portail Immers’Tez. Aucune inscription réelle n’a été créée.\n\nExemple : BTS Tourisme — journée entière, horaires à confirmer.\n\nL’équipe Immers’Tez\nGroupe Tézenas du Montcel"},toRecipients:[{emailAddress:{address:to}}]},saveToSentItems:true})});
        if(result.status!==202)return json({error:result.status===403?"mail_permission_denied":"mail_send_failed",graphHttp:result.status},502);
        return json({status:"accepted",graphHttp:202,sender,recipient:to,test:true,reservationCreated:false},202);
      }catch{return json({error:"mail_test_failed"},502);}
    }
    if(url.pathname === "/public/planning") {
      if(request.method==="OPTIONS")return new Response(null,{status:204,headers});
      if(request.method!=="GET")return json({error:"method_not_allowed"},405);
      try {
        const g=await businessGraph(env);
        const [classes,slots,rows]=await Promise.all([g.read("Formations",classFields),g.read("Disponibilites",slotFields),g.read("Reservations",["DisponibiliteID","Statut"])]);
        const limit=cutoff();
        const active=classes.filter(c=>c.Actif===true&&c.Voie==="Voie scolaire");
        const allowed=new Set(active.map(c=>c.FormationID));
        return json({status:"ok",reservations:false,classes:active.map(c=>({id:c.FormationID,title:c.Intitule,label:c.Niveau,order:c.Ordre})),slots:slots.filter(s=>allowed.has(s.FormationID)&&s.Ouvert===true&&/^\d{4}-\d{2}-\d{2}$/.test(s.DateImmersion)&&s.DateImmersion>limit&&Number.isInteger(s.Capacite)&&s.Capacite>0).map(s=>({classId:s.FormationID,date:s.DateImmersion,capacity:s.Capacite,remaining:remaining(s,rows)}))});
      }catch{return json({error:"planning_unavailable"},502);}
    }
    if(url.pathname === "/admin/booking-create"){
      if(request.method==="OPTIONS")return new Response(null,{status:204,headers});
      if(request.method!=="POST")return json({error:"method_not_allowed"},405);
      const auth=await authorize(request);if(auth!==200)return json({error:auth===403?"forbidden":"unauthorized"},auth);
      try{
        const raw=await request.text();if(raw.length>12000)return json({error:"invalid_input"},400);
        let b;try{b=JSON.parse(raw);}catch{return json({error:"invalid_input"},400);}
        if(!b||typeof b.requestId!=="string"||!/^v08-[a-f0-9]{32}$/.test(b.requestId)||typeof b.title!=="string"||!["Voie scolaire","Apprentissage"].includes(b.way)||typeof b.date!=="string"||!/^\d{4}-\d{2}-\d{2}$/.test(b.date))return json({error:"invalid_input"},400);
        const email=typeof b.email==="string"?b.email.trim().toLowerCase():"";
        if(email.length>254||(!/^[a-z0-9.!#$%&'*+/=?^_+-]+@(groupetezenas\.org|tezenas\.org)$/.test(email)&&email!=="ludovic.lorito@trustteam.fr"))return json({error:"test_address_not_allowed"},400);
        const candidate=validateCandidate({Nom:"RECETTE V08",Prenom:"Candidat fictif",Email:email,Telephone:"0600000000",Etablissement:"LYCEE FICTIF",Civilite:"Madame",DateNaissance:"2009-01-01",VilleResidence:"VILLE FICTIVE",CodePostalResidence:"42000",Terminale:"Générale",SpecialitesBac:"Test",AdresseEtablissement:"Adresse fictive",EmailEtablissement:"lycee@example.invalid",TelephoneEtablissement:"0400000000",CiviliteChef:"Monsieur",ChefEtablissement:"CHEF FICTIF",Acceptation:true});
        if(!candidate)return json({error:"invalid_input"},400);
        const g=await businessGraph(env);
        const rows=await g.read("Reservations",bookingFields);
        const prior=rows.find(r=>r.DemandeUnique===b.requestId);
        if(prior)return json({status:"ok",saved:true,replayed:true,id:prior.ReservationID,classe:prior.FormationID,mail:prior.StatutMail,statut:prior.Statut});
        if(b.date<=cutoff())return json({error:"deadline"},409);
        const [classes,slots]=await Promise.all([g.read("Formations",classFields),g.read("Disponibilites",slotFields)]);
        const eligible=classes.filter(c=>c.Actif===true&&c.Voie===b.way&&c.Intitule===b.title);
        const choices=slots.filter(s=>s.Ouvert===true&&s.DateImmersion===b.date&&eligible.some(c=>c.FormationID===s.FormationID)&&Number.isInteger(s.Capacite)&&s.Capacite>0&&s.Capacite<=30&&remaining(s,rows)>0).sort((a,b)=>remaining(b,rows)-remaining(a,rows)||String(a.FormationID).localeCompare(String(b.FormationID)));
        if(!choices.length)return json({error:"full_or_closed"},409);
        const used=new Set(rows.map(r=>r.PlaceUnique));
        for(const slot of choices){
          if(slots.filter(x=>x.FormationID===slot.FormationID&&x.DateImmersion===slot.DateImmersion).length!==1)return json({error:"duplicate_slot"},409);
          const cls=eligible.find(c=>c.FormationID===slot.FormationID);
          for(let i=1;i<=slot.Capacite;i++){
            const place=slot.DisponibiliteID+"#"+i;if(used.has(place))continue;
            const id="TEST-V08-"+crypto.randomUUID(),now=new Date().toISOString();
            try{await g.add("Reservations",{...candidate,Title:id,ReservationID:id,DisponibiliteID:slot.DisponibiliteID,FormationID:cls.FormationID,AnneeID:slot.AnneeID,IntituleFormation:cls.Intitule,Voie:cls.Voie,DateImmersion:slot.DateImmersion,Statut:"Test",StatutConvention:"En attente",StatutMail:"En attente",DateReservation:now,DateAcceptation:now,VersionAcceptation:"RECETTE-V08",PlaceUnique:place,DemandeUnique:b.requestId});}
            catch{return json({error:"write_uncertain",requestId:b.requestId},502);}
            let mail="Incertain";
            try{
              const graph=await fetch("https://graph.microsoft.com/v1.0/users/"+encodeURIComponent("Tezenas-immersion@tezenas.org")+"/sendMail",{method:"POST",headers:{Authorization:"Bearer "+g.token,"Content-Type":"application/json"},body:JSON.stringify({message:{subject:"[RECETTE] Immers’Tez — confirmation de réservation",body:{contentType:"Text",content:"Bonjour,\n\nVotre réservation FICTIVE a été enregistrée pour tester Immers’Tez.\nFormation : "+cls.Intitule+"\nDate : "+slot.DateImmersion+"\nRéférence : "+id+"\n\nAucune inscription réelle ni convention n’est émise.\n\nL’équipe Immers’Tez"},toRecipients:[{emailAddress:{address:email}}]},saveToSentItems:true})});
              mail=graph.status===202?"Accepte Graph":graph.status>=500?"Incertain":"Echec";
            }catch{}
            try{const saved=(await g.read("Reservations",["ReservationID","StatutMail"])).find(x=>x.ReservationID===id);if(saved)await g.patch("Reservations",saved,{StatutMail:mail});else mail="Incertain";}catch{mail="Incertain";}
            return json({status:"ok",saved:true,id,classe:cls.Niveau,mail,remaining:remaining(slot,rows)-1,test:true},201);
          }
        }
        return json({error:"full_or_closed"},409);
      }catch{return json({error:"operation_failed"},502);}
    }
    if(url.pathname === "/admin/reservations" || url.pathname === "/admin/reservation-change" || url.pathname === "/admin/booking-preview"){
      if(request.method==="OPTIONS")return new Response(null,{status:204,headers});
      const auth=await authorize(request);if(auth!==200)return json({error:auth===403?"forbidden":"unauthorized"},auth);
      try{
        const g=await businessGraph(env);
        if(url.pathname==="/admin/reservations"){
          if(request.method!=="GET")return json({error:"method_not_allowed"},405);
          const rows=await g.read("Reservations",bookingFields);
          return json({status:"ok",rows:rows.filter(r=>!isRecipe(r)).map(r=>({id:r.ReservationID,nom:r.Nom,prenom:r.Prenom,formation:r.IntituleFormation,classe:r.FormationID,date:r.DateImmersion,statut:r.Statut,email:r.Email,convention:r.StatutConvention,mail:r.StatutMail,etag:r._etag})),recipeCount:rows.filter(isRecipe).length});
        }
        if(request.method!=="POST")return json({error:"method_not_allowed"},405);
        const raw=await request.text();if(raw.length>12000)return json({error:"invalid_input"},400);
        let b;try{b=JSON.parse(raw);}catch{return json({error:"invalid_input"},400);}
        if(!b||typeof b!=="object")return json({error:"invalid_input"},400);
        if(url.pathname==="/admin/reservation-change"){
          const rows=await g.read("Reservations",["ReservationID","DisponibiliteID","FormationID","Statut","PlaceUnique","DemandeUnique"]);
          if(b.action==="cancel"){
            const row=rows.find(r=>r.ReservationID===b.id&&!isRecipe(r));if(!row)return json({error:"not_found"},404);
            if(cancelled(row))return json({status:"ok",already:true});
            if(row._etag!==b.etag)return json({error:"conflict"},409);
            await g.patch("Reservations",row,{Statut:"Annulee",PlaceUnique:"ANNULEE-"+row.ReservationID});
            return json({status:"ok"});
          }
          if(b.action==="cleanup-v06"){
            const tests=rows.filter(r=>isRecipe(r)&&r.FormationID==="RECETTE"&&r.Statut==="Test"&&/^TEST-[0-9a-f-]{36}$/.test(r.ReservationID)&&/^RECETTE-V06-[12]$/.test(r.PlaceUnique)&&/^test-[a-f0-9]{32}$/.test(r.DemandeUnique));
            let removed=0;for(const row of tests){await g.remove("Reservations",row);removed++;}
            return json({status:"ok",removed});
          }
          return json({error:"invalid_input"},400);
        }
        const candidate=validateCandidate(b.candidate);
        if(!candidate||typeof b.title!=="string"||!["Voie scolaire","Apprentissage"].includes(b.way)||typeof b.date!=="string"||!/^\d{4}-\d{2}-\d{2}$/.test(b.date))return json({error:"invalid_input"},400);
        if(b.date<=cutoff())return json({error:"deadline"},409);
        const [classes,slots,rows]=await Promise.all([g.read("Formations",classFields),g.read("Disponibilites",slotFields),g.read("Reservations",["DisponibiliteID","Statut"])]);
        const eligible=classes.filter(c=>c.Actif===true&&c.Voie===b.way&&c.Intitule===b.title);
        const choices=slots.filter(s=>s.Ouvert===true&&s.DateImmersion===b.date&&eligible.some(c=>c.FormationID===s.FormationID)&&Number.isInteger(s.Capacite)&&remaining(s,rows)>0).sort((a,b)=>remaining(b,rows)-remaining(a,rows)||String(a.FormationID).localeCompare(String(b.FormationID)));
        if(!choices.length)return json({error:"full_or_closed"},409);
        const slot=choices[0],cls=eligible.find(c=>c.FormationID===slot.FormationID);
        return json({status:"ok",preview:true,saved:false,classe:cls.Niveau,formation:cls.Intitule,date:slot.DateImmersion,remaining:remaining(slot,rows)});
      }catch(e){return json({error:e?.message==="conflict"?"conflict":"operation_failed"},e?.message==="conflict"?409:502);}
    }
    if(url.pathname === "/admin/change") {
      if(request.method==="OPTIONS")return new Response(null,{status:204,headers});
      if(request.method!=="POST")return json({error:"method_not_allowed"},405);
      const status=await authorize(request);if(status!==200)return json({error:status===403?"forbidden":"unauthorized"},status);
      try{
        const raw=await request.text();if(raw.length>12000)return json({error:"invalid_input"},400);
        let b;try{b=JSON.parse(raw);}catch{return json({error:"invalid_input"},400);}
        if(!b||typeof b!=="object")return json({error:"invalid_input"},400);
        const g=await businessGraph(env);
        const classes=await g.read("Formations",classFields);
        const cls=classes.find(c=>c.FormationID===b.id);
        if(b.action==="class"){
          const valid=x=>typeof x==="string"&&x.trim().length>0&&x.trim().length<=150;
          if(!valid(b.label)||!valid(b.title)||!["Voie scolaire","Apprentissage"].includes(b.way)||typeof b.active!=="boolean")return json({error:"invalid_input"},400);
          const fields={Title:b.label.trim(),Niveau:b.label.trim(),Intitule:b.title.trim(),Voie:b.way,Actif:b.active};
          if(b.id){if(!cls)return json({error:"not_found"},404);if(b.etag!==cls._etag)return json({error:"conflict"},409);if(b.way!==cls.Voie)return json({error:"way_immutable"},400);await g.patch("Formations",cls,fields);}
          else{if(classes.some(c=>c.Voie===b.way&&c.Niveau.toLowerCase()===fields.Niveau.toLowerCase()))return json({error:"duplicate_class"},409);await g.add("Formations",{...fields,FormationID:crypto.randomUUID(),Ordre:classes.length+1});}
          return json({status:"ok"});
        }
        if(b.action!=="day"||!cls||!Number.isInteger(b.capacity)||b.capacity<1||b.capacity>30||typeof b.open!=="boolean"||typeof b.date!=="string"||!/^\d{4}-\d{2}-\d{2}$/.test(b.date))return json({error:"invalid_input"},400);
        const date=new Date(b.date+"T12:00:00Z");if(!Number.isFinite(date.getTime())||date.toISOString().slice(0,10)!==b.date)return json({error:"invalid_input"},400);
        const today=new Intl.DateTimeFormat("en-CA",{timeZone:"Europe/Paris",year:"numeric",month:"2-digit",day:"2-digit"}).format(new Date());
        if(b.date<today)return json({error:"past_date"},400);
        const y=date.getUTCFullYear()-(date.getUTCMonth()<8?1:0),year=y+"-"+(y+1);
        const slots=await g.read("Disponibilites",slotFields);
        const matches=slots.filter(s=>s.FormationID===b.id&&s.DateImmersion===b.date);
        if(matches.length>1)return json({error:"duplicate_slot"},409);
        const slot=matches[0];
        if(slot){
          if(b.etag!==slot._etag)return json({error:"conflict"},409);
          const reservations=await g.read("Reservations",["DisponibiliteID","Statut"]);
          const count=reservations.filter(r=>r.DisponibiliteID===slot.DisponibiliteID&&!["annule","annulee","annulé","annulée","cancelled"].includes(String(r.Statut||"").toLowerCase())).length;
          if(b.capacity<count)return json({error:"capacity_below_bookings"},409);
          await g.patch("Disponibilites",slot,{Capacite:b.capacity,Ouvert:b.open});
        }else{
          if(!b.open||b.etag)return json({error:"conflict"},409);
          await g.add("Disponibilites",{Title:cls.Niveau+" "+b.date,DisponibiliteID:year+"-"+b.id+"-"+b.date,FormationID:b.id,AnneeID:year,DateImmersion:b.date,Capacite:b.capacity,Ouvert:true});
        }
        return json({status:"ok"});
      }catch(e){return json({error:e?.message==="conflict"?"conflict":"save_failed"},e?.message==="conflict"?409:502);}
    }
    if(url.pathname === "/admin/planning") {
      if(request.method==="OPTIONS")return new Response(null,{status:204,headers});
      if(request.method!=="GET")return json({error:"method_not_allowed"},405);
      const status=await authorize(request);if(status!==200)return json({error:status===403?"forbidden":"unauthorized"},status);
      try{const g=await businessGraph(env);const [classes,slots]=await Promise.all([g.read("Formations",classFields),g.read("Disponibilites",slotFields)]);return json({status:"ok",classes,slots,reservations:false});}catch{return json({error:"planning_unavailable"},502);}
    }
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
        if(url.searchParams.has("importClass")) {
          const code=url.searchParams.get("importClass");
          const entry=CLASS_SEED.find(c=>c[0]===code);if(!entry)return json({error:"unknown_class"},400);
          const [id,label,title,dates]=entry;const classId="INITIAL-"+id;
          const g=await businessGraph(env);
          const [classes,slots]=await Promise.all([g.read("Formations",["FormationID"]),g.read("Disponibilites",["DisponibiliteID"])]);
          let created=0,existing=0;
          if(!classes.some(c=>c.FormationID===classId))await g.add("Formations",{Title:label,FormationID:classId,Intitule:title,Voie:"Voie scolaire",Niveau:label,Actif:true,Ordre:CLASS_SEED.indexOf(entry)+1});
          const limit=cutoff();
          for(const date of dates.split(" ").filter(d=>d>limit)){
            const slotId="2026-2027-"+classId+"-"+date;
            if(slots.some(s=>s.DisponibiliteID===slotId)){existing++;continue;}
            await g.add("Disponibilites",{Title:label+" "+date,DisponibiliteID:slotId,FormationID:classId,AnneeID:"2026-2027",DateImmersion:date,Capacite:2,Ouvert:true});created++;
          }
          return json({status:"ok",classe:label,created,existing,reservations:false});
        }
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
    if(url.pathname === "/" || url.pathname === "/health") return json({service:"tezenas-immersions-api",version:"0.8.0",status:"ok",mode:"demo",reservations:false});
    return json({error:"not_found"},404);
  }
};

export class MutationGate {
  constructor(state,env){this.env=env;this.tail=Promise.resolve();}
  async fetch(request){
    const job=this.tail.then(()=>handler.fetch(request,this.env));
    this.tail=job.then(()=>{},()=>{});
    return job;
  }
}
export default {
  async fetch(request,env){
    const url=new URL(request.url);
    const write=request.method==="POST"&&["/admin/change","/admin/reservation-change","/admin/booking-create"].includes(url.pathname);
    if(request.method==="POST"&&url.pathname==="/diagnostic/graph"&&(url.searchParams.has("importClass")||url.searchParams.has("init")))return new Response(JSON.stringify({error:"bootstrap_disabled"}),{status:403,headers:{"Content-Type":"application/json"}});
    if(write){
      if(!env.MUTATION_GATE)return new Response(JSON.stringify({error:"mutation_gate_unavailable"}),{status:503,headers:{"Content-Type":"application/json","Access-Control-Allow-Origin":"https://tezenas-immersion.github.io"}});
      return env.MUTATION_GATE.get(env.MUTATION_GATE.idFromName("immersions-global")).fetch(request);
    }
    return handler.fetch(request,env);
  }
};
