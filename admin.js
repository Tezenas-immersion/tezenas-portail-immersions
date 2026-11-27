"use strict";
const tenant="a7f315c0-79f6-4c05-92b6-c9396003e258", client="1cbae7de-d9fc-4659-a7da-c0623c96e789", scope="api://281e44b3-c333-4048-96e8-9765b2e6e30f/access_as_user", redirect="https://tezenas-immersion.github.io/tezenas-portail-immersions/admin.html", authority="https://login.microsoftonline.com/"+tenant+"/oauth2/v2.0", api="https://tezenas-immersions-api.tezenas-immersion.workers.dev/admin/session";
const $=id=>document.getElementById(id), status=t=>{$("status").textContent=t;}, b64=bytes=>btoa(String.fromCharCode(...bytes)).replace(/=/g,"").replace(/\+/g,"-").replace(/\//g,"_"), random=()=>b64(crypto.getRandomValues(new Uint8Array(32)));
const storage="gtm.pkce.v1";
function reset(){sessionStorage.removeItem(storage);$("dashboard").hidden=true;$("logout").hidden=true;$("login").hidden=false;$("login").disabled=false;}
$("login").onclick=async()=>{
  $("login").disabled=true;
  try{
    const verifier=random(), state=random(), challenge=b64(new Uint8Array(await crypto.subtle.digest("SHA-256",new TextEncoder().encode(verifier))));
    sessionStorage.setItem(storage,JSON.stringify({verifier,state,created:Date.now()}));
    location.assign(authority+"/authorize?"+new URLSearchParams({client_id:client,response_type:"code",redirect_uri:redirect,response_mode:"query",scope,state,code_challenge:challenge,code_challenge_method:"S256",prompt:"select_account"}));
  }catch{reset();status("Impossible de démarrer la connexion.");}
};
$("logout").onclick=()=>{reset();location.assign(authority+"/logout?"+new URLSearchParams({post_logout_redirect_uri:redirect}));};
(async()=>{
  const params=new URLSearchParams(location.search);
  if(!params.has("code")&&!params.has("error"))return;
  history.replaceState(null,"",location.pathname);
  $("login").disabled=true;
  try{
    const raw=sessionStorage.getItem(storage);sessionStorage.removeItem(storage);
    if(!raw)throw Error("state");
    const pending=JSON.parse(raw);
    if(params.get("state")!==pending.state||Date.now()-pending.created>900000)throw Error("state");
    if(params.has("error")){reset();status("Connexion Microsoft refusée ou annulée ("+(params.get("error")||"erreur")+").");return;}
    status("Vérification de votre identité et de vos droits…");
    const response=await fetch(authority+"/token",{method:"POST",body:new URLSearchParams({client_id:client,grant_type:"authorization_code",code:params.get("code"),redirect_uri:redirect,code_verifier:pending.verifier,scope})});
    if(!response.ok){reset();status("Microsoft n'a pas délivré de jeton ("+response.status+"). Vérifiez le consentement et l'affectation.");return;}
    const data=await response.json();
    if(typeof data.access_token!=="string")throw Error("token");
    const result=await fetch(api,{headers:{Authorization:"Bearer "+data.access_token},cache:"no-store"});
    if(!result.ok){reset();status(result.status===403?"Accès refusé : rôle Administrateur Immersions absent.":result.status===401?"Accès refusé : jeton non valide pour cette API.":"Service indisponible ("+result.status+").");$("logout").hidden=false;return;}
    const verified=await result.json();
    if(verified.status!=="ok"||verified.role!=="Immersions.Admin")throw Error("unexpected response");
    $("dashboard").hidden=false;$("login").hidden=true;$("logout").hidden=false;status("Authentification Microsoft 365 réussie. Rôle Immersions.Admin confirmé côté serveur.");
    await showPlanning(data.access_token);
  }catch{reset();status("La connexion n'a pas pu être validée. Relancez la connexion.");}
})();


async function showPlanning(token){
 const panel=$("dashboard");panel.replaceChildren();
 const el=(tag,t="")=>{const n=document.createElement(tag);n.textContent=t;return n;};
 const btn=(t,fn)=>{const b=el("button",t);b.type="button";b.onclick=fn;return b;};
 const field=(label,input)=>{const l=el("label",label);l.append(input);return l;};
 const inp=(type,value)=>{const n=el("input");n.type=type;n.value=value;return n;};
 const select=el("select"), month=el("h3"), content=el("div"), editor=el("div"), feedback=el("p"), summary=el("p"), nav=el("div");
 const now=new Date();let view=new Date(now.getFullYear(),now.getMonth(),1), data={classes:[],slots:[]}, mode="calendar", busy=false;
 const iso=(d)=>[d.getFullYear(),String(d.getMonth()+1).padStart(2,"0"),String(d.getDate()).padStart(2,"0")].join("-");
 const pretty=d=>new Date(d+"T12:00:00Z").toLocaleDateString("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric",timeZone:"UTC"});
 const endpoint=api.replace("/admin/session","");
 async function call(path,body){
   const response=await fetch(endpoint+path,{method:body?"POST":"GET",headers:{Authorization:"Bearer "+token,...(body?{"Content-Type":"application/json"}:{})},...(body?{body:JSON.stringify(body)}:{}),cache:"no-store"});
   const d=await response.json();if(!response.ok)throw Error(response.status===401||response.status===403?"Session expirée ou accès refusé : reconnectez-vous.":({conflict:"Une modification concurrente a été détectée. Actualisez le calendrier.",capacity_below_bookings:"Capacité inférieure au nombre d’inscrits.",past_date:"Impossible de modifier une date passée.",duplicate_class:"Cette classe existe déjà.",duplicate_slot:"Doublon détecté dans le planning."}[d.error]||"Erreur d’enregistrement ou de lecture : "+response.status));return d;
 }
 async function load(){
   const old=select.value;data=await call("/admin/planning");
   if(!Array.isArray(data.classes)||!Array.isArray(data.slots))throw Error("Planning invalide.");
   data.classes.sort((a,b)=>a.Ordre-b.Ordre); if(typeof previewCourse!=="undefined"){previewCourse.replaceChildren(...data.classes.filter(c=>c.Actif).map(c=>{const o=el("option",c.Niveau+" — "+c.Intitule);o.value=c.FormationID;return o;}));}
   select.replaceChildren(...data.classes.map(c=>{const o=el("option",c.Niveau+" — "+c.Intitule+" · "+c.Voie+(c.Actif?"":" (inactive)"));o.value=c.FormationID;return o;}));
   if(data.classes.some(c=>c.FormationID===old))select.value=old;
   summary.textContent=data.classes.length+" classes · "+data.slots.length+" journées enregistrées · "+data.slots.filter(s=>s.Ouvert).reduce((n,s)=>n+Number(s.Capacite),0)+" places prévues";
 }
 async function save(bodies){
   if(busy)return;busy=true;panel.setAttribute("aria-busy","true");feedback.textContent="Enregistrement…";let done=0;
   try{for(const b of bodies){await call("/admin/change",b);done++;}feedback.textContent=done+" modification(s) enregistrée(s) dans SharePoint.";}
   catch(e){feedback.textContent=done+" modification(s) enregistrée(s). "+e.message;}
   finally{try{await load();render();}catch(e){feedback.textContent+=" "+e.message;}busy=false;panel.removeAttribute("aria-busy");}
 }
 function editDay(id,date){
   if(busy)return;mode="calendar";select.value=id;render();
   const s=data.slots.find(x=>x.FormationID===id&&x.DateImmersion===date), cap=inp("number",s?.Capacite??2), open=inp("checkbox","");cap.min=1;cap.max=30;open.checked=s?s.Ouvert:true;
   editor.replaceChildren(el("h3",pretty(date)),field("Capacité prévue",cap),field("Journée ouverte",open),btn("Enregistrer",()=>save([{action:"day",id,date,capacity:Number(cap.value),open:open.checked,etag:s?._etag}])));editor.scrollIntoView({block:"nearest"});
 }
 function classEdit(c){
   const label=inp("text",c?.Niveau||""), title=inp("text",c?.Intitule||""), way=el("select"), active=inp("checkbox","");active.checked=c?c.Actif:true;
   for(const w of ["Voie scolaire","Apprentissage"]){const o=el("option",w);o.value=w;way.append(o);}if(c){way.value=c.Voie;way.disabled=true;}
   editor.replaceChildren(el("h3",c?"Modifier la classe":"Nouvelle classe"),field("Nom de classe",label),field("Intitulé public du BTS",title),field("Voie",way),field("Classe active",active),btn("Enregistrer la classe",()=>save([{action:"class",id:c?.FormationID,etag:c?._etag,label:label.value,title:title.value,way:way.value,active:active.checked}])));
 }
 function render(){
   content.replaceChildren();editor.replaceChildren();month.textContent=view.toLocaleDateString("fr-FR",{month:"long",year:"numeric"});
   select.hidden=mode!=="calendar";nav.hidden=mode==="classes"||mode==="reservations";
   const id=select.value, days=new Date(view.getFullYear(),view.getMonth()+1,0).getDate(), dateAt=d=>iso(new Date(view.getFullYear(),view.getMonth(),d));
   if(mode==="reservations"){
     const area=el("div");content.append(area);
     area.append(el("p","Chargement des réservations…"));
     call("/admin/reservations").then(d=>{
       area.replaceChildren();
       if(!d.rows.length)area.append(el("p","Aucune réservation réelle enregistrée."));
       const table=el("table"),head=el("tr"),wrap=el("div");wrap.className="tablewrap";
       for(const t of ["Candidat","Formation / classe","Date","État","Convention","Action"])head.append(el("th",t));table.append(head);
       for(const r of d.rows.sort((a,b)=>String(a.date).localeCompare(String(b.date)))){
         const tr=el("tr");
         for(const t of [r.prenom+" "+r.nom,r.formation+" / "+r.classe,r.date,r.statut,r.convention||"—"])tr.append(el("td",t));
         const td=el("td"),cancel=btn("Annuler",async()=>{if(busy||!confirm("Annuler cette réservation et libérer la place ?"))return;busy=true;try{await call("/admin/reservation-change",{action:"cancel",id:r.id,etag:r.etag});feedback.textContent="Réservation annulée.";render();}catch(e){feedback.textContent=e.message;}finally{busy=false;}});
         cancel.disabled=["annule","annulee","annulé","annulée","cancelled"].includes(String(r.statut).toLowerCase());td.append(cancel);tr.append(td);table.append(tr);
       }
       if(d.rows.length){wrap.append(table);area.append(wrap);}
       area.append(el("p",d.recipeCount+" ligne(s) de recette V0.6 exclue(s)."));
     }).catch(e=>{area.replaceChildren(el("p",e.message));});return;
   }
   if(mode==="classes"){for(const c of data.classes){const row=el("div");row.className="classrow";row.append(el("span",c.Niveau+" — "+c.Intitule+" · "+c.Voie+(c.Actif?"":" · inactive")),btn("Modifier",()=>classEdit(c)));content.append(row);}content.append(btn("Ajouter une classe",()=>classEdit(null)));return;}
   if(mode==="global"){
     const wrap=el("div");wrap.className="tablewrap";const table=el("table"), head=el("tr");head.append(el("th","Classe"));for(let d=1;d<=days;d++)head.append(el("th",String(d)));table.append(head);
     for(const c of data.classes){const tr=el("tr");tr.append(el("th",c.Niveau));for(let d=1;d<=days;d++){const td=el("td"), date=dateAt(d), s=data.slots.find(x=>x.FormationID===c.FormationID&&x.DateImmersion===date), b=btn(s?.Ouvert?String(s.Capacite):"—",()=>editDay(c.FormationID,date));b.className=s?.Ouvert?"day open":"day";td.append(b);tr.append(td);}table.append(tr);}wrap.append(table);content.append(wrap);return;
   }
   const grid=el("div");grid.className="calendar";for(const w of ["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"])grid.append(el("small",w));for(let i=0;i<(view.getDay()+6)%7;i++)grid.append(el("span"));
   for(let d=1;d<=days;d++){const date=dateAt(d),s=data.slots.find(x=>x.FormationID===id&&x.DateImmersion===date),b=btn(String(d),()=>editDay(id,date));b.className=s?.Ouvert?"day open":"day";b.append(el("small",s?.Ouvert?s.Capacite+" places":s?"Fermée":"—"));b.disabled=!id;grid.append(b);}content.append(grid);
   const multi=inp("date",""), end=inp("date",""), capacity=inp("number",2), weekdays=el("div");capacity.min=1;capacity.max=30;
   const checks=["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"].map((w,i)=>{const n=inp("checkbox","");n.checked=i<5;weekdays.append(field(w,n));return n;});
   const bulk=el("details");bulk.append(el("summary","Ouvrir plusieurs dates"),field("Du",multi),field("Au",end),weekdays,field("Capacité",capacity),btn("Ouvrir les dates sélectionnées",()=>{
     if(!id||!multi.value||!end.value||multi.value>end.value)return;const start=new Date(multi.value+"T12:00:00"),last=new Date(end.value+"T12:00:00");const bodies=[];
     if((last-start)/86400000>92){feedback.textContent="Sélectionnez une période de 3 mois maximum.";return;}
     for(let d=start;d<=last;d.setDate(d.getDate()+1)){const date=iso(d);if(!checks[(d.getDay()+6)%7].checked||date<iso(new Date()))continue;const s=data.slots.find(x=>x.FormationID===id&&x.DateImmersion===date);if(!s)bodies.push({action:"day",id,date,capacity:Number(capacity.value),open:true});}
     if(bodies.length&&confirm("Ouvrir "+bodies.length+" nouvelles journées ? Les dates existantes seront conservées."))save(bodies);
   }));content.append(bulk);
 }
 nav.className="toolbar";nav.append(btn("‹",()=>{view.setMonth(view.getMonth()-1);render();}),month,btn("›",()=>{view.setMonth(view.getMonth()+1);render();}));
 const tabs=el("div");tabs.className="toolbar";for(const [m,t] of [["calendar","Par classe"],["global","Vue globale"],["classes","Formations et classes"],["reservations","Réservations"]])tabs.append(btn(t,()=>{mode=m;render();}));tabs.append(btn("Actualiser",async()=>{try{await load();render();}catch(e){feedback.textContent=e.message;}}));
 const recipe=el("details"),output=el("p");recipe.append(el("summary","Nettoyage de la recette V0.6"),el("p","Supprime uniquement les inscriptions techniques RECETTE-V06 identifiées par le serveur."),btn("Nettoyer les tests V0.6",async()=>{if(busy||!confirm("Supprimer les réservations fictives V0.6 ?"))return;busy=true;try{const d=await call("/admin/reservation-change",{action:"cleanup-v06"});output.textContent=d.removed+" ligne(s) de recette supprimée(s).";}catch(e){output.textContent=e.message;}finally{busy=false;}}),output);panel.append(recipe);
 const previewBox=el("details"),previewCourse=el("select"),previewDate=inp("date",""),previewResult=el("pre");
 previewBox.append(el("summary","Recette métier V0.7"),el("p","Validation serveur d’un candidat fictif sur le vrai planning, sans enregistrement."),field("Formation",previewCourse),field("Date",previewDate),btn("Vérifier le créneau",async()=>{
   if(busy)return;const c=data.classes.find(x=>x.FormationID===previewCourse.value);if(!c||!previewDate.value)return;
   const candidate={Nom:"RECETTE",Prenom:"FICTIF",Email:"recette@example.invalid",Telephone:"0600000000",Etablissement:"LYCEE FICTIF",Civilite:"Madame",DateNaissance:"2009-01-01",VilleResidence:"VILLE FICTIVE",CodePostalResidence:"42000",Terminale:"Générale",SpecialitesBac:"Test",AdresseEtablissement:"Adresse fictive",EmailEtablissement:"lycee@example.invalid",TelephoneEtablissement:"0400000000",CiviliteChef:"Monsieur",ChefEtablissement:"CHEF FICTIF",Acceptation:true};
   try{previewResult.textContent=JSON.stringify(await call("/admin/booking-preview",{title:c.Intitule,way:c.Voie,date:previewDate.value,candidate}),null,2);}catch(e){previewResult.textContent=e.message;}
 }),previewResult);panel.append(previewBox);

 select.onchange=render;feedback.setAttribute("role","status");panel.append(el("h2","Gestion des immersions"),el("p","Inscriptions publiques désactivées — capacités théoriques."),summary,tabs,select,nav,content,editor,feedback);
 try{await load();render();}catch(e){feedback.textContent=e.message;}
}
