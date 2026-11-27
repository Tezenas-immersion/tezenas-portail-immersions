import {xlsx,zip,saveFile} from "./exports.js?v=0120";
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
 const now=new Date();let view=new Date(now.getFullYear(),now.getMonth(),1), data={classes:[],slots:[]}, mode="overview", busy=false;
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
   data.classes.sort((a,b)=>a.Ordre-b.Ordre); const previousConcurrency=concurrencyCourse.value;concurrencyCourse.replaceChildren(...data.classes.filter(c=>c.Actif&&c.Voie==="Voie scolaire").map(c=>{const o=el("option",c.Niveau+" — "+c.Intitule);o.value=c.FormationID;return o;}));if(data.classes.some(c=>c.FormationID===previousConcurrency))concurrencyCourse.value=previousConcurrency; bookingCourse.replaceChildren(...data.classes.filter(c=>c.Actif).map(c=>{const o=el("option",c.Niveau+" — "+c.Intitule);o.value=c.FormationID;return o;})); if(typeof previewCourse!=="undefined"){previewCourse.replaceChildren(...data.classes.filter(c=>c.Actif).map(c=>{const o=el("option",c.Niveau+" — "+c.Intitule);o.value=c.FormationID;return o;}));}
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
   select.hidden=mode!=="calendar";nav.hidden=mode==="classes"||mode==="reservations"||mode==="overview";
   const id=select.value, days=new Date(view.getFullYear(),view.getMonth()+1,0).getDate(), dateAt=d=>iso(new Date(view.getFullYear(),view.getMonth(),d));
   if(mode==="overview"){
     const area=el("div","Chargement des indicateurs…");content.append(area);
     call("/admin/reservations").then(d=>{
       area.replaceChildren();
       const active=r=>!["annule","annulee","annulé","annulée","cancelled"].includes(String(r.statut).toLowerCase());
       const real=d.rows.filter(r=>!r.test&&active(r)),tests=d.rows.filter(r=>r.test&&active(r));
       const classes=data.classes.filter(c=>c.Actif),ids=new Set(classes.map(c=>c.FormationID));
       const slots=data.slots.filter(s=>s.Ouvert&&ids.has(s.FormationID)),capacity=slots.reduce((n,s)=>n+Number(s.Capacite),0);
       const occupied=slots.reduce((n,s)=>n+d.rows.filter(r=>active(r)&&r.slot===s.DisponibiliteID).length,0);
       const metrics=el("div");metrics.className="metrics";
       for(const [label,value] of [["Places au planning",capacity],["Inscriptions actives réelles",real.length],["Places non occupées au planning",Math.max(0,capacity-occupied)],["Conventions réelles disponibles",real.filter(r=>r.hasPdf).length],["Réservations de recette actives",tests.length],["Annulations réelles",d.rows.filter(r=>!r.test&&!active(r)).length]]){
         const card=el("div");card.append(el("strong",String(value)),el("span",label));metrics.append(card);
       }
       area.append(el("h3","Pilotage — toutes années enregistrées"),metrics,el("p","Capacité du planning ouvert, y compris les dates passées et les dates clôturées à J-7. Les tests occupent des places mais ne sont pas comptés comme inscriptions réelles."));
       for(const title of [...new Set(classes.map(c=>c.Intitule))]){
         const group=new Set(classes.filter(c=>c.Intitule===title).map(c=>c.FormationID));
         const ss=slots.filter(x=>group.has(x.FormationID)),cap=ss.reduce((n,x)=>n+Number(x.Capacite),0),used=real.filter(r=>ss.some(x=>x.DisponibiliteID===r.slot)).length;
         const p=el("progress");p.max=Math.max(1,cap);p.value=used;area.append(el("p",title+" — "+used+" / "+cap+" inscriptions réelles"),p);
       }
     }).catch(e=>{area.textContent=e.message;});return;
   }
   if(mode==="reservations"){
     const area=el("div");content.append(area);
     area.append(el("p","Chargement des réservations…"));
     call("/admin/reservations").then(d=>{
       area.replaceChildren();
       const filters=el("div");filters.className="filters";
       const search=inp("search",""), campus=el("select"), formation=el("select"), classe=el("select"), year=el("select"), state=el("select"), kind=el("select"), from=inp("date",""), to=inp("date","");
       const options=(node,values,all)=>{node.append(el("option",all));node.firstChild.value="";for(const [value,label] of values){const o=el("option",label);o.value=value;node.append(o);}};
       const unique=xs=>[...new Set(xs.filter(Boolean))].sort().map(x=>[x,x]);
       options(campus,unique(d.rows.map(r=>r.way)),"Tous campus");options(formation,unique(d.rows.map(r=>r.formation)),"Toutes formations");
       options(classe,data.classes.map(c=>[c.FormationID,c.Niveau+" — "+c.Intitule]),"Toutes classes");options(year,unique(d.rows.map(r=>r.year)),"Toutes années");
       options(state,[["active","Actives"],["cancelled","Annulées"]],"Tous statuts");
       options(kind,[["real","Réelles"],["test","Recette"]],"Toutes catégories");kind.value="real";
       const isCancelled=r=>["annule","annulee","annulé","annulée","cancelled"].includes(String(r.statut).toLowerCase());
       const fold=x=>String(x||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase();
       const matches=r=>(!search.value||fold(r.nom+" "+r.prenom+" "+r.email).includes(fold(search.value)))&&(!campus.value||r.way===campus.value)&&(!formation.value||r.formation===formation.value)&&(!classe.value||r.classe===classe.value)&&(!year.value||r.year===year.value)&&(!state.value||(state.value==="cancelled")===isCancelled(r))&&(!kind.value||(kind.value==="test")===r.test)&&(!from.value||r.date>=from.value)&&(!to.value||r.date<=to.value);
       for(const [l,n] of [["Recherche",search],["Campus",campus],["Formation",formation],["Classe",classe],["Année scolaire",year],["Statut",state],["Catégorie",kind],["Du",from],["Au",to]])filters.append(field(l,n));
       area.append(filters);
       const count=el("p"),exportBar=el("div");exportBar.className="toolbar";area.append(count,exportBar);
       const selected=()=>d.rows.filter(matches);
       const classLabel=r=>data.classes.find(c=>c.FormationID===r.classe)?.Niveau||r.classe;
       const candidateGroups=[
         ["Identité",["Civilite","Nom","Prenom","DateNaissance","Telephone","Email","VilleResidence","CodePostalResidence"]],
         ["Scolarité",["Terminale","SpecialitesBac","Etablissement","AdresseEtablissement","EmailEtablissement","TelephoneEtablissement","CiviliteChef","ChefEtablissement"]],
         ["Immersion",["IntituleFormation","FormationID","Voie","AnneeID","DateImmersion","Statut","DateReservation","StatutConvention","StatutMail","ReservationID","DisponibiliteID","Acceptation","DateAcceptation","VersionAcceptation"]]
       ];
       const keys=candidateGroups.flatMap(g=>g[1]);
       const labels={Civilite:"Civilité",Prenom:"Prénom",DateNaissance:"Date de naissance",VilleResidence:"Ville",CodePostalResidence:"Code postal",SpecialitesBac:"Spécialités",Etablissement:"Établissement",AdresseEtablissement:"Adresse établissement",EmailEtablissement:"Mail établissement",TelephoneEtablissement:"Téléphone établissement",CiviliteChef:"Civilité direction",ChefEtablissement:"Chef d’établissement",IntituleFormation:"Formation",FormationID:"Classe",AnneeID:"Année scolaire",DateImmersion:"Date d’immersion",DateReservation:"Date de réservation",StatutConvention:"Convention",StatutMail:"Confirmation mail",ReservationID:"Référence",DisponibiliteID:"Créneau",DateAcceptation:"Date de confirmation",VersionAcceptation:"Version de confirmation"};
       const dateFR=x=>/^\d{4}-\d{2}-\d{2}$/.test(String(x))?String(x).split("-").reverse().join("/"):x;
       const display=(k,v)=>["DateReservation","DateAcceptation"].includes(k)&&v&&Number.isFinite(Date.parse(v))?new Date(v).toLocaleString("fr-FR",{timeZone:"Europe/Paris"}):k==="FormationID"?(data.classes.find(c=>c.FormationID===v)?.Niveau||v):dateFR(v);
       const excel=btn("Exporter Excel (.xlsx)",async()=>{if(busy)return;const rows=selected();if(!rows.length)return;if(rows.length>500){feedback.textContent="Limite de 500 lignes : affinez les filtres.";return;}if(!confirm("Exporter "+rows.length+" dossier(s) candidat contenant des données personnelles ?"))return;busy=true;excel.disabled=true;
         try{const out=[];for(const r of rows){const v=await call("/admin/candidate?id="+encodeURIComponent(r.id));out.push(keys.map(k=>display(k,v.candidate[k])));}saveFile(xlsx(keys.map(k=>labels[k]||k),out),"ImmersTez_Inscriptions_"+iso(new Date())+".xlsx");feedback.textContent=rows.length+" dossier(s) exporté(s).";}
         catch(e){feedback.textContent="Export Excel interrompu : "+e.message;alert(feedback.textContent);}finally{busy=false;excel.disabled=false;}
       });
       const archive=btn("Exporter conventions ZIP",async()=>{if(busy)return;const rows=selected().filter(r=>r.hasPdf);if(!rows.length){alert("Aucun PDF disponible dans cette sélection.");return;}if(rows.length>30){alert("Limite de 30 PDF par export : affinez les filtres.");return;}if(!confirm("Télécharger "+rows.length+" convention(s) contenant des données personnelles ?"))return;busy=true;archive.disabled=true;
         try{const files=[],names=new Set();let size=0;for(const r of rows){const res=await fetch(endpoint+"/admin/convention-download?id="+encodeURIComponent(r.id),{headers:{Authorization:"Bearer "+token},cache:"no-store"});if(!res.ok)throw Error("PDF indisponible : "+r.id);const bytes=await res.arrayBuffer();size+=bytes.byteLength;if(size>50000000)throw Error("Limite de 50 Mo : affinez les filtres.");let name=res.headers.get("Content-Disposition")?.match(/filename="([A-Za-z0-9_.-]+)"/)?.[1]||"Convention.pdf";const base=name;let i=2;while(names.has(name.toLowerCase()))name=base.replace(/\.pdf$/i,"_"+i++ +".pdf");names.add(name.toLowerCase());files.push({name,data:bytes});}saveFile(zip(files),"ImmersTez_Conventions_"+iso(new Date())+".zip");feedback.textContent=files.length+" PDF exportés.";}
         catch(e){feedback.textContent="Export ZIP interrompu : "+e.message;alert(feedback.textContent);}finally{busy=false;archive.disabled=false;}
       });exportBar.append(excel,archive); const reset=btn("Réinitialiser les filtres",()=>{for(const n of [search,campus,formation,classe,year,state,from,to])n.value="";kind.value="real";refreshRows();});exportBar.append(reset);

       const table=el("table"),head=el("tr"),wrap=el("div");wrap.className="tablewrap";
       for(const t of ["Candidat","Formation / classe","Date","État","Convention","Mail","Action"])head.append(el("th",t));table.append(head);
       const refreshRows=()=>{for(const row of Array.from(table.querySelectorAll("tr")).slice(1))row.remove();const rows=selected();count.textContent=rows.length?rows.length+" dossier(s) correspondant aux filtres":"Aucune réservation correspondant aux filtres.";excel.disabled=!rows.length;archive.disabled=!rows.some(r=>r.hasPdf);for(const r of rows.sort((a,b)=>String(a.date).localeCompare(String(b.date)))){
         const tr=el("tr");
         for(const t of [r.prenom+" "+r.nom,r.formation+" / "+classLabel(r),dateFR(r.date),isCancelled(r)?"Annulée":r.test?"Recette":r.statut,r.convention==="Recette PDF accepte Graph"?"PDF accepté par Graph":r.convention||"—",r.mail==="Accepte Graph"?"Accepté par Graph":r.mail||"—"])tr.append(el("td",t));
         const td=el("td"),cancel=btn("Annuler",async()=>{if(busy||!confirm("Annuler cette réservation et libérer la place ?"))return;busy=true;try{await call("/admin/reservation-change",{action:"cancel",id:r.id,etag:r.etag});feedback.textContent="Réservation annulée.";render();}catch(e){feedback.textContent=e.message;}finally{busy=false;}});
         cancel.disabled=["annule","annulee","annulé","annulée","cancelled"].includes(String(r.statut).toLowerCase());td.append(cancel);
         const details=btn("Fiche candidat",async()=>{if(busy)return;busy=true;try{
           const d=await call("/admin/candidate?id="+encodeURIComponent(r.id));editor.replaceChildren(el("h3","Fiche candidat — "+r.prenom+" "+r.nom));
           for(const [group,fields] of candidateGroups){editor.append(el("h4",group));const box=el("div");box.className="candidate-grid filters";for(const k of fields){const line=el("div");line.append(el("strong",labels[k]||k),el("p",String(display(k,d.candidate[k])??"—")));box.append(line);}editor.append(box);}
           editor.scrollIntoView({block:"nearest"});
         }catch(e){feedback.textContent=e.message;}finally{busy=false;}});
         td.append(details);
         if(r.id.startsWith("TEST-V08-")&&r.statut==="Test"&&r.convention!=="Non applicable"){
           const generate=btn("Générer la convention",async()=>{if(busy||!confirm("Générer la convention non signée de recette dans SharePoint ? Aucun mail ne sera envoyé."))return;busy=true;generate.disabled=true;generate.textContent="Génération en cours…";feedback.textContent="Génération Word et PDF en cours…";
             try{const res=await fetch(endpoint+"/admin/convention-generate",{method:"POST",headers:{Authorization:"Bearer "+token,"Content-Type":"application/json"},body:JSON.stringify({id:r.id}),cache:"no-store"});const d=await res.json();if(!res.ok)throw Error(d.error||res.status);feedback.textContent="Convention de recette générée et enregistrée dans SharePoint.";render();}
             catch(e){feedback.textContent="Génération : "+e.message;alert("Génération : "+e.message);}finally{busy=false;generate.disabled=false;generate.textContent="Générer la convention";}
           });td.append(generate);
         }
         if(["Recette generee","Recette PDF accepte Graph","Recette PDF envoi incertain","Recette PDF echec"].includes(r.convention)){
           const download=btn("Télécharger PDF",async()=>{if(busy)return;busy=true;try{
             const res=await fetch(endpoint+"/admin/convention-download?id="+encodeURIComponent(r.id),{headers:{Authorization:"Bearer "+token},cache:"no-store"});
             if(!res.ok){const d=await res.json();throw Error(d.error||res.status);}
             const blob=await res.blob();if(blob.type!=="application/pdf")throw Error("Format inattendu");
             const url=URL.createObjectURL(blob);const a=el("a");a.href=url;a.download=res.headers.get("Content-Disposition")?.match(/filename="([A-Za-z0-9_.-]+)"/)?.[1]||"Convention.pdf";document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);
           }catch(e){feedback.textContent=e.message;}finally{busy=false;}});td.append(download);
         }
         if(r.id.startsWith("TEST-V08-")&&r.statut==="Test"&&["En attente","Recette generee","Recette PDF echec"].includes(r.convention)){
           const send=btn("Traiter la convention de recette",async()=>{if(busy||!confirm("Générer si nécessaire et envoyer le PDF de recette uniquement à l’adresse Trustteam autorisée ?"))return;busy=true;send.disabled=true;send.textContent="Traitement en cours…";
             try{const res=await fetch(endpoint+"/admin/convention-send-test",{method:"POST",headers:{Authorization:"Bearer "+token,"Content-Type":"application/json"},body:JSON.stringify({id:r.id}),cache:"no-store"});const d=await res.json();if(!res.ok)throw Error(d.error||res.status);feedback.textContent=d.convention;render();}
             catch(e){feedback.textContent="Convention : "+e.message;alert(feedback.textContent);}finally{busy=false;send.disabled=false;}
           });td.append(send);
         }
         tr.append(td);table.append(tr);
       }
       }; refreshRows();for(const n of [search,campus,formation,classe,year,state,kind,from,to])n.addEventListener("input",refreshRows);wrap.append(table);area.append(wrap);
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
 const tabs=el("div");tabs.className="toolbar";for(const [m,t] of [["overview","Tableau de bord"],["calendar","Par classe"],["global","Vue globale"],["classes","Formations et classes"],["reservations","Réservations"]])tabs.append(btn(t,()=>{mode=m;render();}));tabs.append(btn("Actualiser",async()=>{try{await load();render();}catch(e){feedback.textContent=e.message;}}));
 const recipe=el("details"),output=el("p");recipe.append(el("summary","Nettoyage de la recette V0.6"),el("p","Supprime uniquement les inscriptions techniques RECETTE-V06 identifiées par le serveur."),btn("Nettoyer les tests V0.6",async()=>{if(busy||!confirm("Supprimer les réservations fictives V0.6 ?"))return;busy=true;try{const d=await call("/admin/reservation-change",{action:"cleanup-v06"});output.textContent=d.removed+" ligne(s) de recette supprimée(s).";}catch(e){output.textContent=e.message;}finally{busy=false;}}),output);panel.append(recipe);
 const previewBox=el("details"),previewCourse=el("select"),previewDate=inp("date",""),previewResult=el("pre");
 previewBox.append(el("summary","Recette métier V0.7"),el("p","Validation serveur d’un candidat fictif sur le vrai planning, sans enregistrement."),field("Formation",previewCourse),field("Date",previewDate),btn("Vérifier le créneau",async()=>{
   if(busy)return;const c=data.classes.find(x=>x.FormationID===previewCourse.value);if(!c||!previewDate.value)return;
   const candidate={Nom:"RECETTE",Prenom:"FICTIF",Email:"recette@example.invalid",Telephone:"0600000000",Etablissement:"LYCEE FICTIF",Civilite:"Madame",DateNaissance:"2009-01-01",VilleResidence:"VILLE FICTIVE",CodePostalResidence:"42000",Terminale:"Générale",SpecialitesBac:"Test",AdresseEtablissement:"Adresse fictive",EmailEtablissement:"lycee@example.invalid",TelephoneEtablissement:"0400000000",CiviliteChef:"Monsieur",ChefEtablissement:"CHEF FICTIF",Acceptation:true};
   try{previewResult.textContent=JSON.stringify(await call("/admin/booking-preview",{title:c.Intitule,way:c.Voie,date:previewDate.value,candidate}),null,2);}catch(e){previewResult.textContent=e.message;}
 }),previewResult);panel.append(previewBox);

 const mailBox=el("details"),mailAddress=inp("email",""),mailOutput=el("p");
 mailAddress.placeholder="prenom.nom@groupetezenas.org";mailAddress.maxLength=254;mailAddress.required=true;
 const sendTest=btn("Envoyer un mail de test",async()=>{
   if(busy)return;const email=mailAddress.value.trim();
   if(!mailAddress.reportValidity()||(!/@(groupetezenas\.org|tezenas\.org)$/i.test(email)&&email.toLowerCase()!=="ludovic.lorito@trustteam.fr")){mailOutput.textContent="Utilisez une adresse Tézenas ou votre adresse Trustteam autorisée pour la recette.";return;}
   if(!confirm("Envoyer un mail technique fictif à "+email+" ?"))return;
   busy=true;sendTest.disabled=true;mailOutput.textContent="Envoi en cours…";
   try{
     const r=await fetch(endpoint+"/admin/mail-test",{method:"POST",headers:{Authorization:"Bearer "+token,"Content-Type":"application/json"},body:JSON.stringify({email}),cache:"no-store"});
     const d=await r.json();
     mailOutput.textContent=r.status===202&&d.status==="accepted"?"HTTP 202 — Microsoft Graph a accepté le message. Vérifiez sa réception et les éléments envoyés.":d.error==="mail_permission_denied"?"Graph HTTP 403 — autorisation refusée, vérifiez notamment la propagation RBAC.":"Échec : "+(d.error||r.status)+(d.graphHttp?" (Graph HTTP "+d.graphHttp+")":"");
   }catch{mailOutput.textContent="Résultat incertain. Vérifiez la réception avant de relancer.";}
   finally{busy=false;sendTest.disabled=false;}
 });
 mailOutput.setAttribute("role","status");mailBox.append(el("summary","Recette messagerie V0.7.2"),el("p","Envoi réel depuis la boîte Immers’Tez, vers une adresse Tézenas ou l’adresse Trustteam autorisée pour le test externe. Aucun candidat ni réservation créé."),field("Adresse professionnelle destinataire",mailAddress),sendTest,mailOutput);panel.append(mailBox);
 const bookingBox=el("details"),bookingCourse=el("select"),bookingDate=inp("date",""),bookingEmail=inp("email",""),bookingOutput=el("pre");
 let pendingBooking=null;
 bookingEmail.placeholder="Adresse de recette autorisée";bookingEmail.required=true;
 const createBooking=btn("Créer une réservation fictive",async()=>{
   if(busy)return;const c=data.classes.find(x=>x.FormationID===bookingCourse.value);
   if(!c||!bookingDate.value||!bookingEmail.reportValidity())return;
   const body={title:c.Intitule,way:c.Voie,date:bookingDate.value,email:bookingEmail.value.trim()};
   const key=JSON.stringify(body);
   if(pendingBooking&&pendingBooking.key!==key){bookingOutput.textContent="Une demande précédente est en attente. Vérifiez son résultat avec les mêmes paramètres avant de continuer.";return;}
   if(!confirm("Créer une réservation fictive, envoyer la confirmation et traiter la convention PDF de recette ?"))return;
   if(!pendingBooking)pendingBooking={key,requestId:"v08-"+crypto.randomUUID().replaceAll("-","")};
   busy=true;createBooking.disabled=true;
   try{
     const response=await fetch(endpoint+"/admin/booking-create",{method:"POST",headers:{Authorization:"Bearer "+token,"Content-Type":"application/json"},body:JSON.stringify({...body,requestId:pendingBooking.requestId}),cache:"no-store"});
     const d=await response.json();bookingOutput.textContent=JSON.stringify(d,null,2);
     if(response.ok||[400,409].includes(response.status))pendingBooking=null;
     if(response.ok){await load();render();}
   }catch{bookingOutput.textContent="Résultat incertain. Relancez avec les mêmes paramètres sans actualiser la page.";}
   finally{busy=false;createBooking.disabled=false;}
 });
 bookingBox.append(el("summary","Recette inscription V0.10"),el("p","Candidat fictif imposé par le serveur. Confirmation au destinataire choisi et convention Initial envoyée uniquement à Trustteam. Occupe une place : annulez après le test."),field("Formation",bookingCourse),field("Date",bookingDate),field("Adresse de réception du mail de recette",bookingEmail),createBooking,bookingOutput);panel.append(bookingBox);
 const concurrencyBox=el("details"),concurrencyCourse=el("select"),concurrencyDate=inp("date",""),concurrencyOutput=el("pre");
 let testIds=[];
 const concurrencyButton=btn("Tester les réservations simultanées",async()=>{
   if(busy||testIds.length)return;
   const c=data.classes.find(x=>x.FormationID===concurrencyCourse.value);
   if(!c||!concurrencyDate.value)return;
   busy=true;concurrencyButton.disabled=true;concurrencyOutput.textContent="Contrôle des disponibilités…";
   try{
     const p=await fetch(endpoint+"/public/planning",{cache:"no-store"});
     if(!p.ok)throw Error("Planning indisponible");
     const d=await p.json(),ids=new Set(d.classes.filter(x=>x.title===c.Intitule).map(x=>x.id));
     const free=d.slots.filter(x=>ids.has(x.classId)&&x.date===concurrencyDate.value).reduce((n,x)=>n+Number(x.remaining),0);
     if(!Number.isInteger(free)||free<1||free>4)throw Error("Choisissez une journée Initial avec 1 à 4 places libres.");
     if(!confirm("Envoyer "+(free+1)+" demandes simultanées pour "+free+" places libres ? Aucun mail ne sera envoyé. Les réservations fictives occuperont les places jusqu’à leur annulation."))return;
     const payload={title:c.Intitule,way:c.Voie,date:concurrencyDate.value,capacityTest:true};
     const keys=Array.from({length:free+1},()=>"v08-"+crypto.randomUUID().replaceAll("-",""));
     async function attempt(requestId){
       try{const res=await fetch(endpoint+"/admin/booking-create",{method:"POST",headers:{Authorization:"Bearer "+token,"Content-Type":"application/json"},body:JSON.stringify({...payload,requestId}),cache:"no-store"});const result=await res.json();return {http:res.status,requestId,...result};}
       catch{return {http:0,requestId,error:"resultat_incertain"};}
     }
     const results=await Promise.all(keys.map(attempt));
     testIds=results.filter(x=>x.http===201&&x.capacityTest===true&&x.id).map(x=>x.id);
     const winner=results.find(x=>x.http===201);
     const replay=winner?await attempt(winner.requestId):null;
     const p2=await fetch(endpoint+"/public/planning",{cache:"no-store"});
     let after=null;if(p2.ok){const d2=await p2.json();after=d2.slots.filter(x=>ids.has(x.classId)&&x.date===payload.date).reduce((n,x)=>n+Number(x.remaining),0);}
     const accepted=results.filter(x=>x.http===201&&x.capacityTest===true).length,rejected=results.filter(x=>x.http===409&&x.error==="full_or_closed").length;
     concurrencyOutput.textContent=JSON.stringify({resultat:accepted===free&&rejected===1&&replay?.replayed===true&&replay.id===winner.id&&after===0?"OK":"A VERIFIER",placesAvant:free,acceptes:accepted,refuses:rejected,placesApres:after,tentatives:results,repetition:replay},null,2);
     if(results.some(x=>x.http===0||x.http>=500))concurrencyOutput.textContent+="\nVérifiez les lignes RECETTE CONCURRENCE dans Réservations : certaines écritures peuvent être incertaines.";
   }catch(e){concurrencyOutput.textContent=e.message;}
   finally{busy=false;concurrencyButton.disabled=testIds.length>0;}
 });
 const cleanup=btn("Annuler les réservations de ce test",async()=>{
   if(busy||!testIds.length||!confirm("Libérer les places occupées par les réservations fictives de ce test ?"))return;
   busy=true;cleanup.disabled=true;
   try{
     const d=await call("/admin/reservations");let n=0;
     for(const id of [...testIds]){
       const row=d.rows.find(x=>x.id===id&&x.nom==="RECETTE CONCURRENCE");
       if(!row)continue;
       await call("/admin/reservation-change",{action:"cancel",id,etag:row.etag});testIds=testIds.filter(x=>x!==id);n++;
     }
     concurrencyOutput.textContent+="\n"+n+" réservation(s) annulée(s). Vérifiez le retour des places sur le portail public.";
     if(testIds.length)concurrencyOutput.textContent+="\nCertaines lignes restent à vérifier dans Réservations.";
     mode="reservations";render();
   }catch(e){concurrencyOutput.textContent+="\nNettoyage : "+e.message;}
   finally{busy=false;cleanup.disabled=false;concurrencyButton.disabled=testIds.length>0;}
 });
 concurrencyBox.append(el("summary","Recette concurrence V0.11"),el("p","Test sur un vrai créneau Initial, sans mail ni convention. Les places doivent être libérées après le test."),field("Formation",concurrencyCourse),field("Date",concurrencyDate),concurrencyButton,cleanup,concurrencyOutput);panel.append(concurrencyBox);
 select.onchange=render;feedback.setAttribute("role","status");const technical=el("details");technical.className="technical";technical.append(el("summary","Outils techniques — recette"));for(const child of Array.from(panel.children))if(child.tagName==="DETAILS")technical.append(child);panel.append(el("h2","Gestion des immersions"),el("p","Inscriptions publiques désactivées — planning de démonstration."),summary,tabs,select,nav,content,editor,feedback,technical);
 try{await load();render();}catch(e){feedback.textContent=e.message;}
}
