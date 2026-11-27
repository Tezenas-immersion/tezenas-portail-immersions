"use strict";
const $=id=>document.getElementById(id), node=(tag,text)=>{const e=document.createElement(tag);e.textContent=text;return e;}, button=(text,fn)=>{const b=node("button",text);b.type="button";b.onclick=fn;return b;};
const today=new Date(),year=today.getMonth()>=8?today.getFullYear():today.getFullYear()-1;
$("year").textContent="Année scolaire "+year+"–"+(year+1);
let classes=[],slots=[],way="",course="",view=new Date(today.getFullYear(),today.getMonth(),1),selected="",loaded=false;
const alternate=["BTS Comptabilité et Gestion","BTS Professions immobilières","BTS NDRC","BTS MCO","BTS GPME"];
const iso=(y,m,d)=>[y,String(m+1).padStart(2,"0"),String(d).padStart(2,"0")].join("-");
const fmt=new Intl.DateTimeFormat("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric",timeZone:"UTC"});
function render(){
 $("ways").replaceChildren(...["Voie scolaire","Apprentissage"].map(w=>{const b=button(w==="Voie scolaire"?"Campus Initial — Voie scolaire":"Campus Alternance — Apprentissage",()=>{way=w;course="";selected="";render();});b.classList.toggle("active",way===w);return b;}));
 $("coursesPanel").hidden=!way;
 const names=way==="Apprentissage"?alternate:[...new Set(classes.map(c=>c.title))];
 $("courses").replaceChildren(...names.map(n=>{const b=button(n,()=>{course=n;selected="";const first=slots.filter(s=>classes.some(c=>c.id===s.classId&&c.title===n)).map(s=>s.date).sort()[0];if(first){const [y,m]=first.split("-").map(Number);view=new Date(y,m-1,1);}render();});b.classList.toggle("active",n===course);return b;}));
 $("courseInfo").textContent=way==="Apprentissage"?"Planning de l’alternance en attente de transmission.":way==="Voie scolaire"&&loaded&&!names.length?"Aucune formation publiée pour le moment.":"";
 $("calendarPanel").hidden=!course||way!=="Voie scolaire";
 $("month").textContent=view.toLocaleDateString("fr-FR",{month:"long",year:"numeric"});
 const ids=new Set(classes.filter(c=>c.title===course).map(c=>c.id)),cells=["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"].map(x=>node("small",x));
 for(let i=0;i<(view.getDay()+6)%7;i++)cells.push(node("span",""));
 for(let d=1;d<=new Date(view.getFullYear(),view.getMonth()+1,0).getDate();d++){
  const date=iso(view.getFullYear(),view.getMonth(),d),daySlots=slots.filter(s=>ids.has(s.classId)&&s.date===date),capacity=daySlots.reduce((a,s)=>a+s.remaining,0);
  const b=button(String(d),()=>{selected=date;render();});b.disabled=capacity===0;
  if(capacity){b.classList.add("open");b.append(node("small",capacity+" libres"));b.title=fmt.format(new Date(date+"T12:00:00Z"))+" — "+capacity+" places restantes";}
  if(daySlots.length&&!capacity)b.append(node("small","Complet"));if(selected===date)b.classList.add("active");cells.push(b);
 }
 $("calendar").replaceChildren(...cells);
 $("selection").textContent=selected?course+" — "+fmt.format(new Date(selected+"T12:00:00Z"))+". Les inscriptions ne sont pas encore ouvertes.":"";
 const formPanel=$("formPanel");
 if(formPanel){
   formPanel.hidden=!selected||way!=="Voie scolaire"||!slots.some(s=>ids.has(s.classId)&&s.date===selected&&s.remaining>0);
   $("bookingSummary").textContent=formPanel.hidden?"":course+" — "+fmt.format(new Date(selected+"T12:00:00Z"))+" — journée entière, horaires communiqués ultérieurement.";
   $("preview").textContent="";
 }
}
$("prev").onclick=()=>{view.setMonth(view.getMonth()-1);selected="";render();};
$("next").onclick=()=>{view.setMonth(view.getMonth()+1);selected="";render();};
render();
(async()=>{try{const r=await fetch("https://tezenas-immersions-api.tezenas-immersion.workers.dev/public/planning",{cache:"no-store"});if(!r.ok)throw Error();const d=await r.json();if(!Array.isArray(d.classes)||!Array.isArray(d.slots))throw Error();classes=d.classes.sort((a,b)=>a.order-b.order);slots=d.slots;loaded=true;$("status").textContent="";render();}catch{$("status").textContent="Le planning est momentanément indisponible. Veuillez réessayer ultérieurement.";}})();

const form=$("bookingForm");
if(form){
 const terminale=$("terminale"), birth=$("birth");
 birth.max=iso(today.getFullYear(),today.getMonth(),today.getDate());
 function sync(){
   $("specialiteLabel").textContent=terminale.value==="Technologique"?"Série et option du bac":terminale.value==="Professionnelle"?"Spécialité professionnelle":"Spécialité(s) du bac";
   $("preview").textContent="";
 }
 terminale.addEventListener("change",sync);
 birth.addEventListener("change",()=>{
   const d=birth.value;
   if(!d)return;
   const [y,m,day]=d.split("-").map(Number);
   const age=today.getFullYear()-y-((today.getMonth()+1<m||(today.getMonth()+1===m&&today.getDate()<day))?1:0);
   $("emailLabel").textContent=age<18?"Adresse mail des parents (élève mineur)":"Adresse mail de l’élève";
 });
 form.addEventListener("input",()=>{$("preview").textContent="";});
 form.addEventListener("submit",e=>{
   e.preventDefault();
   if(!selected||$("formPanel").hidden||!form.reportValidity())return;
   const data=Object.fromEntries(new FormData(form));
   if(Object.entries(data).some(([k,v])=>k!=="Acceptation"&&!String(v).trim())){
     $("preview").textContent="Veuillez compléter tous les champs obligatoires.";return;
   }
   $("preview").textContent="Formulaire de démonstration validé. Aucune réservation, donnée ou convention n’a été envoyée ou enregistrée.";
 });
 $("formFields").disabled=false;
}
