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
  const el=(tag,t)=>{const n=document.createElement(tag);n.textContent=t;return n;};
  panel.append(el("h2","Calendriers du Campus Initial"),el("p","Lecture des données SharePoint — inscriptions publiques désactivées."));
  try{
    const r=await fetch(api.replace("/admin/session","/admin/planning"),{headers:{Authorization:"Bearer "+token},cache:"no-store"});
    if(!r.ok)throw Error("planning");
    const d=await r.json();
    if(!Array.isArray(d.classes)||!Array.isArray(d.slots))throw Error("format");
    panel.append(el("p",d.classes.length+" classes · "+d.slots.length+" journées enregistrées"));
    const select=el("select","");select.setAttribute("aria-label","Classe");select.style.cssText="font:inherit;padding:10px;max-width:100%;border:1px solid #ccd9e4;border-radius:8px";
    for(const c of d.classes.sort((a,b)=>a.Ordre-b.Ordre)){const o=el("option",c.Niveau+" — "+c.Intitule);o.value=c.FormationID;select.append(o);}
    const list=el("div","");
    function render(){
      list.replaceChildren();
      const slots=d.slots.filter(s=>s.FormationID===select.value).sort((a,b)=>a.DateImmersion.localeCompare(b.DateImmersion));
      if(!slots.length){list.append(el("p","Aucune journée importée."));return;}
      const monthNames=new Intl.DateTimeFormat("fr-FR",{month:"long",year:"numeric",timeZone:"UTC"});
      const dayNames=new Intl.DateTimeFormat("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric",timeZone:"UTC"});
      let month="";
      for(const slot of slots){const date=new Date(slot.DateImmersion+"T12:00:00Z");const m=monthNames.format(date);if(m!==month){list.append(el("h3",m));month=m;}
        const row=el("p",dayNames.format(date)+" — "+slot.Capacite+" places prévues — "+(slot.Ouvert?"ouverte au planning":"fermée"));row.style.cssText="padding:10px;background:#eef4fa;border-radius:8px;margin:6px 0";list.append(row);}
    }
    select.onchange=render;panel.append(select,list);render();
  }catch{panel.append(el("p","Impossible de charger le planning. Vérifiez le déploiement de l'API V0.3.0 puis reconnectez-vous."));}
}
