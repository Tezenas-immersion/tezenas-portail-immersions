export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = request.headers.get("Origin");
    const headers = new Headers({"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store","X-Content-Type-Options":"nosniff","Vary":"Origin"});
    if(origin === "https://tezenas-immersion.github.io") {
      headers.set("Access-Control-Allow-Origin",origin);
      headers.set("Access-Control-Allow-Methods","GET, OPTIONS");
    }
    const json=(data,status=200)=>new Response(JSON.stringify(data),{status,headers});
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
        return json({status:"ok",entra:true,sharepoint:true,lists:true,write_test:false});
      } catch {return json({status:"error",step},502);}
    }
    if(request.method === "OPTIONS") return new Response(null,{status:204,headers});
    if(request.method !== "GET") return json({error:"method_not_allowed"},405);
    if(url.pathname === "/" || url.pathname === "/health") return json({service:"tezenas-immersions-api",version:"0.1.1",status:"ok",mode:"demo",reservations:false});
    return json({error:"not_found"},404);
  }
};
