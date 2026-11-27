export default {
  async fetch(request) {
    const url = new URL(request.url);
    const origin = request.headers.get("Origin");
    const headers = new Headers({"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store","X-Content-Type-Options":"nosniff","Vary":"Origin"});
    if(origin === "https://tezenas-immersion.github.io") {
      headers.set("Access-Control-Allow-Origin",origin);
      headers.set("Access-Control-Allow-Methods","GET, OPTIONS");
    }
    if(request.method === "OPTIONS") return new Response(null,{status:204,headers});
    if(request.method !== "GET") return new Response(JSON.stringify({error:"method_not_allowed"}),{status:405,headers});
    if(url.pathname === "/" || url.pathname === "/health") return new Response(JSON.stringify({service:"tezenas-immersions-api",version:"0.1.0",status:"ok",mode:"demo",reservations:false}),{headers});
    return new Response(JSON.stringify({error:"not_found"}),{status:404,headers});
  }
};