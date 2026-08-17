export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const pathname = url.pathname;

    // CORS Headers
    const corsHeaders = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization"
    };

    // Preflight Handling
    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 200,
        headers: corsHeaders
      });
    }

    // Router for API endpoints
    if (pathname === "/api/data" || pathname.startsWith("/api/data/")) {
      if (request.method !== "GET") {
        return new Response(JSON.stringify({ error: "Method Not Allowed. Only GET accepted." }), {
          status: 405,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      const program = (url.searchParams.get("program") || "").toUpperCase();
      const urlMap = {
        ESL: env.ESL_URL,
        HSE: env.HSE_URL,
        CPW: env.CPW_URL
      };

      const targetUrl = urlMap[program];
      if (!targetUrl) {
        return new Response(JSON.stringify({ error: "Invalid or missing program parameter. Must be ESL, HSE, or CPW." }), {
          status: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      try {
        const fetchOptions = { redirect: "follow" };
        if (request.cf) {
          fetchOptions.cf = { cacheEverything: true, cacheTtl: 300 };
        }
        const response = await fetch(targetUrl, fetchOptions);
        const data = await response.json();
        return new Response(JSON.stringify(data), {
          status: 200,
          headers: {
            ...corsHeaders,
            "Content-Type": "application/json",
            "Cache-Control": "public, max-age=300, s-maxage=300, stale-while-revalidate=600"
          }
        });
      } catch (err) {
        console.error("Data proxy error:", err);
        return new Response(JSON.stringify({ error: "Failed to retrieve schedule data from backend." }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    if (pathname === "/api/submit" || pathname.startsWith("/api/submit/")) {
      if (request.method !== "POST") {
        return new Response(JSON.stringify({ error: "Method Not Allowed. Only POST accepted." }), {
          status: 405,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      const mode = url.searchParams.get("mode") || env.SUBMIT_MODE || "test";
      const isProd = mode.toLowerCase() === "prod" || mode.toLowerCase() === "production";
      const submitUrl = isProd
        ? (env.SUBMIT_URL_PROD || env.SUBMIT_URL)
        : (env.SUBMIT_URL_TEST || env.SUBMIT_URL);

      if (!submitUrl) {
        return new Response(JSON.stringify({ error: "Target submission URL missing." }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }

      try {
        const reqBody = await request.text();
        const response = await fetch(submitUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/x-www-form-urlencoded"
          },
          body: reqBody
        });

        const resText = await response.text();
        return new Response(resText, {
          status: response.status,
          headers: { ...corsHeaders, "Content-Type": "text/plain" }
        });
      } catch (err) {
        return new Response(JSON.stringify({ error: "Failed to forward request to target endpoint." }), {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    if (pathname === "/" || pathname === "") {
      return new Response(JSON.stringify({
        status: "online",
        message: "ACC Substitute System - Cloudflare Worker API Proxy",
        endpoints: ["/api/data?program=ESL", "/api/submit"]
      }, null, 2), {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    return new Response(JSON.stringify({ error: "Not Found" }), {
      status: 404,
      headers: { ...corsHeaders, "Content-Type": "application/json" }
    });
  }
};
