export default async function handler(req, res) {
  // CORS Headers
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  // Preflight Handling
  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  // Method Enforcement
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed. Only POST requests are accepted." });
  }

  const requestedMode = req.query.mode || process.env.SUBMIT_MODE || "test";
  const isProd = requestedMode.toLowerCase() === "prod" || requestedMode.toLowerCase() === "production";
  const submitUrl = isProd
    ? (process.env.SUBMIT_URL_PROD || process.env.SUBMIT_URL)
    : (process.env.SUBMIT_URL_TEST || process.env.SUBMIT_URL);

  if (!submitUrl) {
    console.error("Target submission URL environment variable is missing.");
    return res.status(500).json({ error: "Server Configuration Error: Target submission URL missing." });
  }

  try {
    let params;

    if (typeof req.body === "string") {
      params = new URLSearchParams(req.body);
    } else if (req.body && typeof req.body === "object") {
      params = new URLSearchParams();
      for (const [key, value] of Object.entries(req.body)) {
        params.append(key, typeof value === "object" ? JSON.stringify(value) : value);
      }
    } else {
      params = new URLSearchParams();
    }

    const response = await fetch(submitUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: params.toString()
    });

    const responseText = await response.text();
    return res.status(response.status).send(responseText);
  } catch (error) {
    console.error("Proxy execution error:", error);
    return res.status(500).json({ error: "Failed to forward request to target endpoint." });
  }
}
