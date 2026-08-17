export default async function handler(req, res) {
  // CORS & Cache Headers
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Cache-Control", "s-maxage=300, stale-while-revalidate=600");

  // Preflight Handling
  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  // Method Enforcement
  if (req.method !== "GET") {
    return res.status(405).json({ error: "Method Not Allowed. Only GET requests are accepted." });
  }

  const { program } = req.query;
  const programUpper = (program || "").toUpperCase();

  const urlMap = {
    ESL: process.env.ESL_URL,
    HSE: process.env.HSE_URL,
    CPW: process.env.CPW_URL
  };

  const targetUrl = urlMap[programUpper];

  if (!targetUrl) {
    return res.status(400).json({ error: "Invalid or missing program parameter. Must be ESL, HSE, or CPW." });
  }

  try {
    const response = await fetch(targetUrl);
    const data = await response.json();
    return res.status(200).json(data);
  } catch (error) {
    console.error(`Failed to fetch schedule data for program ${programUpper}:`, error);
    return res.status(500).json({ error: "Failed to retrieve schedule data from backend database." });
  }
}
