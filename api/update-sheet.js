export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const SHEET_ID = "1Mw46A8j0c5-6VyOY8K7bEFPTnEGLh6nG-lL5jEXe2G0";
  const GID      = "2073814116";

  try {
    const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
    const { updates } = req.body;
    if (!updates?.length) return res.status(400).json({ error: "No updates" });

    const token = await getAccessToken(creds);

    // Get sheet metadata to find column positions and sheet name
    const metaRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}?fields=sheets.properties`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const meta = await metaRes.json();
    const sheet = meta.sheets?.find(s => String(s.properties.sheetId) === GID);
    const sheetName = sheet?.properties?.title;
    if (!sheetName) throw new Error(`Sheet with GID ${GID} not found`);

    // Get header row to find column positions by name
    const headerRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(sheetName)}!1:1`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const headerData = await headerRes.json();
    const headers = headerData.values?.[0] || [];

    // Build batch update
    const data = updates.map(({ row, colName, value }) => {
      const colIdx = headers.findIndex(h =>
        h.trim().toLowerCase() === colName.trim().toLowerCase()
      );
      if (colIdx < 0) throw new Error(`Column "${colName}" not found in headers: ${headers.join(", ")}`);
      const colLetter = idxToCol(colIdx);
      return {
        range: `${sheetName}!${colLetter}${row}`,
        values: [[value]],
      };
    });

    const updateRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values:batchUpdate`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ valueInputOption: "USER_ENTERED", data }),
      }
    );
    const result = await updateRes.json();
    if (!updateRes.ok) throw new Error(result.error?.message || JSON.stringify(result));

    res.status(200).json({ ok: true, updated: updates.length, sheetName });
  } catch (err) {
    console.error("update-sheet error:", err.message);
    res.status(500).json({ error: err.message });
  }
}

const idxToCol = n => {
  let s = ""; n++;
  while (n > 0) { s = String.fromCharCode(64 + (n - 1) % 26 + 1) + s; n = Math.floor((n - 1) / 26); }
  return s;
};

async function getAccessToken(creds) {
  const now = Math.floor(Date.now() / 1000);
  const payload = {
    iss: creds.client_email,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: "https://oauth2.googleapis.com/token",
    iat: now, exp: now + 3600,
  };
  const enc = s => btoa(JSON.stringify(s)).replace(/=/g,"").replace(/\+/g,"-").replace(/\//g,"_");
  const unsigned = `${enc({ alg:"RS256", typ:"JWT" })}.${enc(payload)}`;

  const pem = creds.private_key.replace(/-----.*?-----/g,"").replace(/\s/g,"");
  const keyBytes = Uint8Array.from(atob(pem), c => c.charCodeAt(0));
  const key = await crypto.subtle.importKey(
    "pkcs8", keyBytes.buffer,
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["sign"]
  );
  const sig = await crypto.subtle.sign("RSASSA-PKCS1-v1_5", key, new TextEncoder().encode(unsigned));
  const sigB64 = btoa(String.fromCharCode(...new Uint8Array(sig)))
    .replace(/=/g,"").replace(/\+/g,"-").replace(/\//g,"_");

  const tokenRes = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${unsigned}.${sigB64}`,
  });
  const td = await tokenRes.json();
  if (!td.access_token) throw new Error("Token error: " + JSON.stringify(td));
  return td.access_token;
}
