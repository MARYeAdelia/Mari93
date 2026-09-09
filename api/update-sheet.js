export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const SHEET_ID = "1Mw46A8j0c5-6VyOY8K7bEFPTnEGLh6nG-lL5jEXe2G0";
  const GID      = "2073814116";

  try {
    // Parse service account credentials from env var
    const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
    const { updates } = req.body; // [{ row: N, col: N, value: "..." }, ...]

    if (!updates?.length) return res.status(400).json({ error: "No updates provided" });

    // Get access token via JWT
    const token = await getAccessToken(creds);

    // Convert GID to sheet name via Sheets API
    const metaRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}?fields=sheets.properties`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const meta = await metaRes.json();
    const sheet = meta.sheets?.find(s => String(s.properties.sheetId) === GID);
    const sheetName = sheet?.properties?.title || "Gerencial";

    // Build batch update request
    const data = updates.map(({ row, col, value }) => ({
      range: `${sheetName}!${colLetter(col)}${row}`,
      values: [[value]],
    }));

    const updateRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values:batchUpdate`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ valueInputOption: "USER_ENTERED", data }),
      }
    );

    const result = await updateRes.json();
    if (!updateRes.ok) throw new Error(result.error?.message || "Sheets API error");

    res.status(200).json({ ok: true, updated: updates.length });
  } catch (err) {
    console.error("update-sheet error:", err);
    res.status(500).json({ error: err.message });
  }
}

// Convert column index (0-based) to letter (A, B, C...)
const colLetter = n => {
  let s = "";
  n++;
  while (n > 0) { s = String.fromCharCode(65 + (n-1) % 26) + s; n = Math.floor((n-1) / 26); }
  return s;
};

// Generate JWT and get access token for Google API
async function getAccessToken(creds) {
  const now = Math.floor(Date.now() / 1000);
  const payload = {
    iss: creds.client_email,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600,
  };

  // Build JWT
  const header = btoa(JSON.stringify({ alg: "RS256", typ: "JWT" })).replace(/=/g,"").replace(/\+/g,"-").replace(/\//g,"_");
  const body   = btoa(JSON.stringify(payload)).replace(/=/g,"").replace(/\+/g,"-").replace(/\//g,"_");
  const unsigned = `${header}.${body}`;

  // Sign with private key using Web Crypto
  const pemKey = creds.private_key
    .replace("-----BEGIN PRIVATE KEY-----", "")
    .replace("-----END PRIVATE KEY-----", "")
    .replace(/\s/g, "");
  const keyBytes = Uint8Array.from(atob(pemKey), c => c.charCodeAt(0));
  const cryptoKey = await crypto.subtle.importKey(
    "pkcs8", keyBytes.buffer,
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false, ["sign"]
  );
  const sigBytes = await crypto.subtle.sign(
    "RSASSA-PKCS1-v1_5", cryptoKey,
    new TextEncoder().encode(unsigned)
  );
  const sig = btoa(String.fromCharCode(...new Uint8Array(sigBytes)))
    .replace(/=/g,"").replace(/\+/g,"-").replace(/\//g,"_");

  const jwt = `${unsigned}.${sig}`;

  // Exchange JWT for access token
  const tokenRes = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`,
  });
  const tokenData = await tokenRes.json();
  if (!tokenData.access_token) throw new Error("Failed to get access token: " + JSON.stringify(tokenData));
  return tokenData.access_token;
}
