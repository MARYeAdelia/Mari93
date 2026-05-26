export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-api-key");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const apiKey = req.headers["x-api-key"];
  if (apiKey !== process.env.API_KEY) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  try {
    const dados = req.body;
    const content = Buffer.from(JSON.stringify(dados)).toString("base64");

    const getRes = await fetch(
      "https://api.github.com/repos/MARYeAdelia/Mari93/contents/dados_farmer.json",
      {
        headers: {
          Authorization: `token ${process.env.GITHUB_TOKEN}`,
          Accept: "application/vnd.github.v3+json",
        },
      }
    );
    const getJson = await getRes.json();
    const sha = getJson.sha;

    const putRes = await fetch(
      "https://api.github.com/repos/MARYeAdelia/Mari93/contents/dados_farmer.json",
      {
        method: "PUT",
        headers: {
          Authorization: `token ${process.env.GITHUB_TOKEN}`,
          Accept: "application/vnd.github.v3+json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          message: "Atualizar dados farmer via Power Automate",
          content,
          sha,
        }),
      }
    );

    if (!putRes.ok) {
      const err = await putRes.text();
      return res.status(500).json({ error: err });
    }

    return res.status(200).json({ success: true, updated: new Date().toISOString() });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
