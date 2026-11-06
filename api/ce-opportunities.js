// Returns the rows of the "Continuing Education Opportunities" workbook for /coned.
// Reads the workbook through its "Anyone with the link" sharing link, so no sign-in or secret is needed.
// To point at a different link, set CE_OPPS_SHARE_URL in Vercel. Add ?debug=1 to see each redirect hop.
// Expected columns (row 1): Title | Description | Dates Offered | Category | Credit Hours
const XLSX = require("xlsx");

const SHARE_URL = process.env.CE_OPPS_SHARE_URL ||
  "https://mobilearchdiocese-my.sharepoint.com/:x:/g/personal/apitts_mobarch_org1/IQCDN39mVOlVQapxTM1pl_-8AYdJe4EIWuyRrGChYVCQwdA?e=GwuVaF";
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36";

// Follow redirects by hand so cookies SharePoint sets on the way (FedAuth for anonymous links) are sent on.
async function fetchWithCookies(startUrl, jar, hops) {
  let url = startUrl;
  for (let i = 0; i < 12; i++) {
    const cookie = [...jar].map(([k, v]) => `${k}=${v}`).join("; ");
    const r = await fetch(url, { redirect: "manual", headers: { "User-Agent": UA, ...(cookie ? { Cookie: cookie } : {}) } });
    const set = typeof r.headers.getSetCookie === "function" ? r.headers.getSetCookie() : [];
    for (const c of set) { const [pair] = c.split(";"); const at = pair.indexOf("="); if (at > 0) jar.set(pair.slice(0, at).trim(), pair.slice(at + 1).trim()); }
    const u = new URL(url);
    hops.push({ status: r.status, path: u.host + u.pathname, cookies: set.map(c => c.split("=")[0]) });
    if (r.status >= 300 && r.status < 400 && r.headers.get("location")) { url = new URL(r.headers.get("location"), url).toString(); continue; }
    return r;
  }
  throw new Error("Too many redirects from SharePoint.");
}

module.exports = async (req, res) => {
  // Short CDN cache: the list changes rarely, and this keeps page loads fast.
  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=600");
  const debug = /[?&]debug=1/.test(req.url || "");
  const hops = [], jar = new Map();
  try {
    await fetchWithCookies(SHARE_URL, jar, hops);
    const dl = new URL(SHARE_URL);
    dl.searchParams.set("download", "1");
    const r = await fetchWithCookies(dl.toString(), jar, hops);
    if (!r.ok) throw new Error(`SharePoint returned ${r.status} for the sharing link.`);
    const buf = Buffer.from(await r.arrayBuffer());
    if (buf[0] !== 0x50 || buf[1] !== 0x4b) {
      throw new Error("The sharing link did not return an Excel file. It may have expired or no longer allow anyone with the link to view.");
    }
    const wb = XLSX.read(buf, { type: "buffer" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const values = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" })
      .filter(row => row.some(v => String(v).trim() !== ""));
    res.status(200).json(debug ? { rows: values.length, hops } : { values });
  } catch (e) {
    res.setHeader("Cache-Control", "no-store");
    res.status(502).json({ error: e.message || "The opportunities list could not be read.", ...(debug ? { hops } : {}) });
  }
};
