// Returns the rows of the Deacon Continuing Education Log (Microsoft Forms response workbook)
// for /tracker. Reads the workbook through its "Anyone with the link" sharing link, so no sign-in
// or secret is needed. To point at a different link, set CE_SHARE_URL in Vercel.
const XLSX = require("xlsx");

const SHARE_URL = process.env.CE_SHARE_URL ||
  "https://mobilearchdiocese-my.sharepoint.com/:x:/g/personal/apitts_mobarch_org1/IQAPXz0Az_CtRqCodcNNzu9XAdPyvEoV9rlfmU545ZoBz_I?e=gSLVCs";

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  try {
    const url = new URL(SHARE_URL);
    url.searchParams.set("download", "1");
    const r = await fetch(url, { redirect: "follow" });
    if (!r.ok) throw new Error(`SharePoint returned ${r.status} for the sharing link.`);
    const buf = Buffer.from(await r.arrayBuffer());
    if (buf[0] !== 0x50 || buf[1] !== 0x4b) {
      throw new Error("The sharing link did not return an Excel file. It may have expired or no longer allow anyone with the link to view.");
    }
    const wb = XLSX.read(buf, { type: "buffer" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const values = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
    res.status(200).json({ values });
  } catch (e) {
    res.status(502).json({ error: e.message || "The response log could not be read." });
  }
};
