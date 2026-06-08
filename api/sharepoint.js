const TENANT_ID = "5a52a73d-c443-4d27-9ae2-9165b37b4b8c";
const CLIENT_ID = "fdfa60f4-1061-4ca1-a6f1-c61bc448d387";
const CLIENT_SECRET = "9mm8Q~n09MLvUAYRAHcn28gBgrG5cm~Xu5.VWcpX";
const SITE_HOST = "alineinternational.sharepoint.com";
const SITE_PATH = "/sites/AlineCoreInfo";

async function getToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
  });
  console.log("[getToken] Fetching token...");
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const data = await resp.json();
  if (!data.access_token) {
    console.error("[getToken] FAILED:", JSON.stringify(data));
    throw new Error("Token failed: " + (data.error_description || data.error || JSON.stringify(data)));
  }
  console.log("[getToken] OK");
  return data.access_token;
}

async function getSiteId(token) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_HOST}:${SITE_PATH}`;
  console.log("[getSiteId] Fetching site...");
  const resp = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const data = await resp.json();
  if (!data.id) {
    console.error("[getSiteId] FAILED:", JSON.stringify(data));
    throw new Error("Site not found: " + JSON.stringify(data));
  }
  console.log("[getSiteId] OK:", data.id);
  return data.id;
}

async function getListId(token, siteId, listName) {
  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$filter=displayName eq '${listName}'`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const data = await resp.json();
  const id = data.value?.[0]?.id;
  console.log(`[getListId] ${listName}:`, id || "NOT FOUND");
  return id;
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  const { action } = req.query;
  console.log(`[handler] action=${action} method=${req.method}`);

  try {
    const token = await getToken();
    const siteId = await getSiteId(token);

    // ── LOAD ───────────────────────────────────────────────────────────────
    if (action === "load") {
      const { refCode } = req.query;
      console.log("[load] refCode:", refCode);
      const appListId = await getListId(token, siteId, "WorkBC Applications");
      const appResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${appListId}/items?$filter=fields/ReferenceCode eq '${refCode}'&$expand=fields`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const appData = await appResp.json();
      const f = appData.value?.[0]?.fields;
      if (!f) {
        console.log("[load] Not found");
        return res.status(404).json({ error: "Not found" });
      }

      const jobListId = await getListId(token, siteId, "WorkBC Jobs");
      const jobResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${jobListId}/items?$filter=fields/ReferenceCode eq '${refCode}'&$expand=fields`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const jobData = await jobResp.json();
      const jobs = (jobData.value || []).map(item => ({
        id: item.id,
        positionTitle: item.fields.PositionTitle || "",
        numPositions: item.fields.NumPositions?.toString() || "",
        startDate: item.fields.StartDate?.split("T")[0] || "",
        hoursPerWeek: item.fields.HoursPerWeek?.toString() || "",
        hourlyWage: item.fields.HourlyWage?.toString() || "",
        duties: item.fields.Duties || "",
        skillsRequired: item.fields.SkillsRequired || "",
        trainingProvided: item.fields.TrainingProvided || "",
        status: item.fields.JobStatus || "Draft",
        createdAt: new Date(item.createdDateTime).toLocaleDateString(),
      }));

      return res.json({
        businessInfo: {
          orgName: f.OrgName || "", craNumber: f.CRANumber || "",
          address: f.Address || "", city: f.City || "",
          province: f.Province || "", postalCode: f.PostalCode || "",
          telephone: f.Telephone || "", fax: f.Fax || "",
          email: f.Email || "", sectorType: f.SectorType || "",
          industryType: f.IndustryType || "",
          numEmployees: f.NumEmployees?.toString() || "",
          cewsParticipant: f.CEWSParticipant || "",
          displacesEmployees: f.DisplacesEmployees || "",
          labourStoppage: f.LabourStoppage || "",
          unionConcurrence: f.UnionConcurrence || "",
          thirdPartyLiability: f.ThirdPartyLiability || "",
          otherWageSubsidy: f.OtherWageSubsidy || "",
          wageSubsidyCoverage: f.WageSubsidyCoverage || "",
          workplaces: [],
        },
        documents: {},
        jobs,
      });
    }

    // ── SAVE ───────────────────────────────────────────────────────────────
    if (action === "save" && req.method === "POST") {
      const { refCode, businessInfo, submissionStatus } = req.body;
      console.log("[save] refCode:", refCode);
      const appListId = await getListId(token, siteId, "WorkBC Applications");

      const checkResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${appListId}/items?$filter=fields/ReferenceCode eq '${refCode}'&$expand=fields`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const checkData = await checkResp.json();
      const existing = checkData.value?.[0];

      const fields = {
        ReferenceCode: refCode,
        OrgName: businessInfo.orgName || "",
        CRANumber: businessInfo.craNumber || "",
        Address: businessInfo.address || "",
        City: businessInfo.city || "",
        Province: businessInfo.province || "",
        PostalCode: businessInfo.postalCode || "",
        Telephone: businessInfo.telephone || "",
        Fax: businessInfo.fax || "",
        Email: businessInfo.email || "",
        SectorType: businessInfo.sectorType || "",
        IndustryType: businessInfo.industryType || "",
        NumEmployees: parseFloat(businessInfo.numEmployees) || null,
        CEWSParticipant: businessInfo.cewsParticipant || "",
        DisplacesEmployees: businessInfo.displacesEmployees || "",
        LabourStoppage: businessInfo.labourStoppage || "",
        UnionConcurrence: businessInfo.unionConcurrence || "",
        ThirdPartyLiability: businessInfo.thirdPartyLiability || "",
        OtherWageSubsidy: businessInfo.otherWageSubsidy || "",
        WageSubsidyCoverage: businessInfo.wageSubsidyCoverage || "",
        SubmissionStatus: submissionStatus || "Draft",
        Title: refCode,
      };

      if (existing) {
        console.log("[save] Updating existing item:", existing.id);
        const patchResp = await fetch(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${appListId}/items/${existing.id}/fields`,
          { method: "PATCH", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }, body: JSON.stringify(fields) }
        );
        if (!patchResp.ok) {
          const err = await patchResp.json();
          console.error("[save] PATCH failed:", JSON.stringify(err));
          throw new Error("PATCH failed: " + JSON.stringify(err));
        }
      } else {
        console.log("[save] Creating new item");
        const postResp = await fetch(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${appListId}/items`,
          { method: "POST", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }, body: JSON.stringify({ fields }) }
        );
        if (!postResp.ok) {
          const err = await postResp.json();
          console.error("[save] POST failed:", JSON.stringify(err));
          throw new Error("POST failed: " + JSON.stringify(err));
        }
      }
      console.log("[save] OK");
      return res.json({ ok: true });
    }

    // ── SAVE JOBS ──────────────────────────────────────────────────────────
    if (action === "saveJobs" && req.method === "POST") {
      const { refCode, orgName, jobs } = req.body;
      console.log("[saveJobs] refCode:", refCode, "jobs:", jobs?.length);
      const jobListId = await getListId(token, siteId, "WorkBC Jobs");

      const existingResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${jobListId}/items?$filter=fields/ReferenceCode eq '${refCode}'&$expand=fields`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const existingData = await existingResp.json();
      for (const item of existingData.value || []) {
        await fetch(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${jobListId}/items/${item.id}`,
          { method: "DELETE", headers: { Authorization: `Bearer ${token}` } }
        );
      }
      for (const job of jobs || []) {
        await fetch(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${jobListId}/items`,
          {
            method: "POST",
            headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({ fields: {
              ReferenceCode: refCode, OrgName: orgName || "",
              PositionTitle: job.positionTitle || "",
              NumPositions: parseFloat(job.numPositions) || null,
              StartDate: job.startDate || null,
              HoursPerWeek: parseFloat(job.hoursPerWeek) || null,
              HourlyWage: parseFloat(job.hourlyWage) || null,
              Duties: job.duties || "", SkillsRequired: job.skillsRequired || "",
              TrainingProvided: job.trainingProvided || "",
              JobStatus: job.status || "Draft",
              Title: `${refCode} — ${job.positionTitle || "Untitled"}`,
            }}),
          }
        );
      }
      console.log("[saveJobs] OK");
      return res.json({ ok: true });
    }

    // ── SEND EMAIL ─────────────────────────────────────────────────────────
    if (action === "sendEmail" && req.method === "POST") {
      const { clientEmail, clientOrgName, refCode } = req.body;
      console.log("[sendEmail] to:", clientEmail, "refCode:", refCode);

      const emailBody = `
        <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
          <div style="background:#0D1F3C;padding:24px;border-radius:8px 8px 0 0;">
            <h1 style="color:#C9A84C;margin:0;font-size:22px;">Aline International</h1>
            <p style="color:#fff;margin:4px 0 0;font-size:14px;">WorkBC Wage Subsidy Portal</p>
          </div>
          <div style="background:#f8f5ef;padding:32px;border-radius:0 0 8px 8px;border:1px solid #e5e7eb;">
            <p style="color:#0D1F3C;font-size:16px;">Dear ${clientOrgName},</p>
            <p style="color:#374151;font-size:15px;">
              Thank you for starting your WorkBC Wage Subsidy application with Aline International.
              Your unique reference code is below — please save it to return and complete your application.
            </p>
            <div style="background:#0D1F3C;border-radius:10px;padding:20px;text-align:center;margin:24px 0;">
              <p style="color:#C9A84C;font-size:13px;margin:0 0 8px;letter-spacing:1px;">YOUR REFERENCE CODE</p>
              <p style="color:#fff;font-size:32px;font-weight:700;font-family:monospace;margin:0;letter-spacing:4px;">${refCode}</p>
            </div>
            <p style="color:#374151;font-size:14px;">Use this code anytime to resume your application.</p>
            <hr style="border:none;border-top:1px solid #e5e7eb;margin:24px 0;" />
            <p style="color:#6b7280;font-size:13px;margin:0;">
              Aline International &nbsp;|&nbsp; (866) 592-3243 &nbsp;|&nbsp; www.aline-business.ca<br/>
              #460-3820 Cessna Dr, Richmond, BC V7B 0A2
            </p>
          </div>
        </div>`;

      const emailResp = await fetch(
        `https://graph.microsoft.com/v1.0/users/info@aline-business.ca/sendMail`,
        {
          method: "POST",
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify({
            message: {
              subject: `Your WorkBC Application Reference Code — ${refCode}`,
              body: { contentType: "HTML", content: emailBody },
              toRecipients: [{ emailAddress: { address: clientEmail } }],
              ccRecipients: [{ emailAddress: { address: "alex.chiang@aline-business.ca" } }],
            },
            saveToSentItems: true,
          }),
        }
      );

      if (!emailResp.ok) {
        const err = await emailResp.json();
        console.error("[sendEmail] FAILED:", JSON.stringify(err));
        throw new Error("Email failed: " + JSON.stringify(err));
      }
      console.log("[sendEmail] OK");
      return res.json({ ok: true });
    }

    // ── UPLOAD ─────────────────────────────────────────────────────────────
    if (action === "upload" && req.method === "POST") {
      const { refCode, fileName } = req.query;
      console.log("[upload] refCode:", refCode, "file:", fileName);
      const chunks = [];
      for await (const chunk of req) chunks.push(chunk);
      const fileBuffer = Buffer.concat(chunks);

      const uploadResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/WorkBC Documents/${refCode}/${fileName}:/content`,
        {
          method: "PUT",
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/octet-stream" },
          body: fileBuffer,
        }
      );
      const uploadData = await uploadResp.json();
      if (!uploadData.webUrl) {
        console.error("[upload] FAILED:", JSON.stringify(uploadData));
        throw new Error("Upload failed: " + JSON.stringify(uploadData));
      }
      console.log("[upload] OK:", uploadData.webUrl);
      return res.json({ ok: true, url: uploadData.webUrl });
    }

    return res.status(400).json({ error: "Unknown action: " + action });

  } catch (e) {
    console.error("[handler] ERROR:", e.message);
    return res.status(500).json({ error: e.message });
  }
}
