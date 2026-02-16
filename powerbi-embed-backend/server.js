const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const axios = require("axios");
const { ConfidentialClientApplication } = require("@azure/msal-node");

const app = express();
app.use(
  cors({
    origin: "https://digilabsolutions0.sharepoint.com",
    credentials: true,
  })
);
const PORT = process.env.PORT || 3000;
const TENANT_ID = process.env.TENANT_ID;

// Service Principal for SharePoint/Graph
const SP_APP_ID = process.env.SP_APP_ID;
const SP_APP_SECRET = process.env.SP_APP_SECRET;

// Service Principal for Power BI
const PBI_APP_ID = process.env.PBI_APP_ID;
const PBI_APP_SECRET = process.env.PBI_APP_SECRET;

const msalSP = new ConfidentialClientApplication({
  auth: {
    clientId: SP_APP_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: SP_APP_SECRET,
  },
});

const msalPBI = new ConfidentialClientApplication({
  auth: {
    clientId: PBI_APP_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: PBI_APP_SECRET,
  },
});

async function getGraphToken() {
  const res = await msalSP.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return res.accessToken;
}

async function getPbiToken() {
  const res = await msalPBI.acquireTokenByClientCredential({
    scopes: ["https://analysis.windows.net/powerbi/api/.default"],
  });
  return res.accessToken;
}

let siteId;
async function initSiteId() {
  try {
    const token = await getGraphToken();
    const resp = await axios.get(
      `https://graph.microsoft.com/v1.0/sites?search="PowerBiConfiguration"`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (resp.data.value && resp.data.value.length > 0) {
      siteId = resp.data.value[0].id;
      console.log("Found SharePoint Site ID:", siteId);
    } else {
      console.error(
        "Critical: Could not find SharePoint site 'PowerBiConfiguration'"
      );
    }
  } catch (err) {
    console.error("Initialization Error:", err.message);
  }
}

initSiteId();

async function getRoleForUserFromSP(userEmail, reportId) {
  const token = await getGraphToken();
  const listName = "PowerBISecurityRoles";
  const resp = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${encodeURIComponent(
      listName
    )}/items?$expand=fields($select=Title,ReportId,User)&$select=id,fields`,
    /*`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${encodeURIComponent(
      listName
    )}/items` + `?$expand=fields($select=Title,ReportId,User)`*/ {
      headers: { Authorization: `Bearer ${token}` },
    }
  );
  console.log("--- DEBUG: FULL LIST DATA ---");
  console.log(JSON.stringify(resp.data.value, null, 2));
  console.log("-----------------------------");

  const lowerEmail = userEmail.toLowerCase();
  for (const item of resp.data.value) {
    const { Title, ReportId, User } = item.fields;

    const targetEmail = userEmail.toLowerCase().trim();
    const listReportId = (ReportId || "").trim();

    if (listReportId === reportId.trim() && User) {
      const userArray = Array.isArray(User) ? User : [User];

      for (let u of userArray) {
        const entryValue = typeof u === "object" ? u.Email || u.email || "" : u;
        const cleanEntryValue = String(entryValue).toLowerCase().trim();

        console.log(`Comparing: '${cleanEntryValue}' with '${targetEmail}'`);

        if (cleanEntryValue === targetEmail) {
          console.log(`✅ Matched! Role: ${Title}`);
          return Title;
        }
      }
    }
  }

  console.log(`No role found for ${userEmail} on report ${reportId}`);
  return undefined;
}

app.get("/api/embed-info", async (req, res) => {
  try {
    const {
      userEmail,
      reportId,
      datasetId,
      reportWorkspaceId,
      datasetWorkspaceId,
      hasRLS,
    } = req.query;

    if (
      !userEmail ||
      !reportId ||
      !datasetId ||
      !reportWorkspaceId ||
      !datasetWorkspaceId
    ) {
      return res
        .status(400)
        .json({ error: "Missing required query parameters" });
    }

    let roles = [];
    if (hasRLS === "yes") {
      const role = await getRoleForUserFromSP(userEmail, reportId);
      if (!role) {
        return res
          .status(403)
          .json({ error: "User not authorized in SharePoint Security List" });
      }
      roles = [role];
    }

    const pbiToken = await getPbiToken();

    const reportRes = await axios.get(
      `https://api.powerbi.com/v1.0/myorg/groups/${reportWorkspaceId}/reports/${reportId}`,
      { headers: { Authorization: `Bearer ${pbiToken}` } }
    );

    const targetWorkspaces = [{ id: reportWorkspaceId }];
    if (datasetWorkspaceId !== reportWorkspaceId) {
      targetWorkspaces.push({ id: datasetWorkspaceId });
    }

    const tokenRequest = {
      datasets: [{ id: datasetId }],
      reports: [{ id: reportId }],
      targetWorkspaces: targetWorkspaces,
      accessLevel: "view",
    };

    if (roles.length > 0) {
      tokenRequest.identities = [
        {
          username: userEmail,
          roles: roles,
          datasets: [datasetId],
        },
      ];
    }

    const embedRes = await axios.post(
      `https://api.powerbi.com/v1.0/myorg/GenerateToken`,
      tokenRequest,
      { headers: { Authorization: `Bearer ${pbiToken}` } }
    );

    res.json({
      embedUrl: reportRes.data.embedUrl,
      embedToken: embedRes.data.token,
      reportId: reportRes.data.id,
      usedRoles: roles,
    });
  } catch (err) {
    const pbiError =
      err.response?.data?.error || err.response?.data || err.message;
    console.error("Power BI API Failure:", JSON.stringify(pbiError, null, 2));

    res.status(500).json({
      error: "Failed to generate Power BI Embed Token",
      details: pbiError,
    });
  }
});

app.listen(PORT, () => console.log(`API listening on port ${PORT}`));

/*
const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const axios = require("axios");
const { ConfidentialClientApplication } = require("@azure/msal-node");

const app = express();
app.use(cors({
  origin: "https://digilabsolutions0.sharepoint.com",
  credentials: true
}));
app.use(bodyParser.json());

const PORT = process.env.PORT || 3000;
const TENANT_ID = process.env.TENANT_ID;

const SP_APP_ID = process.env.SP_APP_ID;
const SP_APP_SECRET = process.env.SP_APP_SECRET;

const PBI_APP_ID = process.env.PBI_APP_ID;
const PBI_APP_SECRET = process.env.PBI_APP_SECRET;

const msalSP = new ConfidentialClientApplication({
  auth: {
    clientId: SP_APP_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: SP_APP_SECRET,
  },
});

const msalPBI = new ConfidentialClientApplication({
  auth: {
    clientId: PBI_APP_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: PBI_APP_SECRET,
  },
});

async function getGraphToken() {
  const res = await msalSP.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return res.accessToken;
}

async function getPbiToken() {
  const res = await msalPBI.acquireTokenByClientCredential({
    scopes: ["https://analysis.windows.net/powerbi/api/.default"],
  });
  return res.accessToken;
}

let siteId;
async function initSiteId() {
  const token = await getGraphToken();
  const resp = await axios.get(
    `https://graph.microsoft.com/v1.0/sites?search="PowerBiConfiguration"`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  console.log("Search results:", resp.data.value);
  if (!resp.data.value.length) {
    throw new Error("Site not found via search");
  }
  siteId = resp.data.value[0].id;
  console.log("siteId searched:", siteId);
}

initSiteId().catch((e) => console.error("Failed to init siteId:", e));

async function getRoleForUserFromSP(userEmail, reportId) {
  const token = await getGraphToken();
  const listName = "PowerBISecurityRoles";
  const resp = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${encodeURIComponent(
      listName
    )}/items` + `?$expand=fields($select=Title,ReportId,User)`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  const lowerEmail = userEmail.toLowerCase();
  for (const item of resp.data.value) {
    const { Title, ReportId, User } = item.fields;
    if (ReportId === reportId && User) {
      const users = Array.isArray(User) ? User : [User];
      for (const u of users) {
        const email = (u.Email || u.email || "").toLowerCase();
        if (email === lowerEmail) {
          console.log(`Matched ${email}, role: ${Title}`);
          return Title;
        }
      }
    }
  }

  console.log(`No role found for ${userEmail} on report ${reportId}`);
  return undefined;
}

app.get("/api/embed-info", async (req, res) => {
  try {
    const {
      userEmail,
      reportId,
      datasetId,
      reportWorkspaceId,
      datasetWorkspaceId,
      hasRLS,
    } = req.query;

    if (
      !userEmail ||
      !reportId ||
      !datasetId ||
      !reportWorkspaceId ||
      !datasetWorkspaceId
    ) {
      return res
        .status(400)
        .json({ error: "Missing required query parameters" });
    }

    let roles = ["Manager", "Employee"];
    
    if (hasRLS === "yes") {
      const role = await getRoleForUserFromSP(userEmail, reportId);
      if (role) {
        roles = [role];
      } else {
        return res
          .status(403)
          .json({ error: "User not authorized for this report" });
      }
    } else {
      console.log(`RLS not enabled, using default roles: ${roles}`);
    }

    const pbiToken = await getPbiToken();
    const reportRes = await axios.get(
      `https://api.powerbi.com/v1.0/myorg/groups/${reportWorkspaceId}/reports/${reportId}`,
      { headers: { Authorization: `Bearer ${pbiToken}` } }
    );
    const targetWorkspaces = [{ id: reportWorkspaceId }];

    if (datasetWorkspaceId !== reportWorkspaceId) {
      targetWorkspaces.push({ id: datasetWorkspaceId });
    }

    const embedRes = await axios.post(
      `https://api.powerbi.com/v1.0/myorg/GenerateToken`,
      {
        datasets: [{ id: datasetId }],
        reports: [{ id: reportId }],
        targetWorkspaces: targetWorkspaces,
        accessLevel: "view",
        identities: [{ username: userEmail, roles, datasets: [datasetId] }],
      },
      { headers: { Authorization: `Bearer ${pbiToken}` } }
    );

    res.json({
      embedUrl: reportRes.data.embedUrl,
      embedToken: embedRes.data.token,
      reportId: reportRes.data.id,
      usedRoles: roles,
    });
  } catch (err) {
    console.error("API error:", err.response?.data || err);
    res.status(500).json({
      error: "Error generating embed token",
      details: err.response?.data || err,
    });
  }
});

app.listen(PORT, () => console.log(`API listening on port ${PORT}`));
*/
