#!/usr/bin/env node

// ---------- Imports ----------
import 'dotenv/config';
import express from "express";
import fetch from "node-fetch";
import { v4 as uuidv4 } from "uuid";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";

// ---------- Environment ----------
const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  PORT = process.env.PORT || "8787",
  BASE_PATH = "/mcp",
} = process.env;

console.log("ENV OK?", !!TENANT_ID, !!CLIENT_ID, !!CLIENT_SECRET);
if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
  console.error("Missing required env: TENANT_ID, CLIENT_ID, CLIENT_SECRET");
  process.exit(1);
}

// ---------- Power BI Auth (MSAL) ----------
// IMPORTANT: For Service Principal authentication with Power BI Admin APIs:
// 1. DO NOT add Power BI Service API permissions (Tenant.Read.All, etc.) in Azure AD
//    - Service principals inherit permissions from Power BI tenant settings only
//    - Azure AD permissions conflict with Admin API access
// 2. Required Power BI tenant settings (in Power BI Admin Portal):
//    - Developer Settings: "Allow service principals to use Power BI APIs"
//    - Admin API Settings: "Service principals can access read-only admin APIs"
// 3. Add the service principal to a security group specified in both settings above

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: CLIENT_SECRET,
  },
});

async function getAccessToken() {
  const tokenResponse = await cca.acquireTokenByClientCredential({
    scopes: ["https://analysis.windows.net/powerbi/api/.default"],
  });
  if (!tokenResponse?.accessToken) throw new Error("Failed to acquire Power BI token");
  
  return tokenResponse.accessToken;
}

// Decode JWT (header/payload only) for diagnostics
function decodeJwtNoVerify(jwt) {
  try {
    const [h, p] = jwt.split(".");
    const pad = s => s.replace(/-/g, "+").replace(/_/g, "/") + "===".slice((s.length + 3) % 4);
    const header = JSON.parse(Buffer.from(pad(h), "base64").toString("utf8"));
    const payload = JSON.parse(Buffer.from(pad(p), "base64").toString("utf8"));
    return { header, payload };
  } catch {
    return null;
  }
}

// Centralised PBI fetch with rich logging
async function pbiAdminFetch(pathAndQuery) {
  const token = await getAccessToken();
  const url = `https://api.powerbi.com/v1.0/myorg/admin${pathAndQuery}`;

  console.log("[PBI REQUEST]", url);

  let res, text;
  try {
    res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    text = await res.text();
  } catch (e) {
    console.error("[PBI NETWORK ERROR]", url, String(e));
    return { ok: false, status: 0, errorText: String(e) };
  }

  if (!res.ok) {
    console.error("[PBI ERROR]", res.status, url, text);
    return { ok: false, status: res.status, errorText: text };
  }

  let body = {};
  try { body = text ? JSON.parse(text) : {}; } catch { body = {}; }
  return { ok: true, status: res.status, body };
}

// Simple pager to count tenant-wide assets (datasets/reports)
async function pbiAdminCount(pathBaseWithScope, pageSize = 5000) {
  let skip = 0;
  let total = 0;

  while (true) {
    const sep = pathBaseWithScope.includes("?") ? "&" : "?";
    const path = `${pathBaseWithScope}${sep}$top=${pageSize}&$skip=${skip}`;
    const result = await pbiAdminFetch(path);

    if (!result.ok) {
      return { ok: false, status: result.status, errorText: result.errorText };
    }

    const items = Array.isArray(result.body?.value) ? result.body.value : [];
    total += items.length;
    if (items.length < pageSize) break;
    skip += pageSize;
  }

  return { ok: true, total };
}

// ---------- MCP Server (tools) ----------
function buildMcpServer() {
  const server = new McpServer({
    name: "powerbi-admin-mcp-remote",
    version: "1.3.0",
  });

  const toErrorContent = (prefix, res) => ([
    { type: "text", text: `${prefix}: HTTP ${res.status}\n${res.errorText ?? "No body"}` }
  ]);

  // Diagnostics: quick admin ping
  server.tool(
    "diagnose_admin_ping",
    {
      description: "Checks if the service principal can call /admin/capacities (tenant-level).",
      inputSchema: { type: "object", properties: {} },
    },
    async () => {
      const r = await pbiAdminFetch(`/capacities?$top=1`);
      if (!r.ok) return { content: toErrorContent("Admin ping failed", r) };
      return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
    }
  );

  // Diagnostics: show token claims (aud/app/tenant/roles)
  server.tool(
    "diagnose_token_claims",
    {
      description: "Decodes the service principal token claims to verify audience/app/tenant.",
      inputSchema: { type: "object", properties: {} },
    },
    async () => {
      try {
        const token = await getAccessToken();
        const decoded = decodeJwtNoVerify(token);
        if (!decoded) return { content: [{ type: "text", text: "Unable to decode token" }] };
        const { header, payload } = decoded;
        const subset = {
          header,
          payload: {
            aud: payload.aud,
            appid: payload.appid,
            tid: payload.tid,
            roles: payload.roles,
            iss: payload.iss,
            iat: payload.iat,
            nbf: payload.nbf,
            exp: payload.exp,
          }
        };
        return { content: [{ type: "text", text: JSON.stringify(subset, null, 2) }] };
      } catch (e) {
        return { content: [{ type: "text", text: `Token error: ${String(e)}` }] };
      }
    }
  );

  // List workspaces
  server.tool(
    "list_admin_groups",
    {
      description: "List Power BI workspaces across the organisation (tenant-wide).",
      inputSchema: { type: "object", properties: { top: { type: "number" } } },
    },
    async (input) => {
      const top = input?.top ?? 100;
      const r = await pbiAdminFetch(`/groups?scope=Organization&$top=${top}`);
      if (!r.ok) return { content: toErrorContent("list_admin_groups failed", r) };
      return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
    }
  );

  // Count datasets
  server.tool(
    "count_admin_datasets",
    {
      description: "Return the total number of datasets across the organisation.",
      inputSchema: { type: "object", properties: {} },
    },
    async () => {
      const r = await pbiAdminCount(`/datasets?scope=Organization`);
      if (!r.ok) return { content: toErrorContent("count_admin_datasets failed", r) };
      return { content: [{ type: "text", text: JSON.stringify({ datasetsCount: r.total }, null, 2) }] };
    }
  );
// Get all reports across organization
server.tool(
  "get_admin_reports",
  {
    description: "Returns all reports across the organization with details. Supports pagination.",
    inputSchema: { 
      type: "object", 
      properties: { 
        top: { type: "number", description: "Number of results per page", default: 100 },
        skip: { type: "number", description: "Number of results to skip", default: 0 }
      }
    },
  },
  async (input) => {
    const { top = 100, skip = 0 } = input || {};
    const r = await pbiAdminFetch(`/reports?$top=${top}&$skip=${skip}`);
    if (!r.ok) return { content: toErrorContent("get_admin_reports failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);
  // Count reports
  server.tool(
    "count_admin_reports",
    {
      description: "Return the total number of reports across the organisation.",
      inputSchema: { type: "object", properties: {} },
    },
    async () => {
      const r = await pbiAdminCount(`/reports?scope=Organization`);
      if (!r.ok) return { content: toErrorContent("count_admin_reports failed", r) };
      return { content: [{ type: "text", text: JSON.stringify({ reportsCount: r.total }, null, 2) }] };
    }
  );

  // Combined counts
  server.tool(
    "count_admin_assets",
    {
      description: "Return both dataset and report totals across the organisation.",
      inputSchema: { type: "object", properties: {} },
    },
    async () => {
      const [d, r] = await Promise.all([
        pbiAdminCount(`/datasets?scope=Organization`),
        pbiAdminCount(`/reports?scope=Organization`),
      ]);
      if (!d.ok) return { content: toErrorContent("datasets part failed", d) };
      if (!r.ok) return { content: toErrorContent("reports part failed", r) };
      return { content: [{ type: "text", text: JSON.stringify({ datasetsCount: d.total, reportsCount: r.total }, null, 2) }] };
    }
  );

  // Add these tools to your buildMcpServer() function

// ---------- Capacity Health Tools ----------

// Get all capacities with details
server.tool(
  "get_admin_capacities",
  {
    description: "Returns list of all Premium capacities with detailed info (SKU, state, region).",
    inputSchema: { type: "object", properties: {} },
  },
  async () => {
    const r = await pbiAdminFetch(`/capacities`);
    if (!r.ok) return { content: toErrorContent("get_admin_capacities failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get refreshables for capacity (to identify refresh issues)
server.tool(
  "get_refreshables_for_capacity",
  {
    description: "Returns refreshables (datasets with schedules) for a specific capacity. Useful for identifying refresh health.",
    inputSchema: { 
      type: "object", 
      properties: { 
        capacityId: { type: "string", description: "The capacity ID (GUID)" },
        top: { type: "number", description: "Number of results to return", default: 100 }
      },
      required: ["capacityId"]
    },
  },
  async (input) => {
    const { capacityId, top = 100 } = input;
    const r = await pbiAdminFetch(`/capacities/${capacityId}/refreshables?$top=${top}`);
    if (!r.ok) return { content: toErrorContent("get_refreshables_for_capacity failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// ---------- Workspace (Group) Health Tools ----------

// Get specific workspace details
server.tool(
  "get_admin_workspace_details",
  {
    description: "Returns detailed information about a specific workspace including state, capacity, and type.",
    inputSchema: { 
      type: "object", 
      properties: { 
        groupId: { type: "string", description: "The workspace ID (GUID)" }
      },
      required: ["groupId"]
    },
  },
  async (input) => {
    const { groupId } = input;
    const r = await pbiAdminFetch(`/groups/${groupId}`);
    if (!r.ok) return { content: toErrorContent("get_admin_workspace_details failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get workspace users (permissions audit)
server.tool(
  "get_admin_workspace_users",
  {
    description: "Returns list of users with access to a workspace. Critical for security audits.",
    inputSchema: { 
      type: "object", 
      properties: { 
        groupId: { type: "string", description: "The workspace ID (GUID)" }
      },
      required: ["groupId"]
    },
  },
  async (input) => {
    const { groupId } = input;
    const r = await pbiAdminFetch(`/groups/${groupId}/users`);
    if (!r.ok) return { content: toErrorContent("get_admin_workspace_users failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get unused artifacts (important for cleanup)
server.tool(
  "get_unused_artifacts",
  {
    description: "Returns datasets, reports, and dashboards not used in 30 days for a workspace. Key for identifying stale content.",
    inputSchema: { 
      type: "object", 
      properties: { 
        groupId: { type: "string", description: "The workspace ID (GUID)" }
      },
      required: ["groupId"]
    },
  },
  async (input) => {
    const { groupId } = input;
    const r = await pbiAdminFetch(`/groups/${groupId}/unused`);
    if (!r.ok) return { content: toErrorContent("get_unused_artifacts failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// ---------- Dataset Health Tools ----------

// Get datasets in specific workspace
server.tool(
  "get_datasets_in_workspace",
  {
    description: "Returns all datasets in a specific workspace with refresh and connection details.",
    inputSchema: { 
      type: "object", 
      properties: { 
        groupId: { type: "string", description: "The workspace ID (GUID)" },
        top: { type: "number", description: "Number of results", default: 100 }
      },
      required: ["groupId"]
    },
  },
  async (input) => {
    const { groupId, top = 100 } = input;
    const r = await pbiAdminFetch(`/groups/${groupId}/datasets?$top=${top}`);
    if (!r.ok) return { content: toErrorContent("get_datasets_in_workspace failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get dataset datasources (for connection audits)
server.tool(
  "get_dataset_datasources",
  {
    description: "Returns data sources for a dataset. Essential for auditing connections and credentials.",
    inputSchema: { 
      type: "object", 
      properties: { 
        datasetId: { type: "string", description: "The dataset ID (GUID)" }
      },
      required: ["datasetId"]
    },
  },
  async (input) => {
    const { datasetId } = input;
    const r = await pbiAdminFetch(`/datasets/${datasetId}/datasources`);
    if (!r.ok) return { content: toErrorContent("get_dataset_datasources failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get dataset users (access audit)
server.tool(
  "get_dataset_users",
  {
    description: "Returns users with access to a dataset. Important for security reviews.",
    inputSchema: { 
      type: "object", 
      properties: { 
        datasetId: { type: "string", description: "The dataset ID (GUID)" }
      },
      required: ["datasetId"]
    },
  },
  async (input) => {
    const { datasetId } = input;
    const r = await pbiAdminFetch(`/datasets/${datasetId}/users`);
    if (!r.ok) return { content: toErrorContent("get_dataset_users failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// ---------- Report Health Tools ----------

// Get reports in workspace
server.tool(
  "get_reports_in_workspace",
  {
    description: "Returns all reports in a workspace. Useful for content inventory.",
    inputSchema: { 
      type: "object", 
      properties: { 
        groupId: { type: "string", description: "The workspace ID (GUID)" },
        top: { type: "number", description: "Number of results", default: 100 }
      },
      required: ["groupId"]
    },
  },
  async (input) => {
    const { groupId, top = 100 } = input;
    const r = await pbiAdminFetch(`/groups/${groupId}/reports?$top=${top}`);
    if (!r.ok) return { content: toErrorContent("get_reports_in_workspace failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get report users
server.tool(
  "get_report_users",
  {
    description: "Returns users with access to a report. Critical for sharing audits.",
    inputSchema: { 
      type: "object", 
      properties: { 
        reportId: { type: "string", description: "The report ID (GUID)" }
      },
      required: ["reportId"]
    },
  },
  async (input) => {
    const { reportId } = input;
    const r = await pbiAdminFetch(`/reports/${reportId}/users`);
    if (!r.ok) return { content: toErrorContent("get_report_users failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// ---------- Dashboard Health Tools ----------

// Get dashboards in workspace
server.tool(
  "get_dashboards_in_workspace",
  {
    description: "Returns all dashboards in a workspace.",
    inputSchema: { 
      type: "object", 
      properties: { 
        groupId: { type: "string", description: "The workspace ID (GUID)" }
      },
      required: ["groupId"]
    },
  },
  async (input) => {
    const { groupId } = input;
    const r = await pbiAdminFetch(`/groups/${groupId}/dashboards`);
    if (!r.ok) return { content: toErrorContent("get_dashboards_in_workspace failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get dashboard users
server.tool(
  "get_dashboard_users",
  {
    description: "Returns users with access to a dashboard.",
    inputSchema: { 
      type: "object", 
      properties: { 
        dashboardId: { type: "string", description: "The dashboard ID (GUID)" }
      },
      required: ["dashboardId"]
    },
  },
  async (input) => {
    const { dashboardId } = input;
    const r = await pbiAdminFetch(`/dashboards/${dashboardId}/users`);
    if (!r.ok) return { content: toErrorContent("get_dashboard_users failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// ---------- Dataflow Health Tools ----------

// Get dataflows in workspace
server.tool(
  "get_dataflows_in_workspace",
  {
    description: "Returns all dataflows in a workspace.",
    inputSchema: { 
      type: "object", 
      properties: { 
        groupId: { type: "string", description: "The workspace ID (GUID)" }
      },
      required: ["groupId"]
    },
  },
  async (input) => {
    const { groupId } = input;
    const r = await pbiAdminFetch(`/groups/${groupId}/dataflows`);
    if (!r.ok) return { content: toErrorContent("get_dataflows_in_workspace failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get dataflow datasources
server.tool(
  "get_dataflow_datasources",
  {
    description: "Returns data sources for a dataflow.",
    inputSchema: { 
      type: "object", 
      properties: { 
        dataflowId: { type: "string", description: "The dataflow ID (GUID)" }
      },
      required: ["dataflowId"]
    },
  },
  async (input) => {
    const { dataflowId } = input;
    const r = await pbiAdminFetch(`/dataflows/${dataflowId}/datasources`);
    if (!r.ok) return { content: toErrorContent("get_dataflow_datasources failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// ---------- Security & Compliance Tools ----------

// Get widely shared reports (security risk)
server.tool(
  "get_widely_shared_reports",
  {
    description: "Returns reports shared with entire organization through links. Critical security check.",
    inputSchema: { type: "object", properties: {} },
  },
  async () => {
    const r = await pbiAdminFetch(`/widelySharedArtifacts/linksSharedToWholeOrganization`);
    if (!r.ok) return { content: toErrorContent("get_widely_shared_reports failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get publish to web reports (major security risk)
server.tool(
  "get_published_to_web",
  {
    description: "Returns items published to web (publicly accessible). CRITICAL security check.",
    inputSchema: { type: "object", properties: {} },
  },
  async () => {
    const r = await pbiAdminFetch(`/widelySharedArtifacts/publishedToWeb`);
    if (!r.ok) return { content: toErrorContent("get_published_to_web failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get user artifact access (audit specific user)
server.tool(
  "get_user_artifact_access",
  {
    description: "Returns all Power BI items a specific user has access to. For user access audits.",
    inputSchema: { 
      type: "object", 
      properties: { 
        userId: { type: "string", description: "User principal name or object ID" }
      },
      required: ["userId"]
    },
  },
  async (input) => {
    const { userId } = input;
    const encodedUserId = encodeURIComponent(userId);
    const r = await pbiAdminFetch(`/users/${encodedUserId}/artifactAccess`);
    if (!r.ok) return { content: toErrorContent("get_user_artifact_access failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// ---------- Apps Health Tools ----------

// Get all apps
server.tool(
  "get_admin_apps",
  {
    description: "Returns all apps in the organization. Important for app governance.",
    inputSchema: { 
      type: "object", 
      properties: { 
        top: { type: "number", description: "Number of results", default: 100 }
      }
    },
  },
  async (input) => {
    const { top = 100 } = input;
    const r = await pbiAdminFetch(`/apps?$top=${top}`);
    if (!r.ok) return { content: toErrorContent("get_admin_apps failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get app users
server.tool(
  "get_app_users",
  {
    description: "Returns users with access to an app.",
    inputSchema: { 
      type: "object", 
      properties: { 
        appId: { type: "string", description: "The app ID (GUID)" }
      },
      required: ["appId"]
    },
  },
  async (input) => {
    const { appId } = input;
    const r = await pbiAdminFetch(`/apps/${appId}/users`);
    if (!r.ok) return { content: toErrorContent("get_app_users failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// ---------- Activity & Audit Tools ----------

// Get activity events (audit logs)
server.tool(
  "get_activity_events",
  {
    description: "Returns audit activity events. Essential for compliance and security monitoring. Format: YYYY-MM-DDTHH:mm:ss",
    inputSchema: { 
      type: "object", 
      properties: { 
        startDateTime: { type: "string", description: "Start time in UTC (e.g., 2025-01-01T00:00:00)" },
        endDateTime: { type: "string", description: "End time in UTC (e.g., 2025-01-01T23:59:59)" }
      },
      required: ["startDateTime", "endDateTime"]
    },
  },
  async (input) => {
    const { startDateTime, endDateTime } = input;
    const start = encodeURIComponent(startDateTime);
    const end = encodeURIComponent(endDateTime);
    const r = await pbiAdminFetch(`/activityevents?startDateTime='${start}'&endDateTime='${end}'`);
    if (!r.ok) return { content: toErrorContent("get_activity_events failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// ---------- Workspace Scan Tools (Most Powerful) ----------

// Initiate workspace scan
server.tool(
  "initiate_workspace_scan",
  {
    description: "Initiates a metadata scan for specified workspaces. Returns scanId to check status. Most comprehensive health check tool.",
    inputSchema: { 
      type: "object", 
      properties: { 
        workspaceIds: { 
          type: "array", 
          items: { type: "string" },
          description: "Array of workspace IDs to scan. Omit to scan all workspaces."
        },
        datasetExpressions: { type: "boolean", description: "Include DAX expressions", default: false },
        datasetSchema: { type: "boolean", description: "Include dataset schema", default: true },
        datasourceDetails: { type: "boolean", description: "Include datasource details", default: true },
        getArtifactUsers: { type: "boolean", description: "Include user access info", default: true },
        lineage: { type: "boolean", description: "Include lineage info", default: true }
      }
    },
  },
  async (input) => {
    const {
      workspaceIds,
      datasetExpressions = false,
      datasetSchema = true,
      datasourceDetails = true,
      getArtifactUsers = true,
      lineage = true
    } = input || {};

    const body = {
      ...(workspaceIds && { workspaces: workspaceIds }),
      datasetExpressions,
      datasetSchema,
      datasourceDetails,
      getArtifactUsers,
      lineage
    };

    const token = await getAccessToken();
    const url = `https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo`;
    
    console.log("[PBI REQUEST]", url, JSON.stringify(body));

    let res, text;
    try {
      res = await fetch(url, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(body)
      });
      text = await res.text();
    } catch (e) {
      console.error("[PBI NETWORK ERROR]", url, String(e));
      return { content: [{ type: "text", text: `Network error: ${String(e)}` }] };
    }

    if (!res.ok) {
      console.error("[PBI ERROR]", res.status, url, text);
      return { content: toErrorContent("initiate_workspace_scan failed", { ok: false, status: res.status, errorText: text }) };
    }

    let responseBody = {};
    try { responseBody = text ? JSON.parse(text) : {}; } catch { responseBody = {}; }
    
    return { content: [{ type: "text", text: JSON.stringify(responseBody, null, 2) }] };
  }
);

// Get detailed report information
server.tool(
  "get_admin_report_details",
  {
    description: "Returns detailed information about a specific report including name, description, dataset, web URL, and metadata.",
    inputSchema: { 
      type: "object", 
      properties: { 
        reportId: { type: "string", description: "The report ID (GUID)" }
      },
      required: ["reportId"]
    },
  },
  async (input) => {
    const { reportId } = input;
    const r = await pbiAdminFetch(`/reports/${reportId}`);
    if (!r.ok) return { content: toErrorContent("get_admin_report_details failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get scan status
server.tool(
  "get_scan_status",
  {
    description: "Checks the status of a workspace scan initiated by initiate_workspace_scan.",
    inputSchema: { 
      type: "object", 
      properties: { 
        scanId: { type: "string", description: "The scan ID returned from initiate_workspace_scan" }
      },
      required: ["scanId"]
    },
  },
  async (input) => {
    const { scanId } = input;
    const r = await pbiAdminFetch(`/workspaces/scanStatus/${scanId}`);
    if (!r.ok) return { content: toErrorContent("get_scan_status failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get scan result
server.tool(
  "get_scan_result",
  {
    description: "Gets the full scan result once status is Succeeded. Contains comprehensive workspace metadata.",
    inputSchema: { 
      type: "object", 
      properties: { 
        scanId: { type: "string", description: "The scan ID" }
      },
      required: ["scanId"]
    },
  },
  async (input) => {
    const { scanId } = input;
    const r = await pbiAdminFetch(`/workspaces/scanResult/${scanId}`);
    if (!r.ok) return { content: toErrorContent("get_scan_result failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

// Get modified workspaces (for incremental scans)
server.tool(
  "get_modified_workspaces",
  {
    description: "Returns workspace IDs modified since a specific date. Useful for incremental health checks.",
    inputSchema: { 
      type: "object", 
      properties: { 
        modifiedSince: { type: "string", description: "ISO 8601 date (e.g., 2025-01-01T00:00:00.000Z)" }
      }
    },
  },
  async (input) => {
    const { modifiedSince } = input || {};
    let path = `/workspaces/modified`;
    if (modifiedSince) {
      path += `?modifiedSince=${encodeURIComponent(modifiedSince)}`;
    }
    const r = await pbiAdminFetch(path);
    if (!r.ok) return { content: toErrorContent("get_modified_workspaces failed", r) };
    return { content: [{ type: "text", text: JSON.stringify(r.body, null, 2) }] };
  }
);

  return server;
}

// ---------- Express App ----------
const app = express();
app.use(express.json());

// Logs (helpful debugging)
app.use((req, _res, next) => {
  console.log(
    "[REQ]",
    req.method,
    req.path,
    "ctype:", req.header("content-type"),
    "mcp-session:", req.header("Mcp-Session-Id")
  );
  next();
});

// Health/info
app.get("/health", (_req, res) => res.json({ ok: true }));
app.get(BASE_PATH, (_req, res) =>
  res.json({ ok: true, info: "MCP endpoint. Initialize (no session ok) OR register -> initialize -> tools/call." })
);

// ---------- Session management ----------
const sessions = new Map(); // sid -> { transport, server }
function createSession(sessionId) {
  const server = buildMcpServer();
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: () => sessionId
  });
  server.connect(transport);
  sessions.set(sessionId, { transport, server });
}

// Optional explicit registration (useful for Postman/manual)
app.post("/register", (_req, res) => {
  const sid = uuidv4();
  createSession(sid);
  res.setHeader("Mcp-Session-Id", sid);
  return res.status(201).json({ ok: true, mcpSessionId: sid });
});

// ---------- MCP endpoint ----------
// Supports two modes:
//   A) No Mcp-Session-Id + initialize  -> auto-creates session and proceeds
//   B) Mcp-Session-Id present          -> uses existing session
app.all(BASE_PATH, async (req, res) => {
  let sid = req.header("Mcp-Session-Id");
  const method = req.body?.method;

  // A) Auto-session on initialize with no header (Claude web behavior)
  if (!sid && req.method === "POST" && method === "initialize") {
    sid = uuidv4();
    createSession(sid);
    res.setHeader("Mcp-Session-Id", sid);
  }

  // After auto-create, we should have a session id
  if (!sid) {
    return res.status(400).json({ error: "Bad Request: Mcp-Session-Id header is required" });
  }
  const session = sessions.get(sid);
  if (!session) {
    return res.status(404).json({ error: "Session not found" });
  }

  // Be lenient on Accept header (Claude may omit it); only warn in logs
  const accept = (req.header("Accept") || "").toLowerCase();
  if (!accept.includes("application/json") || !accept.includes("text/event-stream")) {
    console.warn("[WARN] Accept header missing required values; continuing anyway.");
  }

  await session.transport.handleRequest(req, res, req.body);
});

// ---------- Start ----------
app.listen(Number(PORT), () => {
  console.log(`✅ MCP server running on port ${PORT} at ${BASE_PATH}`);
});