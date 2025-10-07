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

// Basic fetch wrapper
async function pbiAdminFetch(pathAndQuery) {
  const token = await getAccessToken();
  const url = `https://api.powerbi.com/v1.0/myorg/admin${pathAndQuery}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const text = await res.text();
  if (!res.ok) {
    console.error("[PBI ERROR]", res.status, text);
    throw new Error(`Power BI API error ${res.status}: ${text}`);
  }
  return { status: res.status, body: text ? JSON.parse(text) : {} };
}

// Simple tenant-wide pager for endpoints that return { value: [...] }
async function pbiAdminCount(pathBaseWithScope, pageSize = 5000) {
  let skip = 0;
  let total = 0;

  while (true) {
    // Ensure we add the paging params correctly
    const sep = pathBaseWithScope.includes("?") ? "&" : "?";
    const path = `${pathBaseWithScope}${sep}$top=${pageSize}&$skip=${skip}`;
    const { body } = await pbiAdminFetch(path);

    const items = Array.isArray(body?.value) ? body.value : [];
    total += items.length;

    if (items.length < pageSize) break; // last page
    skip += pageSize;
  }

  return total;
}

// ---------- MCP Server (tools) ----------
function buildMcpServer() {
  const server = new McpServer({
    name: "powerbi-admin-mcp-remote",
    version: "1.0.0",
  });

  // 1) Existing: list workspaces
  server.tool(
    "list_admin_groups",
    {
      description: "List Power BI workspaces across the organisation (tenant-wide).",
      inputSchema: {
        type: "object",
        properties: { top: { type: "number", description: "Number to return (default 100)" } },
      },
    },
    async (input) => {
      const top = input?.top ?? 100;
      const result = await pbiAdminFetch(`/groups?scope=Organization&$top=${top}`);
      return { content: [{ type: "json", json: result.body }] };
    }
  );

  // 2) NEW: count datasets across the tenant
  server.tool(
    "count_admin_datasets",
    {
      description: "Return the total number of datasets across the organisation.",
      inputSchema: { type: "object", properties: {} },
    },
    async () => {
      const total = await pbiAdminCount(`/datasets?scope=Organization`);
      return { content: [{ type: "json", json: { datasetsCount: total } }] };
    }
  );

  // 3) NEW: count reports across the tenant
  server.tool(
    "count_admin_reports",
    {
      description: "Return the total number of reports across the organisation.",
      inputSchema: { type: "object", properties: {} },
    },
    async () => {
      const total = await pbiAdminCount(`/reports?scope=Organization`);
      return { content: [{ type: "json", json: { reportsCount: total } }] };
    }
  );

  // 4) NEW: combined counts (datasets + reports)
  server.tool(
    "count_admin_assets",
    {
      description: "Return both dataset and report totals across the organisation.",
      inputSchema: { type: "object", properties: {} },
    },
    async () => {
      const [datasetsCount, reportsCount] = await Promise.all([
        pbiAdminCount(`/datasets?scope=Organization`),
        pbiAdminCount(`/reports?scope=Organization`),
      ]);
      return { content: [{ type: "json", json: { datasetsCount, reportsCount } }] };
    }
  );

  return server;
}

// ---------- Express App ----------
const app = express();
app.use(express.json());

// Minimal logs (helpful if anything fails)
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

// Friendly routes
app.get("/health", (_req, res) => res.json({ ok: true }));
app.get(BASE_PATH, (_req, res) =>
  res.json({ ok: true, info: "MCP endpoint. POST JSON-RPC initialize first, then tools/call." })
);

// ---------- MCP Transport (SDK manages sessions) ----------
const transport = new StreamableHTTPServerTransport({
  sessionIdGenerator: () => uuidv4(),
});
const mcpServer = buildMcpServer();
mcpServer.connect(transport);

// Single handler for MCP endpoint
app.all(BASE_PATH, async (req, res) => {
  await transport.handleRequest(req, res, req.body);
});

// ---------- Start ----------
app.listen(Number(PORT), () => {
  console.log(`✅ MCP server running on port ${PORT} at ${BASE_PATH}`);
});
