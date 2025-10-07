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

// Decode JWT (no verify) to inspect header/payload for diagnostics
function decodeJwtNoVerify(jwt) {
  try {
    const [h, p] = jwt.split(".");
    const pad = (s) => s.replace(/-/g, "+").replace(/_/g, "/") + "===".slice((s.length + 3) % 4);
    const header = JSON.parse(Buffer.from(pad(h), "base64").toString("utf8"));
    const payload = JSON.parse(Buffer.from(pad(p), "base64").toString("utf8"));
    return { header, payload };
  } catch {
    return null;
  }
}

// Centralised fetch with rich logging
async function pbiAdminFetch(pathAndQuery) {
  const token = await getAccessToken();
  const url = `https://api.powerbi.com/v1.0/myorg/admin${pathAndQuery}`;

  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const text = await res.text();

  if (!res.ok) {
    console.error("[PBI ERROR]", res.status, pathAndQuery, text);
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
    version: "1.1.0",
  });

  // Helper to return errors to Claude instead of throwing
  const toErrorContent = (prefix, res) => ([
    { type: "text", text: `${prefix}: HTTP ${res.status}\n${res.errorText ?? "No body"}` }
  ]);

  // 0) Diagnostics: quick ping to admin API
  server.tool(
    "diagnose_admin_ping",
    {
      description: "Checks if the service principal can call /admin/capacities (tenant-level).",
      inputSchema: { type: "object", properties: {} },
    },
    async () => {
      const r = await pbiAdminFetch(`/capacities?$top=1`);
      if (!r.ok) return { content: toErrorContent("Admin ping failed", r) };
      return { content: [{ type: "json", json: r.body }] };
    }
  );

  // 0b) Diagnostics: show access token claims (audience/tenant/appid/roles)
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
        // Redact the token but return safe fields
        const subset = {
          header,
          payload: {
            aud: payload.aud,
            appid: payload.appid,
            tid: payload.tid,
            roles: payload.roles,
            aio: payload.aio,
            ver: payload.ver,
            iss: payload.iss,
            iat: payload.iat,
            nbf: payload.nbf,
            exp: payload.exp,
          }
        };
        return { content: [{ type: "json", json: subset }] };
      } catch (e) {
        return { content: [{ type: "text", text: `Token error: ${String(e)}` }] };
      }
    }
  );

  // 1) List workspaces
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
      return { content: [{ type: "json", json: r.body }] };
    }
  );

  // 2) Count datasets
  server.tool(
    "count_admin_datasets",
    {
      description: "Return the total number of datasets across the organisation.",
      inputSchema: { type: "object", properties: {} },
    },
    async () => {
      const r = await pbiAdminCount(`/datasets?scope=Organization`);
      if (!r.ok) return { content: toErrorContent("count_admin_datasets failed", r) };
      return { content: [{ type: "json", json: { datasetsCount: r.total } }] };
    }
  );

  // 3) Count reports
  server.tool(
    "count_admin_reports",
    {
      description: "Return the total number of reports across the organisation.",
      inputSchema: { type: "object", properties: {} },
    },
    async () => {
      const r = await pbiAdminCount(`/reports?scope=Organization`);
      if (!r.ok) return { content: toErrorContent("count_admin_reports failed", r) };
      return { content: [{ type: "json", json: { reportsCount: r.total } }] };
    }
  );

  // 4) Combined counts
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
      return { content: [{ type: "json", json: { datasetsCount: d.total, reportsCount: r.total } }] };
    }
  );

  return server;
}

// ---------- Express App ----------
const app = express();
app.use(express.json());

// Minimal logs
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
