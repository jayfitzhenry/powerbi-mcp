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
  // API_KEY,  // (disabled for now so Claude web can connect)
  PORT = process.env.PORT || "8787",
  BASE_PATH = "/mcp",
} = process.env;

console.log("ENV OK?", !!TENANT_ID, !!CLIENT_ID, !!CLIENT_SECRET);

// ---------- Power BI Auth (MSAL) ----------
if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
  console.error("Missing required env: TENANT_ID, CLIENT_ID, CLIENT_SECRET");
  process.exit(1);
}

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

async function pbiAdminFetch(pathAndQuery) {
  const token = await getAccessToken();
  const url = `https://api.powerbi.com/v1.0/myorg/admin${pathAndQuery}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Power BI API error ${res.status}: ${await res.text()}`);
  return { status: res.status, body: await res.json() };
}

// ---------- MCP Server (tools) ----------
function buildMcpServer() {
  const server = new McpServer({
    name: "powerbi-admin-mcp-remote",
    version: "1.0.0",
  });

  server.tool(
    "list_admin_groups",
    {
      description: "List all Power BI workspaces across the organisation.",
      inputSchema: {
        type: "object",
        properties: { top: { type: "number" } },
      },
    },
    async (input) => {
      const top = input?.top ?? 100;
      const result = await pbiAdminFetch(`/groups?scope=Organization&$top=${top}`);
      return { content: [{ type: "json", json: result.body }] };
    }
  );

  // You can add more tools here (e.g., admin_get, get_activity_events)

  return server;
}

// ---------- Express App ----------
const app = express();
app.use(express.json());

// Request logger (helps debugging)
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

// ---------- Streamable HTTP Transport (let it manage sessions) ----------
const transport = new StreamableHTTPServerTransport({
  // Newer SDKs require options; this generates session ids for you
  sessionIdGenerator: () => uuidv4(),
});

const mcpServer = buildMcpServer();
mcpServer.connect(transport);

// Single handler for the MCP endpoint
app.all(BASE_PATH, async (req, res) => {
  // No custom auth for now -> Claude web can connect
  // (If you later want auth, we can add rules that still allow initialize/session)
  await transport.handleRequest(req, res, req.body);
});

// ---------- Start ----------
app.listen(Number(PORT), () => {
  console.log(`✅ MCP server running on port ${PORT} at ${BASE_PATH}`);
});
