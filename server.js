#!/usr/bin/env node

// ---------- Imports ----------
import 'dotenv/config';
import express from "express";
import fetch from "node-fetch";
import { v4 as uuidv4 } from "uuid";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";

// ---------- Environment Variables ----------
const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  API_KEY,
  PORT = process.env.PORT || "8787",
  BASE_PATH = "/mcp",
} = process.env;

console.log("ENV OK?", !!TENANT_ID, !!CLIENT_ID, !!CLIENT_SECRET, !!API_KEY);

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !API_KEY) {
  console.error("Missing required env: TENANT_ID, CLIENT_ID, CLIENT_SECRET, API_KEY");
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

async function pbiAdminFetch(pathAndQuery) {
  const token = await getAccessToken();
  const url = `https://api.powerbi.com/v1.0/myorg/admin${pathAndQuery}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Power BI API error ${res.status}: ${await res.text()}`);
  return { status: res.status, body: await res.json() };
}

// ---------- MCP Server Definition ----------
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

  return server;
}

// ---------- Express Setup ----------
const app = express();
app.use(express.json());

// Log each request (helpful for debugging)
app.use((req, _res, next) => {
  console.log(
    "[REQ]",
    req.method,
    req.path,
    "auth:", !!req.header("Authorization"),
    "session:", req.header("Mcp-Session-Id"),
    "ctype:", req.header("content-type")
  );
  next();
});

// ---------- Auth Middleware ----------
// Allow "initialize" without auth, require API key for everything else.
app.use((req, res, next) => {
  if (req.path !== BASE_PATH) return next();

  if (req.method === "POST" && req.body?.method === "initialize") {
    return next(); // allow Claude to start without auth
  }

  const isAuthed = req.header("Authorization") === `Bearer ${API_KEY}`;
  if (!isAuthed) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  next();
});

// ---------- MCP Session Handling ----------
const sessions = new Map();

function newTransport() {
  const transport = new StreamableHTTPServerTransport();
  const server = buildMcpServer();
  server.connect(transport);
  return { transport };
}

app.all(BASE_PATH, async (req, res) => {
  const sessionId = req.header("Mcp-Session-Id");
  let session = sessions.get(sessionId);

  const isInit = req.method === "POST" && req.body?.method === "initialize";
  if (!session && !isInit) return res.status(400).json({ error: "Send initialize first" });

  if (isInit) {
    const { transport } = newTransport();
    const id = uuidv4();
    sessions.set(id, { transport });
    res.setHeader("Mcp-Session-Id", id);
    await transport.handleRequest(req, res, req.body);
    return;
  }

  if (req.method === "DELETE") {
    sessions.delete(sessionId);
    return res.status(204).end();
  }

  await session.transport.handleRequest(req, res, req.body);
});

// ---------- Helpful Routes ----------
app.get(BASE_PATH, (_req, res) => {
  res.json({
    ok: true,
    info: "MCP endpoint ready. Send JSON-RPC initialize first, then tools/call.",
  });
});

app.get("/health", (_req, res) => res.json({ ok: true }));

// ---------- Start Server ----------
app.listen(Number(PORT), () => {
  console.log(`✅ MCP server running on port ${PORT} at ${BASE_PATH}`);
});
