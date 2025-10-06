#!/usr/bin/env node
import 'dotenv/config';
import express from "express";
import fetch from "node-fetch";
import { v4 as uuidv4 } from "uuid";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";

// -------- ENV --------
const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  API_KEY,
  PORT = "8787",
  BASE_PATH = "/mcp",
} = process.env;

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !API_KEY) {
  console.error("Missing required env: TENANT_ID, CLIENT_ID, CLIENT_SECRET, API_KEY");
  process.exit(1);
}

// -------- MSAL (Power BI) --------
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

// -------- MCP server definition --------
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
      return {
  content: [{ type: "json", json: result.body }],
};

    }
  );

  return server;
}

// -------- HTTP wiring --------
const app = express();
app.use(express.json());

// Simple API key check
app.use((req, res, next) => {
  if (req.header("Authorization") !== `Bearer ${API_KEY}`) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
});

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

app.listen(Number(PORT), () => {
  console.log(`✅ MCP server running on port ${PORT} at ${BASE_PATH}`);
});
