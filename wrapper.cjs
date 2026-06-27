#!/usr/bin/env node
/**
 * wrapper.cjs
 *
 * Sits between mcp-proxy (the SSE/HTTP bridge) and the real
 * mcp-sharepoint process (stdio JSON-RPC).
 *
 * What it does:
 *  1. Spawns "node src/index.js" (the real mcp-sharepoint server) as a child.
 *  2. Intercepts the "tools/list" response and removes the credential
 *     fields (tenantId, clientId, clientSecret, siteId, driveId) from
 *     each tool's input schema, so Claude never sees or asks for them.
 *  3. Intercepts every "tools/call" request and injects the real
 *     credential values (read from environment variables) into the
 *     arguments before forwarding to the child process.
 *  4. Everything else is passed through untouched.
 *
 * Required environment variables (set these in Easypanel's
 * Environment Variables panel for this service — NOT in the
 * Dockerfile, and NOT in chat):
 *   SP_TENANT_ID
 *   SP_CLIENT_ID
 *   SP_CLIENT_SECRET
 *   SP_SITE_ID
 *   SP_DRIVE_ID
 */

const { spawn } = require("child_process");
const readline = require("readline");

const CRED_FIELDS = ["tenantId", "clientId", "clientSecret", "siteId", "driveId"];

const CREDS = {
  tenantId: process.env.SP_TENANT_ID,
  clientId: process.env.SP_CLIENT_ID,
  clientSecret: process.env.SP_CLIENT_SECRET,
  siteId: process.env.SP_SITE_ID,
  driveId: process.env.SP_DRIVE_ID,
};

for (const [key, val] of Object.entries(CREDS)) {
  if (!val) {
    process.stderr.write(
      `[wrapper] WARNING: missing env var for "${key}" — tool calls needing it will fail.\n`
    );
  }
}

// Spawn the real mcp-sharepoint server
const child = spawn("node", ["src/index.js"], {
  stdio: ["pipe", "pipe", "inherit"], // inherit stderr so its logs still show in Easypanel
});

// --- Outgoing: parent (mcp-proxy) -> wrapper -> child ---
const parentReader = readline.createInterface({ input: process.stdin });

parentReader.on("line", (line) => {
  if (!line.trim()) return;
  let msg;
  try {
    msg = JSON.parse(line);
  } catch {
    // Not JSON, just forward raw
    child.stdin.write(line + "\n");
    return;
  }

  // Inject credentials into every tools/call request
  if (msg.method === "tools/call" && msg.params && msg.params.arguments) {
    msg.params.arguments = { ...msg.params.arguments, ...CREDS };
  }

  child.stdin.write(JSON.stringify(msg) + "\n");
});

// --- Incoming: child -> wrapper -> parent (mcp-proxy) ---
const childReader = readline.createInterface({ input: child.stdout });

childReader.on("line", (line) => {
  if (!line.trim()) return;
  let msg;
  try {
    msg = JSON.parse(line);
  } catch {
    process.stdout.write(line + "\n");
    return;
  }

  // Strip credential fields out of the tools/list response schema
  if (msg.result && Array.isArray(msg.result.tools)) {
    for (const tool of msg.result.tools) {
      if (tool.inputSchema && tool.inputSchema.properties) {
        for (const field of CRED_FIELDS) {
          delete tool.inputSchema.properties[field];
        }
        if (Array.isArray(tool.inputSchema.required)) {
          tool.inputSchema.required = tool.inputSchema.required.filter(
            (f) => !CRED_FIELDS.includes(f)
          );
        }
      }
    }
  }

  process.stdout.write(JSON.stringify(msg) + "\n");
});

child.on("exit", (code) => {
  process.stderr.write(`[wrapper] child exited with code ${code}\n`);
  process.exit(code ?? 1);
});

process.on("SIGTERM", () => child.kill("SIGTERM"));
process.on("SIGINT", () => child.kill("SIGINT"));
