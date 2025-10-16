/**
 * Combined HTTP wrapper + MCP server
 *
 * - Starts an MCP server (stdio transport) and registers tools from `src/tools.ts`
 * - Also exposes a small HTTP API for quick testing:
 *    GET  /customers/:accountNum
 *    GET  /customers? (supports $filter or filter, $select or select)
 *    GET  /vendors/:vendorAccount
 *    GET  /vendors? (supports $filter or filter, $select or select)
 *
 * Run (dev):
 *   npx ts-node --esm scripts/server-http.ts
 *
 * Notes:
 * - Local imports use .ts extensions so ts-node --esm can resolve them.
 * - This will connect the MCP server to process.stdin/stdout (StdioServerTransport).
 *   If you don't want the stdio MCP transport, set env var `DISABLE_MCP=true`.
 */

import dotenv from "dotenv";
dotenv.config();

import express from "express";
import type { Request, Response } from "express";

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { Dynamics365FO } from "../src/main.ts";
import { registerTools } from "../src/tools.ts";

const port = process.env.HTTP_PORT ? parseInt(process.env.HTTP_PORT, 10) : 3000;

const clientId = process.env.CLIENT_ID || "";
const clientSecret = process.env.CLIENT_SECRET || "";
const tenantId = process.env.TENANT_ID || "";
const D365_BASE_URL = process.env.D365_URL || "";

if (!clientId || !clientSecret || !tenantId || !D365_BASE_URL) {
  console.error("Missing required environment variables. See README / .env");
  process.exit(1);
}

// Create FO helper and MCP server
const fo = new Dynamics365FO(clientId, clientSecret, tenantId, D365_BASE_URL);
const server = new McpServer({
  name: "Dynamics365FO",
  version: "1.0.0.0",
});

// Register tools on the MCP server (uses your src/tools.ts)
registerTools(server, fo);

// Optionally connect the MCP server to stdio transport (default enabled).
if (!process.env.DISABLE_MCP || process.env.DISABLE_MCP === "false") {
  (async () => {
    try {
      const transport = new StdioServerTransport();
      await server.connect(transport);
      console.error("MCP server registered and connected on stdio transport.");
    } catch (err) {
      console.error("Failed to connect MCP server on stdio transport:", err);
    }
  })();
} else {
  console.error("MCP stdio transport disabled via DISABLE_MCP=true");
}

const app = express();
app.use(express.json());

// Helper to normalize response shapes to a single record (or null)
function normalizeSingleRecord(resp: any): any | null {
  if (!resp) return null;
  if (Array.isArray(resp)) return resp.length ? resp[0] : null;
  if (Array.isArray(resp.value)) return resp.value.length ? resp.value[0] : null;
  if (typeof resp === "object") {
    if (resp.value && Array.isArray(resp.value) && resp.value.length) return resp.value[0];
    return Object.keys(resp).length ? resp : null;
  }
  return null;
}

// GET single customer by account identifier (CustomerAccount)
app.get("/customers/:account", async (req: Request, res: Response): Promise<void> => {
  try {
    const account = req.params.account;
    const crossCompany = req.query.crossCompany === "true" || req.query["cross-company"] === "true";
    const fields = (req.query.$select ?? req.query.select) ? String(req.query.$select ?? req.query.select).split(",") : undefined;

    // Call the FO helper (getCustomerByAccountNum uses CustomerAccount field)
    const resp = await fo.getCustomerByAccountNum(account, {
      select: fields,
      crossCompany,
    });

    const record = normalizeSingleRecord(resp);

    if (!record) {
      res.status(404).json({ message: "Not found", account });
      return;
    }

    res.json(record);
    return;
  } catch (err) {
    console.error("Error in /customers/:account", err);
    res.status(500).json({ error: (err as Error).message || String(err) });
    return;
  }
});

// Generic customers endpoint supporting raw OData params ($filter/$select) or plain names
app.get("/customers", async (req: Request, res: Response): Promise<void> => {
  try {
    const rawFilter = (req.query.$filter ?? req.query.filter) as string | undefined;
    const rawSelect = (req.query.$select ?? req.query.select) as string | undefined;

    const filter = rawFilter ? String(rawFilter) : undefined;
    const select = rawSelect ? String(rawSelect).split(",") : undefined;
    const top = req.query.top ? parseInt(String(req.query.top), 10) : undefined;
    const crossCompany = req.query.crossCompany === "true" || req.query["cross-company"] === "true";
    const fetchAllPages = req.query.fetchAllPages === "true";

    const response = await fo.getCustomers({
      filter,
      select,
      top,
      crossCompany,
      fetchAllPages,
    });

    const payload = response && (response.value ?? response);
    res.json(payload);
    return;
  } catch (err) {
    console.error("Error in /customers", err);
    res.status(500).json({ error: (err as Error).message || String(err) });
    return;
  }
});

// ===== Vendor routes =====

// GET single vendor by vendor account identifier (VendorAccount)
app.get("/vendors/:vendorAccount", async (req: Request, res: Response): Promise<void> => {
  try {
    const vendorAccount = req.params.vendorAccount;
    const crossCompany = req.query.crossCompany === "true" || req.query["cross-company"] === "true";
    const fields = (req.query.$select ?? req.query.select) ? String(req.query.$select ?? req.query.select).split(",") : undefined;

    const resp = await fo.getVendorByAccountNum(vendorAccount, {
      select: fields,
      crossCompany,
    });

    const record = normalizeSingleRecord(resp);

    if (!record) {
      res.status(404).json({ message: "Not found", vendorAccount });
      return;
    }

    res.json(record);
    return;
  } catch (err) {
    console.error("Error in /vendors/:vendorAccount", err);
    res.status(500).json({ error: (err as Error).message || String(err) });
    return;
  }
});

// Generic vendors endpoint supporting raw OData params ($filter/$select)
app.get("/vendors", async (req: Request, res: Response): Promise<void> => {
  try {
    const rawFilter = (req.query.$filter ?? req.query.filter) as string | undefined;
    const rawSelect = (req.query.$select ?? req.query.select) as string | undefined;

    const filter = rawFilter ? String(rawFilter) : undefined;
    const select = rawSelect ? String(rawSelect).split(",") : undefined;
    const top = req.query.top ? parseInt(String(req.query.top), 10) : undefined;
    const crossCompany = req.query.crossCompany === "true" || req.query["cross-company"] === "true";
    const fetchAllPages = req.query.fetchAllPages === "true";

    const response = await fo.getVendors({
      filter,
      select,
      top,
      crossCompany,
      fetchAllPages,
    });

    const payload = response && (response.value ?? response);
    res.json(payload);
    return;
  } catch (err) {
    console.error("Error in /vendors", err);
    res.status(500).json({ error: (err as Error).message || String(err) });
    return;
  }
});

app.listen(port, () => {
  console.error(`FO HTTP + MCP wrapper listening at http://localhost:${port}`);
  console.error("MCP tools are registered; connect an MCP client over stdio to use them.");
});