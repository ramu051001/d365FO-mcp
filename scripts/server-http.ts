/**
 * Combined HTTP wrapper + MCP server
 *
 * - Starts an MCP server (HTTP transport) and registers tools from `src/tools.ts`
 * - **Exposes the remote MCP endpoint at /mcp**
 * - Also exposes a small HTTP API for quick testing:
 * GET  /customers/:accountNum
 * GET  /customers? (supports $filter or filter, $select or select)
 * GET  /vendors/:vendorAccount
 * GET  /vendors? (supports $filter or filter, $select or select)
 *
 * Notes:
 * - This uses the StreamableHTTPServerTransport for remote access, accessible at /mcp.
 */

import dotenv from "dotenv";
dotenv.config();

import express from "express";
import type { Request, Response } from "express";
import cors from "cors";

import { randomUUID } from "node:crypto"; // Required for sessionIdGenerator

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js"; 

// FIX: Change local imports from .ts to .js to resolve TS5097
import { Dynamics365FO } from "../src/main.js"; 
import { registerTools } from "../src/tools.js";

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

const app = express();
app.use(cors());
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

// --- REMOTE MCP TRANSPORT SETUP ---

// Wrap the asynchronous setup logic in a self-executing async function (IIFE)
(async () => {
    try {
        const mcpTransport = new StreamableHTTPServerTransport({ 
            sessionIdGenerator: randomUUID, 
        });
        
        await server.connect(mcpTransport); 

        // Use 'as any' to bypass the type error (2339) and expose the handler
        app.use("/mcp", (mcpTransport as any).requestHandler());
        
        // Start the Express server after successful MCP setup
        app.listen(port, () => {
            console.error(`FO HTTP + MCP wrapper listening at http://localhost:${port}`);
            console.error("MCP remote server endpoint is ready at /mcp");
        });

    } catch (err) {
        console.error("FATAL ERROR: Failed to start MCP server or listen to port:", err);
    }
})();