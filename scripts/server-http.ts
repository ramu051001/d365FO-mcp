/**
 * Combined HTTP wrapper + MCP server for REMOTE DEPLOYMENT (Azure/Web)
 *
 * - Starts an MCP server using StreamableHTTPServerTransport and exposes it at /mcp
 * - Also exposes a small HTTP API for testing.
 */

import dotenv from "dotenv";
dotenv.config();

import express, { Request, Response, NextFunction } from "express";
import cors from "cors";
import { randomUUID } from "node:crypto";

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";

// FIX: Change imports to .js extensions for compiled environment
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

// -------------------------
// Helper functions
// -------------------------

function normalizeSingleRecord(resp: any): any | null {
  if (!resp) return null;
  if (Array.isArray(resp)) return resp.length ? resp[0] : null;
  if (Array.isArray(resp?.value)) return resp.value.length ? resp.value[0] : null;
  if (typeof resp === "object") {
    if (resp.value && Array.isArray(resp.value) && resp.value.length) return resp.value[0];
    return Object.keys(resp).length ? resp : null;
  }
  return null;
}

// Wrap async routes to properly catch errors
const asyncHandler = (fn: (req: Request, res: Response, next: NextFunction) => Promise<any>) =>
  (req: Request, res: Response, next: NextFunction) => {
    Promise.resolve(fn(req, res, next)).catch(next);
  };

// -------------------------
// EXPRESS ROUTES
// -------------------------

// GET single customer by account identifier (CustomerAccount)
app.get(
  "/customers/:account",
  asyncHandler(async (req: Request, res: Response) => {
    const account = req.params.account;
    const crossCompany = req.query.crossCompany === "true" || req.query["cross-company"] === "true";
    const fields = (req.query.$select ?? req.query.select)
      ? String(req.query.$select ?? req.query.select).split(",")
      : undefined;

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
  })
);

// Generic customers endpoint supporting raw OData params ($filter/$select) or plain names
app.get(
  "/customers",
  asyncHandler(async (req: Request, res: Response) => {
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
  })
);

// GET single vendor by account identifier (VendorAccount)
app.get(
  "/vendors/:vendorAccount",
  asyncHandler(async (req: Request, res: Response) => {
    const vendorAccount = req.params.vendorAccount;
    const crossCompany = req.query.crossCompany === "true" || req.query["cross-company"] === "true";
    const fields = (req.query.$select ?? req.query.select)
      ? String(req.query.$select ?? req.query.select).split(",")
      : undefined;

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
  })
);

// Generic vendors endpoint supporting raw OData params ($filter/$select)
app.get(
  "/vendors",
  asyncHandler(async (req: Request, res: Response) => {
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
  })
);

// -------------------------
// REMOTE MCP TRANSPORT
// -------------------------

(async () => {
  try {
    const mcpTransport = new StreamableHTTPServerTransport({
      sessionIdGenerator: randomUUID,
    });

    await server.connect(mcpTransport);

    // Express-compatible middleware
    app.use("/mcp", (req: Request, res: Response, next: NextFunction) => {
      const handler = (mcpTransport as any).requestHandler || (mcpTransport as any).middleware;
      if (typeof handler === "function") {
        Promise.resolve(handler.call(mcpTransport, req, res, next)).catch(next);
      } else {
        res.status(500).send("MCP handler unavailable");
      }
    });

    app.listen(port, () => {
      console.error(`FO HTTP + MCP wrapper listening at http://localhost:${port}`);
      console.error("MCP remote server endpoint is ready at /mcp");
    });
  } catch (err) {
    console.error("FATAL ERROR: Failed to start MCP server or listen to port:", err);
    process.exit(1);
  }
})();
