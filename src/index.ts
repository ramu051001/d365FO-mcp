import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { Dynamics365FO } from "./main.js";
import { registerTools } from "./tools.js";
import dotenv from "dotenv";

// Load environment variables from .env file
dotenv.config();

// Create server instance
const server = new McpServer({
  name: "Dynamics365FO",
  version: "1.0.0.0",
});

const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const tenantId = process.env.TENANT_ID;
const D365_BASE_URL = process.env.D365_URL;

if (!clientId || !clientSecret || !tenantId || !D365_BASE_URL) {
  console.error(
    "Missing required environment variables. Please check your .env file."
  );
  process.exit(1);
}

// Instantiate the FO-specific class
const d365 = new Dynamics365FO(clientId, clientSecret, tenantId, D365_BASE_URL);

// Register all tools
registerTools(server, d365);

// Start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Dynamics365 FO MCP server running on stdio...");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});