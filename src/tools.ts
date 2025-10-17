import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { Dynamics365FO } from "../src/main.js";

/**
 * Register MCP tools:
 * - fetch-accounts    -> customer lookup / list (configurable entityType: customer|vendor|auto)
 * - fetch-vendors     -> vendor lookup / list
 * - fetch-entity-by-id-> explicit combined search (returns both if present)
 */
export function registerTools(server: McpServer, d365: Dynamics365FO) {
  // Zod "raw shape" (map of validators) — Server.tool expects this shape
  const fetchEntitiesShape = {
    accountNum: z.string().optional(),
    filter: z.string().optional(),
    select: z.array(z.string()).optional(),
    top: z.number().int().positive().optional(),
    orderby: z.string().optional(),
    crossCompany: z.boolean().optional(),
    fetchAllPages: z.boolean().optional(),
    // explicit hint: 'customer' | 'vendor' | 'auto' (auto = default: customer then vendor fallback)
    entityType: z.enum(["customer", "vendor", "auto"]).optional(),
  };

  // small helper to normalize possible response shapes into a single record (or null)
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

  // Fetch customers tool (with configurable entityType and vendor fallback when configured)
  server.tool(
    "fetch-accounts",
    "Fetch accounts from Dynamics 365 FO. Provide accountNum for single lookup or OData options for listing. Optional input.entityType = 'customer'|'vendor'|'auto' (default 'auto').",
    fetchEntitiesShape,
    async (input: any) => {
      try {
        const entityType: "customer" | "vendor" | "auto" =
          (input?.entityType as "customer" | "vendor" | "auto") ?? "auto";

        // If caller specified an account number, do single-id behavior (with ordering based on entityType)
        if (input && input.accountNum) {
          const id = input.accountNum;
          // Helper to run customer lookup
          const lookupCustomer = async () => {
            const resp = await d365.getCustomerByAccountNum(id, {
              select: input.select,
              crossCompany: input.crossCompany,
            });
            return normalizeSingleRecord(resp);
          };
          // Helper to run vendor lookup
          const lookupVendor = async () => {
            const resp = await d365.getVendorByAccountNum(id, {
              select: input.select,
              crossCompany: input.crossCompany,
            });
            return normalizeSingleRecord(resp);
          };

          if (entityType === "vendor") {
            // vendor-first
            const vendorRecord = await lookupVendor();
            if (vendorRecord) {
              return {
                content: [
                  {
                    type: "text",
                    text: JSON.stringify({ type: "vendor", record: vendorRecord }, null, 2),
                  },
                ],
              };
            }
            // if vendor-first didn't find anything, optionally try customer as fallback
            const customerRecord = await lookupCustomer();
            if (customerRecord) {
              return {
                content: [
                  {
                    type: "text",
                    text: JSON.stringify({ type: "customer", record: customerRecord }, null, 2),
                  },
                ],
              };
            }

            return {
              content: [
                {
                  type: "text",
                  text: `No vendor or customer found for AccountNum = ${id}`,
                },
              ],
            };
          } else if (entityType === "customer") {
            // customer-first
            const customerRecord = await lookupCustomer();
            if (customerRecord) {
              return {
                content: [
                  {
                    type: "text",
                    text: JSON.stringify({ type: "customer", record: customerRecord }, null, 2),
                  },
                ],
              };
            }
            // try vendor fallback
            const vendorRecord = await lookupVendor();
            if (vendorRecord) {
              return {
                content: [
                  {
                    type: "text",
                    text: JSON.stringify({ type: "vendor", record: vendorRecord }, null, 2),
                  },
                ],
              };
            }
            return {
              content: [
                {
                  type: "text",
                  text: `No customer or vendor found for AccountNum = ${id}`,
                },
              ],
            };
          } else {
            // auto: existing default behavior (customer then vendor fallback)
            const customerRecord = await lookupCustomer();
            if (customerRecord) {
              return {
                content: [
                  {
                    type: "text",
                    text: JSON.stringify({ type: "customer", record: customerRecord }, null, 2),
                  },
                ],
              };
            }
            const vendorRecord = await lookupVendor();
            if (vendorRecord) {
              return {
                content: [
                  {
                    type: "text",
                    text: JSON.stringify({ type: "vendor", record: vendorRecord }, null, 2),
                  },
                ],
              };
            }
            return {
              content: [
                {
                  type: "text",
                  text: `No customer or vendor found for AccountNum = ${id}`,
                },
              ],
            };
          }
        }

        // Otherwise use general getCustomers with provided OData options (list) - entityType not applied for lists
        // If you want vendor listing via this tool, pass entityType='vendor' and call fetch-vendors instead (explicit).
        const response = await d365.getCustomers({
          filter: input?.filter,
          select: input?.select,
          top: input?.top,
          orderby: input?.orderby,
          crossCompany: input?.crossCompany,
          fetchAllPages: input?.fetchAllPages,
        });

        // Prefer response.value if present
        const payload = response && (response.value ?? response);
        const text = JSON.stringify(payload, null, 2);

        return {
          content: [
            {
              type: "text",
              text,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Error: ${error instanceof Error ? error.message : "Unknown error"}.`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Fetch vendors tool (explicit vendor lookups / listing)
  server.tool(
    "fetch-vendors",
    "Fetch vendors from Dynamics 365 FO. Optionally filter by accountNum or provide OData options.",
    fetchEntitiesShape,
    async (input: any) => {
      try {
        if (input && input.accountNum) {
          const resp = await d365.getVendorByAccountNum(input.accountNum, {
            select: input.select,
            crossCompany: input.crossCompany,
          });

          const record = normalizeSingleRecord(resp);

          if (!record) {
            return {
              content: [
                {
                  type: "text",
                  text: `No vendor found for VendorAccount = ${input.accountNum}`,
                },
              ],
            };
          }

          return {
            content: [
              {
                type: "text",
                text: JSON.stringify(record, null, 2),
              },
            ],
          };
        }

        const response = await d365.getVendors({
          filter: input?.filter,
          select: input?.select,
          top: input?.top,
          orderby: input?.orderby,
          crossCompany: input?.crossCompany,
          fetchAllPages: input?.fetchAllPages,
        });

        const payload = response && (response.value ?? response);
        const text = JSON.stringify(payload, null, 2);

        return {
          content: [
            {
              type: "text",
              text,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Error: ${error instanceof Error ? error.message : "Unknown error"}.`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Explicit combined search tool: returns both matches if both exist (customer + vendor),
  // otherwise returns the one found. Clear and explicit for callers that want both checked.
  server.tool(
    "fetch-entity-by-id",
    "Search customers and vendors for the given account identifier and return matched records with type tags.",
    fetchEntitiesShape,
    async (input: any) => {
      try {
        const id = input?.accountNum;
        if (!id) {
          return {
            content: [
              {
                type: "text",
                text: "Please provide accountNum in the input payload.",
              },
            ],
          };
        }

        // Kick off both lookups in parallel (faster)
        const [custResp, vendResp] = await Promise.allSettled([
          d365.getCustomerByAccountNum(id, { select: input?.select, crossCompany: input?.crossCompany }),
          d365.getVendorByAccountNum(id, { select: input?.select, crossCompany: input?.crossCompany }),
        ]);

        const cust = custResp.status === "fulfilled" ? normalizeSingleRecord(custResp.value) : null;
        const vend = vendResp.status === "fulfilled" ? normalizeSingleRecord(vendResp.value) : null;

        if (cust && vend) {
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({ customer: cust, vendor: vend }, null, 2),
              },
            ],
          };
        } else if (cust) {
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({ type: "customer", record: cust }, null, 2),
              },
            ],
          };
        } else if (vend) {
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({ type: "vendor", record: vend }, null, 2),
              },
            ],
          };
        } else {
          return {
            content: [
              {
                type: "text",
                text: `No customer or vendor found for ${id}`,
              },
            ],
          };
        }
      } catch (err) {
        return {
          content: [
            {
              type: "text",
              text: `Error during lookup: ${err instanceof Error ? err.message : String(err)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );
}