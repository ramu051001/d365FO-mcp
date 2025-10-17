/**
 * Dynamics 365 Finance & Operations (FO) helper
 *
 * - Authenticates with Azure AD using client credentials (MSAL)
 * - Calls FO OData endpoints under /data (e.g., data/CustomersV3)
 * - Supports OData query building ($filter, $select, $top, $orderby, extra)
 * - Supports following @odata.nextLink (fetchAllPages)
 * - Masks tokens in logs
 */

import { ConfidentialClientApplication } from "@azure/msal-node";
import type { Configuration, ClientCredentialRequest } from "@azure/msal-node";
import dotenv from "dotenv";

dotenv.config();

export class Dynamics365FO {
  private clientId: string;
  private clientSecret: string;
  private tenantId: string;
  // This variable will be the cleaned, base D365 URL (no trailing slash, no quotes).
  private d365BaseUrl: string; 
  private msalInstance: ConfidentialClientApplication;
  private accessToken: string | null = null;
  private tokenExpiration: number | null = null;

  /**
   * @param clientId - Azure AD application client ID
   * @param clientSecret - Azure AD application client secret
   * @param tenantId - Azure AD tenant ID
   * @param d365Url - FO environment base URL (e.g., https://org.cloudax.dynamics.com)
   */
  constructor(
    clientId: string,
    clientSecret: string,
    tenantId: string,
    d365Url: string
  ) {
    this.clientId = clientId;
    this.clientSecret = clientSecret;
    this.tenantId = tenantId;
    
    // **CRITICAL FIX:** Guarantee a clean base URL here. 
    // This removes surrounding quotes, leading/trailing whitespace, and any trailing commas or slashes.
    let cleanedUrl = d365Url.trim().replace(/^['"]|['"]$/g, "");
    // Aggressively remove all trailing slashes and commas
    cleanedUrl = cleanedUrl.replace(/[,/]+$/, ""); 
    
    this.d365BaseUrl = cleanedUrl;

    const msalConfig: Configuration = {
      auth: {
        clientId: this.clientId,
        authority: `https://login.microsoftonline.com/${this.tenantId}`,
        clientSecret: this.clientSecret,
      },
    };

    this.msalInstance = new ConfidentialClientApplication(msalConfig);
  }

  /**
   * Acquire or reuse an application access token.
   * This is the function that was updated to fix the 'invalid_scope' error.
   */
  private async authenticate(): Promise<string> {
    try {
      // 1. Use the cleaned d365BaseUrl property
      const resourceOrigin = this.d365BaseUrl;
      
      const tokenRequest: ClientCredentialRequest = {
        // The scope must be constructed as 'https://<D365_BASE_URL>/.default' 
        // We use the cleaned resourceOrigin which does NOT have a trailing slash or comma.
        scopes: [`${resourceOrigin}/.default`],
      };

      if (this.tokenExpiration && Date.now() < this.tokenExpiration) {
        // console.error("✅ Reusing cached token.");
        return this.accessToken as string;
      }

      const response = await this.msalInstance.acquireTokenByClientCredential(
        tokenRequest
      );

      if (response && response.accessToken) {
        this.accessToken = response.accessToken;
        if (response.expiresOn) {
          // Renew a few minutes before expiry
          this.tokenExpiration = response.expiresOn.getTime() - 3 * 60 * 1000;
        }
        console.error("✅ Token acquired successfully and cached.");
        return this.accessToken as string;
      } else {
        throw new Error("Token acquisition failed: response is null or invalid.");
      }
    } catch (error) {
      // Log to stderr
      console.error("❌ Failed to authenticate with Azure AD:", error);
      throw new Error(
        `Failed to authenticate with Dynamics 365 FO: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
  }

  /**
   * Makes an API request to FO.
   */
  private async makeApiRequest(
    endpoint: string,
    method: string = "GET",
    body?: any,
    additionalHeaders?: Record<string, string>
  ): Promise<any> {
    const token = await this.authenticate();

    // ✅ Clean URL handling (no double slashes)
    // d365BaseUrl is guaranteed to have NO trailing slash.
    const url = endpoint.startsWith("http")
      ? endpoint
      : `${this.d365BaseUrl}/${endpoint.replace(/^\//, "")}`;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "Content-Type": "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      ...additionalHeaders,
    };

    // Mask token for logs
    const masked = { ...headers, Authorization: "Bearer [REDACTED]" };
    console.error(`[Dynamics365FO] Request: ${method} ${url}`);
    console.error("[Dynamics365FO] Headers:", masked);
    if (body) console.error("[Dynamics365FO] Body:", body);

    try {
      const response = await fetch(url, {
        method,
        headers,
        body: body ? JSON.stringify(body) : undefined,
      });

      const text = await response.text().catch(() => null);

      if (!response.ok) {
        console.error(`[Dynamics365FO] Request failed: ${response.status} ${response.statusText}`);
        console.error(`[Dynamics365FO] URL: ${url}`);
        console.error(`[Dynamics365FO] Response body:`, text);
        throw new Error(`API request failed: ${response.status} ${response.statusText} - ${text}`);
      }

      if (!text) return null;

      // ✅ Detect HTML payloads (login pages, redirects) - This is the original problem spot.
      if (text.startsWith("<!DOCTYPE html>") || text.includes("<html")) {
        console.error(`[Dynamics365FO] ERROR: Received HTML instead of JSON. This suggests a D365 FO access/permission issue, even with status 200.`);
        console.error(`[Dynamics365FO] HTML snippet: ${text.substring(0, 300)}...`);
        throw new Error(`Response OK but received non-JSON payload: ${text.substring(0, 100)}...`);
      }

      // ✅ Parse as JSON
      try {
        return JSON.parse(text);
      } catch {
        throw new Error(`Response OK but received non-JSON payload: ${text.substring(0, 100)}...`);
      }
    } catch (err) {
      console.error(`[Dynamics365FO] API request to ${url} failed:`, err);
      throw err;
    }
  }

  /**
   * Build an OData query string from options.
   */
  private buildODataQuery(opts?: {
    filter?: string;
    select?: string[];
    top?: number;
    orderby?: string;
    extra?: Record<string, string>;
  }): string {
    const parts: string[] = [];

    if (opts?.filter) parts.push(`$filter=${encodeURIComponent(opts.filter)}`);
    if (opts?.select?.length) parts.push(`$select=${encodeURIComponent(opts.select.join(","))}`);
    if (opts?.top && opts.top > 0) parts.push(`$top=${opts.top}`);
    if (opts?.orderby) parts.push(`$orderby=${encodeURIComponent(opts.orderby)}`);

    if (opts?.extra) {
      for (const [k, v] of Object.entries(opts.extra)) {
        parts.push(`${k}=${encodeURIComponent(v)}`);
      }
    }

    return parts.join("&");
  }

  /**
   * Fetch customers (supports pagination).
   */
  public async getCustomers(options?: {
    filter?: string;
    select?: string[];
    top?: number;
    orderby?: string;
    crossCompany?: boolean;
    fetchAllPages?: boolean;
  }): Promise<any> {
    const baseEndpoint = "data/CustomersV3";
    const extraParams: Record<string, string> = {};

    if (options?.crossCompany) extraParams["cross-company"] = "true";

    const query = this.buildODataQuery({
      filter: options?.filter,
      select: options?.select,
      top: options?.top,
      orderby: options?.orderby,
      extra: extraParams,
    });

    const endpointWithQuery = query ? `${baseEndpoint}?${query}` : baseEndpoint;
    if (!options?.fetchAllPages) {
      return this.makeApiRequest(endpointWithQuery, "GET");
    }

    const results: any[] = [];
    let resp = await this.makeApiRequest(endpointWithQuery, "GET");

    if (Array.isArray(resp?.value)) results.push(...resp.value);
    else if (Array.isArray(resp)) results.push(...resp);
    else return resp;

    let nextLink = resp["@odata.nextLink"] || resp["@odata.nextlink"];
    while (nextLink) {
      const page = await this.makeApiRequest(nextLink, "GET");
      if (Array.isArray(page?.value)) results.push(...page.value);
      else if (Array.isArray(page)) results.push(...page);
      else break;
      
      nextLink = page["@odata.nextLink"] || page["@odata.nextlink"];
    }

    return { value: results };
  }

  public async getCustomersByName(name: string, options?: {
    select?: string[];
    top?: number;
    crossCompany?: boolean;
    fetchAllPages?: boolean;
  }): Promise<any> {
    const safe = name.replace(/'/g, "''");
    const filter = `contains(Name,'${safe}')`;
    return this.getCustomers({ ...options, filter });
  }

  public async getCustomerByAccountNum(accountNum: string, options?: {
    select?: string[];
    crossCompany?: boolean;
  }): Promise<any> {
    const safe = accountNum.replace(/'/g, "''");
    const filter = `CustomerAccount eq '${safe}'`;
    return this.getCustomers({ ...options, filter, top: 1 });
  }

  public async getVendors(options?: {
    filter?: string;
    select?: string[];
    top?: number;
    orderby?: string;
    crossCompany?: boolean;
    fetchAllPages?: boolean;
  }): Promise<any> {
    const baseEndpoint = "data/VendorsV3";
    const extraParams: Record<string, string> = {};
    if (options?.crossCompany) extraParams["cross-company"] = "true";

    const query = this.buildODataQuery({
      filter: options?.filter,
      select: options?.select,
      top: options?.top,
      orderby: options?.orderby,
      extra: extraParams,
    });

    const endpointWithQuery = query ? `${baseEndpoint}?${query}` : baseEndpoint;
    if (!options?.fetchAllPages) return this.makeApiRequest(endpointWithQuery, "GET");

    const results: any[] = [];
    let resp = await this.makeApiRequest(endpointWithQuery, "GET");

    if (Array.isArray(resp?.value)) results.push(...resp.value);
    else if (Array.isArray(resp)) results.push(...resp);
    else return resp;

    let nextLink = resp["@odata.nextLink"] || resp["@odata.nextlink"];
    while (nextLink) {
      const page = await this.makeApiRequest(nextLink, "GET");
      if (Array.isArray(page?.value)) results.push(...page.value);
      else if (Array.isArray(page)) results.push(...page);
      else break;
      
      nextLink = page["@odata.nextLink"] || page["@odata.nextlink"];
    }

    return { value: results };
  }

  public async getVendorByAccountNum(vendorAccount: string, options?: {
    select?: string[];
    crossCompany?: boolean;
  }): Promise<any> {
    const safe = vendorAccount.replace(/'/g, "''");
    const filter = `VendorAccountNumber eq '${safe}'`;
    return this.getVendors({ ...options, filter, top: 1 });
  }
}
