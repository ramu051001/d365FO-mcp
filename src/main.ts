/**
 * Dynamics 365 Finance & Operations (FO) helper
 *
 * - Authenticates with Azure AD using client credentials (MSAL)
 * - Calls FO OData endpoints under /data (e.g., data/CustomersV3)
 * - Supports OData query building ($filter, $select, $top, $orderby, extra)
 * - Supports following @odata.nextLink (fetchAllPages)
 * - Masks tokens in logs
 *
 * Note: This file is FO-specific (no Dataverse/CE code paths).
 */

import { ConfidentialClientApplication } from "@azure/msal-node";
import type { Configuration, ClientCredentialRequest } from "@azure/msal-node";
import dotenv from "dotenv";

dotenv.config();

export class Dynamics365FO {
  private clientId: string;
  private clientSecret: string;
  private tenantId: string;
  private d365Url: string;
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
    this.d365Url = d365Url;

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
   * Scope uses the origin of d365Url (example: https://org.cloudax.dynamics.com/.default)
   */
  private async authenticate(): Promise<string> {
    const tokenRequest: ClientCredentialRequest = {
      scopes: [`${new URL(this.d365Url).origin}/.default`],
    };

    try {
      if (this.tokenExpiration && Date.now() < this.tokenExpiration) {
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
      } else {
        throw new Error("Token acquisition failed: response is null or invalid.");
      }
    } catch (error) {
      // log to stderr (won't corrupt protocol)
      console.error("Token acquisition failed:", error);
      throw new Error(
        `Failed to authenticate with Dynamics 365 FO: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }

    return this.accessToken as string;
  }

  /**
   * Makes an API request to FO.
   * Accepts either a relative endpoint (e.g. "data/CustomersV3?$top=5")
   * or an absolute URL (e.g. @odata.nextLink).
   */
  private async makeApiRequest(
    endpoint: string,
    method: string = "GET",
    body?: any,
    additionalHeaders?: Record<string, string>
  ): Promise<any> {
    const token = await this.authenticate();

    const url = endpoint.match(/^https?:\/\//i)
      ? endpoint
      : `${this.d365Url.endsWith("/") ? this.d365Url.slice(0, -1) : this.d365Url}/${endpoint.replace(/^\//, "")}`;

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

    // IMPORTANT: write logs to stderr so stdout (JSON-RPC) stays clean
    console.error(`[Dynamics365FO] Request: ${method} ${url}`);
    console.error("[Dynamics365FO] Headers:", masked);
    if (body) console.error("[Dynamics365FO] Body:", body);

    try {
      const response = await fetch(url, {
        method,
        headers,
        body: body ? JSON.stringify(body) : undefined,
      });

      // Read text so we can log non-JSON responses in errors
      const text = await response.text().catch(() => null);

      if (!response.ok) {
        console.error(`[Dynamics365FO] Request failed: ${response.status} ${response.statusText}`);
        console.error(`[Dynamics365FO] URL: ${url}`);
        console.error(`[Dynamics365FO] Response body:`, text);
        try {
          const rawHeaders: Record<string, string> = {};
          response.headers.forEach((v, k) => (rawHeaders[k] = v));
          console.error("[Dynamics365FO] Response headers:", rawHeaders);
        } catch (_) {
          // no-op
        }
        throw new Error(`API request failed: ${response.status} ${response.statusText} - ${text}`);
      }

      if (!text) return null;
      return JSON.parse(text);
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
    if (opts?.select && opts.select.length) parts.push(`$select=${encodeURIComponent(opts.select.join(","))}`);
    if (opts?.top && opts.top > 0) parts.push(`$top=${opts.top}`);
    if (opts?.orderby) parts.push(`$orderby=${encodeURIComponent(opts.orderby)}`);

    if (opts?.extra) {
      for (const [k, v] of Object.entries(opts.extra)) {
        parts.push(`${k}=${encodeURIComponent(v)}`);
      }
    }

    return parts.length ? parts.join("&") : "";
  }

  /**
   * Fetches customers from Dynamics 365 FO (CustomersV3).
   *
   * Options:
   *  - filter: OData $filter expression (e.g. "contains(Name,'Contoso')")
   *  - select: array of fields to $select
   *  - top: $top
   *  - orderby: $orderby expression
   *  - crossCompany: adds cross-company=true query param when true
   *  - fetchAllPages: follows @odata.nextLink to combine pages (beware large datasets)
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
    if (options?.crossCompany) {
      extraParams["cross-company"] = "true";
    }

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

    if (Array.isArray(resp?.value)) {
      results.push(...resp.value);
    } else if (Array.isArray(resp)) {
      results.push(...resp);
    } else {
      return resp;
    }

    let nextLink = resp["@odata.nextLink"] || resp["@odata.nextlink"] || null;
    while (nextLink) {
      const page = await this.makeApiRequest(nextLink, "GET");
      if (Array.isArray(page?.value)) {
        results.push(...page.value);
      } else if (Array.isArray(page)) {
        results.push(...page);
      } else {
        break;
      }
      nextLink = page["@odata.nextLink"] || page["@odata.nextlink"] || null;
    }

    return { value: results };
  }

  /**
   * Convenience: fetch customers by name (simple contains).
   * Uses FO field 'Name' by default.
   */
  public async getCustomersByName(name: string, options?: {
    select?: string[];
    top?: number;
    crossCompany?: boolean;
    fetchAllPages?: boolean;
  }): Promise<any> {
    const safe = name.replace(/'/g, "''");
    const filter = `contains(Name,'${safe}')`;

    return this.getCustomers({
      ...options,
      filter,
    });
  }

  /**
   * Convenience: fetch customer by account identifier (uses FO field 'CustomerAccount').
   *
   * Note: FO schema uses `CustomerAccount` for the account identifier (not AccountNum).
   */
  public async getCustomerByAccountNum(accountNum: string, options?: {
    select?: string[];
    crossCompany?: boolean;
  }): Promise<any> {
    // Use the FO field name 'CustomerAccount' for the equality filter
    const safe = accountNum.replace(/'/g, "''");
    const filter = `CustomerAccount eq '${safe}'`;

    return this.getCustomers({
      ...options,
      filter,
      top: 1,
    });
  }

  /**
   * Fetch vendors from Dynamics 365 FO.
   *
   * - collectionName: "Vendors" by default. Some environments use "VendorsV3" — change if required.
   * - identifier field assumed to be "VendorAccount".
   */
  public async getVendors(options?: {
    filter?: string;
    select?: string[];
    top?: number;
    orderby?: string;
    crossCompany?: boolean;
    fetchAllPages?: boolean;
  }): Promise<any> {
    // Change this value if your FO metadata shows a different entity set like "VendorsV3"
    const baseEndpoint = "data/VendorsV3";

    const extraParams: Record<string, string> = {};
    if (options?.crossCompany) {
      extraParams["cross-company"] = "true";
    }

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

    if (Array.isArray(resp?.value)) {
      results.push(...resp.value);
    } else if (Array.isArray(resp)) {
      results.push(...resp);
    } else {
      return resp;
    }

    let nextLink = resp["@odata.nextLink"] || resp["@odata.nextlink"] || null;
    while (nextLink) {
      const page = await this.makeApiRequest(nextLink, "GET");
      if (Array.isArray(page?.value)) {
        results.push(...page.value);
      } else if (Array.isArray(page)) {
        results.push(...page);
      } else {
        break;
      }
      nextLink = page["@odata.nextLink"] || page["@odata.nextlink"] || null;
    }

    return { value: results };
  }

  // Replace the existing getVendorByAccountNum implementation with this:

  /**
   * Convenience: fetch vendor by vendor account identifier (uses FO field 'VendorAccountNumber').
   */
  public async getVendorByAccountNum(vendorAccount: string, options?: {
    select?: string[];
    crossCompany?: boolean;
  }): Promise<any> {
    // Use the FO field name 'VendorAccountNumber' for the equality filter
    const safe = vendorAccount.replace(/'/g, "''");
    const filter = `VendorAccountNumber eq '${safe}'`;

    return this.getVendors({
      ...options,
      filter,
      top: 1,
    });
  }
  
}