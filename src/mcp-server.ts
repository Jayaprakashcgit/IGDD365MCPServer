import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { makeApiCall } from './api.js';
import { EntityManager } from './entityManager.js';
import { RequestHandlerExtra } from '@modelcontextprotocol/sdk/shared/protocol.js';
import { ServerRequest, ServerNotification } from '@modelcontextprotocol/sdk/types.js';

const entityManager = new EntityManager();
// PAGINATION: Define a default page size to prevent oversized payloads.
const DEFAULT_PAGE_SIZE = 5;

// Helper function to safely send notifications
async function safeNotification(context: RequestHandlerExtra<ServerRequest, ServerNotification>, notification: any): Promise<void> {
    try {
        await context.sendNotification(notification);
    } catch (error) {
        console.log('Notification failed (this is normal in test environments):', error);
    }
}

// Helper function to build the OData $filter string from an object
function buildFilterString(filterObject?: Record<string, string>): string | null {
    if (!filterObject || Object.keys(filterObject).length === 0) {
        return null;
    }
    const filterClauses = Object.entries(filterObject).map(([key, value]) => {
        return `${key} eq '${value}'`;
    });
    return filterClauses.join(' and ');
}


// --- Zod Schemas for Tool Arguments ---

const odataQuerySchema = z.object({
    entity: z.string().describe("The OData entity set to query (e.g., CustomersV3, ReleasedProductsV2)."),
    select: z.string().optional().describe("OData $select query parameter to limit the fields returned."),
    filter: z.record(z.string()).optional().describe("Key-value pairs for filtering. e.g., { ProductNumber: 'D0001', dataAreaId: 'usmf' }."),
    expand: z.string().optional().describe("OData $expand query parameter."),
    // PAGINATION: Updated description for 'top' to explain its role in pagination.
    top: z.number().optional().describe(`The number of records to return per page. Defaults to ${DEFAULT_PAGE_SIZE}.`),
    // PAGINATION: Added 'skip' parameter for fetching subsequent pages.
    skip: z.number().optional().describe("The number of records to skip. Used for pagination to get the next set of results."),
    crossCompany: z.boolean().optional().describe("Set to true to query across all companies."),
});

const createCustomerSchema = z.object({
    customerData: z.record(z.unknown()).describe("A JSON object for the new customer. Must include dataAreaId, CustomerAccount, etc."),
});

const updateCustomerSchema = z.object({
    dataAreaId: z.string().describe("The dataAreaId of the customer (e.g., 'usmf')."),
    customerAccount: z.string().describe("The customer account ID to update (e.g., 'PM-001')."),
    updateData: z.record(z.unknown()).describe("A JSON object with the fields to update."),
});

const getEntityCountSchema = z.object({
    entity: z.string().describe("The OData entity set to count (e.g., CustomersV3)."),
    crossCompany: z.boolean().optional().describe("Set to true to count across all companies."),
});

const createSystemUserSchema = z.object({
     userData: z.record(z.unknown()).describe("A JSON object for the new system user. Must include UserID, Alias, Company, etc."),
});

const assignUserRoleSchema = z.object({
    associationData: z.record(z.unknown()).describe("JSON object for the role association. Must include UserId and SecurityRoleIdentifier."),
});

const updatePositionHierarchySchema = z.object({
    positionId: z.string().describe("The ID of the position to update."),
    hierarchyTypeName: z.string().describe("The hierarchy type name (e.g., 'Line')."),
    validFrom: z.string().datetime().describe("The start validity date in ISO 8601 format."),
    validTo: z.string().datetime().describe("The end validity date in ISO 8601 format."),
    updateData: z.record(z.unknown()).describe("A JSON object with the fields to update (e.g., ParentPositionId)."),
});
//Added by JP Start
const createSalesOrderSchema = z.object({
  header: z.object({
    dataAreaId: z.string().describe("Legal entity (e.g., 'usmf')."),
    RequestedShippingDate: z.string().describe("Requested shipping date (ISO 8601 recommended, e.g., '2025-10-20')."),
    CustomerAccount: z.string().describe("Customer account (e.g., 'US-001')."),
    SalesOrderNumber: z.string().optional().describe("Optional. If omitted, D365 will auto-assign.")
  }),
  lines: z.array(z.object({
    ItemNumber: z.string().describe("Item number to add as a line."),
    OrderedSalesQuantity: z.number().describe("Ordered quantity for the line."),
    SiteId: z.string().optional().describe("Optional. If omitted, system defaulting fills it when configured.")
  }))
  .min(1)
  .describe("At least one line.")
});
//Added by JP End
/**
 * Creates and configures the MCP server with all the tools for the D365 API.
 * @returns {McpServer} The configured McpServer instance. 
 */
export const getServer = (): McpServer => {
    const server = new McpServer({
        name: 'd365-fno-mcp-server',
        version: '1.0.0',
    });

    // --- Tool Definitions ---

    server.tool(
        'odataQuery',
        'Executes a generic GET request against a Dynamics 365 OData entity. The entity name does not need to be case-perfect. Responses are paginated.',
        odataQuerySchema.shape,
        async (args: z.infer<typeof odataQuerySchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            
            const correctedEntity = await entityManager.findBestMatch(args.entity);
    
            if (!correctedEntity) {
                return {
                    isError: true,
                    content: [{ type: 'text', text: `Could not find a matching entity for '${args.entity}'. Please provide a more specific name.` }]
                };
            }
            
            const effectiveArgs = { ...args };

            if (effectiveArgs.filter?.dataAreaId && effectiveArgs.crossCompany !== false) {
                if (!effectiveArgs.crossCompany) {
                    await safeNotification(context, {
                        method: "notifications/message",
                        params: { level: "info", data: `Filter on company ('dataAreaId') detected. Automatically enabling cross-company search.` }
                    });
                }
                effectiveArgs.crossCompany = true;
            }

            await safeNotification(context, {
                method: "notifications/message",
                params: { level: "info", data: `Corrected entity name from '${args.entity}' to '${correctedEntity}'.` }
            });
            
            const { entity, ...queryParams } = effectiveArgs;
            const filterString = buildFilterString(queryParams.filter);
            const url = new URL(`${process.env.DYNAMICS_RESOURCE_URL}/data/${correctedEntity}`);

            // PAGINATION: Apply query parameters including the new skip and a default top.
            const topValue = queryParams.top || DEFAULT_PAGE_SIZE;
            url.searchParams.append('$top', topValue.toString());

            if (queryParams.skip) {
                url.searchParams.append('$skip', queryParams.skip.toString());
            }

            if (queryParams.crossCompany) url.searchParams.append('cross-company', 'true');
            if (queryParams.select) url.searchParams.append('$select', queryParams.select);
            if (filterString) url.searchParams.append('$filter', filterString);
            if (queryParams.expand) url.searchParams.append('$expand', queryParams.expand);
            
            return makeApiCall('GET', url.toString(), null, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'createCustomer',
        'Creates a new customer record in CustomersV3.',
        createCustomerSchema.shape,
        async ({ customerData }: z.infer<typeof createCustomerSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/CustomersV3`;
            return makeApiCall('POST', url, customerData as Record<string, unknown>, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'updateCustomer',
        'Updates an existing customer record in CustomersV3 using a PATCH request.',
        updateCustomerSchema.shape,
        async ({ dataAreaId, customerAccount, updateData }: z.infer<typeof updateCustomerSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/CustomersV3(dataAreaId='${dataAreaId}',CustomerAccount='${customerAccount}')`;
            return makeApiCall('PATCH', url, updateData as Record<string, unknown>, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'getEntityCount',
        'Gets the total count of records for a given OData entity.',
        getEntityCountSchema.shape,
        async ({ entity, crossCompany }: z.infer<typeof getEntityCountSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
             const url = new URL(`${process.env.DYNAMICS_RESOURCE_URL}/data/${entity}/$count`);
             if (crossCompany) url.searchParams.append('cross-company', 'true');
             return makeApiCall('GET', url.toString(), null, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'createSystemUser',
        'Creates a new user in SystemUsers.',
        createSystemUserSchema.shape,
        async ({ userData }: z.infer<typeof createSystemUserSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/SystemUsers`;
            return makeApiCall('POST', url, userData as Record<string, unknown>, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'assignUserRole',
        'Assigns a security role to a user in SecurityUserRoleAssociations.',
        assignUserRoleSchema.shape,
        async ({ associationData }: z.infer<typeof assignUserRoleSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/SecurityUserRoleAssociations`;
            return makeApiCall('POST', url, associationData as Record<string, unknown>, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'updatePositionHierarchy',
        'Updates a position in PositionHierarchies.',
        updatePositionHierarchySchema.shape,
        async ({ positionId, hierarchyTypeName, validFrom, validTo, updateData }: z.infer<typeof updatePositionHierarchySchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/PositionHierarchies(PositionId='${positionId}',HierarchyTypeName='${hierarchyTypeName}',ValidFrom=${validFrom},ValidTo=${validTo})`;
            return makeApiCall('PATCH', url, updateData as Record<string, unknown>, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'action_initializeDataManagement',
        'Executes the InitializeDataManagement action on the DataManagementDefinitionGroups entity.',
        z.object({}).shape,
        async (_args: {}, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities.InitializeDataManagement`;
            return makeApiCall('POST', url, {}, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'getODataMetadata',
        'Retrieves the OData $metadata document for the service.',
        z.object({}).shape,
        async (_args: {}, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
             const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/$metadata`;
             return makeApiCall('GET', url.toString(), null, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );
//Added by JP Start 
server.tool(
  'createSalesOrder',
  'Creates a sales order header (SalesOrderHeadersV4) and one or more lines (SalesOrderLinesV3). SalesOrderNumber and SiteId are optional.',
  createSalesOrderSchema.shape,
  async ({ header, lines }: z.infer<typeof createSalesOrderSchema>, context) => {
    // Helper to relay notifications safely and unwrap text from makeApiCall
    const callAndGetText = async (
      method: 'GET' | 'POST' | 'PATCH',
      url: string,
      body: Record<string, unknown> | null
    ) => {
      const res = await makeApiCall(method, url, body, async (n) => { await safeNotification(context, n); });
      const first = res.content?.[0];
      const text = (first && first.type === 'text') ? first.text : JSON.stringify(res);
      return { res, text };
    };

    // 1) Build header payload (omit SalesOrderNumber if not given to let D365 auto-number)
    const headerPayload: Record<string, unknown> = {
      dataAreaId: header.dataAreaId,
      RequestedShippingDate: header.RequestedShippingDate,
      CustomerAccount: header.CustomerAccount
    };
    if (header.SalesOrderNumber) {
      headerPayload.SalesOrderNumber = header.SalesOrderNumber;
    }

    await safeNotification(context, {
      method: "notifications/message",
      params: { level: "info", data: `Creating sales order header${header.SalesOrderNumber ? `: ${header.SalesOrderNumber}` : ''}` }
    });

    // 2) Create Header
    const headerUrl = `${process.env.DYNAMICS_RESOURCE_URL}/data/SalesOrderHeadersV4`;
    const { res: headerRes, text: headerText } = await callAndGetText('POST', headerUrl, headerPayload);
    if ((headerRes as any).isError) {
      return {
        isError: true,
        content: [{ type: 'text', text: `Failed to create Sales Order header.\n\n${headerText}` }]
      };
    }

    await safeNotification(context, {
      method: "notifications/message",
      params: { level: "info", data: `Header created. Creating ${lines.length} line(s)...` }
    });

    // 3) Create Lines
    const lineUrl = `${process.env.DYNAMICS_RESOURCE_URL}/data/SalesOrderLinesV3`;
    const perLineResults: Array<{ index: number; ok: boolean; message: string }> = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];

      const linePayload: Record<string, unknown> = {
        dataAreaId: header.dataAreaId,
        // Use the same order number reference we sent (or whatever D365 assigned).
        // If your environment returns the assigned number in the header response body,
        // you can parse it here and override. For now, we rely on providing SalesOrderNumber
        // when you have one, or D365 linking by context when Site/number defaulting applies.
        SalesOrderNumber: header.SalesOrderNumber, // omitted if undefined
        ItemNumber: line.ItemNumber,
        OrderedSalesQuantity: line.OrderedSalesQuantity
      };
      if (line.SiteId) {
        linePayload.SiteId = line.SiteId;
      }
      // If SalesOrderNumber was not provided and your system requires it on lines,
      // consider fetching the header we just created to read the assigned number.

      const { res, text } = await callAndGetText('POST', lineUrl, linePayload);
      const ok = !(res as any).isError;
      perLineResults.push({
        index: i,
        ok,
        message: ok ? `Line ${i + 1} created.` : `Line ${i + 1} failed.\n${text}`
      });

      await safeNotification(context, {
        method: "notifications/message",
        params: { level: ok ? "info" : "error", data: perLineResults[perLineResults.length - 1].message }
      });
    }

    // 4) Summarize
    const summary = {
      header: {
        SalesOrderNumber: header.SalesOrderNumber ?? '(auto-assigned)',
        created: true
      },
      lines: perLineResults
    };

    return {
      content: [{
        type: 'text',
        text: `Sales order creation summary:\n\n${JSON.stringify(summary, null, 2)}\n\nTip: If any line failed, re-run with only failed items.`
      }]
    };
  }
);
//Added by JP End

    return server;

};
