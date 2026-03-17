import { IConnectionSchema } from "@cognigy/extension-tools";

export const graphConnection: IConnectionSchema = {
    type: "graphCloud",
    label: "Microsoft Graph (App Registration)",
    fields: [
        { fieldName: "tenantId" },
        { fieldName: "clientId" },
        { fieldName: "clientSecret" }
    ]
};
