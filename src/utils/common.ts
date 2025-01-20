export const CommonConfig = {
    TenantPrefix: process.env.TenantPrefix || "",
    SharePointDomain: process.env.SharePointDomain || "sharepoint.com",
    SiteRelativePath: process.env.SiteRelativePath || "",
    IsLocalEnvironment: process.env.AZURE_FUNCTIONS_ENVIRONMENT === "Development" ? true : false,
    UserAssignedManagedIdentityClientId: process.env.UserAssignedManagedIdentityClientId || undefined,
    WebhookHistoryListTitle: process.env.WebhookHistoryListTitle || "webhookHistory",
    UserAgent: process.env.UserAgent || "Yvand/azd-functions-sharepoint-webhooks",
}

// This method awaits on async calls and catches the exception if there is any - https://dev.to/sobiodarlington/better-error-handling-with-async-await-2e5m
export const safeWait = (promise: Promise<any>) => {
    return promise
        .then(data => ([data, undefined]))
        .catch(error => Promise.resolve([undefined, error]));
}

export interface ISubscriptionResponse {
    clientState: string;
    expirationDateTime: string;
    Id: string;
    notificationUrl: string;
    resource: string;
    resourceData: string;
    scenarios: string;
}

export interface ISharePointWeebhookEvent {
    value: ISharePointWeebhookEventValue[];
}

export interface ISharePointWeebhookEventValue {
    subscriptionId: string;
    clientState: string;
    expirationDateTime: string;
    resource: string;
    tenantId: string;
    siteUrl: string;
    webId: string;
}