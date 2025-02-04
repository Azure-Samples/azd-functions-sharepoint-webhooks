export const CommonConfig = {
    TenantPrefix: process.env.TenantPrefix || "",
    SharePointDomain: process.env.SharePointDomain || "sharepoint.com",
    SiteRelativePath: process.env.SiteRelativePath || "",
    IsLocalEnvironment: process.env.AZURE_FUNCTIONS_ENVIRONMENT === "Development" ? true : false,
    UserAssignedManagedIdentityClientId: process.env.UserAssignedManagedIdentityClientId || undefined,
    WebhookHistoryListTitle: process.env.WebhookHistoryListTitle || "webhookHistory",
    UserAgent: process.env.UserAgent || "Yvand/azd-functions-sharepoint-webhooks",
    MinutesFromNowForChanges: Number(process.env.MinutesFromNowForChanges) || -15,
}

// This method awaits on async calls and catches the exception if there is any - https://dev.to/sobiodarlington/better-error-handling-with-async-await-2e5m
export const safeWait = (promise: Promise<any>) => {
    return promise
        .then(data => ([data, undefined]))
        .catch(error => Promise.resolve([undefined, error]));
}

/**
 * Returns the ticks representing the time some minutes from now, compatible with SharePoint change tokens
 * @param minutesFromNow 
 * @returns Ticks value to use in a SharePoint change token
 */
export const GetChangeTokenTicks = (minutesFromNow: number) => {
    const now = new Date();
    // Convert JavaScript date to .NET ticks: https://stackoverflow.com/questions/7966559/how-to-convert-javascript-date-object-to-ticks
    return ((now.getTime() * 10_000 + minutesFromNow * 60_000 * 10_000) + 621355968000000000);
}

export enum WebhookChangeType {
    Add = 1,
    Update = 2,
    Delete = 3,
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