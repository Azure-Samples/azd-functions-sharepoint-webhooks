import { InvocationContext } from "@azure/functions";
import { Logger, FunctionListener, ILogEntry, ILogListener, LogLevel } from "@pnp/logging";
import { CommonConfig, safeWait } from "./common.js";
import { HttpRequestError } from "@pnp/queryable";
import { hOP } from "@pnp/core";

// // Set the logging context passed to the functions
// let logcontext: InvocationContext;
// export function setLoggingContext(context: InvocationContext): void {
//     logcontext = context; 
// }

// Create a listener to write messages to the logging system
const listener: ILogListener = FunctionListener((entry: ILogEntry): void => {
    logMessage(entry);
});
Logger.subscribe(listener);
Logger.activeLogLevel = LogLevel.Verbose;

// Internal function which logs all the messages including the formatted errors, to app insights if possible, and to the console if in local environment
function logMessage(entry: ILogEntry): void {
    let logcontext: InvocationContext = entry.data;
    if (logcontext) {
        switch (entry.level) {
            case LogLevel.Info:
                logcontext.log(entry.message);
                break;
            case LogLevel.Warning:
                logcontext.warn(entry.message);
                break;
            case LogLevel.Error:
                logcontext.error(entry.message);
                break;
            case LogLevel.Verbose:
                logcontext.trace(entry.message);
                break;
            default:
                logcontext.log(entry.message);
                break;
        }
    } else if (CommonConfig.IsLocalEnvironment) {
        console.log(entry.message);
    }
}

export interface IErrorDetails {
    timestamp: string;
    status: string;
    errorMessage: string;
    contextMessage?: string;
    type: string;
    sprequestguid?: string;
    httpStatus?: number;
}

/**
 * Handles the error and logs it
 * @param e 
 * @param contextMessage 
 * @returns formatted error message
 */
export async function handleError(e: Error | HttpRequestError | unknown, logcontext: InvocationContext, contextMessage?: string): Promise<IErrorDetails> {
    let errorDetails: IErrorDetails = { timestamp: new Date().toISOString(), status: "error", errorMessage: "", type: "", contextMessage: contextMessage };
    let level: LogLevel = LogLevel.Error;
    let message = "";

    if (e instanceof Error) {
        if (hOP(e, "isHttpRequestError")) {
            errorDetails.type = "HttpRequestError";
            let [jsonResponse, awaiterror] = await safeWait((<HttpRequestError>e).response.json());
            if (jsonResponse) {
                message += typeof jsonResponse["odata.error"] === "object" ? jsonResponse["odata.error"].message.value : e.message;
            } else {
                message += e.message;
            }

            errorDetails.httpStatus = (<HttpRequestError>e).status;
            if (errorDetails.httpStatus === 404) {
                level = LogLevel.Warning;
            }

            const spCorrelationid = (e as HttpRequestError).response.headers.get("sprequestguid");
            errorDetails.sprequestguid = spCorrelationid || "";
        } else {
            errorDetails.type = e.name;
            message += e.message;
        }
    } else if (typeof e === "string") {
        errorDetails.type = "string";
        message += e;
    }
    else {
        errorDetails.type = "unknown";
        message += errorDetails.errorMessage;
    }
    
    errorDetails.errorMessage = message;
    Logger.log({
        data: logcontext,
        level: level,
        message: JSON.stringify(errorDetails),
    });
    return errorDetails;
}

export function logInfoMessage(logcontext: InvocationContext, message: string, level: LogLevel = LogLevel.Info): void {
    Logger.log({
        data: logcontext,
        level: level,
        message: message,
    });
}
