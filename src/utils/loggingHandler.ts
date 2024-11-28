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

export interface IMessageDocument {
    timestamp: string;
    level: LogLevel;
    message: string;
}

export interface IErrorMessageDocument extends IMessageDocument {
    error: string;
    type: string;
    sprequestguid?: string;
    httpStatus?: number;
}

/**
 * Process the error, record an error message and return a document with details about the error
 * @param e 
 * @param logcontext 
 * @param message 
 * @returns document with details about the error
 */
export async function logError(e: Error | HttpRequestError | unknown, logcontext: InvocationContext, message: string): Promise<IErrorMessageDocument> {
    let errorDocument: IErrorMessageDocument = { timestamp: new Date().toISOString(), level: LogLevel.Error, message: message, error: "", type: "" };
    let errorDetails = "";
    if (e instanceof Error) {
        if (hOP(e, "isHttpRequestError")) {
            errorDocument.type = "HttpRequestError";
            let [jsonResponse, awaiterror] = await safeWait((<HttpRequestError>e).response.json());
            if (jsonResponse) {
                errorDetails += typeof jsonResponse["odata.error"] === "object" ? jsonResponse["odata.error"].message.value : e.message;
            } else {
                errorDetails += e.message;
            }

            errorDocument.httpStatus = (<HttpRequestError>e).status;
            if (errorDocument.httpStatus === 404) {
                errorDocument.level = LogLevel.Warning;
            }

            const spCorrelationId = (e as HttpRequestError).response.headers.get("sprequestguid");
            errorDocument.sprequestguid = spCorrelationId || "";
        } else {
            errorDocument.type = e.name;
            errorDetails += e.message;
        }
    } else if (typeof e === "string") {
        errorDocument.type = "string";
        errorDetails += e;
    }
    else {
        errorDocument.type = "unknown";
        errorDetails += errorDocument.error;
    }
    
    errorDocument.error = errorDetails;
    Logger.log({
        data: logcontext,
        level: errorDocument.level,
        message: JSON.stringify(errorDocument),
    });
    return errorDocument;
}

/**
 * record the message and return a document with additionnal details
 * @param logcontext 
 * @param message 
 * @param level 
 * @returns 
 */
export function logInfo(logcontext: InvocationContext, message: string, level: LogLevel = LogLevel.Info): IMessageDocument {
    const messageResponse: IMessageDocument = { timestamp: new Date().toISOString(), level: level, message: message };
    Logger.log({
        data: logcontext,
        level: messageResponse.level,
        message: JSON.stringify(messageResponse),
    });
    return messageResponse;
}
