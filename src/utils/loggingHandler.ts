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
    writeEntryToLog(entry);
});
Logger.subscribe(listener);
Logger.activeLogLevel = LogLevel.Verbose;

/**
 * Internal function which writes the entry to the log: application insights if possible, or the console if in local environment
 * @param entry 
 */
function writeEntryToLog(entry: ILogEntry): void {
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
 * @param error 
 * @param logcontext 
 * @param message 
 * @returns document with details about the error
 */
export async function logError(logcontext: InvocationContext, error: Error | HttpRequestError | unknown, message: string): Promise<IErrorMessageDocument> {
    let errorDocument: IErrorMessageDocument = { timestamp: new Date().toISOString(), level: LogLevel.Error, message: message, error: "", type: "", httpStatus: 500 };
    let errorDetails = "";
    if (error instanceof Error) {
        if (hOP(error, "isHttpRequestError")) {
            errorDocument.type = "HttpRequestError";
            let [jsonResponse, awaiterror] = await safeWait((<HttpRequestError>error).response.json());
            if (jsonResponse) {
                errorDetails += typeof jsonResponse["odata.error"] === "object" ? jsonResponse["odata.error"].message.value : error.message;
            } else {
                errorDetails += error.message;
            }

            errorDocument.httpStatus = (<HttpRequestError>error).status;
            if (errorDocument.httpStatus === 404) {
                errorDocument.level = LogLevel.Warning;
            }

            const spCorrelationId = (error as HttpRequestError).response.headers.get("sprequestguid");
            errorDocument.sprequestguid = spCorrelationId || "";
        } else if (error instanceof AggregateError) {
            errorDocument.type = error.name;
            errorDetails += `AggregateError with ${error.errors.length} errors: `;
            for (let i = 0; i < error.errors.length; i++) {
                errorDetails += `Error ${i}: ${error.errors[i].name}: ${error.errors[i].message}. `;
            }
        } else {
            errorDocument.type = error.name;
            errorDetails += error.message;
        }
    } else if (typeof error === "string") {
        errorDocument.type = "string";
        errorDetails = error;
    }
    else {
        errorDocument.type = "unknown";
        errorDetails = JSON.stringify(error);
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
export function logMessage(logcontext: InvocationContext, message: string, level: LogLevel = LogLevel.Info): IMessageDocument {
    const messageResponse: IMessageDocument = { timestamp: new Date().toISOString(), level: level, message: message };
    Logger.log({
        data: logcontext,
        level: messageResponse.level,
        message: JSON.stringify(messageResponse),
    });
    return messageResponse;
}
