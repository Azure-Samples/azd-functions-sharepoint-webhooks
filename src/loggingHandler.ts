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

/**
 * Handles the error and logs it
 * @param e 
 * @param currentOperationDetails 
 * @returns formatted error message
 */
export async function handleError(e: Error | HttpRequestError, logcontext: InvocationContext, currentOperationDetails?: string): Promise<string> {
    let message = currentOperationDetails ? `${currentOperationDetails}: ` : "";
    let level: LogLevel = LogLevel.Error;

    if (hOP(e, "isHttpRequestError")) {
        let [jsonResponse, awaiterror] = await safeWait((<HttpRequestError>e).response.json());
        if (jsonResponse) {
            message += typeof jsonResponse["odata.error"] === "object" ? jsonResponse["odata.error"].message.value : e.message;
        } else {
            message += e.message;
        }
        if ((<HttpRequestError>e).status === 404) {
            level = LogLevel.Warning;
        }
    } else {
        message += e.message;
    }

    Logger.log({
        data: logcontext,
        level: level,
        message: message,
    });
    return message;
}