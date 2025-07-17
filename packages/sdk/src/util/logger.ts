// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Interface for customized logger.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export interface Logger {
  /**
   * Writes to error level logging or lower.
   */
  error(message: string): void;
  /**
   * Writes to warning level logging or lower.
   */
  warn(message: string): void;
  /**
   * Writes to info level logging or lower.
   */
  info(message: string): void;
  /**
   * Writes to verbose level logging.
   */
  verbose(message: string): void;
}

/**
 * Log function for customized logging.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export type LogFunction = (level: LogLevel, message: string) => void;

/**
 * Log level.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export enum LogLevel {
  /**
   * Show verbose, information, warning and error message.
   */
  Verbose,
  /**
   * Show information, warning and error message.
   */
  Info,
  /**
   * Show warning and error message.
   */
  Warn,
  /**
   * Show error message.
   */
  Error,
}

/**
 * Update log level helper.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @param { LogLevel } level - log level in configuration
 */
export function setLogLevel(level: LogLevel): void {
  internalLogger.level = level;
}

/**
 * Get log level.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @returns Log level
 */
export function getLogLevel(): LogLevel | undefined {
  return internalLogger.level;
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export class InternalLogger implements Logger {
  public name?: string;
  public level?: LogLevel = undefined;
  public customLogger: Logger | undefined;
  public customLogFunction: LogFunction | undefined;

  private defaultLogger: Logger = {
    verbose: console.debug,
    info: console.info,
    warn: console.warn,
    error: console.error,
  };

  constructor(name?: string, logLevel?: LogLevel) {
    this.name = name;
    this.level = logLevel;
  }

  public error(message: string): void {
    this.log(LogLevel.Error, (x: Logger) => x.error, message);
  }

  public warn(message: string): void {
    this.log(LogLevel.Warn, (x: Logger) => x.warn, message);
  }

  public info(message: string): void {
    this.log(LogLevel.Info, (x: Logger) => x.info, message);
  }

  public verbose(message: string): void {
    this.log(LogLevel.Verbose, (x: Logger) => x.verbose, message);
  }

  private log(
    logLevel: LogLevel,
    logFunction: (x: Logger) => (message: string) => void,
    message: string
  ): void {
    if (message.trim() === "") {
      return;
    }
    const timestamp = new Date().toUTCString();
    let logHeader: string;
    if (this.name) {
      logHeader = `[${timestamp}] : @microsoft/teamsfx - ${this.name} : ${LogLevel[logLevel]} - `;
    } else {
      logHeader = `[${timestamp}] : @microsoft/teamsfx : ${LogLevel[logLevel]} - `;
    }
    const logMessage = `${logHeader}${message}`;
    if (this.level !== undefined && this.level <= logLevel) {
      if (this.customLogger) {
        logFunction(this.customLogger)(logMessage);
      } else if (this.customLogFunction) {
        this.customLogFunction(logLevel, logMessage);
      } else {
        logFunction(this.defaultLogger)(logMessage);
      }
    }
  }
}

/**
 * Logger instance used internally
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export const internalLogger: InternalLogger = new InternalLogger();

/**
 * Set custom logger. Use the output functions if it's set. Priority is higher than setLogFunction.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @param {Logger} logger - custom logger. If it's undefined, custom logger will be cleared.
 *
 * @example
 * ```typescript
 * setLogger({
 *   verbose: console.debug,
 *   info: console.info,
 *   warn: console.warn,
 *   error: console.error,
 * });
 * ```
 */
export function setLogger(logger?: Logger): void {
  internalLogger.customLogger = logger;
}

/**
 * Set custom log function. Use the function if it's set. Priority is lower than setLogger.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @param {LogFunction} logFunction - custom log function. If it's undefined, custom log function will be cleared.
 *
 * @example
 * ```typescript
 * setLogFunction((level: LogLevel, message: string) => {
 *   if (level === LogLevel.Error) {
 *     console.log(message);
 *   }
 * });
 * ```
 */
export function setLogFunction(logFunction?: LogFunction): void {
  internalLogger.customLogFunction = logFunction;
}
