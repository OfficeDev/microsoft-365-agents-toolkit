// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export enum LogLevel {
  Debug = 1,
  Verbose = 2,
  Info = 3,
  Warning = 4,
  Error = 5,
}

export interface LogProvider {
  log(logLevel: LogLevel, message: string): void;
  verbose(message: string): void;
  debug(message: string): void;
  info(message: string): void;
  info(message: Array<{ content: string; color: Colors }>): void;
  warning(message: string): void;
  error(message: string): void;
  logInFile(logLevel: LogLevel, message: string): Promise<void>;
  getLogFilePath(): string;
}

export enum Colors {
  BRIGHT_WHITE = 0,
  WHITE = 1,
  BRIGHT_MAGENTA = 2,
  BRIGHT_GREEN = 3,
  BRIGHT_YELLOW = 4,
  BRIGHT_RED = 5,
  BRIGHT_CYAN = 6,
}
