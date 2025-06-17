// Copyright (c) Microsoft Corporation.Add commentMore actions
// Licensed under the MIT license.

import { ConversationReference } from "@microsoft/agents-activity";
import { IStorage, PagedData } from "./interface";
import * as fs from "fs";
import * as path from "path";

export class LocalConversationReferenceStore implements IStorage {
  private readonly localFileName =
    process.env.TEAMSFX_NOTIFICATION_STORE_FILENAME ?? ".notification.localstore.json";
  private readonly filePath: string;

  constructor(fileDir: string) {
    this.filePath = path.resolve(fileDir, this.localFileName);
  }

  public async write(changes: { [key: string]: Partial<ConversationReference> }): Promise<void> {
    if (!(await this.storeFileExists())) {
      await this.writeToFile(changes);
    } else {
      const data = await this.readFromFile();
      await this.writeToFile(Object.assign(data, changes));
    }
  }

  public async read(keys: string[]): Promise<{ [key: string]: Partial<ConversationReference> }> {
    if (!(await this.storeFileExists())) {
      return {};
    }

    const data = await this.readFromFile();
    const result: { [key: string]: Partial<ConversationReference> } = {};
    for (const key of keys) {
      if (data[key] !== undefined) {
        result[key] = data[key];
      }
    }
    return result;
  }

  public async delete(keys: string[]): Promise<void> {
    if (!(await this.storeFileExists())) {
      return;
    }

    const data = await this.readFromFile();
    for (const key of keys) {
      if (data[key] !== undefined) {
        delete data[key];
      }
    }
    await this.writeToFile(data);
  }

  public async list(
    pageSize?: number,
    continuationToken?: string
  ): Promise<PagedData<Partial<ConversationReference>>> {
    if (!(await this.storeFileExists())) {
      return {
        data: [],
        continuationToken: "",
      };
    }

    const fileData = await this.readFromFile();
    const data: { [key: string]: unknown }[] = Object.entries(fileData).map(
      (entry) => entry[1] as { [key: string]: unknown }
    );
    return {
      data,
      continuationToken: "",
    };
  }

  private storeFileExists(): Promise<boolean> {
    return new Promise((resolve) => {
      try {
        fs.access(this.filePath, (err) => {
          if (err) {
            resolve(false);
          } else {
            resolve(true);
          }
        });
      } catch (error: unknown) {
        resolve(false);
      }
    });
  }

  private readFromFile(): Promise<any> {
    return new Promise((resolve, reject) => {
      try {
        fs.readFile(this.filePath, { encoding: "utf-8" }, (err, rawData) => {
          if (err) {
            reject(err);
          } else {
            resolve(JSON.parse(rawData));
          }
        });
      } catch (error: unknown) {
        reject(error);
      }
    });
  }

  private async writeToFile(data: unknown): Promise<void> {
    return new Promise((resolve, reject) => {
      try {
        const rawData = JSON.stringify(data, undefined, 2);
        fs.writeFile(this.filePath, rawData, { encoding: "utf-8" }, (err) => {
          if (err) {
            reject(err);
          } else {
            resolve();
          }
        });
      } catch (error: unknown) {
        reject(error);
      }
    });
  }
}
