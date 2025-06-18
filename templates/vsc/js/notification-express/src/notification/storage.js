// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const fs = require("fs");
const path = require("path");

class LocalConversationReferenceStore {
  constructor(fileDir) {
    this.localFileName =
      process.env.TEAMSFX_NOTIFICATION_STORE_FILENAME ?? ".notification.localstore.json";
    this.filePath = path.resolve(fileDir, this.localFileName);
  }

  async write(changes) {
    if (!(await this.storeFileExists())) {
      await this.writeToFile(changes);
    } else {
      const data = await this.readFromFile();
      await this.writeToFile(Object.assign(data, changes));
    }
  }

  async read(keys) {
    if (!(await this.storeFileExists())) {
      return {};
    }

    const data = await this.readFromFile();
    const result = {};
    for (const key of keys) {
      if (data[key] !== undefined) {
        result[key] = data[key];
      }
    }
    return result;
  }

  async delete(keys) {
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

  async list(pageSize, continuationToken) {
    if (!(await this.storeFileExists())) {
      return {
        data: [],
        continuationToken: "",
      };
    }

    const fileData = await this.readFromFile();
    const data = Object.entries(fileData).map((entry) => entry[1]);
    return {
      data,
      continuationToken: "",
    };
  }

  storeFileExists() {
    return new Promise((resolve) => {
      try {
        fs.access(this.filePath, (err) => {
          if (err) {
            resolve(false);
          } else {
            resolve(true);
          }
        });
      } catch (error) {
        resolve(false);
      }
    });
  }

  readFromFile() {
    return new Promise((resolve, reject) => {
      try {
        fs.readFile(this.filePath, { encoding: "utf-8" }, (err, rawData) => {
          if (err) {
            reject(err);
          } else {
            resolve(JSON.parse(rawData));
          }
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  async writeToFile(data) {
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
      } catch (error) {
        reject(error);
      }
    });
  }
}

module.exports = {
  LocalConversationReferenceStore,
};
