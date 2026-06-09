// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import fs from "fs-extra";
import * as sinon from "sinon";
import * as util from "util";

import * as commentJson from "comment-json";
import * as localizeUtils from "../../../../src/common/localizeUtils";
import { CreateOrUpdateJsonFileDriver } from "../../../../src/component/driver/file/createOrUpdateJsonFile";
import { InvalidActionInputError } from "../../../../src/error/common";
import { MockedLogProvider } from "../../../plugins/solution/util";

describe("CreateOrUpdateJsonFileDriver", () => {
  const mockedDriverContext = {
    logProvider: new MockedLogProvider(),
  } as any;
  const driver = new CreateOrUpdateJsonFileDriver();

  beforeEach(() => {
    sinon.stub(localizeUtils, "getDefaultString").callsFake((key, ...params) => {
      if (key === "error.yaml.InvalidActionInputError") {
        return util.format("error.yaml.InvalidActionInputError. %s. %s.", ...params);
      } else if (key === "error.common.UnhandledError") {
        return util.format("error.common.UnhandledError. %s. %s", ...params);
      }
      return "";
    });
    sinon.stub(localizeUtils, "getLocalizedString").returns("");
  });

  afterEach(() => {
    sinon.restore();
  });

  describe("run", () => {
    it("invalid args: empty target", async () => {
      const args: any = {
        target: null,
        appsettings: {
          BOT_ID: "BOT_ID",
          BOT_PASSWORD: "BOT_PASSWORD",
        },
      };
      const result = await driver.run(args, mockedDriverContext);
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert(result.error instanceof InvalidActionInputError);
      }
    });

    it("invalid args: appsettings is not object", async () => {
      const args: any = {
        target: "target",
        appsettings: "value",
      };
      const result = await driver.run(args, mockedDriverContext);
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert(result.error instanceof InvalidActionInputError);
      }
    });

    it("exception", async () => {
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readFile").rejects(new Error("exception"));
      const args: any = {
        target: "path",
        appsettings: {
          BOT_ID: "BOT_ID",
          BOT_PASSWORD: "BOT_PASSWORD",
        },
      };
      const result = await driver.run(args, mockedDriverContext);
      chai.assert(result.isErr());
    });

    it("happy path: with target", async () => {
      const target = "path";
      let content = {};
      const appsettings = {
        BOT_ID: "$botId$",
        BOT_PASSWORD: "$bot-password$",
      };
      sinon.stub(fs, "ensureFile").callsFake(async (path) => {
        return;
      });
      sinon.stub(fs, "readFile").callsFake(async (path) => {
        return Buffer.from(JSON.stringify(appsettings));
      });
      sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
        content = data;
        return;
      });
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "existsSync").callsFake((path) => {
        return true;
      });
      const args: any = {
        target,
        appsettings: {
          BOT_ID: "BOT_ID",
          BOT_PASSWORD: "BOT_PASSWORD",
        },
      };
      const result = await driver.run(args, mockedDriverContext);
      chai.assert(result.isOk());
      if (result.isOk()) {
        chai.assert.equal('{\n\t"BOT_ID": "BOT_ID",\n\t"BOT_PASSWORD": "BOT_PASSWORD"\n}', content);
      }
    });

    it("happy path: execute with target", async () => {
      const target = "path";
      let content = {};
      const appsettings = {
        BOT_ID: "$botId$",
        BOT_PASSWORD: "$bot-password$",
      };
      sinon.stub(fs, "ensureFile").callsFake(async (path) => {
        return;
      });
      sinon.stub(fs, "readFile").callsFake(async (path) => {
        return Buffer.from(JSON.stringify(appsettings));
      });
      sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
        content = data;
        return;
      });
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "existsSync").callsFake((path) => {
        return true;
      });
      const args: any = {
        target,
        appsettings: {
          BOT_ID: "BOT_ID",
          BOT_PASSWORD: "BOT_PASSWORD",
        },
      };
      const result = await driver.execute(args, mockedDriverContext);
      chai.assert(result.result.isOk());
      if (result.result.isOk()) {
        chai.assert.equal('{\n\t"BOT_ID": "BOT_ID",\n\t"BOT_PASSWORD": "BOT_PASSWORD"\n}', content);
      }
    });

    it("happy path: with target and customized data", async () => {
      const target = "path";
      let content = {};
      const appsettings = {
        Foo: "Bar",
        My: {
          BOT_ID: "$botId$",
          Foo: "Bar",
        },
      };
      sinon.stub(fs, "ensureFile").callsFake(async (path) => {
        return;
      });
      sinon.stub(fs, "readFile").callsFake(async (path) => {
        return Buffer.from(JSON.stringify(appsettings));
      });
      sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
        content = data;
        return;
      });
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "existsSync").callsFake((path) => {
        return true;
      });
      const args: any = {
        target,
        appsettings: {
          My: {
            BOT_ID: "BOD_ID",
          },
        },
      };
      const result = await driver.run(args, mockedDriverContext);
      chai.assert(result.isOk());
      if (result.isOk()) {
        chai.assert.equal(
          '{\n\t"Foo": "Bar",\n\t"My": {\n\t\t"BOT_ID": "BOD_ID",\n\t\t"Foo": "Bar"\n\t}\n}',
          content
        );
      }
    });

    it("happy path: with appsettings.json", async () => {
      const target = "path";
      let content = {};
      const appsettings = {
        BOT_ID: "$botId$",
        BOT_PASSWORD: "$bot-password$",
      };
      sinon.stub(fs, "ensureFile").callsFake(async (path) => {
        return;
      });
      sinon.stub(fs, "readFile").callsFake(async (path) => {
        return Buffer.from(JSON.stringify(appsettings));
      });
      sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
        content = data;
        return;
      });
      sinon.stub(fs, "pathExists").callsFake(async (path: fs.PathLike) => {
        if (path.toString().indexOf(target) >= 0) {
          return false;
        }
        return true;
      });
      sinon.stub(fs, "existsSync").callsFake((path) => {
        if (path.toString().indexOf(target) >= 0) {
          return false;
        }
        return true;
      });
      sinon.stub(fs, "copyFile").callsFake(async (p1, p2) => {
        return;
      });
      const args: any = {
        target,
        appsettings: {
          BOT_ID: "BOT_ID",
          BOT_PASSWORD: "BOT_PASSWORD",
        },
      };
      const result = await driver.run(args, mockedDriverContext);
      chai.assert(result.isOk());
      if (result.isOk()) {
        chai.assert.equal('{\n\t"BOT_ID": "BOT_ID",\n\t"BOT_PASSWORD": "BOT_PASSWORD"\n}', content);
      }
    });
  });

  it("happy path: without appsettings.json", async () => {
    const target = "path";
    let content = {};
    const appsettings = {
      BOT_ID: "$botId$",
      BOT_PASSWORD: "$bot-password$",
    };
    sinon.stub(fs, "ensureFile").callsFake(async (path) => {
      return;
    });
    sinon.stub(fs, "readFile").callsFake(async (path) => {
      return Buffer.from(JSON.stringify(appsettings));
    });
    sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
      content = data;
      return;
    });
    sinon.stub(fs, "pathExists").resolves(false);
    sinon.stub(fs, "existsSync").callsFake((path) => {
      return false;
    });
    sinon.stub(fs, "copyFile").callsFake(async (p1, p2) => {
      return;
    });
    const args: any = {
      target,
      appsettings: {
        BOT_ID: "BOT_ID",
        BOT_PASSWORD: "BOT_PASSWORD",
      },
    };
    const result = await driver.run(args, mockedDriverContext);
    chai.assert(result.isOk());
    if (result.isOk()) {
      chai.assert.equal('{\n\t"BOT_ID": "BOT_ID",\n\t"BOT_PASSWORD": "BOT_PASSWORD"\n}', content);
    }
  });

  it("happy path: using content with comment json", async () => {
    const target = "path";
    let content = {};
    const jsonContent = commentJson.parse(`{
      // comment string 1
      "BOT_ID": "$botId$",
      "BOT_PASSWORD": "$bot-password$",
      "FOO": "BAR"
      // comment string 2
    }`);
    sinon.stub(fs, "ensureFile").callsFake(async (path) => {
      return;
    });
    sinon.stub(fs, "readFile").callsFake(async (path) => {
      return Buffer.from(commentJson.stringify(jsonContent, null, "\t"));
    });
    sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
      content = data;
      return;
    });
    sinon.stub(fs, "pathExists").resolves(false);
    sinon.stub(fs, "existsSync").callsFake((path) => {
      return false;
    });
    sinon.stub(fs, "copyFile").callsFake(async (p1, p2) => {
      return;
    });
    const args: any = {
      target,
      content: {
        BOT_ID: "BOT_ID",
        BOT_PASSWORD: "BOT_PASSWORD",
        FOO2: "BAR2",
      },
    };
    const result = await driver.run(args, mockedDriverContext);
    chai.assert(result.isOk());
    if (result.isOk()) {
      chai.assert.equal(
        '{\n\t// comment string 1\n\t"BOT_ID": "BOT_ID",\n\t"BOT_PASSWORD": "BOT_PASSWORD",\n\t"FOO": "BAR",\n\t// comment string 2\n\t"FOO2": "BAR2"\n}',
        content
      );
    }
  });

  it("happy path: using content with comment json, boolean and double values", async () => {
    const target = "path";
    let content = {};
    const jsonContent = commentJson.parse(`{
      // comment string 1
      "BOT_ID": "$botId$",
      "BOT_PASSWORD": "$bot-password$",
      "FOO": "BAR",
      "FOO2": true,
      // comment string 2
    }`);
    sinon.stub(fs, "ensureFile").callsFake(async (path) => {
      return;
    });
    sinon.stub(fs, "readFile").callsFake(async (path) => {
      return Buffer.from(commentJson.stringify(jsonContent, null, "\t"));
    });
    sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
      content = data;
      return;
    });
    sinon.stub(fs, "pathExists").resolves(false);
    sinon.stub(fs, "existsSync").callsFake((path) => {
      return false;
    });
    sinon.stub(fs, "copyFile").callsFake(async (p1, p2) => {
      return;
    });
    const args: any = {
      target,
      content: {
        BOT_ID: "BOT_ID",
        BOT_PASSWORD: "BOT_PASSWORD",
        FOO2: false,
        FOO3: 1.2,
      },
    };
    const result = await driver.run(args, mockedDriverContext);
    chai.assert(result.isOk());
    if (result.isOk()) {
      chai.assert.equal(
        '{\n\t// comment string 1\n\t"BOT_ID": "BOT_ID",\n\t"BOT_PASSWORD": "BOT_PASSWORD",\n\t"FOO": "BAR",\n\t"FOO2": false,\n\t// comment string 2\n\t"FOO3": 1.2\n}',
        content
      );
    }
  });

  it("invalid path: using content and appsettings at the same time", async () => {
    const target = "path";
    let content = {};
    const jsonContent = {
      BOT_ID: "$botId$",
      BOT_PASSWORD: "$bot-password$",
    };
    sinon.stub(fs, "ensureFile").callsFake(async (path) => {
      return;
    });
    sinon.stub(fs, "readFile").callsFake(async (path) => {
      return Buffer.from(JSON.stringify(jsonContent));
    });
    sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
      content = data;
      return;
    });
    sinon.stub(fs, "pathExists").callsFake(async (path: fs.PathLike) => {
      if (path.toString().indexOf(target) >= 0) {
        return false;
      }
      return true;
    });
    sinon.stub(fs, "existsSync").callsFake((path) => {
      if (path.toString().indexOf(target) >= 0) {
        return false;
      }
      return true;
    });
    sinon.stub(fs, "copyFile").callsFake(async (p1, p2) => {
      return;
    });
    const args: any = {
      target,
      appsettings: {
        BOT_ID: "BOT_ID",
        BOT_PASSWORD: "BOT_PASSWORD",
      },
      content: {
        BOT_ID: "BOT_ID",
        BOT_PASSWORD: "BOT_PASSWORD",
      },
    };
    const result = await driver.run(args, mockedDriverContext);
    chai.assert(result.isErr());
    chai.expect((result as any).error.name).equals("InvalidActionInputError");
  });

  it("happy path: add nested object", async () => {
    const target = "path";
    let content = {};
    const jsonContent = {
      FOO: {},
    };
    sinon.stub(fs, "ensureFile").callsFake(async (path) => {
      return;
    });
    sinon.stub(fs, "readFile").callsFake(async (path) => {
      return Buffer.from(JSON.stringify(jsonContent));
    });
    sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
      content = data;
      return;
    });
    sinon.stub(fs, "pathExists").resolves(false);
    sinon.stub(fs, "existsSync").callsFake((path) => {
      return false;
    });
    sinon.stub(fs, "copyFile").callsFake(async (p1, p2) => {
      return;
    });
    const args: any = {
      target,
      content: {
        FOO: {
          FOO1: {
            FOO2: "BAR2",
          },
        },
      },
    };
    const result = await driver.run(args, mockedDriverContext);
    chai.assert(result.isOk());
    if (result.isOk()) {
      chai.assert.equal(
        '{\n\t"FOO": {\n\t\t"FOO1": {\n\t\t\t"FOO2": "BAR2"\n\t\t}\n\t}\n}',
        content
      );
    }
  });

  it("happy path: add nested object to empty json", async () => {
    const target = "path";
    let content = {};
    const jsonContent = {
      BOT_ID: "$botId$",
    };
    sinon.stub(fs, "ensureFile").callsFake(async (path) => {
      return;
    });
    sinon.stub(fs, "readFile").callsFake(async (path) => {
      return Buffer.from(JSON.stringify(jsonContent));
    });
    sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
      content = data;
      return;
    });
    sinon.stub(fs, "pathExists").resolves(false);
    sinon.stub(fs, "existsSync").callsFake((path) => {
      return false;
    });
    sinon.stub(fs, "copyFile").callsFake(async (p1, p2) => {
      return;
    });
    const args: any = {
      target,
      content: {
        FOO: {
          FOO1: {
            FOO2: "BAR2",
          },
        },
      },
    };
    const result = await driver.run(args, mockedDriverContext);
    chai.assert(result.isOk());
    if (result.isOk()) {
      chai.assert.equal(
        '{\n\t"BOT_ID": "$botId$",\n\t"FOO": {\n\t\t"FOO1": {\n\t\t\t"FOO2": "BAR2"\n\t\t}\n\t}\n}',
        content
      );
    }
  });

  it("invalid path: no target path", async () => {
    let content = {};
    const jsonContent = {
      BOT_ID: "$botId$",
    };
    sinon.stub(fs, "ensureFile").callsFake(async (path) => {
      return;
    });
    sinon.stub(fs, "readFile").callsFake(async (path) => {
      return Buffer.from(JSON.stringify(jsonContent));
    });
    sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
      content = data;
      return;
    });
    sinon.stub(fs, "pathExists").resolves(false);
    sinon.stub(fs, "existsSync").callsFake((path) => {
      return false;
    });
    sinon.stub(fs, "copyFile").callsFake(async (p1, p2) => {
      return;
    });
    const args: any = {
      content: {
        FOO: {
          FOO1: {
            FOO2: "BAR2",
          },
        },
      },
    };
    const result = await driver.run(args, mockedDriverContext);
    chai.assert(result.isErr());
  });

  it("invalid path: no content and appsettings", async () => {
    const target = "path";
    let content = {};
    const jsonContent = {
      BOT_ID: "$botId$",
    };
    sinon.stub(fs, "ensureFile").callsFake(async (path) => {
      return;
    });
    sinon.stub(fs, "readFile").callsFake(async (path) => {
      return Buffer.from(JSON.stringify(jsonContent));
    });
    sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
      content = data;
      return;
    });
    sinon.stub(fs, "pathExists").resolves(false);
    sinon.stub(fs, "existsSync").callsFake((path) => {
      return false;
    });
    sinon.stub(fs, "copyFile").callsFake(async (p1, p2) => {
      return;
    });
    const args: any = {
      target,
    };
    const result = await driver.run(args, mockedDriverContext);
    chai.assert(result.isErr());
  });

  it("invalid path: content is not object", async () => {
    const target = "path";
    let content = {};
    const jsonContent = {
      BOT_ID: "$botId$",
    };
    sinon.stub(fs, "ensureFile").callsFake(async (path) => {
      return;
    });
    sinon.stub(fs, "readFile").callsFake(async (path) => {
      return Buffer.from(JSON.stringify(jsonContent));
    });
    sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
      content = data;
      return;
    });
    sinon.stub(fs, "pathExists").resolves(false);
    sinon.stub(fs, "existsSync").callsFake((path) => {
      return false;
    });
    sinon.stub(fs, "copyFile").callsFake(async (p1, p2) => {
      return;
    });
    const args: any = {
      target,
      content: "foo",
    };
    const result = await driver.run(args, mockedDriverContext);
    chai.assert(result.isErr());
  });
});
