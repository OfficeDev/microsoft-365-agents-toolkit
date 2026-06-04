import { Context } from "@microsoft/teamsfx-api";
import { AxiosRequestConfig, default as axios } from "axios";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import mockFs from "mock-fs";
import * as sinon from "sinon";
import * as stream from "stream";
import { createContext, setTools } from "../../../../src/common/globalVars";
import { ensureBicepForDriver } from "../../../../src/component/driver/arm/util/bicepChecker";
import { DriverContext } from "../../../../src/component/driver/interface/commonArgs";
import { cpUtils } from "../../../../src/component/utils/depsChecker/cpUtils";
import { MockTools } from "../../../core/utils";

chai.use(chaiAsPromised);

function createFakeAxiosInstance(sandbox: sinon.SinonSandbox) {
  const fakeAxiosInstance = axios.create();
  sandbox.stub(axios, "create").returns(fakeAxiosInstance);
  return fakeAxiosInstance;
}

const mockBicepVersion = "0.4.1318";
const bicepReleaseApiUrl = "https://api.github.com/repos/Azure/bicep/releases";
const bicepDownloadUrlPrefix = "https://github.com/Azure/bicep/releases/download/";

describe("BicepChecker", () => {
  let sandbox: sinon.SinonSandbox;
  let downloaded: boolean;
  let context: Context;

  beforeEach(() => {
    sandbox = sinon.createSandbox();
    // prevent actually touching real file system
    mockFs({});

    downloaded = false;

    sandbox
      .stub(cpUtils, "executeCommand")
      .callsFake(
        async (
          workDir: string | undefined,
          logger: any,
          options: any,
          command: string,
          ...args: string[]
        ): Promise<string> => {
          if (command === "bicep") {
            throw new Error("Global bicep not installed");
          } else if (args.includes("--version")) {
            if (downloaded) {
              return `Bicep CLI version ${mockBicepVersion}`;
            } else {
              throw new Error("bicep command not found");
            }
          } else {
            throw new Error("Not implemented");
          }
        }
      );

    const tools = new MockTools();
    setTools(tools);
    context = createContext();
  });

  afterEach(() => {
    sandbox.restore();
    mockFs.restore();
  });

  it("Timeout for downloading bicep", async () => {
    const axiosInstance = createFakeAxiosInstance(sandbox);
    sandbox
      .stub(axiosInstance, "get")
      .callsFake(async (url: string, config?: AxiosRequestConfig) => {
        if (url === bicepReleaseApiUrl) {
          return {
            data: [{ tag_name: "v" + mockBicepVersion }],
          };
        } else if (url.startsWith(bicepDownloadUrlPrefix)) {
          const reader = new stream.Readable({
            read(size) {
              // mock a timeout error
              // https://nodejs.org/api/stream.html#errors-while-reading
              this.destroy(new Error("Timeout error"));
            },
          });

          return {
            data: reader,
          };
        } else {
          throw new Error(`Not implemented`);
        }
      });

    // If timeout is not handled, there will be unhandled promise rejection but it seems chai has no way to assert that
    await chai
      .expect(ensureBicepForDriver(context as unknown as DriverContext, "v0.9.1"))
      .to.be.rejectedWith(/Unable to download/);
  });
});
