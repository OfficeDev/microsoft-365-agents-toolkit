import {
  app,
  HttpFunctionOptions,
  HttpRequest,
  HttpResponseInit,
  InvocationContext,
} from "@azure/functions";
import { OboAuthProvider } from "../providers/OboAuthProvider";
import { GraphProvider } from "../providers/GraphProvider";
import {
  IContainerClientCreateRequest,
  IContainerUpdateRequest,
} from "../../../common/schemas/ContainerSchemas";
import {
  ApiError,
  InvalidAccessTokenError,
  MissingContainerDisplayNameError,
} from "../common/Errors";
import { AppAuthProvider } from "../providers/AppAuthProvider";
import { JwtProvider } from "../providers/JwtProvider";

export async function registerContainerType(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const jwt = JwtProvider.fromAuthHeader(request.headers.get("Authorization"));
    if (!jwt || !(await jwt.authorize()) || !jwt.tid) {
      throw new InvalidAccessTokenError();
    }

    const graph = new GraphProvider(new OboAuthProvider(jwt));
    const spRootSiteUrl = await graph.getRootSiteUrl();
    if (!spRootSiteUrl) {
      throw new ApiError("Unable to fetch root site url");
    }
    const authProvider = new AppAuthProvider(jwt.tid, spRootSiteUrl);
    const token = await authProvider.getToken();
    const containerTypeId = process.env.SPE_CONTAINER_TYPE_ID!;
    const registerApi = `${spRootSiteUrl}/_api/v2.1/storageContainerTypes/${containerTypeId}/applicationPermissions`;
    const registerPayload = {
      value: [
        {
          appId: process.env.AZURE_CLIENT_ID,
          delegated: ["full"],
          appOnly: ["full"],
        },
        {
          appId: process.env.AZURE_SPA_CLIENT_ID,
          delegated: ["full"],
          appOnly: ["full"],
        },
      ],
    };

    const registerResponse = await fetch(registerApi, {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify(registerPayload),
    });
    return { jsonBody: await registerResponse.json() };
  } catch (error) {
    if (error instanceof ApiError) {
      return { status: error.status, body: error.message };
    }
    return { status: 500, body: `Register Container Type failed: ${error}` };
  }
}

app.http("registerContainerType", {
  methods: ["PUT"],
  authLevel: "anonymous",
  handler: registerContainerType,
});
