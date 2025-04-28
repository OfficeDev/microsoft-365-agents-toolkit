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

export async function containers(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  return request.method === "POST"
    ? createContainer(request, context)
    : listContainers(request, context);
}

export async function listContainers(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const jwt = JwtProvider.fromAuthHeader(request.headers.get("Authorization"));
    const test = !(await jwt!.authorize());
    if (!jwt || !(await jwt.authorize()) || !jwt.tid) {
      throw new InvalidAccessTokenError();
    }
    const authProvider = new AppAuthProvider(jwt.tid);
    const token = await authProvider.getToken();
    const graph = new GraphProvider(authProvider);
    const containers = await graph.listContainers();
    const hydratedContainers = containers.map((c) => graph.getContainer(c.id));
    return { jsonBody: await Promise.all(hydratedContainers) };
  } catch (error) {
    if (error instanceof ApiError) {
      return { status: error.status, body: error.message };
    }
    return { status: 500, body: `List containers failed: ${error}` };
  }
}

export async function createContainer(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const jwt = JwtProvider.fromAuthHeader(request.headers.get("Authorization"));
    if (!jwt || !(await jwt.authorize()) || !jwt.tid) {
      throw new InvalidAccessTokenError();
    }
    const graph = new GraphProvider(new OboAuthProvider(jwt));
    const clientCreateRequest: IContainerClientCreateRequest =
      (await request.json()) as IContainerClientCreateRequest;
    if (!clientCreateRequest.displayName) {
      throw new MissingContainerDisplayNameError();
    }
    return { jsonBody: await graph.createContainer(clientCreateRequest) };
  } catch (error) {
    if (error instanceof ApiError) {
      return { status: error.status, body: error.message };
    }
    return { status: 500, body: `Create container failed: ${error}` };
  }
}

export async function container(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  return request.method === "PATCH"
    ? updateContainer(request, context)
    : getContainer(request, context);
}

export async function updateContainer(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const jwt = JwtProvider.fromAuthHeader(request.headers.get("Authorization"));
    if (!jwt || !(await jwt.authorize()) || !jwt.tid) {
      throw new InvalidAccessTokenError();
    }
    const graph = new GraphProvider(new OboAuthProvider(jwt));
    const id = request.params.id;
    const updateRequest: IContainerUpdateRequest =
      (await request.json()) as IContainerUpdateRequest;
    return { jsonBody: await graph.updateContainer(id, updateRequest) };
  } catch (error) {
    if (error instanceof ApiError) {
      return { status: error.status, body: error.message };
    }
    return { status: 500, body: `Update container failed: ${error}` };
  }
}

export async function getContainer(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const jwt = JwtProvider.fromAuthHeader(request.headers.get("Authorization"));
    if (!jwt || !(await jwt.authorize()) || !jwt.tid) {
      throw new InvalidAccessTokenError();
    }
    const graph = new GraphProvider(new OboAuthProvider(jwt));
    const id = request.params.id;
    return { jsonBody: await graph.getContainer(id) };
  } catch (error) {
    if (error instanceof ApiError) {
      return { status: error.status, body: error.message };
    }
    return { status: 500, body: `Get container failed: ${error}` };
  }
}

app.http("containers", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  handler: containers,
});

app.http("container", {
  handler: container,
  trigger: {
    type: "httpTrigger",
    name: "container",
    authLevel: "anonymous",
    methods: ["GET", "PATCH"],
    route: "containers/{id}",
  },
});
