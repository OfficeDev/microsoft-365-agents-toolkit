export class ApiError extends Error {
  constructor(message: string, public status: number = 500) {
    super(message);
  }
}

export class MissingAccessTokenError extends ApiError {
  constructor() {
    super("No access token provided", 401);
  }
}

export class InvalidAccessTokenError extends ApiError {
  constructor() {
    super("Invalid access token provided", 403);
  }
}
export class MissingContainerDisplayNameError extends ApiError {
  constructor() {
    super("displayName is required", 400);
  }
}

export class MissingContainerIdError extends ApiError {
  constructor() {
    super("containerId is required", 400);
  }
}
