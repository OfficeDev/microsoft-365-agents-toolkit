export interface IContainerClientCreateRequest {
  displayName: string;
  description?: string;
  itemMajorVersionLimit?: number;
  isItemVersioningEnabled?: boolean;
}

export interface IContainerServerCreateRequest extends IContainerClientCreateRequest {
  containerTypeId?: string;
}

export interface IContainerUpdateRequest extends IContainerServerCreateRequest {
  id?: string;
}

export interface IContainer extends IContainerUpdateRequest {
  id: string;
  status?: IContainerStatus;
  createdDateTime?: string;
  customProperties?: IContainerCustomProperties;
  permissions?: IContainerPermissions;
  storageUsedInBytes?: number;
  columns: IContainerColumn[];
  drive?: {
    webUrl: string;
  };
}

export interface IContainerColumn {
  id?: string;
  name?: string;
  displayName?: string;
  description?: string;
  indexed?: boolean;
  text?: {
    maxLength?: number;
  };
  boolean?: unknown;
  dateTime?: {
    format?: IContainerColumnDateTimeFormat;
  };
  currency?: {
    locale?: string;
  };
  choice?: {
    choices?: string[];
  };
  hyperlinkOrPicture?: {
    isPicture?: boolean;
  };
  number?: {
    maximum?: number;
    minimum?: number;
  };
  personOrGroup?: {
    chooseFromType?: IContainerColumnPersonOrGroupType;
  };
}

export type IContainerColumnPersonOrGroupType = "peopleOnly" | "peopleAndGroups";
export type IContainerColumnDateTimeFormat = "dateOnly" | "dateTime";
export type IContainerStatus = "active" | "inactive";

export type IContainerCustomProperties = {
  [key: string]: IContainerCustomProperty;
};

export type IContainerCustomProperty = {
  value: string;
  isSearchable: boolean;
};

export type IContainerPermissions = IContainerPermission[];

export type IContainerPermission = {
  id: string;
  roles: IContainerPermissionRole[];
  grantedToV2: {
    user: {
      userPrincipalName: string;
      email: string | null;
      displayName: string | null;
    };
  };
};

export type IContainerPermissionRole = "reader" | "writer" | "manager" | "owner";
