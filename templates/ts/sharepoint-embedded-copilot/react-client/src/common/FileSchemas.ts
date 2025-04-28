import {
  DriveItem,
  IdentitySet,
  ItemReference,
  Folder,
  File,
} from "@microsoft/microsoft-graph-types";

export class IDriveItem implements DriveItem {
  createdBy?: IdentitySet;
  createdDateTime?: string;
  description?: string;
  lastModifiedBy?: IdentitySet;
  lastModifiedDateTime?: string;
  parentReference?: ItemReference;
  webUrl?: string;
  webDavUrl?: string;
  "@microsoft.graph.downloadUrl"?: string;
  size?: number;
  folder?: Folder;
  file?: File;
  root?: Record<string, never>;
  listItem?: {
    fields: {
      [key: string]: any;
    };
  };

  public constructor(public id: string, public name: string) {}

  public get modifiedByName(): string {
    return this.lastModifiedBy?.user?.displayName || "";
  }

  public get isFolder(): boolean {
    return this.folder !== undefined;
  }

  public get isFile(): boolean {
    return this.file !== undefined;
  }

  public get extension(): string | undefined {
    return this.isFolder ? undefined : IDriveItem.parseExtension(this.name);
  }

  public get isOfficeDocument(): boolean {
    return this.isFolder ? false : IDriveItem.isOfficeExtension(this.extension);
  }

  public get isWordDocument(): boolean {
    return this.isFolder ? false : IDriveItem._wordExtensions.includes(this.extension || "");
  }

  public get isExcelDocument(): boolean {
    return this.isFolder ? false : IDriveItem._excelExtensions.includes(this.extension || "");
  }

  public get isPowerPointDocument(): boolean {
    return this.isFolder ? false : IDriveItem._powerPointExtensions.includes(this.extension || "");
  }

  public get isPdfConvertibleDocument(): boolean {
    return this.isFolder
      ? false
      : IDriveItem._pdfConvertibleExtensions.includes(this.extension || "");
  }

  public get downloadUrl(): string | undefined {
    return this["@microsoft.graph.downloadUrl"];
  }

  public get desktopUrl(): string | undefined {
    if (!this.isOfficeDocument || !this.webDavUrl) {
      return undefined;
    }

    let protocol = "";
    if (this.isWordDocument) {
      protocol = "ms-word";
    } else if (this.isExcelDocument) {
      protocol = "ms-excel";
    } else if (this.isPowerPointDocument) {
      protocol = "ms-powerpoint";
    }
    return `${protocol}:ofe|u|${this.webDavUrl}`;
  }

  public column(id: string) {
    return this.listItem?.fields?.[id] || "";
  }

  private static readonly _wordExtensions: string[] = [
    "doc",
    "docx",
    "docm",
    "dot",
    "dotx",
    "dotm",
  ];
  private static readonly _excelExtensions: string[] = [
    "xls",
    "xlsx",
    "xlsm",
    "xlt",
    "xltx",
    "xltm",
  ];
  private static readonly _powerPointExtensions: string[] = [
    "ppt",
    "pptx",
    "pptm",
    "pot",
    "potx",
    "potm",
    "pps",
    "ppsx",
    "ppsm",
  ];
  private static readonly _pdfConvertibleExtensions: string[] = [
    "csv",
    "doc",
    "docx",
    "odp",
    "ods",
    "odt",
    "pot",
    "potm",
    "potx",
    "pps",
    "ppsx",
    "ppsm",
    "ppt",
    "pptm",
    "pptx",
    "rtf",
    "xls",
    "xlsx",
  ];
  private static readonly _allExtensions: string[] = [
    ...IDriveItem._wordExtensions,
    ...IDriveItem._excelExtensions,
    ...IDriveItem._powerPointExtensions,
  ];

  public static isOfficeExtension(extension: string | undefined): boolean {
    if (!extension) {
      return false;
    }
    return IDriveItem._allExtensions.includes(extension);
  }

  public static parseExtension(name: string | undefined): string | undefined {
    return name?.toLocaleLowerCase().split(".").pop();
  }
}

export abstract class DriveItemConstructor {
  public static from(item: DriveItem): IDriveItem {
    if (!item) {
      throw new Error("Item is required");
    }
    if (!item.id) {
      throw new Error("Item id is required");
    }
    if (!item.name) {
      throw new Error("Item name is required");
    }
    return Object.assign(new IDriveItem(item.id, item.name), item);
  }
}

export abstract class DriveItemArrayConstructor {
  public static from(items: DriveItem[]): IDriveItem[] {
    if (!items) {
      throw new Error("Items are required");
    }
    return items.map((item: DriveItem) => DriveItemConstructor.from(item));
  }
}
