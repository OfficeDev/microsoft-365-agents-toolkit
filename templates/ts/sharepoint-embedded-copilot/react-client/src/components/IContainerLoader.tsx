import { IContainer } from "../../../common/schemas/ContainerSchemas";
import { IDriveItem } from "../common/FileSchemas";

export interface IContainerLoader {
  container: IContainer | undefined;
  parent: IDriveItem | undefined;
  driveItems: IDriveItem[];
}
