import React, { useEffect, useRef, useState } from "react";
import {
  makeStyles,
  shorthands,
  Button,
  Menu,
  MenuButton,
  MenuList,
  MenuItem,
  MenuPopover,
  MenuTrigger,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  Label,
  Input,
  DialogActions,
  Spinner,
  SwitchOnChangeData,
  Switch,
} from "@fluentui/react-components";
import {
  ArrowUpload20Filled,
  Rename20Filled,
  Delete20Filled,
  Open20Filled,
  Globe20Filled,
  PreviewLink20Filled,
  Add20Filled,
  Folder24Filled,
  DocumentPdf20Regular,
  ArrowDownload20Filled,
} from "@fluentui/react-icons";
import { useLoaderData, useRevalidator } from "react-router-dom";
import { IDriveItem } from "../common/FileSchemas";
import { GraphProvider } from "../providers/GraphProvider";
import { getFileTypeIconProps } from "@fluentui/react-file-type-icons";
import { Icon } from "@fluentui/react";
import { ContainerSettingsDialog } from "./ContainerSettingsDialog";
import { ContainersApiProvider } from "../providers/ContainersApiProvider";
import { IContainerLoader } from "./IContainerLoader";

const containersApi = ContainersApiProvider.instance;
const filesApi = GraphProvider.instance;

const useStyles = makeStyles({
  actionBar: {
    columnGap: "2px",
    display: "flex",
    width: "100%",
    fontSize: "10px",
    backgroundColor: "white",
    boxShadow: "0 4px 8px 0 rgba(0, 0, 0, 0.1), 0 6px 20px 0 rgba(0, 0, 0, 0.05)",
    marginBottom: "20px",
    ...shorthands.borderRadius("10px"),
    ...shorthands.padding("10px"),
  },
  processingSwitch: {
    alignContent: "flex-end",
    textAlign: "left",
    alignItems: "left",
    justifyContent: "left",
  },
});

type IPendingUpload = {
  driveId: string;
  parentId: string;
  file: File;
  uploadTask: Promise<IDriveItem>;
};

export interface IContainerActionBarProps {
  containerId: string;
  parentId: string;
  selectedItem?: IDriveItem;
  onFilePreviewSelected?: (file: IDriveItem) => void;
  onItemsUpdated?: () => void;
}

export const ContainerActionBar: React.FunctionComponent<IContainerActionBarProps> = (
  props: IContainerActionBarProps
) => {
  const { container } = useLoaderData() as IContainerLoader;

  const [showContainerSettings] = useState(false);
  const [uploads, setUploads] = useState<Map<string, IPendingUpload>>(
    new Map<string, IPendingUpload>()
  );

  const [showNewFolderDialog, setShowNewFolderDialog] = useState<boolean>(false);
  const [newFolderName, setNewFolderName] = useState<string>("");
  const [showCreatingSpinner, setShowCreatingSpinner] = useState<boolean>(false);

  const [showRenameDialog, setShowRenameDialog] = useState<boolean>(false);
  const [newName, setNewName] = useState<string>("");
  const [showRenamingSpinner, setShowRenamingSpinner] = useState<boolean>(false);

  const [showDeleteDialog, setShowDeleteDialog] = useState<boolean>(false);
  const [showDeletingSpinner, setShowDeletingSpinner] = useState<boolean>(false);

  const uploadFileRef = useRef<HTMLInputElement>(null);

  const onUploadFileClick = () => {
    if (uploadFileRef.current) {
      uploadFileRef.current.click();
    }
  };

  const onUploadFileSelected = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) {
      return;
    }
    for (let i = 0; i < files.length; i++) {
      const upload: IPendingUpload = {
        driveId: props.containerId,
        parentId: props.parentId,
        file: files[i],
        uploadTask: filesApi.uploadFile(props.containerId, files[i], props.parentId),
      };
      const uploadId = `${upload.driveId}/${upload.parentId}/${files[i].name}`;
      uploads.set(uploadId, upload);
      upload.uploadTask.then(() => {
        uploads.delete(uploadId);
        setUploads(new Map<string, IPendingUpload>(uploads));
        if (uploads.size === 0) {
          props.onItemsUpdated?.();
        }
      });
    }
    setUploads(new Map<string, IPendingUpload>(uploads));
  };

  const createNewFolder = async () => {
    setShowCreatingSpinner(true);
    await filesApi.createFolder(props.containerId, props.parentId, newFolderName);
    setShowCreatingSpinner(false);
    setNewFolderName("");
    setShowNewFolderDialog(false);
    props.onItemsUpdated?.();
  };

  const onDownloadClick = async () => {
    if (!props.selectedItem || !props.selectedItem.isFile || !props.selectedItem.downloadUrl) {
      return;
    }
    window.open(props.selectedItem.downloadUrl, "_blank");
  };

  const onRenameClick = () => {
    if (props.selectedItem === undefined) {
      return;
    }
    setNewName(props.selectedItem.name);
    setShowRenameDialog(true);
  };

  const renameItem = async () => {
    if (props.selectedItem === undefined) {
      setShowRenameDialog(false);
      setShowRenamingSpinner(false);
      return;
    }
    setShowRenamingSpinner(true);
    await filesApi.renameItem(props.containerId, props.selectedItem.id, newName);
    setShowRenamingSpinner(false);
    setNewName("");
    setShowRenameDialog(false);
    props.onItemsUpdated?.();
  };

  const onDeleteClick = () => {
    if (props.selectedItem === undefined) {
      return;
    }
    setShowDeleteDialog(true);
  };

  const onDelete = async () => {
    if (props.selectedItem === undefined) {
      setShowDeleteDialog(false);
      setShowDeletingSpinner(false);
      return;
    }
    setShowDeletingSpinner(true);
    await filesApi.deleteItem(props.containerId, props.selectedItem.id);
    setShowDeletingSpinner(false);
    setShowDeleteDialog(false);
    props.onItemsUpdated?.();
  };

  const onNewDocument = async (extension: string) => {
    const newItem = await filesApi.newDocument(props.containerId, props.parentId, extension);
    props.onItemsUpdated?.();
    window.open(newItem.webUrl, "_blank");
  };

  const styles = useStyles();
  return (
    <div className={styles.actionBar}>
      <input
        ref={uploadFileRef}
        type="file"
        multiple
        onChange={onUploadFileSelected}
        style={{ display: "none" }}
      />
      <Button
        onClick={onUploadFileClick}
        appearance="primary"
        icon={<ArrowUpload20Filled />}
        size="small"
      >
        Upload
      </Button>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton icon={<Add20Filled />} appearance="secondary" size="small">
            New
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem
              icon={<Icon {...getFileTypeIconProps({ extension: "docx", size: 20 })} />}
              onClick={() => onNewDocument("docx")}
            >
              Word document
            </MenuItem>
            <MenuItem
              icon={<Icon {...getFileTypeIconProps({ extension: "pptx", size: 20 })} />}
              onClick={() => onNewDocument("pptx")}
            >
              PowerPoint document
            </MenuItem>
            <MenuItem
              icon={<Icon {...getFileTypeIconProps({ extension: "xlsx", size: 20 })} />}
              onClick={() => onNewDocument("xlsx")}
            >
              Excel document
            </MenuItem>
            <MenuItem
              icon={<Folder24Filled primaryFill="#FFCE3D" />}
              onClick={() => setShowNewFolderDialog(true)}
            >
              Folder
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      {props.selectedItem && props.selectedItem.isFile && (
        <>
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <MenuButton icon={<Open20Filled />} appearance="subtle" size="small">
                Open
              </MenuButton>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                {props.selectedItem.isOfficeDocument && (
                  <>
                    {props.selectedItem.desktopUrl && (
                      <MenuItem
                        icon={
                          <Icon
                            {...getFileTypeIconProps({
                              extension: props.selectedItem.extension,
                              size: 20,
                            })}
                          />
                        }
                        onClick={() => window.open(props.selectedItem!.desktopUrl)}
                      >
                        Open in desktop
                      </MenuItem>
                    )}
                    {props.selectedItem.webUrl && (
                      <MenuItem
                        icon={<Globe20Filled />}
                        onClick={() => window.open(props.selectedItem!.webUrl, "_blank")}
                      >
                        Open in web
                      </MenuItem>
                    )}
                  </>
                )}

                {props.selectedItem.isPdfConvertibleDocument && (
                  <MenuItem
                    icon={<DocumentPdf20Regular />}
                    onClick={async () => {
                      const pdfUrl = await filesApi.getPdfUrl(
                        props.selectedItem?.parentReference?.driveId ?? "",
                        props.selectedItem?.id ?? ""
                      );
                      window.open(pdfUrl, "_blank");
                    }}
                  >
                    Open as PDF
                  </MenuItem>
                )}

                <MenuItem
                  icon={<PreviewLink20Filled />}
                  onClick={() => props.onFilePreviewSelected?.(props.selectedItem!)}
                >
                  Preview in web
                </MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
          <Button
            onClick={onDownloadClick}
            icon={<ArrowDownload20Filled />}
            size="small"
            appearance="subtle"
          >
            Download
          </Button>
        </>
      )}

      {props.selectedItem && (
        <>
          {/*<Button icon={<Share20Filled />} size='small' appearance='subtle'>Share</Button>*/}
          <Button
            onClick={onRenameClick}
            icon={<Rename20Filled />}
            size="small"
            appearance="subtle"
          >
            Rename
          </Button>
          <Button
            onClick={onDeleteClick}
            icon={<Delete20Filled />}
            size="small"
            appearance="subtle"
          >
            Delete
          </Button>
        </>
      )}

      {container && (
        <ContainerSettingsDialog isOpen={showContainerSettings} container={container} />
      )}

      {uploads.size > 0 && (
        <Button disabled={true}>
          {uploads.size} files uploading
          <Spinner size="extra-tiny" />
        </Button>
      )}

      <Dialog open={showNewFolderDialog}>
        <DialogSurface>
          {!showCreatingSpinner && (
            <DialogBody>
              <DialogTitle>New Folder</DialogTitle>
              <DialogContent>
                <Label htmlFor="newFolderName">Folder name:</Label>
                <Input
                  placeholder="Folder name"
                  id="newFolderName"
                  aria-label="newFolderName"
                  type="text"
                  name="newFolderName"
                  value={newFolderName}
                  onChange={(e) => setNewFolderName(e.target.value)}
                />
              </DialogContent>
              <DialogActions>
                <Button appearance="primary" onClick={() => createNewFolder()}>
                  Create
                </Button>
                <Button appearance="secondary" onClick={() => setShowNewFolderDialog(false)}>
                  Cancel
                </Button>
              </DialogActions>
            </DialogBody>
          )}
          {showCreatingSpinner && (
            <>
              <Spinner />
              <p>Creating folder...</p>
            </>
          )}
        </DialogSurface>
      </Dialog>

      <Dialog open={showRenameDialog}>
        <DialogSurface>
          {!showRenamingSpinner && (
            <DialogBody>
              <DialogTitle>Rename</DialogTitle>
              <DialogContent>
                <Label htmlFor="newName">New name:</Label>
                <Input
                  placeholder="New name"
                  id="newName"
                  aria-label="newName"
                  type="text"
                  name="newName"
                  value={newName}
                  onChange={(e) => setNewName(e.target.value)}
                />
              </DialogContent>
              <DialogActions>
                <Button appearance="primary" onClick={() => renameItem()}>
                  Rename
                </Button>
                <Button appearance="secondary" onClick={() => setShowRenameDialog(false)}>
                  Cancel
                </Button>
              </DialogActions>
            </DialogBody>
          )}
          {showRenamingSpinner && (
            <>
              <Spinner />
              <p>Renaming...</p>
            </>
          )}
        </DialogSurface>
      </Dialog>

      <Dialog open={showDeleteDialog}>
        <DialogSurface>
          {!showDeletingSpinner && (
            <DialogBody>
              <DialogTitle>Delete</DialogTitle>
              <DialogContent>
                <p>Are you sure you want to delete {props.selectedItem?.name}?</p>
              </DialogContent>
              <DialogActions>
                <Button appearance="primary" onClick={() => onDelete()}>
                  Delete
                </Button>
                <Button appearance="secondary" onClick={() => setShowDeleteDialog(false)}>
                  Cancel
                </Button>
              </DialogActions>
            </DialogBody>
          )}
          {showDeletingSpinner && (
            <>
              <Spinner />
              <p>Deleting {props.selectedItem?.name}...</p>
            </>
          )}
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default ContainerActionBar;
