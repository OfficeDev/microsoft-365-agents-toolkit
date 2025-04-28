import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Input,
  InputOnChangeData,
  InputProps,
  Label,
  Spinner,
  makeStyles,
} from "@fluentui/react-components";
import { IContainer } from "../../../common/schemas/ContainerSchemas";
import { useEffect, useState } from "react";
import { ContainersApiProvider } from "../providers/ContainersApiProvider";
import { Settings20Filled } from "@fluentui/react-icons";

const containersApi = ContainersApiProvider.instance;

const useStyles = makeStyles({
  containerSelectorControls: {
    width: "400px",
  },
  dialogContent: {
    display: "flex",
    flexDirection: "column",
    rowGap: "10px",
    marginBottom: "25px",
  },
});

export type IContainerSettingsDialogProps = {
  container: IContainer;
  isOpen?: boolean;
  onAbort?: () => void;
  onContainerUpdated?: (container: IContainer) => void;
};

export const ContainerSettingsDialog: React.FunctionComponent<IContainerSettingsDialogProps> = (
  props: IContainerSettingsDialogProps
) => {
  const [container, setContainer] = useState<IContainer>(props.container);
  const [id, setId] = useState(props.container.id);
  const [displayName, setDisplayName] = useState(props.container.displayName || "");
  const [description, setDescription] = useState(props.container.description || "");
  const [customProperties, setCustomProperties] = useState(props.container.customProperties || {});
  const [isOpen, setIsOpen] = useState(props.isOpen);
  const [loaded, setLoaded] = useState(false);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    containersApi
      .get(props.container.id)
      .then(setContainer)
      .then(() => setLoaded(true));
  }, [props.container.id]);

  const handleDisplayNameChange: InputProps["onChange"] = (
    event: React.ChangeEvent<HTMLInputElement>,
    data: InputOnChangeData
  ): void => {
    setDisplayName(data?.value);
  };

  const handleDescriptionChange: InputProps["onChange"] = (
    event: React.ChangeEvent<HTMLInputElement>,
    data: InputOnChangeData
  ): void => {
    setDescription(data?.value);
  };

  const styles = useStyles();
  return (
    <>
      <Button icon={<Settings20Filled />} size="small" onClick={() => setIsOpen(true)}>
        Settings
      </Button>
      <Dialog open={isOpen}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Container Settings</DialogTitle>
            <DialogContent className={styles.dialogContent}>
              <Label htmlFor={container.id}>Id:</Label>
              <Input value={id} className={styles.containerSelectorControls} disabled />
              <Label htmlFor={container.displayName}>Name:</Label>
              <Input
                autoFocus
                required
                value={displayName}
                onChange={handleDisplayNameChange}
                className={styles.containerSelectorControls}
              />
              <Label htmlFor={container.description}>Description:</Label>
              <Input
                autoFocus
                value={description}
                onChange={handleDescriptionChange}
                className={styles.containerSelectorControls}
              ></Input>

              {saving && (
                <Spinner size="medium" label="Updated storage Container..." labelPosition="after" />
              )}
            </DialogContent>
            <DialogActions>
              <Button appearance="primary" disabled={saving || container.displayName === ""}>
                Create
              </Button>
              <Button onClick={() => setIsOpen(false)} appearance="secondary">
                Cancel
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
};
