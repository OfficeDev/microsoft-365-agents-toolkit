import React, { useEffect, useState } from "react";
import {
  Dropdown,
  Option,
  OptionOnSelectData,
  SelectionEvents,
  makeStyles,
  useId,
} from "@fluentui/react-components";
import { IContainer } from "../../../common/schemas/ContainerSchemas";
import { ContainersApiProvider } from "../providers/ContainersApiProvider";

const containersApi = ContainersApiProvider.instance;

const useStyles = makeStyles({
  containerSelectorControls: {
    minWidth: "150px",
    maxWidth: "250px",
    width: "200px",
  },
});

export type IContainerSelectorProps = {
  selectedContainerId?: string;
  onContainerSelected?: (container: IContainer) => void;
  refreshTime?: string;
};

export const ContainerSelector: React.FunctionComponent<IContainerSelectorProps> = (
  props: IContainerSelectorProps
) => {
  const [containers, setContainers] = useState<IContainer[] | undefined>();
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedContainerId, setSelectedContainerId] = useState<string | undefined>(
    props.selectedContainerId
  );
  const containerSelector = useId("containerSelector");

  useEffect(() => {
    (async () => {
      setLoading(true);
      containersApi
        .list()
        .then(setContainers)
        .catch(console.error)
        .finally(() => setLoading(false));
    })();
  }, [props.refreshTime]);

  const onContainerDropdownChange = (event: SelectionEvents, data: OptionOnSelectData) => {
    const selected = containers?.find((container) => container.id === data.optionValue);
    if (selected) {
      setSelectedContainerId(selected.id);
      props.onContainerSelected?.(selected);
    }
  };

  const styles = useStyles();
  return (
    <>
      <Dropdown
        id={containerSelector}
        disabled={loading}
        placeholder="Select a Container"
        value={
          containers?.find((container) => container.id === selectedContainerId)?.displayName || ""
        }
        selectedOptions={selectedContainerId ? [selectedContainerId] : []}
        className={styles.containerSelectorControls}
        onOptionSelect={onContainerDropdownChange}
      >
        {containers?.map((option) => (
          <Option key={option.id} value={option.id}>
            {option.displayName}
          </Option>
        ))}
      </Dropdown>
    </>
  );
};
