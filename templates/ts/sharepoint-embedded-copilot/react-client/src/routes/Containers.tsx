import { Form, Link, useActionData, useLoaderData, useNavigate, useSubmit } from "react-router-dom";
import { ILoaderParams } from "../common/ILoaderParams";
import {
  Breadcrumb,
  BreadcrumbButton,
  BreadcrumbItem,
  Button,
  DataGrid,
  DataGridBody,
  DataGridCell,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridRow,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  Input,
  Label,
  OnSelectionChangeData,
  TableCellLayout,
  TableColumnDefinition,
  Tag,
  createTableColumn,
} from "@fluentui/react-components";
import { useState } from "react";
import { Spinner } from "@microsoft/mgt-react";
import { ChatController } from "../providers/ChatController";
import { ContainersApiProvider } from "../providers/ContainersApiProvider";
import {
  IContainer,
  IContainerClientCreateRequest,
} from "../../../common/schemas/ContainerSchemas";

export async function loader({ params }: ILoaderParams): Promise<IContainer[]> {
  return await ContainersApiProvider.instance.list();
}

export async function action({ params, request }: ILoaderParams) {
  const formData = await request.formData();
  const container = Object.fromEntries(formData) as IContainerClientCreateRequest;
  return await ContainersApiProvider.instance.create(container);
}

export const Containers: React.FunctionComponent = () => {
  const [selectedItems, setSelectedItems] = useState<string[]>([]);
  const [displayName, setDisplayName] = useState<string>("");
  const [description, setDescription] = useState<string>("");
  const [showCreateDialog, setShowCreateDialog] = useState<boolean>(false);
  const [showCreatingSpinner, setShowCreatingSpinner] = useState<boolean>(false);
  const navigate = useNavigate();
  const containers = useLoaderData() as IContainer[];
  const container = useActionData() as IContainer | undefined;

  const submit = useSubmit();

  const onSelectionChange = (e: any, data: OnSelectionChangeData) => {
    const selectedIds = Array.from(data.selectedItems) as string[];
    setSelectedItems(selectedIds);

    const selectedContainers = containers.filter((container) => selectedIds.includes(container.id));
    ChatController.instance.selectedContainers = selectedContainers;
  };

  const submitCreateContainer = async () => {
    setShowCreatingSpinner(true);
    const formData = new FormData();
    formData.append("displayName", displayName);
    formData.append("description", description);
    await submit(formData, { method: "POST" });
    setDisplayName("");
    setDescription("");
    setShowCreateDialog(false);
    setShowCreatingSpinner(false);
  };

  const columns: TableColumnDefinition<IContainer>[] = [
    createTableColumn({
      columnId: "displayName",
      renderHeaderCell: () => {
        return "Display name";
      },
      renderCell: (container) => {
        return (
          <TableCellLayout>
            <Link to={container.id}>{container.displayName}</Link>
          </TableCellLayout>
        );
      },
    }),
    createTableColumn({
      columnId: "description",
      renderHeaderCell: () => {
        return "Description";
      },
      renderCell: (container) => {
        return <TableCellLayout>{container.description}</TableCellLayout>;
      },
    }),
  ];

  const columnSizingOptions = {
    displayName: {
      minWidth: 250,
      defaultWidth: 250,
      idealWidth: 250,
    },
    description: {
      minWidth: 190,
      defaultWidth: 190,
    },
  };

  return (
    <div>
      <div className="view-container-breadcrumb">
        <Breadcrumb size="medium">
          <BreadcrumbItem>
            <BreadcrumbButton size="medium" onClick={() => navigate("/containers")}>
              Containers
            </BreadcrumbButton>
          </BreadcrumbItem>
        </Breadcrumb>
      </div>
      <Form>
        <Dialog open={showCreateDialog}>
          <DialogTrigger disableButtonEnhancement>
            <Button appearance="primary" onClick={() => setShowCreateDialog(true)}>
              Create Container
            </Button>
          </DialogTrigger>
          <DialogSurface>
            {!showCreatingSpinner && (
              <DialogBody>
                <DialogTitle>New Container</DialogTitle>
                <DialogContent className="create-container-content">
                  Create a new container
                  <Label>Display name</Label>
                  <Input
                    placeholder="Display name"
                    aria-label="name"
                    type="text"
                    name="displayName"
                    value={displayName}
                    onChange={(e) => setDisplayName(e.target.value)}
                  />
                  <Label>Description</Label>
                  <Input
                    placeholder="Description"
                    aria-label="Description"
                    type="text"
                    name="description"
                    value={description}
                    onChange={(e) => setDescription(e.target.value)}
                  />
                </DialogContent>
                <DialogActions>
                  <Button appearance="primary" type="submit" onClick={submitCreateContainer}>
                    Create
                  </Button>
                  <DialogTrigger disableButtonEnhancement>
                    <Button appearance="secondary" onClick={() => setShowCreateDialog(false)}>
                      Cancel
                    </Button>
                  </DialogTrigger>
                </DialogActions>
              </DialogBody>
            )}
            {showCreatingSpinner && (
              <>
                <Spinner />
                <p>Creating container...</p>
              </>
            )}
          </DialogSurface>
        </Dialog>
      </Form>
      <h2>Recent</h2>
      <DataGrid
        items={containers}
        columns={columns}
        getRowId={(item) => item.id}
        resizableColumns
        selectionMode="multiselect"
        columnSizingOptions={columnSizingOptions}
        selectedItems={selectedItems}
        onSelectionChange={onSelectionChange}
      >
        <DataGridHeader>
          <DataGridRow selectionCell={{ checkboxIndicator: { "aria-label": "Select row" } }}>
            {({ renderHeaderCell }) => (
              <DataGridHeaderCell>
                <b>{renderHeaderCell()}</b>
              </DataGridHeaderCell>
            )}
          </DataGridRow>
        </DataGridHeader>
        <DataGridBody<IContainer>>
          {({ item, rowId }) => (
            <DataGridRow<IContainer>
              key={rowId}
              selectionCell={{ checkboxIndicator: { "aria-label": "Select row" } }}
            >
              {({ renderCell, columnId }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
            </DataGridRow>
          )}
        </DataGridBody>
      </DataGrid>
    </div>
  );
};
