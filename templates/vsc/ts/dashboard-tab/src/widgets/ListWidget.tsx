import "../styles/ListWidget.css";

import { Button, Text } from "@fluentui/react-components";
import { List28Filled, MoreHorizontal32Regular } from "@fluentui/react-icons";
import { BaseWidget } from "@microsoft/teamsfx-react";

import { ListModel } from "../models/listModel";
import { getListData } from "../services/listService";

interface IListWidgetState {
  data: ListModel[];
}

export default class ListWidget extends BaseWidget<any, IListWidgetState> {
  async getData(): Promise<IListWidgetState> {
    return { data: getListData() };
  }

  header(): JSX.Element | undefined {
    return (
      <div>
        <List28Filled />
        <Text>Your List</Text>
        <Button icon={<MoreHorizontal32Regular />} appearance="transparent" />
      </div>
    );
  }

  body(): JSX.Element | undefined {
    return (
      <div className="list-body">
        {this.state.data?.map((t: ListModel) => {
          return (
            <div key={`${t.id}-div`}>
              <div className="divider" />
              <Text className="title">{t.title}</Text>
              <Text className="content">{t.content}</Text>
            </div>
          );
        })}
      </div>
    );
  }

  footer(): JSX.Element | undefined {
    return <Button appearance="primary">View Details</Button>;
  }
}
