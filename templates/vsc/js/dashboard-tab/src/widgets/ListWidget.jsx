import "../styles/ListWidget.css";

import { Button, Text } from "@fluentui/react-components";
import { List28Filled, MoreHorizontal32Regular } from "@fluentui/react-icons";
import { BaseWidget } from "@microsoft/teamsfx-react";

import { getListData } from "../services/listService";

export default class ListWidget extends BaseWidget {
  async getData() {
    return { data: getListData() };
  }

  header() {
    return (
      <div>
        <List28Filled />
        <Text>Your List</Text>
        <Button icon={<MoreHorizontal32Regular />} appearance="transparent" />
      </div>
    );
  }

  body() {
    return (
      <div className="list-body">
        {this.state.data?.map((t) => {
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

  footer() {
    return <Button appearance="primary">View Details</Button>;
  }
}
