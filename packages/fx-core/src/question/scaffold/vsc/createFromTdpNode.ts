// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Inputs, IQTreeNode } from "@microsoft/teamsfx-api";
import { AppDefinition } from "../../../component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { isPersonalApp, needBotCode } from "../../../component/driver/teamsApp/utils/utils";
import { QuestionNames } from "../../constants";
import {
  selectBotIdsQuestion,
  selectTabsContentUrlQuestion,
  selectTabWebsiteUrlQuestion,
} from "../../create";

export class CreateFromTdpCapabilityOptions {}

export function createFromTdpNode(): IQTreeNode {
  const node: IQTreeNode = {
    data: { type: "group" },
    children: [
      {
        condition: (inputs: Inputs) =>
          inputs.teamsAppFromTdp && isPersonalApp(inputs.teamsAppFromTdp),
        data: { type: "group", name: QuestionNames.RepalceTabUrl },
        children: [
          {
            condition: (inputs: Inputs) =>
              (inputs.teamsAppFromTdp?.staticTabs.filter((o: any) => !!o.websiteUrl) || []).length >
              0,
            data: selectTabWebsiteUrlQuestion(),
          },
          {
            condition: (inputs: Inputs) =>
              (inputs.teamsAppFromTdp?.staticTabs.filter((o: any) => !!o.contentUrl) || []).length >
              0,
            data: selectTabsContentUrlQuestion(),
          },
        ],
      },
      {
        condition: (inputs: Inputs) =>
          inputs.teamsAppFromTdp && needBotCode(inputs.teamsAppFromTdp as AppDefinition),
        data: selectBotIdsQuestion(),
      },
    ],
  };
  return node;
}
