// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode, Inputs, OptionItem, SingleSelectQuestion } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../common/localizeUtils";
import { QuestionNames } from "./constants";
import { inputUserEmailQuestion } from "./other";

export enum ShareScopeOption {
  ShareAppWithTenantUsers = "tenant",
  ShareAppWithSpecificUsers = "users",
  ShareAppWithOwners = "owners",
}

export const MAX_SHARE_EMAILS = 20;

export function shareNode(): IQTreeNode {
  return {
    data: {
      type: "group",
    },
    children: [
      {
        data: shareOptionQuestion(),
        children: [
          {
            condition: (inputs: Inputs) => {
              return (
                inputs[QuestionNames.ShareScope] === ShareScopeOption.ShareAppWithOwners ||
                inputs[QuestionNames.ShareScope] === ShareScopeOption.ShareAppWithSpecificUsers
              );
            },
            data: inputUserEmailQuestion(
              getLocalizedString("core.shareOptionQuestion.emails.title"),
              "Email address of specific users or groups separated by comma.",
              false
            ),
          },
        ],
      },
    ],
  };
}

function shareOptionQuestion(): SingleSelectQuestion {
  return {
    name: QuestionNames.ShareScope,
    title: getLocalizedString("core.shareOptionQuestion.title"),
    type: "singleSelect",
    placeholder: getLocalizedString("core.shareOptionQuestion.placeholder"),
    staticOptions: [
      ShareOptions.shareWithTenant(),
      ShareOptions.shareWithUsers(),
      ShareOptions.shareWithOwners(),
    ],
  };
}

export class ShareOptions {
  static shareWithTenant(): OptionItem {
    return {
      id: ShareScopeOption.ShareAppWithTenantUsers,
      label: getLocalizedString("core.shareOptionQuestion.option.shareWithTenant"),
    };
  }

  static shareWithUsers(): OptionItem {
    return {
      id: ShareScopeOption.ShareAppWithSpecificUsers,
      label: getLocalizedString("core.shareOptionQuestion.option.shareWithUsers"),
    };
  }

  static shareWithOwners(): OptionItem {
    return {
      id: ShareScopeOption.ShareAppWithOwners,
      label: getLocalizedString("core.shareOptionQuestion.option.shareWithOwners"),
    };
  }
}

// export function removeSharedAccessNode(): IQTreeNode {
//   return {
//     data: {
//       type: "group",
//     },
//     children: [],
//   };
// }

// export function selectUsersToRemoveSharedAccess(): MultiSelectQuestion {
//   return {
//     name: QuestionNames.RemoveUsers,
//     title: getLocalizedString("core.selectUsersToRemoveShareAccess.title"),
//     type: "multiSelect",
//     cliDescription: getLocalizedString("core.selectUsersToRemoveShareAccess.title"),
//     staticOptions: [],
//     dynamicOptions: async (inputs: Inputs) => {
//       if (!inputs.projectPath) {
//         throw new Error("Project path is not defined");
//       }
//       const mosTokenRes = await TOOLS.tokenProvider.m365TokenProvider.getAccessToken({
//         scopes: [MosServiceScope],
//       });
//       if (mosTokenRes.isErr()) {
//         throw mosTokenRes.error;
//       }
//       const token = mosTokenRes.value;
//       const configRes = await parseShareAppActionYamlConfig(inputs.projectPath);
//       if (configRes.isErr()) {
//         throw configRes.error;
//       }
//       const titleId = configRes.value.titleId;
//       const agent = await PackageService.GetSharedInstance().previewApp(token, titleId);
//       if (agent.isErr()) {
//         throw agent.error;
//       }
//       if (!agent.value.owners || agent.value.owners.length === 0) {
//         throw new Error("No owner found in the agent");
//       }

//       const currentUserInfoRes = await CollaborationUtil.getCurrentUserInfo(
//         TOOLS.tokenProvider.m365TokenProvider
//       );
//       if (currentUserInfoRes.isErr()) {
//         throw currentUserInfoRes.error;
//       }
//       const operatorId = currentUserInfoRes.value.aadId;

//       const options: OptionItem[] = [];
//       for (const user of agent.value.owners) {
//         if (user.entityId === operatorId) {
//           continue;
//         }
//         const userInfo = await CollaborationUtil.getUserInfoFromId(
//           user.entityId,
//           TOOLS.tokenProvider.m365TokenProvider
//         );
//         if (userInfo) {
//           options.push({
//             id: userInfo.aadId,
//             label: userInfo.displayName,
//             description: userInfo.userPrincipalName,
//           });
//         }
//       }
//       return options;
//     },
//     skipValidation: true,
//   };
// }
