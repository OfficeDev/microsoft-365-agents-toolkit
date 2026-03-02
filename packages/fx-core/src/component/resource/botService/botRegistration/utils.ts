// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Qianhao Dong <qidon@microsoft.com>
 */
import { IBotRegistration } from "../appStudio/interfaces/IBotRegistration";

export class Utils {
  public static mergeIBotRegistration(
    local: IBotRegistration,
    remote: IBotRegistration
  ): IBotRegistration {
    return {
      botId: local.botId ?? remote.botId,
      name: local.name || remote.name,
      description: local.description || remote.description,
      iconUrl: local.iconUrl || remote.iconUrl,
      messagingEndpoint: local.messagingEndpoint || remote.messagingEndpoint,
      callingEndpoint: local.callingEndpoint || remote.callingEndpoint,
      configuredChannels: local.configuredChannels || remote.configuredChannels,
      isSingleTenant: local.isSingleTenant || remote.isSingleTenant,
    };
  }
}
