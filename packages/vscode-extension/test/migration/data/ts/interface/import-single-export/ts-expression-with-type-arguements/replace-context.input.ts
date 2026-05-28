import { Context } from "@microsoft/teams-js";
import { Context as ContextAlias } from "@microsoft/teams-js";

interface ITestContext extends Context {
  x: number;
}

class TestContext implements Context {
  constructor(entityId: string, locale: string) {
    this.entityId = entityId;
    this.locale = locale;
  }

  groupId?: string | undefined;
  teamId?: string | undefined;
  teamName?: string | undefined;
  channelId?: string | undefined;
  channelName?: string | undefined;
  channelType?: microsoftTeams.ChannelType | undefined;
  entityId: string;
  subEntityId?: string | undefined;
  locale: string;
  osLocaleInfo?: microsoftTeams.LocaleInfo | undefined;
  upn?: string | undefined;
  tid?: string | undefined;
  theme?: string | undefined;
  isFullScreen?: boolean | undefined;
  teamType?: microsoftTeams.TeamType | undefined;
  teamSiteUrl?: string | undefined;
  teamSiteDomain?: string | undefined;
  teamSitePath?: string | undefined;
  hostTeamTenantId?: string | undefined;
  hostTeamGroupId?: string | undefined;
  channelRelativeUrl?: string | undefined;
  sessionId?: string | undefined;
  userTeamRole?: microsoftTeams.UserTeamRole | undefined;
  chatId?: string | undefined;
  loginHint?: string | undefined;
  userPrincipalName?: string | undefined;
  userObjectId?: string | undefined;
  isTeamArchived?: boolean | undefined;
  hostClientType?: microsoftTeams.HostClientType | undefined;
  frameContext?: microsoftTeams.FrameContexts | undefined;
  sharepoint?: any;
  tenantSKU?: string | undefined;
  userLicenseType?: string | undefined;
  parentMessageId?: string | undefined;
  ringId?: string | undefined;
  appSessionId?: string | undefined;
  isCallingAllowed?: boolean | undefined;
  isPSTNCallingAllowed?: boolean | undefined;
  meetingId?: string | undefined;
  defaultOneNoteSectionId?: string | undefined;
  isMultiWindow?: boolean | undefined;
  appIconPosition?: number | undefined;
  sourceOrigin?: string | undefined;
  userClickTime?: number | undefined;
  teamTemplateId?: string | undefined;
  userFileOpenPreference?: microsoftTeams.FileOpenPreference | undefined;
}

interface ITestContextAlias extends ContextAlias {
  x: number;
}

class TestContextAlias implements ContextAlias {
  constructor(entityId: string, locale: string) {
    this.entityId = entityId;
    this.locale = locale;
  }

  groupId?: string | undefined;
  teamId?: string | undefined;
  teamName?: string | undefined;
  channelId?: string | undefined;
  channelName?: string | undefined;
  channelType?: microsoftTeams.ChannelType | undefined;
  entityId: string;
  subEntityId?: string | undefined;
  locale: string;
  osLocaleInfo?: microsoftTeams.LocaleInfo | undefined;
  upn?: string | undefined;
  tid?: string | undefined;
  theme?: string | undefined;
  isFullScreen?: boolean | undefined;
  teamType?: microsoftTeams.TeamType | undefined;
  teamSiteUrl?: string | undefined;
  teamSiteDomain?: string | undefined;
  teamSitePath?: string | undefined;
  hostTeamTenantId?: string | undefined;
  hostTeamGroupId?: string | undefined;
  channelRelativeUrl?: string | undefined;
  sessionId?: string | undefined;
  userTeamRole?: microsoftTeams.UserTeamRole | undefined;
  chatId?: string | undefined;
  loginHint?: string | undefined;
  userPrincipalName?: string | undefined;
  userObjectId?: string | undefined;
  isTeamArchived?: boolean | undefined;
  hostClientType?: microsoftTeams.HostClientType | undefined;
  frameContext?: microsoftTeams.FrameContexts | undefined;
  sharepoint?: any;
  tenantSKU?: string | undefined;
  userLicenseType?: string | undefined;
  parentMessageId?: string | undefined;
  ringId?: string | undefined;
  appSessionId?: string | undefined;
  isCallingAllowed?: boolean | undefined;
  isPSTNCallingAllowed?: boolean | undefined;
  meetingId?: string | undefined;
  defaultOneNoteSectionId?: string | undefined;
  isMultiWindow?: boolean | undefined;
  appIconPosition?: number | undefined;
  sourceOrigin?: string | undefined;
  userClickTime?: number | undefined;
  teamTemplateId?: string | undefined;
  userFileOpenPreference?: microsoftTeams.FileOpenPreference | undefined;
}
