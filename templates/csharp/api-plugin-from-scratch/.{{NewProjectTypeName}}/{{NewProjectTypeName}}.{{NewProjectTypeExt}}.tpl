<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" Sdk="Microsoft.TeamsFx.Sdk">
  <ItemGroup>
    <ProjectCapability Include="ProjectConfigurationsDeclaredDimensions" />
    {{#DeclarativeCopilot}}
      <ProjectCapability Include="DeclarativeAgent" />
    {{/DeclarativeCopilot}}
    {{^DeclarativeCopilot}}
      <ProjectCapability Include="ApiPlugin" />
    {{/DeclarativeCopilot}}
  </ItemGroup>
</Project>
