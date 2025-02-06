<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" Sdk="Microsoft.TeamsFx.Sdk">
  <ItemGroup>
    <None Include="GenerateApiKey.ps1" />
  </ItemGroup>
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
