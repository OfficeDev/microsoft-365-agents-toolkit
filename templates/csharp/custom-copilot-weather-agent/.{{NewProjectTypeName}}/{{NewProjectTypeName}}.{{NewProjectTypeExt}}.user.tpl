<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="Current" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup Condition="'$(Configuration)|$(Platform)'=='Debug|AnyCPU'">
    <DebuggerFlavor>ProjectDebugger</DebuggerFlavor>
  </PropertyGroup>
  <PropertyGroup>
{{#enableTestToolByDefault}}
    <ActiveDebugProfile>Teams App Test Tool (browser)</ActiveDebugProfile>
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
    <ActiveDebugProfile>Microsoft Teams (browser)</ActiveDebugProfile>
{{/enableTestToolByDefault}}
  </PropertyGroup>
</Project>