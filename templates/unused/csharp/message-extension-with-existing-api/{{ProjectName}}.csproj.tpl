{{^isNewProjectTypeEnabled}}
<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>

  <ItemGroup>
    <ProjectCapability Include="TeamsFx" />
    <ProjectCapability Include="APIME" />
  </ItemGroup>

  <ItemGroup>
    <None Include="appPackage/**/*" />
  </ItemGroup>

</Project>
{{/isNewProjectTypeEnabled}}
{{#isNewProjectTypeEnabled}}
<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" Sdk="Microsoft.TeamsFx.Sdk">
  <ItemGroup>
    <ProjectCapability Include="APIME" />
  </ItemGroup>
  <ItemGroup>
    <ProjectCapability Include="ProjectConfigurationsDeclaredDimensions" />
  </ItemGroup>
</Project>
{{/isNewProjectTypeEnabled}}
