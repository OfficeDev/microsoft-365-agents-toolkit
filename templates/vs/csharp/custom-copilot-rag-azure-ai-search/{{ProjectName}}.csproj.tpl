<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>{{TargetFramework}}</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Azure.Identity" Version="1.13.1" />
  	<PackageReference Include="Azure.Search.Documents" Version="11.6.0" />
    <PackageReference Include="Microsoft.Teams.Api" Version="2.0.*" />
    <PackageReference Include="Microsoft.Teams.Apps" Version="2.0.*" />
    <PackageReference Include="Microsoft.Teams.Plugins.AspNetCore" Version="2.0.*" />
    <PackageReference Include="Microsoft.Teams.Common" Version="2.0.*" />
  </ItemGroup>

  <ItemGroup>
    <Content Include="Prompts\instructions.txt">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
      <CopyToPublishDirectory>PreserveNewest</CopyToPublishDirectory>
    </Content>
  </ItemGroup>

    <!-- Exclude local settings from publish -->
  <ItemGroup>
    <Content Remove="appsettings.Development.json" />
    <Content Include="appsettings.Development.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
      <CopyToPublishDirectory>None</CopyToPublishDirectory>
    </Content>
    <Content Remove="appsettings.Playground.json" />
    <Content Include="appsettings.Playground.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
      <CopyToPublishDirectory>None</CopyToPublishDirectory>
    </Content>
  </ItemGroup>
</Project>
