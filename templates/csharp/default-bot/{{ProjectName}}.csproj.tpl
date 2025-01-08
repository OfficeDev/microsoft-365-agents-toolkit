<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>{{TargetFramework}}</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>

{{^isNewProjectTypeEnabled}}
  <ItemGroup>
    <ProjectCapability Include="TeamsFx" />
  </ItemGroup>

  <ItemGroup>
    <None Include="appPackage/**/*" />
    <None Include="infra/**/*" />
    <None Remove="devTools/**" />
    <Content Remove="devTools/**/*" />
  </ItemGroup>

{{/isNewProjectTypeEnabled}}
  <ItemGroup>
    <PackageReference Include="Microsoft.Agents.Authentication" Version="0.2.3-alpha" />
    <PackageReference Include="Microsoft.Agents.Authentication.Msal" Version="0.2.3-alpha" />
    <PackageReference Include="Microsoft.Agents.Hosting.AspNetCore" Version="0.2.3-alpha" />
    <PackageReference Include="Microsoft.Agents.Protocols" Version="0.2.3-alpha" />
  </ItemGroup>

</Project>
