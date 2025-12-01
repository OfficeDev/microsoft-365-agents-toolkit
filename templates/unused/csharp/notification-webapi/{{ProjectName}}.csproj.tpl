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
    <None Include=".notification.local*.json" />
    <None Include=".notification.playground*.json" />
  </ItemGroup>

  <ItemGroup>
    <PackageReference Include="AdaptiveCards.Templating" Version="2.0.5" />
    <PackageReference Include="Newtonsoft.Json" Version="13.0.3" />
    <PackageReference Include="Microsoft.Agents.Authentication.Msal" Version="1.*" />
    <PackageReference Include="Microsoft.Agents.Hosting.AspNetCore" Version="1.*" />
    <PackageReference Include="Microsoft.Agents.Extensions.Teams" Version="1.*" />
  </ItemGroup>

</Project>
