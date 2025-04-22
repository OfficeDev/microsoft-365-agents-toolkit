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
    <PackageReference Include="AdaptiveCards.Templating" Version="2.0.5" />
    <PackageReference Include="Microsoft.Agents.Authentication.Msal" Version="0.2.*-*" />
    <PackageReference Include="Microsoft.Agents.Hosting.AspNetCore" Version="0.2.*-*" />
    <PackageReference Include="Microsoft.AspNetCore.Components" Version="8.0.14" />
    <PackageReference Include="Microsoft.TeamsFx" Version="3.0.*-*" >
      <!-- Exclude TeamsFx wwwroot static files which are for frontend only. -->
      <ExcludeAssets>contentFiles</ExcludeAssets>
    </PackageReference>
  </ItemGroup>

</Project>
