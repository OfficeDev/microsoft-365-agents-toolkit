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
    <PackageReference Include="AdaptiveCards.Templating" Version="1.3.1" />
    <PackageReference Include="Microsoft.Agents.BotBuilder.App" Version="1.0.0" />
    <PackageReference Include="Microsoft.Agents.BotBuilder.State" Version="1.0.0" />
    <PackageReference Include="Microsoft.Agents.Storage" Version="1.0.0" />
    <PackageReference Include="Microsoft.Agents.Hosting.AspNetCore" Version="1.0.0" />
    <PackageReference Include="Microsoft.AspNetCore.Components" Version="6.0.33" />
    <PackageReference Include="Microsoft.IdentityModel.Protocols.OpenIdConnect" Version="6.36.0" />
    <PackageReference Include="Microsoft.TeamsFx" Version="3.0.0-rc">
      <!-- Exclude TeamsFx wwwroot static files which are for frontend only. -->
      <ExcludeAssets>contentFiles</ExcludeAssets>
    </PackageReference>
  </ItemGroup>

</Project>
