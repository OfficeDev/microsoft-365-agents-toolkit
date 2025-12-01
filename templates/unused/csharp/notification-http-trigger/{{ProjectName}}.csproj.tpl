<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>{{TargetFramework}}</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <AzureFunctionsVersion>v4</AzureFunctionsVersion>
    <DisableFastUpToDateCheck>true</DisableFastUpToDateCheck>
    <OutputType>Exe</OutputType>
    <RootNamespace>{{SafeProjectName}}</RootNamespace>
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
    <PackageReference Include="Microsoft.Azure.Functions.Worker" Version="2.50.0-preview1" />
    <PackageReference Include="Microsoft.Azure.Functions.Worker.Sdk" Version="2.0.5" />
    <PackageReference Include="Microsoft.Azure.Functions.Worker.Extensions.Http.AspNetCore" Version="2.0.0" />
    <PackageReference Include="Microsoft.Agents.Authentication.Msal" Version="1.*" />
    <PackageReference Include="Microsoft.Agents.Hosting.AspNetCore" Version="1.*" />
    <PackageReference Include="Microsoft.Agents.Extensions.Teams" Version="1.*" />
    <!-- Fix system package downgrade -->
    <PackageReference Include="Microsoft.NETCore.Targets" Version="3.1.0" PrivateAssets="All" />
  </ItemGroup>

  <ItemGroup>
    <Content Include="Resources/*.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </Content>
  </ItemGroup>

  <ItemGroup>
    <None Update="host.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
    <None Update="local.settings.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
      <CopyToPublishDirectory>Never</CopyToPublishDirectory>
    </None>
    <None Update="appsettings.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
    <None Update="appsettings.Development.json">
      <DependentUpon>appsettings.json</DependentUpon>
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
      <CopyToPublishDirectory>Never</CopyToPublishDirectory>
    </None>
    <None Update="appsettings.Playground.json">
      <DependentUpon>appsettings.json</DependentUpon>
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
      <CopyToPublishDirectory>Never</CopyToPublishDirectory>
    </None>
  </ItemGroup>

</Project>
