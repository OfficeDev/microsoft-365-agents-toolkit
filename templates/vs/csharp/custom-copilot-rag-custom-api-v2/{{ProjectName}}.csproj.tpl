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
    <None Include="env/**/*" />
  </ItemGroup>

{{/isNewProjectTypeEnabled}}
  <ItemGroup>
    <PackageReference Include="AdaptiveCards" Version="3.1.0" />
    <PackageReference Include="AdaptiveCards.Templating" Version="1.5.0" />
    <PackageReference Include="Azure.Identity" Version="1.13.1" />
    <PackageReference Include="Microsoft.Bot.Builder" Version="4.22.9" />
    <PackageReference Include="Microsoft.Bot.Builder.Integration.AspNet.Core" Version="4.22.9" />
    <PackageReference Include="Microsoft.OpenApi" Version="1.6.19" />
    <PackageReference Include="Microsoft.OpenApi.Readers" Version="1.6.19" />
    <PackageReference Include="Microsoft.Teams.AI" Version="1.8.*" />
    <PackageReference Include="RestSharp" Version="112.0.0" />
    <PackageReference Include="System.Text.Json" Version="8.0.5" />
  </ItemGroup>

  <ItemGroup>
    <Content Include="Prompts\chat\skprompt.txt">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
      <CopyToPublishDirectory>PreserveNewest</CopyToPublishDirectory>
    </Content>

    <Content Remove="apiSpecificationFile\{{OPENAPI_SPEC_PATH}}" />
    <Content Include="apiSpecificationFile\{{OPENAPI_SPEC_PATH}}">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
      <CopyToPublishDirectory>PreserveNewest</CopyToPublishDirectory>
    </Content>

    <!-- Exclude local settings from publish -->
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
