<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>{{TargetFramework}}</TargetFramework>
    <LangVersion>latest</LangVersion>
    <ImplicitUsings>enable</ImplicitUsings>
    <NoWarn>$(NoWarn);SKEXP0110;SKEXP0010</NoWarn>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="AdaptiveCards" Version="3.1.0" />
    <PackageReference Include="Microsoft.SemanticKernel.Agents.AzureAI" Version="1.73.0-preview" />
    <PackageReference Include="Microsoft.SemanticKernel.Agents.Core" Version="1.73.0" />
    <PackageReference Include="Microsoft.SemanticKernel.Connectors.AzureOpenAI" Version="1.73.0" />
    <PackageReference Include="Microsoft.SemanticKernel.Connectors.OpenAI" Version="1.73.0" />
    <PackageReference Include="Microsoft.Agents.Authentication.Msal" Version="1.*" />
    <PackageReference Include="Microsoft.Agents.Hosting.AspNetCore" Version="1.*" />
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
  <ItemGroup>
    <None Include="appsettings.Development.json" />
  </ItemGroup>
</Project>
