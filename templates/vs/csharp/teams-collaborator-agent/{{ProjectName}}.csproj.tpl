<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>{{TargetFramework}}</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Azure.Identity" Version="1.17.1" />
    <PackageReference Include="JsonSchema.Net" Version="7.3.4" />
    <PackageReference Include="Microsoft.Data.SqlClient" Version="6.1.2" />
    <PackageReference Include="Microsoft.Data.Sqlite" Version="9.0.10" />
    <PackageReference Include="Microsoft.Recognizers.Text.DateTime" Version="1.8.13" />
    <!-- Fix vulnerability in transitive dependency from Microsoft.Recognizers.Text -->
    <PackageReference Include="NuGet.CommandLine" Version="5.11.7" />
    <PackageReference Include="Microsoft.Teams.AI" Version="2.0.*" />
    <PackageReference Include="Microsoft.Teams.Api" Version="2.0.*" />
    <PackageReference Include="Microsoft.Teams.Apps" Version="2.0.*" />
    <PackageReference Include="Microsoft.Teams.Plugins.AspNetCore" Version="2.0.*" />
    <PackageReference Include="Microsoft.Teams.Common" Version="2.0.*" />
    <PackageReference Include="Microsoft.Teams.Extensions.Logging" Version="2.0.*" />
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
