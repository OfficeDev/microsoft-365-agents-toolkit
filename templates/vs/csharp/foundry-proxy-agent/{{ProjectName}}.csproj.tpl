<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>{{TargetFramework}}</TargetFramework>
    <LangVersion>latest</LangVersion>
    <ImplicitUsings>enable</ImplicitUsings>
    <NoWarn>$(NoWarn);SKEXP0110;SKEXP0010</NoWarn>
    <UserSecretsId>{{userSecretsId}}</UserSecretsId>
  </PropertyGroup>

  <ItemGroup>
    <Content Remove="appsettings.json" />
  </ItemGroup>

  <ItemGroup>
    <None Include="appsettings.json">
      <CopyToOutputDirectory>Never</CopyToOutputDirectory>
    </None>
  </ItemGroup>

<ItemGroup>
    <PackageReference Include="Azure.Identity" Version="1.17.1" />
    <PackageReference Include="Microsoft.Agents.Authentication.Msal" Version="1.3.176" />
    <PackageReference Include="Microsoft.Agents.Hosting.AspNetCore" Version="1.3.176" />
    <PackageReference Include="Microsoft.Agents.AI" Version="1.0.0-preview.260108.1" />
    <PackageReference Include="Microsoft.Agents.AI.AzureAI" Version="1.0.0-preview.260108.1" />
    <PackageReference Include="Microsoft.Agents.AI.AzureAI.Persistent" Version="1.0.0-preview.260108.1" />
    <PackageReference Include="Azure.AI.Agents.Persistent" Version="1.2.0-beta.8" />
    <PackageReference Include="Microsoft.Extensions.AI" Version="10.2.0" />
  </ItemGroup>


</Project>
