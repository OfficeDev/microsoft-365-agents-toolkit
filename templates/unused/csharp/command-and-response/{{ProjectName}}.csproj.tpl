<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>{{TargetFramework}}</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="AdaptiveCards.Templating" Version="2.0.5" />
    <PackageReference Include="Microsoft.Agents.Authentication.Msal" Version="1.*" />
    <PackageReference Include="Microsoft.Agents.Hosting.AspNetCore" Version="1.*" />
    <PackageReference Include="Microsoft.AspNetCore.Components" Version="8.0.14" />
  </ItemGroup>

</Project>
