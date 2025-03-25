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
  </ItemGroup>

{{/isNewProjectTypeEnabled}}
  <ItemGroup>
    <PackageReference Include="Microsoft.AspNetCore.Authentication.JwtBearer" Version="6.0.33" />
    <PackageReference Include="Microsoft.Bot.Builder" Version="4.22.7" />
    <PackageReference Include="Microsoft.Bot.Builder.Integration.AspNet.Core" Version="4.22.7" />
    <PackageReference Include="AdaptiveCards" Version="2.7.3" />
    <PackageReference Include="AdaptiveCards.Templating" Version="1.4.0" />
    <PackageReference Include="System.Text.RegularExpressions" Version="4.3.1" />
  </ItemGroup>

</Project>
