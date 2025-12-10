<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>{{TargetFramework}}</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>

  <ItemGroup>
	  <PackageReference Include="AdaptiveCards" Version="3.1.0" />
	  <PackageReference Include="Microsoft.Agents.Authentication.Msal" Version="1.*" />
	  <PackageReference Include="Microsoft.Agents.Extensions.Teams" Version="1.*" />
	  <PackageReference Include="Microsoft.Agents.Hosting.AspNetCore" Version="1.*" />
  </ItemGroup>

</Project>
