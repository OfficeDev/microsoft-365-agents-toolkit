<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>{{TargetFramework}}</TargetFramework>
		<Nullable>enable</Nullable>
		<ImplicitUsings>enable</ImplicitUsings>
		<GenerateEmbeddedFilesManifest>true</GenerateEmbeddedFilesManifest>
		<DisableFastUpToDateCheck>true</DisableFastUpToDateCheck>
	</PropertyGroup>

	  <ItemGroup>
    <Content Remove="Web\**" />
  </ItemGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Extensions.FileProviders.Embedded" Version="9.0.0" />
    <PackageReference Include="Azure.Identity" Version="1.13.1" />
    <PackageReference Include="Microsoft.Teams.Api" Version="2.0.0" />
    <PackageReference Include="Microsoft.Teams.Apps" Version="2.0.0" />
    <PackageReference Include="Microsoft.Teams.Plugins.AspNetCore" Version="2.0.0" />
    <PackageReference Include="Microsoft.Teams.Common" Version="2.0.0" />
  </ItemGroup>

  <!-- Exclude local settings from publish -->
  <ItemGroup>
    <Content Include="Web\package.json">
      <Visible>true</Visible>
      <CopyToOutputDirectory>Never</CopyToOutputDirectory>
    </Content>
    <Content Include="Web\tsconfig.json">
      <Visible>true</Visible>
      <CopyToOutputDirectory>Never</CopyToOutputDirectory>
    </Content>
  </ItemGroup>
  
  <ItemGroup>
    <Folder Include="Web\bin\" />
  </ItemGroup>

  <!-- Run npm steps -->
  <Target Name="NpmInstall" BeforeTargets="BeforeBuild">
    <Message Text="Running npm install..." Importance="high" />
    <Exec Command="npm install" WorkingDirectory="Web" />
  </Target>

  <Target Name="NpmBuild" AfterTargets="NpmInstall" BeforeTargets="BeforeBuild">
    <Message Text="Running npm run build..." Importance="high" />
    <Exec Command="npm run build" WorkingDirectory="Web" />
  </Target>

  <!-- Dynamically include built frontend files as embedded resources after npm build -->
  <Target Name="IncludeFrontendAssets" AfterTargets="NpmBuild" BeforeTargets="BeforeBuild">
    <ItemGroup>
      <EmbeddedResource Include="Web\bin\**" />
    </ItemGroup>
    <Message Text="Including %(EmbeddedResource.Identity) as embedded resource" Importance="high" />
  </Target>

</Project>
