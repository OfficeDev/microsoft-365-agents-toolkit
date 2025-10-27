Microsoft Visual Studio Solution File, Format Version 12.00
# Visual Studio Version 17
VisualStudioVersion = 17.14.36616.10 d17.14
MinimumVisualStudioVersion = 10.0.40219.1
Project("{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}") = "{{ProjectName}}", "{{ProjectName}}\{{ProjectName}}.csproj", "{{ProjectGuid}}"
EndProject
Project("{A9E3F50B-275E-4AF7-ADCE-8BE12D41E305}") = "{{NewProjectTypeName}}", "{{NewProjectTypeName}}\{{NewProjectTypeName}}.atkproj", "{{NewProjectTypeProjectGuid}}"
EndProject
Global
	GlobalSection(SolutionConfigurationPlatforms) = preSolution
		Debug|Any CPU = Debug|Any CPU
		Release|Any CPU = Release|Any CPU
	EndGlobalSection
	GlobalSection(ProjectConfigurationPlatforms) = postSolution
		{{ProjectGuid}}.Debug|Any CPU.ActiveCfg = Debug|Any CPU
		{{ProjectGuid}}.Debug|Any CPU.Build.0 = Debug|Any CPU
		{{ProjectGuid}}.Release|Any CPU.ActiveCfg = Release|Any CPU
		{{ProjectGuid}}.Release|Any CPU.Build.0 = Release|Any CPU
		{{NewProjectTypeProjectGuid}}.Debug|Any CPU.ActiveCfg = Debug|Any CPU
		{{NewProjectTypeProjectGuid}}.Debug|Any CPU.Build.0 = Debug|Any CPU
		{{NewProjectTypeProjectGuid}}.Debug|Any CPU.Deploy.0 = Debug|Any CPU
		{{NewProjectTypeProjectGuid}}.Release|Any CPU.ActiveCfg = Release|Any CPU
		{{NewProjectTypeProjectGuid}}.Release|Any CPU.Build.0 = Release|Any CPU
		{{NewProjectTypeProjectGuid}}.Release|Any CPU.Deploy.0 = Release|Any CPU
	EndGlobalSection
	GlobalSection(SolutionProperties) = preSolution
		HideSolutionNode = FALSE
	EndGlobalSection
	GlobalSection(ExtensibilityGlobals) = postSolution
		SolutionGuid = {{SolutionGuid}}
	EndGlobalSection
EndGlobal