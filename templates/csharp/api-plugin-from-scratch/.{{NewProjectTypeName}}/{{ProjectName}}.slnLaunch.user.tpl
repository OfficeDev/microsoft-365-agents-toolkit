[
  {
    "Name": "Copilot (browser)",
    "Projects": [
      {
        "Path": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Name": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Action": "StartWithoutDebugging",
        "DebugTarget": "Copilot (browser)"
      },
      {
{{#PlaceProjectFileInSolutionDir}}
        "Path": "{{ProjectName}}.csproj",
        "Name": "{{ProjectName}}.csproj",
{{/PlaceProjectFileInSolutionDir}}
{{^PlaceProjectFileInSolutionDir}}
        "Path": "{{ProjectName}}\\{{ProjectName}}.csproj",
        "Name": "{{ProjectName}}\\{{ProjectName}}.csproj",
{{/PlaceProjectFileInSolutionDir}}
        "Action": "Start",
        "DebugTarget": "Start Project"
      }
    ]
  },
  {
    "Name": "Copilot (browser) (skip update Teams App)",
    "Projects": [
      {
        "Path": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Name": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Action": "StartWithoutDebugging",
        "DebugTarget": "Copilot (browser) (skip update Teams App)"
      },
      {
{{#PlaceProjectFileInSolutionDir}}
        "Path": "{{ProjectName}}.csproj",
        "Name": "{{ProjectName}}.csproj",
{{/PlaceProjectFileInSolutionDir}}
{{^PlaceProjectFileInSolutionDir}}
        "Path": "{{ProjectName}}\\{{ProjectName}}.csproj",
        "Name": "{{ProjectName}}\\{{ProjectName}}.csproj",
{{/PlaceProjectFileInSolutionDir}}
        "Action": "Start",
        "DebugTarget": "Start Project"
      }
    ]
  }
]