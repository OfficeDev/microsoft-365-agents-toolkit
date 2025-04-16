[
{{#enableTestToolByDefault}}
  {
    "Name": "Microsoft 365 Agents Playground (browser)",
    "Projects": [
      {
        "Path": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Name": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Action": "StartWithoutDebugging",
        "DebugTarget": "Microsoft 365 Agents Playground (browser)"
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
        "DebugTarget": "Microsoft 365 Agents Playground"
      }
    ]
  },
{{/enableTestToolByDefault}}
  {
    "Name": "Microsoft Teams (browser)",
    "Projects": [
      {
        "Path": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Name": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Action": "StartWithoutDebugging",
        "DebugTarget": "Microsoft Teams (browser)"
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
    "Name": "Microsoft Teams (browser) (skip update app)",
    "Projects": [
      {
        "Path": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Name": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Action": "StartWithoutDebugging",
        "DebugTarget": "Microsoft Teams (browser) (skip update app)"
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
    "Name": "Outlook (browser)",
    "Projects": [
      {
        "Path": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Name": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Action": "StartWithoutDebugging",
        "DebugTarget": "Outlook (browser)"
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
    "Name": "Outlook (browser) (skip update app)",
    "Projects": [
      {
        "Path": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Name": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Action": "StartWithoutDebugging",
        "DebugTarget": "Outlook (browser) (skip update app)"
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
{{#enableTestToolByDefault}}
  }
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
  },
  {
    "Name": "Microsoft 365 Agents Playground (browser)",
    "Projects": [
      {
        "Path": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Name": "{{NewProjectTypeName}}\\{{NewProjectTypeName}}.{{NewProjectTypeExt}}",
        "Action": "StartWithoutDebugging",
        "DebugTarget": "Microsoft 365 Agents Playground (browser)"
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
        "DebugTarget": "Microsoft 365 Agents Playground"
      }
    ]
  }
{{/enableTestToolByDefault}}
]