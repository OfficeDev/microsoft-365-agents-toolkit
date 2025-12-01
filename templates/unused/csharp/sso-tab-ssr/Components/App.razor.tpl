<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>{{ProjectName}}</title>
    <base href="/" />
    <link href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css" rel="stylesheet"/>
    <link href="css/site.css" rel="stylesheet" />
    <link href="{{ProjectName}}.styles.css" rel="stylesheet" />
</head>
<body>
    <HeadOutlet />
    <Routes @rendermode="@RenderMode.InteractiveServer" />
    <script src="_framework/blazor.web.js"></script>
    <script type="module" src="https://unpkg.com/@@fluentui/web-components"></script>
</body>
</html>
