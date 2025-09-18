using Microsoft.Teams.Extensions.Logging;
using Microsoft.Teams.Plugins.AspNetCore.Extensions;

var builder = WebApplication.CreateBuilder(args);

builder.AddTeams();

var app = builder.Build();

app.UseTeams();
app.AddTab("test", "Web/bin");

app.Run();