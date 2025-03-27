using {{SafeProjectName}};
using {{SafeProjectName}}.LinkUnfurling;
using Microsoft.Agents.Hosting.AspNetCore;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddHttpClient("WebClient", client => client.Timeout = TimeSpan.FromSeconds(600));
builder.Services.AddHttpContextAccessor();

builder.Services.AddCloudAdapter<AdapterWithErrorHandler>();
builder.AddBot<LinkUnfurlingApp>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

if (app.Environment.IsDevelopment() || app.Environment.EnvironmentName == "TestTool")
{
    app.MapGet("/", () => "Link Unfurling Bot");
    app.UseDeveloperExceptionPage();
    app.MapControllers().AllowAnonymous();
}
else
{
    app.MapControllers();
}

app.Run();