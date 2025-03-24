using {{SafeProjectName}};
using {{SafeProjectName}}.Action;
using Microsoft.Agents.Hosting.AspNetCore;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddHttpClient("WebClient", client => client.Timeout = TimeSpan.FromSeconds(600));
builder.Services.AddHttpContextAccessor();

builder.Services.AddCloudAdapter<AdapterWithErrorHandler>();
builder.AddBot<ActionApp>();

var app = builder.Build();

app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();

if (app.Environment.IsDevelopment())
{
    app.MapGet("/", () => "Message Extension Action");
    app.UseDeveloperExceptionPage();
    app.MapControllers().AllowAnonymous();
}
else
{
    app.MapControllers();
}

app.Run();