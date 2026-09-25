using Microsoft.EntityFrameworkCore;
using OwaspDemo.Data;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();

var sqlServerConn = builder.Configuration.GetConnectionString("SqlServer") ?? "";
var sqliteConn = builder.Configuration.GetConnectionString("Sqlite") ?? "Data Source=owasp_demo.db";

bool useSqlServer = false;
try
{
    using var testConn = new Microsoft.Data.SqlClient.SqlConnection(sqlServerConn);
    testConn.Open();
    useSqlServer = true;
}
catch
{
    useSqlServer = false;
}

builder.Services.AddDbContext<AppDbContext>(options =>
{
    if (useSqlServer)
    {
        options.UseSqlServer(sqlServerConn);
    }
    else
    {
        options.UseSqlite(sqliteConn);
    }
});

var app = builder.Build();

// Ensure DB is created with seed data
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();
}

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseRouting();
app.UseAuthorization();
app.MapStaticAssets();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=SqlInjection}/{action=Index}/{id?}")
    .WithStaticAssets();

app.Run();
