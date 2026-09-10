
using System.Text;
using System.Security.Claims;
using System.IdentityModel.Tokens.Jwt;
using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using OfficeOpenXml;
using System.Reflection;
using Tools.Services;
using Tools.Middleware;
using Microsoft.Extensions.Options;
using Tools.Models;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using ToolsAPI.Models;
using Microsoft.Extensions.FileProviders;

var builder = WebApplication.CreateBuilder(args);
FileStorageHelper.Initialize(builder.Configuration);

builder.Services.AddHttpContextAccessor();
builder.Services.AddDbContext<ERPToolsDbContext>(options =>
    options.UseMySql(builder.Configuration.GetConnectionString("ERPToolsDb"),
        ServerVersion.AutoDetect(builder.Configuration.GetConnectionString("ERPToolsDb"))));
builder.Services.AddScoped<ILoggerService, LoggerService>();
builder.Services.AddScoped<ApiAuditLoggingFilter>();
builder.Services.AddScoped<IAuthorizationService, AuthorizationService>();
builder.Services.AddScoped<IMasterAuthService, MasterAuthService>();

// Register API settings from configuration
builder.Services.Configure<ApiSettings>(builder.Configuration.GetSection("ApiSettings"));
builder.Services.Configure<MasterAuthSettings>(builder.Configuration.GetSection("MasterAuthSettings"));

// Add JWT Auth



builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidateAudience = true,
            ValidateLifetime = true,
            ValidateIssuerSigningKey = true,
            ValidIssuer = builder.Configuration["JwtSettings:Issuer"],
            ValidAudience =  builder.Configuration["JwtSettings:Audience"],
            IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(builder.Configuration["JwtSettings:SecretKey"])),
            NameClaimType = "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name"
        };
        
        // Map custom claims from JWT to ClaimsPrincipal
        options.Events = new JwtBearerEvents
        {
            OnTokenValidated = context =>
            {
                // Extract custom claims and add them to the ClaimsPrincipal
                var token = context.SecurityToken as JwtSecurityToken;
                if (token != null)
                {
                    var identity = context.Principal.Identity as ClaimsIdentity;
                    if (identity != null)
                    {
                        // Add userid claim if not already present
                        if (identity.FindFirst("userid") == null)
                        {
                            var userIdClaim = token.Claims.FirstOrDefault(c => c.Type == "userid");
                            if (userIdClaim != null)
                            {
                                identity.AddClaim(new Claim("userid", userIdClaim.Value));
                            }
                        }
                        
                        // Add roleId claim if not already present
                        if (identity.FindFirst("roleId") == null)
                        {
                            var roleIdClaim = token.Claims.FirstOrDefault(c => c.Type == "roleId");
                            if (roleIdClaim != null)
                            {
                                identity.AddClaim(new Claim("roleId", roleIdClaim.Value));
                            }
                        }
                    }
                }
                return Task.CompletedTask;
            }
        };
    });

// Add services to the container.
builder.Services.AddHttpClient();

// Add custom HttpClient that bypasses SSL verification for development (RPT Service)
builder.Services.AddHttpClient("RptService")
    .ConfigurePrimaryHttpMessageHandler(() =>
    {
        var handler = new HttpClientHandler();
        // Bypass SSL certificate validation for development
        handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true;
        return handler;
    });

builder.Services.AddHttpClient<IDispatchService, DispatchService>();
builder.Services.AddControllers(options =>
{
    options.Filters.Add<ApiAuditLoggingFilter>();
});
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 100 * 1024 * 1024; // 100 MB
});
builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = 100 * 1024 * 1024; // 100 MB
});
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
ExcelPackage.License.SetNonCommercialPersonal("Your Name");
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policyBuilder =>
    {
        policyBuilder.AllowAnyOrigin()
                      .AllowAnyMethod()
                      .AllowAnyHeader();
    });
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseAuthentication();
app.UseHttpsRedirection();
app.UseCors();

// 1. Default wwwroot static files
app.UseStaticFiles();

// 2. Custom storage path (e.g. E:\Files) static files
var storageBasePath = FileStorageHelper.GetStorageBasePath(builder.Configuration);
if (Directory.Exists(storageBasePath))
{
    var fileProvider = new PhysicalFileProvider(storageBasePath);

    // Serve at root URL so http://host:port/{projectId}/{fileName} continues to work seamlessly
    app.UseStaticFiles(new StaticFileOptions
    {
        FileProvider = fileProvider,
        RequestPath = ""
    });

    // Also serve at /Files URL so http://host:port/Files/{projectId}/{fileName} works as well
    app.UseStaticFiles(new StaticFileOptions
    {
        FileProvider = fileProvider,
        RequestPath = "/Files"
    });
}

app.UseAuthorization();

app.MapControllers();

app.Run();
