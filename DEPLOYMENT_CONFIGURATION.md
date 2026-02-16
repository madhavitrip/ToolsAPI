# Deployment Configuration Guide

## Overview
This document explains how the application is configured to handle environment-specific settings and exclude certain files from publication.

## Changes Made

### 1. **Excluded Files from Publish**
The following files are now excluded from publication and should be managed separately on the production server:

- `wwwroot/` - Static files folder (generated at runtime)
- `appsettings.json` - Development configuration
- `appsettings.Development.json` - Development-specific configuration
- `web.config` - IIS configuration (if applicable)

**Configuration Location:** `Tools.csproj`

```xml
<ItemGroup>
    <Content Remove="wwwroot\**" />
</ItemGroup>

<ItemGroup>
    <None Remove="appsettings.json" />
    <None Remove="appsettings.Development.json" />
    <None Remove="web.config" />
</ItemGroup>
```

### 2. **API URL Configuration**
Hardcoded API URLs have been removed and moved to configuration files.

**Configuration Location:** `appsettings.json` and `appsettings.Production.json`

```json
"ApiSettings": {
    "BaseUrl": "http://192.168.10.208/ERPAPI/api",
    "EnvelopeBreakageUrl": "http://192.168.10.208:81/API/api/EnvelopeBreakages/EnvelopeBreakage"
}
```

### 3. **Dependency Injection Setup**
The `ApiSettings` class is registered in `Program.cs`:

```csharp
builder.Services.Configure<ApiSettings>(builder.Configuration.GetSection("ApiSettings"));
```

### 4. **Controller Implementation**
The `EnvelopeBreakagesController` now injects `ApiSettings`:

```csharp
private readonly ApiSettings _apiSettings;

public EnvelopeBreakagesController(
    ERPToolsDbContext context, 
    ILoggerService loggerService, 
    IOptions<ApiSettings> apiSettings)
{
    _context = context;
    _loggerService = loggerService;
    _apiSettings = apiSettings.Value;
}
```

Usage in code:
```csharp
var response = await client.GetAsync($"{_apiSettings.EnvelopeBreakageUrl}?ProjectId={ProjectId}");
```

## Deployment Instructions

### For Development
1. Use the default `appsettings.json` file
2. Run the application normally

### For Production

1. **Create Production Configuration:**
   - Copy `appsettings.Production.json` to your production server
   - Update the values with your production environment details:
     - Database connection string
     - API URLs
     - JWT settings

2. **Publish the Application:**
   ```bash
   dotnet publish -c Release
   ```

3. **Deploy to Production Server:**
   - Copy the published files to your production server
   - Manually place `appsettings.Production.json` on the server (it won't be included in the publish)
   - Ensure the `wwwroot` folder exists and has proper permissions

4. **Run the Application:**
   ```bash
   dotnet Tools.dll --environment Production
   ```

## Environment-Specific Configuration

The application automatically loads the appropriate `appsettings.{Environment}.json` file based on the `ASPNETCORE_ENVIRONMENT` variable:

- **Development:** `appsettings.Development.json`
- **Production:** `appsettings.Production.json`
- **Staging:** `appsettings.Staging.json` (if created)

## Adding New API URLs

To add new API URLs:

1. Add the URL to `appsettings.json` and `appsettings.Production.json`:
   ```json
   "ApiSettings": {
       "NewApiUrl": "http://your-api-url/endpoint"
   }
   ```

2. Update the `ApiSettings` class in `Models/ApiSettings.cs`:
   ```csharp
   public string NewApiUrl { get; set; }
   ```

3. Use in your controller:
   ```csharp
   var response = await client.GetAsync($"{_apiSettings.NewApiUrl}");
   ```

## Benefits

✅ **No Hardcoded URLs** - All URLs are configurable
✅ **Environment-Specific Settings** - Different configs for dev/prod
✅ **Clean Deployments** - wwwroot and config files not overwritten
✅ **Easy Maintenance** - Change URLs without recompiling
✅ **Security** - Sensitive data not in source control
