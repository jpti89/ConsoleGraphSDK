using Microsoft.Extensions.Configuration;

namespace ConsoleApp2.config;

public class Settings
{
    public string? ClientId { get; set; }

    public string? ClientSecret { get; set; }

    public string? TenantId { get; set; }

    public static Settings LoadSettings()
    {
        IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("./appsettings.json", optional: false)
            .AddJsonFile($"appsettings.Development.json", optional: true)
            .AddUserSecrets<Program>()
            .Build();

        return config.GetRequiredSection("Settings").Get<Settings>() ??
            throw new Exception("Could not load app settings. See README for configuration instructions.");
    }
}