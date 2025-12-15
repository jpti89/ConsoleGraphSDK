using ConsoleApp2.config;

internal class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("Graph SDK Console App\n");

        var settings = Settings.LoadSettings();
        var graphService = new GraphService(settings);

        var app = new App(graphService);
        await app.RunAsync();
    }
}
