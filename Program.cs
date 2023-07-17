using Microsoft.Extensions.Hosting;
using SerialPortWorker;
using Serilog;
using Serilog.Events;


Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Debug()
    .MinimumLevel.Override("Microsoft", LogEventLevel.Warning)
    .Enrich.FromLogContext()
    //.WriteTo
    .WriteTo.File(@"C:\Temp\LogFile.txt")
    .CreateLogger();

try
{
    Log.Information("Starting service");

    Host.CreateDefaultBuilder(args)
        .UseWindowsService()
        .ConfigureServices((hostContext, services) =>
        {
            services.AddHostedService<Worker>();
            services.AddHttpClient();
        })
        .UseSerilog()
        .Build()
        .Run();

    return;
}
catch (Exception ex)
{
    Log.Fatal(ex, "There was a problem starting the service");
    return;
}
finally
{
    // Flush all buffered messages to the file before close or crash
    Log.CloseAndFlush();
}
