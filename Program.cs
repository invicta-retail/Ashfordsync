using AshfordSync.Interfaces;
using AshfordSync.Service;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog;
using System.Threading.Tasks;

namespace AshfordSync
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Log.Logger = new LoggerConfiguration()
                   .WriteTo.File("Ashford.log", rollingInterval: RollingInterval.Day)
                   .CreateLogger();

            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);

            var serviceProvider = serviceCollection.BuildServiceProvider();

            await serviceProvider.GetService<MainProcess>().StartService();

            var logger = serviceProvider.GetService<ILogger<Program>>();

            logger.LogInformation("Completed process");
        }

        private static void ConfigureServices(IServiceCollection services)
        {
            services.AddLogging(configure => configure.AddSerilog())
                .AddTransient<MainProcess>()
                .AddTransient<IExcelService, ExcelService>()
                .AddTransient<IReadInventoryService,ReadInventoryService>();
        }
    }
}
