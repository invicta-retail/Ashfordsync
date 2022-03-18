using AshfordSync.Entities;
using AshfordSync.Interfaces;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace AshfordSync
{
    class MainProcess
    {
        private readonly ILogger<MainProcess> _logger;
        private readonly IExcelService _service;
        private static Mutex mutex = null;

        public MainProcess(ILogger<MainProcess> logger, IExcelService service)
        {
            _logger = logger;
            _service = service;
        }

        public async Task StartService()
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };

            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = System.Text.Json.JsonSerializer.Deserialize<Parameters>(jsonParameters, options);

            string appName = "SupplierEmailIntegrationWorker";
            bool createdNew;

            mutex = new Mutex(true, appName, out createdNew);

            if (!createdNew)
            {
                Console.WriteLine(appName + " is already running! Exiting the application.");
                _logger.LogInformation(appName + " is already running! Exiting the application.");
                return;
            }

            #region Files

            if (!System.IO.Directory.Exists(".\\Inbox"))
            {
                System.IO.Directory.CreateDirectory(".\\Inbox");
            }

            string[] files = Directory.GetFiles(".\\Inbox\\");
            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.CreationTime < DateTime.Now.AddDays(-7))
                    fi.Delete();
            }
            #endregion

            await _service.ProcessExcel();
        }
    }
}
