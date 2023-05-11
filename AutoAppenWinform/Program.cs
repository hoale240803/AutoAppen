using AutoAppenWinform.Services;
using AutoAppenWinform.Services.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace AutoAppenWinform
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            try
            {

                // To customize application configuration such as set high DPI settings or default font,
                // see https://aka.ms/applicationconfiguration.
                ApplicationConfiguration.Initialize();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                var host = CreateHostBuilder().Build();
                ServiceProvider = host.Services;

                Application.Run(ServiceProvider.GetRequiredService<AutoAppen>());

                //Application.Run(new AutoAppen());
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        // TODO: 1. Dependency Injection in winform more detail
        // https://stackoverflow.com/questions/70475830/how-to-use-dependency-injection-in-winforms

        public static IServiceProvider ServiceProvider { get; private set; }

        private static IHostBuilder CreateHostBuilder()
        {
            return Host.CreateDefaultBuilder()
            .ConfigureServices((context, services) =>
            {
                services.AddScoped<IGmailService, GmailService>();
                services.AddScoped<AutoAppen>();
            });
        }
    }
}