using FluentNHibernate.Cfg.Db;
using FluentNHibernate.Cfg;
using GitHubActionsReportGenerator.Repositories;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using NHibernate.Driver;
using NHibernate;
using System.Diagnostics;
using System.Net.WebSockets;
using GitHubActionsReportGenerator;
using GitHubActionsReportGenerator.SheetGenerators;
using GitHubActionsReportGenerator.SheetUpdaters;

using IHost host = CreateHostBuilder(args).Build();
using var scope = host.Services.CreateScope();

var services = scope.ServiceProvider;

try
{
    // Get all registered sheet updaters
    var updaters = services.GetServices<ISheetUpdater>();
    foreach (var updater in updaters)
    {
        await updater.Update();
    }
    var test = 1;
}
catch (Exception e)
{
    Console.WriteLine(e.Message);
}

Console.WriteLine("Finished");

IHostBuilder CreateHostBuilder(string[] strings)
{
    return Host.CreateDefaultBuilder()
        .ConfigureServices((_, services) =>
        {
            services.AddTransient<IRunSummaryRepository, RunSummaryRepository>();
            services.AddTransient<FlakyTestsRepository>();

            // Register all sheet updaters
            services.AddTransient<ISheetUpdater, RunSummarySheetUpdater>();
            services.AddTransient<ISheetUpdater, FlakyTestsSheetUpdater>();
            // Add any additional sheet updaters here
            // services.AddTransient<ISheetUpdater, AnotherSheetUpdater>();
            // services.AddTransient<ISheetUpdater, YetAnotherSheetUpdater>();

            // NHibernate session factory registration
            services.AddSingleton<ISessionFactory>(CreateNHibernateSessionFactory());
        });
}

ISessionFactory CreateNHibernateSessionFactory()
{
    return Fluently.Configure()
      .Database(
        MsSqlConfiguration.MsSql2012
            .Driver<MicrosoftDataSqlClientDriver>()
            .ConnectionString("Server=.\\SQLEXPRESS;Database=GHAData;Integrated Security=True;Encrypt=false"))
      .Mappings(m =>
        m.FluentMappings.AddFromAssemblyOf<Program>())
      .BuildSessionFactory();
}