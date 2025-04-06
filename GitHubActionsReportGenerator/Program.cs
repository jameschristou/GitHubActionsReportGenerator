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

using IHost host = CreateHostBuilder(args).Build();
using var scope = host.Services.CreateScope();

var services = scope.ServiceProvider;

try
{
    // do something
    // have a list of sheet generators
    // each sheet generator can execute SQL to get the data it needs
    var repo = services.GetRequiredService<IRunSummaryRepository>();

    var results = await repo.GetRunSummary(DateTime.Now.AddDays(-7));

    var updater = services.GetRequiredService<ISheetUpdater>();
    await updater.Update();
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
            services.AddTransient<ISheetUpdater, RunSummarySheetUpdater>();

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