using BrickAtHeart.LUGTools.PicklistGenerator.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;

namespace BrickAtHeart.LUGTools.PicklistGenerator
{
    public class Program
    {
        public static void Main(string[] args)
        {
            ServiceCollection services = new ServiceCollection();
            ConfigureServices(services, args);

            ServiceProvider serviceProvider = services.BuildServiceProvider();

            ExportFile exportFile = serviceProvider.GetService<ExportFile>();

            if (exportFile != null)
            {
                List<Person> people = exportFile.ReadPeople();
                Dictionary<int, Part> parts = exportFile.ReadParts();
                List<Order> orders = exportFile.ReadOrders(people, parts);

                PerPartPicklist perPart = serviceProvider.GetService<PerPartPicklist>();
                perPart?.Generate(orders);

                PerPersonPicklist perPerson = serviceProvider.GetService<PerPersonPicklist>();
                perPerson?.Generate(orders);
            }
        }

        private static void ConfigureServices(IServiceCollection services, string[] args)
        {
            services.AddLogging(builder =>
            {
                builder.AddConsole();
                builder.AddDebug();
            });

            IConfigurationRoot configuration = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appSettings.json")
                .AddCommandLine(args)
                .Build();

            services.Configure<PicklistGeneratorOptions>(configuration.GetSection(PicklistGeneratorOptions.Section));

            services.AddSingleton<ExportFile>();
            services.AddSingleton<PerPartPicklist>();
            services.AddSingleton<PerPersonPicklist>();
        }
    }
}