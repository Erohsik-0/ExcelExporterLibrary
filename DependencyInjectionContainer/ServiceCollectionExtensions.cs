using ExportExcel.Interfaces;
using ExportExcel.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace ExportExcel.DependencyInjectionContainer
{
    public static class ServiceCollectionExtensions
    {

        public static IServiceCollection AddExcelExporterServices(this IServiceCollection services)
        {

            services.AddScoped<IExcelExporter , ExcelExporterOrchestrator>();
            services.AddScoped<IAsyncExcelExporter , ExcelExporterOrchestrator>();
            services.AddScoped<IExcelImporter , ExcelImporter>();
            services.AddScoped<IAsyncExcelImporter, ExcelImporter>();
            services.AddScoped<IWorksheetManager, WorksheetManager>();
            services.AddScoped<IJsonFlattener , JsonFlattener>();
            services.AddScoped<IStructureAnalyzer , StructureAnalyzer>();
            services.AddScoped<IDataValidator , DataValidator>();
            services.AddScoped<ITypeDetector , TypeDetector>();
            services.AddScoped<IStructureAnalyzer, StructureAnalyzer>();

            services.AddLogging();

            return services;
        }
    }
}
