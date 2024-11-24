using ExcelYukleme.Service;
using Microsoft.AspNetCore.Identity;
namespace ExcelYukleme.Extensions
{
    public static class StartupExtensions
    {
        public static void AddServiceCollection(this IServiceCollection services)
        {
            services.AddScoped<ICalculateService, CalculateService>();
            services.AddScoped<IExcelService, ExcelService>();
            services.Configure<DataProtectionTokenProviderOptions>(options =>
            {
                options.TokenLifespan = TimeSpan.FromMinutes(30);
            });           

        }
    }
}
