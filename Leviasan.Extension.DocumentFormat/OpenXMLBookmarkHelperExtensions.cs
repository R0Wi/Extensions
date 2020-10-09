using Microsoft.Extensions.DependencyInjection;
using WebAPI.Helpers;

namespace Microsoft.Extensions.DependencyInjection
{
    /// <summary>
    /// ASP.NET Core extensions for OpenXML bookmark helper.
    /// </summary>
    public static class OpenXMLBookmarkHelperExtensions
    {
        /// <summary>
        /// Adds the OpenXML bookmark helper.
        /// </summary>
        /// <param name="services"></param>
        public static IServiceCollection AddOpenXMLBookmarkHelper(this IServiceCollection services)
        {
            services.AddScoped<OpenXMLBookmarkHelper>();
            return services;
        }
    }
}
