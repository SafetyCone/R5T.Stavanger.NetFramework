using System;

using Microsoft.Extensions.DependencyInjection;

using R5T.Dacia;


namespace R5T.Stavanger.NetFramework
{
    public static class IServiceCollectionExtensions
    {
        /// <summary>
        /// Adds the <see cref="ShortcutOperator"/> implementation of <see cref="IShortcutOperator"/> as a <see cref="ServiceLifetime.Singleton"/>.
        /// </summary>
        public static IServiceCollection AddNetFrameworkShortcutOperator(this IServiceCollection services,
            ServiceAction<IShortcutPathConventions> addShortcutPathConventions)
        {
            services
                .AddSingleton<IShortcutOperator, ShortcutOperator>()
                .RunServiceAction(addShortcutPathConventions)
                ;

            return services;
        }

        /// <summary>
        /// Adds the <see cref="ShortcutOperator"/> implementation of <see cref="IShortcutOperator"/> as a <see cref="ServiceLifetime.Singleton"/>.
        /// </summary>
        public static ServiceAction<IShortcutOperator> AddNetFrameworkShortcutOperatorAction(this IServiceCollection services,
            ServiceAction<IShortcutPathConventions> addShortcutPathConventions)
        {
            var serviceAction = new ServiceAction<IShortcutOperator>(() => services.AddNetFrameworkShortcutOperator(addShortcutPathConventions));
            return serviceAction;
        }
    }
}
