// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-13 12:30
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-01-30 21:31
// ***********************************************************************
//  <copyright file="DependencyInjection.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DomainCommonExtensions.CommonExtensions;
using DynamicExcelProvider.Abstractions;
using DynamicExcelProvider.Helpers;
using DynamicExcelProvider.Models.Request;
using DynamicExcelProvider.Providers;
using DynamicExcelProvider.WorkXCore.Abstractions;
using DynamicExcelProvider.WorkXCore.Services;
using Microsoft.Extensions.DependencyInjection;
using System;

// ReSharper disable UnusedMethodReturnValue.Local

#endregion

namespace DynamicExcelProvider
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A dependency injection.
    /// </summary>
    /// <remarks>
    /// </remarks>
    /// =================================================================================================
    public static class DependencyInjection
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Register excel data provider in DI.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <typeparam name="TExcelProvider">.</typeparam>
        /// <param name="services">.</param>
        /// <returns>
        ///     An IServiceCollection.
        /// </returns>
        /// =================================================================================================
        private static IServiceCollection RegisterExcelDataSourceProvider<TExcelProvider>(this IServiceCollection services)
            where TExcelProvider : class, IExcelWriteFactoryProvider
        {
            services.AddTransient<ISpreadsheetDocumentService, SpreadsheetDocumentService>();
            services.AddTransient<IExcelWriteFactoryProvider, TExcelProvider>();

            return services;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Register excel data provider in DI.
        /// </summary>
        /// <param name="services">.</param>
        /// <param name="writeOption">(Optional) The write option.</param>
        /// <returns>
        ///     An IServiceCollection.
        /// </returns>
        /// =================================================================================================
        public static IServiceCollection RegisterExcelDataSourceProvider(this IServiceCollection services,
            Action<ExcelWriteProviderOption> writeOption = null)
        {
            SetProviderOptions(writeOption);

            services.RegisterExcelDataSourceProvider<ExcelWriteFactoryProvider>();

            return services;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Sets provider options.
        /// </summary>
        /// <param name="writeOption">The write option.</param>
        /// =================================================================================================
        private static void SetProviderOptions(Action<ExcelWriteProviderOption> writeOption)
        {
            /*
            writeOption = writeOption.IfIsNull(o =>
            {
                o.ApplyMaxColumnNumberPolicy = true;
                o.ApplyMaxRowNumberPolicy = true;
                o.SheetMaxNumberOfRows = ProviderInitInfo.DefaultMaxRowNumber;
            });
            */

            if (writeOption.IsNotNull())
            {
                var wOption = new ExcelWriteProviderOption();
                writeOption!(wOption);

                ProviderInitInfo.SetMaxRowNumberPolicyRule(wOption.ApplyMaxRowNumberPolicy);
                ProviderInitInfo.SetSheetMaxNumberOfRows(wOption.SheetMaxNumberOfRows);
            }
        }
    }
}