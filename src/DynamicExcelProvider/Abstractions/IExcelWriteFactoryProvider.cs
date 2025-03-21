// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-13 11:47
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-02 23:40
// ***********************************************************************
//  <copyright file="IExcelWriteFactoryProvider.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using AggregatedGenericResultMessage.Abstractions;
using DynamicExcelProvider.Models.Request.Configuration;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using DynamicExcelProvider.Models.Request.Export;
using DynamicExcelProvider.WorkXCore.Models;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

#endregion

namespace DynamicExcelProvider.Abstractions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     EXCEL write provider factory from custom input model.
    /// </summary>
    /// =================================================================================================
    public interface IExcelWriteFactoryProvider
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generate file byte[], must be saved as CVS and returned to client.
        /// </summary>
        /// <param name="embeddedModelCollection">
        ///     Predefined and all model collection properties which are available to manipulate with
        ///     them.
        /// </param>
        /// <param name="availablePropInOutput">
        ///     Available export properties which will be present in result file.
        /// </param>
        /// <param name="data">Custom properties with there values in array of array.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The generate CSV from known.
        /// </returns>
        /// =================================================================================================
        Task<IResult<byte[]>> GenerateCsvFromKnownAsync(
            IReadOnlyCollection<PropModel> embeddedModelCollection, IReadOnlyCollection<PropTranslateModel> availablePropInOutput,
            IEnumerable<IReadOnlyList<PropNameValue>> data, CancellationToken cancellationToken = default);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generate file byte[], must be saved as CVS and returned to client.
        /// </summary>
        /// <typeparam name="TDataModel">.</typeparam>
        /// <param name="embeddedModelCollection">
        ///     Predefined and all model collection properties which are available to manipulate with
        ///     them.
        /// </param>
        /// <param name="availablePropInOutput">
        ///     Available export properties which will be present in result file.
        /// </param>
        /// <param name="data">Available data to export.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The CSV asynchronous.
        /// </returns>
        /// =================================================================================================
        Task<IResult<byte[]>> GenerateCsvAsync<TDataModel>(
            IReadOnlyCollection<PropModel> embeddedModelCollection, IReadOnlyCollection<PropTranslateModel> availablePropInOutput,
            IReadOnlyCollection<TDataModel> data, CancellationToken cancellationToken = default) where TDataModel : class;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an asynchronous Excel file.
        /// </summary>
        /// <typeparam name="TDataModel">Type of the data model.</typeparam>
        /// <param name="data">Available data to export.</param>
        /// <param name="cultureId">Culture id.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The generate.
        /// </returns>
        /// =================================================================================================
        Task<IResult<byte[]>> GenerateAsync<TDataModel>(IReadOnlyCollection<TDataModel> data,
            int cultureId, CancellationToken cancellationToken = default) where TDataModel : class;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an asynchronous Excel file.
        /// </summary>
        /// <typeparam name="TDataModel">Type of the data model.</typeparam>
        /// <param name="stream">The stream.</param>
        /// <param name="data">Available data to export.</param>
        /// <param name="cultureId">Culture id.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The generate.
        /// </returns>
        /// =================================================================================================
        Task<IResult> GenerateAsync<TDataModel>(Stream stream, IReadOnlyCollection<TDataModel> data,
            int cultureId, CancellationToken cancellationToken = default) where TDataModel : class;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an asynchronous Excel file.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The generate.
        /// </returns>
        /// =================================================================================================
        Task<IResult<byte[]>> GenerateAsync(ExcelCollectionExportConfiguration request, CancellationToken cancellationToken = default);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an asynchronous Excel file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="request">The request.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The generate.
        /// </returns>
        /// =================================================================================================
        Task<IResult> GenerateAsync(Stream stream, ExcelCollectionExportConfiguration request, CancellationToken cancellationToken = default);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Writes a excel (spreadsheet, worksheet) file.
        /// </summary>
        /// <param name="filePath">Full pathname of the file.</param>
        /// <param name="workBook">The work book.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        IResult Generate(string filePath, WorkbookDefinition workBook);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Writes a excel (spreadsheet, worksheet) file asynchronous.
        /// </summary>
        /// <param name="filePath">Full pathname of the file.</param>
        /// <param name="workBook">The work book.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The write file.
        /// </returns>
        /// =================================================================================================
        Task<IResult> GenerateAsync(string filePath, WorkbookDefinition workBook, CancellationToken cancellationToken = default);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Writes a excel (spreadsheet, worksheet) file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="workBook">The work book.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        IResult Generate(Stream stream, WorkbookDefinition workBook);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Writes a excel (spreadsheet, worksheet) file asynchronous.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="workBook">The work book.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The write file.
        /// </returns>
        /// =================================================================================================
        Task<IResult> GenerateAsync(Stream stream, WorkbookDefinition workBook, CancellationToken cancellationToken = default);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates a template.
        /// </summary>
        /// <typeparam name="T">Generic type parameter.</typeparam>
        /// <param name="lcid">The lcid.</param>
        /// <param name="customOutFields">(Optional) Custom user defined output/result fields.</param>
        /// <returns>
        ///     The template.
        /// </returns>
        /// =================================================================================================
        IResult<byte[]> GenerateTemplate<T>(int lcid, IReadOnlyCollection<string> customOutFields = null) where T : class;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates a template.
        /// </summary>
        /// <typeparam name="T">Generic type parameter.</typeparam>
        /// <param name="stream">The stream.</param>
        /// <param name="lcid">The lcid.</param>
        /// <param name="customOutFields">(Optional) Custom user defined output/result fields.</param>
        /// <returns>
        ///     The template.
        /// </returns>
        /// =================================================================================================
        IResult GenerateTemplate<T>(MemoryStream stream, int lcid,
            IReadOnlyCollection<string> customOutFields = null) where T : class;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates a template asynchronous.
        /// </summary>
        /// <typeparam name="T">Generic type parameter.</typeparam>
        /// <param name="lcid">The lcid.</param>
        /// <param name="customOutFields">(Optional) Custom user defined output/result fields.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The template asynchronous.
        /// </returns>
        /// =================================================================================================
        Task<IResult<byte[]>> GenerateTemplateAsync<T>(int lcid,
            IReadOnlyCollection<string> customOutFields = null, CancellationToken cancellationToken = default) where T : class;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates a template asynchronous.
        /// </summary>
        /// <typeparam name="T">Generic type parameter.</typeparam>
        /// <param name="stream">The stream.</param>
        /// <param name="lcid">The lcid.</param>
        /// <param name="customOutFields">(Optional) Custom user defined output/result fields.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The template asynchronous.
        /// </returns>
        /// =================================================================================================
        Task<IResult> GenerateTemplateAsync<T>(MemoryStream stream, int lcid,
            IReadOnlyCollection<string> customOutFields = null, CancellationToken cancellationToken = default) where T : class;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates a template.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="configuration">The configuration.</param>
        /// <returns>
        ///     The template.
        /// </returns>
        /// =================================================================================================
        IResult GenerateTemplate(Stream stream, ExcelTemplateWriteConfiguration configuration);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates a template asynchronous.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="configuration">The configuration.</param>
        /// <returns>
        ///     The template asynchronous.
        /// </returns>
        /// =================================================================================================
        Task<IResult> GenerateTemplateAsync(Stream stream, ExcelTemplateWriteConfiguration configuration);
    }
}