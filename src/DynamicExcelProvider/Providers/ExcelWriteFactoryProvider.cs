// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-13 11:51
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:18
// ***********************************************************************
//  <copyright file="ExcelWriteFactoryProvider.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using AggregatedGenericResultMessage;
using AggregatedGenericResultMessage.Abstractions;
using AggregatedGenericResultMessage.Extensions.Result.Messages;
using DynamicExcelProvider.Abstractions;
using DynamicExcelProvider.Helpers;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using DynamicExcelProvider.Models.Request.Export;
using DynamicExcelProvider.WorkXCore.Abstractions;
using DynamicExcelProvider.WorkXCore.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

#endregion

// ReSharper disable ClassNeverInstantiated.Global

namespace DynamicExcelProvider.Providers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     An excel write factory provider.
    /// </summary>
    /// <seealso cref="T:DynamicExcelProvider.Abstractions.IExcelWriteFactoryProvider" />
    /// =================================================================================================
    public class ExcelWriteFactoryProvider : IExcelWriteFactoryProvider
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) the spreadsheet document service.
        /// </summary>
        /// =================================================================================================
        private readonly ISpreadsheetDocumentService _spreadsheetDocumentService;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelWriteFactoryProvider" /> class.
        /// </summary>
        /// <param name="spreadsheetDocumentService">The spreadsheet document service.</param>
        /// =================================================================================================
        public ExcelWriteFactoryProvider(ISpreadsheetDocumentService spreadsheetDocumentService)
            => _spreadsheetDocumentService = spreadsheetDocumentService;

        /// <inheritdoc />
        public async Task<IResult<byte[]>> GenerateCsvFromKnownAsync(
            IReadOnlyCollection<PropModel> embeddedModelCollection,
            IReadOnlyCollection<PropTranslateModel> availablePropInOutput,
            IEnumerable<IReadOnlyList<PropNameValue>> data, CancellationToken cancellationToken = default)
        {
            try
            {
                var byteData = await Task.Run(() => DocGenerateParserHelper.Generate(embeddedModelCollection, availablePropInOutput, data), cancellationToken);

                return byteData;
            }
            catch (Exception e)
            {
                return Result<byte[]>
                    .Failure("An error occurred on generate excel file")
                    .AddException(e);
            }
        }

        /// <inheritdoc />
        public async Task<IResult<byte[]>> GenerateCsvAsync<TDataModel>(
            IReadOnlyCollection<PropModel> embeddedModelCollection,
            IReadOnlyCollection<PropTranslateModel> availablePropInOutput,
            IReadOnlyCollection<TDataModel> data, CancellationToken cancellationToken = default) where TDataModel : class
        {
            try
            {
                var byteData = await Task.Run(() => DocGenerateParserHelper.Generate(embeddedModelCollection, 
                    availablePropInOutput, data), cancellationToken);

                return byteData;
            }
            catch (Exception e)
            {
                return Result<byte[]>
                    .Failure("An error occurred on generate excel file")
                    .AddException(e);
            }
        }

        /// <inheritdoc />
        public async Task<IResult<byte[]>> GenerateAsync<TDataModel>(
            IReadOnlyCollection<TDataModel> data,
            int cultureId, CancellationToken cancellationToken = default) where TDataModel : class
        {
            try
            {
                var byteData = await Task.Run(() => DocGenerateParserHelper.Generate(data, cultureId), cancellationToken);

                return byteData;
            }
            catch (Exception e)
            {
                return Result<byte[]>
                    .Failure("An error occurred on generate excel file")
                    .AddException(e);
            }
        }

        /// <inheritdoc />
        public async Task<IResult> GenerateAsync<TDataModel>(
            Stream stream, IReadOnlyCollection<TDataModel> data,
            int cultureId, CancellationToken cancellationToken = default) where TDataModel : class
            => await Task.Run(() => DocGenerateParserHelper.Generate(stream, data, cultureId), cancellationToken);

        /// <inheritdoc />
        public async Task<IResult<byte[]>> GenerateAsync(ExcelCollectionExportConfiguration request, 
            CancellationToken cancellationToken = default)
        {
            try
            {
                var byteData = await Task.Run(() => DocGenerateParserHelper.Generate(request), cancellationToken);

                return byteData;
            }
            catch (Exception e)
            {
                return Result<byte[]>
                    .Failure("An error occurred on generate excel file")
                    .AddException(e);
            }
        }

        /// <inheritdoc />
        public async Task<IResult> GenerateAsync(Stream stream, ExcelCollectionExportConfiguration request, 
            CancellationToken cancellationToken = default)
            => await Task.Run(() => DocGenerateParserHelper.Generate(stream, request), cancellationToken);

        /// <inheritdoc />
        public IResult Generate(string filePath, WorkbookDefinition workBook)
            => _spreadsheetDocumentService.WriteFile(filePath, workBook);

        /// <inheritdoc />
        public async Task<IResult> GenerateAsync(string filePath, WorkbookDefinition workBook, 
            CancellationToken cancellationToken = default)
            => await _spreadsheetDocumentService.WriteFileAsync(filePath, workBook, cancellationToken);

        /// <inheritdoc />
        public IResult Generate(Stream stream, WorkbookDefinition workBook)
            => _spreadsheetDocumentService.WriteFile(stream, workBook);

        /// <inheritdoc />
        public async Task<IResult> GenerateAsync(
            Stream stream, WorkbookDefinition workBook,
            CancellationToken cancellationToken = default)
            => await _spreadsheetDocumentService.WriteFileAsync(stream, workBook, cancellationToken);
    }
}