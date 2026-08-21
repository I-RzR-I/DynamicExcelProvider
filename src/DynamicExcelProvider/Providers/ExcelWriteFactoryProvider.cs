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

using DynamicExcelProvider.Abstractions;
using DynamicExcelProvider.Helpers;
using DynamicExcelProvider.Models.Request.Configuration;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using DynamicExcelProvider.Models.Request.Export;
using DynamicExcelProvider.WorkXCore.Abstractions;
using DynamicExcelProvider.WorkXCore.Models;
using RzR.Extensions.Domain.Validation;
using RzR.ResultMessage;
using RzR.ResultMessage.Abstractions;
using RzR.ResultMessage.Extensions.Result.Messages;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

// ReSharper disable ClassNeverInstantiated.Global

#endregion

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

        #region WRITE FILE

        #region SYNC

        /// <inheritdoc />
        public IResult Generate(Stream stream, ExcelCollectionExportConfiguration request)
            => DocGenerateParserHelper.Generate(stream, request);

        /// <inheritdoc />
        public IResult<byte[]> Generate(ExcelCollectionExportConfiguration request)
        {
            try
            {
                DomainEnsure.IsNotNull(request, nameof(request));

                return DocGenerateParserHelper.Generate(request);
            }
            catch (Exception e)
            {
                return Result<byte[]>
                    .Failure("An error occurred on generate excel file")
                    .AddException(e);
            }
        }

        /// <inheritdoc />
        public IResult Generate(string filePath, WorkbookDefinition workBook)
            => _spreadsheetDocumentService.WriteFile(filePath, workBook);

        /// <inheritdoc />
        public IResult Generate(Stream stream, WorkbookDefinition workBook)
            => _spreadsheetDocumentService.WriteFile(stream, workBook);

        /// <inheritdoc />
        public IResult Generate(Stream stream, DataTable dataTable)
        {
            try
            {
                DomainEnsure.IsNotNull(dataTable, nameof(dataTable));
                DomainEnsure.IsNotNull(stream, nameof(stream));

                var dataSet = new DataSet(nameof(Generate));
                dataSet.Tables.Add(dataTable);

                return DocGenerateParserHelper.Generate(dataSet, stream);
            }
            catch (Exception e)
            {
                return Result
                    .Failure("An error occurred on generate excel file in `Stream`.")
                    .AddException(e);
            }
        }

        /// <inheritdoc />
        public IResult<byte[]> Generate(DataTable dataTable)
        {
            try
            {
                DomainEnsure.IsNotNull(dataTable, nameof(dataTable));

                var dataSet = new DataSet(nameof(Generate));
                dataSet.Tables.Add(dataTable);

                return DocGenerateParserHelper.Generate(dataSet);
            }
            catch (Exception e)
            {
                return Result<byte[]>
                    .Failure("An error occurred on generate excel file bytes.")
                    .AddException(e);
            }
        }

        /// <inheritdoc />
        public IResult Generate(string filePath, DataTable dataTable)
        {
            try
            {
                DomainEnsure.IsNotNullOrEmptyArgNull(filePath, nameof(filePath));
                DomainEnsure.IsNotNull(dataTable, nameof(dataTable));

                var dataSet = new DataSet(nameof(Generate));
                dataSet.Tables.Add(dataTable);

                return DocGenerateParserHelper.Generate(dataSet, filePath);
            }
            catch (Exception e)
            {
                return Result
                    .Failure("An error occurred on generate excel file to file path.")
                    .AddException(e);
            }
        }

        /// <inheritdoc />
        public IResult Generate(Stream stream, DataSet dataSet)
        {
            try
            {
                DomainEnsure.IsNotNull(stream, nameof(stream));
                DomainEnsure.IsNotNull(dataSet, nameof(dataSet));

                return DocGenerateParserHelper.Generate(dataSet, stream);
            }
            catch (Exception e)
            {
                return Result
                    .Failure("An error occurred on generate excel file in `Stream` from DataSet.")
                    .AddException(e);
            }
        }

        /// <inheritdoc />
        public IResult<byte[]> Generate(DataSet dataSet)
        {
            try
            {
                DomainEnsure.IsNotNull(dataSet, nameof(dataSet));

                return DocGenerateParserHelper.Generate(dataSet);
            }
            catch (Exception e)
            {
                return Result<byte[]>
                    .Failure("An error occurred on generate excel file bytes from DataSet.")
                    .AddException(e);
            }
        }

        /// <inheritdoc />
        public IResult Generate(string filePath, DataSet dataSet)
        {
            try
            {
                DomainEnsure.IsNotNullOrEmptyArgNull(filePath, nameof(filePath));
                DomainEnsure.IsNotNull(dataSet, nameof(dataSet));

                return DocGenerateParserHelper.Generate(dataSet, filePath);
            }
            catch (Exception e)
            {
                return Result<byte[]>
                    .Failure("An error occurred on generate excel file from DataSet.")
                    .AddException(e);
            }
        }

        #endregion

        #region ASYNC

        /// <inheritdoc />
        public async Task<IResult<byte[]>> GenerateCsvFromKnownAsync(
            IReadOnlyCollection<PropModel> embeddedModelCollection,
            IReadOnlyCollection<PropTranslateModel> availablePropInOutput,
            IEnumerable<IReadOnlyList<PropNameValue>> data,
            CancellationToken cancellationToken = default, Encoding encoding = null)
        {
            try
            {
                var byteData = await Task.Run(
                    () => DocGenerateParserHelper.GenerateCsv(embeddedModelCollection, availablePropInOutput, data, encoding),
                    cancellationToken);

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
            IReadOnlyCollection<TDataModel> data,
            CancellationToken cancellationToken = default, Encoding encoding = null) where TDataModel : class
        {
            try
            {
                var byteData = await Task.Run(() => DocGenerateParserHelper.GenerateCsv(embeddedModelCollection,
                    availablePropInOutput, data, encoding), cancellationToken);

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
        public async Task<IResult> GenerateAsync(string filePath, WorkbookDefinition workBook,
            CancellationToken cancellationToken = default)
            => await _spreadsheetDocumentService.WriteFileAsync(filePath, workBook, cancellationToken);

        /// <inheritdoc />
        public async Task<IResult> GenerateAsync(Stream stream, WorkbookDefinition workBook,
            CancellationToken cancellationToken = default)
            => await _spreadsheetDocumentService.WriteFileAsync(stream, workBook, cancellationToken);

        /// <inheritdoc />
        public async Task<IResult> GenerateAsync(Stream stream, DataTable dataTable,
            CancellationToken cancellationToken = default)
            => await Task.Run(() => Generate(stream, dataTable), cancellationToken);

        /// <inheritdoc />
        public async Task<IResult<byte[]>> GenerateAsync(DataTable dataTable,
            CancellationToken cancellationToken = default)
            => await Task.Run(() => Generate(dataTable), cancellationToken);

        /// <inheritdoc />
        public async Task<IResult> GenerateAsync(string filePath, DataTable dataTable,
            CancellationToken cancellationToken = default)
            => await Task.Run(() => Generate(filePath, dataTable), cancellationToken);

        /// <inheritdoc />
        public async Task<IResult> GenerateAsync(Stream stream, DataSet dataSet, CancellationToken cancellationToken = default)
            => await Task.Run(() => Generate(stream, dataSet), cancellationToken);

        /// <inheritdoc />
        public async Task<IResult<byte[]>> GenerateAsync(DataSet dataSet, CancellationToken cancellationToken = default)
            => await Task.Run(() => Generate(dataSet), cancellationToken);

        /// <inheritdoc />
        public async Task<IResult> GenerateAsync(string filePath, DataSet dataSet, CancellationToken cancellationToken = default)
            => await Task.Run(() => Generate(filePath, dataSet), cancellationToken);

        #endregion

        #endregion

        #region TEMPLATE

        /// <inheritdoc />
        public IResult<byte[]> GenerateTemplate<T>(int lcid, IReadOnlyCollection<string> customOutFields = null) where T : class
            => DocGenerateParserHelper.GenerateTemplate<T>(lcid, customOutFields);

        /// <inheritdoc />
        public IResult GenerateTemplate<T>(MemoryStream stream, int lcid,
            IReadOnlyCollection<string> customOutFields = null) where T : class
            => DocGenerateParserHelper.GenerateTemplate<T>(stream, lcid, customOutFields);

        /// <inheritdoc />
        public async Task<IResult<byte[]>> GenerateTemplateAsync<T>(int lcid,
            IReadOnlyCollection<string> customOutFields = null,
            CancellationToken cancellationToken = default) where T : class
            => await Task.Run(() => GenerateTemplate<T>(lcid, customOutFields), cancellationToken);

        /// <inheritdoc />
        public async Task<IResult> GenerateTemplateAsync<T>(MemoryStream stream, int lcid,
            IReadOnlyCollection<string> customOutFields = null,
            CancellationToken cancellationToken = default) where T : class
            => await Task.Run(() => GenerateTemplate<T>(stream, lcid, customOutFields), cancellationToken);

        /// <inheritdoc />
        public IResult GenerateTemplate(Stream stream, ExcelTemplateWriteConfiguration configuration)
            => DocGenerateParserHelper.GenerateTemplate(stream, configuration);

        /// <inheritdoc />
        public async Task<IResult> GenerateTemplateAsync(Stream stream, ExcelTemplateWriteConfiguration configuration)
            => await Task.Run(() => DocGenerateParserHelper.GenerateTemplate(stream, configuration));

        #endregion
    }
}