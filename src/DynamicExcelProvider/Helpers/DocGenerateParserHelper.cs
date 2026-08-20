// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-13 12:05
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-02 23:38
// ***********************************************************************
//  <copyright file="DocGenerateParserHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.Extensions;
using DynamicExcelProvider.Helpers.DataTable;
using DynamicExcelProvider.Models.Request.Configuration;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using DynamicExcelProvider.Models.Request.Export;
using DynamicExcelProvider.WorkXCore.Helpers;
using DynamicExcelProvider.WorkXCore.Models;
using RzR.Extensions.Domain.Collections;
using RzR.Extensions.Domain.Primitives;
using RzR.Extensions.Domain.Text;
using RzR.Extensions.Domain.Validation;
using RzR.ResultMessage;
using RzR.ResultMessage.Abstractions;
using RzR.ResultMessage.Extensions.Result.Messages;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

#endregion

[assembly: InternalsVisibleTo("GeneralDocumentGeneratorTests")]

namespace DynamicExcelProvider.Helpers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A document generate parser helper.
    /// </summary>
    /// =================================================================================================
    internal static class DocGenerateParserHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file bytes.
        /// </summary>
        /// <param name="embeddedModelCollection">Collection of embedded models.</param>
        /// <param name="availablePropInOutput">The available property in output.</param>
        /// <param name="data">The data.</param>
        /// <param name="encoding">
        ///     (Optional) Encoding of the generated CSV bytes. When omitted, <c>iso-8859-1</c> is used to
        ///     preserve the historical output. That code page cannot represent characters outside Latin-1
        ///     (for example 'ș', 'Ș' or '€'), which are irreversibly replaced by '?'. Pass an explicit
        ///     Unicode encoding when such characters must survive.
        /// </param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        internal static IResult<byte[]> GenerateCsv(
            IReadOnlyCollection<PropModel> embeddedModelCollection,
            IReadOnlyCollection<PropTranslateModel> availablePropInOutput,
            IEnumerable<IEnumerable<PropNameValue>> data,
            Encoding encoding = null)
        {
            DataTableHelper.InitDataTable(embeddedModelCollection, availablePropInOutput);
            var table = DataTableHelper.CreateTableAndColumns();
            foreach (var record in data) table.AddRecordFromKnown(record);

            return Result<byte[]>
                .Success(GetCsvEncoding(encoding).GetBytes(table.RExtToCSV()));
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file bytes.
        /// </summary>
        /// <typeparam name="TDataModel">Type of the data model.</typeparam>
        /// <param name="embeddedModelCollection">Collection of embedded models.</param>
        /// <param name="availablePropInOutput">The available property in output.</param>
        /// <param name="data">The data.</param>
        /// <param name="encoding">
        ///     (Optional) Encoding of the generated CSV bytes. When omitted, <c>iso-8859-1</c> is used to
        ///     preserve the historical output. That code page cannot represent characters outside Latin-1
        ///     (for example 'ș', 'Ș' or '€'), which are irreversibly replaced by '?'. Pass an explicit
        ///     Unicode encoding when such characters must survive.
        /// </param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        internal static IResult<byte[]> GenerateCsv<TDataModel>(
            IReadOnlyCollection<PropModel> embeddedModelCollection,
            IReadOnlyCollection<PropTranslateModel> availablePropInOutput,
            IReadOnlyCollection<TDataModel> data,
            Encoding encoding = null) where TDataModel : class
        {
            DataTableHelper.InitDataTable(embeddedModelCollection, availablePropInOutput);
            var table = DataTableHelper.CreateTableAndColumns();
            foreach (var record in data) table.AddRecord(record);

            return Result<byte[]>
                .Success(GetCsvEncoding(encoding).GetBytes(table.RExtToCSV()));
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Resolves the encoding used to render CSV output.
        /// </summary>
        /// <param name="encoding">The requested encoding, or null to use the default one.</param>
        /// <returns>
        ///     The requested encoding, or <c>iso-8859-1</c> when none was requested.
        /// </returns>
        /// =================================================================================================
        private static Encoding GetCsvEncoding(Encoding encoding) => encoding ?? Encoding.GetEncoding("iso-8859-1");

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file bytes.
        /// </summary>
        /// <typeparam name="TResult">Type of the result.</typeparam>
        /// <param name="data">The data.</param>
        /// <param name="cultureId">(Optional) Identifier for the culture.</param>
        /// <param name="sheetName">(Optional) Name of the sheet.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        internal static IResult<byte[]> Generate<TResult>(IReadOnlyCollection<TResult> data, int cultureId = 1033,
            string sheetName = "Sheet1")
        {
            var infoDataModel = WorkbookParseBuildHelper.BuildAndParseInternalModel(data.FirstOrDefault(), cultureId);
            if (infoDataModel.IsSuccess.IsFalse())
                return Result<byte[]>.Failure(infoDataModel.GetFirstMessage());

            var (outputProps, embeddedModelCollection) = infoDataModel.Response;

            var ms = new MemoryStream();
            var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(
                sheetName,
                outputProps,
                embeddedModelCollection,
                data,
                false);
            if (wbDef.IsSuccess.IsFalse())
                return Result<byte[]>.Failure(wbDef.GetFirstMessage());

            var document = SpreadsheetDocumentHelper.Instance.Write(ms, wbDef.Response);

            return document.IsSuccess.IsFalse()
                ? Result<byte[]>.Failure(document.Messages.FirstOrDefault()?.Message)
                : Result<byte[]>.Success(ms.ToArray());
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file.
        /// </summary>
        /// <typeparam name="TResult">Type of the result.</typeparam>
        /// <param name="stream">The stream.</param>
        /// <param name="data">The data.</param>
        /// <param name="cultureId">(Optional) Identifier for the culture.</param>
        /// <param name="sheetName">(Optional) Name of the sheet.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        internal static IResult Generate<TResult>(Stream stream, IReadOnlyCollection<TResult> data, int cultureId = 1033,
            string sheetName = "Sheet1")
        {
            var infoDataModel = WorkbookParseBuildHelper.BuildAndParseInternalModel(data.FirstOrDefault(), cultureId);
            if (infoDataModel.IsSuccess.IsFalse())
                return Result.Failure(infoDataModel.GetFirstMessage());

            var (outputProps, embeddedModelCollection) = infoDataModel.Response;

            var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(
                sheetName,
                outputProps,
                embeddedModelCollection,
                data,
                false);
            if (wbDef.IsSuccess.IsFalse())
                return Result.Failure(wbDef.GetFirstMessage());

            var document = SpreadsheetDocumentHelper.Instance.Write(stream, wbDef.Response);

            return document.IsSuccess.IsFalse()
                ? Result.Failure(document.Messages.FirstOrDefault()?.Message)
                : Result.Success();
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file.
        /// </summary>
        /// <param name="exportConfiguration">The export configuration.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        internal static IResult<byte[]> Generate(ExcelCollectionExportConfiguration exportConfiguration)
        {
            var infoDataModel = (IResult<(List<PropTranslateModel>, List<PropModel>)>)WorkbookParseBuildHelper
                .BuildAndParseInternalModelDynamic(
                exportConfiguration.Configuration,
                exportConfiguration.DataCollection.FirstOrDefault());
            if (infoDataModel.IsSuccess.IsFalse())
                return Result<byte[]>.Failure(infoDataModel.GetFirstMessage());

            var (outputProps, embeddedModelCollection) = infoDataModel.Response;

            var ms = new MemoryStream();

            var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(
                exportConfiguration.Configuration.SheetName,
                outputProps,
                embeddedModelCollection,
                exportConfiguration.DataCollection,
                true);
            if (wbDef.IsSuccess.IsFalse())
                return Result<byte[]>.Failure(wbDef.GetFirstMessage());

            var document = SpreadsheetDocumentHelper.Instance.Write(ms, wbDef.Response);

            return document.IsSuccess.IsFalse()
                ? Result<byte[]>.Failure(document.Messages.FirstOrDefault()?.Message)
                : Result<byte[]>.Success(ms.ToArray());
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="exportConfiguration">The export configuration.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        internal static IResult Generate(Stream stream, ExcelCollectionExportConfiguration exportConfiguration)
        {
            var infoDataModel = (IResult<(List<PropTranslateModel>, List<PropModel>)>)WorkbookParseBuildHelper
                .BuildAndParseInternalModelDynamic(
                exportConfiguration.Configuration,
                exportConfiguration.DataCollection.FirstOrDefault());
            if (infoDataModel.IsSuccess.IsFalse())
                return Result.Failure(infoDataModel.GetFirstMessage());

            var (outputProps, embeddedModelCollection) = infoDataModel.Response;

            var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(
                exportConfiguration.Configuration.SheetName,
                outputProps,
                embeddedModelCollection,
                exportConfiguration.DataCollection,
                true);
            if (wbDef.IsSuccess.IsFalse())
                return Result.Failure(wbDef.GetFirstMessage());

            var document = SpreadsheetDocumentHelper.Instance.Write(stream, wbDef.Response);

            return document.IsSuccess.IsFalse()
                ? Result.Failure(document.Messages.FirstOrDefault()?.Message)
                : Result.Success();
        }

        internal static IResult<byte[]> Generate(System.Data.DataSet dataSet)
        {
            try
            {
                DomainEnsure.IsNotNull(dataSet, nameof(dataSet));

                var worksheets = new List<WorksheetDefinition>();
                var idx = 0;
                foreach (System.Data.DataTable table in dataSet.Tables)
                {
                    var parsedTable = table.ConvertToRowData();
                    if (parsedTable.IsSuccess.IsFalse())
                        return Result<byte[]>.Failure(parsedTable.GetFirstMessage());

                    var (sheetName, ptm, pm, resultRows) = parsedTable.Response;

                    var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(
                        sheetName.IfNullOrEmpty($"Sheet{idx += 1}").RExtToCleanSheetName(),
                        ptm as IReadOnlyCollection<PropTranslateModel>,
                        pm as IReadOnlyCollection<PropModel>,
                        resultRows);

                    if (wbDef.IsSuccess.IsFalse())
                        return Result<byte[]>.Failure(wbDef.GetFirstMessage());

                    worksheets.AddRange(wbDef.Response.Worksheets);
                }

                var ms = new MemoryStream();
                var document = SpreadsheetDocumentHelper.Instance.Write(ms, new WorkbookDefinition(worksheets));

                return document.IsSuccess.IsFalse()
                    ? Result<byte[]>.Failure(document.Messages.FirstOrDefault()?.Message)
                    : Result<byte[]>.Success(ms.ToArray());
            }
            catch (Exception e)
            {
                return Result<byte[]>.Failure(e.Message)
                    .AddError(e);
            }
        }

        internal static IResult Generate(System.Data.DataSet dataSet, Stream stream)
        {
            try
            {
                DomainEnsure.IsNotNull(dataSet, nameof(dataSet));

                var worksheets = new List<WorksheetDefinition>();
                var idx = 0;
                foreach (System.Data.DataTable table in dataSet.Tables)
                {
                    var parsedTable = table.ConvertToRowData();
                    if (parsedTable.IsSuccess.IsFalse())
                        return Result.Failure(parsedTable.GetFirstMessage());

                    var (sheetName, ptm, pm, resultRows) = parsedTable.Response;

                    var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(
                        sheetName.IfNullOrEmpty($"Sheet{idx += 1}").RExtToCleanSheetName(),
                        ptm as IReadOnlyCollection<PropTranslateModel>,
                        pm as IReadOnlyCollection<PropModel>,
                        resultRows);
                    if (wbDef.IsSuccess.IsFalse())
                        return Result.Failure(wbDef.GetFirstMessage());

                    worksheets.AddRange(wbDef.Response.Worksheets);
                }

                var document = SpreadsheetDocumentHelper.Instance.Write(stream, new WorkbookDefinition(worksheets));

                return document.IsSuccess.IsFalse()
                    ? Result.Failure(document.GetFirstMessage())
                    : Result.Success();
            }
            catch (Exception e)
            {
                return Result.Failure(e.Message)
                    .AddError(e);
            }
        }

        internal static IResult Generate(System.Data.DataSet dataSet, string filePath)
        {
            try
            {
                DomainEnsure.IsNotNull(dataSet, nameof(dataSet));

                var worksheets = new List<WorksheetDefinition>();
                var idx = 0;
                foreach (System.Data.DataTable table in dataSet.Tables)
                {
                    var parsedTable = table.ConvertToRowData();
                    if (parsedTable.IsSuccess.IsFalse())
                        return Result.Failure(parsedTable.GetFirstMessage());

                    var (sheetName, ptm, pm, resultRows) = parsedTable.Response;

                    var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(
                        sheetName.IfNullOrEmpty($"Sheet{idx += 1}").RExtToCleanSheetName(),
                        ptm as IReadOnlyCollection<PropTranslateModel>,
                        pm as IReadOnlyCollection<PropModel>,
                        resultRows);
                    if (wbDef.IsSuccess.IsFalse())
                        return Result.Failure(wbDef.GetFirstMessage());

                    worksheets.AddRange(wbDef.Response.Worksheets);
                }

                var document = SpreadsheetDocumentHelper.Instance.Write(filePath, new WorkbookDefinition(worksheets));

                return document.IsSuccess.IsFalse()
                    ? Result.Failure(document.GetFirstMessage())
                    : Result.Success();
            }
            catch (Exception e)
            {
                return Result.Failure(e.Message)
                    .AddError(e);
            }
        }

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
        internal static IResult<byte[]> GenerateTemplate<T>(
            int lcid,
            IReadOnlyCollection<string> customOutFields = null)
            where T : class
        {
            const string sheetName = "Sheet1";
            var infoDataModel = WorkbookParseBuildHelper.BuildAndParseInternalModelDynamic(
                new ExcelWriteConfiguration(sheetName, lcid), null, typeof(T));
            if (infoDataModel.IsSuccess.IsFalse())
                return Result<byte[]>.Failure(infoDataModel.GetFirstMessage());

            var (outputProps, embeddedModelCollection) = infoDataModel.Response;

            if (customOutFields.IsNullOrEmptyEnumerable().IsFalse())
                outputProps = outputProps.Where(x => customOutFields!.Contains(x.CommonName)).ToList();

            var ms = new MemoryStream();

            var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(
                sheetName: sheetName,
                outputProps: outputProps,
                embeddedModelCollection: embeddedModelCollection,
                data: new List<T>(),
                isDynamic: true,
                sourceRowType: typeof(T),
                generateSheetValidations: true);
            if (wbDef.IsSuccess.IsFalse())
                return Result<byte[]>.Failure(wbDef.GetFirstMessage());

            var document = SpreadsheetDocumentHelper.Instance.Write(
                stream: ms,
                workBook: wbDef.Response,
                appendSheetValidations: true);

            return document.IsSuccess.IsFalse()
                ? Result<byte[]>.Failure(document.Messages.FirstOrDefault()?.Message)
                : Result<byte[]>.Success(ms.ToArray());
        }

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
        internal static IResult GenerateTemplate<T>(
            MemoryStream stream,
            int lcid,
            IReadOnlyCollection<string> customOutFields = null)
            where T : class
        {
            const string sheetName = "Sheet1";
            var infoDataModel = WorkbookParseBuildHelper.BuildAndParseInternalModelDynamic(
                new ExcelWriteConfiguration(sheetName, lcid), null, typeof(T));
            if (infoDataModel.IsSuccess.IsFalse())
                return Result.Failure(infoDataModel.GetFirstMessage());

            var (outputProps, embeddedModelCollection) = infoDataModel.Response;

            if (customOutFields.IsNullOrEmptyEnumerable().IsFalse())
                outputProps = outputProps.Where(x => customOutFields!.Contains(x.CommonName)).ToList();

            var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(
                sheetName: sheetName,
                outputProps: outputProps,
                embeddedModelCollection: embeddedModelCollection,
                data: new List<T>(),
                isDynamic: true,
                sourceRowType: typeof(T),
                generateSheetValidations: true);
            if (wbDef.IsSuccess.IsFalse())
                return Result.Failure(wbDef.GetFirstMessage());

            var document = SpreadsheetDocumentHelper.Instance.Write(
                stream: stream,
                workBook: wbDef.Response,
                appendSheetValidations: true);

            return document.IsSuccess.IsFalse()
                ? Result.Failure(document.Messages.FirstOrDefault()?.Message)
                : Result.Success();
        }

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
        internal static IResult GenerateTemplate(
            Stream stream,
            ExcelTemplateWriteConfiguration configuration)
        {
            var sheetName = configuration.SheetName.IfNullOrWhiteSpace("Sheet1");
            var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinitionTemplate(
                sheetName,
                configuration.ColumnHeadings,
                configuration.SheetValidations,
                configuration.SheetValidations.IsNullOrEmptyEnumerable().IsFalse());
            if (wbDef.IsSuccess.IsFalse())
                return Result.Failure(wbDef.GetFirstMessage());

            var document = SpreadsheetDocumentHelper.Instance.Write(
                stream: stream,
                workBook: wbDef.Response,
                appendSheetValidations: true);

            return document.IsSuccess.IsFalse()
                ? Result.Failure(document.Messages.FirstOrDefault()?.Message)
                : Result.Success();
        }
    }
}