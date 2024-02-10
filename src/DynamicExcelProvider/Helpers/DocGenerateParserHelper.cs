﻿// ***********************************************************************
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

using AggregatedGenericResultMessage;
using AggregatedGenericResultMessage.Abstractions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.Extensions;
using DynamicExcelProvider.Helpers.DataTable;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using DynamicExcelProvider.Models.Request.Export;
using DynamicExcelProvider.WorkXCore.Helpers;
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
    /// <remarks>
    /// </remarks>
    /// =================================================================================================
    internal static class DocGenerateParserHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file bytes.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="embeddedModelCollection">Collection of embedded models.</param>
        /// <param name="availablePropInOutput">The available property in output.</param>
        /// <param name="data">The data.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        internal static IResult<byte[]> Generate(
            IReadOnlyCollection<PropModel> embeddedModelCollection,
            IReadOnlyCollection<PropTranslateModel> availablePropInOutput, IEnumerable<IEnumerable<PropNameValue>> data)
        {
            DataTableHelper.InitDataTable(embeddedModelCollection, availablePropInOutput);
            var table = DataTableHelper.CreateTableAndColumns();
            foreach (var record in data) table.AddRecordFromKnown(record);

            return Result<byte[]>.Success(Encoding.GetEncoding("iso-8859-1").GetBytes(table.ToCSV()));
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file bytes.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <typeparam name="TDataModel">Type of the data model.</typeparam>
        /// <param name="embeddedModelCollection">Collection of embedded models.</param>
        /// <param name="availablePropInOutput">The available property in output.</param>
        /// <param name="data">The data.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        internal static IResult<byte[]> Generate<TDataModel>(
            IReadOnlyCollection<PropModel> embeddedModelCollection,
            IReadOnlyCollection<PropTranslateModel> availablePropInOutput,
            IReadOnlyCollection<TDataModel> data) where TDataModel : class
        {
            DataTableHelper.InitDataTable(embeddedModelCollection, availablePropInOutput);
            var table = DataTableHelper.CreateTableAndColumns();
            foreach (var record in data) table.AddRecord(record);

            return Result<byte[]>.Success(Encoding.GetEncoding("iso-8859-1").GetBytes(table.ToCSV()));
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file bytes.
        /// </summary>
        /// <remarks>
        /// </remarks>
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
            var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(sheetName, outputProps, embeddedModelCollection, data, false);
            if (wbDef.IsSuccess.IsFalse())
                return Result<byte[]>.Failure(wbDef.GetFirstMessage());

            var document = SpreadsheetDocumentHelper.Instance.Write(ms, wbDef.Response);
            if (document.IsSuccess.IsFalse())
                return Result<byte[]>.Failure(document.Messages.FirstOrDefault()?.Message);

            return Result<byte[]>.Success(ms.ToArray());
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file.
        /// </summary>
        /// <remarks>
        /// </remarks>
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

            var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(sheetName, outputProps, embeddedModelCollection, data, false);
            if (wbDef.IsSuccess.IsFalse())
                return Result.Failure(wbDef.GetFirstMessage());

            var document = SpreadsheetDocumentHelper.Instance.Write(stream, wbDef.Response);
            if (document.IsSuccess.IsFalse())
                return Result.Failure(document.Messages.FirstOrDefault()?.Message);

            return Result.Success();
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="exportConfiguration">The export configuration.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        internal static IResult<byte[]> Generate(ExcelCollectionExportConfiguration exportConfiguration)
        {
            var infoDataModel = (IResult<(List<PropTranslateModel>, List<PropModel>)>)WorkbookParseBuildHelper.BuildAndParseInternalModelDynamic(
                exportConfiguration.Configuration, exportConfiguration.DataCollection.FirstOrDefault());
            if (infoDataModel.IsSuccess.IsFalse())
                return Result<byte[]>.Failure(infoDataModel.GetFirstMessage());

            var (outputProps, embeddedModelCollection) = infoDataModel.Response;

            var ms = new MemoryStream();

            var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(exportConfiguration.Configuration.SheetName, outputProps,
                embeddedModelCollection, exportConfiguration.DataCollection, true);
            if (wbDef.IsSuccess.IsFalse())
                return Result<byte[]>.Failure(wbDef.GetFirstMessage());

            var document = SpreadsheetDocumentHelper.Instance.Write(ms, wbDef.Response);

            if (document.IsSuccess.IsFalse())
                return Result<byte[]>.Failure(document.Messages.FirstOrDefault()?.Message);

            return Result<byte[]>.Success(ms.ToArray());
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates an excel file.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="stream">The stream.</param>
        /// <param name="exportConfiguration">The export configuration.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        internal static IResult Generate(Stream stream, ExcelCollectionExportConfiguration exportConfiguration)
        {
            var infoDataModel = (IResult<(List<PropTranslateModel>, List<PropModel>)>)WorkbookParseBuildHelper.BuildAndParseInternalModelDynamic(
                exportConfiguration.Configuration, exportConfiguration.DataCollection.FirstOrDefault());
            if (infoDataModel.IsSuccess.IsFalse())
                return Result.Failure(infoDataModel.GetFirstMessage());

            var (outputProps, embeddedModelCollection) = infoDataModel.Response;

            var wbDef = WorkbookParseBuildHelper.BuildAndParseToWorkbookDefinition(exportConfiguration.Configuration.SheetName, outputProps,
                embeddedModelCollection, exportConfiguration.DataCollection, true);
            if (wbDef.IsSuccess.IsFalse())
                return Result.Failure(wbDef.GetFirstMessage());

            var document = SpreadsheetDocumentHelper.Instance.Write(stream, wbDef.Response);

            if (document.IsSuccess.IsFalse())
                return Result.Failure(document.Messages.FirstOrDefault()?.Message);

            return Result.Success();
        }
    }
}