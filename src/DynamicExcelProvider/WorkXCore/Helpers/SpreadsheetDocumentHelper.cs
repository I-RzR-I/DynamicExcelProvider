// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-12-29 18:55
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:36
// ***********************************************************************
//  <copyright file="SpreadsheetDocumentHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using AggregatedGenericResultMessage;
using AggregatedGenericResultMessage.Abstractions;
using AggregatedGenericResultMessage.Extensions.Result;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DomainCommonExtensions.ArraysExtensions;
using DomainCommonExtensions.CommonExtensions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.WorkXCore.Extensions;
using DynamicExcelProvider.WorkXCore.Helpers.Spreadsheet.Style;
using DynamicExcelProvider.WorkXCore.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// ReSharper disable ConvertToUsingDeclaration

#endregion

namespace DynamicExcelProvider.WorkXCore.Helpers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A spreadsheet document helper.
    /// </summary>
    /// =================================================================================================
    public class SpreadsheetDocumentHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) pathname of the root folder.
        /// </summary>
        /// =================================================================================================
        private static readonly string _rootFolder = Directory.GetCurrentDirectory();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     The instance.
        /// </summary>
        /// =================================================================================================
        public static SpreadsheetDocumentHelper Instance = new();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Prevents a default instance of the <see cref="SpreadsheetDocumentHelper" /> class from
        ///     being created.
        /// </summary>
        /// =================================================================================================
        private SpreadsheetDocumentHelper() { }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Writes a excel (spreadsheet, worksheet) file.
        /// </summary>
        /// <remarks>
        ///     28-Jan-24.
        /// </remarks>
        /// <param name="filePath">Full pathname of the file.</param>
        /// <param name="workBook">The work book.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        public IResult Write(string filePath, WorkbookDefinition workBook)
        {
            filePath = filePath ?? Path.Combine(_rootFolder, $"{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            return Write(fs, workBook);
        }

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
        public IResult Write(Stream stream, WorkbookDefinition workBook)
        {
            try
            {
                if (workBook.Worksheets.IsNotNull() && workBook.Worksheets.IsNullOrEmptyEnumerable().IsFalse())
                {
                    var sheetValidation = workBook.Worksheets.Select(x => x.Name).ValidateSheetsName();
                    if (sheetValidation.IsSuccess.IsFalse())
                        return Result.Failure(sheetValidation.ToBase().GetFirstMessage());

                    using (var exDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook /*, true*/))
                    {
                        // Add a WorkbookPart to the document.
                        var workbookPart = exDocument.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        // Add style
                        var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();

                        stylePart.Stylesheet = new Stylesheet(
                            SpreadsheetFontHelper.Instance.GenerateFonts(),
                            SpreadsheetFillHelper.Instance.GenerateFills(),
                            SpreadsheetBorderHelper.Instance.GenerateBorders(),
                            new SpreadsheetCellFormatHelper().GenerateCellFormats(GetAllSheetCellDefinitions(workBook.Worksheets)),
                            SpreadsheetColumnHelper.Instance.GenerateColumns());
                        stylePart.Stylesheet.Save();

                        // Save workbook part
                        workbookPart.Workbook.Save();

                        // Add Sheets to the Workbook.
                        var sheets = exDocument.WorkbookPart?.Workbook.AppendChild(new Sheets());

                        foreach (var worksheet in workBook.Worksheets.WithIndex())
                        {
                            // Add a WorksheetPart to the WorkbookPart.
                            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                            var sheetData = new SheetData();
                            worksheetPart.Worksheet = new Worksheet(sheetData);

                            var sheet = exDocument.AddWorksheet(workbookPart, worksheetPart, worksheet, sheetData);

                            if (sheet.IsSuccess.IsFalse()) return Result.Failure(sheet.GetFirstMessage());

                            sheets?.Append(sheet.Response);
                        }

                        SpreadsheetCellFormatHelper.DisposeObjects();
                    }
                }

                return Result.Success();
            }
            catch (Exception e)
            {
                return Result.Failure().WithError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets all sheet cell definitions in this collection.
        /// </summary>
        /// <param name="sheets">The sheets.</param>
        /// <returns>
        ///     An enumerator that allows foreach to be used to process all sheet cell definitions in
        ///     this collection.
        /// </returns>
        /// =================================================================================================
        private static IEnumerable<CellHeaderDefinition> GetAllSheetCellDefinitions(IEnumerable<WorksheetDefinition> sheets)
        {
            var result = new List<CellHeaderDefinition>();
            foreach (var sheet in sheets) result.AddRange(sheet.ColumnHeadings);

            return result;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Finalizes an instance of the <see cref="SpreadsheetDocumentHelper"/> class.
        /// </summary>
        /// =================================================================================================
        ~SpreadsheetDocumentHelper() => Instance = null;
    }
}