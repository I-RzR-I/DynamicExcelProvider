// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-12-29 21:24
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:26
// ***********************************************************************
//  <copyright file="SpreadsheetDocumentExtensions.cs" company="">
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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DomainCommonExtensions.ArraysExtensions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.WorkXCore.Extensions.DataType;
using DynamicExcelProvider.WorkXCore.Helpers;
using DynamicExcelProvider.WorkXCore.Helpers.Spreadsheet.Style;
using DynamicExcelProvider.WorkXCore.Models;
using System;
using System.Collections.Generic;
using System.Linq;

#endregion

namespace DynamicExcelProvider.WorkXCore.Extensions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A spreadsheet document extensions.
    /// </summary>
    /// =================================================================================================
    public static class SpreadsheetDocumentExtensions
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A SpreadsheetDocument extension method that adds a worksheet.
        /// </summary>
        /// <param name="document">The document to act on.</param>
        /// <param name="workbookPart">The workbook part.</param>
        /// <param name="worksheetPart">The worksheet part.</param>
        /// <param name="sheet">The sheet.</param>
        /// <param name="sheetData">Information describing the sheet.</param>
        /// <returns>
        ///     An IResult&lt;Sheet&gt;
        /// </returns>
        /// =================================================================================================
        public static IResult<Sheet> AddWorksheet(
            this SpreadsheetDocument document, WorkbookPart workbookPart,
            WorksheetPart worksheetPart, (WorksheetDefinition, int) sheet,
            SheetData sheetData)
        {
            try
            {
                // Append a new worksheet and associate it with the workbook.
                var currentSheet = CreateDocumentSheet(document?.WorkbookPart?.GetIdOfPart(worksheetPart),
                    (uint)sheet.Item2 + 1, sheet.Item1.Name);

                if (currentSheet.IsSuccess.IsFalse())
                    return Result<Sheet>.Failure(currentSheet.GetFirstMessage());

                // Add info to current sheet
                // Set header row if exist
                var existHeader = sheet.Item1.ColumnHeadings.IsNullOrEmptyEnumerable();
                if (existHeader.IsFalse())
                {
                    var headerSetResult = SetSheetHeaderRow(sheetData, sheet.Item1.ColumnHeadings);

                    if (headerSetResult.IsSuccess.IsFalse())
                        return Result<Sheet>.Failure(headerSetResult.Messages.FirstOrDefault()?.ToString());
                }

                // Add info to current sheet
                if (sheet.Item1.Rows.IsNullOrEmptyEnumerable().IsFalse())
                {
                    var cellDataDef = sheet.Item1.ColumnHeadings.Select(x => x.CellData).ToArray();
                    var sheetBody = SetSheetBodyRows(sheetData, existHeader.IsFalse(), sheet.Item1.Rows, cellDataDef);

                    if (sheetBody.IsSuccess.IsFalse())
                        return Result<Sheet>.Failure(sheetBody.Messages.FirstOrDefault()?.ToString());
                }

                return Result<Sheet>.Success(currentSheet.Response);
            }
            catch (Exception e)
            {
                return Result<Sheet>.Failure().WithError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Creates document sheet.
        /// </summary>
        /// <param name="partId">Identifier for the part.</param>
        /// <param name="sheetId">Identifier for the sheet.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>
        ///     The new document sheet.
        /// </returns>
        /// =================================================================================================
        private static IResult<Sheet> CreateDocumentSheet(string partId, uint sheetId, string sheetName)
        {
            try
            {
                var sheet = new Sheet { Id = partId, SheetId = sheetId, Name = sheetName };

                return Result<Sheet>.Success(sheet);
            }
            catch (Exception e)
            {
                return Result<Sheet>.Failure().WithError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Sets sheet header row.
        /// </summary>
        /// <param name="sheetData">Information describing the sheet.</param>
        /// <param name="columnHeadings">The column headings.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        private static IResult SetSheetHeaderRow(SheetData sheetData, IEnumerable<CellHeaderDefinition> columnHeadings)
        {
            try
            {
                var rowIdx = 1;
                var row = new Row { RowIndex = (uint)rowIdx };
                foreach (var (cell, index) in columnHeadings.WithIndex())
                {
                    var celIdx = index.GetExcelColumnName();

                    if (celIdx.IsSuccess.IsFalse())
                        return Result<Sheet>.Failure(celIdx.GetFirstMessage());

                    var cellReference = $"{celIdx.Response}{rowIdx}";
                    var cellStyleCode = SpreadsheetCellFormatHelper.BuildFormatCellHeaderKey(cell.WrapText.ToInt(),
                        cell.VerticalCellAlignment.ToInt(),
                        cell.HorizontalCellAlignment.ToInt(), cell.IsBold.ToInt(), cell.IsItalic.ToInt());

                    var xCell = new Cell
                    {
                        CellReference = cellReference, CellValue = new CellValue(cell.Name), DataType = CellValues.String, StyleIndex = SpreadsheetCellFormatHelper.GetFormatIdByCode(cellStyleCode)
                    };

                    row.Append(xCell);
                }

                sheetData.Append(row);

                return Result.Success();
            }
            catch (Exception e)
            {
                return Result<Sheet>.Failure().WithError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Sets sheet body rows.
        /// </summary>
        /// <param name="sheetData">Information describing the sheet.</param>
        /// <param name="existHeaderRow">True to exist header row.</param>
        /// <param name="sheetRows">The sheet rows.</param>
        /// <param name="cellDataDefinitions">The cell data definitions.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        private static IResult SetSheetBodyRows(
            SheetData sheetData, bool existHeaderRow,
            IEnumerable<RowDefinition> sheetRows, CellDataDefinition[] cellDataDefinitions)
        {
            try
            {
                var rowIdx = existHeaderRow.IsTrue() ? 2 : 1;
                foreach (var (currentRow, _) in sheetRows.WithIndex())
                {
                    var rowBuild = SpreadsheetCellHelper.SerializeAsRow(currentRow.Cells, cellDataDefinitions, rowIdx);
                    if (rowBuild.IsSuccess.IsFalse())
                        return Result<Sheet>.Failure(rowBuild.GetFirstMessage());

                    sheetData.Append(rowBuild.Response);
                    rowIdx++;
                }

                return Result.Success();
            }
            catch (Exception e)
            {
                return Result<Sheet>.Failure().WithError(e);
            }
        }
    }
}