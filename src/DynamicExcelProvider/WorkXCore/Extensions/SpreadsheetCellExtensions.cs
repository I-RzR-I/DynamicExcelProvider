// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-15 00:31
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:25
// ***********************************************************************
//  <copyright file="SpreadsheetCellExtensions.cs" company="">
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
using DocumentFormat.OpenXml.Spreadsheet;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.WorkXCore.Extensions.DataType;
using DynamicExcelProvider.WorkXCore.Helpers;
using DynamicExcelProvider.WorkXCore.Helpers.Spreadsheet;
using DynamicExcelProvider.WorkXCore.Helpers.Spreadsheet.Style;
using DynamicExcelProvider.WorkXCore.Models;
using System;

#endregion

namespace DynamicExcelProvider.WorkXCore.Extensions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A spreadsheet cell extensions.
    /// </summary>
    /// =================================================================================================
    public static class SpreadsheetCellExtensions
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A Cell extension method that build cell body.
        /// </summary>
        /// <param name="cellInit">The cellInit to act on.</param>
        /// <param name="cellDefinition">The cell definition.</param>
        /// <param name="cellDataDefinition">The cell data definition.</param>
        /// <param name="cellIndex">Zero-based index of the cell.</param>
        /// <param name="rowIndex">Zero-based index of the row.</param>
        /// <returns>
        ///     An IResult&lt;Cell&gt;
        /// </returns>
        /// =================================================================================================
        public static IResult<Cell> BuildCellBody(
            this Cell cellInit, CellValueDefinition cellDefinition,
            CellDataDefinition cellDataDefinition, int cellIndex, int rowIndex)
        {
            try
            {
                var cellIdx = cellIndex.GetExcelColumnName();
                if (cellIdx.IsSuccess.IsFalse())
                    return Result<Cell>.Failure(cellIdx.GetFirstMessage());

                var cellReference = $"{cellIdx.Response}{rowIndex}";

                cellInit.CellReference = cellReference;
                cellInit.DataType = cellDataDefinition.CellDataType.MapCellTypeToSource();

                var cellValue = SpreadsheetCellValueHelper.BuildCellValue(cellDefinition.Value, cellDefinition.DefaultValue,
                    cellDataDefinition.CellDataType, cellDataDefinition.SourceCellDataType);
                if (cellValue.IsSuccess.IsFalse())
                    return Result<Cell>.Failure(cellValue.GetFirstMessage());

                cellInit.CellValue = cellValue.Response;

                var cellStyleCode = SpreadsheetCellFormatHelper.BuildFormatCellBodyKey(
                    cellDataDefinition.WrapText.ToInt(),
                    cellDataDefinition.VerticalCellAlignment.ToInt(),
                    cellDataDefinition.HorizontalCellAlignment.ToInt(),
                    cellDataDefinition.IsBold.ToInt(),
                    cellDataDefinition.IsItalic.ToInt(),
                    GetFormatCode(cellDataDefinition.FormatCode),
                    cellDataDefinition.CellDataType.ToInt(),
                    cellDataDefinition.SourceCellDataType.ToInt());

                cellInit.StyleIndex = SpreadsheetCellFormatHelper.GetFormatIdByCode(cellStyleCode);

                return Result<Cell>.Success(cellInit);
            }
            catch (Exception e)
            {
                return Result<Cell>.Failure(e.Message).WithError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets format code.
        /// </summary>
        /// <param name="format">Describes the format to use.</param>
        /// <returns>
        ///     The format code.
        /// </returns>
        /// =================================================================================================
        private static int GetFormatCode(string format)
        {
            if (CellDataDefinition.StandardFormats.TryGetValue(format, out var code))
                return (int)code;
            else
            {
                return (int)SpreadsheetCustomDataFormatHelper.CustomDataFormat[format];
            }
        }
    }
}