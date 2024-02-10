// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-26 22:06
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:30
// ***********************************************************************
//  <copyright file="SpreadsheetCellStyleAggregatorHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.WorkXCore.Models;
using System.Collections.Generic;

#endregion

// ReSharper disable ClassNeverInstantiated.Global

namespace DynamicExcelProvider.WorkXCore.Helpers.Spreadsheet.Style
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A spreadsheet cell style aggregator helper.
    /// </summary>
    /// =================================================================================================
    internal class SpreadsheetCellStyleAggregatorHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) The instance.
        /// </summary>
        /// =================================================================================================
        internal static readonly SpreadsheetCellStyleAggregatorHelper Instance = new SpreadsheetCellStyleAggregatorHelper();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Prevents a default instance of the <see cref="SpreadsheetCellStyleAggregatorHelper" />
        ///     class from being created.
        /// </summary>
        /// =================================================================================================
        private SpreadsheetCellStyleAggregatorHelper() => FileStyle = new SpreadsheetCellStyleModel();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets the file style.
        /// </summary>
        /// <value>
        ///     The file style.
        /// </value>
        /// =================================================================================================
        private SpreadsheetCellStyleModel FileStyle { get; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates all sheet cell style.
        /// </summary>
        /// <param name="worksheets">The worksheets.</param>
        /// <returns>
        ///     All sheet cell style.
        /// </returns>
        /// =================================================================================================
        internal SpreadsheetCellStyleModel GenerateAllSheetCellStyle(IEnumerable<WorksheetDefinition> worksheets)
        {
            foreach (var sheet in worksheets) GenerateAllCellStyle(sheet.ColumnHeadings);

            return FileStyle;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates all cell style.
        /// </summary>
        /// <param name="sheetCellDefinitions">The sheet cell definitions.</param>
        /// <returns>
        ///     All cell style.
        /// </returns>
        /// =================================================================================================
        internal SpreadsheetCellStyleModel GenerateAllCellStyle(IEnumerable<CellHeaderDefinition> sheetCellDefinitions)
        {
            foreach (var sheetCellDef in sheetCellDefinitions)
            {
                FileStyle.HeaderHAligns[sheetCellDef.HorizontalCellAlignment] = true;
                FileStyle.HeaderVAligns[sheetCellDef.VerticalCellAlignment] = true;

                FileStyle.CellHAligns[sheetCellDef.CellData.HorizontalCellAlignment] = true;
                FileStyle.CellVAligns[sheetCellDef.CellData.VerticalCellAlignment] = true;
                FileStyle.CellFormats[sheetCellDef.CellData.FormatCode] = true;
                FileStyle.CellDataTypes[sheetCellDef.CellData.CellDataType] = true;
                FileStyle.SourceCellDataTypes[sheetCellDef.CellData.SourceCellDataType] = true;
            }

            return FileStyle;
        }
    }
}