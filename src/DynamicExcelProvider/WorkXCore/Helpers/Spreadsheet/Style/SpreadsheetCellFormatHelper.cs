// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-16 19:40
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:29
// ***********************************************************************
//  <copyright file="SpreadsheetCellFormatHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DocumentFormat.OpenXml.Spreadsheet;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.WorkXCore.Enums;
using DynamicExcelProvider.WorkXCore.Extensions;
using DynamicExcelProvider.WorkXCore.Extensions.DataType;
using DynamicExcelProvider.WorkXCore.Models;
using System.Collections.Generic;

#endregion

// ReSharper disable ClassNeverInstantiated.Global
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable PossibleInvalidOperationException
// ReSharper disable RedundantEmptyObjectOrCollectionInitializer

namespace DynamicExcelProvider.WorkXCore.Helpers.Spreadsheet.Style
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A spreadsheet cell format helper.
    /// </summary>
    /// =================================================================================================
    internal class SpreadsheetCellFormatHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) the instance.
        /// </summary>
        /// =================================================================================================
        internal static readonly SpreadsheetCellFormatHelper Instance = new SpreadsheetCellFormatHelper();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="SpreadsheetCellFormatHelper" /> class.
        /// </summary>
        /// =================================================================================================
        internal SpreadsheetCellFormatHelper()
        {
            CellFormatCodes = new Dictionary<string, uint>();
            CellFormats = new CellFormats();
            FileStyles = new SpreadsheetCellStyleModel();
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) the cell format codes.
        /// </summary>
        /// <value>
        ///     The cell format codes.
        /// </value>
        /// =================================================================================================
        private static Dictionary<string, uint> CellFormatCodes { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the file styles.
        /// </summary>
        /// <value>
        ///     The file styles.
        /// </value>
        /// =================================================================================================
        private static SpreadsheetCellStyleModel FileStyles { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) the cell formats.
        /// </summary>
        /// <value>
        ///     The cell formats.
        /// </value>
        /// =================================================================================================
        private static CellFormats CellFormats { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Dispose objects.
        /// </summary>
        /// =================================================================================================
        internal static void DisposeObjects()
        {
            CellFormatCodes = null;
            FileStyles = null;
            CellFormats = null;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates a cell formats (style).
        /// </summary>
        /// <param name="sheetCellDefinitions">The sheet cell definitions.</param>
        /// <returns>
        ///     The cell formats.
        /// </returns>
        /// =================================================================================================
        internal CellFormats GenerateCellFormats(IEnumerable<CellHeaderDefinition> sheetCellDefinitions)
        {
            FileStyles = SpreadsheetCellStyleAggregatorHelper.Instance.GenerateAllCellStyle(sheetCellDefinitions);

            BuildHeaderCellFormats();
            BuildAllBodyCellFormats((uint)CellFormatCodes.Count);

            return CellFormats;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds format cell header key.
        /// </summary>
        /// <param name="wrapTextIdx">Zero-based index of the wrap text.</param>
        /// <param name="verticalAlignIdx">Zero-based index of the vertical align.</param>
        /// <param name="horizontalAlignIdx">Zero-based index of the horizontal align.</param>
        /// <param name="isBoldIdx">Zero-based index of the is bold.</param>
        /// <param name="isItalicIdx">Zero-based index of the is italic.</param>
        /// <param name="fillId">(Optional) Identifier for the fill.</param>
        /// <returns>
        ///     A string.
        /// </returns>
        /// =================================================================================================
        internal static string BuildFormatCellHeaderKey(
            int wrapTextIdx, int verticalAlignIdx, int horizontalAlignIdx,
            int isBoldIdx, int isItalicIdx, int fillId = 3)
            => $"H{wrapTextIdx}{verticalAlignIdx}{horizontalAlignIdx}{isBoldIdx}{isItalicIdx}{fillId}";

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds format cell body key.
        /// </summary>
        /// <param name="wrapTextIdx">Zero-based index of the wrap text.</param>
        /// <param name="verticalAlignIdx">Zero-based index of the vertical align.</param>
        /// <param name="horizontalAlignIdx">Zero-based index of the horizontal align.</param>
        /// <param name="isBoldIdx">Zero-based index of the is bold.</param>
        /// <param name="isItalicIdx">Zero-based index of the is italic.</param>
        /// <param name="formatCodeIdx">Zero-based index of the format code.</param>
        /// <param name="cellDataTypeIdx">Zero-based index of the cell data type.</param>
        /// <param name="sourceCellDataTypeIdx">Zero-based index of the source cell data type.</param>
        /// <param name="fillId">(Optional) Identifier for the fill.</param>
        /// <returns>
        ///     A string.
        /// </returns>
        /// =================================================================================================
        internal static string BuildFormatCellBodyKey(
            int wrapTextIdx, int verticalAlignIdx, int horizontalAlignIdx,
            int isBoldIdx, int isItalicIdx, int formatCodeIdx, int cellDataTypeIdx, int sourceCellDataTypeIdx, int fillId = 0)
            => $"B{wrapTextIdx}{verticalAlignIdx}{horizontalAlignIdx}{isBoldIdx}{isItalicIdx}{fillId}{formatCodeIdx}{cellDataTypeIdx}{sourceCellDataTypeIdx}";

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets format identifier by code.
        /// </summary>
        /// <param name="formatCode">The format code.</param>
        /// <returns>
        ///     The format identifier by code.
        /// </returns>
        /// =================================================================================================
        internal static uint GetFormatIdByCode(string formatCode)
            => CellFormatCodes[formatCode];

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets cell font identifier.
        /// </summary>
        /// <param name="isBold">True if is bold, false if not.</param>
        /// <param name="isItalic">True if is italic, false if not.</param>
        /// <returns>
        ///     The cell font identifier.
        /// </returns>
        /// =================================================================================================
        private static uint GetCellFontId(bool isBold, bool isItalic)
        {
            if (isBold.IsFalse() && isItalic.IsFalse())
                return 0;
            if (isBold.IsTrue() && isItalic.IsTrue())
                return 3;
            if (isBold.IsTrue() && isItalic.IsFalse())
                return 1;
            if (isBold.IsFalse() && isItalic.IsTrue())
                return 2;

            return 0;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets cell number format identifier.
        /// </summary>
        /// <param name="formatCode">The format code.</param>
        /// <param name="cellDataType">Type of the cell data.</param>
        /// <param name="sourceCellDataType">Type of the source cell data.</param>
        /// <returns>
        ///     The cell number format identifier.
        /// </returns>
        /// =================================================================================================
        private static uint GetCellNumFormatId(string formatCode, CellDataType cellDataType, SourceCellDataType sourceCellDataType)
        {
            if (formatCode.IsNullOrEmpty())
                return 0;

            if (cellDataType == CellDataType.String && sourceCellDataType == SourceCellDataType.String)
                return 0;

            if (cellDataType == CellDataType.Number)
            {
                if (sourceCellDataType == SourceCellDataType.Int
                    || sourceCellDataType == SourceCellDataType.Short
                    || sourceCellDataType == SourceCellDataType.Long)
                    return CellDataDefinition.NumberFormats.ContainsKey(formatCode).IsFalse()
                        ? 1
                        : (uint)CellDataDefinition.NumberFormats[formatCode];

                if (sourceCellDataType == SourceCellDataType.Decimal
                    || sourceCellDataType == SourceCellDataType.Float)
                    return CellDataDefinition.NumberFormats.ContainsKey(formatCode).IsFalse()
                        ? 2
                        : (uint)CellDataDefinition.NumberFormats[formatCode];

                return CellDataDefinition.NumberFormats.ContainsKey(formatCode).IsFalse()
                    ? 1
                    : (uint)CellDataDefinition.NumberFormats[formatCode];
            }

            if (cellDataType == CellDataType.Date)
                return CellDataDefinition.DateFormats.ContainsKey(formatCode).IsFalse()
                    ? 14
                    : (uint)CellDataDefinition.DateFormats[formatCode];

            return (uint)CellDataDefinition.StandardFormats[formatCode];
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds all body cell formats.
        /// </summary>
        /// <param name="startIdx">(Optional) The start index.</param>
        /// =================================================================================================
        private void BuildAllBodyCellFormats(uint startIdx = 0)
        {
            var fillId = 0;

            for (var wrapTextIdx = 0; wrapTextIdx <= 1; wrapTextIdx++) // true or false
            {
                var verticalAlignValues = FileStyles.CellVAligns;
                foreach (var verticalAlign in verticalAlignValues)
                {
                    var horizontalAlignValues = FileStyles.CellHAligns;
                    foreach (var horizontalAlign in horizontalAlignValues)
                        for (var isBoldIdx = 0; isBoldIdx <= 1; isBoldIdx++) // true or false
                        for (var isItalicIdx = 0; isItalicIdx <= 1; isItalicIdx++) // true or false
                        {
                            var formatCodes = FileStyles.CellFormats;
                            foreach (var formatCode in formatCodes)
                            {
                                var cellDataTypeValues = FileStyles.CellDataTypes;
                                foreach (var cellDataType in cellDataTypeValues)
                                {
                                    var sourceCellDataTypeValues = FileStyles.SourceCellDataTypes;
                                    foreach (var sourceCellDataTypeValue in sourceCellDataTypeValues)
                                    {
                                        var format = new CellFormat(
                                            new Alignment
                                            {
                                                WrapText = wrapTextIdx.ToBool(),
                                                Vertical = verticalAlign.Key.MapVerticalCellAlignmentToSource(),
                                                Horizontal = horizontalAlign.Key.MapHorizontalCellAlignmentToSource()
                                            }
                                        )
                                        {
                                            FontId = GetCellFontId(isBoldIdx.ToBool(), isItalicIdx.ToBool()),
                                            FillId = (uint)fillId,
                                            ApplyNumberFormat = true,
                                            NumberFormatId = GetCellNumFormatId(
                                                formatCode.Key,
                                                cellDataType.Key,
                                                sourceCellDataTypeValue.Key)
                                        };

                                        var key = BuildFormatCellBodyKey(wrapTextIdx, verticalAlign.Key.ToInt(),
                                            horizontalAlign.Key.ToInt(), isBoldIdx, isItalicIdx,
                                            CellDataDefinition.StandardFormats[formatCode.Key],
                                            cellDataType.Key.ToInt(), sourceCellDataTypeValue.Key.ToInt(), fillId);

                                        CellFormatCodes.Add(key, startIdx);
                                        CellFormats.Append(format);

                                        startIdx++;
                                    }
                                }
                            }
                        }
                }
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds header cell formats.
        /// </summary>
        /// =================================================================================================
        private void BuildHeaderCellFormats()
        {
            var fillId = 3;
            uint idx = 0;

            for (var wrapTextIdx = 0; wrapTextIdx <= 1; wrapTextIdx++) // true or false
            {
                var verticalAlignValues = FileStyles.HeaderVAligns;
                foreach (var verticalAlign in verticalAlignValues)
                {
                    var horizontalAlignValues = FileStyles.HeaderHAligns;
                    foreach (var horizontalAlign in horizontalAlignValues)
                        for (var isBoldIdx = 0; isBoldIdx <= 1; isBoldIdx++) // true or false
                        for (var isItalicIdx = 0; isItalicIdx <= 1; isItalicIdx++) // true or false
                        {
                            var format = new CellFormat(
                                new Alignment
                                {
                                    WrapText = wrapTextIdx.ToBool(),
                                    Vertical = verticalAlign.Key.MapVerticalCellAlignmentToSource(),
                                    Horizontal = horizontalAlign.Key.MapHorizontalCellAlignmentToSource()
                                }
                            ) { FontId = GetCellFontId(isBoldIdx.ToBool(), isItalicIdx.ToBool()), FillId = (uint)fillId };

                            var key = BuildFormatCellHeaderKey(wrapTextIdx, verticalAlign.Key.ToInt(),
                                horizontalAlign.Key.ToInt(), isBoldIdx, isItalicIdx, fillId);

                            CellFormatCodes.Add(key, idx);
                            CellFormats.Append(format);

                            idx++;
                        }
                }
            }
        }
    }
}