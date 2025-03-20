// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-22 21:10
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:37
// ***********************************************************************
//  <copyright file="CellDataDefinition.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DomainCommonExtensions.CommonExtensions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.WorkXCore.Enums;
using DynamicExcelProvider.WorkXCore.Helpers.Spreadsheet;
using System;
using System.Collections.Generic;

// ReSharper disable RedundantDefaultMemberInitializer
// ReSharper disable AutoPropertyCanBeMadeGetOnly.Global
// ReSharper disable MemberCanBePrivate.Global

#endregion

namespace DynamicExcelProvider.WorkXCore.Models
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A cell data definition.
    /// </summary>
    /// =================================================================================================
    public class CellDataDefinition
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) The standard formats.
        /// </summary>
        /// =================================================================================================
        public static readonly Dictionary<string, uint> StandardFormats = new()
        {
            { "General", 0 },
            { "0", 1 },
            { "0.00", 2 },
            { "#,##0", 3 },
            { "#,##0.00", 4 },
            { "0%", 9 },
            { "0.00%", 10 },
            { "0.00E+00", 11 },
            { "# ?/?", 12 },
            { "# ??/??", 13 },
            { "mm-dd-yy", 14 },
            { "d-mmm-yy", 15 },
            { "d-mmm", 16 },
            { "mmm-yy", 17 },
            { "h:mm AM/PM", 18 },
            { "h:mm:ss AM/PM", 19 },
            { "h:mm", 20 },
            { "h:mm:ss", 21 },
            { "h/d/yy h:mm", 22 },
            { "#,##0;(#,##0)", 37 },
            { "#,##0;[Red](#,##0)", 38 },
            { "#,##0.00;(#,##0.00)", 39 },
            { "#,##0.00;[Red](#,##0.00)", 40 },
            { "mm:ss", 45 },
            { "[h]:mm:ss", 46 },
            { "mmss.0", 47 },
            { "##0.0E+0", 48 },
            { "@", 49 }
        };

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) The number formats.
        /// </summary>
        /// =================================================================================================
        internal static readonly Dictionary<string, uint> NumberFormats = new()
        {
            { "General", 0 },
            { "0", 1 },
            { "0.00", 2 },
            { "#,##0", 3 },
            { "#,##0.00", 4 },
            { "0%", 9 },
            { "0.00%", 10 },
            { "0.00E+00", 11 },
            { "#,##0;(#,##0)", 37 },
            { "#,##0.00;(#,##0.00)", 39 },
            { "##0.0E+0", 48 }
        };

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) The date and time formats.
        /// </summary>
        /// =================================================================================================
        internal static readonly Dictionary<string, uint> DateFormats = new()
        {
            { "mm-dd-yy", 14 },
            { "d-mmm-yy", 15 },
            { "d-mmm", 16 },
            { "mmm-yy", 17 },
            { "h:mm AM/PM", 18 },
            { "h:mm:ss AM/PM", 19 },
            { "h:mm", 20 },
            { "h:mm:ss", 21 },
            { "h/d/yy h:mm", 22 },
            { "mm:ss", 45 },
            { "[h]:mm:ss", 46 },
            { "mmss.0", 47 },
            { "@", 49 }
        };

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     The format code.
        /// </summary>
        /// =================================================================================================
        private string _formatCode = "0";

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellDataDefinition" /> class.
        /// </summary>
        /// =================================================================================================
        public CellDataDefinition() { }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellDataDefinition" /> class.
        /// </summary>
        /// <param name="cellDataType">The type of the cell data.</param>
        /// <param name="sourceCellDataType">The type of the source cell data.</param>
        /// =================================================================================================
        public CellDataDefinition(CellDataType cellDataType, SourceCellDataType sourceCellDataType)
        {
            CellDataType = cellDataType;
            SourceCellDataType = sourceCellDataType;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellDataDefinition" /> class.
        /// </summary>
        /// <param name="cellDataType">The type of the cell data.</param>
        /// <param name="sourceCellDataType">The type of the source cell data.</param>
        /// <param name="formatCode">The format code.</param>
        /// =================================================================================================
        public CellDataDefinition(CellDataType cellDataType, SourceCellDataType sourceCellDataType, string formatCode)
        {
            CellDataType = cellDataType;
            SourceCellDataType = sourceCellDataType;
            FormatCode = formatCode;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellDataDefinition" /> class.
        /// </summary>
        /// <param name="cellDataType">The type of the cell data.</param>
        /// <param name="sourceCellDataType">The type of the source cell data.</param>
        /// <param name="horizontalCellAlignment">The horizontal cell alignment.</param>
        /// <param name="verticalCellAlignment">The vertical cell alignment.</param>
        /// =================================================================================================
        public CellDataDefinition(
            CellDataType cellDataType, SourceCellDataType sourceCellDataType,
            HorizontalCellAlignment horizontalCellAlignment, VerticalCellAlignment verticalCellAlignment)
        {
            CellDataType = cellDataType;
            SourceCellDataType = sourceCellDataType;
            HorizontalCellAlignment = horizontalCellAlignment;
            VerticalCellAlignment = verticalCellAlignment;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellDataDefinition" /> class.
        /// </summary>
        /// <param name="cellDataType">The type of the cell data.</param>
        /// <param name="sourceCellDataType">The type of the source cell data.</param>
        /// <param name="formatCode">The format code.</param>
        /// <param name="horizontalCellAlignment">The horizontal cell alignment.</param>
        /// <param name="verticalCellAlignment">The vertical cell alignment.</param>
        /// =================================================================================================
        public CellDataDefinition(
            CellDataType cellDataType, SourceCellDataType sourceCellDataType, string formatCode,
            HorizontalCellAlignment horizontalCellAlignment, VerticalCellAlignment verticalCellAlignment)
        {
            CellDataType = cellDataType;
            SourceCellDataType = sourceCellDataType;
            FormatCode = formatCode;
            HorizontalCellAlignment = horizontalCellAlignment;
            VerticalCellAlignment = verticalCellAlignment;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the format code.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">
        ///     Thrown when one or more arguments are outside the required range.
        /// </exception>
        /// <value>
        ///     The format code.
        /// </value>
        /// =================================================================================================
        public string FormatCode
        {
            get => _formatCode;
            set
            {
                if (value.IsNull())
                    _formatCode = value;
                else if (StandardFormats.ContainsKey(value).IsTrue())
                {
                    _formatCode = value;
                }
                else if(SpreadsheetCustomDataFormatHelper.CustomDataFormat.ContainsKey(value))
                {
                    _formatCode = value;
                }
                else
                {
                    SpreadsheetCustomDataFormatHelper.SetCustomDataFormat(value);
                    _formatCode = value;

                    //throw new ArgumentOutOfRangeException(string.Format(MessagesInfo.InvalidCellFormat, value));
                }
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a value indicating whether the wrap text.
        /// </summary>
        /// <value>
        ///     True if wrap text, false if not.
        /// </value>
        /// =================================================================================================
        public bool WrapText { get; set; } = false;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a value indicating whether this object is bold.
        /// </summary>
        /// <value>
        ///     True if this object is bold, false if not.
        /// </value>
        /// =================================================================================================
        public bool IsBold { get; set; } = false;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a value indicating whether this object is italic.
        /// </summary>
        /// <value>
        ///     True if this object is italic, false if not.
        /// </value>
        /// =================================================================================================
        public bool IsItalic { get; set; } = false;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the type of the cell data.
        /// </summary>
        /// <value>
        ///     The type of the cell data.
        /// </value>
        /// =================================================================================================
        public CellDataType CellDataType { get; set; } = CellDataType.String;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the type of the source cell data.
        /// </summary>
        /// <value>
        ///     The type of the source cell data.
        /// </value>
        /// =================================================================================================
        public SourceCellDataType SourceCellDataType { get; set; } = SourceCellDataType.String;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the horizontal cell alignment.
        /// </summary>
        /// <value>
        ///     The horizontal cell alignment.
        /// </value>
        /// =================================================================================================
        public HorizontalCellAlignment HorizontalCellAlignment { get; set; } = HorizontalCellAlignment.Center;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the vertical cell alignment.
        /// </summary>
        /// <value>
        ///     The vertical cell alignment.
        /// </value>
        /// =================================================================================================
        public VerticalCellAlignment VerticalCellAlignment { get; set; } = VerticalCellAlignment.Top;
    }
}