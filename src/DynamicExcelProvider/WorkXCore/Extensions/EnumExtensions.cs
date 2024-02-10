// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-13 03:50
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:24
// ***********************************************************************
//  <copyright file="EnumExtensions.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DocumentFormat.OpenXml.Spreadsheet;
using DynamicExcelProvider.WorkXCore.Enums;
using System.ComponentModel;
using System.Runtime.CompilerServices;

#endregion

[assembly: InternalsVisibleTo("WorkXCoreFuncTests")]

namespace DynamicExcelProvider.WorkXCore.Extensions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     An enum extensions.
    /// </summary>
    /// =================================================================================================
    internal static class EnumExtensions
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A CellDataType extension method that map cell type to source.
        /// </summary>
        /// <param name="clientCellType">The clientCellType to act on.</param>
        /// <returns>
        ///     The CellValues.
        /// </returns>
        /// =================================================================================================
        internal static CellValues MapCellTypeToSource(this CellDataType clientCellType)
            => clientCellType switch
            {
                CellDataType.String => CellValues.String,
                CellDataType.Number => CellValues.Number,
                CellDataType.Boolean => CellValues.Boolean,
                CellDataType.Date => CellValues.Date,
                _ => throw new InvalidEnumArgumentException(nameof(clientCellType), (int)clientCellType, typeof(CellDataType))
            };

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A VerticalCellAlignment extension method that map vertical cell alignment to source.
        /// </summary>
        /// <param name="verticalCellAlignment">The verticalCellAlignment to act on.</param>
        /// <returns>
        ///     The VerticalAlignmentValues.
        /// </returns>
        /// =================================================================================================
        internal static VerticalAlignmentValues MapVerticalCellAlignmentToSource(this VerticalCellAlignment verticalCellAlignment)
            => verticalCellAlignment switch
            {
                VerticalCellAlignment.Center => VerticalAlignmentValues.Center,
                VerticalCellAlignment.Bottom => VerticalAlignmentValues.Bottom,
                VerticalCellAlignment.Top => VerticalAlignmentValues.Top,
                _ => throw new InvalidEnumArgumentException(nameof(verticalCellAlignment), (int)verticalCellAlignment, typeof(VerticalAlignmentValues))
            };

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A HorizontalCellAlignment extension method that map horizontal cell alignment to source.
        /// </summary>
        /// <param name="horizontalCellAlignment">The horizontalCellAlignment to act on.</param>
        /// <returns>
        ///     The HorizontalAlignmentValues.
        /// </returns>
        /// =================================================================================================
        internal static HorizontalAlignmentValues MapHorizontalCellAlignmentToSource(this HorizontalCellAlignment horizontalCellAlignment)
            => horizontalCellAlignment switch
            {
                HorizontalCellAlignment.Center => HorizontalAlignmentValues.Center,
                HorizontalCellAlignment.Left => HorizontalAlignmentValues.Left,
                HorizontalCellAlignment.Right => HorizontalAlignmentValues.Right,
                HorizontalCellAlignment.General => HorizontalAlignmentValues.General,
                _ => throw new InvalidEnumArgumentException(nameof(horizontalCellAlignment), (int)horizontalCellAlignment, typeof(HorizontalAlignmentValues))
            };
    }
}