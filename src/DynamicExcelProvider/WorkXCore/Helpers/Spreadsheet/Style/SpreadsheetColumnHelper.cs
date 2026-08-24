// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-29 20:18
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:31
// ***********************************************************************
//  <copyright file="SpreadsheetColumnHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DocumentFormat.OpenXml.Spreadsheet;
using DynamicExcelProvider.WorkXCore.Models;
using RzR.Extensions.Domain.Primitives;
using System.Collections.Generic;

// ReSharper disable ArrangeObjectCreationWhenTypeEvident

#endregion

namespace DynamicExcelProvider.WorkXCore.Helpers.Spreadsheet.Style
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A spreadsheet style helper.
    /// </summary>
    /// =================================================================================================
    internal class SpreadsheetColumnHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) The instance.
        /// </summary>
        /// =================================================================================================
        internal static readonly SpreadsheetColumnHelper Instance = new SpreadsheetColumnHelper();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Prevents a default instance of the <see cref="SpreadsheetColumnHelper" /> class from
        ///     being created.
        /// </summary>
        /// =================================================================================================
        private SpreadsheetColumnHelper() { }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) The widest column Excel accepts, in characters of the default font. Anything
        ///     larger is clamped to this rather than written out, which also caps infinity.
        /// </summary>
        /// =================================================================================================
        private const double MaxColumnWidth = 255d;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds the column definitions for a worksheet from its header definitions.
        /// </summary>
        /// <remarks>
        ///     Only headers that declare a <see cref="CellHeaderDefinition.Width" /> produce a
        ///     <see cref="Column" />; the rest are left out so the spreadsheet application applies its own
        ///     default width. Returns <see langword="null" /> when no header declares a width, so the caller
        ///     can omit the element entirely rather than emit an empty one.
        ///     <para>
        ///         The result belongs in the worksheet, immediately before its
        ///         <see cref="DocumentFormat.OpenXml.Spreadsheet.SheetData" />. It is not a stylesheet element.
        ///     </para>
        /// </remarks>
        /// <param name="columnHeadings">The header definitions of the worksheet, in column order.</param>
        /// <returns>
        ///     The column definitions, or <see langword="null" /> when no width was declared.
        /// </returns>
        /// =================================================================================================
        internal Columns GenerateColumns(IEnumerable<CellHeaderDefinition> columnHeadings)
        {
            if (columnHeadings.IsNull()) return null;

            var columns = new Columns();
            var index = 0u;

            foreach (var heading in columnHeadings)
            {
                index++;

                if (heading.IsNull() || (heading.Width > 0).IsFalse()) continue;

                columns.Append(new Column
                {
                    Min = index,
                    Max = index,
                    Width = heading.Width > MaxColumnWidth ? MaxColumnWidth : heading.Width,
                    CustomWidth = true
                });
            }

            return columns.HasChildren ? columns : null;
        }
    }
}