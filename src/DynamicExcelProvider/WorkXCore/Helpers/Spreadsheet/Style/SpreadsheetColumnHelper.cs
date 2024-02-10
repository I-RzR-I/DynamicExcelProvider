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
        ///     Generates the columns width.
        /// </summary>
        /// <returns>
        ///     The columns.
        /// </returns>
        /// =================================================================================================
        internal Column GenerateColumns() => new() { Width = 100, CustomWidth = true };
    }
}