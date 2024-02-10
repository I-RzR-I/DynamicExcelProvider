// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-29 22:11
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:32
// ***********************************************************************
//  <copyright file="SpreadsheetFillHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

#endregion

namespace DynamicExcelProvider.WorkXCore.Helpers.Spreadsheet.Style
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A spreadsheet fill helper.
    /// </summary>
    /// =================================================================================================
    internal class SpreadsheetFillHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) The instance.
        /// </summary>
        /// =================================================================================================
        internal static readonly SpreadsheetFillHelper Instance = new SpreadsheetFillHelper();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Prevents a default instance of the <see cref="SpreadsheetFillHelper" /> class from being
        ///     created.
        /// </summary>
        /// =================================================================================================
        private SpreadsheetFillHelper() { }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates the fills.
        /// </summary>
        /// <returns>
        ///     The fills.
        /// </returns>
        /// =================================================================================================
        internal Fills GenerateFills()
        {
            var fills = new Fills(
                // Index 0
                new Fill(new PatternFill { PatternType = PatternValues.None }),

                // Index 1
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),

                // Index 2
                new Fill(new PatternFill { PatternType = PatternValues.Gray0625 }),

                // Index 3
                new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue { Value = "66666666" } }) { PatternType = PatternValues.Solid })
            );

            return fills;
        }
    }
}