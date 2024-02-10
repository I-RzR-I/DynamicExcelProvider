// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-29 22:22
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:27
// ***********************************************************************
//  <copyright file="SpreadsheetBorderHelper.cs" company="">
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
    ///     A spreadsheet border helper.
    /// </summary>
    /// =================================================================================================
    internal class SpreadsheetBorderHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) The instance.
        /// </summary>
        /// =================================================================================================
        internal static readonly SpreadsheetBorderHelper Instance = new SpreadsheetBorderHelper();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Prevents a default instance of the <see cref="SpreadsheetBorderHelper" /> class from
        ///     being created.
        /// </summary>
        /// =================================================================================================
        private SpreadsheetBorderHelper() { }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates the borders.
        /// </summary>
        /// <returns>
        ///     The borders.
        /// </returns>
        /// =================================================================================================
        internal Borders GenerateBorders()
        {
            var borders = new Borders(
                // index 0 default
                new Border(),

                // index 1 black border
                new Border(
                    new LeftBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                    new RightBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                    new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                    new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                    new DiagonalBorder())
            );

            return borders;
        }
    }
}