// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-29 22:20
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:34
// ***********************************************************************
//  <copyright file="SpreadsheetFontHelper.cs" company="">
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
    ///     A spreadsheet font helper.
    /// </summary>
    /// =================================================================================================
    internal class SpreadsheetFontHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) The instance.
        /// </summary>
        /// =================================================================================================
        internal static readonly SpreadsheetFontHelper Instance = new SpreadsheetFontHelper();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Prevents a default instance of the <see cref="SpreadsheetFontHelper" /> class from being
        ///     created.
        /// </summary>
        /// =================================================================================================
        private SpreadsheetFontHelper() { }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates the fonts.
        /// </summary>
        /// <param name="fontSize">(Optional) Size of the font.</param>
        /// <returns>
        ///     The fonts.
        /// </returns>
        /// =================================================================================================
        internal Fonts GenerateFonts(double fontSize = 10D)
        {
            var fonts = new Fonts(
                // Index 0 - default
                new Font(
                    new FontSize { Val = fontSize },
                    new FontName { Val = "Arial Unicode" }
                ),

                // Index 1, default + bold
                new Font(
                    new FontSize { Val = fontSize },
                    new FontName { Val = "Arial Unicode" },
                    new Bold()
                ),

                // Index 2, default + italic
                new Font(
                    new FontSize { Val = fontSize },
                    new FontName { Val = "Arial Unicode" },
                    new Italic()
                ),

                // Index 3, default + italic + bold
                new Font(
                    new FontSize { Val = fontSize },
                    new FontName { Val = "Arial Unicode" },
                    new Bold(),
                    new Italic()
                ),

                // Index 4, Calibri
                new Font(
                    new FontSize { Val = fontSize },
                    new FontName { Val = "Calibri" }
                ),

                // Index 5, Calibri + italic + bold
                new Font(
                    new FontSize { Val = fontSize },
                    new FontName { Val = "Calibri" },
                    new Bold()
                ),

                // Index 6, Calibri + italic
                new Font(
                    new FontSize { Val = fontSize },
                    new FontName { Val = "Calibri" },
                    new Italic()
                ),

                // Index 7, Calibri + italic + bold
                new Font(
                    new FontSize { Val = fontSize },
                    new FontName { Val = "Calibri" },
                    new Bold(),
                    new Italic()
                )
            );

            return fonts;
        }
    }
}