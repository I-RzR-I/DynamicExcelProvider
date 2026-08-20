// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2025-10-20 19:10
// 
//  Last Modified By : RzR
//  Last Modified On : 2025-10-20 19:27
// ***********************************************************************
//  <copyright file="StringExtensions.cs" company="RzR SOFT & TECH">
//   Copyright © RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using RzR.Extensions.Domain.Text;
using System.Collections.Generic;
using System.Linq;

#endregion

namespace DynamicExcelProvider.Extensions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A string extensions.
    /// </summary>
    /// =================================================================================================
    internal static class StringExtensions
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) The fallback sheet name used when the sanitized name is empty.
        /// </summary>
        /// =================================================================================================
        private const string DefaultSheetName = "Sheet1";

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) The symbols Excel rejects anywhere inside a sheet name.
        /// </summary>
        /// =================================================================================================
        private static readonly HashSet<char> SheetNameBannedSymbols = new HashSet<char>
        {
            '\\', '/', '*', '[', ']', ':', '?'
        };

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A string extension method that avoid sheet name banned symbol.
        ///     Only the symbols Excel actually rejects are removed, so characters such as
        ///     '|', '+', '&amp;', '(' or ')' are preserved in the caller supplied name.
        /// </summary>
        /// <param name="source">The source to act on.</param>
        /// <returns>
        ///     A string.
        /// </returns>
        /// =================================================================================================
        internal static string RExtAvoidSheetNameBannedSymbol(this string source)
        {
            if (string.IsNullOrEmpty(source))
                return string.Empty;

            var cleaned = new string(source.Where(x => !SheetNameBannedSymbols.Contains(x)).ToArray());

            return cleaned.TrimStart('\'');
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A string extension method that truncate sheet name.
        ///     The result never exceeds the Excel sheet name limit of 31 characters and always
        ///     preserves the requested suffix, truncating the source value to fit around it.
        /// </summary>
        /// <param name="source">The source to act on.</param>
        /// <param name="concatWith">(Optional) The suffix appended to the truncated source.</param>
        /// <returns>
        ///     A string.
        /// </returns>
        /// =================================================================================================
        private static string RExtTruncateSheetName(this string source, string concatWith = "")
        {
            const int maxSheetNameLength = 31;
            var value = source.IfNullThenEmpty();
            var suffix = concatWith.IfNullThenEmpty();

            if (suffix.Length >= maxSheetNameLength)
                suffix = suffix.Substring(0, maxSheetNameLength - 1);

            var budget = maxSheetNameLength - suffix.Length;

            return (value.Length > budget ? value.Substring(0, budget) : value) + suffix;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A string extension method that converts this object to a clean sheet name.
        ///     When the source sanitizes to an empty value the library default name is used instead,
        ///     so the result always satisfies the sheet name validation rules.
        /// </summary>
        /// <param name="source">The source to act on.</param>
        /// <param name="concatWith">(Optional) The suffix appended to the cleaned source.</param>
        /// <returns>
        ///     The given data converted to a string.
        /// </returns>
        /// =================================================================================================
        internal static string RExtToCleanSheetName(this string source, string concatWith = "")
        {
            var cleanName = source.RExtAvoidSheetNameBannedSymbol();

            if (string.IsNullOrEmpty(cleanName))
                cleanName = DefaultSheetName;

            return cleanName.RExtTruncateSheetName(concatWith.RExtAvoidSheetNameBannedSymbol());
        }
    }
}