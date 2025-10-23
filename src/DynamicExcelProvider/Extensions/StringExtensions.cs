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

using DomainCommonExtensions.DataTypeExtensions;
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
        ///     A string extension method that avoid sheet name banned symbol.
        /// </summary>
        /// <param name="source">The source to act on.</param>
        /// <returns>
        ///     A string.
        /// </returns>
        /// =================================================================================================
        internal static string RExtAvoidSheetNameBannedSymbol(this string source)
        {
            var banned = new List<char>
            {
                '\\', '/', '*', '[', ']', ':', '?', '|', '+',
                '@', '#', '£', '^', '&', '(', ')'
            };

            foreach (var item in source.ToArray())
            {
                if (banned.Contains(item))
                    source = source.Replace(item, '\0');
            }

            return source.Replace("\0", "");
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A string extension method that truncate sheet name.
        /// </summary>
        /// <param name="source">The source to act on.</param>
        /// <param name="concatWith">(Optional)</param>
        /// <returns>
        ///     A string.
        /// </returns>
        /// =================================================================================================
        internal static string RExtTruncateSheetName(this string source, string concatWith = "")
        {
            var sheetName = source.Truncate(
                concatWith.IsPresent() && source.Length >= 31
                    ? source.Length - concatWith.Length
                    : 31);

            return concatWith.IsPresent() ? sheetName + concatWith : sheetName;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A string extension method that converts this object to a clean sheet name.
        /// </summary>
        /// <param name="source">The source to act on.</param>
        /// <param name="concatWith">(Optional)</param>
        /// <returns>
        ///     The given data converted to a string.
        /// </returns>
        /// =================================================================================================
        internal static string RExtToCleanSheetName(this string source, string concatWith = "")
            => source.RExtAvoidSheetNameBannedSymbol().RExtTruncateSheetName(concatWith.RExtAvoidSheetNameBannedSymbol());
    }
}