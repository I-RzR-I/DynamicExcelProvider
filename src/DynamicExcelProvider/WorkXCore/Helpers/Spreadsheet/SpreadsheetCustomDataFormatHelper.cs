// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2025-03-18 19:27
// 
//  Last Modified By : RzR
//  Last Modified On : 2025-03-20 14:56
// ***********************************************************************
//  <copyright file="SpreadsheetCustomDataFormatHelper.cs" company="RzR SOFT & TECH">
//   Copyright © RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DocumentFormat.OpenXml.Spreadsheet;
using DomainCommonExtensions.ArraysExtensions;
using DomainCommonExtensions.CommonExtensions.TypeParam;
using DomainCommonExtensions.DataTypeExtensions;
using System.Collections.Generic;
using System.Linq;

#endregion

namespace DynamicExcelProvider.WorkXCore.Helpers.Spreadsheet
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A spreadsheet custom data format helper.
    /// </summary>
    /// =================================================================================================
    internal static class SpreadsheetCustomDataFormatHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the custom data format.
        /// </summary>
        /// <value>
        ///     The custom data format.
        /// </value>
        /// =================================================================================================
        internal static Dictionary<string, uint> CustomDataFormat { get; set; } = new Dictionary<string, uint>();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Sets custom data format.
        /// </summary>
        /// <param name="format">Describes the format to use.</param>
        /// =================================================================================================
        internal static void SetCustomDataFormat(string format)
        {
            var max = CustomDataFormat.Values.OrderByDescending(u => u).FirstOrDefault();
            var next = max.IfFuncIsTrue<uint>(164, () => max == 0);

            CustomDataFormat[format] = next + 1;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Sets custom data formats.
        /// </summary>
        /// <param name="formats">The formats.</param>
        /// =================================================================================================
        internal static void SetCustomDataFormats(IEnumerable<string> formats)
        {
            var total = CustomDataFormat.Count.IfFuncIsTrue(164, () => CustomDataFormat.Count == 0);
            foreach (var format in formats.WithIndex()) 
                CustomDataFormat[format.item] = (uint)(total + format.index + 1);
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets custom format.
        /// </summary>
        /// <param name="id">The identifier.</param>
        /// <returns>
        ///     The custom format.
        /// </returns>
        /// =================================================================================================
        internal static string GetCustomFormat(uint id)
            => CustomDataFormat.ContainsValue(id) ? CustomDataFormat.FirstOrDefault(x => x.Value == id).Key : null;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets custom format.
        /// </summary>
        /// <param name="format">Describes the format to use.</param>
        /// <returns>
        ///     The custom format.
        /// </returns>
        /// =================================================================================================
        internal static string GetCustomFormat(string format)
            => CustomDataFormat.ContainsKey(format).IsTrue() ? format : null;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Generates a numbering format.
        /// </summary>
        /// <returns>
        ///     The numbering format.
        /// </returns>
        /// =================================================================================================
        internal static NumberingFormats GenerateNumberingFormat()
        {
            var numFormats = new NumberingFormats();
            foreach (var format in CustomDataFormat)
            {
                numFormats.Append(new NumberingFormat
                {
                    NumberFormatId = format.Value, 
                    FormatCode = format.Key
                });
            }

            numFormats.Count = (uint)numFormats.ChildElements.Count;

            return numFormats;
        }
    }
}