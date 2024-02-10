// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-12 21:24
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:26
// ***********************************************************************
//  <copyright file="SpreadsheetExtensions.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using AggregatedGenericResultMessage;
using AggregatedGenericResultMessage.Abstractions;
using AggregatedGenericResultMessage.Extensions.Result;
using AggregatedGenericResultMessage.Models;
using DomainCommonExtensions.ArraysExtensions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.WorkXCore.Helpers.Resources;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

// ReSharper disable MergeIntoPattern
// ReSharper disable PossibleMultipleEnumeration

#endregion

[assembly: InternalsVisibleTo("WorkXCoreFuncTests")]

namespace DynamicExcelProvider.WorkXCore.Extensions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A spreadsheet extensions.
    /// </summary>
    /// =================================================================================================
    internal static class SpreadsheetExtensions
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Validates the sheets name described by sheetsName.
        /// </summary>
        /// <param name="sheetsName">Name of the sheets.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        internal static IResult ValidateSheetsName(this IEnumerable<string> sheetsName)
        {
            try
            {
                var duplicates = sheetsName.GetDuplicates().ToList();
                if (duplicates.Count.IsGreaterThanZero())
                    return Result.Failure(string.Format(MessagesInfo.SheetNameExist, duplicates.FirstOrDefault()))
                        .WithError(new MessageDataModel("Duplicates", duplicates.ToArray()));

                foreach (var db in sheetsName)
                {
                    var isCorrectName = ValidateSheetName(db);
                    if (isCorrectName.IsSuccess.IsFalse())
                        return Result.Failure(isCorrectName.ToBase().GetFirstMessage());
                }

                return Result.Success();
            }
            catch (Exception e)
            {
                return Result.Failure().WithError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets excel column name.
        /// </summary>
        /// <param name="columnIndex">Zero-based index of the column.</param>
        /// <returns>
        ///     The excel column name.
        /// </returns>
        /// =================================================================================================
        internal static IResult<string> GetExcelColumnName(this int columnIndex)
        {
            // A - Z
            if (columnIndex >= 0 && columnIndex <= 25) return Result<string>.Success(((char)('A' + columnIndex)).ToString());

            // AA - ZZ
            if (columnIndex >= 26 && columnIndex <= 701)
            {
                var firstChar = (char)('A' + (columnIndex / 26) - 1);
                var secondChar = (char)('A' + (columnIndex % 26));

                return Result<string>.Success(string.Format(CultureInfo.InvariantCulture, "{0}{1}", firstChar, secondChar));
            }

            // 17576
            // AAA - ZZZ
            if (columnIndex >= 702 && columnIndex <= 18277)
            {
                var fc = (columnIndex - 702) / 676;
                var sc = (columnIndex - 702) % 676 / 26;
                var tc = (columnIndex - 702) % 676 % 26;

                var firstChar = (char)('A' + fc);
                var secondChar = (char)('A' + sc);
                var thirdChar = (char)('A' + tc);

                return Result<string>.Success(string.Format(CultureInfo.InvariantCulture, "{0}{1}{2}", firstChar, secondChar, thirdChar));
            }

            return Result<string>.Failure(string.Format(MessagesInfo.ColumnReferenceOutOfRange, columnIndex));
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Validates the sheet name described by sheetName.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        private static IResult ValidateSheetName(string sheetName)
        {
            try
            {
                var validSheetName = new Regex(@"^[^'*\[\]/\\:?][^*\[\]/\\:?]{0,30}$");
                if (!validSheetName.IsMatch(sheetName))
                    return Result.Failure(string.Format(MessagesInfo.InvalidSheetName, sheetName));

                return Result.Success();
            }
            catch (Exception e)
            {
                return Result.Failure().WithError(e);
            }
        }
    }
}