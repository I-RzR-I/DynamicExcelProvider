// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2025-10-19 19:10
// 
//  Last Modified By : RzR
//  Last Modified On : 2025-10-19 19:43
// ***********************************************************************
//  <copyright file="ProviderInitInfo.cs" company="RzR SOFT & TECH">
//   Copyright © RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

namespace DynamicExcelProvider.Helpers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     An excel provider initialize information helper.
    /// </summary>
    /// =================================================================================================
    internal static class ProviderInitInfo
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     The absolute maximum row number allowed on a sheet.
        ///     Excel supports 1.048.576 rows, one of them is reserved for the column headings row.
        /// </summary>
        /// =================================================================================================
        private const int AbsoluteMaxRowNumber = 1_048_575;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     The default maximum row number.
        /// </summary>
        /// =================================================================================================
        internal static int DefaultMaxRowNumber => 1000000;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets a value indicating whether the apply maximum row number policy.
        ///     The initializer is the no dependency injection default, the document generation helpers are
        ///     reachable without any registration. It must stay in sync with the defaults assigned by the
        ///     parameterless ExcelWriteProviderOption constructor.
        /// </summary>
        /// <value>
        ///     True if apply maximum row number policy, false if not.
        ///     Default = true.
        /// </value>
        /// =================================================================================================
        internal static bool ApplyMaxRowNumberPolicy { get; private set; } = true;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets the sheet maximum number of rows.
        ///     The initializer is the no dependency injection default and must stay in sync with the defaults
        ///     assigned by the parameterless ExcelWriteProviderOption constructor.
        /// </summary>
        /// <value>
        ///     The sheet maximum number of rows.
        /// </value>
        /// =================================================================================================
        public static int SheetMaxNumberOfRows { get; private set; } = DefaultMaxRowNumber;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Sets maximum row number policy rule.
        /// </summary>
        /// <param name="optionValue">True to option value.</param>
        /// =================================================================================================
        internal static void SetMaxRowNumberPolicyRule(bool optionValue)
            => ApplyMaxRowNumberPolicy = optionValue;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Sets sheet maximum number of rows.
        ///     A value lower than or equal to zero falls back to the default, a value greater than the Excel
        ///     sheet capacity is capped so the generated document stays valid.
        /// </summary>
        /// <param name="optionValue">The requested maximum number of rows per sheet.</param>
        /// =================================================================================================
        internal static void SetSheetMaxNumberOfRows(int optionValue)
        {
            if (optionValue <= 0)
            {
                SheetMaxNumberOfRows = DefaultMaxRowNumber;

                return;
            }

            SheetMaxNumberOfRows = optionValue > AbsoluteMaxRowNumber
                ? AbsoluteMaxRowNumber
                : optionValue;
        }
    }
}