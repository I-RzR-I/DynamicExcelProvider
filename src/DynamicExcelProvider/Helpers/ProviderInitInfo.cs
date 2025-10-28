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
        ///     The default maximum row number.
        /// </summary>
        /// =================================================================================================
        internal static int DefaultMaxRowNumber => 1000000;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets a value indicating whether the apply maximum row number policy.
        /// </summary>
        /// <value>
        ///     True if apply maximum row number policy, false if not.
        /// </value>
        /// =================================================================================================
        internal static bool ApplyMaxRowNumberPolicy { get; private set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets the sheet maximum number of rows.
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
        /// </summary>
        /// <param name="optionValue">True to option value.</param>
        /// =================================================================================================
        internal static void SetSheetMaxNumberOfRows(int optionValue)
            => SheetMaxNumberOfRows = optionValue;
    }
}