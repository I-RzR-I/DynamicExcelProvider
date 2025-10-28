// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2025-10-18 00:10
// 
//  Last Modified By : RzR
//  Last Modified On : 2025-10-19 20:03
// ***********************************************************************
//  <copyright file="ExcelWriteProviderOption.cs" company="RzR SOFT & TECH">
//   Copyright © RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DomainCommonExtensions.CommonExtensions.TypeParam;
using DynamicExcelProvider.Helpers;

#endregion

namespace DynamicExcelProvider.Models.Request
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     An excel write provider option.
    /// </summary>
    /// =================================================================================================
    public class ExcelWriteProviderOption
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a value indicating whether the apply maximum row number policy.
        ///     If the rule is applied, the maximum number of rows will be 1mln.
        ///     All data sets will be sliced ​​into multiple sheets with the maximum number of rows.
        ///     <code>
        ///            Product sheet with 2 mln rows => Product_1, Product_2
        ///     </code>
        /// </summary>
        /// <value>
        ///     True if apply maximum row number policy, false if not.
        ///     Default = true.
        /// </value>
        /// =================================================================================================
        public bool ApplyMaxRowNumberPolicy { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the sheet maximum number of rows.
        /// </summary>
        /// <value>
        ///     The sheet maximum number of rows.
        ///     Default value is 1mln.
        /// </value>
        /// =================================================================================================
        public int SheetMaxNumberOfRows { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelWriteProviderOption"/> class.
        /// </summary>
        /// =================================================================================================
        public ExcelWriteProviderOption()
        {
            ApplyMaxRowNumberPolicy = true;
            SheetMaxNumberOfRows = ProviderInitInfo.DefaultMaxRowNumber;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelWriteProviderOption"/> class.
        /// </summary>
        /// <param name="applyMaxRowNumberPolicy">
        ///     True if apply maximum row number policy, false if not. Default = true.
        /// </param>
        /// <param name="sheetMaxNumberOfRows">
        ///     The sheet maximum number of rows. Default value is 1mln.
        /// </param>
        /// =================================================================================================
        public ExcelWriteProviderOption(
            bool applyMaxRowNumberPolicy, 
            int sheetMaxNumberOfRows)
        {
            ApplyMaxRowNumberPolicy = applyMaxRowNumberPolicy.IfIsNull(true);
            SheetMaxNumberOfRows = sheetMaxNumberOfRows.IfIsNull(ProviderInitInfo.DefaultMaxRowNumber);
        }
    }
}