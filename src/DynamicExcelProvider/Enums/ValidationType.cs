// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-10-03 19:11
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-10-03 19:11
// ***********************************************************************
//  <copyright file="ValidationType.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

namespace DynamicExcelProvider.Enums
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Values that represent validation types.
    /// </summary>
    /// =================================================================================================
    public enum ValidationType
    {
        /// <summary>
        ///     An enum constant representing the none option.
        /// </summary>
        None,

        /// <summary>
        ///     An enum constant representing the whole option.
        /// </summary>
        Whole,

        /// <summary>
        ///     An enum constant representing the decimal option.
        /// </summary>
        Decimal,

        /// <summary>
        ///     An enum constant representing the list option.
        /// </summary>
        List,

        /// <summary>
        ///     An enum constant representing the date option.
        /// </summary>
        Date,

        /// <summary>
        ///     An enum constant representing the text length option.
        /// </summary>
        TextLength,

        /// <summary>
        ///     An enum constant representing the custom option.
        /// </summary>
        Custom
    }
}