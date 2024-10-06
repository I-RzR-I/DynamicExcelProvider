// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-10-03 19:10
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-10-03 19:10
// ***********************************************************************
//  <copyright file="ValidationOperatorType.cs" company="">
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
    ///     Values that represent validation operator types.
    /// </summary>
    /// =================================================================================================
    public enum ValidationOperatorType
    {
        /// <summary>
        ///     An enum constant representing the none option.
        /// </summary>
        None,

        /// <summary>
        ///     An enum constant representing the between option.
        /// </summary>
        Between,

        /// <summary>
        ///     An enum constant representing the not between option.
        /// </summary>
        NotBetween,

        /// <summary>
        ///     An enum constant representing the equal option.
        /// </summary>
        Equal,

        /// <summary>
        ///     An enum constant representing the not equal option.
        /// </summary>
        NotEqual,

        /// <summary>
        ///     An enum constant representing the less than option.
        /// </summary>
        LessThan,

        /// <summary>
        ///     An enum constant representing the less than or equal option.
        /// </summary>
        LessThanOrEqual,

        /// <summary>
        ///     An enum constant representing the greater than option.
        /// </summary>
        GreaterThan,

        /// <summary>
        ///     An enum constant representing the greater than or equal option.
        /// </summary>
        GreaterThanOrEqual
    }
}