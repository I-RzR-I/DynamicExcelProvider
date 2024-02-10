// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-12-28 23:34
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:20
// ***********************************************************************
//  <copyright file="CellDataType.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

namespace DynamicExcelProvider.WorkXCore.Enums
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Values that represent cell data types.
    /// </summary>
    /// =================================================================================================
    public enum CellDataType
    {
        /// <summary>An enum constant representing the boolean option.</summary>
        Boolean,

        /// <summary>An enum constant representing the date option.</summary>
        Date,

        /// <summary>An enum constant representing the number option.</summary>
        Number,

        /// <summary>An enum constant representing the string option.</summary>
        String
    }
}