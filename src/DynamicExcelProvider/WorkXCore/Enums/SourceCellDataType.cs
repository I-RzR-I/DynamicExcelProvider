// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-13 04:00
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:21
// ***********************************************************************
//  <copyright file="SourceCellDataType.cs" company="">
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
    ///     Values that represent source cell data types.
    /// </summary>
    /// =================================================================================================
    public enum SourceCellDataType
    {
        /// <summary>An enum constant representing the date time option.</summary>
        DateTime,

        /// <summary>An enum constant representing the string option.</summary>
        String,

        /// <summary>An enum constant representing the decimal option.</summary>
        Decimal,

        /// <summary>An enum constant representing the float option.</summary>
        Float,

        /// <summary>An enum constant representing the long option.</summary>
        Long,

        /// <summary>An enum constant representing the int option.</summary>
        Int,

        /// <summary>An enum constant representing the short option.</summary>
        Short,

        /// <summary>An enum constant representing the boolean option.</summary>
        Boolean
    }
}