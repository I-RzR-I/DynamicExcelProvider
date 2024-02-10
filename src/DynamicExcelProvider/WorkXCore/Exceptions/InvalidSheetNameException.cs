// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-11 07:13
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:22
// ***********************************************************************
//  <copyright file="InvalidSheetNameException.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.WorkXCore.Helpers.Resources;
using System;

#endregion

namespace DynamicExcelProvider.WorkXCore.Exceptions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Exception for signalling invalid sheet name errors.
    /// </summary>
    /// <seealso cref="T:Exception"/>
    /// =================================================================================================
    public class InvalidSheetNameException : Exception
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="InvalidSheetNameException" /> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// =================================================================================================
        public InvalidSheetNameException(string name)
            : base(string.Format(MessagesInfo.InvalidSheetName, name))
        {
        }
    }
}