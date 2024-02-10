// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-12 18:55
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:22
// ***********************************************************************
//  <copyright file="WorksheetAlreadyExistsException.cs" company="">
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
    ///     Exception for signalling worksheet already exists errors.
    /// </summary>
    /// <seealso cref="T:Exception"/>
    /// =================================================================================================
    public class WorksheetAlreadyExistsException : Exception
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="WorksheetAlreadyExistsException" /> class.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// =================================================================================================
        public WorksheetAlreadyExistsException(string sheetName)
            : base(string.Format(MessagesInfo.SheetNameExist, sheetName))
        {
        }
    }
}