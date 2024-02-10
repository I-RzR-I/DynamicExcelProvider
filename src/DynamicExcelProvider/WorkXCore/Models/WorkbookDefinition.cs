// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-12-28 23:01
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:41
// ***********************************************************************
//  <copyright file="WorkbookDefinition.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System.Collections.Generic;

#endregion

namespace DynamicExcelProvider.WorkXCore.Models
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A workbook definition. Workbook definition model.
    /// </summary>
    /// =================================================================================================
    public class WorkbookDefinition
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     The worksheets.
        /// </summary>
        /// =================================================================================================
        public IEnumerable<WorksheetDefinition> Worksheets;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="WorkbookDefinition" /> class.
        /// </summary>
        /// =================================================================================================
        public WorkbookDefinition()
        {
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="WorkbookDefinition" /> class.
        /// </summary>
        /// <param name="worksheets">The worksheets.</param>
        /// =================================================================================================
        public WorkbookDefinition(IEnumerable<WorksheetDefinition> worksheets) => Worksheets = worksheets;
    }
}