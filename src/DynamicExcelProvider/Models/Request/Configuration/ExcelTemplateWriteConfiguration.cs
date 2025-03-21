// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2025-03-21 11:36
// 
//  Last Modified By : RzR
//  Last Modified On : 2025-03-21 18:53
// ***********************************************************************
//  <copyright file="ExcelTemplateWriteConfiguration.cs" company="RzR SOFT & TECH">
//   Copyright © RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.Models.Request.Configuration.Template;
using DynamicExcelProvider.WorkXCore.Models;
using System.Collections.Generic;

#endregion

namespace DynamicExcelProvider.Models.Request.Configuration
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     An excel template write configuration.
    /// </summary>
    /// =================================================================================================
    public class ExcelTemplateWriteConfiguration
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the name.
        /// </summary>
        /// <value>
        ///     The name.
        /// </value>
        /// =================================================================================================
        public string SheetName { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the column headings.
        /// </summary>
        /// <value>
        ///     The column headings.
        /// </value>
        /// =================================================================================================
        public IReadOnlyCollection<CellHeaderDefinition> ColumnHeadings { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the sheet validations.
        /// </summary>
        /// <value>
        ///     The sheet validations.
        /// </value>
        /// =================================================================================================
        public IReadOnlyCollection<TemplateDataValidation> SheetValidations { get; set; }
    }
}