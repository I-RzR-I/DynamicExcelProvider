// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-12-28 22:58
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:41
// ***********************************************************************
//  <copyright file="WorksheetDefinition.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

#endregion

namespace DynamicExcelProvider.WorkXCore.Models
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A worksheet definition. Worksheet definition model.
    /// </summary>
    /// =================================================================================================
    public class WorksheetDefinition
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="WorksheetDefinition" /> class.
        /// </summary>
        /// =================================================================================================
        public WorksheetDefinition()
        {
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="WorksheetDefinition" /> class.
        /// </summary>
        /// <param name="workSheetName">Name of the work sheet.</param>
        /// <param name="columnHeadings">The column headings.</param>
        /// <param name="rows">The rows.</param>
        /// =================================================================================================
        public WorksheetDefinition(
            string workSheetName, IEnumerable<CellHeaderDefinition> columnHeadings,
            IEnumerable<RowDefinition> rows)
        {
            Name = workSheetName;
            ColumnHeadings = columnHeadings;
            Rows = rows;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="WorksheetDefinition" /> class.
        /// </summary>
        /// <param name="workSheetName">Name of the work sheet.</param>
        /// <param name="rows">The rows.</param>
        /// =================================================================================================
        public WorksheetDefinition(string workSheetName, IEnumerable<RowDefinition> rows)
        {
            Name = workSheetName;
            Rows = rows;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the name.
        /// </summary>
        /// <value>
        ///     The name.
        /// </value>
        /// =================================================================================================
        public string Name { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the column headings.
        /// </summary>
        /// <value>
        ///     The column headings.
        /// </value>
        /// =================================================================================================
        public IEnumerable<CellHeaderDefinition> ColumnHeadings { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the rows.
        /// </summary>
        /// <value>
        ///     The rows.
        /// </value>
        /// =================================================================================================
        public IEnumerable<RowDefinition> Rows { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the sheet validations.
        /// </summary>
        /// <value>
        ///     The sheet validations.
        /// </value>
        /// =================================================================================================
        public DataValidations SheetValidations { get; set; }
    }
}