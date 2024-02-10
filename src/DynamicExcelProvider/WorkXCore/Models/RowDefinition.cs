// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-12-28 23:01
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:39
// ***********************************************************************
//  <copyright file="RowDefinition.cs" company="">
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
    ///     A row definition. Row definition model.
    /// </summary>
    /// =================================================================================================
    public class RowDefinition
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="RowDefinition" /> class.
        /// </summary>
        /// =================================================================================================
        public RowDefinition()
        {
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="RowDefinition" /> class.
        /// </summary>
        /// <param name="cells">The cells.</param>
        /// =================================================================================================
        public RowDefinition(IEnumerable<CellValueDefinition> cells) => Cells = cells;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the cells.
        /// </summary>
        /// <value>
        ///     The cells.
        /// </value>
        /// =================================================================================================
        public IEnumerable<CellValueDefinition> Cells { get; set; }
    }
}