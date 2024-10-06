// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-29 00:01
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:40
// ***********************************************************************
//  <copyright file="SpreadsheetCellStyleModel.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.WorkXCore.Enums;
using System.Collections.Generic;

#endregion

namespace DynamicExcelProvider.WorkXCore.Models
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A data Model for the spreadsheet cell style.
    /// </summary>
    /// =================================================================================================
    internal class SpreadsheetCellStyleModel
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     List of types of the cell data.
        /// </summary>
        /// <value>
        ///     A list of types of the cell data.
        /// </value>
        /// =================================================================================================
        internal Dictionary<CellDataType, bool> CellDataTypes { get; set; } = new();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     The cell formats.
        /// </summary>
        /// <value>
        ///     The cell formats.
        /// </value>
        /// =================================================================================================
        internal Dictionary<string, bool> CellFormats { get; set; } = new();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     The cell horizontal aligns.
        /// </summary>
        /// <value>
        ///     The cell horizontal aligns.
        /// </value>
        /// =================================================================================================
        internal Dictionary<HorizontalCellAlignment, bool> CellHAligns { get; set; } = new();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     The cell vertical aligns.
        /// </summary>
        /// <value>
        ///     The cell vertical aligns.
        /// </value>
        /// =================================================================================================
        internal Dictionary<VerticalCellAlignment, bool> CellVAligns { get; set; } = new();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     The header horizontal aligns.
        /// </summary>
        /// <value>
        ///     The header horizontal aligns.
        /// </value>
        /// =================================================================================================
        internal Dictionary<HorizontalCellAlignment, bool> HeaderHAligns { get; set; } = new();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     The header vertical aligns.
        /// </summary>
        /// <value>
        ///     The header vertical aligns.
        /// </value>
        /// =================================================================================================
        internal Dictionary<VerticalCellAlignment, bool> HeaderVAligns { get; set; } = new();

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     List of types of the source cell data.
        /// </summary>
        /// <value>
        ///     A list of types of the source cell data.
        /// </value>
        /// =================================================================================================
        internal Dictionary<SourceCellDataType, bool> SourceCellDataTypes { get; set; } = new();
    }
}