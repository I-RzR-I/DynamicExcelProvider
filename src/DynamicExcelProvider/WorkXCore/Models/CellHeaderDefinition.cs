// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-19 19:25
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:38
// ***********************************************************************
//  <copyright file="CellHeaderDefinition.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.WorkXCore.Enums;

#endregion

// ReSharper disable RedundantDefaultMemberInitializer

namespace DynamicExcelProvider.WorkXCore.Models
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A cell header definition.
    /// </summary>
    /// =================================================================================================
    public class CellHeaderDefinition
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellHeaderDefinition" /> class.
        /// </summary>
        /// =================================================================================================
        public CellHeaderDefinition() { }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellHeaderDefinition" /> class.
        /// </summary>
        /// <param name="colName">Name of the column.</param>
        /// <param name="cellData">Information describing the cell.</param>
        /// =================================================================================================
        public CellHeaderDefinition(string colName, CellDataDefinition cellData)
        {
            Name = colName;
            CellData = cellData;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellHeaderDefinition" /> class.
        /// </summary>
        /// <param name="colName">Name of the column.</param>
        /// <param name="isBold">True if this object is bold, false if not.</param>
        /// <param name="isItalic">True if this object is italic, false if not.</param>
        /// <param name="cellData">Information describing the cell.</param>
        /// =================================================================================================
        public CellHeaderDefinition(string colName, bool isBold, bool isItalic, CellDataDefinition cellData)
        {
            Name = colName;
            IsBold = isBold;
            IsItalic = isItalic;
            CellData = cellData;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellHeaderDefinition" /> class.
        /// </summary>
        /// <param name="colName">Name of the column.</param>
        /// <param name="isBold">True if this object is bold, false if not.</param>
        /// <param name="isItalic">True if this object is italic, false if not.</param>
        /// =================================================================================================
        public CellHeaderDefinition(string colName, bool isBold, bool isItalic)
        {
            Name = colName;
            IsBold = isBold;
            IsItalic = isItalic;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the name ot hte column.
        /// </summary>
        /// <value>
        ///     The name.
        /// </value>
        /// =================================================================================================
        public string Name { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a value indicating whether the wrap text.
        /// </summary>
        /// <value>
        ///     True if wrap text, false if not.
        /// </value>
        /// =================================================================================================
        public bool WrapText { get; set; } = false;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a value indicating whether this object is bold.
        /// </summary>
        /// <value>
        ///     True if this object is bold, false if not.
        /// </value>
        /// =================================================================================================
        public bool IsBold { get; set; } = false;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a value indicating whether this object is italic.
        /// </summary>
        /// <value>
        ///     True if this object is italic, false if not.
        /// </value>
        /// =================================================================================================
        public bool IsItalic { get; set; } = false;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the horizontal cell alignment.
        /// </summary>
        /// <value>
        ///     The horizontal cell alignment.
        /// </value>
        /// =================================================================================================
        public HorizontalCellAlignment HorizontalCellAlignment { get; set; } = HorizontalCellAlignment.Center;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the vertical cell alignment.
        /// </summary>
        /// <value>
        ///     The vertical cell alignment.
        /// </value>
        /// =================================================================================================
        public VerticalCellAlignment VerticalCellAlignment { get; set; } = VerticalCellAlignment.Top;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets information describing the cell.
        /// </summary>
        /// <value>
        ///     Information describing the cell.
        /// </value>
        /// =================================================================================================
        public CellDataDefinition CellData { get; set; }
    }
}