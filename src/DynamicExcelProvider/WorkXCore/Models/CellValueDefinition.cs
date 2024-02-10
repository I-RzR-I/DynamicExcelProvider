// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-12-28 23:01
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:39
// ***********************************************************************
//  <copyright file="CellValueDefinition.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

namespace DynamicExcelProvider.WorkXCore.Models
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A cell definition. Cell definition model.
    /// </summary>
    /// =================================================================================================
    public class CellValueDefinition
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellValueDefinition" /> class.
        /// </summary>
        /// =================================================================================================
        public CellValueDefinition()
        {
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellValueDefinition" /> class.
        /// </summary>
        /// <param name="value">The cell value.</param>
        /// =================================================================================================
        public CellValueDefinition(object value) => Value = value;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="CellValueDefinition" /> class.
        /// </summary>
        /// <param name="value">The cell value.</param>
        /// <param name="defaultValue">The default cell value.</param>
        /// =================================================================================================
        public CellValueDefinition(object value, object defaultValue)
        {
            Value = value;
            DefaultValue = defaultValue;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the cell value.
        /// </summary>
        /// <value>
        ///     The value.
        /// </value>
        /// =================================================================================================
        public object Value { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the default cell value.
        /// </summary>
        /// <value>
        ///     The default value.
        /// </value>
        /// =================================================================================================
        public object DefaultValue { get; set; }
    }
}