// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-15 22:43
// 
//  Last Modified By : RzR
//  Last Modified On : 2023-03-15 22:50
// ***********************************************************************
//  <copyright file="PropNameValue.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

namespace DynamicExcelProvider.Models.Request.Configuration.Property
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Property name value model.
    /// </summary>
    /// =================================================================================================
    public class PropNameValue
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Property name.
        /// </summary>
        /// <value>
        ///     The name.
        /// </value>
        /// =================================================================================================
        public string Name { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Property value.
        /// </summary>
        /// <value>
        ///     The value.
        /// </value>
        /// =================================================================================================
        public object Value { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="PropNameValue"/> class.
        /// </summary>
        /// =================================================================================================
        public PropNameValue() { }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="PropNameValue"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        /// =================================================================================================
        public PropNameValue(string name, object value)
        {
            Name = name;
            Value = value;
        }
    }
}