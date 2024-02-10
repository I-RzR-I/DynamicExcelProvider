// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-15 22:43
// 
//  Last Modified By : RzR
//  Last Modified On : 2023-03-15 22:50
// ***********************************************************************
//  <copyright file="PropModel.cs" company="">
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
    ///     General/common property model.
    /// </summary>
    /// =================================================================================================
    public class PropModel
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     General/common property name.
        /// </summary>
        /// <value>
        ///     The name of the common.
        /// </value>
        /// =================================================================================================
        public string CommonName { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Property type.
        /// </summary>
        /// <value>
        ///     The type of the data.
        /// </value>
        /// =================================================================================================
        public string DataType { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Is nullable type.
        /// </summary>
        /// <value>
        ///     True if this object is nullable, false if not.
        /// </value>
        /// =================================================================================================
        public bool IsNullable { get; set; }
    }
}