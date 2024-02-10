// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-15 22:43
// 
//  Last Modified By : RzR
//  Last Modified On : 2023-03-15 22:50
// ***********************************************************************
//  <copyright file="PropTranslateModel.cs" company="">
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
    ///     Property translated model.
    /// </summary>
    /// =================================================================================================
    public class PropTranslateModel
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
        ///     Translated/localization property name.
        /// </summary>
        /// <value>
        ///     The name of the translate.
        /// </value>
        /// =================================================================================================
        public string TranslateName { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Result property order.
        /// </summary>
        /// <value>
        ///     The order.
        /// </value>
        /// =================================================================================================
        public int Order { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Property value result format.
        /// </summary>
        /// <value>
        ///     The format.
        /// </value>
        /// =================================================================================================
        public string Format { get; set; }

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
    }
}