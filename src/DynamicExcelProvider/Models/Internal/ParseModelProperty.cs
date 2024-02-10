// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-15 22:37
// 
//  Last Modified By : RzR
//  Last Modified On : 2023-03-15 22:52
// ***********************************************************************
//  <copyright file="ParseModelProperty.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System.Globalization;
// ReSharper disable UnusedAutoPropertyAccessor.Global

#endregion

namespace DynamicExcelProvider.Models.Internal
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Parsing properties model.
    /// </summary>
    /// =================================================================================================
    public class ParseModelProperty
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     General/common/embedded property name.
        /// </summary>
        /// <value>
        ///     The name of the embedded.
        /// </value>
        /// =================================================================================================
        public string EmbeddedName { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Localization property name.
        /// </summary>
        /// <value>
        ///     The name of the current.
        /// </value>
        /// =================================================================================================
        public string CurrentName { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets localization culture.
        /// </summary>
        /// <value>
        ///     Information describing the culture.
        /// </value>
        /// =================================================================================================
        public CultureInfo CultureInfo { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets order of the property into result.
        /// </summary>
        /// <value>
        ///     The order.
        /// </value>
        /// =================================================================================================
        public int Order { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a value indicating whether the property is included in result model.
        /// </summary>
        /// <value>
        ///     <see langword="true" /> if include it; otherwise, <see langword="false" />.
        /// </value>
        /// =================================================================================================
        public bool InResult { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the format code.
        /// </summary>
        /// <value>
        ///     The format code.
        /// </value>
        /// =================================================================================================
        public string FormatCode { get; set; }

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