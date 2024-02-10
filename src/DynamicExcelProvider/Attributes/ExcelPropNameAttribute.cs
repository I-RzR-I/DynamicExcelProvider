// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-14 22:11
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:44
// ***********************************************************************
//  <copyright file="ExcelPropNameAttribute.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System;
using System.Globalization;

// ReSharper disable RedundantDefaultMemberInitializer

#endregion

namespace DynamicExcelProvider.Attributes
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Excel property name attribute.
    /// </summary>
    /// <seealso cref="T:Attribute" />
    /// =================================================================================================
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public sealed class ExcelPropNameAttribute : Attribute
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="DynamicExcelProvider.Attributes.ExcelPropNameAttribute" />
        ///     class.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="propertyName">Required. Current name of the property.</param>
        /// <param name="cultureInfoId">Required. Current culture id LCID.</param>
        /// <param name="inResult">
        ///     Required. If set to <see langword="true" />, then the property will be included in export
        ///     result; otherwise, no.
        /// </param>
        /// <param name="order">(Optional) Optional. The default value is 0.</param>
        /// <param name="formatCode">(Optional) Optional. Format code.</param>
        /// <param name="wrapText">(Optional) True if wrap text, false if not.</param>
        /// <param name="isBold">(Optional) True if this object is bold, false if not.</param>
        /// <param name="isItalic">(Optional) True if this object is italic, false if not.</param>
        /// =================================================================================================
        public ExcelPropNameAttribute(string propertyName, int cultureInfoId, bool inResult, int order = 0, 
            string formatCode = null, bool wrapText = false, bool isBold = false, bool isItalic = false)
        {
            PropertyName = propertyName;
            CultureInfo = new CultureInfo(cultureInfoId);
            Order = order;
            InResult = inResult;
            FormatCode = formatCode;
            WrapText = wrapText;
            IsBold = isBold;
            IsItalic = isItalic;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="DynamicExcelProvider.Attributes.ExcelPropNameAttribute" />
        ///     class.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="propertyName">Required. Current name of the property.</param>
        /// <param name="cultureInfoId">Required. Current culture id LCID.</param>
        /// <param name="wrapText">(Optional) True if wrap text, false if not.</param>
        /// <param name="isBold">(Optional) True if this object is bold, false if not.</param>
        /// <param name="isItalic">(Optional) True if this object is italic, false if not.</param>
        /// =================================================================================================
        public ExcelPropNameAttribute(string propertyName, int cultureInfoId, bool wrapText = false,
            bool isBold = false, bool isItalic = false)
        {
            PropertyName = propertyName;
            CultureInfo = new CultureInfo(cultureInfoId);
            Order = 0;
            InResult = false;
            WrapText = wrapText;
            IsBold = isBold;
            IsItalic = isItalic;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the localized property name.
        /// </summary>
        /// <value>
        ///     The name of the property.
        /// </value>
        /// =================================================================================================
        public string PropertyName { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets localization culture.
        /// </summary>
        /// <value>
        ///     Information describing the culture.
        /// </value>
        /// =================================================================================================
        public CultureInfo CultureInfo { get; }

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