// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-10-03 19:03
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-10-04 19:02
// ***********************************************************************
//  <copyright file="ExcelPropValidationAttribute.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.Enums;
using System;
using System.Collections.Generic;

#endregion

namespace DynamicExcelProvider.Attributes
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Attribute for excel property validation.
    /// </summary>
    /// <seealso cref="T:Attribute"/>
    /// =================================================================================================
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelPropValidationAttribute : Attribute
    {
        /// <inheritdoc/>
        public ExcelPropValidationAttribute(
            ValidationType validationType,
            ValidationOperatorType operatorType,
            string[] allowedValues,
            object minValue,
            object maxValue,
            string errorMessage = null,
            string promptMessage = null,
            bool allowEmpty = true
            /*bool showListInDropDown = true*/)
        {
            ValidationType = validationType;
            OperatorType = operatorType;
            AllowedValues = allowedValues;
            MinValue = minValue;
            MaxValue = maxValue;
            ErrorMessage = errorMessage;
            PromptMessage = promptMessage;
            AllowEmpty = allowEmpty;
            //ShowListInDropDown = showListInDropDown;
            ShowListInDropDown = false;
        }

        /// <inheritdoc/>
        public ExcelPropValidationAttribute(
            ValidationType validationType,
            ValidationOperatorType operatorType,
            string[] allowedValues,
            string errorMessage = null,
            string promptMessage = null,
            bool allowEmpty = true
            /*bool showListInDropDown = true*/)
        {
            ValidationType = validationType;
            OperatorType = operatorType;
            AllowedValues = allowedValues;
            ErrorMessage = errorMessage;
            PromptMessage = promptMessage;
            AllowEmpty = allowEmpty;
            //ShowListInDropDown = showListInDropDown;
            ShowListInDropDown = false;
        }

        /// <inheritdoc/>
        public ExcelPropValidationAttribute(
            ValidationType validationType,
            ValidationOperatorType operatorType,
            object minValue,
            object maxValue,
            string errorMessage = null,
            string promptMessage = null,
            bool allowEmpty = true)
        {
            ValidationType = validationType;
            OperatorType = operatorType;
            MinValue = minValue;
            MaxValue = maxValue;
            ErrorMessage = errorMessage;
            PromptMessage = promptMessage;
            AllowEmpty = allowEmpty;
        }

        /// <inheritdoc/>
        public ExcelPropValidationAttribute(
            ValidationType validationType,
            ValidationOperatorType operatorType,
            object minValue,
            string errorMessage = null,
            string promptMessage = null,
            bool allowEmpty = true)
        {
            ValidationType = validationType;
            OperatorType = operatorType;
            MinValue = minValue;
            ErrorMessage = errorMessage;
            PromptMessage = promptMessage;
            AllowEmpty = allowEmpty;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelPropValidationAttribute"/> class.
        /// </summary>
        /// =================================================================================================
        public ExcelPropValidationAttribute()
        {
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the type of the validation.
        /// </summary>
        /// <value>
        ///     The type of the validation.
        /// </value>
        /// =================================================================================================
        public ValidationType ValidationType { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the type of the operator.
        /// </summary>
        /// <value>
        ///     The type of the operator.
        /// </value>
        /// =================================================================================================
        public ValidationOperatorType OperatorType { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the allowed values.
        /// </summary>
        /// <value>
        ///     The allowed values.
        /// </value>
        /// =================================================================================================
        public IEnumerable<string> AllowedValues { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the minimum value (INT_32).
        /// </summary>
        /// <value>
        ///     The minimum value.
        /// </value>
        /// =================================================================================================
        public object MinValue { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the maximum value (INT_32).
        /// </summary>
        /// <value>
        ///     The maximum value.
        /// </value>
        /// =================================================================================================
        public object MaxValue { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a message describing the error.
        /// </summary>
        /// <value>
        ///     A message describing the error.
        /// </value>
        /// =================================================================================================
        public string ErrorMessage { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a message describing the prompt.
        /// </summary>
        /// <value>
        ///     A message describing the prompt.
        /// </value>
        /// =================================================================================================
        public string PromptMessage { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a value indicating whether we allow empty.
        /// </summary>
        /// <value>
        ///     True if allow empty, false if not.
        /// </value>
        /// =================================================================================================
        public bool AllowEmpty { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a value indicating whether the list in drop down is shown.
        /// </summary>
        /// <value>
        ///     True if show list in drop down, false if not.
        /// </value>
        /// =================================================================================================
        internal bool ShowListInDropDown { get; set; }
    }
}