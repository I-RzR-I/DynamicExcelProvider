// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-10-03 22:29
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-10-04 18:50
// ***********************************************************************
//  <copyright file="EnumMapper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DocumentFormat.OpenXml.Spreadsheet;
using DynamicExcelProvider.Enums;

#endregion

namespace DynamicExcelProvider.Mapper
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     An enum mapper.
    /// </summary>
    /// =================================================================================================
    internal static class EnumMapper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Map operator.
        /// </summary>
        /// <param name="operatorType">Type of the operator.</param>
        /// <returns>
        ///     The DataValidationOperatorValues.
        /// </returns>
        /// =================================================================================================
        internal static DataValidationOperatorValues MapOperator(ValidationOperatorType operatorType)
            => operatorType switch
            {
                ValidationOperatorType.Between => DataValidationOperatorValues.Between,
                ValidationOperatorType.Equal => DataValidationOperatorValues.Equal,
                ValidationOperatorType.GreaterThan => DataValidationOperatorValues.GreaterThan,
                ValidationOperatorType.GreaterThanOrEqual => DataValidationOperatorValues.GreaterThanOrEqual,
                ValidationOperatorType.LessThan => DataValidationOperatorValues.LessThan,
                ValidationOperatorType.NotBetween => DataValidationOperatorValues.NotBetween,
                ValidationOperatorType.LessThanOrEqual => DataValidationOperatorValues.LessThanOrEqual,
                ValidationOperatorType.NotEqual => DataValidationOperatorValues.NotEqual,
                _ => DataValidationOperatorValues.Between
            };

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Map validation.
        /// </summary>
        /// <param name="validationType">Type of the validation.</param>
        /// <returns>
        ///     The DataValidationValues.
        /// </returns>
        /// =================================================================================================
        internal static DataValidationValues MapValidation(ValidationType validationType)
            => validationType switch
            {
                ValidationType.Custom => DataValidationValues.Custom,
                ValidationType.Date => DataValidationValues.Date,
                ValidationType.Decimal => DataValidationValues.Decimal,
                ValidationType.List => DataValidationValues.List,
                ValidationType.TextLength => DataValidationValues.TextLength,
                ValidationType.Whole => DataValidationValues.Whole,
                ValidationType.None => DataValidationValues.None,
                _ => DataValidationValues.None
            };
    }
}