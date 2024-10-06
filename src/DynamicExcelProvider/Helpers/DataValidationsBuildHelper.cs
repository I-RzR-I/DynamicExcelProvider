﻿// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-10-02 23:42
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-10-04 19:03
// ***********************************************************************
//  <copyright file="DataValidationsBuildHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DomainCommonExtensions.ArraysExtensions;
using DomainCommonExtensions.CommonExtensions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.Attributes;
using DynamicExcelProvider.Enums;
using DynamicExcelProvider.Mapper;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using DynamicExcelProvider.WorkXCore.Extensions;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

#endregion

// ReSharper disable PossibleMultipleEnumeration

namespace DynamicExcelProvider.Helpers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A data validations build helper.
    /// </summary>
    /// =================================================================================================
    internal static class DataValidationsBuildHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds sheet data validations.
        /// </summary>
        /// <param name="properties">[in,out] The properties info.</param>
        /// <param name="outputProps">[in,out] The output properties.</param>
        /// <returns>
        ///     The DataValidations.
        /// </returns>
        /// =================================================================================================
        internal static DataValidations BuildSheetDataValidations(ref IEnumerable<PropertyInfo> properties, ref IReadOnlyCollection<PropTranslateModel> outputProps)
        {
            try
            {
                var dataValidations = new DataValidations();

                foreach (var prop in outputProps.WithIndex())
                {
                    var propInfo = properties.FirstOrDefault(x => x.Name == prop.item.CommonName);
                    var validationAttribute = (ExcelPropValidationAttribute)propInfo!.GetCustomAttribute(typeof(ExcelPropValidationAttribute));
                    if (validationAttribute.IsNotNull())
                    {
                        var sheetColumnLetter = SpreadsheetExtensions.GetExcelColumnName(prop.index).Response;

                        var dataValidation = new DataValidation
                        {
                            SequenceOfReferences = new ListValue<StringValue>
                            {
                                InnerText = $"{sheetColumnLetter}2:{sheetColumnLetter}1048576"
                            },
                            Type = new EnumValue<DataValidationValues>(EnumMapper.MapValidation(validationAttribute.ValidationType)),
                            Operator = validationAttribute.ValidationType != ValidationType.List
                                ? new EnumValue<DataValidationOperatorValues>(EnumMapper.MapOperator(validationAttribute.OperatorType))
                                : null,
                            AllowBlank = new BooleanValue(validationAttribute.AllowEmpty.IsTrue())
                        };

                        if (validationAttribute.PromptMessage.IsNullOrEmpty().IsFalse())
                        {
                            dataValidation.Prompt = validationAttribute.PromptMessage;
                            dataValidation.ShowInputMessage = new BooleanValue(true);
                        }

                        if (validationAttribute.ErrorMessage.IsNullOrEmpty().IsFalse())
                        {
                            dataValidation.Error = validationAttribute.ErrorMessage;
                            dataValidation.ShowErrorMessage = new BooleanValue(true);
                        }

                        if (validationAttribute.MinValue.IsNull().IsFalse())
                            dataValidation.Formula1 = new Formula1($"{validationAttribute.MinValue}");

                        if (validationAttribute.MaxValue.IsNull().IsFalse())
                            dataValidation.Formula2 = new Formula2($"{validationAttribute.MaxValue}");

                        if (validationAttribute.AllowedValues.IsNullOrEmptyEnumerable().IsFalse())
                            dataValidation.Formula1 = new Formula1($"\"{validationAttribute.AllowedValues.ListToString(",")}\"");

                        if (validationAttribute.ShowListInDropDown.IsTrue())
                            dataValidation.ShowDropDown = new BooleanValue(true);

                        dataValidations.Append(dataValidation);
                    }
                }

                return dataValidations;
            }
            catch
            {
                return null;
            }
        }
    }
}