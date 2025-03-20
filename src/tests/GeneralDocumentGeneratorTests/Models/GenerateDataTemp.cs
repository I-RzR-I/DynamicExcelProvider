// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2024-10-06 16:53
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-10-06 16:53
// ***********************************************************************
//  <copyright file="GenerateDataTemp.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

using DynamicExcelProvider.Attributes;
using DynamicExcelProvider.Enums;

namespace GeneralDocumentGeneratorTests.Models
{
    public class GenerateDataTemp
    {
        [ExcelPropValidation(ValidationType = ValidationType.Whole,
            OperatorType = ValidationOperatorType.Between,
            //MinValue = 1,
            //MaxValue = 2, 
            ErrorMessage = "Only number", 
            PromptMessage = "Accept only digits")]
        [ExcelPropName("Identificator", 1048, true, 0)]
        public int Id { get; set; }

        [ExcelPropValidation(ValidationType = ValidationType.Whole,
            OperatorType = ValidationOperatorType.Between,
            MinValue = 1,
            MaxValue = 255, 
            ErrorMessage = "Only number from 1 - 255", 
            PromptMessage = "Accept only digits")]
        [ExcelPropName("Cod client", 1048, true, 1)]
        public int ClientCodeId { get; set; }

        [ExcelPropValidation(ValidationType.Whole, ValidationOperatorType.Equal, minValue: 15, errorMessage: "Only 15 val", promptMessage: "Accept only 15")]
        [ExcelPropName("IdOperatie", 1048, true, 2)]
        public int OperationId { get; set; }

        [ExcelPropName("Nume", 1048, true, 3)]
        public string Name { get; set; }

        [ExcelPropValidation(ValidationType.TextLength, ValidationOperatorType.LessThan, minValue: 5, promptMessage: "Code, length < 5", errorMessage: "Maximum length!!")]
        [ExcelPropName("Cod", 1048, true, 4)]
        public string Code { get; set; }

        [ExcelPropValidation(ValidationType.List, ValidationOperatorType.None, allowedValues: new[] { "Yes", "No" }, promptMessage: "Yes/No", errorMessage: "Only yes/no")]
        [ExcelPropName("EsteActiv", 1048, true, 7)]
        public bool? IsActive { get; set; }

        public int? TempId { get; set; }
    }
}