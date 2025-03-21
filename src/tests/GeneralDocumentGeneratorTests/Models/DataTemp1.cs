// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2024-02-05 19:41
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 20:48
// ***********************************************************************
//  <copyright file="DataTemp1.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System;
using DynamicExcelProvider.Attributes;
using DynamicExcelProvider.Enums;

#endregion

namespace GeneralDocumentGeneratorTests.Models
{
    public class DataTemp1
    {
        [ExcelPropValidation(ValidationType.Whole, ValidationOperatorType.Between, minValue: 1, maxValue: 20, errorMessage: "Only 1-20 val", promptMessage: "Accept only 1-20")]
        [ExcelPropName("Identificator", 1048, true, 0)]
        public int Id { get; set; }

        [ExcelPropValidation(ValidationType.Whole, ValidationOperatorType.Equal, minValue: 15, errorMessage: "Only 15 val", promptMessage: "Accept only 15")]
        [ExcelPropName("IdOperatie", 1048, true, 1)]
        public int OperationId { get; set; }

        [ExcelPropName("Nume", 1048, true, 34)]
        public string Name { get; set; }

        [ExcelPropValidation(ValidationType.TextLength, ValidationOperatorType.LessThan, minValue: 5, promptMessage: "Code, max length 5", errorMessage: "Maximum length!!")]
        [ExcelPropName("Cod", 1048, true, 3)]
        public string Code { get; set; }

        [ExcelPropValidation(ValidationType.Date, ValidationOperatorType.Between, minValue: "2025-03-12", maxValue: "2025-03-20", promptMessage: "Accept date only")]
        [ExcelPropName("Data Inceput", 1048, true, 5)]
        public DateTime StartDate { get; set; }

        [ExcelPropName("Data Sfarsit", 1048, true, 6)]
        public DateTime? EndDate { get; set; }

        [ExcelPropName("DataOperatie", 1048, true, 2, "h/d/yy h:mm")]
        public DateTime? OperationDate { get; set; }

        [ExcelPropValidation(ValidationType.List, ValidationOperatorType.None, allowedValues: new[] { "Yes", "No" }, promptMessage: "Yes/No", errorMessage: "Only yes/no")]
        [ExcelPropName("EsteActiv", 1048, true, 7)]
        public bool? IsActive { get; set; }

        public int? TempId { get; set; }
    }
}