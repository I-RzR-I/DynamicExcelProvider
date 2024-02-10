// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2024-02-01 20:43
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 20:48
// ***********************************************************************
//  <copyright file="DataTemp.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System;
using DynamicExcelProvider.Attributes;

#endregion

namespace GeneralDocumentGeneratorTests.Models
{
    public class DataTemp
    {
        [ExcelPropName("Identificator", 1048, inResult: true, order: 0, isBold: true)]
        [ExcelPropName("Ид", 1049, inResult: true, order: 0, isBold: true)]
        [ExcelPropName("Identificato", 1040, inResult: true, order: 0, isBold: true)]
        [ExcelPropName("पहचान की", 1081, inResult: true, order: 0, isBold: true)]
        public int Id { get; set; }

        [ExcelPropName("Nume", 1048, inResult: true, order: 1, isBold: true)]
        [ExcelPropName("Название", 1049, inResult: true, order: 1, isBold: true)]
        [ExcelPropName("नाम", 1081, inResult: true, order: 1, isBold: true)]
        public string Name { get; set; }

        [ExcelPropName("DataInceput", 1048, inResult: true, order: 2, isBold: true)]
        [ExcelPropName("Начало", 1049, inResult: true, order: 3, isBold: true)]
        [ExcelPropName("आरंभ तिथि", 1081, inResult: false)]
        public DateTime StartDate { get; set; }

        [ExcelPropName("DataSfarsit", 1048, inResult: true, order: 3, formatCode: "h/d/yy h:mm", isBold: true)]
        [ExcelPropName("Конец", 1049, inResult: true, order: 4, isBold: true)]
        [ExcelPropName("दिनांक समाप्त", 1081, inResult: false)]
        public DateTime? EndDate { get; set; }

        [ExcelPropName("EsteActiv", 1048, inResult: true, order: 4, isBold: false)]
        [ExcelPropName("Активный", 1049, inResult: true, order: 2, isBold: true)]
        [ExcelPropName("यह सक्रिय है", 1081, inResult: false)]
        public bool? IsActive { get; set; }

        [ExcelPropName("TemporaneoId", 1040, inResult: true, order: 1, isBold: true)]
        [ExcelPropName("अस्थायी आईडी", 1081, inResult: true, isBold: true)]
        public int? TempId { get; set; }
    }
}