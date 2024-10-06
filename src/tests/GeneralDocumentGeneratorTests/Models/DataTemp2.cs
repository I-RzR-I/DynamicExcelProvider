// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2024-02-06 20:27
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 20:48
// ***********************************************************************
//  <copyright file="DataTemp2.cs" company="">
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
    public class DataTemp2
    {
        [ExcelPropName("Identificator", 1048, true, 0)]
        public int Id { get; set; }

        [ExcelPropName("IdOperatie", 1048, true, 1)]
        public int OperationId { get; set; }

        [ExcelPropName("Nume", 1048, true, 34)]
        public string Name { get; set; }

        [ExcelPropName("Cod", 1048, true, 3)] 
        public string Code { get; set; }

        [ExcelPropName("DataInceput", 1048, true, 5)]
        public DateTime StartDate { get; set; }

        [ExcelPropName("DataSfarsit", 1048, true, 6)]
        public DateTime? EndDate { get; set; }

        [ExcelPropName("DataOperatie", 1048, true, 2, "h/d/yy h:mm")]
        public DateTime? OperationDate { get; set; }

        [ExcelPropName("EsteActiv", 1048, true, 7)]
        public bool? IsActive { get; set; }

        public int? TempId { get; set; }
    }
}