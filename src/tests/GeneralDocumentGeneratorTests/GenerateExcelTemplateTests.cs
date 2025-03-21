// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2024-10-06 16:47
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-10-06 16:47
// ***********************************************************************
//  <copyright file="GenerateExcelTemplateTests.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

using DynamicExcelProvider.Helpers;
using GeneralDocumentGeneratorTests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.WorkXCore.Enums;
using DynamicExcelProvider.WorkXCore.Models;
using System.Collections.Generic;
using System.Linq;
using DynamicExcelProvider.Enums;
using DynamicExcelProvider.Models.Request.Configuration;
using DynamicExcelProvider.Models.Request.Configuration.Template;

namespace GeneralDocumentGeneratorTests
{
    [TestClass]
    public class GenerateExcelTemplateTests
    {
        [DataRow(1048)]
        [DataRow(1033)]
        [TestMethod]
        public void Generator_xlsx_Empty_Template_LCID_Test(int lcid)
        {
            var bytesArray = DocGenerateParserHelper.GenerateTemplate<DataTemp1>(lcid);

            if (bytesArray.IsSuccess.IsFalse())
                throw new Exception(bytesArray.GetFirstMessage());

            Assert.IsTrue(bytesArray.IsSuccess);
            Assert.IsNotNull(bytesArray.Response);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"Generator_xlsx_Empty_Template_LCID_{lcid}_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(bytesArray.Response);

            Assert.IsNotNull(fs);
            Assert.IsNotNull(fs.Length > 0);
        }

        [DataRow(1048)]
        [DataRow(1033)]
        [TestMethod]
        public void Generator_xlsx_Empty_Template_Stream_Test(int lcid)
        {
            var ms = new MemoryStream();
            var bytesArray = DocGenerateParserHelper.GenerateTemplate<DataTemp1>(ms, lcid);

            if (bytesArray.IsSuccess.IsFalse())
                throw new Exception(bytesArray.GetFirstMessage());

            Assert.IsTrue(bytesArray.IsSuccess);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"Generator_xlsx_Empty_Template_Stream_{lcid}_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(ms.ToArray());

            Assert.IsNotNull(fs);
            Assert.IsNotNull(fs.Length > 0);
        }

        [DataRow(1048)]
        [DataRow(1033)]
        [TestMethod]
        public void Generator_xlsx_Empty_Template_Stream_PublicSet_Test(int lcid)
        {
            var ms = new MemoryStream();
            var bytesArray = DocGenerateParserHelper.GenerateTemplate<GenerateDataTemp>(ms, lcid);

            if (bytesArray.IsSuccess.IsFalse())
                throw new Exception(bytesArray.GetFirstMessage());

            Assert.IsTrue(bytesArray.IsSuccess);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"G_xlsx_Empty_T_Stream_PublicSet_{lcid}_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(ms.ToArray());

            Assert.IsNotNull(fs);
            Assert.IsNotNull(fs.Length > 0);
        }

        [DataRow(1048)]
        [DataRow(1033)]
        [TestMethod]
        public void Generator_xlsx_Empty_Template_Stream_PublicSet_CustomFields_Test(int lcid)
        {
            var ms = new MemoryStream();
            var bytesArray = DocGenerateParserHelper.GenerateTemplate<GenerateDataTemp>(
                ms, lcid, new[]
                {
                    nameof(GenerateDataTemp.Id),
                    nameof(GenerateDataTemp.OperationId),
                    nameof(GenerateDataTemp.Name)
                });

            if (bytesArray.IsSuccess.IsFalse())
                throw new Exception(bytesArray.GetFirstMessage());

            Assert.IsTrue(bytesArray.IsSuccess);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"G_xlsx_Empty_T_Stream_PublicSet_CustomFields_{lcid}_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(ms.ToArray());

            Assert.IsNotNull(fs);
            Assert.IsNotNull(fs.Length > 0);
        }

        [TestMethod]
        public void Generate_Template_With_ExcelTemplateWriteConfiguration_Test()
        {
            var templateSheetConfig = new ExcelTemplateWriteConfiguration()
            {
                SheetName = "Sheet1",
                ColumnHeadings = new List<CellHeaderDefinition>
                {
                    new CellHeaderDefinition("Name", true, false,
                        new CellDataDefinition(CellDataType.String, SourceCellDataType.String)),
                    new CellHeaderDefinition("Count", true, true,
                        new CellDataDefinition(CellDataType.Number, SourceCellDataType.Int)),
                    new CellHeaderDefinition("Date", true, false,
                        new CellDataDefinition(CellDataType.Date, SourceCellDataType.DateTime))
                },
                SheetValidations = new List<TemplateDataValidation>()
                {
                    new TemplateDataValidation()
                    {
                        PropertyIndex = 0,
                        ValidationType = ValidationType.TextLength,
                        OperatorType = ValidationOperatorType.LessThanOrEqual,
                        MinValue = 10,
                        PromptMessage = "Text length <= 10",
                        ErrorMessage = "Max text length exceeded"
                    },
                    new TemplateDataValidation()
                    {
                        PropertyIndex = 1,
                        ValidationType = ValidationType.Whole,
                        OperatorType = ValidationOperatorType.Between,
                        MinValue = 10,
                        MaxValue = 15,
                        PromptMessage = "Int value between 10 and 15",
                        ErrorMessage = "Int value not in range",
                    },
                    new TemplateDataValidation()
                    {
                        PropertyIndex = 2,
                        ValidationType = ValidationType.Date,
                        OperatorType = ValidationOperatorType.Between,
                        MinValue = DateTime.Now.AddDays(-2).ToString("yyyy-MM-dd"),
                        MaxValue = DateTime.Now.ToString("yyyy-MM-dd"),
                        PromptMessage = "Date range between today -2 and today",
                        ErrorMessage = "Date not in range",
                    }
                }
            };

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"Generate_Template_With_ExcelTemplateWriteConfiguration_Test_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            var res = DocGenerateParserHelper.GenerateTemplate(fs, templateSheetConfig);

            if (res.IsSuccess.IsFalse())
                throw new Exception(res.Messages.FirstOrDefault()?.ToString());

            Assert.IsTrue(res.IsSuccess);
        }
    }
}