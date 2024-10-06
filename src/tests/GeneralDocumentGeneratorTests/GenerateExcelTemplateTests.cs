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
    }
}