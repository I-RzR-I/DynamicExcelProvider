// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2024-02-01 21:04
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 20:48
// ***********************************************************************
//  <copyright file="ExcelXLSXDocGeneratorParserHelperTests.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

// ReSharper disable InconsistentNaming

#region U S A G E S

using System;
using System.Collections.Generic;
using System.IO;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.Helpers;
using DynamicExcelProvider.Models.Request.Configuration;
using DynamicExcelProvider.Models.Request.Export;
using GeneralDocumentGeneratorTests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#endregion

namespace GeneralDocumentGeneratorTests
{
    [TestClass]
    public class ExcelXLSXDocGeneratorParserHelperTests
    {
        [TestMethod]
        [DataRow(1048)]
        [DataRow(1049)]
        [DataRow(1040)]
        [DataRow(1081)]
        public void Generator_xlsx_Test_1(int lcid)
        {
            var rnd = new Random();

            var tmp = new List<DataTemp>
            {
                new DataTemp
                {
                    Id = 1,
                    Name = "Test name 1",
                    IsActive = true,
                    StartDate = DateTime.Now.AddDays(-1),
                    EndDate = DateTime.Now.AddDays(1),
                    TempId = -11
                },
                new DataTemp
                {
                    Id = 2,
                    Name = "Test name 2",
                    IsActive = true,
                    StartDate = DateTime.Now.AddDays(1),
                    EndDate = DateTime.Now.AddDays(2)
                },
                new DataTemp
                {
                    Id = 3,
                    Name = "Test name 3",
                    IsActive = false,
                    StartDate = DateTime.Now.AddDays(2),
                    EndDate = DateTime.Now.AddDays(3),
                    TempId = null
                },
                new DataTemp
                {
                    Id = 3,
                    Name = "Test name 3",
                    IsActive = null,
                    StartDate = DateTime.Now.AddDays(2),
                    EndDate = DateTime.Now.AddDays(3),
                    TempId = null
                }
            };
            var records = new List<DataTemp>();

            for (var i = 0; i < 10; i++) records.Add(tmp[rnd.Next(tmp.Count)]);

            var dataBytes = DocGenerateParserHelper.Generate(records, lcid);
            if (dataBytes.IsSuccess.IsFalse())
                throw new Exception(dataBytes.GetFirstMessage());

            Assert.IsTrue(dataBytes.IsSuccess);
            Assert.IsNotNull(dataBytes.Response);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"Generator_xlsx_Test_1_{lcid}_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(dataBytes.Response);

            Assert.IsNotNull(fs);
            Assert.IsNotNull(fs.Length > 0);
        }

        [TestMethod]
        public void Generator_xlsx_Test_2()
        {
            var rnd = new Random();

            var tmp = new List<DataTemp>
            {
                new DataTemp
                {
                    Id = 1,
                    Name = "Test name 1",
                    IsActive = true,
                    StartDate = DateTime.Now.AddDays(-1),
                    EndDate = DateTime.Now.AddDays(1),
                    TempId = -11
                },
                new DataTemp
                {
                    Id = 2,
                    Name = "Test name 2",
                    IsActive = true,
                    StartDate = DateTime.Now.AddDays(1),
                    EndDate = DateTime.Now.AddDays(2)
                },
                new DataTemp
                {
                    Id = 3,
                    Name = "Test name 3",
                    IsActive = false,
                    StartDate = DateTime.Now.AddDays(2),
                    EndDate = DateTime.Now.AddDays(3),
                    TempId = null
                },
                new DataTemp
                {
                    Id = 3,
                    Name = "Test name 3",
                    IsActive = null,
                    StartDate = DateTime.Now.AddDays(2),
                    EndDate = DateTime.Now.AddDays(3),
                    TempId = null
                }
            };
            var records = new List<DataTemp>();

            for (var i = 0; i < 10; i++) records.Add(tmp[rnd.Next(tmp.Count)]);

            var bytesArray = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration
                {
                    LCID = 1048,
                    SheetName = "TempSheet1"
                },
                DataCollection = records
            });

            if (bytesArray.IsSuccess.IsFalse())
                throw new Exception(bytesArray.GetFirstMessage());

            Assert.IsTrue(bytesArray.IsSuccess);
            Assert.IsNotNull(bytesArray.Response);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"Generator_xlsx_Test_2_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(bytesArray.Response);

            Assert.IsNotNull(fs);
            Assert.IsNotNull(fs.Length > 0);
        }

        [TestMethod]
        [DataRow(10)]
        [DataRow(100)]
        [DataRow(1000)]
        [DataRow(5000)]
        [DataRow(10000)]
        [DataRow(20000)]
        [DataRow(50000)]
        [DataRow(100000)]
        public void Generator_xlsx_X_rows_Test(int rows)
        {
            var rnd = new Random();

            var tmp = new List<DataTemp>
            {
                new DataTemp
                {
                    Id = 1,
                    Name = "Test name 1",
                    IsActive = true,
                    StartDate = DateTime.Now.AddDays(-1),
                    EndDate = DateTime.Now.AddDays(1),
                    TempId = -11
                },
                new DataTemp
                {
                    Id = 2,
                    Name = "Test name 2",
                    IsActive = true,
                    StartDate = DateTime.Now.AddDays(1),
                    EndDate = DateTime.Now.AddDays(2)
                },
                new DataTemp
                {
                    Id = 3,
                    Name = "Test name 3",
                    IsActive = false,
                    StartDate = DateTime.Now.AddDays(2),
                    EndDate = DateTime.Now.AddDays(3),
                    TempId = null
                },
                new DataTemp
                {
                    Id = 3,
                    Name = "Test name 3",
                    IsActive = null,
                    StartDate = DateTime.Now.AddDays(2),
                    EndDate = DateTime.Now.AddDays(3),
                    TempId = null
                }
            };
            var records = new List<DataTemp>();

            for (var i = 0; i < rows; i++) records.Add(tmp[rnd.Next(tmp.Count)]);

            var bytesArray = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration
                {
                    LCID = 1048,
                    SheetName = "TempSheet1"
                },
                DataCollection = records
            });

            if (bytesArray.IsSuccess.IsFalse())
                throw new Exception(bytesArray.GetFirstMessage());

            Assert.IsTrue(bytesArray.IsSuccess);
            Assert.IsNotNull(bytesArray.Response);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"Generator_xlsx_Test_x_{rows}_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(bytesArray.Response);

            Assert.IsNotNull(fs);
            Assert.IsNotNull(fs.Length > 0);
        }

        [TestMethod]
        [DataRow(10)]
        [DataRow(100)]
        [DataRow(1000)]
        [DataRow(5000)]
        [DataRow(10000)]
        [DataRow(20000)]
        [DataRow(50000)]
        [DataRow(100000)]
        public void Generator_xlsx_X_rows_10_col_Test(int rows)
        {
            var rnd = new Random();

            var tmp = new List<DataTemp1>
            {
                new DataTemp1
                {
                    Id = 1,
                    Name = "Test name 1",
                    IsActive = true,
                    StartDate = DateTime.Now.AddDays(-1),
                    EndDate = DateTime.Now.AddDays(1),
                    Code = "C!1",
                    TempId = -11,
                    OperationDate = DateTime.Now
                },
                new DataTemp1
                {
                    Id = 2,
                    Name = "Test name 2",
                    IsActive = true,
                    StartDate = DateTime.Now.AddDays(1),
                    EndDate = DateTime.Now.AddDays(2),
                    Code = "C!2",
                    OperationDate = DateTime.Now.AddDays(10)
                },
                new DataTemp1
                {
                    Id = 4,
                    Name = "Test name 3",
                    IsActive = false,
                    StartDate = DateTime.Now.AddDays(2),
                    EndDate = DateTime.Now.AddDays(3),
                    TempId = null,
                    Code = "C!4",
                    OperationDate = DateTime.Now.AddYears(-1)
                },
                new DataTemp1
                {
                    Id = 3,
                    Name = "Test name 3",
                    IsActive = null,
                    StartDate = DateTime.Now.AddDays(2),
                    EndDate = DateTime.Now.AddDays(3),
                    TempId = null,
                    Code = "C!3",
                    OperationDate = DateTime.Now.AddYears(1)
                }
            };
            var records = new List<DataTemp1>();

            for (var i = 0; i < rows; i++) records.Add(tmp[rnd.Next(tmp.Count)]);

            var bytesArray = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration
                {
                    LCID = 1048,
                    SheetName = "TempSheet1"
                },
                DataCollection = records
            });

            if (bytesArray.IsSuccess.IsFalse())
                throw new Exception(bytesArray.GetFirstMessage());

            Assert.IsTrue(bytesArray.IsSuccess);
            Assert.IsNotNull(bytesArray.Response);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"Generator_xlsx_Test_x_{rows}_10_col_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(bytesArray.Response);

            Assert.IsNotNull(fs);
            Assert.IsNotNull(fs.Length > 0);
        }
    }
}