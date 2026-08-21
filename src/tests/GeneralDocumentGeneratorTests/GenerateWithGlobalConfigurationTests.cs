// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2025-10-21 19:10
// 
//  Last Modified By : RzR
//  Last Modified On : 2025-10-21 19:23
// ***********************************************************************
//  <copyright file="GenerateWithGlobalConfigurationTests.cs" company="RzR SOFT & TECH">
//   Copyright © RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider;
using DynamicExcelProvider.Abstractions;
using DynamicExcelProvider.Models.Request.Configuration;
using DynamicExcelProvider.Models.Request.Export;
using GeneralDocumentGeneratorTests.Models;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

#endregion

namespace GeneralDocumentGeneratorTests
{
    [TestClass]
    public class GenerateWithGlobalConfigurationTests
    {
        private IExcelWriteFactoryProvider _excelWriteFactoryProvider;

        [TestInitialize]
        public void TestInit()
        {
            var services = new ServiceCollection();
            services.RegisterExcelDataSourceProvider(option =>
            {
                option.ApplyMaxRowNumberPolicy = true;
                option.SheetMaxNumberOfRows = 10;
            });
            var sp = services.BuildServiceProvider();

            _excelWriteFactoryProvider = sp.GetRequiredService<IExcelWriteFactoryProvider>();
        }

        [DataRow(25)]
        [TestMethod]
        public async Task Generator_Xlsx_X_Rows_WithMaxRowInSheet_Test(int rows)
        {
            var rnd = new Random(DateTime.Now.Millisecond);

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

            var config = (new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration
                {
                    LCID = 1048,
                    SheetName = "TempSheet1"
                },
                DataCollection = records
            });
            var data = await _excelWriteFactoryProvider.GenerateAsync(config);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"Generator_Xlsx_X_Rows_WithMaxRowInSheet_Test{rows}_{Guid.NewGuid():N}.xlsx");
            await using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(data.Response);

            Assert.IsNotNull(fs);
            Assert.IsTrue(fs.Length > 0);
        }
    }
}