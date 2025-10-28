// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2025-10-27 13:10
// 
//  Last Modified By : RzR
//  Last Modified On : 2025-10-27 13:18
// ***********************************************************************
//  <copyright file="GenerateExcelFromDataTableTests.cs" company="RzR SOFT & TECH">
//   Copyright © RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

using DynamicExcelProvider;
using DynamicExcelProvider.Abstractions;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace GeneralDocumentGeneratorTests
{
    [TestClass]
    public class GenerateExcelFromDataTableTests
    {
        private IExcelWriteFactoryProvider _excelWriteFactoryProvider;
        private DataTable _table;

        [TestInitialize]
        public void InitTable()
        {
            var dataTable = new DataTable("TempTableName");
            var columns = new List<DataColumn>()
            {
                new DataColumn("Id", typeof(int)),
                new DataColumn("Code", typeof(string)),
                new DataColumn("Name", typeof(string)),
                new DataColumn("Created On", typeof(DateTime)),
                new DataColumn("Is Active", typeof(bool))
            };

            dataTable.Columns.AddRange(columns.ToArray());

            var rnd = new Random(DateTime.Now.Millisecond);

            for (var i = 0; i < 20; i++)
            {
                var b = rnd.Next() > (int.MaxValue / 2);
                var date = DateTime.Now.AddDays(i);
                dataTable.Rows.Add(i, $"Bob_{i}_Code", $"Bob_{i}_Name", date, b);
            }

            _table = dataTable;


            var services = new ServiceCollection();
            services.RegisterExcelDataSourceProvider(option =>
            {
                option.ApplyMaxRowNumberPolicy = true;
                option.SheetMaxNumberOfRows = 10;
            });
            var sp = services.BuildServiceProvider();

            _excelWriteFactoryProvider = sp.GetRequiredService<IExcelWriteFactoryProvider>();
        }
        
        [TestMethod]
        public void WriteFileFromDataTable_Test()
        {
            var data = _excelWriteFactoryProvider.Generate(_table);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"WriteFileFromDataTable_{0}_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(data.Response);

            Assert.IsNotNull(fs);
            Assert.IsNotNull(fs.Length > 0);
        }

        [TestCleanup]
        public void CleanUp() => _table = null;
    }
}