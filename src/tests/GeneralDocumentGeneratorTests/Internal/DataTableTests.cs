// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2025-10-26 18:10
// 
//  Last Modified By : RzR
//  Last Modified On : 2025-10-26 18:54
// ***********************************************************************
//  <copyright file="DataTableTests.cs" company="RzR SOFT & TECH">
//   Copyright © RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DomainCommonExtensions.ArraysExtensions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

// ReSharper disable PossibleMultipleEnumeration

#endregion

namespace GeneralDocumentGeneratorTests.Internal
{
    [TestClass]
    public class DataTableTests
    {
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
                /*
                var newRow = dataTable.NewRow();
                newRow["Id"] = i;
                dataTable.Rows.Add(newRow);
                */
                var b = rnd.Next() > (int.MaxValue / 2);
                var date = DateTime.Now.AddDays(i);
                dataTable.Rows.Add(i, $"Bob_{i}_Code", $"Bob_{i}_Name", date, b);
            }

            _table = dataTable;
        }

        [TestMethod]
        public void ConvertToRawData_Test()
        {
            var result = _table.ConvertToRowData();
            
            Assert.IsNotNull(result);
            Assert.IsTrue(result.IsSuccess);

            var (sheetName, ptm, pm, resultRows) = result.Response;

            Assert.IsNotNull(sheetName);
            Assert.IsTrue(ptm.IsNullOrEmptyEnumerable().IsFalse());
            Assert.IsTrue(ptm.Count() == 5);
            Assert.IsTrue(pm.IsNullOrEmptyEnumerable().IsFalse());
            Assert.IsTrue(pm.Count() == 5);
            Assert.IsTrue(resultRows.IsNullOrEmptyEnumerable().IsFalse());
            Assert.IsTrue(resultRows.Count() == 20);
        }

        [TestCleanup]
        public void CleanUp() => _table = null;
    }
}