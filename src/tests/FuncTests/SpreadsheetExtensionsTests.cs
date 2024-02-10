// ***********************************************************************
//  Assembly         : RzR.Shared.Export.FuncTests
//  Author           : RzR
//  Created On       : 2024-01-28 15:31
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-01-28 16:26
// ***********************************************************************
//  <copyright file="SpreadsheetExtensionsTests.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System.Collections.Generic;
using DynamicExcelProvider.WorkXCore.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#endregion

namespace WorkXCoreFuncTests
{
    [TestClass]
    public class SpreadsheetExtensionsTests
    {
        [TestMethod]
        public void ValidateSheetDuplicateNamesTest()
        {
            var sheets = new List<string> { "sheet1", "sheet1" };
            var validate = sheets.ValidateSheetsName();

            Assert.IsFalse(validate.IsSuccess);
        }

        [TestMethod]
        public void ValidateSheetsNameTest()
        {
            var sheets = new List<string> { "sheet1", "sheet2" };
            var validate = sheets.ValidateSheetsName();

            Assert.IsTrue(validate.IsSuccess);
        }

        [TestMethod]
        public void ValidateSheetsWrongNameTest()
        {
            var sheets = new List<string> { "sheet1!", "sheet.2$_sheet.2$_sheet.2$_sheet.2$_sheet.2$_sheet.2$" };
            var validate = sheets.ValidateSheetsName();

            Assert.IsFalse(validate.IsSuccess);
        }

        [TestMethod]
        [DataRow(0, "A")]
        [DataRow(1, "B")]
        [DataRow(702, "AAA")]
        public void GetExcelColumnNameTest(int idx, string excepted)
        {
            var name = idx.GetExcelColumnName();

            Assert.IsTrue(name.IsSuccess);
            Assert.IsTrue(name.IsSuccess);
            Assert.AreEqual(excepted, name.Response);
        }
    }
}