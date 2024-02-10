// ***********************************************************************
//  Assembly         : RzR.Shared.Export.FuncTests
//  Author           : RzR
//  Created On       : 2024-01-14 22:23
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-01-14 22:23
// ***********************************************************************
//  <copyright file="ExcelColumnGeneratorTest.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

using DomainCommonExtensions.CommonExtensions;
using DynamicExcelProvider.WorkXCore.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace WorkXCoreFuncTests
{
    [TestClass]
    public class ExcelColumnGeneratorTest
    {
        [DataRow(0, "A")]
        [DataRow(1, "B")]
        [DataRow(25, "Z")]
        [DataRow(26, "AA")]
        [DataRow(27, "AB")]
        [DataRow(28, "AC")]
        [DataRow(676, "ZA")]
        [DataRow(701, "ZZ")]
        [DataRow(702, "AAA")]
        [DataRow(703, "AAB")]
        [DataRow(704, "AAC")]
        [DataRow(18277, "ZZZ")]
        [TestMethod]
        public void GenerateAToZ(int index, string result)
        {
            var ex = index.GetExcelColumnName();

            Assert.IsTrue(ex.IsNotNull());
            Assert.IsTrue(ex.IsSuccess);
            Assert.AreEqual(result, ex.Response);
        }
    }
}