// ***********************************************************************
//  Assembly         : RzR.Shared.Export.FuncTests
//  Author           : RzR
//  Created On       : 2024-01-28 11:59
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-01-28 12:03
// ***********************************************************************
//  <copyright file="BoolExtensionsTests.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.WorkXCore.Extensions.DataType;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#endregion

namespace WorkXCoreFuncTests
{
    [TestClass]
    public class BoolExtensionsTests
    {
        [TestMethod]
        [DataRow(false, 0)]
        [DataRow(true, 1)]
        public void CastBoolToInt(bool source, int excepted)
        {
            Assert.AreEqual(excepted, source.ToInt());
        }

        [TestMethod]
        [DataRow(false, 0)]
        [DataRow(true, 1)]
        public void CastNullableBoolToInt(bool? source, int excepted)
        {
            Assert.AreEqual(excepted, source.ToInt());
        }
    }
}