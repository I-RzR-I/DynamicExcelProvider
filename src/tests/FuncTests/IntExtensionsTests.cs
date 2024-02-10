// ***********************************************************************
//  Assembly         : RzR.Shared.Export.FuncTests
//  Author           : RzR
//  Created On       : 2024-01-28 11:52
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-01-28 11:58
// ***********************************************************************
//  <copyright file="IntExtensionsTests.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System;
using DynamicExcelProvider.WorkXCore.Extensions.DataType;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#endregion

namespace WorkXCoreFuncTests
{
    [TestClass]
    public class IntExtensionsTests
    {
        [TestMethod]
        [DataRow(0, false)]
        [DataRow(1, true)]
        [DataRow(10, true)]
        public static void CastIntToBool(int source, bool excepted)
        {
            Assert.AreEqual(source.ToBool(), excepted);
        }

        [TestMethod]
        public static void CastIntToBoolException()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => (-1).ToBool());
        }

        [TestMethod]
        [DataRow(0, false)]
        [DataRow(1, true)]
        [DataRow(10, true)]
        public static void CastNullableIntToBool(int? source, bool excepted)
        {
            Assert.AreEqual(source.ToBool(), excepted);
        }

        [TestMethod]
        public static void CastNullableIntToBoolException()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => ((int?)-1).ToBool());
        }
    }
}