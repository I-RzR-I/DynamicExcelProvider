// ***********************************************************************
//  Assembly         : RzR.Shared.Export.FuncTests
//  Author           : RzR
//  Created On       : 2024-01-28 11:17
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-01-28 11:17
// ***********************************************************************
//  <copyright file="EnumMapTests.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

using DocumentFormat.OpenXml.Spreadsheet;
using DynamicExcelProvider.WorkXCore.Enums;
using DynamicExcelProvider.WorkXCore.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace WorkXCoreFuncTests
{
    [TestClass]
    public class EnumMapTests
    {
        [TestMethod]
        [DataRow(CellDataType.String, "str")]
        [DataRow(CellDataType.Boolean, "b")]
        [DataRow(CellDataType.Number, "n")]
        [DataRow(CellDataType.Date, "d")]
        public void MapCellTypeToSourceTests(CellDataType clientCellType, string sourceCellValue)
        {
            var result = clientCellType.MapCellTypeToSource();

            Assert.AreEqual(new CellValues(sourceCellValue), result);
        }

        [TestMethod]
        [DataRow(HorizontalCellAlignment.Center, "center")]
        [DataRow(HorizontalCellAlignment.Left, "left")]
        [DataRow(HorizontalCellAlignment.Right, "right")]
        [DataRow(HorizontalCellAlignment.General, "general")]
        public void MapHorizontalCellAlignmentToSourceTests(HorizontalCellAlignment horizontalCellAlignment, string sourceHorizontalCellAlignment)
        {
            var result = horizontalCellAlignment.MapHorizontalCellAlignmentToSource();

            Assert.AreEqual(new HorizontalAlignmentValues(sourceHorizontalCellAlignment), result);
        }

        [TestMethod]
        [DataRow(VerticalCellAlignment.Center, "center")]
        [DataRow(VerticalCellAlignment.Bottom, "bottom")]
        [DataRow(VerticalCellAlignment.Top, "top")]
        public void MapVerticalCellAlignmentToSourceTests(VerticalCellAlignment verticalCellAlignment, string sourceVerticalCellAlignment)
        {
            var result = verticalCellAlignment.MapVerticalCellAlignmentToSource();

            Assert.AreEqual(new VerticalAlignmentValues(sourceVerticalCellAlignment), result);
        }
    }
}