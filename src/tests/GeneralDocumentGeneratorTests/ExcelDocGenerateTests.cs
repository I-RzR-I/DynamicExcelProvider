// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2024-01-12 18:28
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 20:48
// ***********************************************************************
//  <copyright file="ExcelDocGenerateTests.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.Helpers;
using DynamicExcelProvider.WorkXCore.Enums;
using DynamicExcelProvider.WorkXCore.Helpers;
using DynamicExcelProvider.WorkXCore.Models;
using GeneralDocumentGeneratorTests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#endregion

#pragma warning disable SCS0005

namespace GeneralDocumentGeneratorTests
{
    [TestClass]
    public class ExcelDocGenerateTests
    {
        [TestMethod]
        public void CreateFileToPath_Test()
        {
            var wbd = new WorkbookDefinition
            {
                Worksheets = new List<WorksheetDefinition>
                {
                    new WorksheetDefinition
                    {
                        Name = "Sheet1",
                        ColumnHeadings = new List<CellHeaderDefinition>
                        {
                            new CellHeaderDefinition("Name", true, false,
                                new CellDataDefinition(CellDataType.String, SourceCellDataType.String)),
                            new CellHeaderDefinition("Count", true, true,
                                new CellDataDefinition(CellDataType.Number, SourceCellDataType.Int,
                                    HorizontalCellAlignment.Left, VerticalCellAlignment.Bottom)),
                            new CellHeaderDefinition("Date", true, false,
                                new CellDataDefinition(CellDataType.Date, SourceCellDataType.DateTime, "mm-dd-yy"))
                        },
                        Rows = new List<RowDefinition>
                        {
                            new RowDefinition
                            {
                                Cells = new List<CellValueDefinition>
                                {
                                    new CellValueDefinition("string 1"),
                                    new CellValueDefinition(123),
                                    new CellValueDefinition(DateTime.Now)
                                }
                            }
                        }
                    },
                    new WorksheetDefinition
                    {
                        Name = "Sheet2",
                        ColumnHeadings = new List<CellHeaderDefinition>
                        {
                            new CellHeaderDefinition("NameX", true, false,
                                new CellDataDefinition(CellDataType.String, SourceCellDataType.String)
                                    { IsBold = true, IsItalic = true }),
                            new CellHeaderDefinition("CountX", true, false,
                                new CellDataDefinition(CellDataType.Number, SourceCellDataType.Int) { IsBold = true })
                        },
                        Rows = new List<RowDefinition>
                        {
                            new RowDefinition
                            {
                                Cells = new List<CellValueDefinition>
                                {
                                    new CellValueDefinition("string 12"),
                                    new CellValueDefinition(1234)
                                }
                            }
                        }
                    }
                }
            };

            var res = SpreadsheetDocumentHelper.Instance.Write((string)null, wbd);

            if (res.IsSuccess.IsFalse())
                throw new Exception(res.Messages.FirstOrDefault()?.ToString());

            Assert.IsTrue(res.IsSuccess);
        }

        [TestMethod]
        public void CreateFileToFromStream_Test()
        {
            var wbd = new WorkbookDefinition
            {
                Worksheets = new List<WorksheetDefinition>
                {
                    new WorksheetDefinition
                    {
                        Name = "Sheet1",
                        ColumnHeadings = new List<CellHeaderDefinition>
                        {
                            new CellHeaderDefinition("Name", true, false,
                                new CellDataDefinition(CellDataType.String, SourceCellDataType.String)),
                            new CellHeaderDefinition("Count", true, true,
                                new CellDataDefinition(CellDataType.Number, SourceCellDataType.Int)),
                            new CellHeaderDefinition("Date", true, false,
                                new CellDataDefinition(CellDataType.Date, SourceCellDataType.DateTime))
                        },
                        Rows = new List<RowDefinition>
                        {
                            new RowDefinition
                            {
                                Cells = new List<CellValueDefinition>
                                {
                                    new CellValueDefinition("string 1"),
                                    new CellValueDefinition(123),
                                    new CellValueDefinition(DateTime.Now)
                                }
                            },
                            new RowDefinition
                            {
                                Cells = new List<CellValueDefinition>
                                {
                                    new CellValueDefinition("string 2"),
                                    new CellValueDefinition(225),
                                    new CellValueDefinition(DateTime.Now.AddDays(-1))
                                }
                            }
                        }
                    },
                    new WorksheetDefinition
                    {
                        Name = "Sheet2",
                        ColumnHeadings = new List<CellHeaderDefinition>
                        {
                            new CellHeaderDefinition("NameX", true, false,
                                new CellDataDefinition(CellDataType.String, SourceCellDataType.String)
                                    { IsBold = true, IsItalic = true }),
                            new CellHeaderDefinition("CountX", true, false,
                                new CellDataDefinition(CellDataType.Number, SourceCellDataType.Int) { IsBold = true })
                        },
                        Rows = new List<RowDefinition>
                        {
                            new RowDefinition
                            {
                                Cells = new List<CellValueDefinition>
                                {
                                    new CellValueDefinition("string 12"),
                                    new CellValueDefinition(1234)
                                }
                            }
                        }
                    }
                }
            };

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"CreateFileToFromStream_Test_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            var res = SpreadsheetDocumentHelper.Instance.Write(fs, wbd);

            if (res.IsSuccess.IsFalse())
                throw new Exception(res.Messages.FirstOrDefault()?.ToString());

            Assert.IsTrue(res.IsSuccess);
        }

        [TestMethod]
        public void CreateFileToFromStream_x_rows_Test()
        {
            var rows = new List<RowDefinition>();

            for (var i = 0; i < 100; i++)
                rows.Add(new RowDefinition
                {
                    Cells = new List<CellValueDefinition>
                    {
                        new CellValueDefinition($"string {i}"),
                        new CellValueDefinition(i),
                        new CellValueDefinition(DateTime.Now.AddDays(i))
                    }
                });

            var wbd = new WorkbookDefinition
            {
                Worksheets = new List<WorksheetDefinition>
                {
                    new WorksheetDefinition
                    {
                        Name = "Sheet1",
                        ColumnHeadings = new List<CellHeaderDefinition>
                        {
                            new CellHeaderDefinition("Name", true, false,
                                new CellDataDefinition(CellDataType.String, SourceCellDataType.String)),
                            new CellHeaderDefinition("Count", true, true,
                                new CellDataDefinition(CellDataType.Number, SourceCellDataType.Int)),
                            new CellHeaderDefinition("Date", true, false,
                                new CellDataDefinition(CellDataType.Date, SourceCellDataType.DateTime))
                        },
                        Rows = rows
                    }
                }
            };

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"CreateFileToFromStream_x_rows_Test_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            var res = SpreadsheetDocumentHelper.Instance.Write(fs, wbd);

            if (res.IsSuccess.IsFalse())
                throw new Exception(res.Messages.FirstOrDefault()?.ToString());

            Assert.IsTrue(res.IsSuccess);
        }

        [TestMethod]
        [DataRow(10)]
        [DataRow(100)]
        [DataRow(1000)]
        [DataRow(10000)]
        public void CreateFileToFromStream_x_rows_with_dynamic_input_Test(int rowsCount)
        {
            var rnd = new Random();
            var tmp = new List<DataTemp2>
            {
                new DataTemp2
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
                new DataTemp2
                {
                    Id = 2,
                    Name = "Test name 2",
                    IsActive = true,
                    StartDate = DateTime.Now.AddDays(1),
                    EndDate = DateTime.Now.AddDays(2),
                    Code = "C!2",
                    OperationDate = DateTime.Now.AddDays(10)
                },
                new DataTemp2
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
                new DataTemp2
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
            var records = new List<DataTemp2>();

            for (var i = 0; i < rowsCount; i++) records.Add(tmp[rnd.Next(tmp.Count)]);

            var cells = new List<CellHeaderDefinition>
            {
                new CellHeaderDefinition("Id", true, false,
                    new CellDataDefinition(CellDataType.Number, SourceCellDataType.Int)),
                new CellHeaderDefinition("Name", true, true,
                    new CellDataDefinition(CellDataType.String, SourceCellDataType.String)),
                new CellHeaderDefinition("StartDate", true, false,
                    new CellDataDefinition(CellDataType.Date, SourceCellDataType.DateTime)),
                new CellHeaderDefinition("IsActive", true, false,
                    new CellDataDefinition(CellDataType.Boolean, SourceCellDataType.Boolean))
            };

            var wbd = new WorkbookDefinition
            {
                Worksheets = new List<WorksheetDefinition>
                {
                    new WorksheetDefinition
                    {
                        Name = "Sheet1",
                        ColumnHeadings = cells,
                        Rows = WorkbookParseBuildHelper.ParseAndBuildRowsFromSource(cells, records).Response
                    }
                }
            };

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"CreateFileToFromStream_x_rows_Test_{Guid.NewGuid():N}.xlsx");
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            var res = SpreadsheetDocumentHelper.Instance.Write(fs, wbd);

            if (res.IsSuccess.IsFalse())
                throw new Exception(res.Messages.FirstOrDefault()?.ToString());

            Assert.IsTrue(res.IsSuccess);
        }
    }
}