// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2024-02-07 16:06
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 20:48
// ***********************************************************************
//  <copyright file="ExcelServiceDocGenerationTests.cs" company="">
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
using System.Threading.Tasks;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider;
using DynamicExcelProvider.Abstractions;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using GeneralDocumentGeneratorTests.Models;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#endregion

namespace GeneralDocumentGeneratorTests
{
    [TestClass]
    public class ExcelServiceDocGenerationTests
    {
        private IExcelWriteFactoryProvider _excelWriteFactoryProvider;

        [TestInitialize]
        public void TestInit()
        {
            var services = new ServiceCollection();
            services.RegisterExcelDataSourceProvider();
            var sp = services.BuildServiceProvider();

            _excelWriteFactoryProvider = sp.GetRequiredService<IExcelWriteFactoryProvider>();
        }

        [TestMethod]
        public async Task GenerateCsvFromKnownAsync_Test()
        {
            var data = new List<List<PropNameValue>>();

            var generalInputModel = new List<PropModel>
            {
                new PropModel
                {
                    DataType = "int",
                    IsNullable = false,
                    CommonName = "Id"
                    //Format = null
                },
                new PropModel
                {
                    DataType = "string",
                    IsNullable = false,
                    //Format = null,
                    CommonName = "Name"
                },
                new PropModel
                {
                    DataType = "date",
                    IsNullable = false,
                    //Format = "dd.MM.yyyy",
                    CommonName = "StartDate"
                },
                new PropModel
                {
                    DataType = "date",
                    IsNullable = true,
                    //Format = "dd.MM.yyyy",
                    CommonName = "EndDate"
                },
                new PropModel
                {
                    DataType = "bool",
                    IsNullable = false,
                    //Format = null,
                    CommonName = "IsActive"
                }
            };

            var translated = new List<PropTranslateModel>
            {
                new PropTranslateModel
                {
                    CommonName = "Id",
                    TranslateName = "IdInregistrare",
                    Order = 0
                },
                new PropTranslateModel
                {
                    CommonName = "Name",
                    TranslateName = "Nume",
                    Order = 1
                },
                new PropTranslateModel
                {
                    CommonName = "StartDate",
                    TranslateName = "DataInceput",
                    Order = 2
                },
                new PropTranslateModel
                {
                    CommonName = "EndDate",
                    TranslateName = "DataSfarsit",
                    Order = 4
                }
            };

            data.Add(new List<PropNameValue>
            {
                new PropNameValue { Name = "IdInregistrare", Value = 1 },
                new PropNameValue { Name = "Nume", Value = "Test name 1" },
                new PropNameValue { Name = "DataInceput", Value = $"{DateTime.Now}" },
                new PropNameValue { Name = "DataSfarsit", Value = $"{DateTime.Now.AddDays(1)}" }
            });
            data.Add(new List<PropNameValue>
            {
                new PropNameValue { Name = "IdInregistrare", Value = 2 },
                new PropNameValue { Name = "Nume", Value = "Test name 2" },
                new PropNameValue { Name = "DataInceput", Value = $"{DateTime.Now.AddDays(1)}" },
                new PropNameValue { Name = "DataSfarsit", Value = $"{DateTime.Now.AddDays(2)}" }
            });

            var dataBytes =
                await _excelWriteFactoryProvider.GenerateCsvFromKnownAsync(generalInputModel, translated, data);

            if (dataBytes.IsSuccess.IsFalse())
                throw new Exception(dataBytes.GetFirstMessage());

            Assert.IsTrue(dataBytes.IsSuccess);
            Assert.IsNotNull(dataBytes.Response);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"GenerateCsvFromKnownAsync_Test_{Guid.NewGuid():N}.csv");
            await using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(dataBytes.Response);

            Assert.IsNotNull(fs);
            Assert.IsNotNull(fs.Length > 0);
        }

        [TestMethod]
        public async Task GenerateCsvAsync_Test()
        {
            var generalInputModel = new List<PropModel>
            {
                new PropModel
                {
                    DataType = "int",
                    IsNullable = false,
                    CommonName = "Id"
                    //Format = null
                },
                new PropModel
                {
                    DataType = "string",
                    IsNullable = false,
                    //Format = null,
                    CommonName = "Name"
                },
                new PropModel
                {
                    DataType = "date",
                    IsNullable = false,
                    //Format = "dd.MM.yyyy",
                    CommonName = "StartDate"
                },
                new PropModel
                {
                    DataType = "date",
                    IsNullable = true,
                    //Format = "dd.MM.yyyy",
                    CommonName = "EndDate"
                },
                new PropModel
                {
                    DataType = "bool",
                    IsNullable = false,
                    //Format = null,
                    CommonName = "IsActive"
                }
            };

            var translated = new List<PropTranslateModel>
            {
                new PropTranslateModel
                {
                    CommonName = "Id",
                    TranslateName = "Identificator",
                    Order = 0
                },
                new PropTranslateModel
                {
                    CommonName = "Name",
                    TranslateName = "Nume",
                    Order = 1
                },
                new PropTranslateModel
                {
                    CommonName = "StartDate",
                    TranslateName = "Inceput",
                    Order = 2
                },
                new PropTranslateModel
                {
                    CommonName = "EndDate",
                    TranslateName = "Sfarsit",
                    Order = 4
                }
            };

            var records = new List<DataTempRo>
            {
                new DataTempRo
                {
                    Identificator = 1,
                    Nume = "Test name 1",
                    Activ = true,
                    Inceput = DateTime.Now,
                    Sfarsit = DateTime.Now.AddDays(1)
                },
                new DataTempRo
                {
                    Identificator = 2,
                    Nume = "Test name 2",
                    Activ = true,
                    Inceput = DateTime.Now.AddDays(1),
                    Sfarsit = DateTime.Now.AddDays(2)
                },
                new DataTempRo
                {
                    Identificator = 3,
                    Nume = "Test name 3",
                    Activ = true,
                    Inceput = DateTime.Now.AddDays(2),
                    Sfarsit = DateTime.Now.AddDays(3)
                }
            };

            var dataBytes = await _excelWriteFactoryProvider.GenerateCsvAsync(generalInputModel, translated, records);

            if (dataBytes.IsSuccess.IsFalse())
                throw new Exception(dataBytes.GetFirstMessage());

            Assert.IsTrue(dataBytes.IsSuccess);
            Assert.IsNotNull(dataBytes.Response);

            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                $"GenerateCsvAsync_Test_{Guid.NewGuid():N}.csv");
            await using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.Write(dataBytes.Response);

            Assert.IsNotNull(fs);
            Assert.IsNotNull(fs.Length > 0);
        }
    }
}