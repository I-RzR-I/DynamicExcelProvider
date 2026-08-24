using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DynamicExcelProvider.Attributes;
using DynamicExcelProvider.Helpers;
using DynamicExcelProvider.Models.Request.Configuration;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using DynamicExcelProvider.Models.Request.Export;
using DynamicExcelProvider.WorkXCore.Enums;
using DynamicExcelProvider.WorkXCore.Models;
using GeneralDocumentGeneratorTests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneralDocumentGeneratorTests
{
    [TestClass]
    public class GeneratedDocumentContentTests
    {
        private const int NoSplitMaxRows = 1_000;

        private const string GeneratedOutputFolder = "GeneratedOutput";

        private readonly List<string> _writtenFiles = new List<string>();

        public TestContext TestContext { get; set; }

        [TestInitialize]
        public void ApplyExplicitProviderConfiguration()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);
            _writtenFiles.Clear();
        }

        [TestCleanup]
        public void RestoreProviderConfiguration()
        {
            ConfigureRowPolicy(true, ProviderInitInfo.DefaultMaxRowNumber);

            foreach (var file in _writtenFiles)
            {
                try
                {
                    if (File.Exists(file)) File.Delete(file);
                }
                catch (IOException)
                {
                }
            }

            _writtenFiles.Clear();
        }

        [TestMethod]
        public void Generate_WhenAllRowsFitOneChunk_ProducesExactlyOneSheetWithoutNumericSuffix()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);
            const string sheetName = "Chunky";
            var records = BuildDataTempRecords(5);

            var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration { LCID = 1048, SheetName = sheetName },
                DataCollection = records
            });

            if (!result.IsSuccess)
                Assert.Fail($"Generating 5 rows into a sheet limit of {NoSplitMaxRows} must succeed. " +
                            $"Reported: {result.GetFirstMessage()}");

            SaveForInspection(result.Response, "SingleChunk_NoSuffix");

            var sheetNames = ReadSheetNames(result.Response, "single chunk export");

            Assert.AreEqual(1, sheetNames.Count);
            Assert.AreEqual(sheetName, sheetNames[0]);

            var rows = ReadSheetRows(result.Response, sheetNames[0], "single chunk export");
            Assert.AreEqual(6, rows.Count);
        }

        [TestMethod]
        public void Generate_WhenRowsExceedSheetLimit_NamesChunksWithOneBasedSuffixes()
        {
            ConfigureRowPolicy(true, 4);
            const string sheetName = "Chunky";
            var records = BuildDataTempRecords(10);

            var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration { LCID = 1048, SheetName = sheetName },
                DataCollection = records
            });

            if (!result.IsSuccess)
                Assert.Fail($"Generating 10 rows into sheets of 4 must succeed. Reported: {result.GetFirstMessage()}");

            SaveForInspection(result.Response, "MultiChunk_OneBasedSuffixes");

            var sheetNames = ReadSheetNames(result.Response, "multi chunk export");

            Assert.AreEqual(3, sheetNames.Count);

            Assert.AreEqual($"{sheetName}_1|{sheetName}_2|{sheetName}_3", string.Join("|", sheetNames));

            foreach (var name in sheetNames)
                Assert.IsFalse(name.EndsWith("_0", StringComparison.Ordinal));

            var expectedDataRows = new[] { 4, 4, 2 };
            for (var i = 0; i < sheetNames.Count; i++)
            {
                var rows = ReadSheetRows(result.Response, sheetNames[i], "multi chunk export");
                Assert.AreEqual(expectedDataRows[i] + 1, rows.Count);
            }
        }

        [TestMethod]
        public void Generate_WhenRowCountEqualsTheSheetLimitExactly_ProducesOneUnsuffixedSheet()
        {
            const int maxRowsPerSheet = 4;
            ConfigureRowPolicy(true, maxRowsPerSheet);
            const string sheetName = "Chunky";
            var records = BuildDataTempRecords(maxRowsPerSheet);

            var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration { LCID = 1048, SheetName = sheetName },
                DataCollection = records
            });

            if (!result.IsSuccess)
                Assert.Fail($"Generating {maxRowsPerSheet} rows into a sheet limit of {maxRowsPerSheet} must " +
                            $"succeed. Reported: {result.GetFirstMessage()}");

            SaveForInspection(result.Response, "ExactSingleChunk");

            var sheetNames = ReadSheetNames(result.Response, "exact single chunk export");

            Assert.AreEqual(1, sheetNames.Count);

            Assert.AreEqual(sheetName, string.Join("|", sheetNames));

            var rows = ReadSheetRows(result.Response, sheetNames[0], "exact single chunk export");
            Assert.AreEqual(maxRowsPerSheet + 1, rows.Count);
        }

        [TestMethod]
        public void Generate_WhenRowCountIsExactlyTwiceTheSheetLimit_ProducesTwoFullSheetsAndNoEmptyThird()
        {
            const int maxRowsPerSheet = 4;
            ConfigureRowPolicy(true, maxRowsPerSheet);
            const string sheetName = "Chunky";
            var records = BuildDataTempRecords(maxRowsPerSheet * 2);

            var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration { LCID = 1048, SheetName = sheetName },
                DataCollection = records
            });

            if (!result.IsSuccess)
                Assert.Fail($"Generating {maxRowsPerSheet * 2} rows into sheets of {maxRowsPerSheet} must " +
                            $"succeed. Reported: {result.GetFirstMessage()}");

            SaveForInspection(result.Response, "ExactTwoChunks");

            var sheetNames = ReadSheetNames(result.Response, "exact two chunk export");

            Assert.AreEqual(2, sheetNames.Count);

            Assert.AreEqual($"{sheetName}_1|{sheetName}_2", string.Join("|", sheetNames));

            for (var i = 0; i < sheetNames.Count; i++)
            {
                var rows = ReadSheetRows(result.Response, sheetNames[i], "exact two chunk export");
                Assert.AreEqual(maxRowsPerSheet + 1, rows.Count);
            }
        }

        [TestMethod]
        public void GenerateTemplate_ReturnsANonEmptyWorkbookWithOneSheetAndAHeaderRow()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);
            const int lcid = 1048;

            var result = DocGenerateParserHelper.GenerateTemplate<DataTemp1>(lcid);

            if (!result.IsSuccess)
                Assert.Fail($"GenerateTemplate<DataTemp1>({lcid}) must succeed. Reported: {result.GetFirstMessage()}");
            Assert.IsNotNull(result.Response);
            Assert.IsTrue(result.Response.Length > 0);

            SaveForInspection(result.Response, "Template_HeaderOnly");

            var sheetNames = ReadSheetNames(result.Response, $"GenerateTemplate<DataTemp1>({lcid})");
            Assert.AreEqual(1, sheetNames.Count);

            var rows = ReadSheetRows(result.Response, sheetNames[0], $"GenerateTemplate<DataTemp1>({lcid})");
            Assert.AreEqual(1, rows.Count);

            var expectedHeaders = new[]
            {
                "Identificator", "IdOperatie", "DataOperatie", "Cod",
                "Data Inceput", "Data Sfarsit", "EsteActiv", "Nume"
            };

            AssertSameHeaderSet(expectedHeaders, rows[0]);
        }

        [TestMethod]
        public void Generate_WithDoubleDecimalLongAndShortProperties_KeepsEveryColumnAndEveryValue()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);
            var record = new NumericScaleModel
            {
                Ticks = 9007199254740993L,
                Rank = 32000,
                Ratio = 0.1d,
                Amount = 99.99m,
                Label = "row-1"
            };

            var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration { LCID = 1033, SheetName = "Numbers" },
                DataCollection = new List<NumericScaleModel> { record }
            });

            if (!result.IsSuccess)
                Assert.Fail("Exporting a model with long/short/double/decimal properties must succeed. " +
                            $"Reported: {result.GetFirstMessage()}");

            var sheetNames = ReadSheetNames(result.Response, "numeric model export");
            Assert.AreEqual(1, sheetNames.Count);

            var rows = ReadSheetRows(result.Response, sheetNames[0], "numeric model export");
            Assert.AreEqual(2, rows.Count);

            var expectedHeaders = new[] { "Ticks", "Rank", "Ratio", "Amount", "Label" };
            AssertSameHeaderSet(expectedHeaders, rows[0]);

            var header = rows[0];
            var data = rows[1];
            Assert.AreEqual(header.Count, data.Count);

            Assert.AreEqual(0.1d, ReadDouble(header, data, "Ratio"), 0d);
            Assert.AreEqual(99.99m, ReadDecimal(header, data, "Amount"));
            Assert.AreEqual(9007199254740993L, ReadLong(header, data, "Ticks"));
            Assert.AreEqual(32000L, ReadLong(header, data, "Rank"));
            Assert.AreEqual("row-1", CellFor(header, data, "Label"));
        }

        [TestMethod]
        public void Generate_ToAPathHoldingALargerFile_LeavesNoStaleTailBytes()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);
            var dataSet = BuildOverwriteDataSet();

            var expected = DocGenerateParserHelper.Generate(dataSet);
            if (!expected.IsSuccess)
                Assert.Fail($"The in-memory reference generation must succeed. Reported: {expected.GetFirstMessage()}");
            Assert.IsTrue(expected.Response.Length > 0);

            var filePath = NewFilePath("Overwrite_LargerExistingFile", "xlsx");
            var filler = new byte[256 * 1024];
            for (var i = 0; i < filler.Length; i++) filler[i] = 0xAB;
            File.WriteAllBytes(filePath, filler);

            var result = DocGenerateParserHelper.Generate(BuildOverwriteDataSet(), filePath);

            if (!result.IsSuccess)
                Assert.Fail($"Writing over an existing file must succeed. Reported: {result.GetFirstMessage()}");

            var onDisk = new FileInfo(filePath).Length;
            Assert.IsTrue(onDisk < filler.Length);
            var tail = File.ReadAllBytes(filePath);
            Assert.AreNotEqual((byte)0xAB, tail[tail.Length - 1]);

            var written = File.ReadAllBytes(filePath);
            var sheetNames = ReadSheetNames(written, "workbook written over a larger existing file");
            Assert.AreEqual(1, sheetNames.Count);
            Assert.AreEqual("OverwriteMe", sheetNames[0]);

            var rows = ReadSheetRows(written, sheetNames[0], "workbook written over a larger existing file");
            Assert.AreEqual(4, rows.Count);
        }

        [TestMethod]
        public void GenerateCsv_WhenAHeaderContainsAComma_QuotesItAndKeepsTheFieldCountStable()
        {
            const string headerWithComma = "Last, First";
            const string valueWithComma = "Doe, Jane";

            var embeddedModel = new List<PropModel>
            {
                new PropModel { CommonName = "Id", DataType = "int", IsNullable = false },
                new PropModel { CommonName = "Label", DataType = "string", IsNullable = false }
            };
            var translated = new List<PropTranslateModel>
            {
                new PropTranslateModel { CommonName = "Id", TranslateName = "Id", Order = 0 },
                new PropTranslateModel { CommonName = "Label", TranslateName = headerWithComma, Order = 1 }
            };
            var data = new List<List<PropNameValue>>
            {
                new List<PropNameValue>
                {
                    new PropNameValue { Name = "Id", Value = 1 },
                    new PropNameValue { Name = headerWithComma, Value = valueWithComma }
                }
            };

            var result = DocGenerateParserHelper.GenerateCsv(embeddedModel, translated,
                (IEnumerable<IEnumerable<PropNameValue>>)data);

            if (!result.IsSuccess)
                Assert.Fail("Generating a CSV with a comma in a column name must succeed. " +
                            $"Reported: {result.GetFirstMessage()}");
            Assert.IsTrue(result.Response.Length > 0);

            SaveForInspection(result.Response, "Csv_QuotedHeader", "csv");

            var text = Encoding.GetEncoding("iso-8859-1").GetString(result.Response);
            var headerLine = text.Split('\n')[0].TrimEnd('\r');

            Assert.IsTrue(headerLine.Contains("\"" + headerWithComma + "\"", StringComparison.Ordinal));

            var records = ParseCsv(text);
            Assert.IsTrue(records.Count >= 2);

            Assert.AreEqual(2, records[0].Count);
            Assert.AreEqual(headerWithComma, records[0][1]);

            Assert.AreEqual(records[0].Count, records[1].Count);
            Assert.AreEqual(valueWithComma, records[1][1]);
        }

        [TestMethod]
        public void Generate_WithParenthesesInTheSheetName_KeepsThemIntact()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);
            const string sheetName = "Sales (2024)";
            var records = BuildDataTempRecords(3);

            var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration { LCID = 1048, SheetName = sheetName },
                DataCollection = records
            });

            if (!result.IsSuccess)
                Assert.Fail($"Exporting to a sheet named '{sheetName}' must succeed: parentheses are legal in " +
                            $"an Excel sheet name. Reported: {result.GetFirstMessage()}");

            var sheetNames = ReadSheetNames(result.Response, "parenthesised sheet name export");

            Assert.AreEqual(1, sheetNames.Count);

            Assert.AreEqual(sheetName, sheetNames[0]);

            var rows = ReadSheetRows(result.Response, sheetNames[0], "parenthesised sheet name export");
            Assert.AreEqual(4, rows.Count);
        }

        [TestMethod]
        public void Generate_WithAnEmptySheetName_FallsBackToSheet1InsteadOfFailing()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);
            var records = BuildDataTempRecords(3);

            var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration { LCID = 1048, SheetName = string.Empty },
                DataCollection = records
            });

            if (!result.IsSuccess)
                Assert.Fail("An empty sheet name must fall back to a default name, not fail the export and " +
                            $"discard the data. Reported: {result.GetFirstMessage()}");

            var sheetNames = ReadSheetNames(result.Response, "empty sheet name export");

            Assert.AreEqual(1, sheetNames.Count);

            Assert.AreEqual("Sheet1", sheetNames[0]);

            var rows = ReadSheetRows(result.Response, sheetNames[0], "empty sheet name export");
            Assert.AreEqual(4, rows.Count);
        }

        [TestMethod]
        public void Generate_WithADecimalValue_UsesANumberFormatThatDisplaysTheFractionalPart()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);
            const string sheetName = "Money";
            var record = new NumericScaleModel
            {
                Ticks = 1L,
                Rank = 1,
                Ratio = 0.1d,
                Amount = 99.99m,
                Label = "row-1"
            };

            var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration { LCID = 1033, SheetName = sheetName },
                DataCollection = new List<NumericScaleModel> { record }
            });

            if (!result.IsSuccess)
                Assert.Fail($"Exporting a decimal value must succeed. Reported: {result.GetFirstMessage()}");

            var rows = ReadSheetRows(result.Response, sheetName, "decimal number format export");
            Assert.AreEqual(2, rows.Count);
            Assert.AreEqual(99.99m, ReadDecimal(rows[0], rows[1], "Amount"));

            var numberFormatId = ReadDataCellNumberFormatId(result.Response, sheetName, "Amount",
                "decimal number format export");

            Assert.AreNotEqual(1u, numberFormatId);

            var decimalCapable = new HashSet<uint> { 0, 2, 4, 10, 11, 39, 48 };
            Assert.IsTrue(decimalCapable.Contains(numberFormatId) || numberFormatId >= 164);
        }

        [TestMethod]
        [Ignore("Known defect: generation is not thread safe (shared static style state). Enable with the concurrency fix.")]
        public void Generate_FromMultipleThreads_ProducesAValidWorkbookEveryTime()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);
            var failures = new List<string>();
            var gate = new object();

            Parallel.For(0, 16, i =>
            {
                var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
                {
                    Configuration = new ExcelWriteConfiguration { LCID = 1048, SheetName = $"Sheet{i}" },
                    DataCollection = BuildDataTempRecords(5)
                });

                if (result.IsSuccess && result.Response != null && result.Response.Length > 0) return;

                lock (gate) failures.Add($"iteration {i}: {result.GetFirstMessage()}");
            });

            Assert.AreEqual(0, failures.Count);
        }

        private static void ConfigureRowPolicy(bool applyPolicy, int maxRowsPerSheet)
        {
            ProviderInitInfo.SetMaxRowNumberPolicyRule(applyPolicy);
            ProviderInitInfo.SetSheetMaxNumberOfRows(maxRowsPerSheet);
        }

        private static List<DataTemp> BuildDataTempRecords(int count)
        {
            var records = new List<DataTemp>(count);
            var start = new DateTime(2024, 1, 1, 8, 0, 0, DateTimeKind.Unspecified);

            for (var i = 0; i < count; i++)
            {
                records.Add(new DataTemp
                {
                    Id = i + 1,
                    Name = $"Row {i + 1}",
                    IsActive = i % 2 == 0,
                    StartDate = start.AddDays(i),
                    EndDate = start.AddDays(i + 1),
                    TempId = i
                });
            }

            return records;
        }

        private static System.Data.DataSet BuildOverwriteDataSet()
        {
            var table = new System.Data.DataTable("OverwriteMe");
            table.Columns.Add(new System.Data.DataColumn("Id", typeof(int)));
            table.Columns.Add(new System.Data.DataColumn("Label", typeof(string)));

            for (var i = 1; i <= 3; i++) table.Rows.Add(i, $"Label {i}");

            var dataSet = new System.Data.DataSet("OverwriteSet");
            dataSet.Tables.Add(table);

            return dataSet;
        }

        private string NewFilePath(string prefix, string extension)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"{prefix}_{Guid.NewGuid():N}.{extension}");
            _writtenFiles.Add(path);

            return path;
        }

        private string SaveForInspection(byte[] content, string name, string extension = "xlsx")
        {
            var folder = Path.Combine(Directory.GetCurrentDirectory(), GeneratedOutputFolder);
            Directory.CreateDirectory(folder);

            var path = Path.Combine(folder, $"{name}.{extension}");
            File.WriteAllBytes(path, content);

            TestContext?.WriteLine($"Generated document saved to: {path}");
            TestContext?.AddResultFile(path);

            return path;
        }

        public class NumericScaleModel
        {
            public long Ticks { get; set; }

            public short Rank { get; set; }

            public double Ratio { get; set; }

            public decimal Amount { get; set; }

            public string Label { get; set; }
        }

        private static List<string> ReadSheetNames(byte[] bytes, string context)
        {
            Assert.IsNotNull(bytes);
            Assert.IsTrue(bytes.Length > 0);

            try
            {
                using var ms = new MemoryStream(bytes);
                using var document = SpreadsheetDocument.Open(ms, false);

                return document.WorkbookPart.Workbook
                    .Descendants<Sheet>()
                    .Select(x => x.Name?.Value)
                    .ToList();
            }
            catch (AssertFailedException)
            {
                throw;
            }
            catch (Exception e)
            {
                Assert.Fail($"{context}: the {bytes.Length} byte(s) produced could not be opened as an xlsx " +
                            $"workbook ({e.GetType().Name}: {e.Message}).");

                return null;
            }
        }

        private static List<List<string>> ReadSheetRows(byte[] bytes, string sheetName, string context)
        {
            Assert.IsNotNull(bytes);
            Assert.IsTrue(bytes.Length > 0);

            try
            {
                using var ms = new MemoryStream(bytes);
                using var document = SpreadsheetDocument.Open(ms, false);

                var workbookPart = document.WorkbookPart;
                var sheet = workbookPart.Workbook.Descendants<Sheet>()
                    .FirstOrDefault(x => x.Name?.Value == sheetName);

                Assert.IsNotNull(sheet);

                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);

                return worksheetPart.Worksheet
                    .Descendants<SheetData>()
                    .First()
                    .Elements<Row>()
                    .Select(row => row.Elements<Cell>()
                        .Select(cell => cell.CellValue == null ? string.Empty : cell.CellValue.InnerText)
                        .ToList())
                    .ToList();
            }
            catch (AssertFailedException)
            {
                throw;
            }
            catch (Exception e)
            {
                Assert.Fail($"{context}: sheet '{sheetName}' could not be read back " +
                            $"({e.GetType().Name}: {e.Message}).");

                return null;
            }
        }

        private static uint ReadDataCellNumberFormatId(byte[] bytes, string sheetName, string columnName, string context)
        {
            Assert.IsNotNull(bytes);
            Assert.IsTrue(bytes.Length > 0);

            using var ms = new MemoryStream(bytes);
            using var document = SpreadsheetDocument.Open(ms, false);

            var workbookPart = document.WorkbookPart;
            var sheet = workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(x => x.Name?.Value == sheetName);
            Assert.IsNotNull(sheet);

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
            var sheetRows = worksheetPart.Worksheet
                .Descendants<SheetData>()
                .First()
                .Elements<Row>()
                .ToList();

            Assert.IsTrue(sheetRows.Count >= 2);

            var headerCells = sheetRows[0].Elements<Cell>().ToList();
            var columnIndex = headerCells.FindIndex(x =>
                (x.CellValue == null ? string.Empty : x.CellValue.InnerText) == columnName);
            Assert.IsTrue(columnIndex >= 0);

            var dataCells = sheetRows[1].Elements<Cell>().ToList();
            Assert.IsTrue(columnIndex < dataCells.Count);

            var cell = dataCells[columnIndex];
            Assert.IsNotNull(cell.StyleIndex);

            var stylesPart = workbookPart.WorkbookStylesPart;
            Assert.IsNotNull(stylesPart);
            Assert.IsNotNull(stylesPart.Stylesheet?.CellFormats);

            var styleIndex = (int)cell.StyleIndex.Value;
            var cellFormats = stylesPart.Stylesheet.CellFormats.Elements<CellFormat>().ToList();
            Assert.IsTrue(styleIndex < cellFormats.Count);

            var format = cellFormats[styleIndex];

            return format.NumberFormatId?.Value ?? 0u;
        }

        private static void AssertSameHeaderSet(IEnumerable<string> expected, List<string> actualHeaderRow)
        {
            var expectedSorted = expected.ToList();
            expectedSorted.Sort(StringComparer.Ordinal);

            var actualSorted = actualHeaderRow.ToList();
            actualSorted.Sort(StringComparer.Ordinal);

            Assert.AreEqual(expectedSorted.Count, actualSorted.Count);

            Assert.AreEqual(string.Join("|", expectedSorted), string.Join("|", actualSorted));
        }

        private static string CellFor(List<string> header, List<string> data, string columnName)
        {
            var index = header.IndexOf(columnName);
            Assert.IsTrue(index >= 0);
            Assert.IsTrue(index < data.Count);

            return data[index];
        }

        private static double ReadDouble(List<string> header, List<string> data, string columnName)
        {
            var raw = CellFor(header, data, columnName);
            Assert.IsTrue(double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value));

            return value;
        }

        private static decimal ReadDecimal(List<string> header, List<string> data, string columnName)
        {
            var raw = CellFor(header, data, columnName);
            Assert.IsTrue(decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value));

            return value;
        }

        private static long ReadLong(List<string> header, List<string> data, string columnName)
        {
            var raw = CellFor(header, data, columnName);
            Assert.IsTrue(decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value));
            Assert.AreEqual(decimal.Truncate(value), value);

            return (long)value;
        }

        private static List<List<string>> ParseCsv(string text)
        {
            var records = new List<List<string>>();
            var record = new List<string>();
            var field = new StringBuilder();
            var inQuotes = false;
            var index = 0;

            while (index < text.Length)
            {
                var current = text[index];

                if (inQuotes)
                {
                    if (current == '"')
                    {
                        if (index + 1 < text.Length && text[index + 1] == '"')
                        {
                            field.Append('"');
                            index += 2;

                            continue;
                        }

                        inQuotes = false;
                        index++;

                        continue;
                    }

                    field.Append(current);
                    index++;

                    continue;
                }

                if (current == '"')
                {
                    inQuotes = true;
                    index++;

                    continue;
                }

                if (current == ',')
                {
                    record.Add(field.ToString());
                    field.Clear();
                    index++;

                    continue;
                }

                if (current == '\r' || current == '\n')
                {
                    if (current == '\r' && index + 1 < text.Length && text[index + 1] == '\n') index++;

                    record.Add(field.ToString());
                    field.Clear();
                    records.Add(record);
                    record = new List<string>();
                    index++;

                    continue;
                }

                field.Append(current);
                index++;
            }

            if (field.Length > 0 || record.Count > 0)
            {
                record.Add(field.ToString());
                records.Add(record);
            }

            return records;
        }

        public class ColumnWidthModel
        {
            [ExcelPropName("Ident", 1033, true, 0, width: 12)]
            public int Id { get; set; }

            [ExcelPropName("Full name", 1033, true, 1, width: 40.5)]
            public string Name { get; set; }

            [ExcelPropName("Note", 1033, true, 2)]
            public string Note { get; set; }
        }

        [TestMethod]
        public void Generate_WithDeclaredColumnWidths_EmitsColsInTheWorksheetBeforeSheetData()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);

            var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration { LCID = 1033, SheetName = "Widths" },
                DataCollection = new List<ColumnWidthModel>
                {
                    new ColumnWidthModel { Id = 1, Name = "first", Note = "n" }
                }
            });

            if (!result.IsSuccess)
                Assert.Fail($"Exporting a model that declares column widths must succeed. " +
                            $"Reported: {result.GetFirstMessage()}");

            SaveForInspection(result.Response, "ColumnWidths_Declared");

            using var ms = new MemoryStream(result.Response);
            using var doc = SpreadsheetDocument.Open(ms, false);

            var worksheetPart = doc.WorkbookPart.WorksheetParts.Single();
            var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();

            if (columns == null)
                Assert.Fail("The worksheet must carry a <cols> element when a width is declared; " +
                            "a spreadsheet application ignores column widths written anywhere else.");

            var declared = columns.Elements<Column>().ToList();
            Assert.AreEqual(2, declared.Count);

            Assert.AreEqual(1u, declared[0].Min.Value);
            Assert.AreEqual(1u, declared[0].Max.Value);
            Assert.AreEqual(12d, declared[0].Width.Value, 0d);

            Assert.AreEqual(2u, declared[1].Min.Value);
            Assert.AreEqual(2u, declared[1].Max.Value);
            Assert.AreEqual(40.5d, declared[1].Width.Value, 0d);

            var children = worksheetPart.Worksheet.ChildElements.ToList();
            var colsIndex = children.FindIndex(x => x is Columns);
            var dataIndex = children.FindIndex(x => x is SheetData);
            Assert.IsTrue(colsIndex >= 0 && colsIndex < dataIndex);
        }

        [TestMethod]
        public void Generate_WithNoDeclaredColumnWidths_OmitsColsEntirely()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);

            var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration { LCID = 1048, SheetName = "NoWidths" },
                DataCollection = BuildDataTempRecords(3)
            });

            if (!result.IsSuccess)
                Assert.Fail($"Exporting a model without widths must succeed. Reported: {result.GetFirstMessage()}");

            SaveForInspection(result.Response, "ColumnWidths_None");

            using var ms = new MemoryStream(result.Response);
            using var doc = SpreadsheetDocument.Open(ms, false);

            var worksheetPart = doc.WorkbookPart.WorksheetParts.Single();
            Assert.IsNull(worksheetPart.Worksheet.GetFirstChild<Columns>());
        }
        public class DegenerateWidthModel
        {
            [ExcelPropName("Zero", 1033, inResult: true, order: 0, width: 0)]
            public int Zero { get; set; }

            [ExcelPropName("Negative", 1033, inResult: true, order: 1, width: -5)]
            public int Negative { get; set; }

            [ExcelPropName("Oversized", 1033, inResult: true, order: 2, width: 300)]
            public int Oversized { get; set; }
        }

        [TestMethod]
        public void Generate_WithZeroOrNegativeWidths_OmitsThoseColumnsInsteadOfHidingThem()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);

            var result = DocGenerateParserHelper.Generate(new ExcelCollectionExportConfiguration
            {
                Configuration = new ExcelWriteConfiguration { LCID = 1033, SheetName = "Degenerate" },
                DataCollection = new List<DegenerateWidthModel>
                {
                    new DegenerateWidthModel { Zero = 1, Negative = 2, Oversized = 3 }
                }
            });

            if (!result.IsSuccess)
                Assert.Fail($"A degenerate width must not fail the export. Reported: {result.GetFirstMessage()}");

            SaveForInspection(result.Response, "ColumnWidths_Degenerate");

            using var ms = new MemoryStream(result.Response);
            using var doc = SpreadsheetDocument.Open(ms, false);

            var columns = doc.WorkbookPart.WorksheetParts.Single().Worksheet.GetFirstChild<Columns>();
            if (columns == null)
                Assert.Fail("The oversized column still declares a width, so <cols> must be present.");

            var declared = columns.Elements<Column>().ToList();

            Assert.AreEqual(1, declared.Count);
            Assert.AreEqual(3u, declared[0].Min.Value);
            Assert.AreEqual(255d, declared[0].Width.Value, 0d);
        }

        [TestMethod]
        public void GenerateTemplate_WithNaNWidth_OmitsTheColumnRatherThanWritingNaN()
        {
            ConfigureRowPolicy(true, NoSplitMaxRows);

            using var stream = new MemoryStream();

            var result = DocGenerateParserHelper.GenerateTemplate(stream, new ExcelTemplateWriteConfiguration
            {
                SheetName = "NaNWidth",
                ColumnHeadings = new List<CellHeaderDefinition>
                {
                    new CellHeaderDefinition("Broken", true, false,
                        new CellDataDefinition(CellDataType.String, SourceCellDataType.String)) { Width = double.NaN },
                    new CellHeaderDefinition("Fine", true, false,
                        new CellDataDefinition(CellDataType.String, SourceCellDataType.String)) { Width = 20 }
                }
            });

            if (!result.IsSuccess)
                Assert.Fail($"A NaN width must not fail the export. Reported: {result.GetFirstMessage()}");

            var bytes = stream.ToArray();
            SaveForInspection(bytes, "ColumnWidths_NaN");

            using var read = new MemoryStream(bytes);
            using var doc = SpreadsheetDocument.Open(read, false);

            var columns = doc.WorkbookPart.WorksheetParts.Single().Worksheet.GetFirstChild<Columns>();
            if (columns == null)
                Assert.Fail("The second column declares a valid width, so <cols> must be present.");

            var declared = columns.Elements<Column>().ToList();

            Assert.AreEqual(1, declared.Count);
            Assert.AreEqual(2u, declared[0].Min.Value);
            Assert.AreEqual(20d, declared[0].Width.Value, 0d);
        }

    }
}
