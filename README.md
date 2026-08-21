> **Note** This repository is developed using .netstandard2.0.

[![NuGet Version](https://img.shields.io/nuget/v/DynamicExcelProvider.svg?style=flat&logo=nuget)](https://www.nuget.org/packages/DynamicExcelProvider/)
[![Nuget Downloads](https://img.shields.io/nuget/dt/DynamicExcelProvider.svg?style=flat&logo=nuget)](https://www.nuget.org/packages/DynamicExcelProvider)

The current repository is based on `DocumentFormat.OpenXml` and aims to make exporting to an Excel list easier. So in other words, it is a wrapper on a previously specified library. The idea to create this repository has been initiated and its roots have grown more and more in a long period. 

Countless times I faced the problem/need to implement some functionalities to export map a list to Excel (`csv` or `xlsx`) file with different columns names (in some cases related to specific language), dynamic number of columns, or in user-specified order, etc. As a result, the implemented solution has some basic functionalities previously mentioned.

For more information about that, follow the info from using doc.

**In case you wish to use it in your project, u can install the package from <a href="https://www.nuget.org/packages/DynamicExcelProvider" target="_blank">nuget.org</a>** or specify what version you want:

> `Install-Package DynamicExcelProvider -Version x.x.x.x`

## What you get

- Column names per language, taken from the model itself. You annotate the property once per culture and ask for a specific LCID at export time.
- Control over which properties end up in the file, and in which order.
- `xlsx` from a typed collection, from a `DataTable`, from a `DataSet` (one worksheet per table), or from a workbook you build by hand.
- `csv` with an encoding you choose, and with the usual formula-injection traps handled.
- Header-only templates, with optional data validations on the columns.
- Column width per column, in Excel's own unit.
- Big lists split over several sheets when they pass the row limit.
- Every call returns a result object instead of throwing, so you check `IsSuccess` and move on.

## Quick start

Register the provider at startup:

```csharp
using Microsoft.Extensions.DependencyInjection;

services.RegisterExcelDataSourceProvider();
```

Annotate the model. One `ExcelPropName` per culture, the LCID is the integer culture id (1033 is en-US, 1048 is ro-RO):

```csharp
using DynamicExcelProvider.Attributes;

public class InvoiceLine
{
    [ExcelPropName("Product", 1033, inResult: true, order: 0, width: 30)]
    [ExcelPropName("Produs", 1048, inResult: true, order: 0, width: 30)]
    public string Product { get; set; }

    [ExcelPropName("Quantity", 1033, inResult: true, order: 1)]
    [ExcelPropName("Cantitate", 1048, inResult: true, order: 1)]
    public int Quantity { get; set; }

    [ExcelPropName("Issued at", 1033, inResult: true, order: 2, formatCode: "dd/MM/yyyy")]
    [ExcelPropName("Emis la", 1048, inResult: true, order: 2, formatCode: "dd/MM/yyyy")]
    public DateTime IssuedAt { get; set; }

    // never exported, whatever the culture
    [ExcelPropName("Internal id", 1033, inResult: false)]
    public Guid InternalId { get; set; }
}
```

Inject `IExcelWriteFactoryProvider` and export:

```csharp
using DynamicExcelProvider.Abstractions;

public class InvoiceExportService
{
    private readonly IExcelWriteFactoryProvider _excelProvider;

    public InvoiceExportService(IExcelWriteFactoryProvider excelProvider)
        => _excelProvider = excelProvider;

    public async Task<byte[]> ExportAsync(IReadOnlyCollection<InvoiceLine> lines, CancellationToken ct)
    {
        var result = await _excelProvider.GenerateAsync(lines, 1033, ct);

        if (!result.IsSuccess)
            return null; // result carries the messages, log them or map them to your own error

        return result.Response;
    }
}
```

Matching is done on the LCID integer, not on a `CultureInfo` instance, so it also works when the app runs in globalization-invariant mode.

### The other entry points

Same provider, different input. These are the shapes, each one has stream / byte array / file path variants:

```csharp
// dynamic column set, described at runtime instead of by attributes
IResult<byte[]> Generate(ExcelCollectionExportConfiguration request);
Task<IResult<byte[]>> GenerateAsync(ExcelCollectionExportConfiguration request, CancellationToken ct = default);

// ADO.NET, a DataSet gives one worksheet per table
Task<IResult<byte[]>> GenerateAsync(DataTable dataTable, CancellationToken ct = default);
Task<IResult<byte[]>> GenerateAsync(DataSet dataSet, CancellationToken ct = default);

// full manual control over sheets, headers and cells
IResult Generate(string filePath, WorkbookDefinition workBook);

// header row only, for a file the user fills in and sends back
IResult<byte[]> GenerateTemplate<T>(int lcid, IReadOnlyCollection<string> customOutFields = null);
```

On templates you can also put `[ExcelPropValidation]` on a property to get a real Excel data validation on that column, a value list or a min/max range. See [usage](docs/usage.md) for the parameters.

## Column width

`width` is the column width in characters of the default font. That is Excel's own unit, not pixels. Leave it out or pass `0` and the spreadsheet application decides on its own, which is the behaviour you had before this option existed.

```csharp
[ExcelPropName("Full name", 1033, true, 1, width: 40)]
public string Name { get; set; }
```

On the low-level path it is a property on the header definition:

```csharp
new CellHeaderDefinition { Name = "Full name", Width = 40 }
```

## Options

```csharp
services.RegisterExcelDataSourceProvider(option =>
{
    option.ApplyMaxRowNumberPolicy = true;      // default: true
    option.SheetMaxNumberOfRows = 1_000_000;    // default: 1_000_000
});
```

With the policy on, a data set larger than `SheetMaxNumberOfRows` is split over several sheets, suffixed `Sheet_1`, `Sheet_2` and so on. Counting starts at 1. If everything fits in one sheet, the sheet keeps its configured name with no suffix at all.

## CSV

The CSV methods take the encoding as the last parameter, after the cancellation token:

```csharp
Task<IResult<byte[]>> GenerateCsvAsync<TDataModel>(
    IReadOnlyCollection<PropModel> embeddedModelCollection,
    IReadOnlyCollection<PropTranslateModel> availablePropInOutput,
    IReadOnlyCollection<TDataModel> data,
    CancellationToken cancellationToken = default,
    Encoding encoding = null) where TDataModel : class;
```

`GenerateCsvFromKnownAsync` has the same tail. Things worth knowing:

- The default encoding is ISO-8859-1, kept that way so existing consumers do not break. It cannot represent anything outside Latin-1, and those characters come out as `?`. If your data is not Latin-1, pass `Encoding.UTF8` explicitly.
- Every field is quoted, headers included, and a quote inside a value is doubled.
- A value starting with `=`, `+`, `-`, `@`, a tab or a carriage return gets an apostrophe in front of it, so the spreadsheet does not treat it as a formula. Values that parse as numbers are left alone, so negative numbers stay numbers.


## Content
1. [USING](docs/usage.md)
1. [CHANGELOG](docs/CHANGELOG.md)
1. [BRANCH-GUIDE](docs/branch-guide.md)
