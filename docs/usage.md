From the beginning to use available functionalities, u must install the package

> `Install-Package DynamicExcelProvider -Version x.x.x.x`

The [README](../README.md) has the short version, install and one working export. This page is the long one: the full method surface, the attribute parameters, the behaviour you cannot guess from a signature.

In the application `Startup.cs` register service:
```csharp
public void ConfigureServices(IServiceCollection services)
        {
            ...
            
            services.RegisterExcelDataSourceProvider();
            // or use global configuration
            services.RegisterExcelDataSourceProvider(option =>
            {
                option.ApplyMaxRowNumberPolicy = true;
                option.SheetMaxNumberOfRows = X;
            });
            
            ...
        }
```

In your class inject service generator:
```csharp
public class GenerateFiles
    {
        private readonly IExcelWriteFactoryProvider _excelWriteFactoryProvider;

        public GenerateFiles(IExcelWriteFactoryProvider excelWriteFactoryProvider)
        {
            _excelWriteFactoryProvider = excelWriteFactoryProvider;
        }
    }
```

## Available methods to generate file

There are many overloads, but only a few shapes. Every entry point either hands you the bytes back, writes into a `Stream` you own, or writes to a file path, and almost every one of them has a `GenerateAsync` twin with a `CancellationToken` at the end. Grouped by what you feed it:

### CSV

```csharp
Task<IResult<byte[]>> GenerateCsvFromKnownAsync(
            IReadOnlyCollection<PropModel> embeddedModelCollection,
            IReadOnlyCollection<PropTranslateModel> availablePropInOutput, 
            IEnumerable<IReadOnlyList<PropNameValue>> data, 
            CancellationToken cancellationToken = default,
            Encoding encoding = null);
            
Task<IResult<byte[]>> GenerateCsvAsync<TDataModel>(
            IReadOnlyCollection<PropModel> embeddedModelCollection,
            IReadOnlyCollection<PropTranslateModel> availablePropInOutput,
            IReadOnlyCollection<TDataModel> data, 
            CancellationToken cancellationToken = default,
            Encoding encoding = null);
            
```

### From a typed collection, column names taken from the attributes

```csharp
Task<IResult<byte[]>> GenerateAsync<TDataModel>(
            IReadOnlyCollection<TDataModel> data,
            int cultureId, 
            CancellationToken cancellationToken = default);
            
Task<IResult> GenerateAsync<TDataModel>(
            Stream stream, 
            IReadOnlyCollection<TDataModel> data,
            int cultureId, 
            CancellationToken cancellationToken = default);
            
```

### From a configuration described at runtime

Use this one when the columns are not known at compile time, when they come from user settings or from a database.

```csharp
Task<IResult<byte[]>> GenerateAsync(
            ExcelCollectionExportConfiguration request, 
            CancellationToken cancellationToken = default);
            
Task<IResult> GenerateAsync(
            Stream stream, 
            ExcelCollectionExportConfiguration request, 
            CancellationToken cancellationToken = default);
            
```

### From a workbook you build yourself

`WorkbookDefinition` is the low level path: you declare the sheets, the header cells and the rows, and nothing is inferred from a model.

```csharp
IResult Generate(
            string filePath, 
            WorkbookDefinition workBook);
            
Task<IResult> GenerateAsync(
            string filePath, 
            WorkbookDefinition workBook, 
            CancellationToken cancellationToken = default);
            
IResult Generate(
            Stream stream, 
            WorkbookDefinition workBook);
            
Task<IResult> GenerateAsync(
            Stream stream, 
            WorkbookDefinition workBook, 
            CancellationToken cancellationToken = default);
```

### From a `DataTable` or a `DataSet`

A `DataSet` produces one worksheet per table.

```csharp
IResult<byte[]> Generate(DataTable dataTable);
IResult Generate(Stream stream, DataTable dataTable);
IResult Generate(string filePath, DataTable dataTable);

IResult<byte[]> Generate(DataSet dataSet);
IResult Generate(Stream stream, DataSet dataSet);
IResult Generate(string filePath, DataSet dataSet);
```

All six have a `GenerateAsync` twin with the same parameters plus `CancellationToken cancellationToken = default` on the end.

### Templates

A template is the header row and nothing under it, for a file the user fills in and sends back to you.

```csharp
IResult<byte[]> GenerateTemplate<T>(
            int lcid,
            IReadOnlyCollection<string> customOutFields = null) where T : class;

IResult GenerateTemplate<T>(
            MemoryStream stream,
            int lcid,
            IReadOnlyCollection<string> customOutFields = null) where T : class;

Task<IResult<byte[]>> GenerateTemplateAsync<T>(
            int lcid,
            IReadOnlyCollection<string> customOutFields = null,
            CancellationToken cancellationToken = default) where T : class;

Task<IResult> GenerateTemplateAsync<T>(
            MemoryStream stream,
            int lcid,
            IReadOnlyCollection<string> customOutFields = null,
            CancellationToken cancellationToken = default) where T : class;

IResult GenerateTemplate(Stream stream, ExcelTemplateWriteConfiguration configuration);

Task<IResult> GenerateTemplateAsync(Stream stream, ExcelTemplateWriteConfiguration configuration);
```

`customOutFields` limits the template to the named columns. Left `null`, the header row contains everything the model exposes for that LCID. The generic overloads that write into a stream ask for a `MemoryStream`, the configuration ones take any `Stream`.

## What comes back

Nothing is thrown at you on the normal path. Every method returns `IResult` or `IResult<byte[]>`, so the calling convention is the same everywhere: look at `IsSuccess`, take `Response` when it is true, read `Messages` when it is not.

```csharp
var result = await _excelWriteFactoryProvider.GenerateAsync(invoiceLines, 1033, cancellationToken);

if (!result.IsSuccess)
{
    _logger.LogWarning("Excel export failed: {Reason}", result.Messages.FirstOrDefault()?.Message);

    return null;
}

return result.Response;
```

The non generic `IResult` works the same way, it just has no `Response` to give you, the file is already in the stream or on disk.

## Defining the columns with `ExcelPropName`

This is how most people use the library. The attribute is repeatable (`AllowMultiple = true`), so you put one on the property for every culture you want to support:

```csharp
[ExcelPropName(propertyName, cultureInfoId, inResult, order: 0, formatCode: null,
               wrapText: false, isBold: false, isItalic: false, width: 0)]
```

- `propertyName` - the header text written in the file for this culture.
- `cultureInfoId` - the LCID, as an integer. 1033 is en-US, 1048 is ro-RO.
- `inResult` - `false` keeps the property out of the file.
- `order` - the position of the column.
- `formatCode` - the Excel number or date format, for example `"dd/MM/yyyy"` or `"#,##0.00"`.
- `wrapText`, `isBold`, `isItalic` - the style flags.
- `width` - the column width, in characters of the default font. That is Excel's own unit, not pixels. `0` or left out and the spreadsheet application decides.

```csharp
public class InvoiceLine
{
    [ExcelPropName("Issued at", 1033, inResult: true, order: 0, formatCode: "dd/MM/yyyy", width: 14)]
    [ExcelPropName("Emis la", 1048, inResult: true, order: 0, formatCode: "dd/MM/yyyy", width: 14)]
    public DateTime IssuedAt { get; set; }

    [ExcelPropName("Total", 1033, inResult: true, order: 1, formatCode: "#,##0.00")]
    [ExcelPropName("Total", 1048, inResult: true, order: 1, formatCode: "#,##0.00")]
    public decimal Total { get; set; }

    // never exported, whatever the culture
    [ExcelPropName("Internal id", 1033, inResult: false)]
    public Guid InternalId { get; set; }
}
```

### The three argument trap

There is a second constructor on the same attribute, `(string propertyName, int cultureInfoId, bool wrapText = false, bool isBold = false, bool isItalic = false)`. A positional call with three arguments binds to that one, not to the one you had in mind:

```csharp
[ExcelPropName("Name", 1033, true)]   // this sets wrapText, and InResult stays false
```

The compiler is happy, the export runs, and the column is simply not in the file. Nothing tells you why. Use named arguments and the ambiguity disappears:

```csharp
[ExcelPropName("Name", 1033, inResult: true, order: 1, width: 40)]
```

Same advice for `order:` and `width:` in general. It costs nothing and it survives the next constructor overload.

## Culture matching

Matching is done on the LCID integer, not on a `CultureInfo` instance, and it is exact. A neutral LCID does not match a specific one. Annotate with 9 (`en`), ask for 1033 (`en-US`) at export time, and the columns are still there but they come out under the C# property names instead of your translations. No error, nothing in the messages, just the wrong headers. Annotate and request the same value.

Because no `CultureInfo` is built while the attribute is constructed, reflecting over an annotated model stays safe under globalization-invariant mode, where the culture lookup itself would throw.

## Column width

On the attribute it is the `width` parameter described above. On the low level path it is a property of the header definition:

```csharp
new CellHeaderDefinition { Name = "Full name", Width = 40 }
```

Only the headers that declare a `Width` produce a column definition in the file. The rest keep the default width of the application, exactly as before this option existed.

## Row limit and sheet names

```csharp
services.RegisterExcelDataSourceProvider(option =>
{
    option.ApplyMaxRowNumberPolicy = true;      // default: true
    option.SheetMaxNumberOfRows = 1_000_000;    // default: 1_000_000
});
```

- With the policy on, data larger than `SheetMaxNumberOfRows` is split over several sheets, suffixed `_1`, `_2` and so on. Counting starts at 1.
- When everything fits in one sheet, the sheet keeps its configured name, with no suffix at all.
- A value of `0` or less falls back to the default, and anything larger is capped at 1,048,575, which is the Excel row limit minus the header row.

## Validations on a template

`ExcelPropValidation` turns a template column into a real Excel data validation, a value list or a range the user cannot type outside of. One per property, unlike `ExcelPropName`:

```csharp
[ExcelPropValidation(validationType, operatorType, allowedValues, minValue, maxValue,
                     errorMessage = null, promptMessage = null, allowEmpty = true)]
```

`ValidationType`: `None`, `Whole`, `Decimal`, `List`, `Date`, `TextLength`, `Custom`.

`ValidationOperatorType`: `None`, `Between`, `NotBetween`, `Equal`, `NotEqual`, `LessThan`, `LessThanOrEqual`, `GreaterThan`, `GreaterThanOrEqual`.

```csharp
[ExcelPropName("Status", 1033, inResult: true, order: 2, width: 18)]
[ExcelPropValidation(ValidationType.List, ValidationOperatorType.None,
    new[] { "New", "Paid", "Cancelled" },
    errorMessage: "Pick one of the listed values.")]
public string Status { get; set; }
```

There are shorter constructors for the cases where you need only the list, or only a min and a max. See the `ExcelPropValidationAttribute` class for the rest.

## CSV

The CSV methods accept an optional `encoding`. When it is left `null` the library keeps its historical default, `ISO-8859-1`, which is what Excel on Windows expects when a `.csv` is opened by double-click. That code page cannot represent characters outside Latin-1 (for example the Romanian `ș` or the `€` sign) and replaces them with `?`, so pass `Encoding.UTF8` when the exported text needs them.

Two more things the CSV writer does on its own:

- Every field is quoted, headers included, and a quote inside a value is doubled.
- A value starting with `=`, `+`, `-`, `@`, a tab or a carriage return gets an apostrophe in front, so a spreadsheet opening the file will not evaluate it as a formula. Values that parse as numbers are left alone, so a negative number stays a number.

## Styles

For more flexibility in the new file generation was added header table style and cell styles like: `Bold`, `Italic`, `WrapText`, `HorizontalAlignment`, `VerticalAlignment`, `CellDataType`, `SourceCellDataType`, `FormatCode`.

## Manual definition or attributes

Method parameters supply can be generated manually or can be defined by property attribute `ExcelPropName`, for more detailed information see the `ExcelPropNameAttribute` class with available descriptions/comments.

In case when parameters and rows are defined manually you have more flexibility and configuration possibilities.

Definition through attribute is easier and with no lost time on custom definition and parsing data. For more detailed configuration, please consult the test project `GeneralDocumentGeneratorTests`.
