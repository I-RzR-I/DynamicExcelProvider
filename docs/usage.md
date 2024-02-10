From the beginning to use available functionalities, u must install the package

> `Install-Package DynamicExcelProvider -Version x.x.x.x`

In the application `Startup.cs` register service:
```csharp
public void ConfigureServices(IServiceCollection services)
        {
            ...
            
            services.RegisterExcelDataSourceProvider();
            
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

Available methods to generate file:
```csharp
Task<IResult<byte[]>> GenerateCsvFromKnownAsync(
            IReadOnlyCollection<PropModel> embeddedModelCollection,
            IReadOnlyCollection<PropTranslateModel> availablePropInOutput, 
            IEnumerable<IReadOnlyList<PropNameValue>> data, 
            CancellationToken cancellationToken = default);
            
Task<IResult<byte[]>> GenerateCsvAsync<TDataModel>(
            IReadOnlyCollection<PropModel> embeddedModelCollection,
            IReadOnlyCollection<PropTranslateModel> availablePropInOutput,
            IReadOnlyCollection<TDataModel> data, 
            CancellationToken cancellationToken = default);
            
Task<IResult<byte[]>> GenerateAsync<TDataModel>(
            IReadOnlyCollection<TDataModel> data,
            int cultureId, 
            CancellationToken cancellationToken = default);
            
Task<IResult> GenerateAsync<TDataModel>(
            Stream stream, 
            IReadOnlyCollection<TDataModel> data,
            int cultureId, 
            CancellationToken cancellationToken = default);
            
Task<IResult<byte[]>> GenerateAsync(
            ExcelCollectionExportConfiguration request, 
            CancellationToken cancellationToken = default);
            
Task<IResult> GenerateAsync(
            Stream stream, 
            ExcelCollectionExportConfiguration request, 
            CancellationToken cancellationToken = default);
            
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

For more flexibility in the new file generation was added header table style and cell styles like: `Bold`, `Italic`, `WrapText`, `HorizontalAlignment`, `VerticalAlignment`, `CellDataType`, `SourceCellDataType`, `FormatCode`.

Method parameters supply can be generated manually or can be defined by property attribute `ExcelPropName`, for more detailed information see the `ExcelPropNameAttribute` class with available descriptions/comments.

In case when parameters and rows are defined manually you have more flexibility and configuration possibilities.

Definition through attribute is easier and with no lost time on custom definition and parsing data. For more detailed configuration, please consult the test project `GeneralDocumentGeneratorTests`.
