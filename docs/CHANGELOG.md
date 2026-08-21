### **v3.0.0.7667** [[RzR](mailto:108324929+I-RzR-I@users.noreply.github.com)] 21-08-2026
* [c156822] (RzR) -> Auto commit uncommited files
* [9352925] (RzR) -> Add configurable column width and emit it into the worksheet.
* [12d5ad0] (RzR) -> Fix culture resolution failing under globalization-invariant mode.
* [a2727ff] (RzR) -> Add content level assertions over generated documents.
* [ede3339] (RzR) -> Harden CSV output and expose the output encoding.
* [ee26063] (RzR) -> Fix file write truncation and empty workbook handling.
* [4c5e4f1] (RzR) -> Fix numeric type mapping and DBNull handling.
* [0d80d2c] (RzR) -> Fix sheet naming, row-limit policy default and number format.
* [7b6bc32] (RzR) -> Upgrade reference packages version, migrate namespaces and fix errors.

### **v2.1.0.4942** [[RzR](mailto:108324929+I-RzR-I@users.noreply.github.com)] 28-10-2025
* [d6fb76f] (RzR) -> Auto commit uncommited files
* [74ebbfa] (RzR) -> Fix project name in scripts.
* [92206fc] (RzR) -> Add script generation and adjust docs.
* [859d865] (RzR) -> Add new export methods from `DataTable` and `DataSet`.
* [0e37bee] (RzR) -> Add max row limit per sheet implementation.
* [b6dbce2] (RzR) -> Add max row limit per sheet configure option.
* [e2e3cfb] (RzR) -> Reorganize methods and add new method.
* [d964bc5] (RzR) -> Rename internal extension methods.

### **v2.0.0.0** 
-> Add template generation based on user defined configuration (fields and validations) `GenerateTemplateAsync` and `GenerateTemplate`; <br />
-> Add custom user defined fields on template generation based on class type (GenerateTemplate/Async&lt;T&gt;); <br />
-> Remove cast to int on data validation attribute; <br />
-> Add new test to verify data validation on DATE buid; <br />
-> Adjust cell data validation build flow. Add/adjust DATE format on validation; <br />
-> Adjust store and data format to use custom defined; <br />
-> Add custom numbering formats; <br />
-> Upgrade reference package version: `AggregatedGenericResultMessage`, `DocumentFormat.OpenXml`, `DomainCommonExtensions`.<br />

### **v1.2.0.0** 
-> Upgrade reference package version: `AggregatedGenericResultMessage`, `DocumentFormat.OpenXml`, `DomainCommonExtensions`.

### **v1.1.1.6701** 
-> Update reference package version, fixing CVE (`CVE-2024-43485`);

## **v1.1.0.0**
-> Add template generation; <br />
-> Add attribute with column validation on generate template and complete it; <br />
-> Small code clean and fixes; <br />
-> Upgrade reference libraries.
