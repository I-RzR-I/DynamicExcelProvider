> **Note** This repository is developed using .netstandard2.0.

[![NuGet Version](https://img.shields.io/nuget/v/DynamicExcelProvider.svg?style=flat&logo=nuget)](https://www.nuget.org/packages/DynamicExcelProvider/)
[![Nuget Downloads](https://img.shields.io/nuget/dt/DynamicExcelProvider.svg?style=flat&logo=nuget)](https://www.nuget.org/packages/DynamicExcelProvider)

The current repository is based on `DocumentFormat.OpenXml` and aims to make exporting to an Excel list easier. So in other words, it is a wrapper on a previously specified library. The idea to create this repository has been initiated and its roots have grown more and more in a long period. 

Countless times I faced the problem/need to implement some functionalities to export map a list to Excel (`csv` or `xlsx`) file with different columns names (in some cases related to specific language), dynamic number of columns, or in user-specified order, etc. As a result, the implemented solution has some basic functionalities previously mentioned.

For more information about that, follow the info from using doc.

**In case you wish to use it in your project, u can install the package from <a href="https://www.nuget.org/packages/DynamicExcelProvider" target="_blank">nuget.org</a>** or specify what version you want:

> `Install-Package DynamicExcelProvider -Version x.x.x.x`

## Content
1. [USING](docs/usage.md)
1. [CHANGELOG](docs/CHANGELOG.md)
1. [BRANCH-GUIDE](docs/branch-guide.md)