// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-02-06 16:49
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-06 20:14
// ***********************************************************************
//  <copyright file="WorkbookParseBuildHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using AggregatedGenericResultMessage;
using AggregatedGenericResultMessage.Abstractions;
using AggregatedGenericResultMessage.Extensions.Result.Messages;
using DomainCommonExtensions.ArraysExtensions;
using DomainCommonExtensions.CommonExtensions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.Extensions;
using DynamicExcelProvider.Helpers.Attributes;
using DynamicExcelProvider.Models.Internal;
using DynamicExcelProvider.Models.Request.Configuration;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using DynamicExcelProvider.Models.Request.Configuration.Template;
using DynamicExcelProvider.WorkXCore.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using TypeExtensions = DynamicExcelProvider.Extensions.TypeExtensions;

// ReSharper disable AccessToModifiedClosure
// ReSharper disable AssignNullToNotNullAttribute
// ReSharper disable PossibleMultipleEnumeration

#endregion

namespace DynamicExcelProvider.Helpers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A workbook parse and build helper.
    /// </summary>
    /// =================================================================================================
    public static class WorkbookParseBuildHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds and parse internal model.
        /// </summary>
        /// <typeparam name="TResult">Type of the result.</typeparam>
        /// <param name="data">The data.</param>
        /// <param name="cultureId">Identifier for the culture.</param>
        /// <returns>
        ///     A list of.
        /// </returns>
        /// =================================================================================================
        internal static IResult<(List<PropTranslateModel>, List<PropModel>)> BuildAndParseInternalModel<TResult>(
            TResult data,
            int cultureId)
        {
            try
            {
                var outputProps = new List<PropTranslateModel>();
                var embeddedModelCollection = new List<PropModel>();
                var props = data.GetProperties();

                foreach (var propertyInfo in props)
                {
                    embeddedModelCollection.Add(new PropModel
                    {
                        CommonName = propertyInfo.Name,
                        DataType = TypeHelper.GetNonNullableType(propertyInfo.PropertyType).ToString(),
                        IsNullable = TypeExtensions.IsNullablePropType(propertyInfo.PropertyType)
                    });
                }

                var propertiesAttributes = PropNameAttributeHelper.GetPropNameAttributeByPassMissing<TResult>(cultureId)
                    .Where(x => x.InResult).OrderBy(x => x.Order)
                    .ToList();
                foreach (var p in propertiesAttributes)
                {
                    outputProps.Add(new PropTranslateModel
                    {
                        CommonName = p.EmbeddedName,
                        Order = p.Order,
                        TranslateName = p.CurrentName,
                        Format = p.FormatCode,
                        IsItalic = p.IsItalic,
                        IsBold = p.IsBold,
                        WrapText = p.WrapText
                    });
                }

                return Result<(List<PropTranslateModel>, List<PropModel>)>
                    .Success((outputProps, embeddedModelCollection));
            }
            catch (Exception e)
            {
                return Result<(List<PropTranslateModel>, List<PropModel>)>
                    .Failure(e.Message)
                    .AddError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds and parse internal model dynamic.
        /// </summary>
        /// <param name="exportConfiguration">The export configuration.</param>
        /// <param name="firstRow">The first row.</param>
        /// <param name="sourceRowType">(Optional) Type of the source row.</param>
        /// <returns>
        ///     A list of.
        /// </returns>
        /// =================================================================================================
        internal static IResult<(List<PropTranslateModel>, List<PropModel>)> BuildAndParseInternalModelDynamic(
            ExcelWriteConfiguration exportConfiguration, dynamic firstRow, Type sourceRowType = null)
        {
            try
            {
                if (firstRow == null && sourceRowType.IsNull())
                    return Result<(List<PropTranslateModel>, List<PropModel>)>
                        .Failure("Can't define row type");

                var outputProps = new List<PropTranslateModel>();
                var embeddedModelCollection = new List<PropModel>();
                var rowType = firstRow?.GetType() == null ? sourceRowType : firstRow.GetType();
                var props = ((IEnumerable<PropertyInfo>)rowType!.GetProperties())
                    .Where(x => TypeExtensions.IsSimpleType(x.PropertyType))
                    .ToList();

                foreach (var propertyInfo in props)
                {
                    var prop = propertyInfo;
                    if (TypeExtensions.IsSimpleType(prop.PropertyType))
                        embeddedModelCollection.Add(new PropModel
                        {
                            CommonName = prop.Name,
                            DataType = TypeHelper.GetNonNullableType(prop.PropertyType).ToString(),
                            IsNullable = TypeExtensions.IsNullablePropType(prop.PropertyType)
                        });
                }

                var propertiesAttributes = ((IEnumerable<ParseModelProperty>)PropNameAttributeHelper
                        .GetPropNameAttributeByPassMissing(exportConfiguration.LCID, rowType))
                    .Where(x => x.InResult).OrderBy(x => x.Order)
                    .ToList();
                foreach (var p in propertiesAttributes)
                {
                    outputProps.Add(new PropTranslateModel
                    {
                        CommonName = p.EmbeddedName,
                        Order = p.Order,
                        TranslateName = p.CurrentName,
                        Format = p.FormatCode,
                        IsItalic = p.IsItalic,
                        IsBold = p.IsBold,
                        WrapText = p.WrapText
                    });
                }

                return Result<(List<PropTranslateModel>, List<PropModel>)>
                    .Success((outputProps, embeddedModelCollection));
            }
            catch (Exception e)
            {
                return Result<(List<PropTranslateModel>, List<PropModel>)>
                    .Failure(e.Message).AddError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds and parse to workbook definition.
        /// </summary>
        /// <typeparam name="TResult">Type of the result.</typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="outputProps">The output properties.</param>
        /// <param name="embeddedModelCollection">Collection of embedded models.</param>
        /// <param name="data">The data.</param>
        /// <param name="isDynamic">True if is dynamic, false if not.</param>
        /// <param name="sourceRowType">(Optional) Type of the source row.</param>
        /// <param name="generateSheetValidations">(Optional) True to generate sheet validations.</param>
        /// <returns>
        ///     An IResult&lt;WorkbookDefinition&gt;
        /// </returns>
        /// =================================================================================================
        internal static IResult<WorkbookDefinition> BuildAndParseToWorkbookDefinition<TResult>(
            string sheetName, IReadOnlyCollection<PropTranslateModel> outputProps,
            IReadOnlyCollection<PropModel> embeddedModelCollection, IReadOnlyCollection<TResult> data,
            bool isDynamic, Type sourceRowType = null, bool generateSheetValidations = false)
        {
            try
            {
                if (data.IsNullOrEmptyEnumerable() && sourceRowType.IsNull())
                    return Result<WorkbookDefinition>
                        .Failure("Can't define row type");

                var rowType = data.IsNullOrEmptyEnumerable() ? sourceRowType : data.FirstOrDefault()?.GetType();

                IEnumerable<PropertyInfo> props = rowType?.GetProperties()
                    .Where(x => TypeExtensions.IsSimpleType(x.PropertyType))
                    .ToList();

                var wbDef = new WorkbookDefinition
                {
                    Worksheets = new List<WorksheetDefinition>(1)
                    {
                        new WorksheetDefinition
                        {
                            Name = sheetName,
                            ColumnHeadings = outputProps
                                .OrderBy(x => x.Order)
                                .Select(x => new CellHeaderDefinition
                                {
                                    Name = x.TranslateName,
                                    IsItalic = x.IsItalic,
                                    IsBold = x.IsBold,
                                    WrapText = x.WrapText,
                                    CellData = new CellDataDefinition
                                    {
                                        CellDataType = DataTypeHelper.GetColumnType(embeddedModelCollection
                                            .FirstOrDefault(a => a.CommonName == x.CommonName)?.DataType),
                                        SourceCellDataType = DataTypeHelper.GetSourceColumnType(embeddedModelCollection
                                            .FirstOrDefault(a => a.CommonName == x.CommonName)?.DataType),
                                        FormatCode = x.Format ?? "0"
                                    }
                                }),
                            Rows = data.Select(row => new RowDefinition
                            {
                                Cells = outputProps.OrderBy(x => x.Order)
                                    .Select(cell => new CellValueDefinition
                                    {
                                        DefaultValue = default,
                                        Value = isDynamic.IsFalse()
                                            ? row.GetPropertiesInfoFromT()
                                                .FirstOrDefault(a => a.Name == cell.CommonName)?
                                                .GetGetMethod(true).Invoke(row, new object[] { })
                                            : props
                                                .FirstOrDefault(a => a.Name == cell.CommonName)?
                                                .GetGetMethod(true).Invoke(row, new object[] { })
                                    })
                            }),
                            SheetValidations =
                                generateSheetValidations.IsTrue()
                                    ? DataValidationsBuildHelper.BuildSheetDataValidations(ref props, ref outputProps)
                                    : null
                        }
                    }
                };

                return Result<WorkbookDefinition>.Success(wbDef);
            }
            catch (Exception e)
            {
                return Result<WorkbookDefinition>
                    .Failure(e.Message)
                    .AddError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds and parse to workbook definition template.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="columnHeadings">The column headings.</param>
        /// <param name="sheetValidations">The sheet validations.</param>
        /// <param name="generateSheetValidations">(Optional) True to generate sheet validations.</param>
        /// <returns>
        ///     An IResult&lt;WorkbookDefinition&gt;
        /// </returns>
        /// =================================================================================================
        internal static IResult<WorkbookDefinition> BuildAndParseToWorkbookDefinitionTemplate(
            string sheetName, IEnumerable<CellHeaderDefinition> columnHeadings,
            IEnumerable<TemplateDataValidation> sheetValidations,
            bool generateSheetValidations = false)
        {
            try
            {
                var wbDef = new WorkbookDefinition
                {
                    Worksheets = new List<WorksheetDefinition>(1)
                    {
                        new WorksheetDefinition
                        {
                            Name = sheetName,
                            ColumnHeadings = columnHeadings,
                            Rows = new List<RowDefinition>(),
                            SheetValidations =
                                generateSheetValidations.IsTrue()
                                    ? DataValidationsBuildHelper.BuildSheetDataValidations(ref sheetValidations)
                                    : null
                        }
                    }
                };

                return Result<WorkbookDefinition>.Success(wbDef);
            }
            catch (Exception e)
            {
                return Result<WorkbookDefinition>
                    .Failure(e.Message)
                    .AddError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Parse and build rows from source.
        /// </summary>
        /// <param name="cells">The cells.</param>
        /// <param name="dataRows">The data rows.</param>
        /// <returns>
        ///     An IResult&lt;IEnumerable&lt;RowDefinition&gt;&gt;
        /// </returns>
        /// =================================================================================================
        public static IResult<IEnumerable<RowDefinition>> ParseAndBuildRowsFromSource(
            IEnumerable<CellHeaderDefinition> cells, 
            IEnumerable<dynamic> dataRows)
        {
            try
            {
                var resultRows = new List<RowDefinition>();
                foreach (var row in dataRows)
                {
                    var rowType = row?.GetType();
                    var props = ((IEnumerable<PropertyInfo>)rowType!.GetProperties())
                        .Where(x => TypeExtensions.IsSimpleType(x.PropertyType)).ToList();
                    var rowCells = new List<CellValueDefinition>();
                    foreach (var cell in cells)
                    {
                        var value = props
                            .FirstOrDefault(a => a.Name == cell.Name)?
                            .GetGetMethod(true).Invoke(row, new object[] { });
                        var rowCell = new CellValueDefinition(value, null);
                        rowCells.Add(rowCell);
                    }

                    resultRows.Add(new RowDefinition(rowCells));
                }

                return Result<IEnumerable<RowDefinition>>.Success(resultRows);
            }
            catch (Exception e)
            {
                return Result<IEnumerable<RowDefinition>>
                    .Failure(e.Message)
                    .AddError(e);
            }
        }
    }
}