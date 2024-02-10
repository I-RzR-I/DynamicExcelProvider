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
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.Extensions;
using DynamicExcelProvider.Models.Internal;
using DynamicExcelProvider.Models.Request.Configuration;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using DynamicExcelProvider.WorkXCore.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
// ReSharper disable PossibleMultipleEnumeration

#endregion

namespace DynamicExcelProvider.Helpers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A workbook parse and build helper.
    /// </summary>
    /// <remarks>
    /// </remarks>
    /// =================================================================================================
    public static class WorkbookParseBuildHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds and parse internal model.
        /// </summary>
        /// <remarks>
        /// </remarks>
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
                    embeddedModelCollection.Add(new PropModel
                    {
                        CommonName = propertyInfo.Name, 
                        DataType = TypeHelper.GetNonNullableType(propertyInfo.PropertyType).ToString(), 
                        IsNullable = propertyInfo.PropertyType.IsNullablePropType()
                    });

                var propertiesAttributes = PropNameAttributeHelper.GetPropNameAttributeByPassMissing<TResult>(cultureId)
                    .Where(x => x.InResult).OrderBy(x => x.Order)
                    .ToList();
                foreach (var p in propertiesAttributes)
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

                return Result<(List<PropTranslateModel>, List<PropModel>)>
                    .Success((outputProps, embeddedModelCollection));
            }
            catch (Exception e)
            {
                return Result<(List<PropTranslateModel>, List<PropModel>)>
                    .Failure().AddError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds and parse internal model dynamic.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="exportConfiguration">The export configuration.</param>
        /// <param name="firstRow">The first row.</param>
        /// <returns>
        ///     A list of.
        /// </returns>
        /// =================================================================================================
        internal static IResult<(List<PropTranslateModel>, List<PropModel>)> BuildAndParseInternalModelDynamic(
            ExcelWriteConfiguration exportConfiguration,
            dynamic firstRow)
        {
            try
            {
                var outputProps = new List<PropTranslateModel>();
                var embeddedModelCollection = new List<PropModel>();
                var rowType = firstRow?.GetType();
                var props = ((IEnumerable<PropertyInfo>)rowType!.GetProperties())
                    .Where(x => x.PropertyType.IsSimpleType())
                    .ToList();

                foreach (var propertyInfo in props)
                {
                    var prop = propertyInfo;
                    if (prop.PropertyType.IsSimpleType())
                        embeddedModelCollection.Add(new PropModel
                        {
                            CommonName = prop.Name, 
                            DataType = TypeHelper.GetNonNullableType(prop.PropertyType).ToString(), 
                            IsNullable = prop.PropertyType.IsNullablePropType()
                        });
                }

                var propertiesAttributes = ((IEnumerable<ParseModelProperty>)PropNameAttributeHelper.GetPropNameAttributeByPassMissing(exportConfiguration.LCID, rowType))
                    .Where(x => x.InResult).OrderBy(x => x.Order)
                    .ToList();
                foreach (var p in propertiesAttributes)
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

                return Result<(List<PropTranslateModel>, List<PropModel>)>
                    .Success((outputProps, embeddedModelCollection));
            }
            catch (Exception e)
            {
                return Result<(List<PropTranslateModel>, List<PropModel>)>
                    .Failure().AddError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds and parse to workbook definition.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <typeparam name="TResult">Type of the result.</typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="outputProps">The output properties.</param>
        /// <param name="embeddedModelCollection">Collection of embedded models.</param>
        /// <param name="data">The data.</param>
        /// <param name="isDynamic">True if is dynamic, false if not.</param>
        /// <returns>
        ///     An IResult&lt;WorkbookDefinition&gt;
        /// </returns>
        /// =================================================================================================
        internal static IResult<WorkbookDefinition> BuildAndParseToWorkbookDefinition<TResult>(
            string sheetName, IReadOnlyCollection<PropTranslateModel> outputProps,
            IReadOnlyCollection<PropModel> embeddedModelCollection, IReadOnlyCollection<TResult> data, bool isDynamic)
        {
            try
            {
                IEnumerable<PropertyInfo> props = data.FirstOrDefault().GetProperties()
                    .Where(x => x.PropertyType.IsSimpleType())
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
                            })
                        }
                    }
                };

                return Result<WorkbookDefinition>.Success(wbDef);
            }
            catch (Exception e)
            {
                return Result<WorkbookDefinition>
                    .Failure().AddError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Parse and build rows from source.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="cells">The cells.</param>
        /// <param name="dataRows">The data rows.</param>
        /// <returns>
        ///     An IResult&lt;IEnumerable&lt;RowDefinition&gt;&gt;
        /// </returns>
        /// =================================================================================================
        public static IResult<IEnumerable<RowDefinition>> ParseAndBuildRowsFromSource(IEnumerable<CellHeaderDefinition> cells, IEnumerable<dynamic> dataRows)
        {
            try
            {
                var resultRows = new List<RowDefinition>();
                foreach (var row in dataRows)
                {
                    var rowType = row?.GetType();
                    var props = ((IEnumerable<PropertyInfo>)rowType!.GetProperties())
                        .Where(x => x.PropertyType.IsSimpleType()).ToList();
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
                    .Failure().AddError(e);
            }
        }
    }
}