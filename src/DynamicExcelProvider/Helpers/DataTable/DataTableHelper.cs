// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-13 12:04
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:12
// ***********************************************************************
//  <copyright file="DataTableHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DomainCommonExtensions.CommonExtensions;
using DynamicExcelProvider.Extensions;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

// ReSharper disable PossibleMultipleEnumeration
// ReSharper disable PossibleInvalidOperationException

#endregion

namespace DynamicExcelProvider.Helpers.DataTable
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A data table helper.
    /// </summary>
    /// =================================================================================================
    internal static class DataTableHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Information describing the general table.
        /// </summary>
        /// =================================================================================================
        private static IReadOnlyCollection<PropModel> _generalTableData;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     The available property in output.
        /// </summary>
        /// =================================================================================================
        private static IEnumerable<PropTranslateModel> _availablePropInOutput;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes the data table.
        /// </summary>
        /// <param name="generalTableData">Information describing the general table.</param>
        /// <param name="availablePropInOutput">The available property in output.</param>
        /// =================================================================================================
        internal static void InitDataTable(
            IReadOnlyCollection<PropModel> generalTableData,
            IEnumerable<PropTranslateModel> availablePropInOutput)
        {
            _generalTableData = generalTableData;
            _availablePropInOutput = availablePropInOutput;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Creates table and columns.
        /// </summary>
        /// <param name="tableName">(Optional) Name of the table.</param>
        /// <param name="cultureId">(Optional) Identifier for the culture.</param>
        /// <returns>
        ///     The new table and columns.
        /// </returns>
        /// =================================================================================================
        internal static System.Data.DataTable CreateTableAndColumns(string tableName = null, int? cultureId = null)
        {
            var table = string.IsNullOrEmpty(tableName)
                ? new System.Data.DataTable()
                : new System.Data.DataTable(tableName);

            table.Locale = new CultureInfo(cultureId.IsNull() ? 1033 : cultureId.Value!);

            foreach (var col in _availablePropInOutput.OrderBy(x => x.Order))
            {
                var columnInfo = _generalTableData.FirstOrDefault(x => x.CommonName == col.CommonName);
                if (columnInfo != null) table.Columns.Add(col.TranslateName, DataTypeHelper.GetColumnType(columnInfo.DataType, columnInfo.IsNullable));
            }

            return table;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A System.Data.DataTable extension method that adds a record from known to 'record'.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="record">The record.</param>
        /// =================================================================================================
        internal static void AddRecordFromKnown(this System.Data.DataTable table, IEnumerable<PropNameValue> record)
        {
            var row = table.NewRow();
            var idx = 0;
            foreach (var item in _availablePropInOutput.OrderBy(x => x.Order))
            {
                var columnInfo = _generalTableData.FirstOrDefault(x => x.CommonName == item.CommonName);
                if (columnInfo == null) continue;
                {
                    var dataType = DataTypeHelper.GetColumnType(columnInfo.DataType, columnInfo.IsNullable);
                    var data = record.FirstOrDefault(x => x.Name == item.TranslateName)?.Value;
                    var convertedData = Convert.ChangeType(data, dataType);
                    row[idx] = convertedData;
                    idx++;
                }
            }

            table.Rows.Add(row);
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A System.Data.DataTable extension method that adds a record to 'record'.
        /// </summary>
        /// <typeparam name="TRow">Type of the row.</typeparam>
        /// <param name="table">The table.</param>
        /// <param name="record">The record.</param>
        /// =================================================================================================
        internal static void AddRecord<TRow>(this System.Data.DataTable table, TRow record) where TRow : class
        {
            var recordProps = record.GetPropertiesInfoFromT();
            var row = table.NewRow();
            var idx = 0;
            foreach (var item in _availablePropInOutput.OrderBy(x => x.Order))
            {
                var columnInfo = _generalTableData.FirstOrDefault(x => x.CommonName == item.CommonName);

                if (columnInfo == null) continue;
                var dataType = DataTypeHelper.GetColumnType(columnInfo.DataType, columnInfo.IsNullable);
                var dataProp = recordProps.FirstOrDefault(x => x.Name == item.TranslateName);

                if (dataProp == null) continue;
                {
                    var propValue = dataProp.GetGetMethod(true).Invoke(record, new object[] { });

                    var convertedData = Convert.ChangeType(propValue, dataType);
                    row[idx] = convertedData;
                    idx++;
                }
            }

            table.Rows.Add(row);
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A System.Data.DataTable extension method that adds a record from t to 'record'.
        /// </summary>
        /// <typeparam name="TRow">Type of the row.</typeparam>
        /// <param name="table">The table.</param>
        /// <param name="record">The record.</param>
        /// =================================================================================================
        internal static void AddRecordFromT<TRow>(this System.Data.DataTable table, TRow record)
        {
            var recordProps = record.GetPropertiesInfoFromT();
            var row = table.NewRow();
            var idx = 0;
            foreach (var item in _availablePropInOutput.OrderBy(x => x.Order))
            {
                var columnInfo = _generalTableData.FirstOrDefault(x => x.CommonName == item.CommonName);

                if (columnInfo == null) continue;
                var dataType = DataTypeHelper.GetColumnType(columnInfo.DataType, columnInfo.IsNullable);
                var dataProp = recordProps.FirstOrDefault(x => x.Name == item.CommonName);

                if (dataProp == null) continue;
                {
                    var propValue = dataProp.GetGetMethod(true).Invoke(record, new object[] { });

                    var convertedData = Convert.ChangeType(propValue, dataType);
                    row[idx] = convertedData;
                    idx++;
                }
            }

            table.Rows.Add(row);
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A System.Data.DataTable extension method that adds a record from dynamic.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="record">The record.</param>
        /// <param name="recordProps">The record properties.</param>
        /// =================================================================================================
        internal static void AddRecordFromDynamic(this System.Data.DataTable table, dynamic record, IEnumerable<PropertyInfo> recordProps)
        {
            var row = table.NewRow();
            var idx = 0;
            foreach (var item in _availablePropInOutput.OrderBy(x => x.Order))
            {
                var columnInfo = _generalTableData.FirstOrDefault(x => x.CommonName == item.CommonName);

                if (columnInfo == null) continue;
                var dataType = DataTypeHelper.GetColumnType(columnInfo.DataType, columnInfo.IsNullable);
                var propertyInfos = recordProps.ToList();
                var dataProp = propertyInfos.FirstOrDefault(x => x.Name == item.CommonName);

                if (dataProp == null) continue;
                {
                    var propValue = dataProp.GetGetMethod(true).Invoke(record, new object[] { });

                    var convertedData = Convert.ChangeType(propValue, dataType);
                    row[idx] = convertedData;
                    idx++;
                }
            }

            table.Rows.Add(row);
        }
    }
}