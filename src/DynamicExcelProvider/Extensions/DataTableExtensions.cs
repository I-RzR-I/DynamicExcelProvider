// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-13 12:05
// 
//  Last Modified By : RzR
//  Last Modified On : 2023-03-13 14:23
// ***********************************************************************
//  <copyright file="DataTableExtensions.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.Helpers;
using DynamicExcelProvider.Models.Request.Configuration.Property;
using RzR.Extensions.Domain.Primitives;
using RzR.ResultMessage;
using RzR.ResultMessage.Abstractions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Text;

// ReSharper disable InconsistentNaming

#endregion

namespace DynamicExcelProvider.Extensions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Data table extensions.
    /// </summary>
    /// <remarks>
    /// </remarks>
    /// =================================================================================================
    internal static class DataTableExtensions
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Converts the passed in data table to a CSV-style string.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="table">Table to convert.</param>
        /// <returns>
        ///     Resulting CSV-style string.
        /// </returns>
        /// =================================================================================================
        internal static string RExtToCSV(this DataTable table) => RExtToCSV(table, ",", true);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Converts the passed in data table to a CSV-style string.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="table">Table to convert.</param>
        /// <param name="includeHeader">
        ///     true - include headers<br />
        ///     false - do not include header column.
        /// </param>
        /// <returns>
        ///     Resulting CSV-style string.
        /// </returns>
        /// =================================================================================================
        internal static string RExtToCSV(this DataTable table, bool includeHeader) => RExtToCSV(table, ",", includeHeader);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A DataTable extension method that converts a table to a row data.
        /// </summary>
        /// <param name="table">Table to convert.</param>
        /// <returns>
        ///     The given data converted to a row data.
        /// </returns>
        /// =================================================================================================
        internal static IResult<(
            string, 
            IEnumerable<PropTranslateModel>,
            IEnumerable<PropModel>,
            IEnumerable<IReadOnlyCollection<PropNameValue>>)> ConvertToRowData(this DataTable table)
        {
            var sheetName = (table.TableName ?? "Sheet1").RExtToCleanSheetName();
            var outputProps = new List<PropTranslateModel>();
            var embeddedModelCollection = new List<PropModel>();
            var rows = new List<List<PropNameValue>>();

            var columnNr = table.Columns.Count;
            for (var i = 0; i < columnNr; i++)
            {
                var column = table.Columns[i];

                embeddedModelCollection.Add(new PropModel
                {
                    CommonName = column.ColumnName,
                    DataType = TypeHelper.GetNonNullableType(column.DataType).ToString(),
                    IsNullable = column.AllowDBNull
                }); 
                
                outputProps.Add(new PropTranslateModel
                {
                    CommonName = column.ColumnName,
                    Order = i,
                    TranslateName = column.ColumnName,
                    Format = "General",
                    IsItalic = false,
                    IsBold = false,
                    WrapText = false
                });
            }

            for (var i = 0; i < table.Rows.Count; i++)
            {
                var dictRow = new List<PropNameValue>();
                var rowData = table.Rows[i];
                for (var j = 0; j < columnNr; j++)
                {
                    var columnName = table.Columns[j].ColumnName;
                    var columnValue = rowData[j];

                    dictRow.Add(new PropNameValue(columnName, columnValue.IsDbNull() ? null : columnValue));
                }
                rows.Add(dictRow);
            }

            return Result<(string, IEnumerable<PropTranslateModel>, IEnumerable<PropModel>, IEnumerable<IReadOnlyCollection<PropNameValue>>)>
                .Success((sheetName, outputProps, embeddedModelCollection, rows));
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Converts the passed in data table to a CSV-style string.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="table">Table to convert.</param>
        /// <param name="delimiter">.</param>
        /// <param name="includeHeader">
        ///     true - include headers<br />
        ///     false - do not include header column.
        /// </param>
        /// <returns>
        ///     Resulting CSV-style string.
        /// </returns>
        /// =================================================================================================
        private static string RExtToCSV(this DataTable table, string delimiter, bool includeHeader)
        {
            var result = new StringBuilder();

            if (includeHeader)
            {
                foreach (DataColumn column in table.Columns)
                {
                    result.Append(CsvField(column.ColumnName));
                    result.Append(delimiter);
                }

                if (result.Length >= delimiter.Length) result.Length -= delimiter.Length;
                result.Append(Environment.NewLine);
            }

            foreach (DataRow row in table.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    result.Append(CsvField(item));
                    result.Append(delimiter);
                }

                if (result.Length >= delimiter.Length) result.Length -= delimiter.Length;
                result.Append(Environment.NewLine);
            }

            return result.ToString();
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Formats a single value as a CSV field. The value is quoted and its embedded quotes are
        ///     doubled (RFC 4180), and a leading character that a spreadsheet would interpret as the start
        ///     of a formula is neutralized with a leading apostrophe.
        /// </summary>
        /// <param name="item">The value to format.</param>
        /// <returns>
        ///     The value rendered as a CSV field, or an empty string when the value is missing.
        /// </returns>
        /// =================================================================================================
        private static string CsvField(object item)
        {
            if (item.IsNull() || item.IsDbNull()) return string.Empty;

            var value = Convert.ToString(item, CultureInfo.InvariantCulture) ?? string.Empty;

            if (value.Length > 0
                && "=+-@\t\r".IndexOf(value[0]) >= 0
                && !double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                value = "'" + value;

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}