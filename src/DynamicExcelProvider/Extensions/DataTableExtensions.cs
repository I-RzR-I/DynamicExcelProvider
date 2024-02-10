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

using System;
using System.Data;
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
        internal static string ToCSV(this DataTable table) => ToCSV(table, ",", true);

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
        internal static string ToCSV(this DataTable table, bool includeHeader) => ToCSV(table, ",", includeHeader);

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
        private static string ToCSV(this DataTable table, string delimiter, bool includeHeader)
        {
            var result = new StringBuilder();

            if (includeHeader)
            {
                foreach (DataColumn column in table.Columns)
                {
                    result.Append(column.ColumnName);
                    result.Append(delimiter);
                }

                result.Remove(--result.Length, 0);
                result.Append(Environment.NewLine);
            }

            foreach (DataRow row in table.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    if (item is DBNull)
                    {
                        result.Append(delimiter);
                    }
                    else
                    {
                        var itemAsString = item.ToString();
                        itemAsString = itemAsString.Replace("\"", "\"\"");
                        itemAsString = "\"" + itemAsString + "\"";
                        result.Append(itemAsString + delimiter);
                    }
                }

                result.Remove(--result.Length, 0);
                result.Append(Environment.NewLine);
            }

            return result.ToString();
        }
    }
}