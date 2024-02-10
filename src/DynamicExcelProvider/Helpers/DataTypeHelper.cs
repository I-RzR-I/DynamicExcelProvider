// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-13 12:04
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-03 10:54
// ***********************************************************************
//  <copyright file="DataTypeHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.WorkXCore.Enums;
using System;

// ReSharper disable UnusedParameter.Global
// ReSharper disable RedundantCaseLabel
// ReSharper disable RedundantAssignment
// ReSharper disable MethodOverloadWithOptionalParameter

#endregion

namespace DynamicExcelProvider.Helpers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A data type helper.
    /// </summary>
    /// =================================================================================================
    internal static class DataTypeHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets column type.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="isNullable">(Optional) True if is nullable, false if not.</param>
        /// <returns>
        ///     The column type.
        /// </returns>
        /// =================================================================================================
        internal static Type GetColumnType(string type, bool isNullable = false)
        {
            type = type.StartsWith("System.")
                ? type.Replace("System.", "")
                : type;

            Type columnType;
            switch (type.ToLower())
            {
                case "tinyint":
                case "smallint":
                case "short":
                    columnType = typeof(short);
                    break;

                case "int":
                case "int32":
                    columnType = typeof(int);
                    break;

                case "int64":
                case "long":
                case "bigint":
                    columnType = typeof(long);
                    break;

                case "bool":
                case "bit":
                case "boolean":
                    columnType = typeof(bool);
                    break;

                case "decimal":
                case "numeric":
                    columnType = typeof(decimal);
                    break;

                case "double":
                    columnType = typeof(double);
                    break;

                case "float":
                case "real":
                    columnType = typeof(float);
                    break;

                case "date":
                case "datetime":
                    columnType = typeof(DateTime);
                    break;

                case "string":
                case "char":
                case "varchar":
                case "nvarchar":
                case "text":
                default:
                    columnType = typeof(string);
                    break;
            }

            return columnType;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets column type.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns>
        ///     The column type.
        /// </returns>
        /// =================================================================================================
        internal static CellDataType GetColumnType(string type)
        {
            type = type.StartsWith("System.")
                ? type.Replace("System.", "")
                : type;

            var columnType = CellDataType.String;
            switch (type.ToLower())
            {
                case "tinyint":
                case "smallint":
                case "short":
                case "int":
                case "int32":
                case "int64":
                case "long":
                case "bigint":
                case "decimal":
                case "numeric":
                case "double":
                case "float":
                case "real":
                    columnType = CellDataType.Number;
                    break;

                case "bool":
                case "bit":
                case "boolean":
                    columnType = CellDataType.Boolean;
                    break;

                case "date":
                case "datetime":
                    columnType = CellDataType.Date;
                    break;

                case "string":
                case "char":
                case "varchar":
                case "nvarchar":
                case "text":
                default:
                    columnType = CellDataType.String;
                    break;
            }

            return columnType;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets source column type.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns>
        ///     The source column type.
        /// </returns>
        /// =================================================================================================
        internal static SourceCellDataType GetSourceColumnType(string type)
        {
            type = type.StartsWith("System.")
                ? type.Replace("System.", "")
                : type;

            var columnType = SourceCellDataType.String;
            switch (type.ToLower())
            {
                case "tinyint":
                case "smallint":
                case "short":
                case "int":
                case "int32":
                case "int64":
                case "long":
                case "bigint":
                case "decimal":
                case "numeric":
                case "double":
                case "float":
                case "real":
                    columnType = SourceCellDataType.Int;
                    break;

                case "bool":
                case "bit":
                case "boolean":
                    columnType = SourceCellDataType.Boolean;
                    break;

                case "date":
                case "datetime":
                    columnType = SourceCellDataType.DateTime;
                    break;

                case "string":
                case "char":
                case "varchar":
                case "nvarchar":
                case "text":
                default:
                    columnType = SourceCellDataType.String;
                    break;
            }

            return columnType;
        }
    }
}