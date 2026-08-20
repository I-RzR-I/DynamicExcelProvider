// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-15 14:53
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:35
// ***********************************************************************
//  <copyright file="SpreadsheetCellValueHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DocumentFormat.OpenXml.Spreadsheet;
using DynamicExcelProvider.WorkXCore.Enums;
using DynamicExcelProvider.WorkXCore.Helpers.Resources;
using RzR.Extensions.Domain.Primitives;
using RzR.ResultMessage;
using RzR.ResultMessage.Abstractions;
using RzR.ResultMessage.Extensions.Result;
using System;
using System.Globalization;

#endregion

namespace DynamicExcelProvider.WorkXCore.Helpers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A spreadsheet cell value helper.
    /// </summary>
    /// =================================================================================================
    internal static class SpreadsheetCellValueHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Builds cell value.
        /// </summary>
        /// <param name="value">The value to act on.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <param name="cellDataType">Type of the cell data.</param>
        /// <param name="sourceCellDataType">Type of the source cell data.</param>
        /// <returns>
        ///     An IResult&lt;CellValue&gt;
        /// </returns>
        /// =================================================================================================
        internal static IResult<CellValue> BuildCellValue(
            object value, object defaultValue,
            CellDataType cellDataType, SourceCellDataType sourceCellDataType)
        {
            try
            {
                if (value.IsNullOrDbNull())
                {
                    if (defaultValue.IsNullOrDbNull().IsFalse())
                    {
                        var castDefaultValue = defaultValue.CastObjectToCellValue(sourceCellDataType);

                        return castDefaultValue.IsSuccess.IsFalse()
                            ? Result<CellValue>.Failure(castDefaultValue.GetFirstMessage())
                            : castDefaultValue;
                    }

                    return Result<CellValue>.Success(new CellValue());
                }

                return value.CastObjectToCellValue(sourceCellDataType);
            }
            catch (Exception e)
            {
                return Result<CellValue>.Failure(e.Message).WithError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     An object extension method that cast object to cell value.
        /// </summary>
        /// <param name="value">The value to act on.</param>
        /// <param name="sourceCellDataType">Type of the source cell data.</param>
        /// <returns>
        ///     An IResult&lt;CellValue&gt;
        /// </returns>
        /// =================================================================================================
        private static IResult<CellValue> CastObjectToCellValue(this object value, SourceCellDataType sourceCellDataType)
        {
            try
            {
                return sourceCellDataType switch
                {
                    SourceCellDataType.DateTime 
                        => Result<CellValue>.Success(new CellValue(Convert.ToDateTime(value, CultureInfo.InvariantCulture))),
                    SourceCellDataType.String 
                        => Result<CellValue>.Success(new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture))),
                    SourceCellDataType.Decimal
                        => Result<CellValue>.Success(new CellValue(Convert.ToDecimal(value, CultureInfo.InvariantCulture))),
                    SourceCellDataType.Float
                        => Result<CellValue>.Success(new CellValue(Convert.ToDouble(value, CultureInfo.InvariantCulture))),
                    SourceCellDataType.Long 
                        => Result<CellValue>.Success(new CellValue(Convert.ToDecimal(value, CultureInfo.InvariantCulture))),
                    SourceCellDataType.Int 
                        => Result<CellValue>.Success(new CellValue(Convert.ToInt32(value, CultureInfo.InvariantCulture))),
                    SourceCellDataType.Short
                        => Result<CellValue>.Success(new CellValue(Convert.ToInt32(value, CultureInfo.InvariantCulture))),
                    SourceCellDataType.Boolean
                        => Result<CellValue>.Success(new CellValue(Convert.ToBoolean(value, CultureInfo.InvariantCulture))),
                    _ => Result<CellValue>.Failure(string.Format(MessagesInfo.InvalidDataSourceType, sourceCellDataType))
                };
            }
            catch (Exception e)
            {
                return Result<CellValue>.Failure(e.Message).WithError(e);
            }
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     An object extension method that determines whether the supplied value represents a missing
        ///     value. Values coming from a <see cref="System.Data.DataTable" /> use
        ///     <see cref="DBNull.Value" /> instead of a CLR null reference, so both forms are treated the
        ///     same way.
        /// </summary>
        /// <param name="value">The value to act on.</param>
        /// <returns>
        ///     True if the value is null or <see cref="DBNull.Value" />, false if not.
        /// </returns>
        /// =================================================================================================
        private static bool IsNullOrDbNull(this object value) => value.IsNull() || value.IsDbNull();
    }
}