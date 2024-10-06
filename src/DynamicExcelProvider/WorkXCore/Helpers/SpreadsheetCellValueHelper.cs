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

using AggregatedGenericResultMessage;
using AggregatedGenericResultMessage.Abstractions;
using AggregatedGenericResultMessage.Extensions.Result;
using DocumentFormat.OpenXml.Spreadsheet;
using DomainCommonExtensions.CommonExtensions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.WorkXCore.Enums;
using DynamicExcelProvider.WorkXCore.Helpers.Resources;
using System;

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
                if (value.IsNull())
                {
                    if (defaultValue.IsNotNull())
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
                    SourceCellDataType.DateTime => Result<CellValue>.Success(new CellValue((DateTime)value)),
                    SourceCellDataType.String => Result<CellValue>.Success(new CellValue((string)value)),
                    SourceCellDataType.Decimal => Result<CellValue>.Success(new CellValue((decimal)value)),
                    SourceCellDataType.Long => Result<CellValue>.Success(new CellValue((int)value)),
                    SourceCellDataType.Int => Result<CellValue>.Success(new CellValue((int)value)),
                    SourceCellDataType.Short => Result<CellValue>.Success(new CellValue((int)value)),
                    SourceCellDataType.Boolean => Result<CellValue>.Success(new CellValue((bool)value)),
                    _ => Result<CellValue>.Failure(string.Format(MessagesInfo.InvalidDataSourceType, sourceCellDataType))
                };
            }
            catch (Exception e)
            {
                return Result<CellValue>.Failure(e.Message).WithError(e);
            }
        }
    }
}