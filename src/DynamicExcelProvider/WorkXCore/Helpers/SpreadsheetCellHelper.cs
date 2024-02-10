// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-15 00:18
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:34
// ***********************************************************************
//  <copyright file="SpreadsheetCellHelper.cs" company="">
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
using DomainCommonExtensions.ArraysExtensions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.WorkXCore.Extensions;
using DynamicExcelProvider.WorkXCore.Models;
using System;
using System.Collections.Generic;

#endregion

namespace DynamicExcelProvider.WorkXCore.Helpers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A spreadsheet cell helper.
    /// </summary>
    /// =================================================================================================
    public static class SpreadsheetCellHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Serialize as row.
        /// </summary>
        /// <param name="cellDefinitions">The cell definitions.</param>
        /// <param name="cellDataDefinitions">The cell data definitions.</param>
        /// <param name="rowIndex">Zero-based index of the row.</param>
        /// <returns>
        ///     An IResult&lt;Row&gt;
        /// </returns>
        /// =================================================================================================
        internal static IResult<Row> SerializeAsRow(
            IEnumerable<CellValueDefinition> cellDefinitions,
            CellDataDefinition[] cellDataDefinitions, int rowIndex)
        {
            try
            {
                var row = new Row { RowIndex = (uint)rowIndex };
                foreach (var (currentCellItem, currentCellIndex) in cellDefinitions.WithIndex())
                {
                    var xCell = new Cell();

                    var cellBuildResult = xCell.BuildCellBody(currentCellItem, cellDataDefinitions[currentCellIndex],
                        currentCellIndex, rowIndex);
                    if (cellBuildResult.IsSuccess.IsFalse())
                        return Result<Row>.Failure(cellBuildResult.GetFirstMessage());

                    row.Append(cellBuildResult.Response);
                }

                return Result<Row>.Success(row);
            }
            catch (Exception e)
            {
                return Result<Row>.Failure().WithError(e);
            }
        }
    }
}