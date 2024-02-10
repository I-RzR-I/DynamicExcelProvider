// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-28 19:47
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-01-28 20:19
// ***********************************************************************
//  <copyright file="ISpreadsheetDocumentService.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using AggregatedGenericResultMessage.Abstractions;
using DynamicExcelProvider.WorkXCore.Models;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

#endregion

namespace DynamicExcelProvider.WorkXCore.Abstractions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Interface for spreadsheet document service.
    /// </summary>
    /// =================================================================================================
    public interface ISpreadsheetDocumentService
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Writes a excel (spreadsheet, worksheet) file.
        /// </summary>
        /// <param name="filePath">Full pathname of the file.</param>
        /// <param name="workBook">The work book.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        IResult WriteFile(string filePath, WorkbookDefinition workBook);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Writes a excel (spreadsheet, worksheet) file asynchronous.
        /// </summary>
        /// <param name="filePath">Full pathname of the file.</param>
        /// <param name="workBook">The work book.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The write file.
        /// </returns>
        /// =================================================================================================
        Task<IResult> WriteFileAsync(string filePath, WorkbookDefinition workBook, CancellationToken cancellationToken = default);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Writes a excel (spreadsheet, worksheet) file.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="workBook">The work book.</param>
        /// <returns>
        ///     An IResult.
        /// </returns>
        /// =================================================================================================
        IResult WriteFile(Stream stream, WorkbookDefinition workBook);

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Writes a excel (spreadsheet, worksheet) file asynchronous.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="workBook">The work book.</param>
        /// <param name="cancellationToken">
        ///     (Optional) A token that allows processing to be cancelled.
        /// </param>
        /// <returns>
        ///     The write file.
        /// </returns>
        /// =================================================================================================
        Task<IResult> WriteFileAsync(Stream stream, WorkbookDefinition workBook, CancellationToken cancellationToken = default);
    }
}