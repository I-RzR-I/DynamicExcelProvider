// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-28 19:46
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-01-30 00:24
// ***********************************************************************
//  <copyright file="SpreadsheetDocumentService.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using AggregatedGenericResultMessage.Abstractions;
using DynamicExcelProvider.WorkXCore.Abstractions;
using DynamicExcelProvider.WorkXCore.Helpers;
using DynamicExcelProvider.WorkXCore.Models;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

#endregion

// ReSharper disable ClassWithVirtualMembersNeverInherited.Global

namespace DynamicExcelProvider.WorkXCore.Services
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A service for accessing spreadsheet documents information.
    /// </summary>
    /// <seealso cref="T:DynamicExcelProvider.WorkXCore.Abstractions.ISpreadsheetDocumentService"/>
    ///
    /// ### <inheritdoc cref="ISpreadsheetDocumentService"/>
    /// =================================================================================================
    public class SpreadsheetDocumentService : ISpreadsheetDocumentService
    {
        /// <inheritdoc/>
        public IResult WriteFile(string filePath, WorkbookDefinition workBook)
        {
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            return WriteFile(fs, workBook);
        }

        /// <inheritdoc/>
        public virtual async Task<IResult> WriteFileAsync(
            string filePath, WorkbookDefinition workBook,
            CancellationToken cancellationToken = default)
        {
            using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            return await WriteFileAsync(fs, workBook, cancellationToken);
        }

        /// <inheritdoc/>
        public IResult WriteFile(Stream stream, WorkbookDefinition workBook)
            => SpreadsheetDocumentHelper.Instance.Write(stream, workBook);

        /// <inheritdoc/>
        public async Task<IResult> WriteFileAsync(
            Stream stream, WorkbookDefinition workBook,
            CancellationToken cancellationToken = default)
            => await Task.Run(() => SpreadsheetDocumentHelper.Instance.Write(stream, workBook), cancellationToken);
    }
}