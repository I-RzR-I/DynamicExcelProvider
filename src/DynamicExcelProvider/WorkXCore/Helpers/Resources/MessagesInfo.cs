// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-12 21:47
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-01-28 11:15
// ***********************************************************************
//  <copyright file="MessagesInfo.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

namespace DynamicExcelProvider.WorkXCore.Helpers.Resources
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>Information about the messages.</summary>
    /// =================================================================================================
    internal static class MessagesInfo
    {
        /// <summary>(Immutable) the sheet name exist.</summary>
        internal const string SheetNameExist = "The supplied name ({0}) is not a valid XLSX worksheet name.";

        /// <summary>(Immutable) name of the invalid sheet.</summary>
        internal const string InvalidSheetName = "The supplied name ({0}) is not a valid XLSX worksheet name.";

        /// <summary>(Immutable) the column reference out of range.</summary>
        internal const string ColumnReferenceOutOfRange = "Column reference ({0}) is out of range.";

        /// <summary>(Immutable) the invalid cell format.</summary>
        internal const string InvalidCellFormat = "The supplied cell format ({0}) is not present in StandardFormats.";

        /// <summary>(Immutable) type of the invalid data source.</summary>
        internal const string InvalidDataSourceType = "The supplied data type ({0}) is not supported.";

        /// <summary>(Immutable) the invalid in to bool.</summary>
        internal const string InvalidInToToBool = "Source data must be greater than 0!";
    }
}