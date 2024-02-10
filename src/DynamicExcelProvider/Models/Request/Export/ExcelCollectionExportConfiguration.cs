// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-18 11:38
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-02 23:42
// ***********************************************************************
//  <copyright file="ExcelCollectionExportConfiguration.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.Models.Request.Configuration;
using System.Collections.Generic;

#endregion

namespace DynamicExcelProvider.Models.Request.Export
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     An excel collection export configuration.
    /// </summary>
    /// =================================================================================================
    public class ExcelCollectionExportConfiguration
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets a collection of data.
        /// </summary>
        /// <value>
        ///     A collection of data.
        /// </value>
        /// =================================================================================================
        public IReadOnlyCollection<dynamic> DataCollection { get; set; }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the configuration.
        /// </summary>
        /// <value>
        ///     The configuration.
        /// </value>
        /// =================================================================================================
        public ExcelWriteConfiguration Configuration { get; set; }
    }
}