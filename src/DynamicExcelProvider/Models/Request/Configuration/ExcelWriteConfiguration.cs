// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-15 22:39
// 
//  Last Modified By : RzR
//  Last Modified On : 2023-03-15 22:39
// ***********************************************************************
//  <copyright file="ExcelWriteConfiguration.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

// ReSharper disable InconsistentNaming
namespace DynamicExcelProvider.Models.Request.Configuration
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Excel write configuration.
    /// </summary>
    /// =================================================================================================
    public class ExcelWriteConfiguration
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Excel sheet name.
        /// </summary>
        /// =================================================================================================
        public string SheetName = "Sheet1";

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets or sets the lcid.
        /// </summary>
        /// <value>
        ///     The lcid.
        /// </value>
        /// =================================================================================================
        public int LCID { get; set; } = 1033;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelWriteConfiguration"/> class.
        /// </summary>
        /// =================================================================================================
        public ExcelWriteConfiguration()
        {
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelWriteConfiguration"/> class.
        /// </summary>
        /// <param name="sheetName">Excel sheet name.</param>
        /// =================================================================================================
        public ExcelWriteConfiguration(string sheetName) => SheetName = sheetName;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelWriteConfiguration"/> class.
        /// </summary>
        /// <param name="sheetName">Excel sheet name.</param>
        /// <param name="lcid">The lcid.</param>
        /// =================================================================================================
        public ExcelWriteConfiguration(string sheetName, int lcid)
        {
            SheetName = sheetName;
            LCID = lcid;
        }
    }
}