// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-15 19:15
// 
//  Last Modified By : RzR
//  Last Modified On : 2023-03-15 22:55
// ***********************************************************************
//  <copyright file="ObjectExtensions.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System.Collections.Generic;
using System.Reflection;

#endregion

namespace DynamicExcelProvider.Extensions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Object extensions.
    /// </summary>
    /// <remarks>
    ///     30-Jan-24.
    /// </remarks>
    /// =================================================================================================
    internal static class ObjectExtensions
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Get properties from object.
        /// </summary>
        /// <param name="obj">Object.</param>
        /// <returns>
        ///     An enumerator that allows foreach to be used to process the properties in this collection.
        /// </returns>
        /// =================================================================================================
        internal static IEnumerable<PropertyInfo> GetProperties(this object obj) => obj.GetType().GetProperties();
    }
}