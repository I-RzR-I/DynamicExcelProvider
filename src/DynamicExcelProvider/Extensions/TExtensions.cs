// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-13 12:06
// 
//  Last Modified By : RzR
//  Last Modified On : 2023-03-15 22:54
// ***********************************************************************
//  <copyright file="TExtensions.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System.Collections.Generic;
using System.Linq;
using System.Reflection;
// ReSharper disable InconsistentNaming

#endregion

namespace DynamicExcelProvider.Extensions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     T extensions.
    /// </summary>
    /// =================================================================================================
    internal static class TExtensions
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Append object to array.
        /// </summary>
        /// <typeparam name="T">Type of array.</typeparam>
        /// <param name="first">First item.</param>
        /// <param name="items">All item array.</param>
        /// <returns>
        ///     A T[].
        /// </returns>
        /// =================================================================================================
        internal static T[] AppendTo<T>(this T first, params T[] items)
        {
            var result = new T[items.Length + 1];
            result[0] = first;
            items.CopyTo(result, 1);

            return result;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Get list of properties from TSource.
        /// </summary>
        /// <typeparam name="TSource">Source type.</typeparam>
        /// <param name="source">Source.</param>
        /// <returns>
        ///     The properties information from t.
        /// </returns>
        /// =================================================================================================
        internal static IList<PropertyInfo> GetPropertiesInfoFromT<TSource>(this TSource source)
        {
            try
            {
                return typeof(TSource).GetProperties()
                    .ToList();
            }
            catch { return null; }
        }
    }
}