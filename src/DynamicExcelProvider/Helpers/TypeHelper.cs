// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-17 08:37
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-03 10:50
// ***********************************************************************
//  <copyright file="TypeHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DynamicExcelProvider.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

#endregion

namespace DynamicExcelProvider.Helpers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A type helper.
    /// </summary>
    /// =================================================================================================
    internal class TypeHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     (Immutable) dictionary of nullable types.
        /// </summary>
        /// =================================================================================================
        internal static readonly Dictionary<Type, Type> NullableTypeDict = new Dictionary<Type, Type>
        {
            [typeof(byte?)] = typeof(byte),
            [typeof(sbyte?)] = typeof(sbyte),
            [typeof(short?)] = typeof(short),
            [typeof(ushort?)] = typeof(ushort),
            [typeof(int?)] = typeof(int),
            [typeof(uint?)] = typeof(uint),
            [typeof(long?)] = typeof(long),
            [typeof(ulong?)] = typeof(ulong),
            [typeof(float?)] = typeof(float),
            [typeof(double?)] = typeof(double),
            [typeof(decimal?)] = typeof(decimal),
            [typeof(bool?)] = typeof(bool),
            [typeof(char?)] = typeof(char),
            [typeof(Guid?)] = typeof(Guid),
            [typeof(DateTime?)] = typeof(DateTime),
            [typeof(DateTimeOffset?)] = typeof(DateTimeOffset),
            [typeof(TimeSpan?)] = typeof(TimeSpan)
        };

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Get non nullable type.
        /// </summary>
        /// <param name="type">.</param>
        /// <returns>
        ///     The non nullable type.
        /// </returns>
        /// =================================================================================================
        internal static Type GetNonNullableType(Type type) => type.IsNullablePropType()
            ? NullableTypeDict.FirstOrDefault(x => x.Key == type).Value
            : type;
    }
}