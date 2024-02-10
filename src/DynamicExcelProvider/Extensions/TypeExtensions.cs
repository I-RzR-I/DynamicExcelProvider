// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-17 08:36
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-01-30 21:27
// ***********************************************************************
//  <copyright file="TypeExtensions.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DomainCommonExtensions.CommonExtensions;
using System;
using System.Collections.Generic;

#endregion

namespace DynamicExcelProvider.Extensions
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     A type extensions.
    /// </summary>
    /// <remarks>
    /// </remarks>
    /// =================================================================================================
    public static class TypeExtensions
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A Type extension method that query if 'type' is nullable property type.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <exception cref="ArgumentNullException">
        ///     Thrown when one or more required arguments are null.
        /// </exception>
        /// <param name="type">The type to act on.</param>
        /// <returns>
        ///     True if nullable property type, false if not.
        /// </returns>
        /// =================================================================================================
        public static bool IsNullablePropType(this Type type)
        {
            if (type.IsNull()) throw new ArgumentNullException(nameof(type));

            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A Type extension method that query if 'type' is simple type.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="type">The type to act on.</param>
        /// <returns>
        ///     True if simple type, false if not.
        /// </returns>
        /// =================================================================================================
        public static bool IsSimpleType(this Type type)
        {
            var underlyingType = Nullable.GetUnderlyingType(type);
            type = underlyingType ?? type;
            var simpleTypes = new List<Type>
            {
                typeof(byte),
                typeof(sbyte),
                typeof(short),
                typeof(ushort),
                typeof(int),
                typeof(uint),
                typeof(long),
                typeof(ulong),
                typeof(float),
                typeof(double),
                typeof(decimal),
                typeof(bool),
                typeof(string),
                typeof(char),
                typeof(Guid),
                typeof(DateTime),
                typeof(DateTimeOffset),
                typeof(byte[])
            };

            return simpleTypes.Contains(type) || type.IsEnum;
        }
    }
}