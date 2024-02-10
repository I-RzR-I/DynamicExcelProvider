// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-25 09:12
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:24
// ***********************************************************************
//  <copyright file="IntExtensions.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DomainCommonExtensions.CommonExtensions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.WorkXCore.Helpers.Resources;
using System;
using System.Runtime.CompilerServices;

#endregion

[assembly: InternalsVisibleTo("WorkXCoreFuncTests")]

namespace DynamicExcelProvider.WorkXCore.Extensions.DataType
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     An int extensions.
    /// </summary>
    /// <remarks>
    ///     25-Jan-24.
    /// </remarks>
    /// =================================================================================================
    internal static class IntExtensions
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     An int? extension method that converts a source to a bool.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">
        ///     Thrown when one or more arguments are outside the required range.
        /// </exception>
        /// <param name="source">The source to act on.</param>
        /// <returns>
        ///     Source as a bool.
        /// </returns>
        /// =================================================================================================
        internal static bool ToBool(this int source)
        {
            if (source < 0)
                throw new ArgumentOutOfRangeException(nameof(source), source, MessagesInfo.InvalidInToToBool);

            return source.IsGreaterThanZero();
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     An int? extension method that converts a source to a bool.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">
        ///     Thrown when one or more arguments are outside the required range.
        /// </exception>
        /// <param name="source">The source to act on.</param>
        /// <returns>
        ///     Source as a bool.
        /// </returns>
        /// =================================================================================================
        internal static bool ToBool(this int? source)
        {
            if (source.IsNull() || source < 0)
                throw new ArgumentOutOfRangeException(nameof(source), source, MessagesInfo.InvalidInToToBool);

            return source > 0;
        }
    }
}