// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2024-01-24 22:41
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 00:23
// ***********************************************************************
//  <copyright file="BoolExtensions.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DomainCommonExtensions.DataTypeExtensions;
using System.Runtime.CompilerServices;

#endregion

[assembly: InternalsVisibleTo("WorkXCoreFuncTests")]

namespace DynamicExcelProvider.WorkXCore.Extensions.DataType
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     An extensions.
    /// </summary>
    /// =================================================================================================
    internal static class BoolExtensions
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A bool extension method that converts a source to an int.
        /// </summary>
        /// <param name="source">The source to act on.</param>
        /// <returns>
        ///     Source as an int.
        /// </returns>
        /// =================================================================================================
        internal static int ToInt(this bool source)
            => source.IsTrue() ? 1 : 0;

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     A bool? extension method that converts a source to an int.
        /// </summary>
        /// <param name="source">The source to act on.</param>
        /// <returns>
        ///     Source as an int.
        /// </returns>
        /// =================================================================================================
        internal static int ToInt(this bool? source)
            => source.IsTrue() ? 1 : 0;
    }
}