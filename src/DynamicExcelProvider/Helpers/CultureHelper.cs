// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2026-08-24
//
//  Last Modified By : RzR
//  Last Modified On : 2026-08-24
// ***********************************************************************
//  <copyright file="CultureHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
//
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System;
using System.Globalization;

#endregion

namespace DynamicExcelProvider.Helpers
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Turns a culture identifier into a culture without ever throwing.
    /// </summary>
    /// =================================================================================================
    internal static class CultureHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Resolves an LCID to a culture, falling back to the invariant culture when the identifier
        ///     cannot be resolved on the current runtime.
        /// </summary>
        /// <param name="lcid">The culture identifier.</param>
        /// <returns>
        ///     The resolved culture, or <see cref="CultureInfo.InvariantCulture" />.
        /// </returns>
        /// =================================================================================================
        internal static CultureInfo SafeFromLcid(int lcid)
        {
            try
            {
                return CultureInfo.GetCultureInfo(lcid);
            }
            catch (CultureNotFoundException)
            {
                return CultureInfo.InvariantCulture;
            }
            catch (ArgumentOutOfRangeException)
            {
                return CultureInfo.InvariantCulture;
            }
        }
    }
}
