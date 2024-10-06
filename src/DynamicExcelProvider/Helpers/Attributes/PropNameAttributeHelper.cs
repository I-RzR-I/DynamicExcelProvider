// ***********************************************************************
//  Assembly         : RzR.Shared.Export.DynamicExcelProvider
//  Author           : RzR
//  Created On       : 2023-03-15 08:24
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-03 10:52
// ***********************************************************************
//  <copyright file="PropNameAttributeHelper.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using DomainCommonExtensions.ArraysExtensions;
using DomainCommonExtensions.CommonExtensions;
using DomainCommonExtensions.DataTypeExtensions;
using DynamicExcelProvider.Attributes;
using DynamicExcelProvider.Models.Internal;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

// ReSharper disable ConstantConditionalAccessQualifier
// ReSharper disable ConstantNullCoalescingCondition

#endregion

namespace DynamicExcelProvider.Helpers.Attributes
{
    /// -------------------------------------------------------------------------------------------------
    /// <summary>
    ///     Helper for property name attribute.
    /// </summary>
    /// =================================================================================================
    internal static class PropNameAttributeHelper
    {
        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Enumerates property name attribute by property in this collection.
        /// </summary>
        /// <typeparam name="T">Generic type parameter.</typeparam>
        /// <param name="propertyName">The propertyName to act on.</param>
        /// <returns>
        ///     An enumerator that allows foreach to be used to process property name attribute by
        ///     property in this collection.
        /// </returns>
        /// =================================================================================================
        internal static IEnumerable<ExcelPropNameAttribute> PropNameAttributeByProperty<T>(this string propertyName)
        {
            var resultList = new List<ExcelPropNameAttribute>();
            try
            {
                var props = typeof(T).GetProperty(propertyName)
                    ?.GetCustomAttributes(typeof(ExcelPropNameAttribute), false).Cast<ExcelPropNameAttribute>();
                if (props.IsNull()) return null;

                resultList.AddRange(props!);
            }
            catch { return null; }

            return resultList;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Enumerates property name attribute by property in this collection.
        /// </summary>
        /// <typeparam name="T">Generic type parameter.</typeparam>
        /// <param name="propertyName">The propertyName to act on.</param>
        /// <param name="cultureId">Identifier for the culture.</param>
        /// <returns>
        ///     An enumerator that allows foreach to be used to process property name attribute by
        ///     property in this collection.
        /// </returns>
        /// =================================================================================================
        internal static IEnumerable<ExcelPropNameAttribute> PropNameAttributeByProperty<T>(this string propertyName, int cultureId)
        {
            var resultList = new List<ExcelPropNameAttribute>();
            try
            {
                var props = typeof(T).GetProperty(propertyName)
                    ?.GetCustomAttributes(typeof(ExcelPropNameAttribute), false).Cast<ExcelPropNameAttribute>()
                    .Where(x => Equals(x.CultureInfo, new CultureInfo(cultureId)));
                if (props.IsNull()) return null;

                resultList.AddRange(props!);
            }
            catch { return null; }

            return resultList;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Enumerates property name attribute in this collection.
        /// </summary>
        /// <typeparam name="T">Generic type parameter.</typeparam>
        /// <returns>
        ///     An enumerator that allows foreach to be used to process property name attribute in this
        ///     collection.
        /// </returns>
        /// =================================================================================================
        internal static IEnumerable<ExcelPropNameAttribute> PropNameAttribute<T>()
        {
            var resultList = new List<ExcelPropNameAttribute>();
            try
            {
                foreach (var propertyInfo in typeof(T).GetProperties())
                {
                    var props = propertyInfo
                        ?.GetCustomAttributes(typeof(ExcelPropNameAttribute), false).Cast<ExcelPropNameAttribute>();
                    if (props.IsNull()) return null;

                    resultList.AddRange(props);
                }
            }
            catch { return null; }

            return resultList;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Enumerates property name attribute in this collection.
        /// </summary>
        /// <typeparam name="T">Generic type parameter.</typeparam>
        /// <param name="cultureId">Identifier for the culture.</param>
        /// <returns>
        ///     An enumerator that allows foreach to be used to process property name attribute in this
        ///     collection.
        /// </returns>
        /// =================================================================================================
        internal static IEnumerable<ExcelPropNameAttribute> PropNameAttribute<T>(int cultureId)
        {
            var resultList = new List<ExcelPropNameAttribute>();
            try
            {
                foreach (var propertyInfo in typeof(T).GetProperties())
                {
                    var props = propertyInfo
                        ?.GetCustomAttributes(typeof(ExcelPropNameAttribute), false).Cast<ExcelPropNameAttribute>()
                        .Where(x => Equals(x.CultureInfo, new CultureInfo(cultureId)));
                    if (props.IsNull()) return null;

                    resultList.AddRange(props);
                }
            }
            catch { return null; }

            return resultList;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets the property name attributes in this collection.
        /// </summary>
        /// <typeparam name="T">Generic type parameter.</typeparam>
        /// <param name="cultureId">Identifier for the culture.</param>
        /// <returns>
        ///     An enumerator that allows foreach to be used to process the property name attributes in
        ///     this collection.
        /// </returns>
        /// =================================================================================================
        internal static IEnumerable<ParseModelProperty> GetPropNameAttribute<T>(int cultureId)
        {
            var resultList = new List<ParseModelProperty>();
            try
            {
                foreach (var propertyInfo in typeof(T).GetProperties())
                {
                    var props = propertyInfo
                        ?.GetCustomAttributes(typeof(ExcelPropNameAttribute), false).Cast<ExcelPropNameAttribute>()
                        .Where(x => Equals(x.CultureInfo, new CultureInfo(cultureId)))
                        .Select(x => new ParseModelProperty
                        {
                            EmbeddedName = propertyInfo.Name.IsNullOrEmpty() ? x.PropertyName : propertyInfo.Name,
                            CurrentName = x.PropertyName,
                            Order = x.Order,
                            InResult = x.InResult,
                            CultureInfo = x.CultureInfo,
                            FormatCode = x.FormatCode,
                            IsItalic = x.IsItalic,
                            IsBold = x.IsBold,
                            WrapText = x.WrapText
                        });
                    if (props.IsNull()) return null;

                    resultList.AddRange(props);
                }
            }
            catch { return null; }

            return resultList;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets the property name attributes in this collection.
        /// </summary>
        /// <typeparam name="T">Generic type parameter.</typeparam>
        /// <param name="cultureId">(Optional) Identifier for the culture.</param>
        /// <returns>
        ///     An enumerator that allows foreach to be used to process the property name attributes in
        ///     this collection.
        /// </returns>
        /// =================================================================================================
        internal static IEnumerable<ParseModelProperty> GetPropNameAttributeByPassMissing<T>(int cultureId = 1033)
        {
            var resultList = new List<ParseModelProperty>();
            try
            {
                var hasAttributeByCulture = typeof(T).GetProperties().Any(x => x.GetCustomAttributes(typeof(ExcelPropNameAttribute), false).Cast<ExcelPropNameAttribute>()
                    .Where(z => Equals(z.CultureInfo, new CultureInfo(cultureId))).ToList().Count > 0);
                foreach (var (item, index) in typeof(T).GetProperties().WithIndex())
                {
                    var propAttributes = item.GetCustomAttributes(typeof(ExcelPropNameAttribute), false).Cast<ExcelPropNameAttribute>().ToList();
                    if (propAttributes.IsNullOrEmptyEnumerable().IsFalse())
                    {
                        var existAttributeWithCulture = propAttributes.Any(x => Equals(x.CultureInfo, new CultureInfo(cultureId)));
                        if (existAttributeWithCulture.IsTrue())
                        {
                            var propAttribute = propAttributes.FirstOrDefault(x => Equals(x.CultureInfo, new CultureInfo(cultureId)));
                            resultList.Add(new ParseModelProperty()
                            {
                                CultureInfo = new CultureInfo(cultureId),
                                CurrentName = propAttribute!.PropertyName.IsNullOrEmpty() ? item.Name : propAttribute?.PropertyName,
                                InResult = propAttribute?.InResult ?? true,
                                EmbeddedName = item.Name,
                                Order = propAttribute?.Order ?? index,
                                FormatCode = propAttribute.FormatCode,
                                IsItalic = propAttribute.IsItalic,
                                IsBold = propAttribute.IsBold,
                                WrapText = propAttribute.WrapText
                            });
                        }
                        else
                        {
                            resultList.Add(new ParseModelProperty()
                            {
                                CultureInfo = new CultureInfo(cultureId),
                                CurrentName = item.Name,
                                InResult = true,
                                EmbeddedName = item.Name,
                                Order = index
                            });

                        }
                    }
                    else
                    {
                        if (hasAttributeByCulture.IsFalse())
                        {
                            resultList.Add(new ParseModelProperty()
                            {
                                CultureInfo = new CultureInfo(cultureId),
                                CurrentName = item.Name,
                                InResult = true,
                                EmbeddedName = item.Name,
                                Order = index
                            });
                        }
                    }
                }
            }
            catch { return null; }

            return resultList;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets the property name attributes in this collection.
        /// </summary>
        /// <param name="cultureId">Identifier for the culture.</param>
        /// <param name="type">The type.</param>
        /// <returns>
        ///     An enumerator that allows foreach to be used to process the property name attributes in
        ///     this collection.
        /// </returns>
        /// =================================================================================================
        internal static IEnumerable<ParseModelProperty> GetPropNameAttribute(int cultureId, Type type)
        {
            var resultList = new List<ParseModelProperty>();
            try
            {
                foreach (var propertyInfo in type.GetProperties())
                {
                    var props = propertyInfo
                        ?.GetCustomAttributes(typeof(ExcelPropNameAttribute), false).Cast<ExcelPropNameAttribute>()
                        .Where(x => Equals(x.CultureInfo, new CultureInfo(cultureId)))
                        .Select(x => new ParseModelProperty
                        {
                            EmbeddedName = propertyInfo.Name.IsNullOrEmpty() ? x.PropertyName : propertyInfo.Name,
                            CurrentName = x.PropertyName,
                            Order = x.Order,
                            InResult = x.InResult,
                            CultureInfo = x.CultureInfo,
                            FormatCode = x.FormatCode,
                            IsItalic = x.IsItalic,
                            IsBold = x.IsBold,
                            WrapText = x.WrapText
                        });
                    if (props.IsNull()) return null;

                    resultList.AddRange(props);
                }
            }
            catch { return null; }

            return resultList;
        }

        /// -------------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets the property name attributes in this collection.
        /// </summary>
        /// <param name="cultureId">Identifier for the culture.</param>
        /// <param name="type">The type.</param>
        /// <returns>
        ///     An enumerator that allows foreach to be used to process the property name attributes in
        ///     this collection.
        /// </returns>
        /// =================================================================================================
        internal static IEnumerable<ParseModelProperty> GetPropNameAttributeByPassMissing(int cultureId, Type type)
        {
            var resultList = new List<ParseModelProperty>();
            try
            {
                var hasAttributeByCulture = type.GetProperties().Any(x => x.GetCustomAttributes(typeof(ExcelPropNameAttribute), false).Cast<ExcelPropNameAttribute>()
                    .Where(z => Equals(z.CultureInfo, new CultureInfo(cultureId))).ToList().Count > 0);
                foreach (var (item, index) in type.GetProperties().WithIndex())
                {
                    var propAttributes = item.GetCustomAttributes(typeof(ExcelPropNameAttribute), false).Cast<ExcelPropNameAttribute>().ToList();
                    if (propAttributes.IsNullOrEmptyEnumerable().IsFalse())
                    {
                        var existAttributeWithCulture = propAttributes.Any(x => Equals(x.CultureInfo, new CultureInfo(cultureId)));
                        if (existAttributeWithCulture.IsTrue())
                        {
                            var propAttribute = propAttributes.FirstOrDefault(x => Equals(x.CultureInfo, new CultureInfo(cultureId)));
                            resultList.Add(new ParseModelProperty()
                            {
                                CultureInfo = new CultureInfo(cultureId),
                                CurrentName = propAttribute!.PropertyName.IsNullOrEmpty() ? item.Name : propAttribute.PropertyName,
                                InResult = propAttribute.InResult,
                                EmbeddedName = item.Name,
                                Order = propAttribute?.Order ?? index,
                                FormatCode = propAttribute.FormatCode,
                                IsItalic = propAttribute.IsItalic,
                                IsBold = propAttribute.IsBold,
                                WrapText = propAttribute.WrapText
                            });
                        }
                        else
                        {
                            resultList.Add(new ParseModelProperty()
                            {
                                CultureInfo = new CultureInfo(cultureId),
                                CurrentName = item.Name,
                                InResult = true,
                                EmbeddedName = item.Name,
                                Order = index
                            });

                        }
                    }
                    else
                    {
                        if (hasAttributeByCulture.IsFalse())
                        {
                            resultList.Add(new ParseModelProperty()
                            {
                                CultureInfo = new CultureInfo(cultureId),
                                CurrentName = item.Name,
                                InResult = true,
                                EmbeddedName = item.Name,
                                Order = index
                            });
                        }
                    }
                }
            }
            catch { return null; }

            return resultList;
        }
    }
}