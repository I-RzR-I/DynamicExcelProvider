// ***********************************************************************
//  Assembly         : RzR.Shared.Export.GeneralDocumentGeneratorTests
//  Author           : RzR
//  Created On       : 2024-02-01 20:43
// 
//  Last Modified By : RzR
//  Last Modified On : 2024-02-07 20:48
// ***********************************************************************
//  <copyright file="DataTempRo.cs" company="">
//   Copyright (c) RzR. All rights reserved.
//  </copyright>
// 
//  <summary>
//  </summary>
// ***********************************************************************

#region U S A G E S

using System;

#endregion

namespace GeneralDocumentGeneratorTests.Models
{
    public class DataTempRo
    {
        public int Identificator { get; set; }
        public string Nume { get; set; }
        public DateTime Inceput { get; set; }
        public DateTime? Sfarsit { get; set; }
        public bool Activ { get; set; }
    }
}