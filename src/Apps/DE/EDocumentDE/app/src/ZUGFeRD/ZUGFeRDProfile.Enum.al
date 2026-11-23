// ------------------------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
// ------------------------------------------------------------------------------------------------
namespace Microsoft.eServices.EDocument.Formats;

/// <summary>
/// The ZUGFeRD / Factur-X profile that the ZUGFeRD export writes.
/// The profile decides the specification identifier (BT-24), the name of the XML file embedded in the PDF
/// and which profile specific elements are written, for example the sub-line structure of EXTENDED.
/// </summary>
enum 13915 "ZUGFeRD Profile"
{
    Access = Public;
    Extensible = false;

    /// <summary>
    /// ZUGFeRD with the XRechnung 3.0 CIUS. This is the default and the behavior before profiles existed.
    /// </summary>
    value(0; XRechnung)
    {
        Caption = 'XRechnung';
    }
    /// <summary>
    /// ZUGFeRD / Factur-X EXTENDED (ZUGFeRD 2.5.2 / Factur-X 1.09.2).
    /// </summary>
    value(1; Extended)
    {
        Caption = 'Extended';
    }
}
