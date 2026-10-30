// ------------------------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
// ------------------------------------------------------------------------------------------------
namespace Microsoft.eServices.EDocument.Formats;

/// <summary>
/// The sub-line type (BT-X-8, LineStatusReasonCode) of an invoice line in the ZUGFeRD EXTENDED profile.
/// It is written only for the EXTENDED profile.
/// </summary>
enum 13916 "ZUGFeRD Line Subtype"
{
    Access = Public;
    Extensible = false;

    /// <summary>
    /// No sub-line type is written. A line with a parent line is written as DETAIL (BR-FXEXT-06).
    /// </summary>
    value(0; " ")
    {
        Caption = ' ', Locked = true;
    }
    /// <summary>
    /// The line counts for the totals and the VAT breakdown.
    /// </summary>
    value(1; Detail)
    {
        Caption = 'Detail';
    }
    /// <summary>
    /// The line is a subtotal of its DETAIL and GROUP children.
    /// </summary>
    value(2; Group)
    {
        Caption = 'Group';
    }
    /// <summary>
    /// The line is for information only and does not count for the totals.
    /// </summary>
    value(3; Information)
    {
        Caption = 'Information';
    }
}
