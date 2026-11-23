// ------------------------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
// ------------------------------------------------------------------------------------------------
namespace Microsoft.eServices.EDocument.Formats;

/// <summary>
/// Buffer for the document notes (BG-1, IncludedNote) that subscribers add to the ZUGFeRD export.
/// </summary>
table 13914 "ZUGFeRD Note"
{
    Access = Public;
    Caption = 'ZUGFeRD Note';
    DataClassification = CustomerContent;
    TableType = Temporary;
    InherentEntitlements = RIMDX;
    InherentPermissions = RIMDX;

    fields
    {
        field(1; "Entry No."; Integer)
        {
            Caption = 'Entry No.';
            ToolTip = 'Specifies the order of the note in the document.';
        }
        field(2; "Subject Code"; Code[10])
        {
            Caption = 'Subject Code';
            ToolTip = 'Specifies the subject of the note (BT-21) as a UNTDID 4451 code, for example AAI, ADU, REG or SUR.';
        }
        field(3; Content; Text[2048])
        {
            Caption = 'Content';
            ToolTip = 'Specifies the text of the note (BT-22).';
        }
    }

    keys
    {
        key(PK; "Entry No.")
        {
            Clustered = true;
        }
    }

    /// <summary>
    /// Adds a note at the end of the buffer.
    /// </summary>
    /// <param name="SubjectCode">The UNTDID 4451 subject code, for example ADU or REG. Can be blank.</param>
    /// <param name="NoteContent">The text of the note.</param>
    procedure AddNote(SubjectCode: Code[10]; NoteContent: Text[2048])
    var
        LastEntryNo: Integer;
    begin
        Rec.Reset();
        if Rec.FindLast() then
            LastEntryNo := Rec."Entry No.";
        Rec.Init();
        Rec."Entry No." := LastEntryNo + 1;
        Rec."Subject Code" := SubjectCode;
        Rec.Content := NoteContent;
        Rec.Insert();
    end;
}
