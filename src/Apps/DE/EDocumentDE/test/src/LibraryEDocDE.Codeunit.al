// ------------------------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
// ------------------------------------------------------------------------------------------------
namespace Microsoft.eServices.EDocument.Formats;

using Microsoft.Bank.BankAccount;
using Microsoft.Bank.DirectDebit;
using Microsoft.eServices.EDocument;
using Microsoft.Sales.Customer;
using Microsoft.Sales.History;
using Microsoft.Sales.Setup;
using System.Reflection;

codeunit 13925 "Library - E-Doc DE"
{
    Access = Internal;
    EventSubscriberInstance = Manual;

    var
        LibraryERM: Codeunit "Library - ERM";
        LibrarySales: Codeunit "Library - Sales";
        LibraryUtility: Codeunit "Library - Utility";
        CapturedPaymentMeansHeaderRecordId: RecordId;
        DefaultCoarseRoutingTxt: Label '99', Locked = true;
        SEPADirectDebitMeansCodeTok: Label '59', Locked = true;
        CapturedEDocumentServiceCode: Code[20];
        EDocumentServiceEventCount: Integer;
        ZUGFeRDSubLineScenarioEnabled: Boolean;
        ZUGFeRDForceExtendedProfile: Boolean;
        ZUGFeRDUsedPartAmount: Decimal;
        ZUGFeRDAdditionalLineNo: Integer;
        ZUGFeRDDespatchAdviceNo: Code[20];
        ZUGFeRDLineIDPrefix: Text;
        ZUGFeRDUsedPartNoteTxt: Label 'The invoice amount includes VAT on the value of the old part.', Locked = true;
        ZUGFeRDRegisterNoteTxt: Label 'Managing director: Jane Doe, Register court: Musterstadt HRB 12345', Locked = true;
        ZUGFeRDUsedPartLineTxt: Label 'Old part value', Locked = true;
        ZUGFeRDUsedPartSubjectCodeTok: Label 'ADU', Locked = true;
        ZUGFeRDRegisterSubjectCodeTok: Label 'REG', Locked = true;

    /// <summary>
    /// Resets the OnAfterFindEDocumentService capture state. Bind this codeunit with
    /// BindSubscription before the export and unbind afterwards.
    /// </summary>
    procedure ClearCapturedEDocumentService()
    begin
        Clear(CapturedEDocumentServiceCode);
        Clear(EDocumentServiceEventCount);
    end;

    /// <summary>
    /// Returns the Code of the E-Document Service carried by the last OnAfterFindEDocumentService event.
    /// </summary>
    /// <returns>The Code of the E-Document Service carried by the last captured event.</returns>
    procedure GetCapturedEDocumentServiceCode(): Code[20]
    begin
        exit(CapturedEDocumentServiceCode);
    end;

    /// <summary>
    /// Returns how many times OnAfterFindEDocumentService was raised while bound.
    /// </summary>
    /// <returns>The number of OnAfterFindEDocumentService events captured while the subscriber was bound.</returns>
    procedure GetEDocumentServiceEventCount(): Integer
    begin
        exit(EDocumentServiceEventCount);
    end;

#if not CLEAN30
#pragma warning disable AL0432
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Export XRechnung Document", 'OnAfterFindEDocumentService', '', false, false)]
    local procedure CaptureXRechnungOnAfterFindEDocumentService(var EDocumentService: Record "E-Document Service"; EDocumentFormat: Code[20])
    begin
        CapturedEDocumentServiceCode := EDocumentService.Code;
        EDocumentServiceEventCount += 1;
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Export ZUGFeRD Document", 'OnAfterFindEDocumentService', '', false, false)]
    local procedure CaptureZUGFeRDOnAfterFindEDocumentService(var EDocumentService: Record "E-Document Service")
    begin
        CapturedEDocumentServiceCode := EDocumentService.Code;
        EDocumentServiceEventCount += 1;
    end;
#pragma warning restore AL0432
#endif

    /// <summary>
    /// Returns the record ID of the header carried by the last payment means event of the XRechnung or
    /// ZUGFeRD export. Bind this codeunit with BindSubscription before the export and unbind afterwards.
    /// </summary>
    /// <returns>The record ID of the header carried by the last captured payment means event.</returns>
    procedure GetCapturedPaymentMeansHeaderRecordId(): RecordId
    begin
        exit(CapturedPaymentMeansHeaderRecordId);
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Export XRechnung Document", OnInsertPaymentMeansOnBeforeAddToRoot, '', false, false)]
    local procedure CaptureXRechnungOnInsertPaymentMeansOnBeforeAddToRoot(var PaymentMeansElement: XmlElement; HeaderRecRef: RecordRef)
    begin
        CapturedPaymentMeansHeaderRecordId := HeaderRecRef.RecordId();
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Export ZUGFeRD Document", OnInsertPaymentMethodOnBeforeAddToRoot, '', false, false)]
    local procedure CaptureZUGFeRDOnInsertPaymentMethodOnBeforeAddToRoot(var PaymentMethodElement: XmlElement; HeaderRecRef: RecordRef)
    begin
        CapturedPaymentMeansHeaderRecordId := HeaderRecRef.RecordId();
    end;

    /// <summary>
    /// Makes the ZUGFeRD event subscribers of this codeunit build a document like an automotive invoice: two header notes
    /// (ADU and REG), every line except the first one as a sub-line of the first line, an additional old-part value line
    /// that is also a sub-line of the first line, the old-part value as paid amount and a despatch advice reference.
    /// Bind this codeunit with BindSubscription before the export and unbind afterwards.
    /// </summary>
    /// <param name="UsedPartAmount">The net amount of the additional old-part value line, also used as paid amount.</param>
    /// <param name="AdditionalLineNo">The line number of the additional line.</param>
    /// <param name="DespatchAdviceNo">The despatch advice number appended to the delivery block.</param>
    procedure EnableZUGFeRDSubLineScenario(UsedPartAmount: Decimal; AdditionalLineNo: Integer; DespatchAdviceNo: Code[20])
    begin
        ZUGFeRDSubLineScenarioEnabled := true;
        ZUGFeRDUsedPartAmount := UsedPartAmount;
        ZUGFeRDAdditionalLineNo := AdditionalLineNo;
        ZUGFeRDDespatchAdviceNo := DespatchAdviceNo;
    end;

    /// <summary>
    /// Switches off all ZUGFeRD scenarios of this codeunit, so that a later bind only captures events.
    /// </summary>
    procedure ResetZUGFeRDScenario()
    begin
        Clear(ZUGFeRDSubLineScenarioEnabled);
        Clear(ZUGFeRDForceExtendedProfile);
        Clear(ZUGFeRDUsedPartAmount);
        Clear(ZUGFeRDAdditionalLineNo);
        Clear(ZUGFeRDDespatchAdviceNo);
        Clear(ZUGFeRDLineIDPrefix);
    end;

    /// <summary>
    /// Makes the profile resolution subscriber of this codeunit return the EXTENDED profile, whatever the service says.
    /// </summary>
    procedure ForceZUGFeRDExtendedProfile()
    begin
        ZUGFeRDForceExtendedProfile := true;
    end;

    /// <summary>
    /// Makes the line identification subscriber of this codeunit prefix every line ID, like an own position numbering.
    /// </summary>
    /// <param name="LineIDPrefix">The prefix for every line ID.</param>
    procedure SetZUGFeRDLineIDPrefix(LineIDPrefix: Text)
    begin
        ZUGFeRDLineIDPrefix := LineIDPrefix;
    end;

    /// <summary>
    /// Returns the content of the ADU note that the sub-line scenario adds.
    /// </summary>
    /// <returns>The content of the ADU note.</returns>
    procedure GetZUGFeRDUsedPartNote(): Text
    begin
        exit(ZUGFeRDUsedPartNoteTxt);
    end;

    /// <summary>
    /// Returns the content of the REG note that the sub-line scenario adds.
    /// </summary>
    /// <returns>The content of the REG note.</returns>
    procedure GetZUGFeRDRegisterNote(): Text
    begin
        exit(ZUGFeRDRegisterNoteTxt);
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Export ZUGFeRD Document", OnAfterResolveZUGFeRDProfile, '', false, false)]
    local procedure ForceExtendedOnAfterResolveZUGFeRDProfile(EDocumentService: Record "E-Document Service"; HeaderRecordRef: RecordRef; var ZUGFeRDProfile: Enum "ZUGFeRD Profile")
    begin
        if ZUGFeRDForceExtendedProfile then
            ZUGFeRDProfile := ZUGFeRDProfile::Extended;
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Export ZUGFeRD Document", OnCollectHeaderNotes, '', false, false)]
    local procedure AddNotesOnCollectHeaderNotes(HeaderRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile"; var TempZUGFeRDNote: Record "ZUGFeRD Note" temporary)
    begin
        if not ZUGFeRDSubLineScenarioEnabled then
            exit;
        TempZUGFeRDNote.AddNote(ZUGFeRDUsedPartSubjectCodeTok, ZUGFeRDUsedPartNoteTxt);
        TempZUGFeRDNote.AddNote(ZUGFeRDRegisterSubjectCodeTok, ZUGFeRDRegisterNoteTxt);
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Export ZUGFeRD Document", OnInsertLineOnBeforeCreateAssociatedDocumentLine, '', false, false)]
    local procedure SetHierarchyOnInsertLineOnBeforeCreateAssociatedDocumentLine(LineRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile"; var LineID: Text; var ParentLineID: Text; var LineSubtype: Enum "ZUGFeRD Line Subtype")
    var
        SalesInvoiceLine: Record "Sales Invoice Line";
        LineNo: Integer;
        FirstLineNo: Integer;
    begin
        if ZUGFeRDLineIDPrefix <> '' then
            LineID := ZUGFeRDLineIDPrefix + LineID;
        if not ZUGFeRDSubLineScenarioEnabled then
            exit;
        // "Line No." has the same field number on Sales Invoice Line and Sales Cr.Memo Line.
        LineNo := LineRecordRef.Field(SalesInvoiceLine.FieldNo("Line No.")).Value();
        FirstLineNo := GetFirstDocumentLineNo(LineRecordRef);
        if LineNo = FirstLineNo then
            exit;
        ParentLineID := ZUGFeRDLineIDPrefix + Format(FirstLineNo);
        if LineNo = ZUGFeRDAdditionalLineNo then
            LineSubtype := LineSubtype::Detail;
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Export ZUGFeRD Document", OnCollectAdditionalInvoiceLines, '', false, false)]
    local procedure AddUsedPartLineOnCollectAdditionalInvoiceLines(HeaderRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile"; var TempAdditionalSalesInvLine: Record "Sales Invoice Line" temporary)
    var
        SalesInvoiceHeader: Record "Sales Invoice Header";
        SalesInvoiceLine: Record "Sales Invoice Line";
    begin
        if not ZUGFeRDSubLineScenarioEnabled then
            exit;
        HeaderRecordRef.SetTable(SalesInvoiceHeader);
        SalesInvoiceLine.SetRange("Document No.", SalesInvoiceHeader."No.");
        SalesInvoiceLine.SetFilter(Type, '<>%1', SalesInvoiceLine.Type::" ");
        SalesInvoiceLine.FindFirst();
        TempAdditionalSalesInvLine := SalesInvoiceLine;
        TempAdditionalSalesInvLine."Line No." := ZUGFeRDAdditionalLineNo;
        TempAdditionalSalesInvLine.Description := ZUGFeRDUsedPartLineTxt;
        TempAdditionalSalesInvLine."Item Reference No." := '';
        TempAdditionalSalesInvLine.Quantity := 1;
        TempAdditionalSalesInvLine."Unit Price" := ZUGFeRDUsedPartAmount;
        TempAdditionalSalesInvLine."Line Discount %" := 0;
        TempAdditionalSalesInvLine."Line Discount Amount" := 0;
        TempAdditionalSalesInvLine."Inv. Discount Amount" := 0;
        TempAdditionalSalesInvLine.Amount := ZUGFeRDUsedPartAmount;
        TempAdditionalSalesInvLine."Amount Including VAT" := GetUsedPartAmountIncludingVAT(SalesInvoiceLine."VAT %");
        TempAdditionalSalesInvLine.Insert();
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Export ZUGFeRD Document", OnCollectAdditionalCrMemoLines, '', false, false)]
    local procedure AddUsedPartLineOnCollectAdditionalCrMemoLines(HeaderRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile"; var TempAdditionalSalesCrMemoLine: Record "Sales Cr.Memo Line" temporary)
    var
        SalesCrMemoHeader: Record "Sales Cr.Memo Header";
        SalesCrMemoLine: Record "Sales Cr.Memo Line";
    begin
        if not ZUGFeRDSubLineScenarioEnabled then
            exit;
        HeaderRecordRef.SetTable(SalesCrMemoHeader);
        SalesCrMemoLine.SetRange("Document No.", SalesCrMemoHeader."No.");
        SalesCrMemoLine.FindFirst();
        TempAdditionalSalesCrMemoLine := SalesCrMemoLine;
        TempAdditionalSalesCrMemoLine."Line No." := ZUGFeRDAdditionalLineNo;
        TempAdditionalSalesCrMemoLine.Description := ZUGFeRDUsedPartLineTxt;
        TempAdditionalSalesCrMemoLine."Item Reference No." := '';
        TempAdditionalSalesCrMemoLine.Quantity := 1;
        TempAdditionalSalesCrMemoLine."Unit Price" := ZUGFeRDUsedPartAmount;
        TempAdditionalSalesCrMemoLine."Line Discount %" := 0;
        TempAdditionalSalesCrMemoLine."Line Discount Amount" := 0;
        TempAdditionalSalesCrMemoLine."Inv. Discount Amount" := 0;
        TempAdditionalSalesCrMemoLine.Amount := ZUGFeRDUsedPartAmount;
        TempAdditionalSalesCrMemoLine."Amount Including VAT" := GetUsedPartAmountIncludingVAT(SalesCrMemoLine."VAT %");
        TempAdditionalSalesCrMemoLine.Insert();
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Export ZUGFeRD Document", OnInsertMonetarySummationOnBeforeCreateElement, '', false, false)]
    local procedure SetPrepaidAmountOnInsertMonetarySummationOnBeforeCreateElement(HeaderRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile"; var PrepaidAmount: Decimal)
    begin
        if ZUGFeRDSubLineScenarioEnabled then
            PrepaidAmount := ZUGFeRDUsedPartAmount;
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Export ZUGFeRD Document", OnAfterInsertApplicableHeaderTradeDelivery, '', false, false)]
    local procedure AddDespatchAdviceOnAfterInsertApplicableHeaderTradeDelivery(var DeliveryElement: XmlElement; HeaderRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile")
    var
        RAMNamespaceTok: Label 'urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100', Locked = true;
    begin
        if not ZUGFeRDSubLineScenarioEnabled then
            exit;
        DeliveryElement.Add(
            XmlElement.Create('DespatchAdviceReferencedDocument', RAMNamespaceTok,
                XmlElement.Create('IssuerAssignedID', RAMNamespaceTok, ZUGFeRDDespatchAdviceNo)));
    end;

    /// <summary>
    /// Returns the old-part amount including VAT at the given VAT rate, rounded like a posted line in local currency.
    /// </summary>
    /// <param name="VATPercent">The VAT rate.</param>
    /// <returns>The old-part amount including VAT.</returns>
    procedure GetUsedPartAmountIncludingVAT(VATPercent: Decimal): Decimal
    begin
        exit(Round(ZUGFeRDUsedPartAmount * (1 + VATPercent / 100)));
    end;

    local procedure GetFirstDocumentLineNo(LineRecordRef: RecordRef): Integer
    var
        SalesInvoiceLine: Record "Sales Invoice Line";
        SalesCrMemoLine: Record "Sales Cr.Memo Line";
        DocumentNo: Code[20];
    begin
        // "Document No." has the same field number on Sales Invoice Line and Sales Cr.Memo Line.
        DocumentNo := LineRecordRef.Field(SalesInvoiceLine.FieldNo("Document No.")).Value();
        if LineRecordRef.Number() = Database::"Sales Cr.Memo Line" then begin
            SalesCrMemoLine.SetRange("Document No.", DocumentNo);
            SalesCrMemoLine.FindFirst();
            exit(SalesCrMemoLine."Line No.");
        end;
        SalesInvoiceLine.SetRange("Document No.", DocumentNo);
        SalesInvoiceLine.SetFilter(Type, '<>%1', SalesInvoiceLine.Type::" ");
        SalesInvoiceLine.FindFirst();
        exit(SalesInvoiceLine."Line No.");
    end;

    /// <summary>
    /// Creates a payment method with the SEPA direct debit payment means code '59'.
    /// </summary>
    /// <returns>The Code of the new payment method.</returns>
    procedure CreateDirectDebitPaymentMethod(): Code[10]
    begin
        exit(CreatePaymentMethodWithMeansCode(SEPADirectDebitMeansCodeTok));
    end;

    /// <summary>
    /// Creates a payment method with the given UNCL4461 payment means code.
    /// </summary>
    /// <param name="PaymentMeansCode">The payment means code to set up on the payment method.</param>
    /// <returns>The Code of the new payment method.</returns>
    procedure CreatePaymentMethodWithMeansCode(PaymentMeansCode: Code[3]): Code[10]
    var
        PaymentMethod: Record "Payment Method";
    begin
        LibraryERM.CreatePaymentMethod(PaymentMethod);
        PaymentMethod.Validate("Payment Means Code", PaymentMeansCode);
        PaymentMethod.Modify(true);
        exit(PaymentMethod.Code);
    end;

    /// <summary>
    /// Creates a customer with a bank account with an IBAN and a SEPA direct debit mandate for it.
    /// </summary>
    /// <param name="SEPADirectDebitMandate">Returns the new mandate.</param>
    /// <param name="CustomerBankAccount">Returns the new customer bank account the mandate refers to.</param>
    /// <param name="PaymentMethodCode">The payment method to set on the customer.</param>
    /// <returns>The No. of the new customer.</returns>
    procedure CreateCustomerWithDirectDebitMandate(var SEPADirectDebitMandate: Record "SEPA Direct Debit Mandate"; var CustomerBankAccount: Record "Customer Bank Account"; PaymentMethodCode: Code[10]): Code[20]
    var
        Customer: Record Customer;
    begin
        LibrarySales.CreateCustomer(Customer);
        exit(AddDirectDebitMandateToCustomer(SEPADirectDebitMandate, CustomerBankAccount, Customer."No.", PaymentMethodCode));
    end;

    /// <summary>
    /// Gives an existing customer a bank account with an IBAN and a SEPA direct debit mandate for it.
    /// Use this when the test needs a customer built by its own suite, for example one with the
    /// address and VAT data a format requires.
    /// </summary>
    /// <param name="SEPADirectDebitMandate">Returns the new mandate.</param>
    /// <param name="CustomerBankAccount">Returns the new customer bank account the mandate refers to.</param>
    /// <param name="CustomerNo">The customer that gets the bank account and the mandate.</param>
    /// <param name="PaymentMethodCode">The payment method to set on the customer.</param>
    /// <returns>The No. of the customer.</returns>
    procedure AddDirectDebitMandateToCustomer(var SEPADirectDebitMandate: Record "SEPA Direct Debit Mandate"; var CustomerBankAccount: Record "Customer Bank Account"; CustomerNo: Code[20]; PaymentMethodCode: Code[10]): Code[20]
    var
        Customer: Record Customer;
        SalesReceivablesSetup: Record "Sales & Receivables Setup";
    begin
        LibraryUtility.UpdateSetupNoSeriesCode(Database::"Sales & Receivables Setup", SalesReceivablesSetup.FieldNo("Direct Debit Mandate Nos."));
        Customer.Get(CustomerNo);
        Customer.Validate("Payment Method Code", PaymentMethodCode);
        Customer.Modify(true);
        LibrarySales.CreateCustomerBankAccount(CustomerBankAccount, Customer."No.");
        CustomerBankAccount.IBAN := LibraryUtility.GenerateMOD97CompliantCode();
        CustomerBankAccount.Modify(true);
        LibrarySales.CreateCustomerMandate(SEPADirectDebitMandate, Customer."No.", CustomerBankAccount.Code, WorkDate(), CalcDate('<1Y>', WorkDate()));
        Customer.Validate("Preferred Bank Account Code", CustomerBankAccount.Code);
        Customer.Modify(true);
        exit(Customer."No.");
    end;

    procedure CreateValidRoutingNo(): Text[50]
    var
        FineRouting: Text[20];
        CheckDigit: Text[2];
    begin
        FineRouting := GenerateAlphanumFineRouting();
        CheckDigit := ComputeCheckDigit(DefaultCoarseRoutingTxt, FineRouting);
        exit(CopyStr(DefaultCoarseRoutingTxt + '-' + FineRouting + '-' + CheckDigit, 1, 50));
    end;

    local procedure GenerateAlphanumFineRouting(): Text[20]
    begin
        // CreateGuid() produces hex chars (0-9, A-F) which are valid alphanumeric fine routing characters
        exit(CopyStr(DelChr(Format(CreateGuid()), '=', '{-}'), 1, 20));
    end;

    local procedure ComputeCheckDigit(CoarseRouting: Text; FineRouting: Text): Text[2]
    var
        NumericString: Text;
        Remainder: Integer;
        CheckDigitValue: Integer;
    begin
        NumericString := ConvertToNumericString(CoarseRouting + FineRouting) + '00';
        Remainder := ComputeMod97(NumericString);
        CheckDigitValue := 98 - Remainder;
        if CheckDigitValue < 10 then
            exit(CopyStr('0' + Format(CheckDigitValue), 1, 2));
        exit(CopyStr(Format(CheckDigitValue), 1, 2));
    end;

    local procedure ConvertToNumericString(Input: Text): Text
    var
        TypeHelper: Codeunit "Type Helper";
        UpperInput: Text;
        Result: Text;
        Ch: Char;
        i: Integer;
    begin
        UpperInput := UpperCase(Input);
        for i := 1 to StrLen(UpperInput) do begin
            Ch := UpperInput[i];
            if TypeHelper.IsLatinLetter(Ch) then
                Result += Format(Ch - 55)
            else
                Result += Format(Ch - 48);
        end;
        exit(Result);
    end;

    local procedure ComputeMod97(NumericString: Text): Integer
    var
        Remainder: Integer;
        DigitValue: Integer;
        i: Integer;
    begin
        Remainder := 0;
        for i := 1 to StrLen(NumericString) do begin
            Evaluate(DigitValue, CopyStr(NumericString, i, 1));
            Remainder := (Remainder * 10 + DigitValue) mod 97;
        end;
        exit(Remainder);
    end;
}
