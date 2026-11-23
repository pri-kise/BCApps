// ------------------------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
// ------------------------------------------------------------------------------------------------
namespace Microsoft.eServices.EDocument.Formats;

using Microsoft.Bank.BankAccount;
using Microsoft.Bank.DirectDebit;
using Microsoft.CRM.Team;
using Microsoft.eServices.EDocument;
using Microsoft.Finance.Currency;
using Microsoft.Finance.GeneralLedger.Journal;
using Microsoft.Finance.GeneralLedger.Setup;
using Microsoft.Finance.VAT.Ledger;
using Microsoft.Finance.VAT.Setup;
using Microsoft.Foundation.Address;
using Microsoft.Foundation.Company;
using Microsoft.Foundation.PaymentTerms;
using Microsoft.Foundation.Reporting;
using Microsoft.Foundation.UOM;
using Microsoft.Inventory.Item;
using Microsoft.Inventory.Item.Attribute;
using Microsoft.Inventory.Item.Catalog;
using Microsoft.Inventory.Location;
using Microsoft.Peppol;
using Microsoft.Projects.Project.Job;
using Microsoft.Sales.Customer;
using Microsoft.Sales.History;
using Microsoft.Sales.Receivables;
using Microsoft.Service.History;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Telemetry;
using System.Utilities;

codeunit 13917 "Export ZUGFeRD Document"
{
    TableNo = "Record Export Buffer";
    InherentEntitlements = X;
    InherentPermissions = X;

    var
        CompanyInformation: Record "Company Information";
        GeneralLedgerSetup: Record "General Ledger Setup";
        EDocumentService: Record "E-Document Service";
        TempAdditionalSalesInvLine: Record "Sales Invoice Line" temporary;
        TempAdditionalSalesCrMemoLine: Record "Sales Cr.Memo Line" temporary;
        FeatureTelemetry: Codeunit "Feature Telemetry";
        PeppolVATHelper: Codeunit "PEPPOL VAT Helper";
        EDocumentDEHelper: Codeunit "E-Document DE Helper";
        EDocItemChargeMapping: Codeunit "E-Doc. Item Charge Mapping";
        DEPaymentMeansHelper: Codeunit "DE Payment Means Helper";
        SourceDocumentRecordRef: RecordRef;
        ItemChargeStructures: Dictionary of [Integer, Integer];
        LineLevelItemChargeAmounts: Dictionary of [Integer, Decimal];
        LineLevelItemChargeLineNos: Dictionary of [Integer, List of [Integer]];
        FeatureNameTok: Label 'E-document ZUGFeRD Format', Locked = true;
        StartEventNameTok: Label 'E-document ZUGFeRD export started', Locked = true;
        EndEventNameTok: Label 'E-document ZUGFeRD export completed', Locked = true;
        GLNSchemeIDTok: Label '0088', Locked = true;
        XmlNamespaceRSM: Text;
        XmlNamespaceRAM: Text;
        XmlNamespaceUDT: Text;
        XmlNamespaceQDT: Text;
        ItemGTINCache: Dictionary of [Code[20], Code[14]];
#if not CLEAN30
        EDocumentServiceProvided: Boolean;
#endif
        DocumentLanguageCode: Code[10];
        CurrentZUGFeRDProfile: Enum "ZUGFeRD Profile";
        VATBreakdownKeys: List of [Text];
        VATBreakdownBasisAmounts: Dictionary of [Text, Decimal];
        VATBreakdownVATAmounts: Dictionary of [Text, Decimal];
        VATBreakdownCategories: Dictionary of [Text, Text];
        VATBreakdownRates: Dictionary of [Text, Decimal];
        VATBreakdownBusPostingGroups: Dictionary of [Text, Code[20]];
        VATBreakdownProdPostingGroups: Dictionary of [Text, Code[20]];
        XRechnungGuidelineIDTok: Label 'urn:cen.eu:en16931:2017#compliant#urn:xeinkauf.de:kosit:xrechnung_3.0', Locked = true;
        ExtendedGuidelineIDTok: Label 'urn:cen.eu:en16931:2017#conformant#urn:factur-x.eu:1p0:extended', Locked = true;
        XRechnungAttachmentNameTok: Label 'xrechnung.xml', Locked = true;
        FacturXAttachmentNameTok: Label 'factur-x.xml', Locked = true;
        LineSubtypeDetailTok: Label 'DETAIL', Locked = true;
        LineSubtypeGroupTok: Label 'GROUP', Locked = true;
        LineSubtypeInformationTok: Label 'INFORMATION', Locked = true;
        AdditionalLineNoExistsErr: Label 'The additional line %1 for the ZUGFeRD export uses a line number that the document already has. Use a line number that no line of the document uses.', Comment = '%1 = Line No.';

    trigger OnRun()
    var
        ZUGFeRDReportIntegration: Codeunit "ZUGFeRD Report Integration";
    begin
        BindSubscription(ZUGFeRDReportIntegration);

        ExportSalesDocument(Rec);

        UnbindSubscription(ZUGFeRDReportIntegration);
        ResetProvidedService();
    end;


    /// <summary>
    /// Use this procedure to check if the current report print is for the ZUGFeRD export.
    /// </summary>
    /// <returns>true when the XML should be embedded</returns>
    procedure IsZUGFeRDPrintProcess() Result: Boolean
    begin
        Result := false;
        OnIsZUGFeRDPrintProcess(Result);
    end;

    [InternalEvent(false)]
    local procedure OnIsZUGFeRDPrintProcess(var Result: Boolean)
    begin
    end;

    /// <summary>
    /// Creates a ZUGFeRD XML document from the sales invoice and adds it as an attachment to the rendering payload.
    /// </summary>
    /// <param name="SalesInvoiceHeader">The sales invoice header record to export.</param>
    /// <param name="RenderingPayload">The JSON object to add the XML attachment to.</param>
    procedure CreateAndAddXMLAttachmentToRenderingPayload(var SalesInvoiceHeader: Record "Sales Invoice Header"; var RenderingPayload: JsonObject)
    var
        TempBlob: Codeunit "Temp Blob";
        XmlOutStream: OutStream;
    begin
        TempBlob.CreateOutStream(XmlOutStream, TextEncoding::UTF8);
        CreateXML(SalesInvoiceHeader, XmlOutStream);

        AddXMLAttachmentToRenderingPayload(TempBlob, RenderingPayload);
    end;

    /// <summary>
    /// Creates a ZUGFeRD XML document from the sales credit memo and adds it as an attachment to the rendering payload.
    /// </summary>
    /// <param name="SalesCrMemoHeader">The sales credit memo header record to export.</param>
    /// <param name="RenderingPayload">The JSON object to add the XML attachment to.</param>
    procedure CreateAndAddXMLAttachmentToRenderingPayload(var SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var RenderingPayload: JsonObject)
    var
        TempBlob: Codeunit "Temp Blob";
        XmlOutStream: OutStream;
    begin
        TempBlob.CreateOutStream(XmlOutStream, TextEncoding::UTF8);
        CreateXML(SalesCrMemoHeader, XmlOutStream);

        AddXMLAttachmentToRenderingPayload(TempBlob, RenderingPayload);
    end;

    /// <summary>
    /// Creates a ZUGFeRD XML document from the service invoice and adds it as an attachment to the rendering payload.
    /// </summary>
    /// <param name="ServiceInvoiceHeader">The service invoice header record to export.</param>
    /// <param name="RenderingPayload">The JSON object to add the XML attachment to.</param>
    procedure CreateAndAddXMLAttachmentToRenderingPayload(var ServiceInvoiceHeader: Record "Service Invoice Header"; var RenderingPayload: JsonObject)
    var
        TempBlob: Codeunit "Temp Blob";
        XmlOutStream: OutStream;
    begin
        TempBlob.CreateOutStream(XmlOutStream, TextEncoding::UTF8);
        CreateXML(ServiceInvoiceHeader, XmlOutStream);

        AddXMLAttachmentToRenderingPayload(TempBlob, RenderingPayload);
    end;

    /// <summary>
    /// Creates a ZUGFeRD XML document from the service credit memo and adds it as an attachment to the rendering payload.
    /// </summary>
    /// <param name="ServiceCrMemoHeader">The service credit memo header record to export.</param>
    /// <param name="RenderingPayload">The JSON object to add the XML attachment to.</param>
    procedure CreateAndAddXMLAttachmentToRenderingPayload(var ServiceCrMemoHeader: Record "Service Cr.Memo Header"; var RenderingPayload: JsonObject)
    var
        TempBlob: Codeunit "Temp Blob";
        XmlOutStream: OutStream;
    begin
        TempBlob.CreateOutStream(XmlOutStream, TextEncoding::UTF8);
        CreateXML(ServiceCrMemoHeader, XmlOutStream);

        AddXMLAttachmentToRenderingPayload(TempBlob, RenderingPayload);
    end;

    local procedure AddXMLAttachmentToRenderingPayload(var XmlAttachmentTempBlob: Codeunit "Temp Blob"; var RenderingPayload: JsonObject)
    var
        PDFDocument: Codeunit "PDF Document";
        DataType: Enum "PDF Attach. Data Relationship";
        XmlInStream: InStream;
        Name: Text;
        MimeType: Text;
        Description: Text;
        DescriptionLbl: Label 'This is the e-invoicing xml document';
    begin
        PDFDocument.Initialize();
        Name := GetXMLAttachmentName();
        DataType := Enum::"PDF Attach. Data Relationship"::Alternative;
        MimeType := 'text/xml';
        Description := DescriptionLbl;

        XmlAttachmentTempBlob.CreateInStream(XmlInStream, TextEncoding::UTF8);
        PDFDocument.AddAttachment(Name, DataType, MimeType, XmlInStream, Description, true);

        RenderingPayload := PDFDocument.ToJson(RenderingPayload);
    end;

    procedure ExportSalesDocument(var RecordExportBuffer: Record "Record Export Buffer")
    var
        SalesInvoiceHeader: Record "Sales Invoice Header";
        SalesCrMemoHeader: Record "Sales Cr.Memo Header";
        ServiceInvoiceHeader: Record "Service Invoice Header";
        ServiceCrMemoHeader: Record "Service Cr.Memo Header";
        TempBlob: Codeunit "Temp Blob";
        FileOutStream: OutStream;
        FileInStream: InStream;
    begin
        FeatureTelemetry.LogUsage('0000EXD', FeatureNameTok, StartEventNameTok);

        case RecordExportBuffer.RecordID.TableNo of
            Database::"Sales Invoice Header":
                begin
                    SalesInvoiceHeader.Get(RecordExportBuffer.RecordID);
                    if not GenerateSalesInvoicePDFAttachment(SalesInvoiceHeader, TempBlob) then
                        exit;
                end;
            Database::"Sales Cr.Memo Header":
                begin
                    SalesCrMemoHeader.Get(RecordExportBuffer.RecordID);
                    if not GenerateSalesCrMemoPDFAttachment(SalesCrMemoHeader, TempBlob) then
                        exit;
                end;
            Database::"Service Invoice Header":
                begin
                    ServiceInvoiceHeader.Get(RecordExportBuffer.RecordID);
                    if not GenerateServiceInvoicePDFAttachment(ServiceInvoiceHeader, TempBlob) then
                        exit;
                end;
            Database::"Service Cr.Memo Header":
                begin
                    ServiceCrMemoHeader.Get(RecordExportBuffer.RecordID);
                    if not GenerateServiceCrMemoPDFAttachment(ServiceCrMemoHeader, TempBlob) then
                        exit;
                end;
        end;
        TempBlob.CreateInStream(FileInStream);
        RecordExportBuffer."File Content".CreateOutStream(FileOutStream, TextEncoding::UTF8);
        CopyStream(FileOutStream, FileInStream);
        RecordExportBuffer.Modify();
        FeatureTelemetry.LogUsage('0000EXE', FeatureNameTok, EndEventNameTok);
    end;

    procedure GenerateSalesInvoicePDFAttachment(SalesInvoiceHeader: Record "Sales Invoice Header"; var TempBlob: Codeunit "Temp Blob"): Boolean
    var
        ReportSelections: Record "Report Selections";
        IsHandled: Boolean;
    begin
        IsHandled := false;
        OnBeforeGenerateSalesInvoicePDFAttachment(SalesInvoiceHeader, TempBlob, IsHandled);
        if IsHandled then
            exit(TempBlob.HasValue());
        SalesInvoiceHeader.SetRange("No.", SalesInvoiceHeader."No.");
        ReportSelections.GetPdfReportForCust(
            TempBlob, "Report Selection Usage"::"S.Invoice",
            SalesInvoiceHeader, SalesInvoiceHeader."Bill-to Customer No.");
        exit(TempBlob.HasValue());
    end;

    procedure GenerateSalesCrMemoPDFAttachment(SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var TempBlob: Codeunit "Temp Blob"): Boolean
    var
        ReportSelections: Record "Report Selections";
        IsHandled: Boolean;
    begin
        IsHandled := false;
        OnBeforeGenerateSalesCrMemoPDFAttachment(SalesCrMemoHeader, TempBlob, IsHandled);
        if IsHandled then
            exit(TempBlob.HasValue());
        SalesCrMemoHeader.SetRange("No.", SalesCrMemoHeader."No.");
        ReportSelections.GetPdfReportForCust(
            TempBlob, "Report Selection Usage"::"S.Cr.Memo",
            SalesCrMemoHeader, SalesCrMemoHeader."Bill-to Customer No.");
        exit(TempBlob.HasValue());
    end;

    procedure GenerateServiceInvoicePDFAttachment(ServiceInvoiceHeader: Record "Service Invoice Header"; var TempBlob: Codeunit "Temp Blob"): Boolean
    var
        ReportSelections: Record "Report Selections";
        IsHandled: Boolean;
    begin
        IsHandled := false;
        OnBeforeGenerateServiceInvoicePDFAttachment(ServiceInvoiceHeader, TempBlob, IsHandled);
        if IsHandled then
            exit(TempBlob.HasValue());
        ServiceInvoiceHeader.SetRange("No.", ServiceInvoiceHeader."No.");
        ReportSelections.GetPdfReportForCust(
            TempBlob, "Report Selection Usage"::"SM.Invoice",
            ServiceInvoiceHeader, ServiceInvoiceHeader."Bill-to Customer No.");
        exit(TempBlob.HasValue());
    end;

    procedure GenerateServiceCrMemoPDFAttachment(ServiceCrMemoHeader: Record "Service Cr.Memo Header"; var TempBlob: Codeunit "Temp Blob"): Boolean
    var
        ReportSelections: Record "Report Selections";
        IsHandled: Boolean;
    begin
        IsHandled := false;
        OnBeforeGenerateServiceCrMemoPDFAttachment(ServiceCrMemoHeader, TempBlob, IsHandled);
        if IsHandled then
            exit(TempBlob.HasValue());
        ServiceCrMemoHeader.SetRange("No.", ServiceCrMemoHeader."No.");
        ReportSelections.GetPdfReportForCust(
            TempBlob, "Report Selection Usage"::"SM.Credit Memo",
            ServiceCrMemoHeader, ServiceCrMemoHeader."Bill-to Customer No.");
        exit(TempBlob.HasValue());
    end;

    procedure CreateXML(SalesInvoiceHeader: Record "Sales Invoice Header"; var FileOutstream: Outstream)
    var
        SalesInvLine: Record "Sales Invoice Line";
        Currency: Record Currency;
        CurrencyCode: Code[10];
        RootXMLNode: XmlElement;
        XMLDoc: XmlDocument;
        XMLDocText: Text;
        LineAmounts: Dictionary of [Text, Decimal];
        LineVATAmount: Dictionary of [Decimal, Decimal];
        LineAmount: Dictionary of [Decimal, Decimal];
        LineDiscAmount: Dictionary of [Decimal, Decimal];
    begin
        Clear(ItemGTINCache);
        GetSetups();
#if not CLEAN30
#pragma warning disable AL0432
        FindEDocumentService();
#pragma warning restore AL0432
#endif
        ClearAdditionalLines();
        if not DocumentLinesExist(SalesInvoiceHeader, SalesInvLine) then
            exit;

        DocumentLanguageCode := SalesInvoiceHeader."Language Code";
        SourceDocumentRecordRef.Close();
        SourceDocumentRecordRef.GetTable(SalesInvoiceHeader);
        CurrentZUGFeRDProfile := ResolveZUGFeRDProfile(EDocumentService, SourceDocumentRecordRef);
        CollectAdditionalLines(SalesInvLine);
        XmlDocument.ReadFrom(GetInvoiceXMLHeader(), XMLDoc);
        XmlDoc.GetRoot(RootXMLNode);

        InitializeNamespaces();
        ClassifyItemCharges(SalesInvoiceHeader, SalesInvLine);
        CurrencyCode := GetCurrencyCode(SalesInvoiceHeader."Currency Code", Currency);
        CalculateLineAmounts(SalesInvoiceHeader, SalesInvLine, Currency, LineAmounts);
        InsertVATAmounts(SalesInvLine, LineVATAmount, LineAmount, LineDiscAmount, SalesInvoiceHeader."Prices Including VAT", Currency);
        InsertHeaderData(RootXMLNode, SalesInvoiceHeader);
        InsertSupplyChainTradeTransaction(RootXMLNode, SalesInvoiceHeader, SalesInvLine, CurrencyCode, Currency, LineAmount, LineVATAmount, LineAmounts, LineDiscAmount);
        OnCreateXMLOnBeforeSalesInvXmlDocumentWriteToFile(XMLDoc, SalesInvoiceHeader);
        XMLDoc.WriteTo(XMLDocText);
        FileOutstream.WriteText(XMLDocText);
        Clear(XMLDoc);
        SourceDocumentRecordRef.Close();
        ClearAdditionalLines();
    end;

    procedure CreateXML(SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var FileOutstream: Outstream)
    var
        SalesCrMemoLine: Record "Sales Cr.Memo Line";
        Currency: Record Currency;
        CurrencyCode: Code[10];
        RootXMLNode: XmlElement;
        XMLDoc: XmlDocument;
        XMLDocText: Text;
        LineAmounts: Dictionary of [Text, Decimal];
        LineVATAmount: Dictionary of [Decimal, Decimal];
        LineAmount: Dictionary of [Decimal, Decimal];
        LineDiscAmount: Dictionary of [Decimal, Decimal];
    begin
        Clear(ItemGTINCache);
        GetSetups();
#if not CLEAN30
#pragma warning disable AL0432
        FindEDocumentService();
#pragma warning restore AL0432
#endif
        ClearAdditionalLines();
        if not DocumentLinesExist(SalesCrMemoHeader, SalesCrMemoLine) then
            exit;

        DocumentLanguageCode := SalesCrMemoHeader."Language Code";
        SourceDocumentRecordRef.Close();
        SourceDocumentRecordRef.GetTable(SalesCrMemoHeader);
        CurrentZUGFeRDProfile := ResolveZUGFeRDProfile(EDocumentService, SourceDocumentRecordRef);
        CollectAdditionalLines(SalesCrMemoLine);
        XmlDocument.ReadFrom(GetInvoiceXMLHeader(), XMLDoc);
        XmlDoc.GetRoot(RootXMLNode);

        InitializeNamespaces();
        ClassifyItemCharges(SalesCrMemoHeader, SalesCrMemoLine);
        CurrencyCode := GetCurrencyCode(SalesCrMemoHeader."Currency Code", Currency);
        CalculateLineAmounts(SalesCrMemoHeader, SalesCrMemoLine, Currency, LineAmounts);
        InsertVATAmounts(SalesCrMemoLine, LineVATAmount, LineAmount, LineDiscAmount, SalesCrMemoHeader."Prices Including VAT", Currency);
        InsertHeaderData(RootXMLNode, SalesCrMemoHeader);
        InsertSupplyChainTradeTransaction(RootXMLNode, SalesCrMemoHeader, SalesCrMemoLine, CurrencyCode, Currency, LineAmount, LineVATAmount, LineAmounts, LineDiscAmount);
        OnCreateXMLOnBeforeSalesCrMemoXmlDocumentWriteToFile(XMLDoc, SalesCrMemoHeader);
        XMLDoc.WriteTo(XMLDocText);
        FileOutstream.WriteText(XMLDocText);
        Clear(XMLDoc);
        SourceDocumentRecordRef.Close();
        ClearAdditionalLines();
    end;

    procedure CreateXML(ServiceInvoiceHeader: Record "Service Invoice Header"; var FileOutstream: Outstream)
    var
        SalesInvoiceHeader: Record "Sales Invoice Header";
        ServiceInvoiceLine: Record "Service Invoice Line";
        TempSalesInvLine: Record "Sales Invoice Line" temporary;
        Currency: Record Currency;
        CurrencyCode: Code[10];
        RootXMLNode: XmlElement;
        XMLDoc: XmlDocument;
        XMLDocText: Text;
        LineAmounts: Dictionary of [Text, Decimal];
        LineVATAmount: Dictionary of [Decimal, Decimal];
        LineAmount: Dictionary of [Decimal, Decimal];
        LineDiscAmount: Dictionary of [Decimal, Decimal];
    begin
        Clear(ItemGTINCache);
        GetSetups();
#if not CLEAN30
#pragma warning disable AL0432
        FindEDocumentService();
#pragma warning restore AL0432
#endif
        ClearItemChargeClassification();
        ClearAdditionalLines();
        TransferToSalesInvoiceHeader(ServiceInvoiceHeader, SalesInvoiceHeader);
        SalesInvoiceHeader."Company Bank Account Code" := ServiceInvoiceHeader."Company Bank Account Code";
        ServiceInvoiceLine.SetRange("Document No.", ServiceInvoiceHeader."No.");
        if ServiceInvoiceLine.FindSet() then
            repeat
                TransferToSalesInvoiceLine(ServiceInvoiceLine, TempSalesInvLine);
                TempSalesInvLine.Insert();
            until ServiceInvoiceLine.Next() = 0;
        if not DocumentLinesExist(SalesInvoiceHeader, TempSalesInvLine) then
            exit;

        DocumentLanguageCode := SalesInvoiceHeader."Language Code";
        SourceDocumentRecordRef.Close();
        SourceDocumentRecordRef.GetTable(ServiceInvoiceHeader);
        CurrentZUGFeRDProfile := ResolveZUGFeRDProfile(EDocumentService, SourceDocumentRecordRef);
        CollectAdditionalLines(TempSalesInvLine);
        XmlDocument.ReadFrom(GetInvoiceXMLHeader(), XMLDoc);
        XmlDoc.GetRoot(RootXMLNode);

        InitializeNamespaces();
        CurrencyCode := GetCurrencyCode(SalesInvoiceHeader."Currency Code", Currency);
        CalculateLineAmounts(SalesInvoiceHeader, TempSalesInvLine, Currency, LineAmounts);
        InsertVATAmounts(TempSalesInvLine, LineVATAmount, LineAmount, LineDiscAmount, SalesInvoiceHeader."Prices Including VAT", Currency);
        InsertHeaderData(RootXMLNode, SalesInvoiceHeader);
        InsertSupplyChainTradeTransaction(RootXMLNode, SalesInvoiceHeader, TempSalesInvLine, CurrencyCode, Currency, LineAmount, LineVATAmount, LineAmounts, LineDiscAmount);
        OnCreateXMLOnBeforeSalesInvXmlDocumentWriteToFile(XMLDoc, SalesInvoiceHeader);
        XMLDoc.WriteTo(XMLDocText);
        FileOutstream.WriteText(XMLDocText);
        Clear(XMLDoc);
        SourceDocumentRecordRef.Close();
        ClearAdditionalLines();
    end;

    procedure CreateXML(ServiceCrMemoHeader: Record "Service Cr.Memo Header"; var FileOutstream: Outstream)
    var
        SalesCrMemoHeader: Record "Sales Cr.Memo Header";
        TempSalesCrMemoLine: Record "Sales Cr.Memo Line" temporary;
        ServiceCrMemoLine: Record "Service Cr.Memo Line";
        Currency: Record Currency;
        CurrencyCode: Code[10];
        RootXMLNode: XmlElement;
        XMLDoc: XmlDocument;
        XMLDocText: Text;
        LineAmounts: Dictionary of [Text, Decimal];
        LineVATAmount: Dictionary of [Decimal, Decimal];
        LineAmount: Dictionary of [Decimal, Decimal];
        LineDiscAmount: Dictionary of [Decimal, Decimal];
    begin
        Clear(ItemGTINCache);
        GetSetups();
#if not CLEAN30
#pragma warning disable AL0432
        FindEDocumentService();
#pragma warning restore AL0432
#endif
        ClearItemChargeClassification();
        ClearAdditionalLines();
        TransferToSalesCrMemoHeader(ServiceCrMemoHeader, SalesCrMemoHeader);
        SalesCrMemoHeader."Company Bank Account Code" := ServiceCrMemoHeader."Company Bank Account Code";
        ServiceCrMemoLine.SetRange("Document No.", ServiceCrMemoHeader."No.");
        if ServiceCrMemoLine.FindSet() then
            repeat
                TransferToSalesCrMemoLine(ServiceCrMemoLine, TempSalesCrMemoLine);
                TempSalesCrMemoLine.Insert();
            until ServiceCrMemoLine.Next() = 0;
        if not DocumentLinesExist(SalesCrMemoHeader, TempSalesCrMemoLine) then
            exit;

        DocumentLanguageCode := SalesCrMemoHeader."Language Code";
        SourceDocumentRecordRef.Close();
        SourceDocumentRecordRef.GetTable(ServiceCrMemoHeader);
        CurrentZUGFeRDProfile := ResolveZUGFeRDProfile(EDocumentService, SourceDocumentRecordRef);
        CollectAdditionalLines(TempSalesCrMemoLine);
        XmlDocument.ReadFrom(GetInvoiceXMLHeader(), XMLDoc);
        XmlDoc.GetRoot(RootXMLNode);

        InitializeNamespaces();
        CurrencyCode := GetCurrencyCode(SalesCrMemoHeader."Currency Code", Currency);
        CalculateLineAmounts(SalesCrMemoHeader, TempSalesCrMemoLine, Currency, LineAmounts);
        InsertVATAmounts(TempSalesCrMemoLine, LineVATAmount, LineAmount, LineDiscAmount, SalesCrMemoHeader."Prices Including VAT", Currency);
        InsertHeaderData(RootXMLNode, SalesCrMemoHeader);
        InsertSupplyChainTradeTransaction(RootXMLNode, SalesCrMemoHeader, TempSalesCrMemoLine, CurrencyCode, Currency, LineAmount, LineVATAmount, LineAmounts, LineDiscAmount);
        OnCreateXMLOnBeforeSalesCrMemoXmlDocumentWriteToFile(XMLDoc, SalesCrMemoHeader);
        XMLDoc.WriteTo(XMLDocText);
        FileOutstream.WriteText(XMLDocText);
        Clear(XMLDoc);
        SourceDocumentRecordRef.Close();
        ClearAdditionalLines();
    end;

    local procedure DocumentLinesExist(SalesInvoiceHeader: Record "Sales Invoice Header"; var SalesInvLine: Record "Sales Invoice Line"): Boolean
    begin
        SalesInvLine.SetRange("Document No.", SalesInvoiceHeader."No.");
        SalesInvLine.SetFilter(Type, '<>%1', SalesInvLine.Type::" ");
        SalesInvLine.SetFilter("No.", '<>%1', '');
        SalesInvLine.SetFilter(Quantity, '<>0');
        OnDocumentLinesExistOnAfterFilterSalesInvLine(SalesInvoiceHeader, SalesInvLine);
        if SalesInvLine.FindSet() then
            exit(true);
        exit(false);
    end;

    local procedure DocumentLinesExist(SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var SalesCrMemoLine: Record "Sales Cr.Memo Line"): Boolean
    begin
        SalesCrMemoLine.SetRange("Document No.", SalesCrMemoHeader."No.");
        SalesCrMemoLine.SetFilter(Type, '<>%1', SalesCrMemoLine.Type::" ");
        SalesCrMemoLine.SetFilter("No.", '<>%1', '');
        SalesCrMemoLine.SetFilter(Quantity, '<>0');
        OnDocumentLinesExistOnAfterFilterSalesCrMemoLine(SalesCrMemoHeader, SalesCrMemoLine);
        if SalesCrMemoLine.FindSet() then
            exit(true);
        exit(false);
    end;

    local procedure InitializeNamespaces()
    begin
        XmlNamespaceRSM := 'urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100';
        XmlNamespaceRAM := 'urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100';
        XmlNamespaceUDT := 'urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100';
        XmlNamespaceQDT := 'urn:un:unece:uncefact:data:standard:QualifiedDataType:100';
    end;

    local procedure GetInvoiceXMLHeader(): Text
    begin
        exit('<?xml version="1.0" encoding="UTF-8"?>' +
        '<rsm:CrossIndustryInvoice xmlns:rsm="urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100" ' +
        'xmlns:qdt="urn:un:unece:uncefact:data:standard:QualifiedDataType:100" ' +
        'xmlns:ram="urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100" ' +
        'xmlns:xs="http://www.w3.org/2001/XMLSchema" ' +
        'xmlns:udt="urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100" />');
    end;

    local procedure InsertHeaderData(var RootXMLNode: XmlElement; SalesInvoiceHeader: Record "Sales Invoice Header")
    var
        ExchangedDocumentContextElement, GuidelineSpecifiedDocumentContextParameterElement, BusinessProcessContextElement, ExchangedDocumentElement, IssueDateTimeElement : XmlElement;
    begin
        ExchangedDocumentContextElement := XmlElement.Create('ExchangedDocumentContext', XmlNamespaceRSM);
        BusinessProcessContextElement := XmlElement.Create('BusinessProcessSpecifiedDocumentContextParameter', XmlNamespaceRAM);
        BusinessProcessContextElement.Add(XmlElement.Create('ID', XmlNamespaceRAM, 'urn:fdc:peppol.eu:2017:poacc:billing:01:1.0'));
        ExchangedDocumentContextElement.Add(BusinessProcessContextElement);

        GuidelineSpecifiedDocumentContextParameterElement := XmlElement.Create('GuidelineSpecifiedDocumentContextParameter', XmlNamespaceRAM);
        GuidelineSpecifiedDocumentContextParameterElement.Add(XmlElement.Create('ID', XmlNamespaceRAM, GetGuidelineID()));
        ExchangedDocumentContextElement.Add(GuidelineSpecifiedDocumentContextParameterElement);
        RootXMLNode.Add(ExchangedDocumentContextElement);
        ExchangedDocumentElement := XmlElement.Create('ExchangedDocument', XmlNamespaceRSM);
        ExchangedDocumentElement.Add(XmlElement.Create('ID', XmlNamespaceRAM, SalesInvoiceHeader."No."));
        ExchangedDocumentElement.Add(XmlElement.Create('TypeCode', XmlNamespaceRAM, '380')); // Invoice type code

        IssueDateTimeElement := XmlElement.Create('IssueDateTime', XmlNamespaceRAM);
        IssueDateTimeElement.Add(XmlElement.Create('DateTimeString', XmlNamespaceUDT, XmlAttribute.Create('format', '102'), FormatDate(SalesInvoiceHeader."Document Date")));
        ExchangedDocumentElement.Add(IssueDateTimeElement);
        InsertLanguageID(ExchangedDocumentElement, SalesInvoiceHeader."Language Code");
        InsertHeaderNotes(ExchangedDocumentElement, SalesInvoiceHeader);
        RootXMLNode.Add(ExchangedDocumentElement);
        OnAfterInsertSalesInvHeaderData(RootXMLNode, SalesInvoiceHeader);
    end;

    local procedure InsertHeaderData(var RootXMLNode: XmlElement; SalesCrMemoHeader: Record "Sales Cr.Memo Header")
    var
        ExchangedDocumentContextElement, GuidelineSpecifiedDocumentContextParameterElement, BusinessProcessContextElement, ExchangedDocumentElement, IssueDateTimeElement : XmlElement;
    begin
        ExchangedDocumentContextElement := XmlElement.Create('ExchangedDocumentContext', XmlNamespaceRSM);
        BusinessProcessContextElement := XmlElement.Create('BusinessProcessSpecifiedDocumentContextParameter', XmlNamespaceRAM);
        BusinessProcessContextElement.Add(XmlElement.Create('ID', XmlNamespaceRAM, 'urn:fdc:peppol.eu:2017:poacc:billing:01:1.0'));
        ExchangedDocumentContextElement.Add(BusinessProcessContextElement);

        GuidelineSpecifiedDocumentContextParameterElement := XmlElement.Create('GuidelineSpecifiedDocumentContextParameter', XmlNamespaceRAM);
        GuidelineSpecifiedDocumentContextParameterElement.Add(XmlElement.Create('ID', XmlNamespaceRAM, GetGuidelineID()));
        ExchangedDocumentContextElement.Add(GuidelineSpecifiedDocumentContextParameterElement);
        RootXMLNode.Add(ExchangedDocumentContextElement);
        ExchangedDocumentElement := XmlElement.Create('ExchangedDocument', XmlNamespaceRSM);
        ExchangedDocumentElement.Add(XmlElement.Create('ID', XmlNamespaceRAM, SalesCrMemoHeader."No."));
        ExchangedDocumentElement.Add(XmlElement.Create('TypeCode', XmlNamespaceRAM, '381')); // Credit memo type code

        IssueDateTimeElement := XmlElement.Create('IssueDateTime', XmlNamespaceRAM);
        IssueDateTimeElement.Add(XmlElement.Create('DateTimeString', XmlNamespaceUDT, XmlAttribute.Create('format', '102'), FormatDate(SalesCrMemoHeader."Document Date")));
        ExchangedDocumentElement.Add(IssueDateTimeElement);
        InsertLanguageID(ExchangedDocumentElement, SalesCrMemoHeader."Language Code");
        InsertHeaderNotes(ExchangedDocumentElement, SalesCrMemoHeader);
        RootXMLNode.Add(ExchangedDocumentElement);
        OnAfterInsertSalesCrMemoHeaderData(RootXMLNode, SalesCrMemoHeader);
    end;

    local procedure InsertApplicableHeaderTradeAgreement(var RootXMLNode: XmlElement; RecordVariant: Variant)
    var
        Customer: Record Customer;
        SalesInvoiceHeader: Record "Sales Invoice Header";
        SalesCrMemoHeader: Record "Sales Cr.Memo Header";
        TempBodyReportSelections: Record "Report Selections" temporary;
        ReportSelections: Record "Report Selections";
        DataTypeManagement: Codeunit "Data Type Management";
        HeaderRecordRef: RecordRef;
        HeaderTradeAgreementElement, SellerTradePartyElement, BuyerTradePartyElement, SpecifiedTaxRegistrationElement, IDElement : XmlElement;
        SellerOrderReferencedDocumentElement, BuyerOrderReferencedDocumentElement : XmlElement;
        ContactElement : XmlElement;
        SellerIDAttr, BuyerIDAttr : XmlAttribute;
        CustomerNo: Code[20];
        BuyerName2: Text[50];
        BuyerCounty: Text[30];
        QuoteNo: Code[20];
        ShipmentMethodCode: Code[10];
        SellerRegistrationNo: Text;
        CustomerGLN: Code[13];
        CustomerName: Text[100];
        Address: Text[100];
        Address2: Text[100];
        PostCode: Text[20];
        City: Text[50];
        CountryCode: Code[10];
        AddressCountryCode: Code[10];
        VATRegistrationNo: Text[20];
        CustomerRegistrationNo: Text[50];
        YourReference: Text[35];
        Contact: Text[100];
        CustomerEmail: Text[250];
        PhoneNumber: Text[30];
        OrderNo: Code[20];
        ExternalDocumentNo: Code[35];
        SellerStreetName: Text;
        SellerAdditionalStreetName: Text;
        SellerCityName: Text;
        SellerContactName: Text;
        SellerEmailAddress: Text;
        SellerPhoneNumber: Text;
        SellerPostalZone: Text;
        SellerCountryCode: Code[10];
        RespCentrCode: Code[10];
    begin
        if not DataTypeManagement.GetRecordRef(RecordVariant, HeaderRecordRef) then
            exit;
        case HeaderRecordRef.Number of
            Database::"Sales Invoice Header":
                begin
                    HeaderRecordRef.SetTable(SalesInvoiceHeader);
                    CustomerNo := SalesInvoiceHeader."Bill-to Customer No.";
                    CustomerName := SalesInvoiceHeader."Bill-to Name";
                    Address := SalesInvoiceHeader."Bill-to Address";
                    Address2 := SalesInvoiceHeader."Bill-to Address 2";
                    PostCode := SalesInvoiceHeader."Bill-to Post Code";
                    City := SalesInvoiceHeader."Bill-to City";
                    AddressCountryCode := SalesInvoiceHeader."Bill-to Country/Region Code";
                    BuyerName2 := SalesInvoiceHeader."Bill-to Name 2";
                    BuyerCounty := SalesInvoiceHeader."Bill-to County";
                    if UseSellToAsBuyer() then begin
                        CustomerNo := SalesInvoiceHeader."Sell-to Customer No.";
                        CustomerName := SalesInvoiceHeader."Sell-to Customer Name";
                        Address := SalesInvoiceHeader."Sell-to Address";
                        Address2 := SalesInvoiceHeader."Sell-to Address 2";
                        PostCode := SalesInvoiceHeader."Sell-to Post Code";
                        City := SalesInvoiceHeader."Sell-to City";
                        AddressCountryCode := SalesInvoiceHeader."Sell-to Country/Region Code";
                        BuyerName2 := SalesInvoiceHeader."Sell-to Customer Name 2";
                        BuyerCounty := SalesInvoiceHeader."Sell-to County";
                    end;
                    QuoteNo := SalesInvoiceHeader."Quote No.";
                    ShipmentMethodCode := SalesInvoiceHeader."Shipment Method Code";
                    CountryCode := SalesInvoiceHeader."VAT Country/Region Code";
                    VATRegistrationNo := SalesInvoiceHeader."VAT Registration No.";
                    YourReference := SalesInvoiceHeader."Your Reference";
                    Contact := SalesInvoiceHeader."Sell-to Contact";
                    ReportSelections.FindEmailBodyUsageForCust("Report Selection Usage"::"S.Invoice", CustomerNo, TempBodyReportSelections);
                    CustomerEmail := ReportSelections.GetEmailAddressExt("Report Selection Usage"::"S.Invoice".AsInteger(), RecordVariant, CustomerNo, TempBodyReportSelections);
                    PhoneNumber := SalesInvoiceHeader."Sell-to Phone No.";
                    OrderNo := SalesInvoiceHeader."Order No.";
                    ExternalDocumentNo := SalesInvoiceHeader."External Document No.";
                    RespCentrCode := SalesInvoiceHeader."Responsibility Center";
                    GetSellerContactInfo(SalesInvoiceHeader, SellerContactName, SellerPhoneNumber, SellerEmailAddress);
                end;
            Database::"Sales Cr.Memo Header":
                begin
                    HeaderRecordRef.SetTable(SalesCrMemoHeader);
                    CustomerNo := SalesCrMemoHeader."Bill-to Customer No.";
                    CustomerName := SalesCrMemoHeader."Bill-to Name";
                    Address := SalesCrMemoHeader."Bill-to Address";
                    Address2 := SalesCrMemoHeader."Bill-to Address 2";
                    PostCode := SalesCrMemoHeader."Bill-to Post Code";
                    City := SalesCrMemoHeader."Bill-to City";
                    AddressCountryCode := SalesCrMemoHeader."Bill-to Country/Region Code";
                    BuyerName2 := SalesCrMemoHeader."Bill-to Name 2";
                    BuyerCounty := SalesCrMemoHeader."Bill-to County";
                    if UseSellToAsBuyer() then begin
                        CustomerNo := SalesCrMemoHeader."Sell-to Customer No.";
                        CustomerName := SalesCrMemoHeader."Sell-to Customer Name";
                        Address := SalesCrMemoHeader."Sell-to Address";
                        Address2 := SalesCrMemoHeader."Sell-to Address 2";
                        PostCode := SalesCrMemoHeader."Sell-to Post Code";
                        City := SalesCrMemoHeader."Sell-to City";
                        AddressCountryCode := SalesCrMemoHeader."Sell-to Country/Region Code";
                        BuyerName2 := SalesCrMemoHeader."Sell-to Customer Name 2";
                        BuyerCounty := SalesCrMemoHeader."Sell-to County";
                    end;
                    ShipmentMethodCode := SalesCrMemoHeader."Shipment Method Code";
                    CountryCode := SalesCrMemoHeader."VAT Country/Region Code";
                    VATRegistrationNo := SalesCrMemoHeader."VAT Registration No.";
                    YourReference := SalesCrMemoHeader."Your Reference";
                    Contact := SalesCrMemoHeader."Sell-to Contact";
                    ReportSelections.FindEmailBodyUsageForCust("Report Selection Usage"::"S.Cr.Memo", CustomerNo, TempBodyReportSelections);
                    CustomerEmail := ReportSelections.GetEmailAddressExt("Report Selection Usage"::"S.Cr.Memo".AsInteger(), RecordVariant, CustomerNo, TempBodyReportSelections);
                    PhoneNumber := SalesCrMemoHeader."Sell-to Phone No.";
                    OrderNo := SalesCrMemoHeader."Return Order No.";
                    ExternalDocumentNo := SalesCrMemoHeader."External Document No.";
                    RespCentrCode := SalesCrMemoHeader."Responsibility Center";
                    GetSellerContactInfo(SalesCrMemoHeader, SellerContactName, SellerPhoneNumber, SellerEmailAddress);
                end;
        end;

        GetSellerPostalAddr(RespCentrCode, SellerStreetName, SellerAdditionalStreetName, SellerCityName, SellerPostalZone, SellerCountryCode);
        Customer.SetLoadFields("Use GLN in Electronic Document", GLN, "Registration Number", "Use Reg. No. in E-Document", "VAT Registration No.", "Our Account No.");
        if Customer.Get(CustomerNo) then begin
            if Customer."Use GLN in Electronic Document" then
                CustomerGLN := Customer.GLN;
            if Customer."Use Reg. No. in E-Document" and (Customer.GLN = '') and (Customer."VAT Registration No." = '') and (Customer."Registration Number" <> '') then
                CustomerRegistrationNo := Customer."Registration Number";
        end;
        HeaderTradeAgreementElement := XmlElement.Create('ApplicableHeaderTradeAgreement', XmlNamespaceRAM);
        // CII R74: no empty elements. XRechnung requires BT-10; the check on release reports a missing value.
        AddElementIfNotEmpty(HeaderTradeAgreementElement, 'BuyerReference', GetBuyerReference(RecordVariant));

        // Seller
        SellerTradePartyElement := XmlElement.Create('SellerTradeParty', XmlNamespaceRAM);
        // BT-29 seller identifier, EXTENDED only: our account no. at the customer.
        if IsExtendedProfile() then
            AddElementIfNotEmpty(SellerTradePartyElement, 'ID', Customer."Our Account No.");
        if CompanyInformation."Use GLN in Electronic Document" and (CompanyInformation.GLN <> '') then begin
            SellerIDAttr := XmlAttribute.Create('schemeID', GLNSchemeIDTok);
            SellerTradePartyElement.Add(XmlElement.Create('GlobalID', XmlNamespaceRAM, SellerIDAttr, CompanyInformation.GLN));
        end;
        SellerTradePartyElement.Add(XmlElement.Create('Name', XmlNamespaceRAM, CompanyInformation.Name));

        // BT-30 / BT-28, EXTENDED only. The registration no. is not repeated when it is written as tax number (FC) below.
        if IsExtendedProfile() then begin
            if not (CompanyInformation."Use Reg. No. in E-Document" and (CompanyInformation."VAT Registration No." = '') and (CompanyInformation.GLN = '')) then
                SellerRegistrationNo := CompanyInformation."Registration No.";
            InsertLegalOrganization(SellerTradePartyElement, SellerRegistrationNo, CompanyInformation."Name 2");
        end;

        // Seller Contact
        if SellerPhoneNumber <> '' then begin
            ContactElement := XmlElement.Create('DefinedTradeContact', XmlNamespaceRAM);
            AddElementIfNotEmpty(ContactElement, 'PersonName', SellerContactName);
            ContactElement.Add(XmlElement.Create('TelephoneUniversalCommunication', XmlNamespaceRAM,
                XmlElement.Create('CompleteNumber', XmlNamespaceRAM, SellerPhoneNumber)));
            if SellerEmailAddress <> '' then
                ContactElement.Add(XmlElement.Create('EmailURIUniversalCommunication', XmlNamespaceRAM,
                    XmlElement.Create('URIID', XmlNamespaceRAM, SellerEmailAddress)));
            SellerTradePartyElement.Add(ContactElement);
        end;

        // Seller Address
        SellerTradePartyElement.Add(CreatePostalTradeAddressElement(SellerPostalZone, SellerStreetName, SellerAdditionalStreetName, SellerCityName, SellerCountryCode, GetCountyIfExtended(GetSellerCounty(RespCentrCode))));

        //Seller E-Mail
        if CompanyInformation."E-Mail" <> '' then
            SellerTradePartyElement.Add(XmlElement.Create('URIUniversalCommunication', XmlNamespaceRAM,
                XmlElement.Create('URIID', XmlNamespaceRAM, XmlAttribute.Create('schemeID', 'EM'), CompanyInformation."E-Mail")));

        if CompanyInformation."VAT Registration No." <> '' then begin
            SellerIDAttr := XmlAttribute.Create('schemeID', 'VA');
            IDElement := XmlElement.Create('ID', XmlNamespaceRAM, SellerIDAttr, GetVATRegistrationNo(CompanyInformation."VAT Registration No.", SellerCountryCode));
            SpecifiedTaxRegistrationElement := XmlElement.Create('SpecifiedTaxRegistration', XmlNamespaceRAM);
            SpecifiedTaxRegistrationElement.Add(IDElement);
            SellerTradePartyElement.Add(SpecifiedTaxRegistrationElement);
        end else
            if CompanyInformation."Use Reg. No. in E-Document" and
               (CompanyInformation.GLN = '') and
               (CompanyInformation."Registration No." <> '')
            then begin
                SellerIDAttr := XmlAttribute.Create('schemeID', 'FC');
                IDElement := XmlElement.Create('ID', XmlNamespaceRAM, SellerIDAttr, CompanyInformation."Registration No.");
                SpecifiedTaxRegistrationElement := XmlElement.Create('SpecifiedTaxRegistration', XmlNamespaceRAM);
                SpecifiedTaxRegistrationElement.Add(IDElement);
                SellerTradePartyElement.Add(SpecifiedTaxRegistrationElement);
            end;
        HeaderTradeAgreementElement.Add(SellerTradePartyElement);

        // Buyer
        BuyerTradePartyElement := XmlElement.Create('BuyerTradeParty', XmlNamespaceRAM);
        if CustomerGLN <> '' then
            BuyerTradePartyElement.Add(XmlElement.Create('GlobalID', XmlNamespaceRAM, XmlAttribute.Create('schemeID', GLNSchemeIDTok), CustomerGLN));
        BuyerTradePartyElement.Add(XmlElement.Create('Name', XmlNamespaceRAM, CustomerName));

        // BT-47 as today; BT-45 trading name (Name 2) EXTENDED only.
        InsertLegalOrganization(BuyerTradePartyElement, CustomerRegistrationNo, BuyerName2);

        // Buyer Contact
        InsertBuyerContact(BuyerTradePartyElement, Contact, PhoneNumber, CustomerEmail);

        // Buyer Address. BT-55 is the country of the address; the VAT country (prefix of BT-48) can differ.
        BuyerTradePartyElement.Add(CreatePostalTradeAddressElement(PostCode, Address, Address2, City, GetCountryRegionCode(AddressCountryCode), GetCountyIfExtended(BuyerCounty)));

        // Buyer E-Mail
        if CustomerEmail <> '' then
            BuyerTradePartyElement.Add(XmlElement.Create('URIUniversalCommunication', XmlNamespaceRAM,
                XmlElement.Create('URIID', XmlNamespaceRAM, XmlAttribute.Create('schemeID', 'EM'), CustomerEmail)));

        if VATRegistrationNo <> '' then begin
            BuyerIDAttr := XmlAttribute.Create('schemeID', 'VA');
            IDElement := XmlElement.Create('ID', XmlNamespaceRAM, BuyerIDAttr, GetVATRegistrationNo(VATRegistrationNo, CountryCode));
            SpecifiedTaxRegistrationElement := XmlElement.Create('SpecifiedTaxRegistration', XmlNamespaceRAM);
            SpecifiedTaxRegistrationElement.Add(IDElement);
            BuyerTradePartyElement.Add(SpecifiedTaxRegistrationElement);
        end;
        HeaderTradeAgreementElement.Add(BuyerTradePartyElement);
        InsertDeliveryTerms(HeaderTradeAgreementElement, ShipmentMethodCode);

        // Seller Order Reference
        if OrderNo <> '' then begin
            SellerOrderReferencedDocumentElement := XmlElement.Create('SellerOrderReferencedDocument', XmlNamespaceRAM);
            SellerOrderReferencedDocumentElement.Add(XmlElement.Create('IssuerAssignedID', XmlNamespaceRAM, OrderNo));
            HeaderTradeAgreementElement.Add(SellerOrderReferencedDocumentElement);
        end;

        if ExternalDocumentNo <> '' then begin
            BuyerOrderReferencedDocumentElement := XmlElement.Create('BuyerOrderReferencedDocument', XmlNamespaceRAM);
            BuyerOrderReferencedDocumentElement.Add(XmlElement.Create('IssuerAssignedID', XmlNamespaceRAM, ExternalDocumentNo));
            HeaderTradeAgreementElement.Add(BuyerOrderReferencedDocumentElement);
        end;

        // BT-X-403 quotation and BT-11 project, EXTENDED only. Subscribers append after the project.
        if IsExtendedProfile() then
            InsertReferencedDocument(HeaderTradeAgreementElement, 'QuotationReferencedDocument', QuoteNo, 0, 0D);
        InsertProcuringProject(HeaderTradeAgreementElement, HeaderRecordRef);

        OnAfterInsertApplicableHeaderTradeAgreement(HeaderTradeAgreementElement, HeaderRecordRef);
        RootXMLNode.Add(HeaderTradeAgreementElement);
    end;

    local procedure InsertApplicableHeaderTradeDelivery(var RootXMLNode: XmlElement; SalesInvoiceHeader: Record "Sales Invoice Header")
    var
        DeliveryElement: XmlElement;
    begin
        DeliveryElement := XmlElement.Create('ApplicableHeaderTradeDelivery', XmlNamespaceRAM);

        if CurrentZUGFeRDProfile = CurrentZUGFeRDProfile::Extended then begin
            InsertShipToTradeParty(DeliveryElement, SalesInvoiceHeader."Sell-to Customer No.", SalesInvoiceHeader."Ship-to Code", SalesInvoiceHeader."Ship-to Name",
                SalesInvoiceHeader."Ship-to Address", SalesInvoiceHeader."Ship-to Address 2", SalesInvoiceHeader."Ship-to Post Code", SalesInvoiceHeader."Ship-to City",
                SalesInvoiceHeader."Ship-to Country/Region Code", SalesInvoiceHeader."Ship-to County", SalesInvoiceHeader."Ship-to Contact", SalesInvoiceHeader."Ship-to Phone No.");
            InsertShipFromTradeParty(DeliveryElement, SalesInvoiceHeader."Location Code");
        end else
            // XRechnung profile: unchanged sell-to address. The ship-to fix for all profiles is in PR #12462 / #12268.
            InsertShipToTradeParty(DeliveryElement, SalesInvoiceHeader."Sell-to Customer No.", SalesInvoiceHeader."Ship-to Code", SalesInvoiceHeader."Sell-to Customer Name",
                SalesInvoiceHeader."Sell-to Address", '', SalesInvoiceHeader."Sell-to Post Code", SalesInvoiceHeader."Sell-to City",
                SalesInvoiceHeader."Sell-to Country/Region Code", '', '', '');
        InsertActualDeliveryDate(DeliveryElement, SalesInvoiceHeader."Shipment Date");

        RaiseOnAfterInsertApplicableHeaderTradeDelivery(DeliveryElement, SalesInvoiceHeader);
        RootXMLNode.Add(DeliveryElement);
    end;

    local procedure InsertApplicableHeaderTradeDelivery(var RootXMLNode: XmlElement; SalesCrMemoHeader: Record "Sales Cr.Memo Header")
    var
        DeliveryElement: XmlElement;
    begin
        DeliveryElement := XmlElement.Create('ApplicableHeaderTradeDelivery', XmlNamespaceRAM);

        if CurrentZUGFeRDProfile = CurrentZUGFeRDProfile::Extended then begin
            InsertShipToTradeParty(DeliveryElement, SalesCrMemoHeader."Sell-to Customer No.", SalesCrMemoHeader."Ship-to Code", SalesCrMemoHeader."Ship-to Name",
                SalesCrMemoHeader."Ship-to Address", SalesCrMemoHeader."Ship-to Address 2", SalesCrMemoHeader."Ship-to Post Code", SalesCrMemoHeader."Ship-to City",
                SalesCrMemoHeader."Ship-to Country/Region Code", SalesCrMemoHeader."Ship-to County", SalesCrMemoHeader."Ship-to Contact", SalesCrMemoHeader."Ship-to Phone No.");
            InsertShipFromTradeParty(DeliveryElement, SalesCrMemoHeader."Location Code");
        end else
            // XRechnung profile: unchanged sell-to address. The ship-to fix for all profiles is in PR #12462 / #12268.
            InsertShipToTradeParty(DeliveryElement, SalesCrMemoHeader."Sell-to Customer No.", SalesCrMemoHeader."Ship-to Code", SalesCrMemoHeader."Sell-to Customer Name",
                SalesCrMemoHeader."Sell-to Address", '', SalesCrMemoHeader."Sell-to Post Code", SalesCrMemoHeader."Sell-to City",
                SalesCrMemoHeader."Sell-to Country/Region Code", '', '', '');
        InsertActualDeliveryDate(DeliveryElement, SalesCrMemoHeader."Shipment Date");

        RaiseOnAfterInsertApplicableHeaderTradeDelivery(DeliveryElement, SalesCrMemoHeader);
        RootXMLNode.Add(DeliveryElement);
    end;

    /// <summary>
    /// BG-13 deliver-to. The contact (DefinedTradeContact) and the county are EXTENDED only; callers pass them blank otherwise.
    /// </summary>
    local procedure InsertShipToTradeParty(var DeliveryElement: XmlElement; CustomerNo: Code[20]; ShipToCode: Code[10]; Name: Text; Address: Text; Address2: Text; PostCode: Text; City: Text; CountryRegionCode: Code[10]; County: Text; ContactName: Text; PhoneNo: Text)
    var
        ShipToPartyElement, ContactElement : XmlElement;
    begin
        ShipToPartyElement := XmlElement.Create('ShipToTradeParty', XmlNamespaceRAM);
        InsertDeliveryGLN(ShipToPartyElement, CustomerNo, ShipToCode);
        AddElementIfNotEmpty(ShipToPartyElement, 'Name', Name);
        if (ContactName <> '') or (PhoneNo <> '') then begin
            ContactElement := XmlElement.Create('DefinedTradeContact', XmlNamespaceRAM);
            AddElementIfNotEmpty(ContactElement, 'PersonName', ContactName);
            if PhoneNo <> '' then
                ContactElement.Add(XmlElement.Create('TelephoneUniversalCommunication', XmlNamespaceRAM,
                    XmlElement.Create('CompleteNumber', XmlNamespaceRAM, PhoneNo)));
            ShipToPartyElement.Add(ContactElement);
        end;
        ShipToPartyElement.Add(CreatePostalTradeAddressElement(PostCode, Address, Address2, City, GetCountryRegionCode(CountryRegionCode), County));
        DeliveryElement.Add(ShipToPartyElement);
    end;

    /// <summary>
    /// BG-X-30 ship-from, EXTENDED only: the location the goods leave from, when its address differs from the seller's.
    /// </summary>
    local procedure InsertShipFromTradeParty(var DeliveryElement: XmlElement; LocationCode: Code[10])
    var
        Location: Record Location;
        ShipFromPartyElement: XmlElement;
    begin
        if LocationCode = '' then
            exit;
        Location.SetLoadFields(Name, Address, "Address 2", "Post Code", City, "Country/Region Code", County);
        if not Location.Get(LocationCode) then
            exit;
        if Location.Address = '' then
            exit;
        if (Location.Address = CompanyInformation.Address) and (Location."Post Code" = CompanyInformation."Post Code") and (Location.City = CompanyInformation.City) then
            exit;
        ShipFromPartyElement := XmlElement.Create('ShipFromTradeParty', XmlNamespaceRAM);
        AddElementIfNotEmpty(ShipFromPartyElement, 'Name', Location.Name);
        ShipFromPartyElement.Add(CreatePostalTradeAddressElement(Location."Post Code", Location.Address, Location."Address 2", Location.City, GetCountryRegionCode(Location."Country/Region Code"), Location.County));
        DeliveryElement.Add(ShipFromPartyElement);
    end;

    local procedure InsertActualDeliveryDate(var DeliveryElement: XmlElement; ShipmentDate: Date)
    var
        ActualDeliveryDateElement, OccurrenceDateTimeElement : XmlElement;
    begin
        // CII R74: without a shipment date no delivery date, instead of the placeholder 17530101.
        if ShipmentDate = 0D then
            exit;
        ActualDeliveryDateElement := XmlElement.Create('ActualDeliverySupplyChainEvent', XmlNamespaceRAM);
        OccurrenceDateTimeElement := XmlElement.Create('OccurrenceDateTime', XmlNamespaceRAM);
        OccurrenceDateTimeElement.Add(XmlElement.Create('DateTimeString', XmlNamespaceUDT, XmlAttribute.Create('format', '102'), FormatDate(ShipmentDate)));
        ActualDeliveryDateElement.Add(OccurrenceDateTimeElement);
        DeliveryElement.Add(ActualDeliveryDateElement);
    end;

    /// <summary>
    /// Postal address (BG-5, BG-8, BG-15, …). CII R74: empty parts are not written; CountryID is mandatory.
    /// </summary>
    local procedure CreatePostalTradeAddressElement(PostCode: Text; LineOne: Text; LineTwo: Text; City: Text; CountryRegionCode: Code[10]; County: Text) PostalTradeAddressElement: XmlElement
    begin
        PostalTradeAddressElement := XmlElement.Create('PostalTradeAddress', XmlNamespaceRAM);
        AddElementIfNotEmpty(PostalTradeAddressElement, 'PostcodeCode', PostCode);
        AddElementIfNotEmpty(PostalTradeAddressElement, 'LineOne', LineOne);
        AddElementIfNotEmpty(PostalTradeAddressElement, 'LineTwo', LineTwo);
        AddElementIfNotEmpty(PostalTradeAddressElement, 'CityName', City);
        PostalTradeAddressElement.Add(XmlElement.Create('CountryID', XmlNamespaceRAM, GetCountryISOCode(CountryRegionCode)));
        AddElementIfNotEmpty(PostalTradeAddressElement, 'CountrySubDivisionName', County);
    end;

    local procedure AddElementIfNotEmpty(var ParentElement: XmlElement; ElementName: Text; Value: Text)
    begin
        if Value = '' then
            exit;
        ParentElement.Add(XmlElement.Create(ElementName, XmlNamespaceRAM, Value));
    end;

    /// <summary>
    /// Decision D-05 (proposal, open): in EXTENDED the buyer (BG-7) is the customer whose VAT data was posted, as
    /// G/L Setup "Bill-to/Sell-to VAT Calc." selects it. BR-IC-02 / BR-AE-02 check the buyer's VAT ID (BT-48).
    /// The XRechnung profile keeps the bill-to customer.
    /// </summary>
    local procedure UseSellToAsBuyer(): Boolean
    begin
        if CurrentZUGFeRDProfile <> CurrentZUGFeRDProfile::Extended then
            exit(false);
        exit(GeneralLedgerSetup."Bill-to/Sell-to VAT Calc." = GeneralLedgerSetup."Bill-to/Sell-to VAT Calc."::"Sell-to/Buy-from No.");
    end;

    /// <summary>
    /// BG-X-36 invoicee, EXTENDED only. Decision D-06 (proposal, open): written only when the buyer is the sell-to customer
    /// and the bill-to customer is a different one, so that the invoice recipient is not lost. BR-FXEXT-03: only the VAT ID.
    /// </summary>
    local procedure InsertInvoiceeTradeParty(var SettlementElement: XmlElement; SellToCustomerNo: Code[20]; BillToCustomerNo: Code[20]; Name: Text; Address: Text; Address2: Text; PostCode: Text; City: Text; CountryRegionCode: Code[10]; County: Text)
    var
        Customer: Record Customer;
        InvoiceePartyElement, SpecifiedTaxRegistrationElement : XmlElement;
    begin
        if not UseSellToAsBuyer() then
            exit;
        if (BillToCustomerNo = '') or (BillToCustomerNo = SellToCustomerNo) then
            exit;
        Customer.SetLoadFields("Use GLN in Electronic Document", GLN, "VAT Registration No.", "Country/Region Code");
        if not Customer.Get(BillToCustomerNo) then
            Clear(Customer);
        InvoiceePartyElement := XmlElement.Create('InvoiceeTradeParty', XmlNamespaceRAM);
        if Customer."Use GLN in Electronic Document" and (Customer.GLN <> '') then
            InvoiceePartyElement.Add(XmlElement.Create('GlobalID', XmlNamespaceRAM, XmlAttribute.Create('schemeID', GLNSchemeIDTok), Customer.GLN));
        AddElementIfNotEmpty(InvoiceePartyElement, 'Name', Name);
        InvoiceePartyElement.Add(CreatePostalTradeAddressElement(PostCode, Address, Address2, City, GetCountryRegionCode(CountryRegionCode), County));
        if Customer."VAT Registration No." <> '' then begin
            SpecifiedTaxRegistrationElement := XmlElement.Create('SpecifiedTaxRegistration', XmlNamespaceRAM);
            SpecifiedTaxRegistrationElement.Add(XmlElement.Create('ID', XmlNamespaceRAM, XmlAttribute.Create('schemeID', 'VA'), GetVATRegistrationNo(Customer."VAT Registration No.", GetCountryRegionCode(Customer."Country/Region Code"))));
            InvoiceePartyElement.Add(SpecifiedTaxRegistrationElement);
        end;
        SettlementElement.Add(InvoiceePartyElement);
    end;

    local procedure InsertDeliveryGLN(var ShipToPartyElement: XmlElement; CustomerNo: Code[20]; ShipToCode: Code[10])
    var
        Customer: Record Customer;
        ShipToAddress: Record "Ship-to Address";
        DeliveryGLN: Code[13];
    begin
        Customer.SetLoadFields("Use GLN in Electronic Document", GLN);
        if not Customer.Get(CustomerNo) then
            exit;
        if not Customer."Use GLN in Electronic Document" then
            exit;
        ShipToAddress.SetLoadFields(GLN);
        if (ShipToCode <> '') and ShipToAddress.Get(CustomerNo, ShipToCode) then
            DeliveryGLN := ShipToAddress.GLN;
        if DeliveryGLN = '' then
            DeliveryGLN := Customer.GLN;
        if DeliveryGLN <> '' then
            ShipToPartyElement.Add(XmlElement.Create('GlobalID', XmlNamespaceRAM, XmlAttribute.Create('schemeID', GLNSchemeIDTok), DeliveryGLN));
    end;

    local procedure InsertApplicableHeaderTradeSettlement(var RootXMLNode: XmlElement; var SalesInvHeader: Record "Sales Invoice Header"; var SalesInvLine: Record "Sales Invoice Line"; CurrencyCode: Code[10]; var LineAmounts: Dictionary of [Text, Decimal]; var LineDiscAmount: Dictionary of [Decimal, Decimal])
    var
        SettlementElement, MonetarySummationElement : XmlElement;
        ItemChargeTotalAmount: Decimal;
        ItemChargeAllowanceTotalAmount: Decimal;
        PrepaidAmount: Decimal;
    begin
        SettlementElement := XmlElement.Create('ApplicableHeaderTradeSettlement', XmlNamespaceRAM);
        // BT-90 Bank assigned creditor identifier. The CII HeaderTradeSettlement sequence puts
        // CreditorReferenceID before InvoiceCurrencyCode, so it is added first.
        InsertCreditorReference(SettlementElement, SalesInvHeader."Payment Method Code", SalesInvHeader."Company Bank Account Code");
        InsertPaymentReference(SettlementElement, SalesInvHeader."Payment Reference", SalesInvHeader."No.");

        InsertTaxCurrencyCode(SettlementElement, SalesInvHeader."Currency Code", "Gen. Journal Document Type"::Invoice, SalesInvHeader."No.", SalesInvHeader."Posting Date");
        SettlementElement.Add(XmlElement.Create('InvoiceCurrencyCode', XmlNamespaceRAM, CurrencyCode));
        InsertInvoiceeTradeParty(SettlementElement, SalesInvHeader."Sell-to Customer No.", SalesInvHeader."Bill-to Customer No.", SalesInvHeader."Bill-to Name",
            SalesInvHeader."Bill-to Address", SalesInvHeader."Bill-to Address 2", SalesInvHeader."Bill-to Post Code", SalesInvHeader."Bill-to City",
            SalesInvHeader."Bill-to Country/Region Code", SalesInvHeader."Bill-to County");
        InsertCurrencyExchange(SettlementElement, SalesInvHeader."Currency Code", SalesInvHeader."Currency Factor", SalesInvHeader."Posting Date");
        InsertPaymentMethod(SettlementElement, SalesInvHeader."Payment Method Code", SalesInvHeader."Company Bank Account Code", SalesInvHeader."Direct Debit Mandate ID", SalesInvHeader);
        InsertTradeTax(SettlementElement, SalesInvLine);
        InsertInvDiscountAllowanceCharge(SettlementElement, SalesInvLine, LineDiscAmount, LineAmounts);
        InsertItemChargeAllowanceCharges(SettlementElement, SalesInvLine);

        InsertPaymentTerms(SettlementElement, SalesInvHeader."Payment Terms Code", SalesInvHeader."Due Date", GetDirectDebitMandateID(SalesInvHeader."Payment Method Code", SalesInvHeader."Direct Debit Mandate ID"),
            SalesInvHeader."Document Date", SalesInvHeader."Pmt. Discount Date", SalesInvHeader."Payment Discount %", GetPaymentDiscountBasis(LineAmounts, SalesInvLine.FieldName(Amount), SalesInvLine.FieldName("Amount Including VAT")),
            GetPaymentDiscountAmount("Gen. Journal Document Type"::Invoice, SalesInvHeader."No.", SalesInvHeader."Bill-to Customer No."), SalesInvHeader."Currency Code");
        GetDocumentLevelItemChargeTotals(SalesInvLine, ItemChargeTotalAmount, ItemChargeAllowanceTotalAmount);
        MonetarySummationElement := XmlElement.Create('SpecifiedTradeSettlementHeaderMonetarySummation', XmlNamespaceRAM);
        MonetarySummationElement.Add(XmlElement.Create('LineTotalAmount', XmlNamespaceRAM, FormatDecimal(LineAmounts.Get(SalesInvLine.FieldName(Amount)) + LineAmounts.Get(SalesInvLine.FieldName("Inv. Discount Amount")) - ItemChargeTotalAmount + ItemChargeAllowanceTotalAmount)));
        if ItemChargeTotalAmount > 0 then
            MonetarySummationElement.Add(XmlElement.Create('ChargeTotalAmount', XmlNamespaceRAM, FormatDecimal(ItemChargeTotalAmount)));
        MonetarySummationElement.Add(XmlElement.Create('AllowanceTotalAmount', XmlNamespaceRAM, FormatDecimal(LineAmounts.Get(SalesInvLine.FieldName("Inv. Discount Amount")) + ItemChargeAllowanceTotalAmount)));
        MonetarySummationElement.Add(XmlElement.Create('TaxBasisTotalAmount', XmlNamespaceRAM, FormatDecimal(LineAmounts.Get(SalesInvLine.FieldName(Amount)))));
        MonetarySummationElement.Add(XmlElement.Create('TaxTotalAmount', XmlNamespaceRAM, XmlAttribute.Create('currencyID', CurrencyCode), FormatDecimal(LineAmounts.Get(SalesInvLine.FieldName("Amount Including VAT")) - LineAmounts.Get(SalesInvLine.FieldName(Amount)))));
        InsertTaxTotalInLCY(MonetarySummationElement, SalesInvHeader."Currency Code", "Gen. Journal Document Type"::Invoice, SalesInvHeader."No.", SalesInvHeader."Posting Date");
        MonetarySummationElement.Add(XmlElement.Create('GrandTotalAmount', XmlNamespaceRAM, FormatDecimal(LineAmounts.Get(SalesInvLine.FieldName("Amount Including VAT")))));
        PrepaidAmount := GetPrepaidAmount(SalesInvHeader);
        MonetarySummationElement.Add(XmlElement.Create('TotalPrepaidAmount', XmlNamespaceRAM, FormatDecimal(PrepaidAmount)));
        MonetarySummationElement.Add(XmlElement.Create('DuePayableAmount', XmlNamespaceRAM, FormatDecimal(LineAmounts.Get(SalesInvLine.FieldName("Amount Including VAT")) - PrepaidAmount)));

        SettlementElement.Add(MonetarySummationElement);
        RootXMLNode.Add(SettlementElement);
    end;

    local procedure InsertApplicableHeaderTradeSettlement(var RootXMLNode: XmlElement; var SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var SalesCrMemoLine: Record "Sales Cr.Memo Line"; CurrencyCode: Code[10]; var LineAmounts: Dictionary of [Text, Decimal]; var LineDiscAmount: Dictionary of [Decimal, Decimal])
    var
        SettlementElement, MonetarySummationElement : XmlElement;
        ItemChargeTotalAmount: Decimal;
        ItemChargeAllowanceTotalAmount: Decimal;
        PrepaidAmount: Decimal;
    begin
        SettlementElement := XmlElement.Create('ApplicableHeaderTradeSettlement', XmlNamespaceRAM);

        InsertTaxCurrencyCode(SettlementElement, SalesCrMemoHeader."Currency Code", "Gen. Journal Document Type"::"Credit Memo", SalesCrMemoHeader."No.", SalesCrMemoHeader."Posting Date");
        SettlementElement.Add(XmlElement.Create('InvoiceCurrencyCode', XmlNamespaceRAM, CurrencyCode));
        InsertInvoiceeTradeParty(SettlementElement, SalesCrMemoHeader."Sell-to Customer No.", SalesCrMemoHeader."Bill-to Customer No.", SalesCrMemoHeader."Bill-to Name",
            SalesCrMemoHeader."Bill-to Address", SalesCrMemoHeader."Bill-to Address 2", SalesCrMemoHeader."Bill-to Post Code", SalesCrMemoHeader."Bill-to City",
            SalesCrMemoHeader."Bill-to Country/Region Code", SalesCrMemoHeader."Bill-to County");
        InsertCurrencyExchange(SettlementElement, SalesCrMemoHeader."Currency Code", SalesCrMemoHeader."Currency Factor", SalesCrMemoHeader."Posting Date");
        InsertPaymentMethod(SettlementElement, SalesCrMemoHeader."Payment Method Code", SalesCrMemoHeader."Company Bank Account Code", '', SalesCrMemoHeader);
        InsertTradeTax(SettlementElement, SalesCrMemoLine);
        InsertInvDiscountAllowanceCharge(SettlementElement, SalesCrMemoLine, LineDiscAmount, LineAmounts);
        InsertItemChargeAllowanceCharges(SettlementElement, SalesCrMemoLine);

        InsertPaymentTerms(SettlementElement, SalesCrMemoHeader."Payment Terms Code", SalesCrMemoHeader."Due Date", '',
            SalesCrMemoHeader."Document Date", SalesCrMemoHeader."Pmt. Discount Date", SalesCrMemoHeader."Payment Discount %", GetPaymentDiscountBasis(LineAmounts, SalesCrMemoLine.FieldName(Amount), SalesCrMemoLine.FieldName("Amount Including VAT")),
            GetPaymentDiscountAmount("Gen. Journal Document Type"::"Credit Memo", SalesCrMemoHeader."No.", SalesCrMemoHeader."Bill-to Customer No."), SalesCrMemoHeader."Currency Code");
        GetDocumentLevelItemChargeTotals(SalesCrMemoLine, ItemChargeTotalAmount, ItemChargeAllowanceTotalAmount);
        MonetarySummationElement := XmlElement.Create('SpecifiedTradeSettlementHeaderMonetarySummation', XmlNamespaceRAM);
        MonetarySummationElement.Add(XmlElement.Create('LineTotalAmount', XmlNamespaceRAM, FormatDecimal(LineAmounts.Get(SalesCrMemoLine.FieldName(Amount)) + LineAmounts.Get(SalesCrMemoLine.FieldName("Inv. Discount Amount")) - ItemChargeTotalAmount + ItemChargeAllowanceTotalAmount)));
        if ItemChargeTotalAmount > 0 then
            MonetarySummationElement.Add(XmlElement.Create('ChargeTotalAmount', XmlNamespaceRAM, FormatDecimal(ItemChargeTotalAmount)));
        MonetarySummationElement.Add(XmlElement.Create('AllowanceTotalAmount', XmlNamespaceRAM, FormatDecimal(LineAmounts.Get(SalesCrMemoLine.FieldName("Inv. Discount Amount")) + ItemChargeAllowanceTotalAmount)));
        MonetarySummationElement.Add(XmlElement.Create('TaxBasisTotalAmount', XmlNamespaceRAM, FormatDecimal(LineAmounts.Get(SalesCrMemoLine.FieldName(Amount)))));
        MonetarySummationElement.Add(XmlElement.Create('TaxTotalAmount', XmlNamespaceRAM, XmlAttribute.Create('currencyID', CurrencyCode), FormatDecimal(LineAmounts.Get(SalesCrMemoLine.FieldName("Amount Including VAT")) - LineAmounts.Get(SalesCrMemoLine.FieldName(Amount)))));
        InsertTaxTotalInLCY(MonetarySummationElement, SalesCrMemoHeader."Currency Code", "Gen. Journal Document Type"::"Credit Memo", SalesCrMemoHeader."No.", SalesCrMemoHeader."Posting Date");
        MonetarySummationElement.Add(XmlElement.Create('GrandTotalAmount', XmlNamespaceRAM, FormatDecimal(LineAmounts.Get(SalesCrMemoLine.FieldName("Amount Including VAT")))));
        PrepaidAmount := GetPrepaidAmount(SalesCrMemoHeader);
        MonetarySummationElement.Add(XmlElement.Create('TotalPrepaidAmount', XmlNamespaceRAM, FormatDecimal(PrepaidAmount)));
        MonetarySummationElement.Add(XmlElement.Create('DuePayableAmount', XmlNamespaceRAM, FormatDecimal(LineAmounts.Get(SalesCrMemoLine.FieldName("Amount Including VAT")) - PrepaidAmount)));

        SettlementElement.Add(MonetarySummationElement);
        InsertPrecedingInvoiceReference(SettlementElement, SalesCrMemoHeader."Applies-to Doc. Type", SalesCrMemoHeader."Applies-to Doc. No.");
        RootXMLNode.Add(SettlementElement);
    end;

    local procedure InsertBuyerContact(var BuyerTradePartyElement: XmlElement; ContactName: Text; PhoneNumber: Text; EmailAddress: Text)
    var
        ContactElement: XmlElement;
    begin
        if (ContactName = '') and (PhoneNumber = '') and (EmailAddress = '') then
            exit;

        ContactElement := XmlElement.Create('DefinedTradeContact', XmlNamespaceRAM);

        if ContactName <> '' then
            ContactElement.Add(XmlElement.Create('PersonName', XmlNamespaceRAM, ContactName));

        if PhoneNumber <> '' then
            ContactElement.Add(XmlElement.Create('TelephoneUniversalCommunication', XmlNamespaceRAM,
                XmlElement.Create('CompleteNumber', XmlNamespaceRAM, PhoneNumber)));

        if EmailAddress <> '' then
            ContactElement.Add(XmlElement.Create('EmailURIUniversalCommunication', XmlNamespaceRAM,
                XmlElement.Create('URIID', XmlNamespaceRAM, EmailAddress)));

        BuyerTradePartyElement.Add(ContactElement);
    end;

    /// <summary>
    /// VAT breakdown (BG-23). Decision D-15: one breakdown per VAT category, rate and exemption reason, not per rate only,
    /// so that e.g. E 0 % and Z 0 % are not merged (BR-E-08, BR-Z-08). The amounts are summed from the lines, as
    /// InsertVATAmounts does for LineAmount / LineVATAmount.
    /// </summary>
    local procedure InsertTradeTax(var SettlementElement: XmlElement; var SalesInvLine: Record "Sales Invoice Line")
    begin
        ClearVATBreakdown();
        if SalesInvLine.FindSet() then
            repeat
                AddToVATBreakdown(SalesInvLine."Tax Category", SalesInvLine."VAT Bus. Posting Group", SalesInvLine."VAT Prod. Posting Group", SalesInvLine."VAT %", SalesInvLine.Amount, SalesInvLine."Amount Including VAT" - SalesInvLine.Amount);
            until SalesInvLine.Next() = 0;
        TempAdditionalSalesInvLine.Reset();
        if TempAdditionalSalesInvLine.FindSet() then
            repeat
                AddToVATBreakdown(TempAdditionalSalesInvLine."Tax Category", TempAdditionalSalesInvLine."VAT Bus. Posting Group", TempAdditionalSalesInvLine."VAT Prod. Posting Group", TempAdditionalSalesInvLine."VAT %", TempAdditionalSalesInvLine.Amount, TempAdditionalSalesInvLine."Amount Including VAT" - TempAdditionalSalesInvLine.Amount);
            until TempAdditionalSalesInvLine.Next() = 0;
        InsertVATBreakdown(SettlementElement);

        SalesInvLine.SetRange("VAT %");
        SalesInvLine.SetRange("VAT Calculation Type");
    end;

    local procedure InsertTradeTax(var SettlementElement: XmlElement; var SalesCrMemoLine: Record "Sales Cr.Memo Line")
    begin
        ClearVATBreakdown();
        if SalesCrMemoLine.FindSet() then
            repeat
                AddToVATBreakdown(SalesCrMemoLine."Tax Category", SalesCrMemoLine."VAT Bus. Posting Group", SalesCrMemoLine."VAT Prod. Posting Group", SalesCrMemoLine."VAT %", SalesCrMemoLine.Amount, SalesCrMemoLine."Amount Including VAT" - SalesCrMemoLine.Amount);
            until SalesCrMemoLine.Next() = 0;
        TempAdditionalSalesCrMemoLine.Reset();
        if TempAdditionalSalesCrMemoLine.FindSet() then
            repeat
                AddToVATBreakdown(TempAdditionalSalesCrMemoLine."Tax Category", TempAdditionalSalesCrMemoLine."VAT Bus. Posting Group", TempAdditionalSalesCrMemoLine."VAT Prod. Posting Group", TempAdditionalSalesCrMemoLine."VAT %", TempAdditionalSalesCrMemoLine.Amount, TempAdditionalSalesCrMemoLine."Amount Including VAT" - TempAdditionalSalesCrMemoLine.Amount);
            until TempAdditionalSalesCrMemoLine.Next() = 0;
        InsertVATBreakdown(SettlementElement);

        SalesCrMemoLine.SetRange("VAT %");
        SalesCrMemoLine.SetRange("VAT Calculation Type");
    end;

    local procedure ClearVATBreakdown()
    begin
        Clear(VATBreakdownKeys);
        Clear(VATBreakdownBasisAmounts);
        Clear(VATBreakdownVATAmounts);
        Clear(VATBreakdownCategories);
        Clear(VATBreakdownRates);
        Clear(VATBreakdownBusPostingGroups);
        Clear(VATBreakdownProdPostingGroups);
    end;

    local procedure AddToVATBreakdown(TaxCategory: Code[10]; VATBusPostingGroup: Code[20]; VATProdPostingGroup: Code[20]; VATPercent: Decimal; BasisAmount: Decimal; VATAmount: Decimal)
    var
        CategoryID: Text;
        VATEXCode: Text;
        VATClauseDescription: Text;
        BreakdownKey: Text;
    begin
        CategoryID := GetTaxCategoryID(TaxCategory, VATBusPostingGroup, VATProdPostingGroup);
        PeppolVATHelper.GetVATClauseInfo(VATBusPostingGroup, VATProdPostingGroup, DocumentLanguageCode, VATEXCode, VATClauseDescription);
        BreakdownKey := CategoryID + '|' + FormatFiveDecimal(VATPercent) + '|' + VATEXCode;
        if not VATBreakdownKeys.Contains(BreakdownKey) then begin
            VATBreakdownKeys.Add(BreakdownKey);
            VATBreakdownBasisAmounts.Add(BreakdownKey, 0);
            VATBreakdownVATAmounts.Add(BreakdownKey, 0);
            VATBreakdownCategories.Add(BreakdownKey, CategoryID);
            VATBreakdownRates.Add(BreakdownKey, VATPercent);
            // The VAT clause of the first line of the group gives the exemption reason text.
            VATBreakdownBusPostingGroups.Add(BreakdownKey, VATBusPostingGroup);
            VATBreakdownProdPostingGroups.Add(BreakdownKey, VATProdPostingGroup);
        end;
        VATBreakdownBasisAmounts.Set(BreakdownKey, VATBreakdownBasisAmounts.Get(BreakdownKey) + BasisAmount);
        VATBreakdownVATAmounts.Set(BreakdownKey, VATBreakdownVATAmounts.Get(BreakdownKey) + VATAmount);
    end;

    local procedure InsertVATBreakdown(var SettlementElement: XmlElement)
    var
        BreakdownKey: Text;
        VATEXCode: Text;
        VATClauseDescription: Text;
    begin
        foreach BreakdownKey in VATBreakdownKeys do begin
            PeppolVATHelper.GetVATClauseInfo(VATBreakdownBusPostingGroups.Get(BreakdownKey), VATBreakdownProdPostingGroups.Get(BreakdownKey), DocumentLanguageCode, VATEXCode, VATClauseDescription);
            InsertTaxElement(SettlementElement, FormatDecimal(VATBreakdownVATAmounts.Get(BreakdownKey)), FormatDecimal(VATBreakdownBasisAmounts.Get(BreakdownKey)),
                VATBreakdownCategories.Get(BreakdownKey), FormatFiveDecimal(VATBreakdownRates.Get(BreakdownKey)), VATEXCode, VATClauseDescription);
        end;
    end;

    local procedure InsertTaxElement(var SettlementElement: XmlElement; CalculatedAmount: Text; BasisAmount: Text; CategoryCode: Text; RateApplicablePercent: Text; VATEXCode: Text; VATClauseDescription: Text)
    var
        TaxElement: XmlElement;
    begin
        TaxElement := XmlElement.Create('ApplicableTradeTax', XmlNamespaceRAM);
        TaxElement.Add(XmlElement.Create('CalculatedAmount', XmlNamespaceRAM, CalculatedAmount));
        TaxElement.Add(XmlElement.Create('TypeCode', XmlNamespaceRAM, 'VAT'));
        if VATClauseDescription <> '' then
            TaxElement.Add(XmlElement.Create('ExemptionReason', XmlNamespaceRAM, VATClauseDescription));
        TaxElement.Add(XmlElement.Create('BasisAmount', XmlNamespaceRAM, BasisAmount));
        TaxElement.Add(XmlElement.Create('CategoryCode', XmlNamespaceRAM, CategoryCode));
        if VATEXCode <> '' then
            TaxElement.Add(XmlElement.Create('ExemptionReasonCode', XmlNamespaceRAM, VATEXCode));
        TaxElement.Add(XmlElement.Create('RateApplicablePercent', XmlNamespaceRAM, RateApplicablePercent));
        SettlementElement.Add(TaxElement);
    end;

    local procedure InsertCategoryTradeTax(var RootElement: XmlElement; CategoryCode: Text; RateApplicablePercent: Text)
    var
        TaxElement: XmlElement;
    begin
        TaxElement := XmlElement.Create('CategoryTradeTax', XmlNamespaceRAM);
        TaxElement.Add(XmlElement.Create('TypeCode', XmlNamespaceRAM, 'VAT'));
        TaxElement.Add(XmlElement.Create('CategoryCode', XmlNamespaceRAM, CategoryCode));
        TaxElement.Add(XmlElement.Create('RateApplicablePercent', XmlNamespaceRAM, RateApplicablePercent));
        RootElement.Add(TaxElement);
    end;

    procedure InsertSupplyChainTradeTransaction(var RootXMLNode: XmlElement; SalesInvoiceHeader: Record "Sales Invoice Header"; var SalesInvoiceLine: Record "Sales Invoice Line"; CurrencyCode: Code[10]; Currency: Record Currency; var LineAmount: Dictionary of [Decimal, Decimal]; var LineVATAmount: Dictionary of [Decimal, Decimal]; var LineAmounts: Dictionary of [Text, Decimal]; var LineDiscAmount: Dictionary of [Decimal, Decimal])
    var
        SupplyChainTradeTransactionElement: XmlElement;
    begin
        SupplyChainTradeTransactionElement := XmlElement.Create('SupplyChainTradeTransaction', XmlNamespaceRSM);
        if SalesInvoiceLine.FindSet() then
            repeat
                if not IsItemChargeAllowanceCharge(SalesInvoiceLine."Line No.") then
                    InsertInvoiceLine(SupplyChainTradeTransactionElement, SalesInvoiceLine, Currency, CurrencyCode, SalesInvoiceHeader."Prices Including VAT");
            until SalesInvoiceLine.Next() = 0;
        TempAdditionalSalesInvLine.Reset();
        if TempAdditionalSalesInvLine.FindSet() then
            repeat
                InsertInvoiceLine(SupplyChainTradeTransactionElement, TempAdditionalSalesInvLine, Currency, CurrencyCode, SalesInvoiceHeader."Prices Including VAT");
            until TempAdditionalSalesInvLine.Next() = 0;
        InsertApplicableHeaderTradeAgreement(SupplyChainTradeTransactionElement, SalesInvoiceHeader);
        InsertApplicableHeaderTradeDelivery(SupplyChainTradeTransactionElement, SalesInvoiceHeader);
        InsertApplicableHeaderTradeSettlement(SupplyChainTradeTransactionElement, SalesInvoiceHeader, SalesInvoiceLine, CurrencyCode, LineAmounts, LineDiscAmount);

        RootXMLNode.Add(SupplyChainTradeTransactionElement);
    end;

    local procedure InsertBillingSpecifiedPeriod(var RootElement: XmlElement; StartDate: Date; EndDate: Date);
    var
        BillingSpecifiedPeriodElement: XmlElement;
        StartDateElement: XmlElement;
        EndDateElement: XmlElement;
    begin
        BillingSpecifiedPeriodElement := XmlElement.Create('BillingSpecifiedPeriod', XmlNamespaceRAM);
        StartDateElement := XmlElement.Create('StartDateTime', XmlNamespaceRAM);
        StartDateElement.Add(XmlElement.Create('DateTimeString', XmlNamespaceUDT, XmlAttribute.Create('format', '102'), FormatDate(StartDate)));
        BillingSpecifiedPeriodElement.Add(StartDateElement);
        EndDateElement := XmlElement.Create('EndDateTime', XmlNamespaceRAM);
        EndDateElement.Add(XmlElement.Create('DateTimeString', XmlNamespaceUDT, XmlAttribute.Create('format', '102'), FormatDate(EndDate)));
        BillingSpecifiedPeriodElement.Add(EndDateElement);
        RootElement.Add(BillingSpecifiedPeriodElement);
    end;

    local procedure InsertGlobalID(var SpecifiedTradeProductElement: XmlElement; ItemNo: Code[20])
    var
        GTIN: Code[14];
    begin
        GTIN := GetItemGTIN(ItemNo);
        if GTIN = '' then
            exit;

        SpecifiedTradeProductElement.Add(XmlElement.Create('GlobalID', XmlNamespaceRAM, XmlAttribute.Create('schemeID', '0160'), GTIN));
    end;

    local procedure GetItemGTIN(ItemNo: Code[20]) GTIN: Code[14]
    var
        Item: Record Item;
    begin
        if ItemGTINCache.Get(ItemNo, GTIN) then
            exit;

        Item.SetLoadFields(Item.GTIN);
        if Item.Get(ItemNo) then
            GTIN := Item.GTIN;
        ItemGTINCache.Add(ItemNo, GTIN);
    end;

    local procedure InsertInvoiceLine(var SupplyChainTradeTransactionElement: XmlElement; var SalesInvoiceLine: Record "Sales Invoice Line"; Currency: Record Currency; CurrencyCode: Code[10]; PricesIncVAT: Boolean)
    var
        LineRecordRef: RecordRef;
        InvoiceLineElement: XmlElement;
        AssociatedDocumentLineElement: XmlElement;
        SpecifiedLineTradeAgreementElement: XmlElement;
        NetPriceProductTradePriceElement: XmlElement;
        ChargeAmountElement: XmlElement;
        SpecifiedLineTradeDeliveryElement: XmlElement;
        BilledQuantityElement: XmlElement;
        SpecifiedLineTradeSettlementElement: XmlElement;
        SpecifiedTradeSettlementLineMonetarySummationElement: XmlElement;
        IsHandled: Boolean;
    begin
        IsHandled := false;
        OnBeforeInsertInvoiceLine(SupplyChainTradeTransactionElement, SalesInvoiceLine, Currency, CurrencyCode, PricesIncVAT, IsHandled);
        if not IsHandled then begin
            InvoiceLineElement := XmlElement.Create('IncludedSupplyChainTradeLineItem', XmlNamespaceRAM);
            if PricesIncVAT then
                ExcludeVAT(SalesInvoiceLine, Currency."Amount Rounding Precision");
            LineRecordRef.GetTable(SalesInvoiceLine);
            AssociatedDocumentLineElement := CreateAssociatedDocumentLineElement(LineRecordRef, SalesInvoiceLine."Line No.");
            InvoiceLineElement.Add(AssociatedDocumentLineElement);

            InsertTradeProduct(InvoiceLineElement, SalesInvoiceLine.Type = SalesInvoiceLine.Type::Item, SalesInvoiceLine."No.", SalesInvoiceLine."Item Reference No.",
                SalesInvoiceLine.Description, SalesInvoiceLine."Description 2", SalesInvoiceLine."Variant Code");

            SpecifiedLineTradeAgreementElement := XmlElement.Create('SpecifiedLineTradeAgreement', XmlNamespaceRAM);
            InsertLineReferences(SpecifiedLineTradeAgreementElement, SalesInvoiceLine."Order No.", SalesInvoiceLine."Order Line No.", SalesInvoiceLine."Blanket Order No.", SalesInvoiceLine."Blanket Order Line No.");
            NetPriceProductTradePriceElement := XmlElement.Create('NetPriceProductTradePrice', XmlNamespaceRAM);
            ChargeAmountElement := XmlElement.Create('ChargeAmount', XmlNamespaceRAM, FormatDecimalUnlimited(GetInvoiceLineUnitPrice(SalesInvoiceLine)));
            NetPriceProductTradePriceElement.Add(ChargeAmountElement);
            SpecifiedLineTradeAgreementElement.Add(NetPriceProductTradePriceElement);
            InvoiceLineElement.Add(SpecifiedLineTradeAgreementElement);

            SpecifiedLineTradeDeliveryElement := XmlElement.Create('SpecifiedLineTradeDelivery', XmlNamespaceRAM);
            BilledQuantityElement := CreateBilledQuantityElement(SalesInvoiceLine);
            SpecifiedLineTradeDeliveryElement.Add(BilledQuantityElement);
            InsertLineDeliveryNote(SpecifiedLineTradeDeliveryElement, SalesInvoiceLine);
            InvoiceLineElement.Add(SpecifiedLineTradeDeliveryElement);

            // Trade Settlement - VAT
            SpecifiedLineTradeSettlementElement := XmlElement.Create('SpecifiedLineTradeSettlement', XmlNamespaceRAM);
            SpecifiedLineTradeSettlementElement.Add(CreateLineTradeTaxElement(SalesInvoiceLine."Tax Category", SalesInvoiceLine."VAT Bus. Posting Group", SalesInvoiceLine."VAT Prod. Posting Group", SalesInvoiceLine."VAT %"));

            if SalesInvoiceLine."Shipment Date" <> 0D then
                InsertBillingSpecifiedPeriod(SpecifiedLineTradeSettlementElement, SalesInvoiceLine."Shipment Date", SalesInvoiceLine."Shipment Date");

            InsertLineDiscount(SpecifiedLineTradeSettlementElement, SalesInvoiceLine."Line Discount Amount", SalesInvoiceLine."Line Discount %");
            InsertLineLevelItemChargeAllowanceCharges(SpecifiedLineTradeSettlementElement, SalesInvoiceLine);

            SpecifiedTradeSettlementLineMonetarySummationElement := XmlElement.Create('SpecifiedTradeSettlementLineMonetarySummation', XmlNamespaceRAM);
            SpecifiedTradeSettlementLineMonetarySummationElement.Add(XmlElement.Create('LineTotalAmount', XmlNamespaceRAM, FormatDecimal(SalesInvoiceLine.Amount + SalesInvoiceLine."Inv. Discount Amount" + GetLineLevelItemChargeAmount(SalesInvoiceLine))));
            SpecifiedLineTradeSettlementElement.Add(SpecifiedTradeSettlementLineMonetarySummationElement);

            InvoiceLineElement.Add(SpecifiedLineTradeSettlementElement);
            OnBeforeAddInvoiceLineElement(InvoiceLineElement, SalesInvoiceLine, Currency, CurrencyCode, PricesIncVAT);
            SupplyChainTradeTransactionElement.Add(InvoiceLineElement);
        end;
    end;

    procedure InsertSupplyChainTradeTransaction(var RootXMLNode: XmlElement; SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var SalesCrMemoLine: Record "Sales Cr.Memo Line"; CurrencyCode: Code[10]; Currency: Record Currency; var LineAmount: Dictionary of [Decimal, Decimal]; var LineVATAmount: Dictionary of [Decimal, Decimal]; var LineAmounts: Dictionary of [Text, Decimal]; var LineDiscAmount: Dictionary of [Decimal, Decimal])
    var
        SupplyChainTradeTransactionElement: XmlElement;
    begin
        SupplyChainTradeTransactionElement := XmlElement.Create('SupplyChainTradeTransaction', XmlNamespaceRSM);
        if SalesCrMemoLine.FindSet() then
            repeat
                if not IsItemChargeAllowanceCharge(SalesCrMemoLine."Line No.") then
                    InsertCrMemoLine(SupplyChainTradeTransactionElement, SalesCrMemoLine, Currency, CurrencyCode, SalesCrMemoHeader."Prices Including VAT");
            until SalesCrMemoLine.Next() = 0;
        TempAdditionalSalesCrMemoLine.Reset();
        if TempAdditionalSalesCrMemoLine.FindSet() then
            repeat
                InsertCrMemoLine(SupplyChainTradeTransactionElement, TempAdditionalSalesCrMemoLine, Currency, CurrencyCode, SalesCrMemoHeader."Prices Including VAT");
            until TempAdditionalSalesCrMemoLine.Next() = 0;
        InsertApplicableHeaderTradeAgreement(SupplyChainTradeTransactionElement, SalesCrMemoHeader);
        InsertApplicableHeaderTradeDelivery(SupplyChainTradeTransactionElement, SalesCrMemoHeader);
        InsertApplicableHeaderTradeSettlement(SupplyChainTradeTransactionElement, SalesCrMemoHeader, SalesCrMemoLine, CurrencyCode, LineAmounts, LineDiscAmount);

        RootXMLNode.Add(SupplyChainTradeTransactionElement);
    end;

    local procedure InsertCrMemoLine(var SupplyChainTradeTransactionElement: XmlElement; var SalesCrMemoLine: Record "Sales Cr.Memo Line"; Currency: Record Currency; CurrencyCode: Code[10]; PricesIncVAT: Boolean)
    var
        LineRecordRef: RecordRef;
        CrMemoLineElement: XmlElement;
        AssociatedDocumentLineElement: XmlElement;
        SpecifiedLineTradeAgreementElement: XmlElement;
        NetPriceProductTradePriceElement: XmlElement;
        ChargeAmountElement: XmlElement;
        SpecifiedLineTradeDeliveryElement: XmlElement;
        BilledQuantityElement: XmlElement;
        SpecifiedLineTradeSettlementElement: XmlElement;
        SpecifiedTradeSettlementLineMonetarySummationElement: XmlElement;
        IsHandled: Boolean;
    begin
        IsHandled := false;
        OnBeforeInsertCrMemoLine(SupplyChainTradeTransactionElement, SalesCrMemoLine, Currency, CurrencyCode, PricesIncVAT, IsHandled);
        if not IsHandled then begin
            CrMemoLineElement := XmlElement.Create('IncludedSupplyChainTradeLineItem', XmlNamespaceRAM);
            if PricesIncVAT then
                ExcludeVAT(SalesCrMemoLine, Currency."Amount Rounding Precision");
            LineRecordRef.GetTable(SalesCrMemoLine);
            AssociatedDocumentLineElement := CreateAssociatedDocumentLineElement(LineRecordRef, SalesCrMemoLine."Line No.");
            CrMemoLineElement.Add(AssociatedDocumentLineElement);

            InsertTradeProduct(CrMemoLineElement, SalesCrMemoLine.Type = SalesCrMemoLine.Type::Item, SalesCrMemoLine."No.", SalesCrMemoLine."Item Reference No.",
                SalesCrMemoLine.Description, SalesCrMemoLine."Description 2", SalesCrMemoLine."Variant Code");

            SpecifiedLineTradeAgreementElement := XmlElement.Create('SpecifiedLineTradeAgreement', XmlNamespaceRAM);
            InsertLineReferences(SpecifiedLineTradeAgreementElement, SalesCrMemoLine."Order No.", SalesCrMemoLine."Order Line No.", SalesCrMemoLine."Blanket Order No.", SalesCrMemoLine."Blanket Order Line No.");
            NetPriceProductTradePriceElement := XmlElement.Create('NetPriceProductTradePrice', XmlNamespaceRAM);
            ChargeAmountElement := XmlElement.Create('ChargeAmount', XmlNamespaceRAM, FormatDecimalUnlimited(GetInvoiceLineUnitPrice(SalesCrMemoLine)));
            NetPriceProductTradePriceElement.Add(ChargeAmountElement);
            SpecifiedLineTradeAgreementElement.Add(NetPriceProductTradePriceElement);
            CrMemoLineElement.Add(SpecifiedLineTradeAgreementElement);

            SpecifiedLineTradeDeliveryElement := XmlElement.Create('SpecifiedLineTradeDelivery', XmlNamespaceRAM);
            BilledQuantityElement := CreateBilledQuantityElement(SalesCrMemoLine);
            SpecifiedLineTradeDeliveryElement.Add(BilledQuantityElement);
            CrMemoLineElement.Add(SpecifiedLineTradeDeliveryElement);

            // Trade Settlement - VAT
            SpecifiedLineTradeSettlementElement := XmlElement.Create('SpecifiedLineTradeSettlement', XmlNamespaceRAM);

            SpecifiedLineTradeSettlementElement.Add(CreateLineTradeTaxElement(SalesCrMemoLine."Tax Category", SalesCrMemoLine."VAT Bus. Posting Group", SalesCrMemoLine."VAT Prod. Posting Group", SalesCrMemoLine."VAT %"));

            if SalesCrMemoLine."Shipment Date" <> 0D then
                InsertBillingSpecifiedPeriod(SpecifiedLineTradeSettlementElement, SalesCrMemoLine."Shipment Date", SalesCrMemoLine."Shipment Date");

            InsertLineDiscount(SpecifiedLineTradeSettlementElement, SalesCrMemoLine."Line Discount Amount", SalesCrMemoLine."Line Discount %");
            InsertLineLevelItemChargeAllowanceCharges(SpecifiedLineTradeSettlementElement, SalesCrMemoLine);

            SpecifiedTradeSettlementLineMonetarySummationElement := XmlElement.Create('SpecifiedTradeSettlementLineMonetarySummation', XmlNamespaceRAM);
            SpecifiedTradeSettlementLineMonetarySummationElement.Add(XmlElement.Create('LineTotalAmount', XmlNamespaceRAM, FormatDecimal(SalesCrMemoLine.Amount + SalesCrMemoLine."Inv. Discount Amount" + GetLineLevelItemChargeAmount(SalesCrMemoLine))));
            SpecifiedLineTradeSettlementElement.Add(SpecifiedTradeSettlementLineMonetarySummationElement);

            CrMemoLineElement.Add(SpecifiedLineTradeSettlementElement);
            OnBeforeAddCrMemoLineElement(CrMemoLineElement, SalesCrMemoLine, Currency, CurrencyCode, PricesIncVAT);
            SupplyChainTradeTransactionElement.Add(CrMemoLineElement);
        end;
    end;

    local procedure InsertCreditorReference(var SettlementElement: XmlElement; PaymentMethodCode: Code[10]; CompanyBankAccountCode: Code[20])
    var
        CreditorNo: Code[35];
    begin
        if not DEPaymentMeansHelper.IsDirectDebit(DEPaymentMeansHelper.GetPaymentMeansCode(PaymentMethodCode)) then
            exit;
        CreditorNo := DEPaymentMeansHelper.GetCreditorNo(CompanyBankAccountCode);
        if CreditorNo = '' then
            exit;
        SettlementElement.Add(XmlElement.Create('CreditorReferenceID', XmlNamespaceRAM, CreditorNo));
    end;

    local procedure GetDirectDebitMandateID(PaymentMethodCode: Code[10]; DirectDebitMandateID: Code[35]): Code[35]
    begin
        if DEPaymentMeansHelper.IsDirectDebit(DEPaymentMeansHelper.GetPaymentMeansCode(PaymentMethodCode)) then
            exit(DirectDebitMandateID);
        exit('');
    end;

    local procedure InsertPaymentTerms(var RootXMLNode: XmlElement; PaymentTermsCode: Code[10]; DueDate: Date; DirectDebitMandateID: Code[35]; DocumentDate: Date; PmtDiscountDate: Date; PaymentDiscountPercent: Decimal; DiscountBasisAmount: Decimal; DiscountAmount: Decimal; DocumentCurrencyCode: Code[10])
    var
        PaymentTerms: Record "Payment Terms";
        PaymentTermsElement: XmlElement;
        PaymentTermsDescriptionElement: XmlElement;
        DueDateElement: XmlElement;
    begin
        PaymentTermsElement := XmlElement.Create('SpecifiedTradePaymentTerms', XmlNamespaceRAM);
        if PaymentTermsCode <> '' then
            if PaymentTerms.Get(PaymentTermsCode) and (PaymentTerms.Description <> '') then begin
                PaymentTermsDescriptionElement := XmlElement.Create('Description', XmlNamespaceRAM, PaymentTerms.Description);
                PaymentTermsElement.Add(PaymentTermsDescriptionElement);
            end;

        // CII R74: without a due date no DueDateDateTime, instead of the placeholder 17530101.
        if DueDate <> 0D then begin
            DueDateElement := XmlElement.Create('DueDateDateTime', XmlNamespaceRAM);
            DueDateElement.Add(XmlElement.Create('DateTimeString', XmlNamespaceUDT, XmlAttribute.Create('format', '102'), FormatDate(DueDate)));
            PaymentTermsElement.Add(DueDateElement);
        end;
        // BT-89 Mandate reference identifier. The CII TradePaymentTerms sequence puts DirectDebitMandateID
        // after DueDateDateTime, so it is added last.
        if DirectDebitMandateID <> '' then
            PaymentTermsElement.Add(XmlElement.Create('DirectDebitMandateID', XmlNamespaceRAM, DirectDebitMandateID));
        InsertPaymentDiscountTerms(PaymentTermsElement, DocumentDate, PmtDiscountDate, PaymentDiscountPercent, DiscountBasisAmount, DiscountAmount, DocumentCurrencyCode);

        if PaymentTermsElement.GetChildElements().Count() = 0 then
            exit;
        RootXMLNode.Add(PaymentTermsElement);
    end;

    /// <summary>
    /// Payment discount (Skonto), BT-X-282..BT-X-287, EXTENDED only. EN 16931 and XRechnung have no structured discount;
    /// the KoSIT text convention #SKONTO# for BT-20 is decision D-18 and not written here.
    /// </summary>
    local procedure InsertPaymentDiscountTerms(var PaymentTermsElement: XmlElement; DocumentDate: Date; PmtDiscountDate: Date; PaymentDiscountPercent: Decimal; DiscountBasisAmount: Decimal; DiscountAmount: Decimal; DocumentCurrencyCode: Code[10])
    var
        Currency: Record Currency;
        DiscountTermsElement: XmlElement;
        BasisDateElement: XmlElement;
    begin
        if CurrentZUGFeRDProfile <> CurrentZUGFeRDProfile::Extended then
            exit;
        if (PaymentDiscountPercent = 0) or (PmtDiscountDate = 0D) then
            exit;
        if DiscountAmount = 0 then begin
            Currency.Initialize(DocumentCurrencyCode);
            DiscountAmount := Round(DiscountBasisAmount * PaymentDiscountPercent / 100, Currency."Amount Rounding Precision");
        end;

        DiscountTermsElement := XmlElement.Create('ApplicableTradePaymentDiscountTerms', XmlNamespaceRAM);
        if DocumentDate <> 0D then begin
            // BT-X-282 basis date and BT-X-283 period: the discount applies until basis date + period.
            BasisDateElement := XmlElement.Create('BasisDateTime', XmlNamespaceRAM);
            BasisDateElement.Add(XmlElement.Create('DateTimeString', XmlNamespaceUDT, XmlAttribute.Create('format', '102'), FormatDate(DocumentDate)));
            DiscountTermsElement.Add(BasisDateElement);
            DiscountTermsElement.Add(XmlElement.Create('BasisPeriodMeasure', XmlNamespaceRAM, XmlAttribute.Create('unitCode', 'DAY'), Format(PmtDiscountDate - DocumentDate)));
        end;
        DiscountTermsElement.Add(XmlElement.Create('BasisAmount', XmlNamespaceRAM, FormatDecimal(DiscountBasisAmount)));
        DiscountTermsElement.Add(XmlElement.Create('CalculationPercent', XmlNamespaceRAM, FormatFiveDecimal(PaymentDiscountPercent)));
        DiscountTermsElement.Add(XmlElement.Create('ActualDiscountAmount', XmlNamespaceRAM, FormatDecimal(Abs(DiscountAmount))));
        PaymentTermsElement.Add(DiscountTermsElement);
    end;

    local procedure GetPaymentDiscountBasis(var LineAmounts: Dictionary of [Text, Decimal]; AmountKey: Text; AmountIncludingVATKey: Text): Decimal
    begin
        // The discount base follows G/L Setup "Pmt. Disc. Excl. VAT", as in the posting.
        if GeneralLedgerSetup."Pmt. Disc. Excl. VAT" then
            exit(LineAmounts.Get(AmountKey));
        exit(LineAmounts.Get(AmountIncludingVATKey));
    end;

    local procedure GetPaymentDiscountAmount(DocumentType: Enum "Gen. Journal Document Type"; DocumentNo: Code[20]; CustomerNo: Code[20]): Decimal
    var
        CustLedgerEntry: Record "Cust. Ledger Entry";
    begin
        // The discount as posted. Without a ledger entry (e.g. a plain print before posting) it is calculated from base and percentage.
        CustLedgerEntry.SetLoadFields("Original Pmt. Disc. Possible");
        CustLedgerEntry.SetRange("Document Type", DocumentType);
        CustLedgerEntry.SetRange("Document No.", DocumentNo);
        CustLedgerEntry.SetRange("Customer No.", CustomerNo);
        if CustLedgerEntry.FindFirst() then
            exit(CustLedgerEntry."Original Pmt. Disc. Possible");
        exit(0);
    end;

    /// <summary>
    /// BT-83 remittance information, EXTENDED only: the payment reference of the invoice, or the invoice number.
    /// </summary>
    local procedure InsertPaymentReference(var SettlementElement: XmlElement; PaymentReference: Code[50]; DocumentNo: Code[20])
    begin
        if CurrentZUGFeRDProfile <> CurrentZUGFeRDProfile::Extended then
            exit;
        if PaymentReference <> '' then
            SettlementElement.Add(XmlElement.Create('PaymentReference', XmlNamespaceRAM, PaymentReference))
        else
            AddElementIfNotEmpty(SettlementElement, 'PaymentReference', DocumentNo);
    end;

    local procedure InsertPaymentMethod(var RootXMLNode: XmlElement; PaymentMethodCode: Code[10]; CompanyBankAccountCode: Code[20]; DirectDebitMandateID: Code[35]; RecordVariant: Variant)
    var
        DataTypeManagement: Codeunit "Data Type Management";
        HeaderRecordRef: RecordRef;
        PaymentMethodElement, PaymentMethodTypeCodeElement : XmlElement;
        PaymentMeansCode: Code[3];
    begin
        PaymentMeansCode := DEPaymentMeansHelper.GetPaymentMeansCode(PaymentMethodCode);
        PaymentMethodElement := XmlElement.Create('SpecifiedTradeSettlementPaymentMeans', XmlNamespaceRAM);
        PaymentMethodTypeCodeElement := XmlElement.Create('TypeCode', XmlNamespaceRAM, PaymentMeansCode);
        PaymentMethodElement.Add(PaymentMethodTypeCodeElement);
        InsertPaymentMeansInformation(PaymentMethodElement, PaymentMethodCode);

        case PaymentMeansCode of
            '30', '58':
                InsertCreditTransferPayment(PaymentMethodElement, CompanyBankAccountCode);
            '49', '59':
                InsertDirectDebitPayment(PaymentMethodElement, DirectDebitMandateID);
        end;

        // The original source document, so that a subscriber sees the service header and not the sales header it was transferred to.
        // The passed header is only used when InsertSupplyChainTradeTransaction is called directly, without CreateXML.
        if SourceDocumentRecordRef.Number() <> 0 then
            OnInsertPaymentMethodOnBeforeAddToRoot(PaymentMethodElement, SourceDocumentRecordRef)
        else
            if DataTypeManagement.GetRecordRef(RecordVariant, HeaderRecordRef) then
                OnInsertPaymentMethodOnBeforeAddToRoot(PaymentMethodElement, HeaderRecordRef);
        RootXMLNode.Add(PaymentMethodElement);
    end;

    local procedure InsertDirectDebitPayment(var PaymentMethodElement: XmlElement; DirectDebitMandateID: Code[35])
    var
        SEPADirectDebitMandate: Record "SEPA Direct Debit Mandate";
        CustomerBankAccount: Record "Customer Bank Account";
    begin
        // BG-17 CREDIT TRANSFER (PayeePartyCreditorFinancialAccount, PayeeSpecifiedCreditorFinancialInstitution) must not
        // be sent for direct debit (BR-DE-25-b). The creditor is identified by BT-90 CreditorReferenceID instead.

        // BT-91 PayerPartyDebtorFinancialAccount = account being charged (customer, direct debit only), resolved via the mandate.
        if DirectDebitMandateID <> '' then begin
            SEPADirectDebitMandate.SetLoadFields(SEPADirectDebitMandate."Customer No.", SEPADirectDebitMandate."Customer Bank Account Code");
            if SEPADirectDebitMandate.Get(DirectDebitMandateID) then begin
                CustomerBankAccount.SetLoadFields(CustomerBankAccount.IBAN);
                if CustomerBankAccount.Get(SEPADirectDebitMandate."Customer No.", SEPADirectDebitMandate."Customer Bank Account Code") then
                    AddDebtorAccount(PaymentMethodElement, CustomerBankAccount.IBAN);
            end;
        end;
    end;

    /// <summary>
    /// BT-82 payment means text, EXTENDED only: the description of the payment method.
    /// </summary>
    local procedure InsertPaymentMeansInformation(var PaymentMethodElement: XmlElement; PaymentMethodCode: Code[10])
    var
        PaymentMethod: Record "Payment Method";
    begin
        if CurrentZUGFeRDProfile <> CurrentZUGFeRDProfile::Extended then
            exit;
        if PaymentMethodCode = '' then
            exit;
        PaymentMethod.SetLoadFields(Description);
        if PaymentMethod.Get(PaymentMethodCode) then
            AddElementIfNotEmpty(PaymentMethodElement, 'Information', PaymentMethod.Description);
    end;

    local procedure InsertCreditTransferPayment(var PaymentMethodElement: XmlElement; CompanyBankAccountCode: Code[20])
    var
        IBAN: Text[50];
        SWIFTCode: Code[20];
        AccountName: Text[100];
    begin
        // PayeePartyCreditorFinancialAccount = company bank account (payment recipient), for invoices and credit memos alike.
        // Customer bank data is intentionally not used as payee here, even though a customer IBAN/BIC could technically be read
        GetBankAccountPaymentDetails(CompanyBankAccountCode, IBAN, SWIFTCode, AccountName);
        // BT-85 account name, EXTENDED only.
        if CurrentZUGFeRDProfile <> CurrentZUGFeRDProfile::Extended then
            AccountName := '';
        AddCreditorAccount(PaymentMethodElement, IBAN, AccountName);
        AddCreditorFinancialInstitution(PaymentMethodElement, SWIFTCode);
    end;

    local procedure AddCreditorAccount(var PaymentMethodElement: XmlElement; IBAN: Text[50]; AccountName: Text[100])
    var
        PaymentMethodIBANElement: XmlElement;
    begin
        if IBAN = '' then
            exit;
        PaymentMethodIBANElement := XmlElement.Create('PayeePartyCreditorFinancialAccount', XmlNamespaceRAM);
        PaymentMethodIBANElement.Add(XmlElement.Create('IBANID', XmlNamespaceRAM, GetIBAN(IBAN)));
        AddElementIfNotEmpty(PaymentMethodIBANElement, 'AccountName', AccountName);
        PaymentMethodElement.Add(PaymentMethodIBANElement);
    end;

    local procedure AddDebtorAccount(var PaymentMethodElement: XmlElement; IBAN: Text[50])
    var
        DebtorAccountElement: XmlElement;
    begin
        if IBAN = '' then
            exit;
        DebtorAccountElement := XmlElement.Create('PayerPartyDebtorFinancialAccount', XmlNamespaceRAM);
        DebtorAccountElement.Add(XmlElement.Create('IBANID', XmlNamespaceRAM, GetIBAN(IBAN)));
        PaymentMethodElement.Add(DebtorAccountElement);
    end;

    local procedure AddCreditorFinancialInstitution(var PaymentMethodElement: XmlElement; SWIFTCode: Code[20])
    var
        PaymentMethodBICElement: XmlElement;
    begin
        if SWIFTCode = '' then
            exit;
        PaymentMethodBICElement := XmlElement.Create('PayeeSpecifiedCreditorFinancialInstitution', XmlNamespaceRAM);
        PaymentMethodBICElement.Add(XmlElement.Create('BICID', XmlNamespaceRAM, GetIBAN(SWIFTCode)));
        PaymentMethodElement.Add(PaymentMethodBICElement);
    end;

    local procedure InsertInvDiscountAllowanceCharge(var RootXMLNode: XmlElement; var SalesInvLine: Record "Sales Invoice Line"; var LineDiscAmount: Dictionary of [Decimal, Decimal]; var LineAmounts: Dictionary of [Text, Decimal])
    var
        InvDiscountAmount: Decimal;
    begin
        InvDiscountAmount := LineAmounts.Get(SalesInvLine.FieldName("Inv. Discount Amount"));
        if InvDiscountAmount = 0 then
            exit;
        if SalesInvLine.FindSet() then
            repeat
                InsertInvDiscountAllowanceChargeForLine(RootXMLNode, SalesInvLine, LineDiscAmount);
            until SalesInvLine.Next() = 0;
        TempAdditionalSalesInvLine.Reset();
        if TempAdditionalSalesInvLine.FindSet() then
            repeat
                InsertInvDiscountAllowanceChargeForLine(RootXMLNode, TempAdditionalSalesInvLine, LineDiscAmount);
            until TempAdditionalSalesInvLine.Next() = 0;
    end;

    local procedure InsertInvDiscountAllowanceChargeForLine(var RootXMLNode: XmlElement; SalesInvLine: Record "Sales Invoice Line"; var LineDiscAmount: Dictionary of [Decimal, Decimal])
    begin
        if not LineDiscAmount.ContainsKey(SalesInvLine."VAT %") then
            exit;
        InsertAllowanceCharge(
                   RootXMLNode, 'Document discount',
                   GetTaxCategoryID(SalesInvLine."Tax Category", SalesInvLine."VAT Bus. Posting Group", SalesInvLine."VAT Prod. Posting Group"),
                   LineDiscAmount.Get(SalesInvLine."VAT %"), SalesInvLine."VAT %", true);
        LineDiscAmount.Remove(SalesInvLine."VAT %");
    end;

    local procedure InsertInvDiscountAllowanceCharge(var RootXMLNode: XmlElement; var SalesCrMemoLine: Record "Sales Cr.Memo Line"; var LineDiscAmount: Dictionary of [Decimal, Decimal]; var LineAmounts: Dictionary of [Text, Decimal])
    var
        InvDiscountAmount: Decimal;
    begin
        InvDiscountAmount := LineAmounts.Get(SalesCrMemoLine.FieldName("Inv. Discount Amount"));
        if InvDiscountAmount = 0 then
            exit;
        if SalesCrMemoLine.FindSet() then
            repeat
                InsertInvDiscountAllowanceChargeForLine(RootXMLNode, SalesCrMemoLine, LineDiscAmount);
            until SalesCrMemoLine.Next() = 0;
        TempAdditionalSalesCrMemoLine.Reset();
        if TempAdditionalSalesCrMemoLine.FindSet() then
            repeat
                InsertInvDiscountAllowanceChargeForLine(RootXMLNode, TempAdditionalSalesCrMemoLine, LineDiscAmount);
            until TempAdditionalSalesCrMemoLine.Next() = 0;
    end;

    local procedure InsertInvDiscountAllowanceChargeForLine(var RootXMLNode: XmlElement; SalesCrMemoLine: Record "Sales Cr.Memo Line"; var LineDiscAmount: Dictionary of [Decimal, Decimal])
    begin
        if not LineDiscAmount.ContainsKey(SalesCrMemoLine."VAT %") then
            exit;
        InsertAllowanceCharge(
                   RootXMLNode, 'Document discount',
                   GetTaxCategoryID(SalesCrMemoLine."Tax Category", SalesCrMemoLine."VAT Bus. Posting Group", SalesCrMemoLine."VAT Prod. Posting Group"),
                   LineDiscAmount.Get(SalesCrMemoLine."VAT %"), SalesCrMemoLine."VAT %", true);
        LineDiscAmount.Remove(SalesCrMemoLine."VAT %");
    end;

    local procedure InsertAllowanceCharge(var RootXMLNode: XmlElement; AllowanceChargeReason: Text; TaxCategory: Text; Amount: Decimal; Percent: Decimal; InsertCategoryTax: Boolean)
    var
        AllowanceChargeElement: XmlElement;
    begin
        AllowanceChargeElement := XmlElement.Create('SpecifiedTradeAllowanceCharge', XmlNamespaceRAM);
        AllowanceChargeElement.Add(XmlElement.Create('ChargeIndicator', XmlNamespaceRAM, XmlElement.Create('Indicator', XmlNamespaceUDT, false)));
        AllowanceChargeElement.Add(XmlElement.Create('ActualAmount', XmlNamespaceRAM, FormatDecimal(Amount)));
        AllowanceChargeElement.Add(XmlElement.Create('Reason', XmlNamespaceRAM, AllowanceChargeReason));
        if InsertCategoryTax then
            InsertCategoryTradeTax(AllowanceChargeElement, TaxCategory, FormatFiveDecimal(Percent));
        RootXMLNode.Add(AllowanceChargeElement);
    end;

    local procedure CalculateLineAmounts(SalesInvoiceHeader: Record "Sales Invoice Header"; var SalesInvLine: Record "Sales Invoice Line"; Currency: Record Currency; var LineAmounts: Dictionary of [Text, Decimal])
    var
        TotalInvDiscountAmount: Decimal;
        TotalAmount: Decimal;
        TotalAmountIncludingVAT: Decimal;
    begin
        repeat
            if SalesInvoiceHeader."Prices Including VAT" then
                SalesInvLine."Inv. Discount Amount" := Round(SalesInvLine."Inv. Discount Amount" / (1 + SalesInvLine."VAT %" / 100), Currency."Amount Rounding Precision");
            TotalInvDiscountAmount += SalesInvLine."Inv. Discount Amount";
        until SalesInvLine.Next() = 0;

        SalesInvLine.CalcSums(Amount, "Amount Including VAT");
        TotalAmount := SalesInvLine.Amount;
        TotalAmountIncludingVAT := SalesInvLine."Amount Including VAT";

        TempAdditionalSalesInvLine.Reset();
        if TempAdditionalSalesInvLine.FindSet() then
            repeat
                TotalAmount += TempAdditionalSalesInvLine.Amount;
                TotalAmountIncludingVAT += TempAdditionalSalesInvLine."Amount Including VAT";
                if SalesInvoiceHeader."Prices Including VAT" then
                    TotalInvDiscountAmount += Round(TempAdditionalSalesInvLine."Inv. Discount Amount" / (1 + TempAdditionalSalesInvLine."VAT %" / 100), Currency."Amount Rounding Precision")
                else
                    TotalInvDiscountAmount += TempAdditionalSalesInvLine."Inv. Discount Amount";
            until TempAdditionalSalesInvLine.Next() = 0;

        if not LineAmounts.ContainsKey(SalesInvLine.FieldName(Amount)) then
            LineAmounts.Add(SalesInvLine.FieldName(Amount), TotalAmount);
        if not LineAmounts.ContainsKey(SalesInvLine.FieldName("Amount Including VAT")) then
            LineAmounts.Add(SalesInvLine.FieldName("Amount Including VAT"), TotalAmountIncludingVAT);
        if not LineAmounts.ContainsKey(SalesInvLine.FieldName("Inv. Discount Amount")) then
            LineAmounts.Add(SalesInvLine.FieldName("Inv. Discount Amount"), TotalInvDiscountAmount);
        OnAfterCalculateInvoiceLineAmounts(SalesInvoiceHeader, SalesInvLine, Currency, LineAmounts);
    end;

    local procedure CalculateLineAmounts(SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var SalesCrMemoLine: Record "Sales Cr.Memo Line"; Currency: Record Currency; var LineAmounts: Dictionary of [Text, Decimal])
    var
        TotalInvDiscountAmount: Decimal;
        TotalAmount: Decimal;
        TotalAmountIncludingVAT: Decimal;
    begin
        repeat
            if SalesCrMemoHeader."Prices Including VAT" then
                SalesCrMemoLine."Inv. Discount Amount" := Round(SalesCrMemoLine."Inv. Discount Amount" / (1 + SalesCrMemoLine."VAT %" / 100), Currency."Amount Rounding Precision");
            TotalInvDiscountAmount += SalesCrMemoLine."Inv. Discount Amount";
        until SalesCrMemoLine.Next() = 0;

        SalesCrMemoLine.CalcSums(Amount, "Amount Including VAT");
        TotalAmount := SalesCrMemoLine.Amount;
        TotalAmountIncludingVAT := SalesCrMemoLine."Amount Including VAT";

        TempAdditionalSalesCrMemoLine.Reset();
        if TempAdditionalSalesCrMemoLine.FindSet() then
            repeat
                TotalAmount += TempAdditionalSalesCrMemoLine.Amount;
                TotalAmountIncludingVAT += TempAdditionalSalesCrMemoLine."Amount Including VAT";
                if SalesCrMemoHeader."Prices Including VAT" then
                    TotalInvDiscountAmount += Round(TempAdditionalSalesCrMemoLine."Inv. Discount Amount" / (1 + TempAdditionalSalesCrMemoLine."VAT %" / 100), Currency."Amount Rounding Precision")
                else
                    TotalInvDiscountAmount += TempAdditionalSalesCrMemoLine."Inv. Discount Amount";
            until TempAdditionalSalesCrMemoLine.Next() = 0;

        if not LineAmounts.ContainsKey(SalesCrMemoLine.FieldName(Amount)) then
            LineAmounts.Add(SalesCrMemoLine.FieldName(Amount), TotalAmount);
        if not LineAmounts.ContainsKey(SalesCrMemoLine.FieldName("Amount Including VAT")) then
            LineAmounts.Add(SalesCrMemoLine.FieldName("Amount Including VAT"), TotalAmountIncludingVAT);
        if not LineAmounts.ContainsKey(SalesCrMemoLine.FieldName("Inv. Discount Amount")) then
            LineAmounts.Add(SalesCrMemoLine.FieldName("Inv. Discount Amount"), TotalInvDiscountAmount);
        OnAfterCalculateCrMemoLineAmounts(SalesCrMemoHeader, SalesCrMemoLine, Currency, LineAmounts);
    end;

    local procedure ExcludeVAT(var SalesInvLine: Record "Sales Invoice Line"; RoundingPrecision: Decimal)
    begin
        SalesInvLine."Line Discount Amount" := Round(SalesInvLine."Line Discount Amount" / (1 + SalesInvLine."VAT %" / 100), RoundingPrecision);
        SalesInvLine."Unit Price" := SalesInvLine."Unit Price" / (1 + SalesInvLine."VAT %" / 100);
        SalesInvLine."Inv. Discount Amount" := Round(SalesInvLine."Inv. Discount Amount" / (1 + SalesInvLine."VAT %" / 100), RoundingPrecision);
    end;

    local procedure ExcludeVAT(var SalesCrMemoLine: Record "Sales Cr.Memo Line"; RoundingPrecision: Decimal)
    begin
        SalesCrMemoLine."Line Discount Amount" := Round(SalesCrMemoLine."Line Discount Amount" / (1 + SalesCrMemoLine."VAT %" / 100), RoundingPrecision);
        SalesCrMemoLine."Unit Price" := SalesCrMemoLine."Unit Price" / (1 + SalesCrMemoLine."VAT %" / 100);
        SalesCrMemoLine."Inv. Discount Amount" := Round(SalesCrMemoLine."Inv. Discount Amount" / (1 + SalesCrMemoLine."VAT %" / 100), RoundingPrecision);
    end;

    local procedure InsertVATAmounts(var SalesInvLine: Record "Sales Invoice Line"; var LineVATAmount: Dictionary of [Decimal, Decimal]; var LineAmount: Dictionary of [Decimal, Decimal]; var LineDiscAmount: Dictionary of [Decimal, Decimal]; PricesIncVAT: Boolean; Currency: Record Currency)
    begin
        if SalesInvLine.FindSet() then
            repeat
                AddLineVATAmounts(SalesInvLine, LineVATAmount, LineAmount, LineDiscAmount, PricesIncVAT, Currency);
            until SalesInvLine.Next() = 0;
        TempAdditionalSalesInvLine.Reset();
        if TempAdditionalSalesInvLine.FindSet() then
            repeat
                AddLineVATAmounts(TempAdditionalSalesInvLine, LineVATAmount, LineAmount, LineDiscAmount, PricesIncVAT, Currency);
            until TempAdditionalSalesInvLine.Next() = 0;
    end;

    local procedure AddLineVATAmounts(SalesInvLine: Record "Sales Invoice Line"; var LineVATAmount: Dictionary of [Decimal, Decimal]; var LineAmount: Dictionary of [Decimal, Decimal]; var LineDiscAmount: Dictionary of [Decimal, Decimal]; PricesIncVAT: Boolean; Currency: Record Currency)
    begin
        AddAmountForVAT(SalesInvLine."VAT %", SalesInvLine."Amount Including VAT" - SalesInvLine.Amount, LineVATAmount);
        AddAmountForVAT(SalesInvLine."VAT %", SalesInvLine.Amount, LineAmount);
        if PricesIncVAT then
            ExcludeVAT(SalesInvLine, Currency."Amount Rounding Precision");
        AddAmountForVAT(SalesInvLine."VAT %", SalesInvLine."Inv. Discount Amount", LineDiscAmount);
    end;

    local procedure InsertVATAmounts(var SalesCrMemoLine: Record "Sales Cr.Memo Line"; var LineVATAmount: Dictionary of [Decimal, Decimal]; var LineAmount: Dictionary of [Decimal, Decimal]; var LineDiscAmount: Dictionary of [Decimal, Decimal]; PricesIncVAT: Boolean; Currency: Record Currency)
    begin
        if SalesCrMemoLine.FindSet() then
            repeat
                AddLineVATAmounts(SalesCrMemoLine, LineVATAmount, LineAmount, LineDiscAmount, PricesIncVAT, Currency);
            until SalesCrMemoLine.Next() = 0;
        TempAdditionalSalesCrMemoLine.Reset();
        if TempAdditionalSalesCrMemoLine.FindSet() then
            repeat
                AddLineVATAmounts(TempAdditionalSalesCrMemoLine, LineVATAmount, LineAmount, LineDiscAmount, PricesIncVAT, Currency);
            until TempAdditionalSalesCrMemoLine.Next() = 0;
    end;

    local procedure AddLineVATAmounts(SalesCrMemoLine: Record "Sales Cr.Memo Line"; var LineVATAmount: Dictionary of [Decimal, Decimal]; var LineAmount: Dictionary of [Decimal, Decimal]; var LineDiscAmount: Dictionary of [Decimal, Decimal]; PricesIncVAT: Boolean; Currency: Record Currency)
    begin
        AddAmountForVAT(SalesCrMemoLine."VAT %", SalesCrMemoLine."Amount Including VAT" - SalesCrMemoLine.Amount, LineVATAmount);
        AddAmountForVAT(SalesCrMemoLine."VAT %", SalesCrMemoLine.Amount, LineAmount);
        if PricesIncVAT then
            ExcludeVAT(SalesCrMemoLine, Currency."Amount Rounding Precision");
        AddAmountForVAT(SalesCrMemoLine."VAT %", SalesCrMemoLine."Inv. Discount Amount", LineDiscAmount);
    end;

    local procedure AddAmountForVAT(VATPercent: Decimal; NewAmount: Decimal; var TotalAmounts: Dictionary of [Decimal, Decimal])
    begin
        if not TotalAmounts.ContainsKey(VATPercent) then
            TotalAmounts.Add(VATPercent, NewAmount)
        else
            TotalAmounts.Set(VATPercent, TotalAmounts.Get(VATPercent) + NewAmount);
    end;

    local procedure GetBuyerReference(RecordVariant: Variant): Text
    var
        SalesInvoiceHeader: Record "Sales Invoice Header";
        DataTypeManagement: Codeunit "Data Type Management";
        HeaderRecordRef: RecordRef;
        BuyerReferenceFieldRef: FieldRef;
        CustomerNoFieldRef: FieldRef;
        YourReferenceFieldRef: FieldRef;
        BuyerReference: Text[100];
        BillToCustomerNo: Code[20];
        YourReference: Text[35];
    begin
        if not DataTypeManagement.GetRecordRef(RecordVariant, HeaderRecordRef) then
            exit('');
        if not (HeaderRecordRef.Number() in
            [Database::"Sales Invoice Header",
            Database::"Sales Cr.Memo Header",
            Database::"Service Invoice Header",
            Database::"Service Cr.Memo Header"])
        then
            exit('');
        BuyerReferenceFieldRef := HeaderRecordRef.Field(SalesInvoiceHeader.FieldNo("Buyer Reference"));
        CustomerNoFieldRef := HeaderRecordRef.Field(SalesInvoiceHeader.FieldNo("Bill-to Customer No."));
        YourReferenceFieldRef := HeaderRecordRef.Field(SalesInvoiceHeader.FieldNo("Your Reference"));
        BuyerReference := BuyerReferenceFieldRef.Value();
        BillToCustomerNo := CustomerNoFieldRef.Value();
        YourReference := YourReferenceFieldRef.Value();
        exit(EDocumentDEHelper.GetBuyerReferenceValue(BuyerReference, BillToCustomerNo, YourReference));
    end;

    local procedure GetIBAN(IBAN: Text[50]) IBANFormatted: Text[50]
    begin
        // Format IBAN to remove spaces and ensure it is in uppercase
        if IBAN = '' then
            exit('');
        IBANFormatted := UpperCase(DelChr(IBAN, '=', ' '));
        exit(CopyStr(IBANFormatted, 1, 50));
    end;

    local procedure GetSellerPostalAddr(RespCentercode: Code[10]; var StreetName: Text; var SupplierAdditionalStreetName: Text; var CityName: Text; var PostalZone: Text; var CountryRegionCode: Code[10])
    var
        RespCenter: Record "Responsibility Center";
    begin
        if RespCenter.Get(RespCentercode) then begin
            StreetName := RespCenter.Address;
            SupplierAdditionalStreetName := RespCenter."Address 2";
            CityName := RespCenter.City;
            PostalZone := RespCenter."Post Code";
            CountryRegionCode := GetCountryRegionCode(RespCenter."Country/Region Code");
            exit;
        end;
        StreetName := CompanyInformation.Address;
        SupplierAdditionalStreetName := CompanyInformation."Address 2";
        CityName := CompanyInformation.City;
        PostalZone := CompanyInformation."Post Code";
        CountryRegionCode := CompanyInformation."Country/Region Code";
    end;

    local procedure GetSellerCounty(RespCenterCode: Code[10]): Text
    var
        RespCenter: Record "Responsibility Center";
    begin
        RespCenter.SetLoadFields(County);
        if RespCenter.Get(RespCenterCode) then
            exit(RespCenter.County);
        exit(CompanyInformation.County);
    end;

    local procedure GetSellerContactInfo(SalesInvoiceHeader: Record "Sales Invoice Header"; var ContactName: Text; var PhoneNumber: Text; var EmailAddress: Text)
    begin
        if SetSellerContactFromSalesPerson(SalesInvoiceHeader."Salesperson Code", ContactName, PhoneNumber, EmailAddress) then
            exit;
        SetSellerContactFromCompanyInformation(ContactName, PhoneNumber, EmailAddress);
    end;

    local procedure GetSellerContactInfo(SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var ContactName: Text; var PhoneNumber: Text; var EmailAddress: Text)
    begin
        if SetSellerContactFromSalesPerson(SalesCrMemoHeader."Salesperson Code", ContactName, PhoneNumber, EmailAddress) then
            exit;
        SetSellerContactFromCompanyInformation(ContactName, PhoneNumber, EmailAddress);
    end;

    local procedure SetSellerContactFromSalesPerson(SalesPersonCode: Code[20]; var ContactName: Text; var PhoneNumber: Text; var EmailAddress: Text): Boolean
    var
        Salesperson: Record "Salesperson/Purchaser";
    begin
        if SalesPersonCode = '' then
            exit(false);
        Salesperson.SetLoadFields(Name, "Phone No.", "E-Mail");
        if not Salesperson.Get(SalesPersonCode) then
            exit(false);
        ContactName := Salesperson.Name;
        PhoneNumber := Salesperson."Phone No.";
        EmailAddress := Salesperson."E-Mail";
        exit(true);
    end;

    local procedure SetSellerContactFromCompanyInformation(var ContactName: Text; var PhoneNumber: Text; var EmailAddress: Text)
    begin
        ContactName := CompanyInformation."Contact Person";
        PhoneNumber := CompanyInformation."Phone No.";
        EmailAddress := CompanyInformation."E-Mail";
    end;

    #region ItemCharge
    local procedure ClassifyItemCharges(SalesInvoiceHeader: Record "Sales Invoice Header"; var SalesInvLine: Record "Sales Invoice Line")
    var
        ChargeSalesInvLine: Record "Sales Invoice Line";
        TargetSalesInvLine: Record "Sales Invoice Line";
        Structure: Enum "Item Charge E-Doc. Structure";
        ExportedLineExists: Boolean;
    begin
        ClearItemChargeClassification();

        ChargeSalesInvLine.CopyFilters(SalesInvLine);
        ChargeSalesInvLine.SetRange(Type, ChargeSalesInvLine.Type::"Charge (Item)");
        if not ChargeSalesInvLine.FindSet() then
            exit;

        ExportedLineExists := ExportedNonChargeLineExists(SalesInvLine);

        repeat
            Structure := EDocItemChargeMapping.GetItemChargeStructure(EDocumentService, SalesInvoiceHeader, ChargeSalesInvLine, TargetSalesInvLine);
            // The export skips lines that the classification cannot see, such as lines without a number, without a
            // quantity, or removed by a subscriber. Turning the item charges into allowances/charges would then leave
            // the document without any invoice line, which BR-16 does not allow.
            if not ExportedLineExists then
                Structure := Structure::"Line with Unit Code";
            // A line level allowance/charge needs an invoice line to live in. Without one it degrades to document level, so that the charge is never lost.
            if (Structure = Structure::"Line Allowance/Charge") and not IsExportedLine(SalesInvLine, TargetSalesInvLine."Line No.") then
                Structure := Structure::"Document Allowance/Charge";
            ItemChargeStructures.Add(ChargeSalesInvLine."Line No.", Structure.AsInteger());
            if Structure = Structure::"Line Allowance/Charge" then
                AddLineLevelItemCharge(TargetSalesInvLine."Line No.", ChargeSalesInvLine."Line No.", ChargeSalesInvLine.Amount);
        until ChargeSalesInvLine.Next() = 0;
    end;

    local procedure ClassifyItemCharges(SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var SalesCrMemoLine: Record "Sales Cr.Memo Line")
    var
        ChargeSalesCrMemoLine: Record "Sales Cr.Memo Line";
        TargetSalesCrMemoLine: Record "Sales Cr.Memo Line";
        Structure: Enum "Item Charge E-Doc. Structure";
        ExportedLineExists: Boolean;
    begin
        ClearItemChargeClassification();

        ChargeSalesCrMemoLine.CopyFilters(SalesCrMemoLine);
        ChargeSalesCrMemoLine.SetRange(Type, ChargeSalesCrMemoLine.Type::"Charge (Item)");
        if not ChargeSalesCrMemoLine.FindSet() then
            exit;

        ExportedLineExists := ExportedNonChargeLineExists(SalesCrMemoLine);

        repeat
            Structure := EDocItemChargeMapping.GetItemChargeStructure(EDocumentService, SalesCrMemoHeader, ChargeSalesCrMemoLine, TargetSalesCrMemoLine);
            // The export skips lines that the classification cannot see, such as lines without a number, without a
            // quantity, or removed by a subscriber. Turning the item charges into allowances/charges would then leave
            // the document without any credit memo line, which BR-16 does not allow.
            if not ExportedLineExists then
                Structure := Structure::"Line with Unit Code";
            // A line level allowance/charge needs a credit memo line to live in. Without one it degrades to document level, so that the charge is never lost.
            if (Structure = Structure::"Line Allowance/Charge") and not IsExportedLine(SalesCrMemoLine, TargetSalesCrMemoLine."Line No.") then
                Structure := Structure::"Document Allowance/Charge";
            ItemChargeStructures.Add(ChargeSalesCrMemoLine."Line No.", Structure.AsInteger());
            if Structure = Structure::"Line Allowance/Charge" then
                AddLineLevelItemCharge(TargetSalesCrMemoLine."Line No.", ChargeSalesCrMemoLine."Line No.", ChargeSalesCrMemoLine.Amount);
        until ChargeSalesCrMemoLine.Next() = 0;
    end;

    local procedure ClearItemChargeClassification()
    begin
        Clear(ItemChargeStructures);
        Clear(LineLevelItemChargeAmounts);
        Clear(LineLevelItemChargeLineNos);
    end;

    local procedure AddLineLevelItemCharge(TargetLineNo: Integer; ChargeLineNo: Integer; ChargeAmount: Decimal)
    var
        ChargeLineNos: List of [Integer];
    begin
        if LineLevelItemChargeAmounts.ContainsKey(TargetLineNo) then
            LineLevelItemChargeAmounts.Set(TargetLineNo, LineLevelItemChargeAmounts.Get(TargetLineNo) + ChargeAmount)
        else
            LineLevelItemChargeAmounts.Add(TargetLineNo, ChargeAmount);

        if not LineLevelItemChargeLineNos.Get(TargetLineNo, ChargeLineNos) then
            LineLevelItemChargeLineNos.Add(TargetLineNo, ChargeLineNos);
        ChargeLineNos.Add(ChargeLineNo);
    end;

    local procedure ExportedNonChargeLineExists(var SalesInvLine: Record "Sales Invoice Line"): Boolean
    var
        ExportedSalesInvLine: Record "Sales Invoice Line";
    begin
        ExportedSalesInvLine.CopyFilters(SalesInvLine);
        ExportedSalesInvLine.SetFilter(Type, '<>%1', ExportedSalesInvLine.Type::"Charge (Item)");
        exit(not ExportedSalesInvLine.IsEmpty());
    end;

    local procedure ExportedNonChargeLineExists(var SalesCrMemoLine: Record "Sales Cr.Memo Line"): Boolean
    var
        ExportedSalesCrMemoLine: Record "Sales Cr.Memo Line";
    begin
        ExportedSalesCrMemoLine.CopyFilters(SalesCrMemoLine);
        ExportedSalesCrMemoLine.SetFilter(Type, '<>%1', ExportedSalesCrMemoLine.Type::"Charge (Item)");
        exit(not ExportedSalesCrMemoLine.IsEmpty());
    end;

    local procedure IsExportedLine(var SalesInvLine: Record "Sales Invoice Line"; LineNo: Integer): Boolean
    var
        ExportedSalesInvLine: Record "Sales Invoice Line";
    begin
        if LineNo = 0 then
            exit(false);
        ExportedSalesInvLine.CopyFilters(SalesInvLine);
        ExportedSalesInvLine.SetRange("Line No.", LineNo);
        exit(not ExportedSalesInvLine.IsEmpty());
    end;

    local procedure IsExportedLine(var SalesCrMemoLine: Record "Sales Cr.Memo Line"; LineNo: Integer): Boolean
    var
        ExportedSalesCrMemoLine: Record "Sales Cr.Memo Line";
    begin
        if LineNo = 0 then
            exit(false);
        ExportedSalesCrMemoLine.CopyFilters(SalesCrMemoLine);
        ExportedSalesCrMemoLine.SetRange("Line No.", LineNo);
        exit(not ExportedSalesCrMemoLine.IsEmpty());
    end;

    local procedure InsertItemChargeAllowanceCharges(var SettlementElement: XmlElement; var SalesInvLine: Record "Sales Invoice Line")
    var
        ChargeSalesInvLine: Record "Sales Invoice Line";
    begin
        if ItemChargeStructures.Count() = 0 then
            exit;

        ChargeSalesInvLine.CopyFilters(SalesInvLine);
        ChargeSalesInvLine.SetRange(Type, ChargeSalesInvLine.Type::"Charge (Item)");
        if not ChargeSalesInvLine.FindSet() then
            exit;

        repeat
            if IsDocumentLevelItemCharge(ChargeSalesInvLine."Line No.") then
                InsertItemChargeAllowanceCharge(SettlementElement, ChargeSalesInvLine, true);
        until ChargeSalesInvLine.Next() = 0;
    end;

    local procedure InsertItemChargeAllowanceCharges(var SettlementElement: XmlElement; var SalesCrMemoLine: Record "Sales Cr.Memo Line")
    var
        ChargeSalesCrMemoLine: Record "Sales Cr.Memo Line";
    begin
        if ItemChargeStructures.Count() = 0 then
            exit;

        ChargeSalesCrMemoLine.CopyFilters(SalesCrMemoLine);
        ChargeSalesCrMemoLine.SetRange(Type, ChargeSalesCrMemoLine.Type::"Charge (Item)");
        if not ChargeSalesCrMemoLine.FindSet() then
            exit;

        repeat
            if IsDocumentLevelItemCharge(ChargeSalesCrMemoLine."Line No.") then
                InsertItemChargeAllowanceCharge(SettlementElement, ChargeSalesCrMemoLine, true);
        until ChargeSalesCrMemoLine.Next() = 0;
    end;

    local procedure InsertLineLevelItemChargeAllowanceCharges(var LineTradeSettlementElement: XmlElement; var SalesInvLine: Record "Sales Invoice Line")
    var
        ChargeSalesInvLine: Record "Sales Invoice Line";
        ChargeLineNos: List of [Integer];
        ChargeLineNo: Integer;
    begin
        if not LineLevelItemChargeLineNos.Get(SalesInvLine."Line No.", ChargeLineNos) then
            exit;

        foreach ChargeLineNo in ChargeLineNos do begin
            ChargeSalesInvLine.Get(SalesInvLine."Document No.", ChargeLineNo);
            InsertItemChargeAllowanceCharge(LineTradeSettlementElement, ChargeSalesInvLine, false);
        end;
    end;

    local procedure InsertLineLevelItemChargeAllowanceCharges(var LineTradeSettlementElement: XmlElement; var SalesCrMemoLine: Record "Sales Cr.Memo Line")
    var
        ChargeSalesCrMemoLine: Record "Sales Cr.Memo Line";
        ChargeLineNos: List of [Integer];
        ChargeLineNo: Integer;
    begin
        if not LineLevelItemChargeLineNos.Get(SalesCrMemoLine."Line No.", ChargeLineNos) then
            exit;

        foreach ChargeLineNo in ChargeLineNos do begin
            ChargeSalesCrMemoLine.Get(SalesCrMemoLine."Document No.", ChargeLineNo);
            InsertItemChargeAllowanceCharge(LineTradeSettlementElement, ChargeSalesCrMemoLine, false);
        end;
    end;

    local procedure InsertItemChargeAllowanceCharge(var ParentXMLNode: XmlElement; ChargeSalesInvLine: Record "Sales Invoice Line"; InsertCategoryTax: Boolean)
    var
        AllowanceChargeElement: XmlElement;
        ReasonCode: Code[10];
        ReasonText: Text[100];
    begin
        GetItemChargeReason(ChargeSalesInvLine, ReasonCode, ReasonText);
        AllowanceChargeElement := XmlElement.Create('SpecifiedTradeAllowanceCharge', XmlNamespaceRAM);
        AllowanceChargeElement.Add(
            XmlElement.Create('ChargeIndicator', XmlNamespaceRAM,
                XmlElement.Create('Indicator', XmlNamespaceUDT, EDocumentDEHelper.IsCharge(ChargeSalesInvLine.Amount))));
        AllowanceChargeElement.Add(XmlElement.Create('ActualAmount', XmlNamespaceRAM, FormatDecimal(EDocumentDEHelper.GetReportedAmount(ChargeSalesInvLine.Amount))));
        if ReasonCode <> '' then
            AllowanceChargeElement.Add(XmlElement.Create('ReasonCode', XmlNamespaceRAM, ReasonCode));
        AllowanceChargeElement.Add(XmlElement.Create('Reason', XmlNamespaceRAM, ReasonText));
        if InsertCategoryTax then
            InsertCategoryTradeTax(
                AllowanceChargeElement,
                GetTaxCategoryID(ChargeSalesInvLine."Tax Category", ChargeSalesInvLine."VAT Bus. Posting Group", ChargeSalesInvLine."VAT Prod. Posting Group"),
                FormatFiveDecimal(ChargeSalesInvLine."VAT %"));
        ParentXMLNode.Add(AllowanceChargeElement);
    end;

    local procedure InsertItemChargeAllowanceCharge(var ParentXMLNode: XmlElement; ChargeSalesCrMemoLine: Record "Sales Cr.Memo Line"; InsertCategoryTax: Boolean)
    var
        AllowanceChargeElement: XmlElement;
        ReasonCode: Code[10];
        ReasonText: Text[100];
    begin
        GetItemChargeReason(ChargeSalesCrMemoLine, ReasonCode, ReasonText);
        AllowanceChargeElement := XmlElement.Create('SpecifiedTradeAllowanceCharge', XmlNamespaceRAM);
        AllowanceChargeElement.Add(
            XmlElement.Create('ChargeIndicator', XmlNamespaceRAM,
                XmlElement.Create('Indicator', XmlNamespaceUDT, EDocumentDEHelper.IsCharge(ChargeSalesCrMemoLine.Amount))));
        AllowanceChargeElement.Add(XmlElement.Create('ActualAmount', XmlNamespaceRAM, FormatDecimal(EDocumentDEHelper.GetReportedAmount(ChargeSalesCrMemoLine.Amount))));
        if ReasonCode <> '' then
            AllowanceChargeElement.Add(XmlElement.Create('ReasonCode', XmlNamespaceRAM, ReasonCode));
        AllowanceChargeElement.Add(XmlElement.Create('Reason', XmlNamespaceRAM, ReasonText));
        if InsertCategoryTax then
            InsertCategoryTradeTax(
                AllowanceChargeElement,
                GetTaxCategoryID(ChargeSalesCrMemoLine."Tax Category", ChargeSalesCrMemoLine."VAT Bus. Posting Group", ChargeSalesCrMemoLine."VAT Prod. Posting Group"),
                FormatFiveDecimal(ChargeSalesCrMemoLine."VAT %"));
        ParentXMLNode.Add(AllowanceChargeElement);
    end;

    local procedure GetItemChargeReason(ChargeSalesInvLine: Record "Sales Invoice Line"; var ReasonCode: Code[10]; var ReasonText: Text[100])
    begin
        EDocItemChargeMapping.GetItemChargeReason(ChargeSalesInvLine."No.", ReasonCode, ReasonText);
        if ReasonText = '' then
            ReasonText := ChargeSalesInvLine.Description;
        if (ReasonText = '') and (ReasonCode = '') then
            ReasonText := ChargeSalesInvLine."No.";
    end;

    local procedure GetItemChargeReason(ChargeSalesCrMemoLine: Record "Sales Cr.Memo Line"; var ReasonCode: Code[10]; var ReasonText: Text[100])
    begin
        EDocItemChargeMapping.GetItemChargeReason(ChargeSalesCrMemoLine."No.", ReasonCode, ReasonText);
        if ReasonText = '' then
            ReasonText := ChargeSalesCrMemoLine.Description;
        if (ReasonText = '') and (ReasonCode = '') then
            ReasonText := ChargeSalesCrMemoLine."No.";
    end;

    local procedure GetDocumentLevelItemChargeTotals(var SalesInvLine: Record "Sales Invoice Line"; var TotalChargeAmount: Decimal; var TotalAllowanceAmount: Decimal)
    var
        ChargeSalesInvLine: Record "Sales Invoice Line";
    begin
        TotalChargeAmount := 0;
        TotalAllowanceAmount := 0;
        if ItemChargeStructures.Count() = 0 then
            exit;

        ChargeSalesInvLine.CopyFilters(SalesInvLine);
        ChargeSalesInvLine.SetRange(Type, ChargeSalesInvLine.Type::"Charge (Item)");
        if not ChargeSalesInvLine.FindSet() then
            exit;

        repeat
            if IsDocumentLevelItemCharge(ChargeSalesInvLine."Line No.") then
                if EDocumentDEHelper.IsCharge(ChargeSalesInvLine.Amount) then
                    TotalChargeAmount += EDocumentDEHelper.GetReportedAmount(ChargeSalesInvLine.Amount)
                else
                    TotalAllowanceAmount += EDocumentDEHelper.GetReportedAmount(ChargeSalesInvLine.Amount);
        until ChargeSalesInvLine.Next() = 0;
    end;

    local procedure GetDocumentLevelItemChargeTotals(var SalesCrMemoLine: Record "Sales Cr.Memo Line"; var TotalChargeAmount: Decimal; var TotalAllowanceAmount: Decimal)
    var
        ChargeSalesCrMemoLine: Record "Sales Cr.Memo Line";
    begin
        TotalChargeAmount := 0;
        TotalAllowanceAmount := 0;
        if ItemChargeStructures.Count() = 0 then
            exit;

        ChargeSalesCrMemoLine.CopyFilters(SalesCrMemoLine);
        ChargeSalesCrMemoLine.SetRange(Type, ChargeSalesCrMemoLine.Type::"Charge (Item)");
        if not ChargeSalesCrMemoLine.FindSet() then
            exit;

        repeat
            if IsDocumentLevelItemCharge(ChargeSalesCrMemoLine."Line No.") then
                if EDocumentDEHelper.IsCharge(ChargeSalesCrMemoLine.Amount) then
                    TotalChargeAmount += EDocumentDEHelper.GetReportedAmount(ChargeSalesCrMemoLine.Amount)
                else
                    TotalAllowanceAmount += EDocumentDEHelper.GetReportedAmount(ChargeSalesCrMemoLine.Amount);
        until ChargeSalesCrMemoLine.Next() = 0;
    end;

    local procedure GetLineLevelItemChargeAmount(var SalesInvLine: Record "Sales Invoice Line") TotalChargeAmount: Decimal
    begin
        if not LineLevelItemChargeAmounts.Get(SalesInvLine."Line No.", TotalChargeAmount) then
            exit(0);
    end;

    local procedure GetLineLevelItemChargeAmount(var SalesCrMemoLine: Record "Sales Cr.Memo Line") TotalChargeAmount: Decimal
    begin
        if not LineLevelItemChargeAmounts.Get(SalesCrMemoLine."Line No.", TotalChargeAmount) then
            exit(0);
    end;

    local procedure CreateBilledQuantityElement(SalesInvoiceLine: Record "Sales Invoice Line") BilledQuantityElement: XmlElement
    begin
        if IsItemChargeInvoiceLine(SalesInvoiceLine."Line No.") then begin
            BilledQuantityElement := XmlElement.Create('BilledQuantity', XmlNamespaceRAM, FormatDecimalUnlimited(EDocItemChargeMapping.GetFallbackQuantity(GetLineNetAmount(SalesInvoiceLine))));
            BilledQuantityElement.SetAttribute('unitCode', EDocItemChargeMapping.GetFallbackUnitOfMeasureCode(SalesInvoiceLine."No."));
            exit;
        end;

        BilledQuantityElement := XmlElement.Create('BilledQuantity', XmlNamespaceRAM, FormatDecimalUnlimited(SalesInvoiceLine.Quantity));
        BilledQuantityElement.SetAttribute('unitCode', GetUoMCode(SalesInvoiceLine."Unit of Measure Code"));
    end;

    local procedure CreateBilledQuantityElement(SalesCrMemoLine: Record "Sales Cr.Memo Line") BilledQuantityElement: XmlElement
    begin
        if IsItemChargeInvoiceLine(SalesCrMemoLine."Line No.") then begin
            BilledQuantityElement := XmlElement.Create('BilledQuantity', XmlNamespaceRAM, FormatDecimalUnlimited(EDocItemChargeMapping.GetFallbackQuantity(GetLineNetAmount(SalesCrMemoLine))));
            BilledQuantityElement.SetAttribute('unitCode', EDocItemChargeMapping.GetFallbackUnitOfMeasureCode(SalesCrMemoLine."No."));
            exit;
        end;

        BilledQuantityElement := XmlElement.Create('BilledQuantity', XmlNamespaceRAM, FormatDecimalUnlimited(SalesCrMemoLine.Quantity));
        BilledQuantityElement.SetAttribute('unitCode', GetUoMCode(SalesCrMemoLine."Unit of Measure Code"));
    end;

    local procedure GetInvoiceLineUnitPrice(SalesInvoiceLine: Record "Sales Invoice Line"): Decimal
    begin
        if not IsItemChargeInvoiceLine(SalesInvoiceLine."Line No.") then
            exit(SalesInvoiceLine."Unit Price");

        exit(EDocItemChargeMapping.GetFallbackUnitPrice(GetLineNetAmount(SalesInvoiceLine)));
    end;

    local procedure GetInvoiceLineUnitPrice(SalesCrMemoLine: Record "Sales Cr.Memo Line"): Decimal
    begin
        if not IsItemChargeInvoiceLine(SalesCrMemoLine."Line No.") then
            exit(SalesCrMemoLine."Unit Price");

        exit(EDocItemChargeMapping.GetFallbackUnitPrice(GetLineNetAmount(SalesCrMemoLine)));
    end;

    local procedure GetLineNetAmount(SalesInvoiceLine: Record "Sales Invoice Line"): Decimal
    begin
        exit(SalesInvoiceLine.Amount + SalesInvoiceLine."Inv. Discount Amount");
    end;

    local procedure GetLineNetAmount(SalesCrMemoLine: Record "Sales Cr.Memo Line"): Decimal
    begin
        exit(SalesCrMemoLine.Amount + SalesCrMemoLine."Inv. Discount Amount");
    end;

    local procedure IsItemChargeAllowanceCharge(LineNo: Integer): Boolean
    var
        Structure: Enum "Item Charge E-Doc. Structure";
    begin
        if not TryGetItemChargeStructure(LineNo, Structure) then
            exit(false);
        exit(Structure <> Structure::"Line with Unit Code");
    end;

    local procedure IsItemChargeInvoiceLine(LineNo: Integer): Boolean
    var
        Structure: Enum "Item Charge E-Doc. Structure";
    begin
        if not TryGetItemChargeStructure(LineNo, Structure) then
            exit(false);
        exit(Structure = Structure::"Line with Unit Code");
    end;

    local procedure IsDocumentLevelItemCharge(LineNo: Integer): Boolean
    var
        Structure: Enum "Item Charge E-Doc. Structure";
    begin
        if not TryGetItemChargeStructure(LineNo, Structure) then
            exit(false);
        exit(Structure = Structure::"Document Allowance/Charge");
    end;

    local procedure TryGetItemChargeStructure(LineNo: Integer; var Structure: Enum "Item Charge E-Doc. Structure"): Boolean
    begin
        if not ItemChargeStructures.ContainsKey(LineNo) then
            exit(false);
        Structure := Enum::"Item Charge E-Doc. Structure".FromInteger(ItemChargeStructures.Get(LineNo));
        exit(true);
    end;
    #endregion

    #region Profile
    /// <summary>
    /// Resolves the ZUGFeRD profile for a document. The export and any check that depends on the profile use this
    /// procedure, so that both always see the same profile.
    /// </summary>
    /// <param name="EDocService">The E-Document Service that triggered the export. A blank service gives the XRechnung profile.</param>
    /// <param name="HeaderRecordRef">The posted source document: sales or service invoice or credit memo.</param>
    /// <returns>The profile to export.</returns>
    procedure ResolveZUGFeRDProfile(EDocService: Record "E-Document Service"; HeaderRecordRef: RecordRef) ResolvedProfile: Enum "ZUGFeRD Profile"
    begin
        ResolvedProfile := EDocService."ZUGFeRD Profile";
        OnAfterResolveZUGFeRDProfile(EDocService, HeaderRecordRef, ResolvedProfile);
    end;

    local procedure GetGuidelineID(): Text
    begin
        if CurrentZUGFeRDProfile = CurrentZUGFeRDProfile::Extended then
            exit(ExtendedGuidelineIDTok);
        exit(XRechnungGuidelineIDTok);
    end;

    local procedure GetXMLAttachmentName(): Text
    begin
        // BR-HYBRID: the EN 16931 based profiles embed factur-x.xml, the XRechnung profile embeds xrechnung.xml.
        if CurrentZUGFeRDProfile = CurrentZUGFeRDProfile::Extended then
            exit(FacturXAttachmentNameTok);
        exit(XRechnungAttachmentNameTok);
    end;

    local procedure GetHeaderRecordRef(RecordVariant: Variant; var HeaderRecordRef: RecordRef)
    var
        DataTypeManagement: Codeunit "Data Type Management";
    begin
        // The original source document, so that a subscriber sees the service header and not the sales header it was transferred to.
        // The passed header is only used when InsertSupplyChainTradeTransaction is called directly, without CreateXML.
        if SourceDocumentRecordRef.Number() <> 0 then begin
            HeaderRecordRef := SourceDocumentRecordRef;
            exit;
        end;
        DataTypeManagement.GetRecordRef(RecordVariant, HeaderRecordRef);
    end;

    local procedure InsertHeaderNotes(var ExchangedDocumentElement: XmlElement; RecordVariant: Variant)
    var
        TempZUGFeRDNote: Record "ZUGFeRD Note" temporary;
        HeaderRecordRef: RecordRef;
        IncludedNoteElement: XmlElement;
    begin
        GetHeaderRecordRef(RecordVariant, HeaderRecordRef);
        AddWorkDescriptionNote(HeaderRecordRef, TempZUGFeRDNote);
        OnCollectHeaderNotes(HeaderRecordRef, CurrentZUGFeRDProfile, TempZUGFeRDNote);
        TempZUGFeRDNote.Reset();
        if not TempZUGFeRDNote.FindSet() then
            exit;
        repeat
            // CII R74: no empty elements.
            if TempZUGFeRDNote.Content <> '' then begin
                IncludedNoteElement := XmlElement.Create('IncludedNote', XmlNamespaceRAM);
                IncludedNoteElement.Add(XmlElement.Create('Content', XmlNamespaceRAM, TempZUGFeRDNote.Content));
                if TempZUGFeRDNote."Subject Code" <> '' then
                    IncludedNoteElement.Add(XmlElement.Create('SubjectCode', XmlNamespaceRAM, TempZUGFeRDNote."Subject Code"));
                ExchangedDocumentElement.Add(IncludedNoteElement);
            end;
        until TempZUGFeRDNote.Next() = 0;
    end;

    local procedure CreateAssociatedDocumentLineElement(LineRecordRef: RecordRef; LineNo: Integer) AssociatedDocumentLineElement: XmlElement
    var
        LineSubtype: Enum "ZUGFeRD Line Subtype";
        LineID: Text;
        LineNote: Text;
        ParentLineID: Text;
    begin
        LineID := Format(LineNo);
        OnInsertLineOnBeforeCreateAssociatedDocumentLine(LineRecordRef, CurrentZUGFeRDProfile, LineID, ParentLineID, LineSubtype);
        if LineID = '' then
            LineID := Format(LineNo);

        AssociatedDocumentLineElement := XmlElement.Create('AssociatedDocumentLineDocument', XmlNamespaceRAM);
        AssociatedDocumentLineElement.Add(XmlElement.Create('LineID', XmlNamespaceRAM, LineID));
        // Sub-lines (BT-X-304 ParentLineID, BT-X-8 LineStatusReasonCode) exist only in the EXTENDED profile.
        if CurrentZUGFeRDProfile <> CurrentZUGFeRDProfile::Extended then
            exit;
        // BR-FXEXT-06: a line with a parent line must have a sub-line type.
        if (ParentLineID <> '') and (LineSubtype = LineSubtype::" ") then
            LineSubtype := LineSubtype::Detail;
        if ParentLineID <> '' then
            AssociatedDocumentLineElement.Add(XmlElement.Create('ParentLineID', XmlNamespaceRAM, ParentLineID));
        if LineSubtype <> LineSubtype::" " then
            AssociatedDocumentLineElement.Add(XmlElement.Create('LineStatusReasonCode', XmlNamespaceRAM, GetLineSubtypeCode(LineSubtype)));
        // BT-127 line note, after the sub-line data.
        LineNote := GetAttachedLineText(LineRecordRef);
        if LineNote <> '' then
            AssociatedDocumentLineElement.Add(XmlElement.Create('IncludedNote', XmlNamespaceRAM, XmlElement.Create('Content', XmlNamespaceRAM, LineNote)));
    end;

    local procedure GetLineSubtypeCode(LineSubtype: Enum "ZUGFeRD Line Subtype"): Text
    begin
        case LineSubtype of
            LineSubtype::Detail:
                exit(LineSubtypeDetailTok);
            LineSubtype::Group:
                exit(LineSubtypeGroupTok);
            LineSubtype::Information:
                exit(LineSubtypeInformationTok);
        end;
        exit('');
    end;

    local procedure CollectAdditionalLines(var SalesInvLine: Record "Sales Invoice Line")
    var
        ExportedLineNos: List of [Integer];
    begin
        OnCollectAdditionalInvoiceLines(SourceDocumentRecordRef, CurrentZUGFeRDProfile, TempAdditionalSalesInvLine);
        TempAdditionalSalesInvLine.Reset();
        if TempAdditionalSalesInvLine.IsEmpty() then
            exit;

        // The line number identifies a line in the item charge classification and is the default line ID (BT-126),
        // so an additional line must not reuse the number of a document line.
        if SalesInvLine.FindSet() then
            repeat
                ExportedLineNos.Add(SalesInvLine."Line No.");
            until SalesInvLine.Next() = 0;
        if SalesInvLine.FindSet() then;
        TempAdditionalSalesInvLine.FindSet();
        repeat
            if ExportedLineNos.Contains(TempAdditionalSalesInvLine."Line No.") then
                Error(AdditionalLineNoExistsErr, TempAdditionalSalesInvLine."Line No.");
        until TempAdditionalSalesInvLine.Next() = 0;
    end;

    local procedure CollectAdditionalLines(var SalesCrMemoLine: Record "Sales Cr.Memo Line")
    var
        ExportedLineNos: List of [Integer];
    begin
        OnCollectAdditionalCrMemoLines(SourceDocumentRecordRef, CurrentZUGFeRDProfile, TempAdditionalSalesCrMemoLine);
        TempAdditionalSalesCrMemoLine.Reset();
        if TempAdditionalSalesCrMemoLine.IsEmpty() then
            exit;

        // The line number identifies a line in the item charge classification and is the default line ID (BT-126),
        // so an additional line must not reuse the number of a document line.
        if SalesCrMemoLine.FindSet() then
            repeat
                ExportedLineNos.Add(SalesCrMemoLine."Line No.");
            until SalesCrMemoLine.Next() = 0;
        if SalesCrMemoLine.FindSet() then;
        TempAdditionalSalesCrMemoLine.FindSet();
        repeat
            if ExportedLineNos.Contains(TempAdditionalSalesCrMemoLine."Line No.") then
                Error(AdditionalLineNoExistsErr, TempAdditionalSalesCrMemoLine."Line No.");
        until TempAdditionalSalesCrMemoLine.Next() = 0;
    end;

    local procedure ClearAdditionalLines()
    begin
        TempAdditionalSalesInvLine.Reset();
        TempAdditionalSalesInvLine.DeleteAll();
        TempAdditionalSalesCrMemoLine.Reset();
        TempAdditionalSalesCrMemoLine.DeleteAll();
    end;

    local procedure GetPrepaidAmount(RecordVariant: Variant) PrepaidAmount: Decimal
    var
        HeaderRecordRef: RecordRef;
    begin
        GetHeaderRecordRef(RecordVariant, HeaderRecordRef);
        OnInsertMonetarySummationOnBeforeCreateElement(HeaderRecordRef, CurrentZUGFeRDProfile, PrepaidAmount);
    end;

    local procedure RaiseOnAfterInsertApplicableHeaderTradeDelivery(var DeliveryElement: XmlElement; RecordVariant: Variant)
    var
        HeaderRecordRef: RecordRef;
    begin
        GetHeaderRecordRef(RecordVariant, HeaderRecordRef);
        OnAfterInsertApplicableHeaderTradeDelivery(DeliveryElement, HeaderRecordRef, CurrentZUGFeRDProfile);
    end;
    #endregion

    #region Additional data
    // FP4, FP5, FP8, FP9 of the EXTENDED prototype. The data is written for the EXTENDED profile only, so that the
    // XRechnung output stays unchanged (decision D-17: which EN terms also go into XRechnung is open).

    local procedure IsExtendedProfile(): Boolean
    begin
        exit(CurrentZUGFeRDProfile = CurrentZUGFeRDProfile::Extended);
    end;

    /// <summary>
    /// BG-31 item information. EXTENDED adds BT-X-533 ModelID (variant code), BT-154 Description (Description 2),
    /// BT-X-536 ModelName (variant description) and, for items, the item data (attributes, tariff no., origin, manufacturer).
    /// </summary>
    local procedure InsertTradeProduct(var LineElement: XmlElement; IsItem: Boolean; No: Code[20]; ItemReferenceNo: Code[50]; Description: Text; Description2: Text; VariantCode: Code[10])
    var
        ItemVariant: Record "Item Variant";
        ProductElement: XmlElement;
        HasVariant: Boolean;
    begin
        ProductElement := XmlElement.Create('SpecifiedTradeProduct', XmlNamespaceRAM);
        if IsItem then
            InsertGlobalID(ProductElement, No);
        if No <> '' then
            ProductElement.Add(XmlElement.Create('SellerAssignedID', XmlNamespaceRAM, No));
        if ItemReferenceNo <> '' then
            ProductElement.Add(XmlElement.Create('BuyerAssignedID', XmlNamespaceRAM, ItemReferenceNo));
        if IsExtendedProfile() and IsItem and (VariantCode <> '') then begin
            ItemVariant.SetLoadFields(Description);
            HasVariant := ItemVariant.Get(No, VariantCode);
            ProductElement.Add(XmlElement.Create('ModelID', XmlNamespaceRAM, VariantCode));
        end;
        ProductElement.Add(XmlElement.Create('Name', XmlNamespaceRAM, Description));
        if IsExtendedProfile() then begin
            AddElementIfNotEmpty(ProductElement, 'Description', Description2);
            if HasVariant then
                AddElementIfNotEmpty(ProductElement, 'ModelName', ItemVariant.Description);
            if IsItem then
                InsertItemData(ProductElement, No);
        end;
        LineElement.Add(ProductElement);
    end;

    /// <summary>
    /// BG-32 item attributes, BT-158 classification (tariff no., list HS), BT-159 country of origin, BG-X-94 manufacturer.
    /// </summary>
    local procedure InsertItemData(var ProductElement: XmlElement; ItemNo: Code[20])
    var
        Item: Record Item;
        ItemAttribute: Record "Item Attribute";
        ItemAttributeValue: Record "Item Attribute Value";
        ItemAttributeValueMapping: Record "Item Attribute Value Mapping";
        Manufacturer: Record Manufacturer;
        CharacteristicElement, ClassificationElement, ManufacturerElement : XmlElement;
    begin
        Item.SetLoadFields("Tariff No.", "Country/Region of Origin Code", "Manufacturer Code");
        if not Item.Get(ItemNo) then
            exit;

        ItemAttributeValueMapping.SetRange("Table ID", Database::Item);
        ItemAttributeValueMapping.SetRange("No.", ItemNo);
        if ItemAttributeValueMapping.FindSet() then
            repeat
                if ItemAttribute.Get(ItemAttributeValueMapping."Item Attribute ID") and ItemAttributeValue.Get(ItemAttributeValueMapping."Item Attribute ID", ItemAttributeValueMapping."Item Attribute Value ID") then
                    if (ItemAttribute.Name <> '') and (ItemAttributeValue.Value <> '') then begin
                        CharacteristicElement := XmlElement.Create('ApplicableProductCharacteristic', XmlNamespaceRAM);
                        CharacteristicElement.Add(XmlElement.Create('Description', XmlNamespaceRAM, ItemAttribute.Name));
                        CharacteristicElement.Add(XmlElement.Create('Value', XmlNamespaceRAM, ItemAttributeValue.Value));
                        ProductElement.Add(CharacteristicElement);
                    end;
            until ItemAttributeValueMapping.Next() = 0;

        if Item."Tariff No." <> '' then begin
            ClassificationElement := XmlElement.Create('DesignatedProductClassification', XmlNamespaceRAM);
            ClassificationElement.Add(XmlElement.Create('ClassCode', XmlNamespaceRAM, XmlAttribute.Create('listID', 'HS'), Item."Tariff No."));
            ProductElement.Add(ClassificationElement);
        end;

        if Item."Country/Region of Origin Code" <> '' then
            ProductElement.Add(XmlElement.Create('OriginTradeCountry', XmlNamespaceRAM,
                XmlElement.Create('ID', XmlNamespaceRAM, GetCountryISOCode(Item."Country/Region of Origin Code"))));

        if Item."Manufacturer Code" <> '' then begin
            Manufacturer.SetLoadFields(Name);
            if Manufacturer.Get(Item."Manufacturer Code") and (Manufacturer.Name <> '') then begin
                ManufacturerElement := XmlElement.Create('ManufacturerTradeParty', XmlNamespaceRAM);
                ManufacturerElement.Add(XmlElement.Create('Name', XmlNamespaceRAM, Manufacturer.Name));
                ProductElement.Add(ManufacturerElement);
            end;
        end;
    end;

    /// <summary>
    /// BG-X-81 seller order and the blanket order as contract of the line, EXTENDED only.
    /// They precede NetPriceProductTradePrice in the XSD sequence.
    /// </summary>
    local procedure InsertLineReferences(var LineTradeAgreementElement: XmlElement; OrderNo: Code[20]; OrderLineNo: Integer; BlanketOrderNo: Code[20]; BlanketOrderLineNo: Integer)
    begin
        if not IsExtendedProfile() then
            exit;
        InsertReferencedDocument(LineTradeAgreementElement, 'SellerOrderReferencedDocument', OrderNo, OrderLineNo, 0D);
        InsertReferencedDocument(LineTradeAgreementElement, 'ContractReferencedDocument', BlanketOrderNo, BlanketOrderLineNo, 0D);
    end;

    /// <summary>
    /// BT-X-85 delivery date and BG-X-83 delivery note of the line, EXTENDED only: the posted shipment the invoice line
    /// comes from (combined invoices), either linked directly or through the item ledger entries.
    /// </summary>
    local procedure InsertLineDeliveryNote(var LineTradeDeliveryElement: XmlElement; SalesInvoiceLine: Record "Sales Invoice Line")
    var
        TempSalesShipmentLine: Record "Sales Shipment Line" temporary;
        SalesShipmentHeader: Record "Sales Shipment Header";
        ShipmentNo: Code[20];
        ShipmentLineNo: Integer;
    begin
        if not IsExtendedProfile() then
            exit;
        if SalesInvoiceLine."Shipment No." <> '' then begin
            ShipmentNo := SalesInvoiceLine."Shipment No.";
            ShipmentLineNo := SalesInvoiceLine."Shipment Line No.";
        end else
            if SalesInvoiceLine.Type = SalesInvoiceLine.Type::Item then begin
                SalesInvoiceLine.GetSalesShptLines(TempSalesShipmentLine);
                if TempSalesShipmentLine.FindFirst() then begin
                    ShipmentNo := TempSalesShipmentLine."Document No.";
                    ShipmentLineNo := TempSalesShipmentLine."Line No.";
                end;
            end;
        if ShipmentNo = '' then
            exit;
        SalesShipmentHeader.SetLoadFields("Posting Date");
        if not SalesShipmentHeader.Get(ShipmentNo) then
            Clear(SalesShipmentHeader);
        InsertActualDeliveryDate(LineTradeDeliveryElement, SalesShipmentHeader."Posting Date");
        InsertReferencedDocument(LineTradeDeliveryElement, 'DeliveryNoteReferencedDocument', ShipmentNo, ShipmentLineNo, SalesShipmentHeader."Posting Date");
    end;

    local procedure InsertReferencedDocument(var ParentElement: XmlElement; ElementName: Text; IssuerAssignedID: Text; LineNo: Integer; IssueDate: Date)
    var
        ReferencedDocumentElement, IssueDateElement : XmlElement;
    begin
        if IssuerAssignedID = '' then
            exit;
        ReferencedDocumentElement := XmlElement.Create(ElementName, XmlNamespaceRAM);
        ReferencedDocumentElement.Add(XmlElement.Create('IssuerAssignedID', XmlNamespaceRAM, IssuerAssignedID));
        if LineNo <> 0 then
            ReferencedDocumentElement.Add(XmlElement.Create('LineID', XmlNamespaceRAM, Format(LineNo)));
        if IssueDate <> 0D then begin
            IssueDateElement := XmlElement.Create('FormattedIssueDateTime', XmlNamespaceRAM);
            IssueDateElement.Add(XmlElement.Create('DateTimeString', XmlNamespaceQDT, XmlAttribute.Create('format', '102'), FormatDate(IssueDate)));
            ReferencedDocumentElement.Add(IssueDateElement);
        end;
        ParentElement.Add(ReferencedDocumentElement);
    end;

    /// <summary>
    /// Line VAT (BG-30). EXTENDED adds the exemption reason of the line, BT-X-96 / BT-X-97.
    /// </summary>
    local procedure CreateLineTradeTaxElement(TaxCategory: Code[10]; VATBusPostingGroup: Code[20]; VATProdPostingGroup: Code[20]; VATPercent: Decimal) ApplicableTradeTaxElement: XmlElement
    var
        VATEXCode: Text;
        VATClauseDescription: Text;
    begin
        if IsExtendedProfile() then
            PeppolVATHelper.GetVATClauseInfo(VATBusPostingGroup, VATProdPostingGroup, DocumentLanguageCode, VATEXCode, VATClauseDescription);
        ApplicableTradeTaxElement := XmlElement.Create('ApplicableTradeTax', XmlNamespaceRAM);
        ApplicableTradeTaxElement.Add(XmlElement.Create('TypeCode', XmlNamespaceRAM, 'VAT'));
        AddElementIfNotEmpty(ApplicableTradeTaxElement, 'ExemptionReason', VATClauseDescription);
        ApplicableTradeTaxElement.Add(XmlElement.Create('CategoryCode', XmlNamespaceRAM, GetTaxCategoryID(TaxCategory, VATBusPostingGroup, VATProdPostingGroup)));
        AddElementIfNotEmpty(ApplicableTradeTaxElement, 'ExemptionReasonCode', VATEXCode);
        ApplicableTradeTaxElement.Add(XmlElement.Create('RateApplicablePercent', XmlNamespaceRAM, FormatFiveDecimal(VATPercent)));
    end;

    /// <summary>
    /// Line discount (BG-27). EXTENDED adds BT-138 percentage and BT-140 reason code 95 (discount).
    /// </summary>
    local procedure InsertLineDiscount(var LineTradeSettlementElement: XmlElement; LineDiscountAmount: Decimal; LineDiscountPercent: Decimal)
    var
        AllowanceChargeElement: XmlElement;
    begin
        if LineDiscountAmount = 0 then
            exit;
        AllowanceChargeElement := XmlElement.Create('SpecifiedTradeAllowanceCharge', XmlNamespaceRAM);
        AllowanceChargeElement.Add(XmlElement.Create('ChargeIndicator', XmlNamespaceRAM, XmlElement.Create('Indicator', XmlNamespaceUDT, false)));
        if IsExtendedProfile() and (LineDiscountPercent <> 0) then
            AllowanceChargeElement.Add(XmlElement.Create('CalculationPercent', XmlNamespaceRAM, FormatFiveDecimal(LineDiscountPercent)));
        AllowanceChargeElement.Add(XmlElement.Create('ActualAmount', XmlNamespaceRAM, FormatDecimal(LineDiscountAmount)));
        if IsExtendedProfile() then
            AllowanceChargeElement.Add(XmlElement.Create('ReasonCode', XmlNamespaceRAM, '95'));
        AllowanceChargeElement.Add(XmlElement.Create('Reason', XmlNamespaceRAM, 'Line Discount'));
        LineTradeSettlementElement.Add(AllowanceChargeElement);
    end;

    /// <summary>
    /// BT-127 line note, EXTENDED only: the text lines attached to the line (extended texts), which are not exported as lines.
    /// </summary>
    local procedure GetAttachedLineText(LineRecordRef: RecordRef) AttachedText: Text
    var
        SalesCrMemoLine: Record "Sales Cr.Memo Line";
        SalesInvoiceLine: Record "Sales Invoice Line";
        AttachedSalesCrMemoLine: Record "Sales Cr.Memo Line";
        AttachedSalesInvoiceLine: Record "Sales Invoice Line";
    begin
        case LineRecordRef.Number() of
            Database::"Sales Invoice Line":
                begin
                    LineRecordRef.SetTable(SalesInvoiceLine);
                    AttachedSalesInvoiceLine.SetLoadFields(Description);
                    AttachedSalesInvoiceLine.SetRange("Document No.", SalesInvoiceLine."Document No.");
                    AttachedSalesInvoiceLine.SetRange("Attached to Line No.", SalesInvoiceLine."Line No.");
                    AttachedSalesInvoiceLine.SetRange(Type, AttachedSalesInvoiceLine.Type::" ");
                    if AttachedSalesInvoiceLine.FindSet() then
                        repeat
                            AddLineText(AttachedText, AttachedSalesInvoiceLine.Description);
                        until AttachedSalesInvoiceLine.Next() = 0;
                end;
            Database::"Sales Cr.Memo Line":
                begin
                    LineRecordRef.SetTable(SalesCrMemoLine);
                    AttachedSalesCrMemoLine.SetLoadFields(Description);
                    AttachedSalesCrMemoLine.SetRange("Document No.", SalesCrMemoLine."Document No.");
                    AttachedSalesCrMemoLine.SetRange("Attached to Line No.", SalesCrMemoLine."Line No.");
                    AttachedSalesCrMemoLine.SetRange(Type, AttachedSalesCrMemoLine.Type::" ");
                    if AttachedSalesCrMemoLine.FindSet() then
                        repeat
                            AddLineText(AttachedText, AttachedSalesCrMemoLine.Description);
                        until AttachedSalesCrMemoLine.Next() = 0;
                end;
        end;
    end;

    local procedure AddLineText(var AttachedText: Text; LineText: Text)
    begin
        if LineText = '' then
            exit;
        if AttachedText <> '' then
            AttachedText += ' ';
        AttachedText += LineText;
    end;

    /// <summary>
    /// BT-X-4 document language (ISO 639-1), EXTENDED only. It precedes IncludedNote in ExchangedDocument.
    /// </summary>
    local procedure InsertLanguageID(var ExchangedDocumentElement: XmlElement; LanguageCode: Code[10])
    var
        Language: Codeunit Language;
    begin
        if not IsExtendedProfile() then
            exit;
        if LanguageCode = '' then
            exit;
        AddElementIfNotEmpty(ExchangedDocumentElement, 'LanguageID', LowerCase(Language.GetTwoLetterISOLanguageName(Language.GetLanguageIdOrDefault(LanguageCode))));
    end;

    /// <summary>
    /// BT-22 note from the work description of the document, EXTENDED only. It comes before the notes of subscribers.
    /// </summary>
    local procedure AddWorkDescriptionNote(HeaderRecordRef: RecordRef; var TempZUGFeRDNote: Record "ZUGFeRD Note" temporary)
    var
        SalesCrMemoHeader: Record "Sales Cr.Memo Header";
        SalesInvoiceHeader: Record "Sales Invoice Header";
        WorkDescription: Text;
    begin
        if not IsExtendedProfile() then
            exit;
        case HeaderRecordRef.Number() of
            Database::"Sales Invoice Header":
                begin
                    HeaderRecordRef.SetTable(SalesInvoiceHeader);
                    WorkDescription := SalesInvoiceHeader.GetWorkDescription();
                end;
            Database::"Sales Cr.Memo Header":
                begin
                    HeaderRecordRef.SetTable(SalesCrMemoHeader);
                    WorkDescription := SalesCrMemoHeader.GetWorkDescription();
                end;
        end;
        if WorkDescription <> '' then
            TempZUGFeRDNote.AddNote('', CopyStr(WorkDescription, 1, MaxStrLen(TempZUGFeRDNote.Content)));
    end;

    /// <summary>
    /// BT-X-145 Incoterms, EXTENDED only, when the shipment method code is an Incoterms 2020 code.
    /// ApplicableTradeDeliveryTerms precedes the order references in ApplicableHeaderTradeAgreement.
    /// </summary>
    local procedure InsertDeliveryTerms(var HeaderTradeAgreementElement: XmlElement; ShipmentMethodCode: Code[10])
    var
        DeliveryTermsElement: XmlElement;
    begin
        if not IsExtendedProfile() then
            exit;
        if not (ShipmentMethodCode in ['EXW', 'FCA', 'CPT', 'CIP', 'DAP', 'DPU', 'DDP', 'FAS', 'FOB', 'CFR', 'CIF']) then
            exit;
        DeliveryTermsElement := XmlElement.Create('ApplicableTradeDeliveryTerms', XmlNamespaceRAM);
        DeliveryTermsElement.Add(XmlElement.Create('DeliveryTypeCode', XmlNamespaceRAM, ShipmentMethodCode));
        HeaderTradeAgreementElement.Add(DeliveryTermsElement);
    end;

    /// <summary>
    /// BT-11 project (SpecifiedProcuringProject), EXTENDED only, when all lines of the document belong to one project.
    /// </summary>
    local procedure InsertProcuringProject(var HeaderTradeAgreementElement: XmlElement; HeaderRecordRef: RecordRef)
    var
        Job: Record Job;
        SalesCrMemoLine: Record "Sales Cr.Memo Line";
        SalesInvoiceLine: Record "Sales Invoice Line";
        ProjectElement: XmlElement;
        JobNo: Code[20];
        DocumentNo: Code[20];
    begin
        if not IsExtendedProfile() then
            exit;
        DocumentNo := HeaderRecordRef.Field(3).Value();
        case HeaderRecordRef.Number() of
            Database::"Sales Invoice Header":
                begin
                    SalesInvoiceLine.SetLoadFields("Job No.");
                    SalesInvoiceLine.SetRange("Document No.", DocumentNo);
                    SalesInvoiceLine.SetFilter("Job No.", '<>%1', '');
                    if not SalesInvoiceLine.FindFirst() then
                        exit;
                    JobNo := SalesInvoiceLine."Job No.";
                    SalesInvoiceLine.SetFilter("Job No.", '<>%1&<>%2', '', JobNo);
                    if not SalesInvoiceLine.IsEmpty() then
                        exit;
                end;
            Database::"Sales Cr.Memo Header":
                begin
                    SalesCrMemoLine.SetLoadFields("Job No.");
                    SalesCrMemoLine.SetRange("Document No.", DocumentNo);
                    SalesCrMemoLine.SetFilter("Job No.", '<>%1', '');
                    if not SalesCrMemoLine.FindFirst() then
                        exit;
                    JobNo := SalesCrMemoLine."Job No.";
                    SalesCrMemoLine.SetFilter("Job No.", '<>%1&<>%2', '', JobNo);
                    if not SalesCrMemoLine.IsEmpty() then
                        exit;
                end;
            else
                exit;
        end;
        Job.SetLoadFields(Description);
        if not Job.Get(JobNo) then
            Clear(Job);
        ProjectElement := XmlElement.Create('SpecifiedProcuringProject', XmlNamespaceRAM);
        ProjectElement.Add(XmlElement.Create('ID', XmlNamespaceRAM, JobNo));
        // Name is mandatory in the XSD.
        if Job.Description <> '' then
            ProjectElement.Add(XmlElement.Create('Name', XmlNamespaceRAM, Job.Description))
        else
            ProjectElement.Add(XmlElement.Create('Name', XmlNamespaceRAM, JobNo));
        HeaderTradeAgreementElement.Add(ProjectElement);
    end;

    /// <summary>
    /// BT-30 legal registration ID and BT-28 / BT-45 trading name. The registration ID is written as today; the trading
    /// name ("Name 2") only for EXTENDED.
    /// </summary>
    local procedure InsertLegalOrganization(var TradePartyElement: XmlElement; RegistrationID: Text; TradingName: Text)
    var
        LegalOrganizationElement: XmlElement;
    begin
        if not IsExtendedProfile() then
            TradingName := '';
        if (RegistrationID = '') and (TradingName = '') then
            exit;
        LegalOrganizationElement := XmlElement.Create('SpecifiedLegalOrganization', XmlNamespaceRAM);
        AddElementIfNotEmpty(LegalOrganizationElement, 'ID', RegistrationID);
        AddElementIfNotEmpty(LegalOrganizationElement, 'TradingBusinessName', TradingName);
        TradePartyElement.Add(LegalOrganizationElement);
    end;

    local procedure GetCountyIfExtended(County: Text): Text
    begin
        if IsExtendedProfile() then
            exit(County);
        exit('');
    end;

    /// <summary>
    /// BT-6 VAT accounting currency, EXTENDED only, for documents in foreign currency with posted VAT entries.
    /// BT-6 requires BT-111 (VAT total in that currency), see InsertTaxTotalInLCY.
    /// </summary>
    local procedure InsertTaxCurrencyCode(var SettlementElement: XmlElement; DocumentCurrencyCode: Code[10]; DocumentType: Enum "Gen. Journal Document Type"; DocumentNo: Code[20]; PostingDate: Date)
    var
        VATAmountLCY: Decimal;
    begin
        if not GetVATAmountLCY(DocumentCurrencyCode, DocumentType, DocumentNo, PostingDate, VATAmountLCY) then
            exit;
        SettlementElement.Add(XmlElement.Create('TaxCurrencyCode', XmlNamespaceRAM, GeneralLedgerSetup."LCY Code"));
    end;

    /// <summary>
    /// BT-111 VAT total in local currency (from the VAT entries of the document), EXTENDED only.
    /// </summary>
    local procedure InsertTaxTotalInLCY(var MonetarySummationElement: XmlElement; DocumentCurrencyCode: Code[10]; DocumentType: Enum "Gen. Journal Document Type"; DocumentNo: Code[20]; PostingDate: Date)
    var
        VATAmountLCY: Decimal;
    begin
        if not GetVATAmountLCY(DocumentCurrencyCode, DocumentType, DocumentNo, PostingDate, VATAmountLCY) then
            exit;
        MonetarySummationElement.Add(XmlElement.Create('TaxTotalAmount', XmlNamespaceRAM, XmlAttribute.Create('currencyID', GeneralLedgerSetup."LCY Code"), FormatDecimal(VATAmountLCY)));
    end;

    local procedure GetVATAmountLCY(DocumentCurrencyCode: Code[10]; DocumentType: Enum "Gen. Journal Document Type"; DocumentNo: Code[20]; PostingDate: Date; var VATAmountLCY: Decimal): Boolean
    var
        VATEntry: Record "VAT Entry";
    begin
        if not IsExtendedProfile() then
            exit(false);
        if (DocumentCurrencyCode = '') or (DocumentCurrencyCode = GeneralLedgerSetup."LCY Code") then
            exit(false);
        VATEntry.SetRange(Type, VATEntry.Type::Sale);
        VATEntry.SetRange("Document Type", DocumentType);
        VATEntry.SetRange("Document No.", DocumentNo);
        VATEntry.SetRange("Posting Date", PostingDate);
        if VATEntry.IsEmpty() then
            exit(false);
        VATEntry.CalcSums(Amount);
        VATAmountLCY := Abs(VATEntry.Amount);
        exit(true);
    end;

    /// <summary>
    /// BG-X-41 exchange rate, EXTENDED only: 1 unit of the document currency in local currency, at the posting date.
    /// It follows the invoicee and precedes the payment means.
    /// </summary>
    local procedure InsertCurrencyExchange(var SettlementElement: XmlElement; DocumentCurrencyCode: Code[10]; CurrencyFactor: Decimal; PostingDate: Date)
    var
        CurrencyExchangeElement, RateDateElement : XmlElement;
    begin
        if not IsExtendedProfile() then
            exit;
        if (DocumentCurrencyCode = '') or (CurrencyFactor = 0) then
            exit;
        CurrencyExchangeElement := XmlElement.Create('TaxApplicableTradeCurrencyExchange', XmlNamespaceRAM);
        CurrencyExchangeElement.Add(XmlElement.Create('SourceCurrencyCode', XmlNamespaceRAM, DocumentCurrencyCode));
        CurrencyExchangeElement.Add(XmlElement.Create('TargetCurrencyCode', XmlNamespaceRAM, GeneralLedgerSetup."LCY Code"));
        CurrencyExchangeElement.Add(XmlElement.Create('ConversionRate', XmlNamespaceRAM, FormatDecimalUnlimited(Round(1 / CurrencyFactor, 0.000001))));
        if PostingDate <> 0D then begin
            RateDateElement := XmlElement.Create('ConversionRateDateTime', XmlNamespaceRAM);
            RateDateElement.Add(XmlElement.Create('DateTimeString', XmlNamespaceUDT, XmlAttribute.Create('format', '102'), FormatDate(PostingDate)));
            CurrencyExchangeElement.Add(RateDateElement);
        end;
        SettlementElement.Add(CurrencyExchangeElement);
    end;

    /// <summary>
    /// BG-3 preceding invoice of a credit memo (BT-25, BT-26), EXTENDED only: the invoice the credit memo is applied to.
    /// InvoiceReferencedDocument follows the monetary summation.
    /// </summary>
    local procedure InsertPrecedingInvoiceReference(var SettlementElement: XmlElement; AppliesToDocType: Enum "Gen. Journal Document Type"; AppliesToDocNo: Code[20])
    var
        SalesInvoiceHeader: Record "Sales Invoice Header";
    begin
        if not IsExtendedProfile() then
            exit;
        if (AppliesToDocType <> AppliesToDocType::Invoice) or (AppliesToDocNo = '') then
            exit;
        SalesInvoiceHeader.SetLoadFields("Posting Date");
        if not SalesInvoiceHeader.Get(AppliesToDocNo) then
            Clear(SalesInvoiceHeader);
        InsertReferencedDocument(SettlementElement, 'InvoiceReferencedDocument', AppliesToDocNo, 0, SalesInvoiceHeader."Posting Date");
    end;
    #endregion

    #region CommonFunctions
    local procedure GetSetups()
    begin
        CompanyInformation.Get();
        GeneralLedgerSetup.Get();
        OnAfterGetSetups(CompanyInformation, GeneralLedgerSetup);
    end;

    procedure FormatDate(VarDate: Date): Text[20];
    begin
        if VarDate = 0D then
            exit('17530101');
        exit(Format(VarDate, 0, '<Year4><Month,2><Day,2>'));
    end;

    procedure FormatDecimal(VarDecimal: Decimal): Text
    var
        TypeHelper: Codeunit "Type Helper";
    begin
        exit(Format(VarDecimal, 0, TypeHelper.GetXMLAmountFormatWithTwoDecimalPlaces()))
    end;

    procedure FormatDecimalUnlimited(VarDecimal: Decimal): Text
    begin
        exit(Format(VarDecimal, 0, 9));
    end;

#if not CLEAN29
    [Obsolete('FormatFourDecimal is no longer used internally. Quantity and unit price fields are now exported with unlimited precision per EN 16931. For VAT percentages, use FormatFiveDecimal. For quantity and unit price, use FormatDecimalUnlimited.', '29.0')]
    procedure FormatFourDecimal(VarDecimal: Decimal): Text
    begin
        exit(Format(VarDecimal, 0, '<Precision,4:4><Standard Format,9>'));
    end;
#endif

    procedure FormatFiveDecimal(VarDecimal: Decimal): Text
    begin
        exit(Format(Round(VarDecimal, 0.00001), 0, 9));
    end;

    procedure GetUoMCode(UoMCode: Code[10]): Text;
    var
        UnitofMeasure: Record "Unit of Measure";
    begin
        if not UnitofMeasure.Get(UoMCode) then
            exit(Format(UoMCode));
        if UnitofMeasure."International Standard Code" <> '' then
            exit(UnitofMeasure."International Standard Code");
        exit(UoMCode);
    end;

    local procedure GetCurrencyCode(DocumentCurrencyCode: Code[10]; var Currency: Record Currency): Code[10]
    begin
        if DocumentCurrencyCode = '' then begin
            Currency.InitRoundingPrecision();
            exit(GeneralLedgerSetup."LCY Code");
        end else begin
            Currency.Get(DocumentCurrencyCode);
            Currency.TestField("Amount Rounding Precision");
            Currency.TestField("Unit-Amount Rounding Precision");
            exit(DocumentCurrencyCode);
        end;
    end;

    local procedure GetCountryRegionCode(CountryRegionCode: Code[10]): Code[10]
    begin
        if CountryRegionCode <> '' then
            exit(CountryRegionCode);

        CompanyInformation.TestField("Country/Region Code");
        exit(CompanyInformation."Country/Region Code");
    end;

    local procedure GetCountryISOCode(CountryRegionCode: Code[10]): Code[2]
    var
        CountryRegion: Record "Country/Region";
    begin
        CountryRegion.Get(CountryRegionCode);
        exit(CountryRegion."ISO Code");
    end;

    procedure GetVATRegistrationNo(VATRegistrationNo: Text[20]; CountryRegionCode: Code[10]): Text[30];
    begin
        if CountryRegionCode = '' then
            CountryRegionCode := GetCountryRegionCode(CountryRegionCode);
        if CopyStr(VATRegistrationNo, 1, 2) <> CountryRegionCode then
            exit(CountryRegionCode + VATRegistrationNo);
        exit(VATRegistrationNo);
    end;

    local procedure GetTaxCategoryID(TaxCategory: Code[10]; VATBusPostingGroup: Code[20]; VATProductPostingGroup: Code[20]): Text[10];
    var
        VATPostingSetup: Record "VAT Posting Setup";
    begin
        if TaxCategory <> '' then
            exit(TaxCategory);
        if not VATPostingSetup.Get(VATBusPostingGroup, VATProductPostingGroup) then
            exit('');
        exit(VATPostingSetup."Tax Category");
    end;

    procedure SetEDocumentService(NewEDocumentService: Record "E-Document Service")
    begin
        EDocumentService := NewEDocumentService;
#if not CLEAN30
        EDocumentServiceProvided := true;
#endif
    end;

    local procedure ResetProvidedService()
    begin
        // Clear the per-instance service state at the end of every run so a reused instance never carries
        // the service provided for an earlier export into a later export that does not provide one.
        Clear(EDocumentService);
#if not CLEAN30
        EDocumentServiceProvided := false;
#endif
    end;

#if not CLEAN30
#pragma warning disable AA0228, AL0432
    [Obsolete('The triggering E-Document Service is now provided through SetEDocumentService (threaded via "ZUGFeRD Export Context"). This function is still called on every export until CLEAN30: when a service was provided it only raises OnAfterFindEDocumentService without overwriting that service, otherwise it performs the legacy FindLast lookup. As of CLEAN30 this function and that fallback are removed, so any caller that does not provide a service - for example a customized sales report that never sets the ZUGFeRD Export Context - then exports with a blank E-Document Service.', '30.0')]
    local procedure FindEDocumentService()
    begin
        // A service provided through SetEDocumentService is the service that triggered the export -
        // never overwrite it with a lookup, but still raise the event so existing subscribers keep working.
        // When nothing was provided (for example a customized report that does not use the context)
        // the original lookup runs unchanged.
        if EDocumentServiceProvided then begin
            OnAfterFindEDocumentService(EDocumentService);
            exit;
        end;

        EDocumentService.SetRange("Document Format", EDocumentService."Document Format"::ZUGFeRD);
        if EDocumentService.FindLast() then;
        OnAfterFindEDocumentService(EDocumentService);
    end;
#pragma warning restore AA0228, AL0432
#endif

    local procedure GetBankAccountPaymentDetails(BankAccountCode: Code[20]; var IBAN: Text[50]; var SWIFTCode: Code[20]; var AccountName: Text[100])
    var
        BankAccount: Record "Bank Account";
    begin
        // BT-85 is the name of the account holder, the seller, not the bank account's own name in BC.
        AccountName := CompanyInformation.Name;
        if BankAccountCode <> '' then
            if BankAccount.Get(BankAccountCode) then begin
                IBAN := BankAccount.IBAN;
                SWIFTCode := BankAccount."SWIFT Code";
                exit;
            end;

        IBAN := CompanyInformation.IBAN;
        SWIFTCode := CompanyInformation."SWIFT Code";
    end;
    #endregion

#if not CLEAN30
#pragma warning disable AA0228
    [Obsolete('The triggering E-Document Service is now provided through SetEDocumentService (threaded via "ZUGFeRD Export Context"). This event is STILL raised on every export until CLEAN30 - both when a service was provided and on the legacy FindLast lookup path - so existing subscribers keep working during the deprecation window. It no longer exists as of CLEAN30; move any logic that depends on it to the service provided through SetEDocumentService.', '30.0')]
    [IntegrationEvent(false, false)]
    local procedure OnAfterFindEDocumentService(var EDocumentService: Record "E-Document Service")
    begin
    end;
#pragma warning restore AA0228
#endif

    [IntegrationEvent(false, false)]
    local procedure OnAfterInsertSalesInvHeaderData(var XMLCurrNode: XmlElement; SalesInvoiceHeader: Record "Sales Invoice Header")
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnCreateXMLOnBeforeSalesInvXmlDocumentWriteToFile(var XMLDoc: XmlDocument; SalesInvoiceHeader: Record "Sales Invoice Header")
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnDocumentLinesExistOnAfterFilterSalesInvLine(var SalesInvoiceHeader: Record "Sales Invoice Header"; var SalesInvLine: Record "Sales Invoice Line")
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnAfterInsertSalesCrMemoHeaderData(var XMLCurrNode: XmlElement; SalesCrMemoHeader: Record "Sales Cr.Memo Header")
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnCreateXMLOnBeforeSalesCrMemoXmlDocumentWriteToFile(var XMLDoc: XmlDocument; SalesCrMemoHeader: Record "Sales Cr.Memo Header")
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnDocumentLinesExistOnAfterFilterSalesCrMemoLine(var SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var SalesCrMemoLine: Record "Sales Cr.Memo Line")
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnAfterCalculateInvoiceLineAmounts(var SalesInvoiceHeader: Record "Sales Invoice Header"; var SalesInvLine: Record "Sales Invoice Line"; Currency: Record Currency; var LineAmounts: Dictionary of [Text, Decimal])
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnAfterCalculateCrMemoLineAmounts(var SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var SalesCrMemoLine: Record "Sales Cr.Memo Line"; Currency: Record Currency; var LineAmounts: Dictionary of [Text, Decimal])
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeGenerateSalesInvoicePDFAttachment(SalesInvoiceHeader: Record "Sales Invoice Header"; var TempBlob: Codeunit "Temp Blob"; var IsHandled: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeGenerateSalesCrMemoPDFAttachment(SalesCrMemoHeader: Record "Sales Cr.Memo Header"; var TempBlob: Codeunit "Temp Blob"; var IsHandled: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeGenerateServiceInvoicePDFAttachment(ServiceInvoiceHeader: Record "Service Invoice Header"; var TempBlob: Codeunit "Temp Blob"; var IsHandled: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeGenerateServiceCrMemoPDFAttachment(ServiceCrMemoHeader: Record "Service Cr.Memo Header"; var TempBlob: Codeunit "Temp Blob"; var IsHandled: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnAfterGetSetups(var CompanyInformation: Record "Company Information"; var GeneralLedgerSetup: Record "General Ledger Setup")
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeInsertInvoiceLine(var SupplyChainTradeTransactionElement: XmlElement; var SalesInvoiceLine: Record "Sales Invoice Line"; Currency: Record Currency; CurrencyCode: Code[10]; PricesIncVAT: Boolean; var IsHandled: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeAddInvoiceLineElement(var InvoiceLineElement: XmlElement; var SalesInvoiceLine: Record "Sales Invoice Line"; Currency: Record Currency; CurrencyCode: Code[10]; PricesIncVAT: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeInsertCrMemoLine(var SupplyChainTradeTransactionElement: XmlElement; var SalesCrMemoLine: Record "Sales Cr.Memo Line"; Currency: Record Currency; CurrencyCode: Code[10]; PricesIncVAT: Boolean; var IsHandled: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeAddCrMemoLineElement(var CrMemoLineElement: XmlElement; var SalesCrMemoLine: Record "Sales Cr.Memo Line"; Currency: Record Currency; CurrencyCode: Code[10]; PricesIncVAT: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnAfterInsertApplicableHeaderTradeAgreement(var HeaderTradeAgreementElement: XmlElement; HeaderRecordRef: RecordRef)
    begin
    end;

    /// <summary>
    /// Fires immediately before the generated payment means element is added to the document root.
    /// Subscribe to inspect or extend the payment block, using HeaderRecRef to access any field of the source document.
    /// </summary>
    [IntegrationEvent(false, false)]
    local procedure OnInsertPaymentMethodOnBeforeAddToRoot(var PaymentMethodElement: XmlElement; HeaderRecRef: RecordRef)
    begin
    end;

    /// <summary>
    /// Fires after the ZUGFeRD profile of a document was read from the E-Document Service.
    /// Subscribe to choose the profile per document, for example per customer.
    /// The export and the checks use the same resolution, so they always agree on the profile.
    /// </summary>
    /// <param name="EDocumentService">The E-Document Service that triggered the export. Blank for a plain report print.</param>
    /// <param name="HeaderRecordRef">The posted source document: sales or service invoice or credit memo.</param>
    /// <param name="ZUGFeRDProfile">The profile to export. Change it to override the profile of the service.</param>
    [IntegrationEvent(false, false)]
    local procedure OnAfterResolveZUGFeRDProfile(EDocumentService: Record "E-Document Service"; HeaderRecordRef: RecordRef; var ZUGFeRDProfile: Enum "ZUGFeRD Profile")
    begin
    end;

    /// <summary>
    /// Fires when the document header (ExchangedDocument) is written. Add notes (BG-1) to the buffer; the export writes
    /// them as IncludedNote elements in the order of the buffer, for example a note with subject code ADU or REG.
    /// </summary>
    /// <param name="HeaderRecordRef">The posted source document: sales or service invoice or credit memo.</param>
    /// <param name="ZUGFeRDProfile">The profile that is exported.</param>
    /// <param name="TempZUGFeRDNote">The notes to write. Use AddNote to add a note.</param>
    [IntegrationEvent(false, false)]
    local procedure OnCollectHeaderNotes(HeaderRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile"; var TempZUGFeRDNote: Record "ZUGFeRD Note" temporary)
    begin
    end;

    /// <summary>
    /// Fires for every invoice or credit memo line before its line identification (AssociatedDocumentLineDocument) is written,
    /// also for additional lines. Set the line ID (BT-126) and, for the EXTENDED profile, the parent line (BT-X-304) and the
    /// sub-line type (BT-X-8). ParentLineID must be the line ID of another line of the document (BR-FXEXT-11).
    /// A line with a parent line and no sub-line type is written as DETAIL. Parent line and sub-line type are ignored for other profiles.
    /// </summary>
    /// <param name="LineRecordRef">The line: a Sales Invoice Line or Sales Cr.Memo Line. Service lines are passed as the sales line they are transferred to.</param>
    /// <param name="ZUGFeRDProfile">The profile that is exported.</param>
    /// <param name="LineID">The line ID. The default is the line number.</param>
    /// <param name="ParentLineID">The line ID of the parent line. Blank for a line without a parent.</param>
    /// <param name="LineSubtype">The sub-line type.</param>
    [IntegrationEvent(false, false)]
    local procedure OnInsertLineOnBeforeCreateAssociatedDocumentLine(LineRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile"; var LineID: Text; var ParentLineID: Text; var LineSubtype: Enum "ZUGFeRD Line Subtype")
    begin
    end;

    /// <summary>
    /// Fires once per sales or service invoice before any amount is calculated. Add lines that are not posted document lines
    /// but must be in the e-invoice, for example the old-part value of an exchange part. The export writes them after the
    /// document lines and includes them in the VAT breakdown and the document totals, like posted lines.
    /// Fill at least Line No., Description, Quantity, Unit of Measure Code, Unit Price, Amount, Amount Including VAT, VAT %,
    /// and Tax Category or the VAT posting groups. Amounts follow the Prices Including VAT setting of the document, like posted lines.
    /// The line number must not be used by a line of the document.
    /// </summary>
    /// <param name="HeaderRecordRef">The posted source document: sales or service invoice.</param>
    /// <param name="ZUGFeRDProfile">The profile that is exported.</param>
    /// <param name="TempAdditionalSalesInvLine">The additional lines.</param>
    [IntegrationEvent(false, false)]
    local procedure OnCollectAdditionalInvoiceLines(HeaderRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile"; var TempAdditionalSalesInvLine: Record "Sales Invoice Line" temporary)
    begin
    end;

    /// <summary>
    /// Fires once per sales or service credit memo before any amount is calculated. See OnCollectAdditionalInvoiceLines.
    /// </summary>
    /// <param name="HeaderRecordRef">The posted source document: sales or service credit memo.</param>
    /// <param name="ZUGFeRDProfile">The profile that is exported.</param>
    /// <param name="TempAdditionalSalesCrMemoLine">The additional lines.</param>
    [IntegrationEvent(false, false)]
    local procedure OnCollectAdditionalCrMemoLines(HeaderRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile"; var TempAdditionalSalesCrMemoLine: Record "Sales Cr.Memo Line" temporary)
    begin
    end;

    /// <summary>
    /// Fires before the document totals (SpecifiedTradeSettlementHeaderMonetarySummation) are written.
    /// Set the paid amount (BT-113). The amount due for payment (BT-115) is the invoice total with VAT (BT-112) minus the paid amount.
    /// </summary>
    /// <param name="HeaderRecordRef">The posted source document: sales or service invoice or credit memo.</param>
    /// <param name="ZUGFeRDProfile">The profile that is exported.</param>
    /// <param name="PrepaidAmount">The paid amount. The default is 0.</param>
    [IntegrationEvent(false, false)]
    local procedure OnInsertMonetarySummationOnBeforeCreateElement(HeaderRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile"; var PrepaidAmount: Decimal)
    begin
    end;

    /// <summary>
    /// Fires immediately before the delivery block (ApplicableHeaderTradeDelivery) is added to the document.
    /// Subscribe to append elements that follow the actual delivery date in the XSD sequence, for example the despatch advice reference (BT-16).
    /// </summary>
    /// <param name="DeliveryElement">The delivery block.</param>
    /// <param name="HeaderRecordRef">The posted source document: sales or service invoice or credit memo.</param>
    /// <param name="ZUGFeRDProfile">The profile that is exported.</param>
    [IntegrationEvent(false, false)]
    local procedure OnAfterInsertApplicableHeaderTradeDelivery(var DeliveryElement: XmlElement; HeaderRecordRef: RecordRef; ZUGFeRDProfile: Enum "ZUGFeRD Profile")
    begin
    end;

    /// <summary>
    /// Copies fields with matching numbers, classes, and types between two record variants.
    /// W1 PEPPOL30 only exposes generic Sales Header / Sales Line targets, but ZUGFeRD/XRechnung exports
    /// operate on the posted Sales Invoice/Cr.Memo Header and Line shapes - this local helper provides
    /// the posted-to-posted field copy that the old PEPPOL Management.Transfer*ToSalesInvoice/CrMemo*
    /// helpers used to do.
    /// </summary>
    local procedure TransferRecordFields(FromRecord: Variant; var ToRecord: Variant)
    var
        FromRecRef: RecordRef;
        ToRecRef: RecordRef;
        FromFieldRef: FieldRef;
        ToFieldRef: FieldRef;
        i: Integer;
    begin
        FromRecRef.GetTable(FromRecord);
        ToRecRef.GetTable(ToRecord);
        for i := 1 to FromRecRef.FieldCount do begin
            FromFieldRef := FromRecRef.FieldIndex(i);
            if ToRecRef.FieldExist(FromFieldRef.Number) then begin
                ToFieldRef := ToRecRef.Field(FromFieldRef.Number);
                if (FromFieldRef.Class = ToFieldRef.Class) and (FromFieldRef.Type = ToFieldRef.Type) and (FromFieldRef.Length <= ToFieldRef.Length) then
                    ToFieldRef.Value := FromFieldRef.Value();
            end;
        end;
        ToRecRef.SetTable(ToRecord);
    end;

    local procedure TransferToSalesInvoiceHeader(FromRecord: Variant; var ToSalesInvoiceHeader: Record "Sales Invoice Header")
    var
        ToRecord: Variant;
    begin
        ToRecord := ToSalesInvoiceHeader;
        TransferRecordFields(FromRecord, ToRecord);
        ToSalesInvoiceHeader := ToRecord;
    end;

    local procedure TransferToSalesCrMemoHeader(FromRecord: Variant; var ToSalesCrMemoHeader: Record "Sales Cr.Memo Header")
    var
        ToRecord: Variant;
    begin
        ToRecord := ToSalesCrMemoHeader;
        TransferRecordFields(FromRecord, ToRecord);
        ToSalesCrMemoHeader := ToRecord;
    end;

    local procedure TransferToSalesInvoiceLine(FromRecord: Variant; var ToSalesInvoiceLine: Record "Sales Invoice Line")
    var
        ToRecord: Variant;
    begin
        ToRecord := ToSalesInvoiceLine;
        TransferRecordFields(FromRecord, ToRecord);
        ToSalesInvoiceLine := ToRecord;
    end;

    local procedure TransferToSalesCrMemoLine(FromRecord: Variant; var ToSalesCrMemoLine: Record "Sales Cr.Memo Line")
    var
        ToRecord: Variant;
    begin
        ToRecord := ToSalesCrMemoLine;
        TransferRecordFields(FromRecord, ToRecord);
        ToSalesCrMemoLine := ToRecord;
    end;
}
