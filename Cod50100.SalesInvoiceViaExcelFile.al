codeunit 50100 SalesInvoiceViaExcelFile
{
    trigger OnRun()
    var
        ExcelCount: Integer;
        ColText: array[50] of Text;
        i: Integer;
        SalesOrderNo: Code[20];
    begin
        ImportAndReadExcelFile;
        if ExcelBuffer.IsEmpty then
            exit;

        ExcelCount := 0;
        ExcelBuffer.reset;
        if ExcelBuffer.FindLast() then
            ExcelCount := ExcelBuffer."Row No.";

        ExcelCount := 0;
        ExcelBuffer.reset;
        if ExcelBuffer.FindLast() then
            ExcelCount := ExcelBuffer."Row No.";

        for i := 2 to ExcelCount do begin
            Clear(ColText);
            ExcelBuffer.RESET;
            ExcelBuffer.SetRange("Row No.", i);
            if ExcelBuffer.FindSet() then
                repeat
                    ColText[ExcelBuffer."Column No."] := ExcelBuffer."Cell Value as Text";
                until ExcelBuffer.next = 0;
            SalesOrderNo := '';
            FindAndCreateSalesHeader(ColText, SalesOrderNo);
            CreateSalesLine(ColText, SalesOrderNo);
        end;
        ScheduleSalesInvoicePostProcess;
    end;

    procedure ImportAndReadExcelFile()
    var
        tblob: Codeunit "Temp Blob";
        istream: InStream;
        Sheetname: Text;
        FileMgmt: Codeunit "File Management";
    begin
        Clear(tblob);
        FileMgmt.BLOBImportWithFilter(tblob, 'Import Excel Format', '', '*.xlsx|*.xlsx', 'All files (*.*)|*.*');
        if tblob.HasValue() then begin
            tblob.CreateInStream(istream);
            Sheetname := ExcelBuffer.SelectSheetsNameStream(istream);
            if Sheetname = '' then
                Error('Sheet Name not selected.');
            ExcelBuffer.OpenBookStream(istream, Sheetname);
            ExcelBuffer.ReadSheet();
        end else
            Error('File has not been selected for the Import.');
    end;

    procedure FindAndCreateSalesHeader(ColText: array[50] of Text; var SaleOrderNo: Code[20])
    var
        NotFound: Boolean;
        SalesHeader: Record "Sales Header";
    begin
        NotFound := true;
        if OrderNo.Get(ColText[1], SaleOrderNo) then
            NotFound := false;

        if NotFound then begin
            SalesHeader.Init();
            SalesHeader."Document Type" := SalesHeader."Document Type"::Invoice;
            SalesHeader.Insert(true);
            SaleOrderNo := SalesHeader."No.";
            OrderNo.Add(ColText[1], SalesHeader."No.");
            SalesHeader."External Document No." := ColText[1];
            SalesHeader.Validate("Sell-to Customer No.", ColText[2]);
            Evaluate(SalesHeader."Posting Date", ColText[3]);
            SalesHeader.Modify(true);
        end;
    end;

    procedure CreateSalesLine(ColText: array[50] of Text; SaleOrderNo: Code[20])
    var
        SalesLine: Record "Sales Line";
        LineNno: Integer;
    begin
        GetLastSalesLineNo(SaleOrderNo, LineNno);
        SalesLine.Init();
        SalesLine."Document Type" := SalesLine."Document Type"::Invoice;
        SalesLine."Document No." := SaleOrderNo;
        SalesLine."Line No." := LineNno + 10000;
        SalesLine.Type := SalesLine.Type::Item;
        SalesLine."No." := ColText[4];
        SalesLine.Validate("No.");
        Evaluate(SalesLine.Quantity, ColText[5]);
        SalesLine.Validate(Quantity);
        Evaluate(SalesLine."Unit Price", ColText[6]);
        SalesLine.Validate("Unit Price");
        SalesLine.Insert(true);
    end;

    procedure GetLastSalesLineNo(SaleOrderNo: Code[20]; var LineNno: Integer)
    var
        SalesLine: Record "Sales Line";
    begin
        LineNno := 0;
        SalesLine.SetRange("Document Type", SalesLine."Document Type"::Invoice);
        SalesLine.SetRange("Document No.", SaleOrderNo);
        if SalesLine.FindLast() then
            LineNno := SalesLine."Line No.";
    end;

    procedure ScheduleSalesInvoicePostProcess()
    var
        SalesPostProcess: Codeunit "Sales Post via Job Queue";
        InvNo: Code[20];
        SalesInvNo: Code[20];
        SalesHeader: Record "Sales Header";
    begin
        foreach InvNo in OrderNo.Keys() do begin
            OrderNo.Get(InvNo, SalesInvNo);
            SalesHeader.Get(SalesHeader."Document Type"::Invoice, SalesInvNo);
            SalesPostProcess.EnqueueSalesDocWithUI(SalesHeader, false);
        end;
    end;

    var
        ExcelBuffer: Record "Excel Buffer" temporary;
        OrderNo: Dictionary of [Code[20], Code[20]];
}
