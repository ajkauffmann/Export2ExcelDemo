codeunit 70101 "Export Sales Order 2 Excel"
{
    TableNo = "Sales Header";

    var
        ChoiceTxt: Label 'Open as File,Email as Attachment';
        BookNameTxt: Label 'Export Sales Order';
        SheetNameTxt: Label 'Sales Order';
        HeaderTxt: Label 'Export Customers';
        InitExcelBufferErr: Label 'Could not initialize Excel Buffer. Do you a correct template imported?';


    trigger OnRun()
    begin
        Export2Excel(Rec);
    end;

    local procedure Export2Excel(SalesHeader: Record "Sales Header")
    var
        TempExcelBuf: Record "Excel Buffer" temporary;
        Choice: Integer;
    begin
        Choice := StrMenu(ChoiceTxt);
        if Choice = 0 then
            exit;

        CreateAndFillExcelBuffer(TempExcelBuf, SalesHeader);

        if Choice = 1 then
            TempExcelBuf.OpenExcelFile(BookNameTxt)
        else
            TempExcelBuf.EmailExcelFile(BookNameTxt);
    end;

    local procedure CreateAndFillExcelBuffer(
        var TempExcelBuf: Record "Excel Buffer" temporary;
        var SalesHeader: Record "Sales Header")
    begin
        if not InitExcelBuffer(TempExcelBuf) then
            Error(InitExcelBufferErr);
        FillExcelBuffer(TempExcelBuf, SalesHeader);
        TempExcelBuf.WriteSheet(HeaderTxt, CompanyName(), UserId());
        TempExcelBuf.CloseBook();
    end;

    local procedure InitExcelBuffer(var TempExcelBuf: Record "Excel Buffer" temporary): Boolean
    var
        ExcelTemplate: Record "Excel Template";
        TempBlob: Codeunit "Temp Blob";
        InStr: InStream;
    begin
        ExcelTemplate.FindFirst();
        if not ExcelTemplate.GetTemplateFileAsTempBlob(TempBlob) then
            exit;
        
        TempBlob.CreateInStream(InStr);
        TempExcelBuf.UpdateBookStream(InStr, SheetNameTxt, true);
        exit(true);
    end;

    local procedure FillExcelBuffer(
        var TempExcelBuf: Record "Excel Buffer" temporary;
        var SalesHeader: Record "Sales Header")
    var
        TempExcelBufSheet: Record "Excel Buffer" temporary;
    begin
        FillHeaderData(TempExcelBufSheet, SalesHeader);
        FillLineData(TempExcelBufSheet, SalesHeader);
        TempExcelBuf.WriteAllToCurrentSheet(TempExcelBufSheet);
    end;

    local procedure FillHeaderData(var TempExcelBuf: Record "Excel Buffer" temporary; SalesHeader: Record "Sales Header")
    begin
        TempExcelBuf.EnterCell(TempExcelBuf, 4, 3, SalesHeader."No.", false, false, false);
        TempExcelBuf.EnterCell(TempExcelBuf, 5, 3, SalesHeader."Sell-to Customer Name", false, false, false);
        TempExcelBuf.EnterCell(TempExcelBuf, 6, 3, SalesHeader."Sell-to Contact", false, false, false);
        TempExcelBuf.EnterCell(TempExcelBuf, 7, 3, SalesHeader."Posting Date", false, false, false);
        TempExcelBuf.EnterCell(TempExcelBuf, 8, 3, SalesHeader."Order Date", false, false, false);
        TempExcelBuf.EnterCell(TempExcelBuf, 4, 6, SalesHeader."Due Date", false, false, false);
        TempExcelBuf.EnterCell(TempExcelBuf, 5, 6, SalesHeader."Requested Delivery Date", false, false, false);
        TempExcelBuf.EnterCell(TempExcelBuf, 6, 6, SalesHeader."External Document No.", false, false, false);
    end;

    local procedure FillLineData(var TempExcelBuf: Record "Excel Buffer" temporary; SalesHeader: Record "Sales Header")
    var
        SalesLine: Record "Sales Line";
        NextRowNo: Integer;
    begin
        NextRowNo := 12;
        SalesLine.SetRange("Document Type", SalesHeader."Document Type");
        SalesLine.SetRange("Document No.", SalesHeader."No.");
        if SalesLine.FindSet(false) then
            repeat
                TempExcelBuf.EnterCell(TempExcelBuf, NextRowNo, 2, SalesLine.Type, false, false, false);
                TempExcelBuf.EnterCell(TempExcelBuf, NextRowNo, 3, SalesLine."No.", false, false, false);
                TempExcelBuf.EnterCell(TempExcelBuf, NextRowNo, 4, SalesLine.Description, false, false, false);
                TempExcelBuf.EnterCell(TempExcelBuf, NextRowNo, 5, SalesLine.Quantity, false, false, false);
                TempExcelBuf.EnterCell(TempExcelBuf, NextRowNo, 6, SalesLine."Unit of Measure Code", false, false, false);
                TempExcelBuf.EnterCell(TempExcelBuf, NextRowNo, 7, SalesLine."Unit Price", false, false, false);
                TempExcelBuf.EnterCell(TempExcelBuf, NextRowNo, 8, SalesLine."Line Amount", false, false, false);
                NextRowNo += 1;
            until SalesLine.Next() = 0;
    end;
}