codeunit 70100 "Export Customer 2 Excel"
{
    var
        BookNameTxt: Label 'Export Customers';
        SheetNameTxt: Label 'Customers';
        HeaderTxt: Label 'Export Customers';
        ChoiceTxt: Label 'Open as File,Email as Attachment';


    trigger OnRun()
    begin
        Export2Excel();
    end;

    local procedure Export2Excel()
    var
        TempExcelBuf: Record "Excel Buffer" temporary;
        Choice: Integer;
    begin
        Choice := StrMenu(ChoiceTxt);
        if Choice = 0 then
            exit;

        CreateAndFillExcelBuffer(TempExcelBuf);

        if Choice = 1 then
            DownloadAndOpenExcel(TempExcelBuf)
        else
            TempExcelBuf.EmailExcelFile(BookNameTxt);
    end;

    local procedure CreateAndFillExcelBuffer(var TempExcelBuf: Record "Excel Buffer" temporary)
    begin
        TempExcelBuf.CreateNewBook(SheetNameTxt);
        FillExcelBuffer(TempExcelBuf);
        TempExcelBuf.WriteSheet(HeaderTxt, CompanyName(), UserId());
        TempExcelBuf.CloseBook();
    end;

    local procedure FillExcelBuffer(var TempExcelBuf: Record "Excel Buffer" temporary)
    var
        Customer: Record Customer;
    begin
        if Customer.FindSet() then
            repeat
                FillExcelRow(TempExcelBuf, Customer);
            until Customer.Next() = 0;
    end;

    local procedure FillExcelRow(
        var TempExcelBuf: Record "Excel Buffer" temporary;
        Customer: Record Customer)
    begin
        with Customer do begin
            TempExcelBuf.NewRow();
            TempExcelBuf.AddColumn("No.", false, '', false, false, false, '', TempExcelBuf."Cell Type"::Text);
            TempExcelBuf.AddColumn(Name, false, '', false, false, false, '', TempExcelBuf."Cell Type"::Text);
            TempExcelBuf.AddColumn(Address, false, '', false, false, false, '', TempExcelBuf."Cell Type"::Text);
            TempExcelBuf.AddColumn("Post Code", false, '', false, false, false, '', TempExcelBuf."Cell Type"::Text);
            TempExcelBuf.AddColumn(City, false, '', false, false, false, '', TempExcelBuf."Cell Type"::Text);
            TempExcelBuf.AddColumn("Country/Region Code", false, '', false, false, false, '', TempExcelBuf."Cell Type"::Text);
        end;
    end;

    local procedure DownloadAndOpenExcel(var TempExcelBuf: Record "Excel Buffer" temporary)
    begin
        TempExcelBuf.SetFriendlyFilename(BookNameTxt);
        TempExcelBuf.OpenExcel();
    end;
}