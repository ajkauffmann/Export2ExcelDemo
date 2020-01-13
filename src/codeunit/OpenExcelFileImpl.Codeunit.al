codeunit 70103 OpenExcelFileImpl
{
    procedure CreateAndOpenExcelFile(
        var TempExcelBuf: Record "Excel Buffer" temporary;
        BookName: Text)
    begin
        TempExcelBuf.SetFriendlyFilename(BookName);
        TempExcelBuf.OpenExcel();
    end;
}