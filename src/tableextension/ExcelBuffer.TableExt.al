tableextension 70100 "Excel Buffer Ext" extends "Excel Buffer"
{
    procedure EmailExcelFile(BookName: Text)
    var
        EmailExcelFileImpl: Codeunit EmailExcelFileImpl;
    begin
        EmailExcelFileImpl.CreateAndSendEmail(Rec, BookName);
    end;
}