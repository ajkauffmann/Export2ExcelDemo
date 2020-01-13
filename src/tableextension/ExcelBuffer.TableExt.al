tableextension 70100 "Excel Buffer Ext" extends "Excel Buffer"
{
    procedure OpenExcelFile(BookName: Text)
    var
        OpenExcelFileImpl: Codeunit OpenExcelFileImpl;
    begin
        OpenExcelFileImpl.CreateAndOpenExcelFile(Rec, BookName);
    end;

    procedure EmailExcelFile(BookName: Text)
    var
        EmailExcelFileImpl: Codeunit EmailExcelFileImpl;
    begin
        EmailExcelFileImpl.CreateAndSendEmail(Rec, BookName);
    end;
}