codeunit 70102 EmailExcelFileImpl
{
    var
        ExcelFileExtensionTok: Label '.xlsx', Locked = true;
        EmailSentTxt: Label 'File has been sent by email';
        Book1Txt: Label 'Book1';

    procedure CreateAndSendEmail(
        var TempExcelBuf: Record "Excel Buffer" temporary;
        BookName: Text)
    var
        SMTPMail: Codeunit "SMTP Mail";
        Recipients: List of [Text];
    begin
        Recipients.Add('bc@cronus.company');
        SMTPMail.CreateMessage(
            'Business Central Mail',
            'bc@cronus.company',
            Recipients,
            'Test Export to Excel and email',
            'This is an text to export data to an Excel file and email it.',
            false);

        AddAttachment(SMTPMail, TempExcelBuf, GetFriendlyFilename(BookName));
        SMTPMail.Send();
        Message(EmailSentTxt);
    end;

    local procedure AddAttachment(
        var SMTPMail: Codeunit "SMTP Mail";
        var TempExcelBuf: Record "Excel Buffer" temporary;
        BookName: Text)
    var
        TempBlob: Codeunit "Temp Blob";
        InStr: InStream;
    begin
        ExportExcelFileToBlob(TempExcelBuf, TempBlob);
        TempBlob.CreateInStream(InStr);
        SMTPMail.AddAttachmentStream(InStr, BookName);
    end;

    local procedure ExportExcelFileToBlob(
        var TempExcelBuf: Record "Excel Buffer" temporary;
        var TempBlob: Codeunit "Temp Blob")
    var
        OutStr: OutStream;
    begin
        TempBlob.CreateOutStream(OutStr);
        TempExcelBuf.SaveToStream(OutStr, true);
    end;

    local procedure GetFriendlyFilename(BookName: Text): Text
    var
        FileManagement: Codeunit "File Management";
    begin
        if BookName = '' then
            exit(Book1Txt + ExcelFileExtensionTok);

        exit(FileManagement.StripNotsupportChrInFileName(BookName) + ExcelFileExtensionTok);
    end;
}