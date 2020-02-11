report 70100 "MyExcelTest"
{
    Caption = 'My Excel Test';
    UsageCategory = Administration;
    ApplicationArea = All;
    ProcessingOnly = true;

    dataset
    {
        dataitem(Customer; Customer)
        {
            trigger OnPreDataItem()
            begin
                WriteExcelWkshtHeader();
            end;

            trigger OnAfterGetRecord()
            begin
                TempExcelBuffer.NewRow(); // increase row number, set col number to 0
                TempExcelBuffer.AddColumn("No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Name, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Address, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn("Post Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(City, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn("Balance (LCY)", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
            end;

            trigger OnPostDataItem()
            begin
                WriteExcelWkshtFooter();
            end;
        }
    }

    var
        TempExcelBuffer: Record "Excel Buffer" temporary;

    trigger OnPostReport()
    begin
        TempExcelBuffer.CreateNewBook('FirstSheet');
        TempExcelBuffer.WriteSheet('My Customers', CompanyName(), UserId());

        // to create an additional sheet in the workbook
        TempExcelBuffer.SelectOrAddSheet('SecondSheet');

        // normally, do TempExcelBuffer.DeleteAll() to clear the buffer, and fill it with 
        // new data for the new sheet before writing it to the book
        // this just creates a second copy of the same sheet because we didnt clear the data
        TempExcelBuffer.WriteSheet('My Customers 2', CompanyName(), UserId());

        TempExcelBuffer.CloseBook(); // to clear internal variables for excel writer objects
        TempExcelBuffer.SetFriendlyFilename('MyCustomerTest'); // otherwise you get 'Book1' as file name
        TempExcelBuffer.OpenExcel(); // doesnt open Excel, but copies the xlsx file to the Downloads folder
    end;

    local procedure WriteExcelWkshtHeader();
    var
        MyCurrentRow: Integer;
        OneWay: Boolean;
        ReportTitleTxt: Label 'My Excel Test Customer List';
    begin
        // creates the Excel Buffer records to create the header cells of the current sheet

        TempExcelBuffer.ClearNewRow(); // sets row number to 0 and column number to 0
        TempExcelBuffer.NewRow(); // increases row number by 1, sets column number to 0 => row 1, col 0
        TempExcelBuffer.AddColumn(CopyStr(CompanyName(), 1, 250), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.NewRow(); // => row 2, col 0
        TempExcelBuffer.AddColumn(ReportTitleTxt, false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.NewRow(); // => row 3, col 0
        TempExcelBuffer.AddColumn(Format(CurrentDateTime()) + ': ' + UserId(), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);

        if OneWay then begin
            TempExcelBuffer.NewRow(); // => row 4, col 0
            TempExcelBuffer.NewRow(); // => row 5, col 0 -> where the column headers go
        end else begin
            // instead of doing NewRow twice, you can also set the pointer to a specific row/col
            // you'll first need the current row number, use F12 to drill down into the Excel Buffer function
            TempExcelBuffer.GetCurrentRow(MyCurrentRow); // from the CurrentDateTime row, so MyCurrentRow is now 3
            // so to skip two rows, you then do:
            TempExcelBuffer.SetCurrent(MyCurrentRow + 2, 0); // => row 5, col 0 -> where the column headers go
            // check out WriteExcelWkshtFooter how to use SetCurrent for skipping columns
        end;

        TempExcelBuffer.AddColumn(Customer.FieldCaption("No."), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Customer.FieldCaption(Name), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Customer.FieldCaption(Address), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Customer.FieldCaption("Post Code"), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Customer.FieldCaption(City), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Customer.FieldCaption("Balance (LCY)"), false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
    end;

    local procedure WriteExcelWkshtFooter();
    var
        MyCurrentRow: Integer;
        TotalBalanceTxt: Label 'Total Balance';
    begin
        TempExcelBuffer.NewRow();
        TempExcelBuffer.GetCurrentRow(MyCurrentRow);
        TempExcelBuffer.SetCurrent(MyCurrentRow, 1);
        TempExcelBuffer.AddColumn(TotalBalanceTxt, false, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.SetCurrent(MyCurrentRow, 5);
        TempExcelBuffer.AddColumn('=SUM(F6:F' + Format(MyCurrentRow - 1) + ')', true, '', true, false, false, '', TempExcelBuffer."Cell Type"::Text);
    end;
}