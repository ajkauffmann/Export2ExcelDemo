pageextension 70101 "Sales Order List Ext" extends "Sales Order List"
{
    actions
    {
        addlast(processing)
        {
            action(Export2Excel)
            {
                ApplicationArea = All;
                Caption = 'Export to Excel';
                Image = ExportToExcel;

                trigger OnAction()
                var
                    SalesHeader: Record "Sales Header";
                begin
                    SalesHeader := Rec;
                    SalesHeader.SetRecFilter();
                    Codeunit.Run(Codeunit::"Export Sales Order 2 Excel", SalesHeader)
                end;
            }
        }
    }
}