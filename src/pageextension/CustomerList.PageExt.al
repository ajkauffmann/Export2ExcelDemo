pageextension 70100 "CustomerListExt" extends "Customer List"
{
    actions
    {
        addlast(processing)
        {
            action(Export2Excel)
            {
                Caption = 'Export to Excel';
                ApplicationArea = All;
                Image = Excel;
                trigger OnAction()
                begin
                    Codeunit.Run(Codeunit::"Export Customer 2 Excel");
                end;
            }
        }
    }
}