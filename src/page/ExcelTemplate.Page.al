page 70100 "Excel Template"
{
    PageType = List;
    SourceTable = "Excel Template";
    UsageCategory = Lists;
    ApplicationArea = All;

    layout
    {
        area(Content)
        {
            repeater(lines)
            {
                field(Name; Name) { ApplicationArea = All; }
                field(HasContent; HasContent)
                {
                    ApplicationArea = All;
                    Caption = 'Has content';
                }
                field(Filename; Filename) { ApplicationArea = All; }
            }
        }
    }

    actions
    {
        area(Processing)
        {
            action(ImportFile)
            {
                ApplicationArea = All;
                Caption = 'Import file';
                Image = Import;

                trigger OnAction()
                begin
                    ImportTemplateFile();
                end;
            }
        }
    }

}