pageextension 59900 "Sales Order Subform IS" extends "Sales Order Subform"
{
    actions
    {
        addafter(EditInExcel)
        {
            action("&Import")
            {
                Caption = 'Import Excel';
                Image = ImportExcel;
                Promoted = true;
                PromotedCategory = Process;
                ApplicationArea = All;
                Visible = true;
                ToolTip = 'Import data from excel.';
                trigger OnAction()
                var
                    ImportExcel: Codeunit ImportExcel;
                begin
                    ImportExcel.ImportExcelData(Rec);
                end;
            }
        }
    }
}
