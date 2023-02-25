codeunit 59900 ImportExcel
{
    procedure ImportExcelData(var Rec: Record "Sales Line")
    var
        InStream: InStream;
        FromFile: Text;
        SheetName: Text;
    begin
        //This procedure allows you to upload a file to Business Central
        UploadIntoStream('Select the Excel file to Import', '', '', FromFile, InStream);

        if FromFile = '' then
            Error('File not found');

        SheetName := ExcelBufferTemp.SelectSheetsNameStream(InStream);

        //Subsequently, it loads the Excel Buffer data type with the information from the excel file.
        ExcelBufferTemp.Reset();
        ExcelBufferTemp.DeleteAll();
        ExcelBufferTemp.OpenBookStream(InStream, SheetName);
        ExcelBufferTemp.ReadSheet();

        //And finally, it invokes the procedure that will allow obtaining the information from each cell of the excel
        //file and inserting it into the Sales Lines.
        InsertExcelData(Rec);
    end;

    local procedure InsertExcelData(var Rec: Record "Sales Line")
    var
        SalesLine: Record "Sales Line";
        RowNo, MaxRowNo : Integer;
    begin
        RowNo := 0;
        MaxRowNo := 0;

        //To know how many lines we are going to iterate over, we calculate the value of the last line
        ExcelBufferTemp.Reset();
        if ExcelBufferTemp.FindLast() then begin
            MaxRowNo := ExcelBufferTemp."Row No.";
        end;

        //We iterate from line or row 2, since the first one is the Headers.
        for RowNo := 2 to MaxRowNo do begin

            SalesLine.Init();
            SalesLine."Document Type" := "Sales Document Type"::Order;

            //With the GetValueAt Cell procedure we obtain the information from each cell 
            //and then create the information from each Sales Line that we will insert.

            Evaluate(SalesLine."Document No.", GetValueAtCell(RowNo, 1));
            Evaluate(SalesLine."Line No.", GetValueAtCell(RowNo, 2));
            SalesLine.Validate(Type, GetType(GetValueAtCell(RowNo, 3)));
            SalesLine.Validate("No.", GetValueAtCell(RowNo, 4));
            SalesLine.Validate("Location Code", GetValueAtCell(RowNo, 5));
            Evaluate(SalesLine.Quantity, GetValueAtCell(RowNo, 6));
            Evaluate(SalesLine."Unit Price", GetValueAtCell(RowNo, 7));

            if Rec."Document No." <> SalesLine."Document No." then
                Error('The Document No%1 of line %2 is different from the Document No%3 of the Header', SalesLine."Document No.", SalesLine."Line No.", Rec."Document No.");

            if not SalesLine.Insert() then
                SalesLine.Modify();

        end;
    end;

    //Method that allows us to obtain the Type of Sales Line.
    local procedure GetType(Text: Text): Enum "Sales Line Type"
    begin

        case Text of
            ' ':
                exit("Sales Line Type"::" ");
            'G/L Account':
                exit("Sales Line Type"::"G/L Account");
            'Item':
                exit("Sales Line Type"::Item);
            'Resource':
                exit("Sales Line Type"::Resource);
            'Fixed Asset':
                exit("Sales Line Type"::"Fixed Asset");
            'Charge (Item)':
                exit("Sales Line Type"::"Charge (Item)");
        end;

        Error('The %1 field in the Type column is not valid.', Text);

    end;

    //Method that allows us to obtain the information from each cell 
    local procedure GetValueAtCell(RowNo: Integer; ColNo: Integer):
            Text
    begin
        ExcelBufferTemp.Reset();
        If ExcelBufferTemp.Get(RowNo, ColNo) then
            exit(ExcelBufferTemp."Cell Value as Text")
        else
            exit('');
    end;

    var
        ExcelBufferTemp: Record "Excel Buffer" temporary;
}
