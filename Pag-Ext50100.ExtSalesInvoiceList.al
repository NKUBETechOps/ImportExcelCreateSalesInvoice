pageextension 50100 "Ext-Sales Invoice List" extends "Sales Invoice List"
{
    actions
    {
        addafter("&Invoice")
        {
            action("Import Sales Invoices from Excel")
            {
                ApplicationArea = All;
                Caption = 'Import Sales Invoices from Excel';
                Image = Import;
                trigger OnAction()
                var
                    SalesInvoiceViaExcelFile: Codeunit SalesInvoiceViaExcelFile;
                begin
                    SalesInvoiceViaExcelFile.Run();
                end;
            }
        }
    }
}
