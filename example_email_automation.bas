Sub email_sheet_as_PDF()

    Dim PDF_export_path As String
    PDF_export_path = "your\path\export.pdf"

    Call Worksheets("output").ExportAsFixedFormat( _ 
        Type:=xlTypePDF, _
        Filename:=PDF_export_path _
    )

    With get_outlook_app().CreateItem(0)
        .To = "you@audience.com"
        .Subject = "Hi GTC!"
        .Body = "Hello."
        .Attachments.Add (PDF_export_path)
        .Display
    End With

    Kill PDF_export_path
    
End Sub

Function get_outlook_app() As Object

    Set get_outlook_app = CreateObject("Outlook.Application")

End Function
