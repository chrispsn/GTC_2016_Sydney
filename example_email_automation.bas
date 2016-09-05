Sub email_sheet_as_PDF(ws As Worksheet, PDF_temp_path As String)

    Call ws.ExportAsFixedFormat( _
        Type:=xlTypePDF, _
        Filename:=PDF_temp_path _
    )

    With get_outlook_app().CreateItem(0)
        .To = "you@audience.com"
        .Subject = "Hi GTC!"
        .Body = "Hello."
        .Attachments.Add (PDF_temp_path)
        .Display
    End With

    Kill PDF_temp_path
    
End Sub

Function get_outlook_app() As Object

    Set get_outlook_app = CreateObject("Outlook.Application")

End Function
