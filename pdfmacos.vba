Sub PDFTabsBetweenStartAndEnd()

Dim X As Long
Dim pdfPath As String

pdfPath = "/Users/yourmacnamehere/Library/Group Containers/UBF8T346G9.Office/MyFolder" 'replace "yourmacnamehere" with the name of your mac'
  
Sheets(Sheets("PDF - Start").Index + 1).Select
For X = Sheets("PDF - Start").Index + 1 To Sheets("PDF - End").Index - 1
If Sheets(X).Visible = True Then
    Sheets(X).Select False
End If
Next
  
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
FileName:=pdfPath & "/" & "\Insert Name Here.pdf", _
Quality:=xlQualityStandard, IncludeDocProperties:=True, _
IgnorePrintAreas:=False, OpenAfterPublish:=False
        
Sheets("PDF - Start").Select

End Sub
