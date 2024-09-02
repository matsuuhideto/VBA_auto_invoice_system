Sub Final()

Dim row
Dim name
Dim body
Dim mailto
Dim price
Dim ordernumber
Dim filelocation

row = 2

Do Until Sheets("order").Cells(row, 1) = ""
If Sheets("order").Cells(row, 2) <> "A" Then
Sheets("order").Cells(row, 8) = "-"
ElseIf Sheets("order").Cells(row, 8) = 0 Then
Sheets("invoice_answer").Range("L4") = Sheets("order").Cells(row, 1).Value
Sheets("invoice_answer").printpreview

ordernumber = Sheets("invoice_answer").Range("L4")
filelocation = "C:\Temp\VBAExercies\Invoice" & ordernumber & ".pdf"
    Sheets("invoice_answer").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        filelocation, Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    Set WordApp = CreateObject("word.Application")
    WordApp.Documents.Open "C:\Temp\VBAExercies\mail_template.docx"
    WordApp.Visible = True
    WordApp.Selection.WholeStory
    body = WordApp.Selection.Text
    WordApp.Quit

name = Sheets("invoice_answer").Range("B5")
price = Sheets("invoice_answer").Range("L36")
body = Replace(body, "xx", name)
body = Replace(body, "yy", price)
mailto = Sheets("invoice_answer").Range("T3")


  Set OutApp = CreateObject("Outlook.Application")
  Set objMsg = OutApp.CreateItem(olMailItem)


    With objMsg
        .to = mailto
        .Subject = "test"
        .body = body
        .attachments.Add filelocation
        .Display

    End With
Sheets("order").Cells(row, 8) = Date

Else
End If

row = row + 1
Loop

End Sub
