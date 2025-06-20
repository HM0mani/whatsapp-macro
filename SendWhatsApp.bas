Sub SendSingleWhatsAppMessage(rowNum As Long)
    Dim ws As Worksheet
    Dim phone As String
    Dim name As String
    Dim consume As Double, produce As Double
    Dim due As Double, earned As Double, net As Double
    Dim template As String, message As String, url As String

    Set ws = ThisWorkbook.Sheets("Sheet1")

    name = ws.Cells(rowNum, 1).Value
    phone = Replace(ws.Cells(rowNum, 2).Value, "+", "") ' Remove "+" if present
    consume = Val(ws.Cells(rowNum, 17).Value)
    produce = Val(ws.Cells(rowNum, 18).Value)

    ' Calculate due and earned without adjusting production
    due = consume * 0.75
    earned = produce * 0.75

    ' Net result + 1 JOD as fee or deduction
    net = due - earned + 1

    ' Select correct template
    If net >= 0 Then
        template = Sheets("Template").Range("A1").Value
    Else
        template = Sheets("Template").Range("B1").Value
    End If

    ' Replace placeholders
    template = Replace(template, "{production}", produce)
    template = Replace(template, "{consumption}", consume)
    template = Replace(template, "{result}", Format(Abs(net), "0.00") & " JOD")

    ' Encode for WhatsApp
    message = Replace(template, Chr(10), "%0A")
    message = Replace(message, " ", "%20")
    
    ' TEST: Clear the message (so nothing is sent)
    message = ""

    ' Build WhatsApp Web URL
    url = "https://web.whatsapp.com/send?phone=" & phone & "&text=" & message

    ' Open in browser
    ThisWorkbook.FollowHyperlink Address:=url

    ' ?? Visually mark the row as "sent"
    ws.Rows(rowNum).Interior.Color = RGB(204, 255, 204) ' light green
    ws.Cells(rowNum, 20).Value = "? Sent" ' optional: add status in column T
End Sub

