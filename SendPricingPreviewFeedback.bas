Attribute VB_Name = "SendPricingPreviewFeedback"
Function SendpricingPreview(Email)
    Sheets("Pricing Tracking File").Select
    notifications_to_go = 0
    notificacao = ""
    cont = 0
    For Each linha In Range(Range("E2"), Range("E2").End(xlDown))
        If linha.EntireRow.Hidden = False Then
            If Cells(linha.Row, "A").Text = "Exploratory" Then
                If Cells(linha.Row, "AC").Text <> "" And Cells(linha.Row, "S").Text <> "" Then
                    If Cells(linha.Row, "AC").Text = Email Then
                        Cells(linha.Row, "A").FormulaR1C1 = "Not for upload"
                        Cells(linha.Row, "AD").FormulaR1C1 = Format(Now(), "MM/DD/YYYY")
                        Eu = usuario()
                        JobCode = Cells(linha.Row, "E").Text
                        compZone = Cells(linha.Row, "F").Text
                        JobProfile = Cells(linha.Row, "H").Text
                        Country = Cells(linha.Row, "K").Text
                        TheCurrency = Cells(linha.Row, "Q").Text
                        TheMin = Cells(linha.Row, "R").Text
                        TheMidpoint = Cells(linha.Row, "S").Text
                        TheMax = Cells(linha.Row, "T").Text
                        notificacao = notificacao + "<tr><td>" + JobCode + "</td><td>" + compZone + "</td><td>" + JobProfile + "</td><td>" + TheCurrency + "</td><td>" + TheMin + "</td><td>" + TheMidpoint + "</td><td>" + TheMax + "</td></tr>"
                        cont = cont + 1
                    End If
                End If
                'notifications_to_go = notifications_to_go + 1
            End If
        End If
    Next
    wordJobs = "job"
    If cont > 1 Then
        wordJobs = "jobs"
    End If
    notificacao = "<style>table {font - family: arial, sans - serif;border-collapse: collapse;width: 100 %;}td, th {border: 1px solid #000000;text-align: left;padding: 8px;}</style><body><p>Hi,</p><p>We have received your pricing preview request. Here's the information you requested:</p><table><tr><th>Job Code</th><th>Comp Market</th><th>Job Profile</th><th>Currency</th></th><th>Minimum</th></th><th>Midpoint</th></th><th>Maximum</th></tr>" + notificacao + "</table>"
    notificacao = notificacao + "<p>Please, note that this is not being uploaded to Workday, we are only sharing it with you for modeling purposes. If you need these ranges on Workday, please submit a case for it <a href = 'https://dell.service-now.com/hrportal?id=hri_sc_cat_item&sys_id=03f9e24adbf63680ae487eb6bf961944'>here</a>.</p>"
    notificacao = notificacao + "<p>Thanks,</p>"
    notificacao = notificacao + "<p>" & Eu & "<br>Name of the Team<br>Name of the company</p></body>"

    notificacao = notificacao + "</body>"
    If cont > 0 Then
        Set objOutlook = Outlook.Application
        Set objMail = objOutlook.CreateItem(olMailItem)
        objMail.To = Email
        objMail.CC = ""
        objMail.Subject = "Market Pricing Preview Feedback"
        objMail.HTMLBody = notificacao
        objMail.Send
        'objMail.Close (olDiscard) 'Draft email
        Set objMail = Nothing
    End If

End Function



Sub Send_Pricing_Preview_Feedback()
    If ActiveSheet.Name <> "Pricing Tracking File" Then
        MsgBox "It seems that you have the wrong sheet selected. Please make sure you're on the adhoc activations working file on the 'Pricing Tracking File' sheet, then select the row for which you'd like to send the notification and try again."
    Else
    
    
    Dim ArrayEmails() As String
a = 0
ReDim Preserve ArrayEmails(a)
ArrayEmails(a) = "Email List"
a = a + 1

    For Each linha In Range(Range("E2"), Range("E2").End(xlDown))
        If linha.EntireRow.Hidden = False Then
            If Cells(linha.Row, "A").Text = "Exploratory" Then
                If Cells(linha.Row, "AC").Text <> "" Then
                    found = 0
                    For Each Email In ArrayEmails
                        If Cells(linha.Row, "AC").Text = Email Then
                            found = 1
                        End If
                    Next
                    If found = 0 Then
                        ReDim Preserve ArrayEmails(a)
                        ArrayEmails(a) = Cells(linha.Row, "AC").Text
                        a = a + 1
                    End If
                End If
            End If
        End If
    Next
    contador = 0
For b = 1 To UBound(ArrayEmails)
    contador = contador + 1
    SendpricingPreview (ArrayEmails(b))
Next
MsgBox "Sent pricing feedback to " & contador & IIf(contador = 1, " requestor.", " requestors.")
    

    End If
End Sub

