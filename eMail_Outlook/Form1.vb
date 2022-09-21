Imports outlook = Microsoft.Office.Interop.Outlook

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim OutlookMessage As outlook.MailItem
        Dim AppOutlook As New outlook.Application
        Try
            OutlookMessage = AppOutlook.CreateItem(outlook.OlItemType.olMailItem)
            Dim Recipents As outlook.Recipients = OutlookMessage.Recipients
            Recipents.Add("someone@somewhere.uk")
            OutlookMessage.Subject = "Sending through Outlook"
            OutlookMessage.Body = "Testing outlook Mail"
            OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
            OutlookMessage.Send()
        Catch ex As Exception
            MessageBox.Show("Mail could not be sent") 'if you dont want this message, simply delete this line
        Finally
            OutlookMessage = Nothing
            AppOutlook = Nothing
        End Try
    End Sub

End Class