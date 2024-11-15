Attribute VB_Name = "Module1"
Public Sub ATTACHMENTS_ALL(MItem As Outlook.MailItem)

    Dim oAttachment As Outlook.Attachment
    Dim SaveFolder As String
    
    SaveFolder = "C:\Users\i.janowska\Documents\OUTLOOKZALACZNIKI\"
    
    For Each oAttachment In MItem.Attachments
        On Error Resume Next ' Ignore errors
        oAttachment.SaveAsFile SaveFolder & oAttachment.DisplayName
        On Error GoTo 0 ' Reset error handling to default after save attempt
    Next
    
End Sub
