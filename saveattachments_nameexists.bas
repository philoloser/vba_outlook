Attribute VB_Name = "Module2"
Public Sub ATTACHMENTS_NAMEEXISTS(MItem As Outlook.MailItem)
    Dim olApp As Outlook.Application
    Dim olNs As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim olAttachment As Outlook.Attachment
    Dim SaveFolder As String
    Dim filePath As String

    ' Initialize Outlook application and folder
    Set olApp = New Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    Set olFolder = olNs.GetDefaultFolder(olFolderInbox)

    ' Specify the folder where you want to save attachments
    SaveFolder = "C:\Users\i.janowska\Documents\OUTLOOKZALACZNIKI\"  ' Change this to your desired folder path

    ' Loop through each item in the Inbox
    For Each olItem In olFolder.Items
        ' Loop through each attachment in the item
        For Each olAttachment In MItem.Attachments
            filePath = SaveFolder & olAttachment.FileName
            
            ' Check if the file already exists
            If Dir(filePath) = "" Then
                ' File does not exist, save the attachment
                olAttachment.SaveAsFile filePath
            Else
                ' File already exists, skip downloading
                Debug.Print "File already exists: " & olAttachment.FileName
            End If
        Next olAttachment
    Next olItem

    ' Cleanup
    Set olAttachment = Nothing
    Set olItem = Nothing
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
End Sub


