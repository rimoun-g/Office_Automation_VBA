VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Find Attachments By Name"
   ClientHeight    =   8528.001
   ClientLeft      =   104
   ClientTop       =   429
   ClientWidth     =   8632.001
   OleObjectBlob   =   "outlook find attachments by name.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGetAttachments_Click()
'On Error GoTo On_Error

    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As attachment
    Dim Report As String
    
     Dim Folder As Outlook.Folder
     
     Dim objNS As Outlook.NameSpace
     Set Folder = Application.ActiveExplorer.CurrentFolder

Dim objFolder As Outlook.MAPIFolder
Set objNS = GetNamespace("MAPI")
Set objFolder = Folder
     
          x = 0
    For Each selItem In objFolder.Items
        If selItem.Class = olMail Then
            Set aMail = selItem
            For Each aAttach In aMail.Attachments
             info = GetAttachmentInfo(aAttach)
             spInfo = Split(info, vbCrLf)
            
            lstAllAttachments.AddItem (LCase(spInfo(0)))
            lstAllAttachments.List(x, 1) = spInfo(1)
             x = x + 1
            Next
        End If
    Next



lblFilesNumber.Caption = (lstAllAttachments.ListCount)


'On_Error:
'    MsgBox "error=" & Err.Number & " " & Err.Description
'      MsgBox "file=" & info

End Sub

Private Sub lstFoundItems_Click()

End Sub

Private Sub txtSearch_Change()
'On Error GoTo On_Error

If txtSearch.Text <> "" Or txtSearch.Text <> " " Then

lstIndex = lstAllAttachments.ListCount



lstFoundItems.Clear



For i = 1 To lstIndex


DamnVar = InStr(lstAllAttachments.List(i - 1, 0), txtSearch.Text)

If DamnVar > 0 Then lstFoundItems.AddItem (lstAllAttachments.List(i - 1, 0))

Next

End If
lblFoundResultsCounter.Caption = "Found Files: " & lstFoundItems.ListCount
'On_Error:
'    MsgBox "error=" & Err.Number & " " & Err.Description
'      MsgBox "file=" & info
End Sub

Private Sub UserForm_Initialize()

lstAllAttachments.ColumnCount = 2

' lstAllAttachments.AddItem ("file one.docx")
' lstAllAttachments.AddItem ("Noname.docx")

End Sub



Public Function GetAttachmentInfo(attachment As attachment)
    Dim Report
    GetAttachmentInfo = ""
   ' Report = Report & "Index: " & attachment.Index & vbCrLf
    Report = Report & attachment.DisplayName & vbCrLf
    'Report = Report & "File Name: " & attachment.FileName & vbCrLf
    'Report = Report & "Block Level: " & attachment.BlockLevel & vbCrLf
    'Report = Report & "Path Name: " & attachment.PathName & vbCrLf
    'Report = Report & "Position: " & attachment.Position & vbCrLf
   Report = Report & attachment.Size
    'Report = Report & "Type: " & attachment.Type & vbCrLf
     
    GetAttachmentInfo = Report
End Function
