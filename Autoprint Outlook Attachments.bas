Attribute VB_Name = "Módulo1"
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
  "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
  ByVal lpFile As String, ByVal lpParameters As String, _
  ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private WithEvents Items As Outlook.Items

Private Sub Application_Startup()
  Dim Ns As Outlook.NameSpace
  Dim Folder As Outlook.MAPIFolder

  Set Ns = Application.GetNamespace("MAPI")
  Set Folder = Ns.GetDefaultFolder(olFolderInbox)
  Set Items = Folder.Items
End Sub

Private Sub Items_ItemAdd(ByVal Item As Object)
  If TypeOf Item Is Outlook.MailItem Then
    PrintAttachments Item
  End If
End Sub

Private Sub PrintAttachments(oMail As Outlook.MailItem)
  On Error Resume Next
  Dim colAtts As Outlook.Attachments
  Dim oAtt As Outlook.Attachment
  Dim sFile As String
  Dim sDirectory As String
  Dim sFileType As String

  sDirectory = "D:\Attachments\"

  Set colAtts = oMail.Attachments
  Set trigSubject = oMail.Subject

  If left$(trigSubject, 6) = “#print” then
MsgBox Left$(trigSubject, 6)
  If colAtts.Count Then
    For Each oAtt In colAtts

' This code looks at the last 4 characters in a filename
      sFileType = LCase$(Right$(oAtt.FileName, 4))

      Select Case sFileType

' Add additional file types below
      Case ".xls", ".doc", "docx", “.pdf”

        sFile = sDirectory & oAtt.FileName
        oAtt.SaveAsFile sFile
        ShellExecute 0, "print", sFile, vbNullString, vbNullString, 0
      End Select
    Next
  End If
  End If
End Sub

