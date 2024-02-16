Option Explicit

' Trebuie creat un subfolder in Inbox intitulat "Confirmate"
' Trebuie creat un subfolder in Inbox intitulat "Respinse"
'------------------------------------------------

' Trebuie activate Macro in Outlook
' File > Options > Trust Center > Trust Center Settings... > Macro Settings
' Enable all macros (not recommended; potentially dangerous code can run)
' -------------------------------------------------

' Trebuie copiate fisierele template .oft intr-un folder
'--------------------------------------------------

' Trebuie modificate urmatoarele variable
' FolderPath
' TemplatePath...
' Sunt marcate cu VVVVVVVVV
' -------------------------------------------------

' Pentru a funtion trebuie modificata functia eveniment din ThisOutlookSession cu urmatorul continut:
'Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
'
'    Dim EntryIDs As Variant
'    Dim EntryID As Variant
'    Dim MailItem As Object
'    Dim NS As NameSpace
'
'    Set NS = Application.GetNamespace("MAPI")
'    EntryIDs = Split(EntryIDCollection, ",")
'
'    For Each EntryID In EntryIDs
'        Set MailItem = NS.GetItemFromID(EntryID)
'        If TypeName(MailItem) = "MailItem" Then
'            ProcessNewMail MailItem
'        End If
'    Next
'End Sub

' ------------------------------------------

' Trebuie adugate doua butoare in ribbon la categoria HOME
' Cutomize Ribbon, in Home -> Add Group, Rename "Confirma/Infirma"
' Se pun trei butoane in grup, respectiv cele trei functii de confirmare, infirmare general si informare participanti
' Se da Rename, unul "Confirma" cu imaginea o bifa verde,
' Celalalt "Respinge" cu cercul rosu cu X alb ca imagine
' Al treilea cu "Respinge participant" cu triunghiul cu semnul exclamarii
' ---------------------------------------------


' Aceasta functie primeste ca argument email primit, il verifica pe rand
' Daca are linku-ri da reply automat de respingere
' Daca are atasamente mai mari de 30MB da email de respingere
' Daca are alte atasamente decat .docx sau .pdf trimite email de respingere
' In toate cele trei situatii muta emailul in subfolderul "Respinge" din Inbox

Public Sub ProcessNewMail(ByVal Item As MailItem)
    Dim Atmt As Attachment
    Dim TotalAttachmentSize As Long
    Dim FolderPath As String
    Dim RejectionFolder As MAPIFolder
    Dim TemplatePathLinkInEmail As String
    Dim TemplatePathPreaMare As String
    Dim TemplatePathFormatAtasamente As String
    
    ' ************************************************
    ' Modifica emailul
    ' Specify the folder path for "Rejection" folder (change as needed)
    
    ' Find the Rejection folder

    
    ' VVVVVVVVV
    FolderPath = "stefan.caravelea@just.ro\Inbox\Respinse" ' Adjust based on your actual folder path
        
    ' VVVVVVVVVVVVVVVVVVVVVV
    TemplatePathLinkInEmail = "C:\Users\scaravelea\OneDrive\Office_Templates\OutlookTemplates\RespingeLinkInEmail.oft" ' Specify the actual path to your template
    TemplatePathPreaMare = "C:\Users\scaravelea\OneDrive\Office_Templates\OutlookTemplates\RespingeAtasamenteMari.oft"
    TemplatePathFormatAtasamente = "C:\Users\scaravelea\OneDrive\Office_Templates\OutlookTemplates\RespingeFormatAtasamente.oft"
    
    Set RejectionFolder = GetFolder(FolderPath)
    Debug.Print RejectionFolder
        
    ' Initialize total attachment size
    TotalAttachmentSize = 0

    ' Calculate total size of attachments
    If Item.Attachments.Count > 0 Then
        For Each Atmt In Item.Attachments
            TotalAttachmentSize = TotalAttachmentSize + Atmt.Size
        Next Atmt
    End If

    ' Check for links, unacceptable attachments, and total attachment size
    If InStr(Item.Body, "http://") > 0 Or InStr(Item.Body, "https://") > 0 Then
            
        
        ' Call a function or perform an action to reply and move the email
        ReplyAndMoveTemplate Item, TemplatePathLinkInEmail, RejectionFolder
        
   ' Chech if total attachements are less then 30 MB
    ElseIf TotalAttachmentSize > 300000000 Then ' 30 MB in bytes
   

        ' Call a function or perform an action to reply and move the email
        ReplyAndMoveTemplate Item, TemplatePathPreaMare, RejectionFolder
        
    Else
        ' Check each attachment for unacceptable types
        Dim HasUnacceptableType As Boolean
        HasUnacceptableType = False
        For Each Atmt In Item.Attachments
            If Not (LCase(Right(Atmt.FileName, 5)) = ".docx" Or LCase(Right(Atmt.FileName, 4)) = ".pdf") Then
                HasUnacceptableType = True
                Exit For
            End If
        Next
        
        ' If the attachements are unacceptable types reply and sent email to rejection folder
        If HasUnacceptableType Then
       
            ReplyAndMoveTemplate Item, TemplatePathFormatAtasamente, RejectionFolder
            
        End If
    End If
End Sub



Sub ReplyAndMoveTemplate(MailItem As MailItem, TemplatePath As String, RejectionFolder As MAPIFolder)
    Dim TemplateMail As MailItem
    Dim ReplyMail As MailItem
    
    ' Create a mail item from the template
    Set TemplateMail = Application.CreateItemFromTemplate(TemplatePath)
    
    ' Create a reply from the original MailItem
    Set ReplyMail = MailItem.Reply
    
    ' Copy the content from the template to the reply
    ReplyMail.HTMLBody = TemplateMail.HTMLBody & ReplyMail.HTMLBody
    
    ' Optional: Copy subject and other properties from the template if needed
    ' Note: Prefix subject with "Re: " if it does not already include it
'    If Not Left(TemplateMail.Subject, 3) = "Re: " Then
'        ReplyMail.Subject = "Re: " & MailItem.Subject
'    Else
'        ReplyMail.Subject = TemplateMail.Subject
'    End If
    
    ReplyMail.Subject = TemplateMail.Subject
    
    ' Send the reply
    ReplyMail.Send
    
    ' Check if the MailItem is already in the RejectionFolder
    If Not MailItem.Parent.EntryID = RejectionFolder.EntryID Then
        ' Mark the original mail as read and move it to the Rejection Folder if it's not already there
        With MailItem
            .UnRead = False
            .Move RejectionFolder
            .UnRead = False
        End With
    Else
        ' Mark as read if it's already in the RejectionFolder
        MailItem.UnRead = False
    End If
    
    ' Clean up
    TemplateMail.Delete
End Sub



Function GetFolder(ByVal FolderPath As String) As MAPIFolder
    Dim FoldersArray As Variant
    Dim i As Integer
    Dim Folder As MAPIFolder
    Dim NS As NameSpace

    Set NS = Application.GetNamespace("MAPI")

    On Error GoTo ErrorHandler
    FoldersArray = Split(FolderPath, "\")
    Set Folder = NS.Folders.Item(FoldersArray(0))
    If Not Folder Is Nothing Then
        Debug.Print "Root folder found: " & Folder.Name ' Debugging line
        For i = 1 To UBound(FoldersArray, 1)
            Set Folder = Folder.Folders.Item(FoldersArray(i))
            If Folder Is Nothing Then
                Debug.Print "Failed to find subfolder: " & FoldersArray(i) ' Debugging line
                Exit For
            Else
                Debug.Print "Subfolder found: " & Folder.Name ' Debugging line
            End If
        Next
    End If

    Set GetFolder = Folder
    Exit Function

ErrorHandler:
    Debug.Print "Error finding folder: " & Err.Description ' Debugging line
    Set GetFolder = Nothing
End Function

' Trimite mesaj de confirmare a inregistrarii
Sub SendConfirmationReplyHTML()
    Dim objItem As Object
    Dim objMail As MailItem
    Dim objReply As MailItem
    Dim objExplorer As Explorer
    Dim objSelection As Selection
    Dim isMailItemOpen As Boolean
    Dim response As VbMsgBoxResult
    Dim FolderPath As String
    Dim ConfirmationFolder As MAPIFolder
    Dim TemplatePath As String
    ' Modifica emailul
    ' Specify the folder path for "Rejection" folder (change as needed)
    
    ' VVVVVVVVVV
    FolderPath = "stefa.caravelea@just.ro\Inbox\Confirmate" ' Adjust based on your actual folder path
    
    ' VVVVVVVVVV
    TemplatePath = "C:\Users\scaravelea\OneDrive\Office_Templates\OutlookTemplates\Confirmare.oft" ' Specify the actual path to your template

    ' Find the Rejection folder
    Set ConfirmationFolder = GetFolder(FolderPath)
    isMailItemOpen = False
   
    ' Check if there's an open item
    On Error Resume Next
    If Not Application.ActiveInspector Is Nothing Then
        If TypeName(Application.ActiveInspector.CurrentItem) = "MailItem" Then
            Set objItem = Application.ActiveInspector.CurrentItem
            isMailItemOpen = True
        End If
    End If
    
    ' If no open mail item, check the selection in the main window
    If Not isMailItemOpen Then
        Set objExplorer = Application.ActiveExplorer
        If Not objExplorer Is Nothing Then
            Set objSelection = objExplorer.Selection
            If objSelection.Count > 0 Then
                If TypeName(objSelection.Item(1)) = "MailItem" Then
                    Set objItem = objSelection.Item(1)
                End If
            End If
        End If
    End If
    
    ' Proceed if a MailItem is identified
    If Not objItem Is Nothing Then
        Set objMail = objItem
        ' Create a reply email
        Set objReply = objMail.Reply
        ' Ask user for confirmation before sending the reply
        response = MsgBox("Doriti sa trimiteti un mesaj de confrimare a primirii?", vbQuestion + vbYesNo, "Confirm Send")
        
        If response = vbYes Then
        ' User confirmed, proceed with creating and sending the reply
            ' Create a reply email
            ReplyAndMoveTemplate objMail, TemplatePath, ConfirmationFolder
        Else
            ' User declined, do not send the reply
            ' MsgBox "Operatiunea de confirmare a fost anulata", vbInformation
        End If
    Else
        MsgBox "Deschideti emailul pentru a folosi aceasta operatiune.", vbExclamation
    End If
End Sub

'Trimite mesaj de respingere a inregistrarii pentru neconformitatea atasamanteleor
Sub SendRejectionReplyHTML()
    Dim objItem As Object
    Dim objMail As MailItem
    Dim objReply As MailItem
    Dim objExplorer As Explorer
    Dim objSelection As Selection
    Dim isMailItemOpen As Boolean
    Dim response As VbMsgBoxResult
    Dim FolderPath As String
    Dim RejectionFolder As MAPIFolder
    Dim TemplatePath As String
    isMailItemOpen = False
    

    ' Modifica emailul
    ' Specify the folder path for "Rejection" folder (change as needed)
    
    ' VVVVVVVV
    FolderPath = "stefan.caravelea@just.ro\Inbox\Respinse" ' Adjust based on your actual folder path
         
    ' VVVVVVVVV
    TemplatePath = "C:\Users\scaravelea\OneDrive\Office_Templates\OutlookTemplates\RespingeAtasamenteNeconforme.oft" ' Specify the actual path to your template
         
    ' Find the Rejection folder
    Set RejectionFolder = GetFolder(FolderPath)
    
    ' Check if there's an open item
    If Not Application.ActiveInspector Is Nothing Then
        If TypeName(Application.ActiveInspector.CurrentItem) = "MailItem" Then
            Set objItem = Application.ActiveInspector.CurrentItem
            isMailItemOpen = True
        End If
    End If
    
    ' If no open mail item, check the selection in the main window
    If Not isMailItemOpen Then
        Set objExplorer = Application.ActiveExplorer
        If Not objExplorer Is Nothing Then
            Set objSelection = objExplorer.Selection
            If objSelection.Count > 0 Then
                If TypeName(objSelection.Item(1)) = "MailItem" Then
                    Set objItem = objSelection.Item(1)
                End If
            End If
        End If
    End If
    
    ' Proceed if a MailItem is identified
    If Not objItem Is Nothing Then
        Set objMail = objItem
        ' Ask user for confirmation before sending the reply
        response = MsgBox("Doriti sa trimiteti un mesaj de respingere a inregistrarii?", vbQuestion + vbYesNo, "Confirm Send")
 
        If response = vbYes Then
            ' User confirmed, proceed with creating and sending the reply
            ' Call the reply and move function
            ' ReplyAndMove objMail, reply_body5, reply_subject5, RejectionFolder
            
            ReplyAndMoveTemplate objMail, TemplatePath, RejectionFolder
        Else
            ' User declined, do not send the reply
            ' MsgBox "Operatiunea de respingere a fost anulata", vbInformation
        End If
    Else
        MsgBox "Please select or open an email to use this feature.", vbExclamation
    End If
End Sub

'Trimite mesaj de respingere a inregistrarii in situatia participantilor
Sub SendRejectionReplyHTMLParticipanti()
    Dim objItem As Object
    Dim objMail As MailItem
    Dim objReply As MailItem
    Dim objExplorer As Explorer
    Dim objSelection As Selection
    Dim isMailItemOpen As Boolean
    Dim response As VbMsgBoxResult
    Dim FolderPath As String
    Dim RejectionFolder As MAPIFolder
    Dim TemplatePath As String
    isMailItemOpen = False
    

    ' Modifica emailul
    ' Specify the folder path for "Rejection" folder (change as needed)
    
    ' VVVVVVVV
    FolderPath = "stefan.caravelea@just.ro\Inbox\Respinse" ' Adjust based on your actual folder path
         
    ' VVVVVVVVV
    TemplatePath = "C:\Users\scaravelea\OneDrive\Office_Templates\OutlookTemplates\RespingeParticipanti.oft" ' Specify the actual path to your template
         
    ' Find the Rejection folder
    Set RejectionFolder = GetFolder(FolderPath)
    
    ' Check if there's an open item
    If Not Application.ActiveInspector Is Nothing Then
        If TypeName(Application.ActiveInspector.CurrentItem) = "MailItem" Then
            Set objItem = Application.ActiveInspector.CurrentItem
            isMailItemOpen = True
        End If
    End If
    
    ' If no open mail item, check the selection in the main window
    If Not isMailItemOpen Then
        Set objExplorer = Application.ActiveExplorer
        If Not objExplorer Is Nothing Then
            Set objSelection = objExplorer.Selection
            If objSelection.Count > 0 Then
                If TypeName(objSelection.Item(1)) = "MailItem" Then
                    Set objItem = objSelection.Item(1)
                End If
            End If
        End If
    End If
    
    ' Proceed if a MailItem is identified
    If Not objItem Is Nothing Then
        Set objMail = objItem
        ' Ask user for confirmation before sending the reply
        response = MsgBox("Doriti sa trimiteti un mesaj de respingere a inregistrarii catre participant?", vbQuestion + vbYesNo, "Confirm Send")
 
        If response = vbYes Then
            ' User confirmed, proceed with creating and sending the reply
            ' Call the reply and move function
            ' ReplyAndMove objMail, reply_body5, reply_subject5, RejectionFolder
            
            ReplyAndMoveTemplate objMail, TemplatePath, RejectionFolder
        Else
            ' User declined, do not send the reply
            ' MsgBox "Operatiunea de respingere a fost anulata", vbInformation
        End If
    Else
        MsgBox "Please select or open an email to use this feature.", vbExclamation
    End If
End Sub





