Option Explicit

' Trebuie creat un subfolder in Inbox intitulat "Respinse"
'------------------------------------------------

' Trebuie activate Macro in Outlook
' File > Options > Trust Center > Trust Center Settings... > Macro Settings
' -------------------------------------------------

' Trebuie modificata variabila FolderPath de mai jod cu emailul corespunzator
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
' Se pun doua butoane in grup, respectiv cele doua macrouri de confirmare si infirmare
' Se da Rename, unul "Confirma inregistrarea" cu imaginea o bifa verde,
' Celalalt "Infirma inregistrarea" cu cercul rosu cu X alb ca imagine



' Aceasta functie primeste ca argument email primit, il verifica pe rand
' Daca are linku-ri da reply automat de respingere
' Daca are atasamente mai mari de 10MB da email de respingere
' Daca are alte atasamente decat .docx sau .pdf trimite email de respingere
' In toate cele trei situatii muta emailul in subfolderul "Respinge" din Inbox

Public Sub ProcessNewMail(ByVal Item As MailItem)
    Dim Atmt As Attachment
    Dim TotalAttachmentSize As Long
    Dim FolderPath As String
    Dim RejectionFolder As MAPIFolder

    ' Mesajul de respingere pentru linkuri in email
    Dim reply_body1 As String
    Dim reply_subject1 As String
    reply_body1 = "<p>Buna ziua,</p>" & _
    "<p>Mesajul dumneavoastra a fost respins de la inregistrare din cauza neindeplinirii conditiilor tehnice pentru inregistrare.</p>" & _
    "<p>Au fost identificate urmatoarele nereguli: a) Corpul emailului trebuie sa contina doar text, fara link-uri externe;</p>" & _
    "<p>La adresa de email <a href=""mailto:tr-timis-reg@just.ro"">tr-timis-reg@just.ro</a> se pot trimite inscrisuri respectand urmatoarele reguli :</p>" & _
    "<ul>" & _
    "<li>a) Corpul emailului trebuie sa contina doar text, fara link-uri externe;</li>" & _
    "<li>b) Formatul admis al atasamentelor este .docx sau .pdf;</li>" & _
    "<li>c) Marimea maxima a tuturor atasamentelor este de 10 MB;</li>" & _
    "<li>d) Rezolutia maxima a atasamentelor este 200 dpi;</li>" & _
    "<li>e) Fundalul paginilor scanate trebuie sa fie alb;</li>" & _
    "<li>f) Atasamentele trebuie sa cuprinda doar text lizibil, fara elemente grafice mari (>20% din suprafata paginii).</li>" & _
    "</ul>" & _
    "<p>Plansele foto pot fi depuse doar in mod fizic la registratura sau prin posta/curier.</p>" & _
    "<p><strong>ATENTIE:</strong></p>" & _
    "<p>Trimiterea repetata a unor email-uri sau faxuri care nu indeplinesc aceste conditii poate rezulta in etichetarea de catre sistem ca posta electronica nedorita (SPAM) a adresei de email/numarului de telefon.</p>" & _
    "<p>Nedeschiderea link-urilor de comunicare emise de instanta poate duce la stergerea automata a email-ului din baza de date ca fiind gresit.</p>" & _
    "<p>Email-urile si faxurile referitoare la dosare aflate in curs de judecata trebuie trimise cu cel putin 24 de ore inaintea sedintei de judecata. In caz contrar exista riscul ca documentele sa fie atasate la dosar dupa terminarea sedintei.</p>" & _
    "<p>Pozele, filmarile, inscrisurile si celelalte mijloace de proba stocate pe suport electronic vor fi depuse la dosar pe suport CD sau DVD, nu pe stick-uri USB.</p>" & _
    "<p>Nu capsati cererile si documentele depuse la registratura sau trimise prin posta pentru a evita degradarea documentelor sau a echipamentelor de scanare si pentru a facilita procesul de scanare.</p>" & _
    "<p>Tribunalul Timis, Compartimentul Registratura," & _
    "Piata Tepes Voda nr. 2, Timisoara, Timis, 300055, " & _
    "Email: <a href=""mailto:tr-timis-reg@just.ro"">tr-timis-reg@just.ro</a></p>"
    
    reply_subject1 = "Auto Reply: Emai-l dumneavoastra a fost respins de la inregistrare. Interzise linkuri in email. Nu raspundeti"


    ' Mesajul de respingere pentru depasirea 10 MB in atasamente
    Dim reply_body2 As String
    Dim reply_subject2 As String
    reply_body2 = "<p>Buna ziua,</p>" & _
    "<p>Mesajul dumneavoastra a fost respins de la inregistrare din cauza neindeplinirii conditiilor tehnice pentru inregistrare.</p>" & _
    "<p>Au fost identificate urmatoarele nereguli: c) Marimea maxima a tuturor atasamentelor este de 10 MB </p>" & _
    "<p>La adresa de email <a href=""mailto:tr-timis-reg@just.ro"">tr-timis-reg@just.ro</a> se pot trimite inscrisuri respectand urmatoarele reguli :</p>" & _
    "<ul>" & _
    "<li>a) Corpul emailului trebuie sa contina doar text, fara link-uri externe;</li>" & _
    "<li>b) Formatul admis al atasamentelor este .docx sau .pdf;</li>" & _
    "<li>c) Marimea maxima a tuturor atasamentelor este de 10 MB;</li>" & _
    "<li>d) Rezolutia maxima a atasamentelor este 200 dpi;</li>" & _
    "<li>e) Fundalul paginilor scanate trebuie sa fie alb;</li>" & _
    "<li>f) Atasamentele trebuie sa cuprinda doar text lizibil, fara elemente grafice mari (>20% din suprafata paginii).</li>" & _
    "</ul>" & _
    "<p>Plansele foto pot fi depuse doar in mod fizic la registratura sau prin posta/curier.</p>" & _
    "<p><strong>ATENTIE:</strong></p>" & _
    "<p>Trimiterea repetata a unor email-uri sau faxuri care nu indeplinesc aceste conditii poate rezulta in etichetarea de catre sistem ca posta electronica nedorita (SPAM) a adresei de email/numarului de telefon.</p>" & _
    "<p>Nedeschiderea link-urilor de comunicare emise de instanta poate duce la stergerea automata a email-ului din baza de date ca fiind gresit.</p>" & _
    "<p>Email-urile si faxurile referitoare la dosare aflate in curs de judecata trebuie trimise cu cel putin 24 de ore inaintea sedintei de judecata. In caz contrar exista riscul ca documentele sa fie atasate la dosar dupa terminarea sedintei.</p>" & _
    "<p>Pozele, filmarile, inscrisurile si celelalte mijloace de proba stocate pe suport electronic vor fi depuse la dosar pe suport CD sau DVD, nu pe stick-uri USB.</p>" & _
    "<p>Nu capsati cererile si documentele depuse la registratura sau trimise prin posta pentru a evita degradarea documentelor sau a echipamentelor de scanare si pentru a facilita procesul de scanare.</p>" & _
    "<p>Tribunalul Timis, Compartimentul Registratura," & _
    "Piata Tepes Voda nr. 2, Timisoara, Timis, 300055, " & _
    "Email: <a href=""mailto:tr-timis-reg@just.ro"">tr-timis-reg@just.ro</a></p>"
    
    reply_subject2 = "Auto Reply: Emai-l dumneavoastra a fost respins de la inregistrare. Atasamente prea mari, peste 10 MB. Nu raspundeti"

    ' Mesajul de respingere pentru atasamente diferite de .docx si .pdf
    Dim reply_body3 As String
    Dim reply_subject3 As String
    reply_body3 = "<p>Buna ziua,</p>" & _
    "<p>Mesajul dumneavoastra a fost respins de la inregistrare din cauza neindeplinirii conditiilor tehnice pentru inregistrare.</p>" & _
    "<p>Au fost identificate urmatoarele nereguli: Formatul admis al atasamentelor este .docx sau .pdf; </p>" & _
    "<p>La adresa de email <a href=""mailto:tr-timis-reg@just.ro"">tr-timis-reg@just.ro</a> se pot trimite inscrisuri respectand urmatoarele reguli :</p>" & _
    "<ul>" & _
    "<li>a) Corpul emailului trebuie sa contina doar text, fara link-uri externe;</li>" & _
    "<li>b) Formatul admis al atasamentelor este .docx sau .pdf;</li>" & _
    "<li>c) Marimea maxima a tuturor atasamentelor este de 10 MB;</li>" & _
    "<li>d) Rezolutia maxima a atasamentelor este 200 dpi;</li>" & _
    "<li>e) Fundalul paginilor scanate trebuie sa fie alb;</li>" & _
    "<li>f) Atasamentele trebuie sa cuprinda doar text lizibil, fara elemente grafice mari (>20% din suprafata paginii).</li>" & _
    "</ul>" & _
    "<p>Plansele foto pot fi depuse doar in mod fizic la registratura sau prin posta/curier.</p>" & _
    "<p><strong>ATENTIE:</strong></p>" & _
    "<p>Trimiterea repetata a unor email-uri sau faxuri care nu indeplinesc aceste conditii poate rezulta in etichetarea de catre sistem ca posta electronica nedorita (SPAM) a adresei de email/numarului de telefon.</p>" & _
    "<p>Nedeschiderea link-urilor de comunicare emise de instanta poate duce la stergerea automata a email-ului din baza de date ca fiind gresit.</p>" & _
    "<p>Email-urile si faxurile referitoare la dosare aflate in curs de judecata trebuie trimise cu cel putin 24 de ore inaintea sedintei de judecata. In caz contrar exista riscul ca documentele sa fie atasate la dosar dupa terminarea sedintei.</p>" & _
    "<p>Pozele, filmarile, inscrisurile si celelalte mijloace de proba stocate pe suport electronic vor fi depuse la dosar pe suport CD sau DVD, nu pe stick-uri USB.</p>" & _
    "<p>Nu capsati cererile si documentele depuse la registratura sau trimise prin posta pentru a evita degradarea documentelor sau a echipamentelor de scanare si pentru a facilita procesul de scanare.</p>" & _
    "<p>Tribunalul Timis, Compartimentul Registratura," & _
    "Piata Tepes Voda nr. 2, Timisoara, Timis, 300055, " & _
    "Email: <a href=""mailto:tr-timis-reg@just.ro"">tr-timis-reg@just.ro</a></p>"
    
    reply_subject3 = "Auto Reply: Emai-l dumneavoastra a fost respins de la inregistrare. Atasamente diferite de .docx si .pdf. Nu raspundeti"


    ' Initialize total attachment size
    TotalAttachmentSize = 0

    ' Calculate total size of attachments
    If Item.Attachments.Count > 0 Then
        For Each Atmt In Item.Attachments
            TotalAttachmentSize = TotalAttachmentSize + Atmt.Size
        Next Atmt
    End If

    ' ##################################################################
    ' Modifica emailul
    ' Specify the folder path for "Rejection" folder (change as needed)
    FolderPath = "stefan.caravelea@just.ro\Inbox\Respinse" ' Adjust based on your actual folder path
    
    ' ###################################################################
    
    ' Find the Rejection folder
    Set RejectionFolder = GetFolder(FolderPath)
    Debug.Print RejectionFolder

    ' Check for links, unacceptable attachments, and total attachment size
    If InStr(Item.body, "http://") > 0 Or InStr(Item.body, "https://") > 0 Then
            
        
        ' Call a function or perform an action to reply and move the email
        ReplyAndMove Item, reply_body1, reply_subject1, RejectionFolder
        
   ' Chech if total attachements are less then 10 MB
    ElseIf TotalAttachmentSize > 300000000 Then ' 30 MB in bytes
   

        ' Call a function or perform an action to reply and move the email
        ReplyAndMove Item, reply_body2, reply_subject2, RejectionFolder
        
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
       
            ReplyAndMove Item, reply_body3, reply_subject3, RejectionFolder
            
        End If
    End If
End Sub


Sub ReplyAndMove(MailItem As MailItem, ReplyMessage As String, SubjectLine As String, RejectionFolder As MAPIFolder)
    With MailItem.Reply
        .HTMLBody = ReplyMessage & .HTMLBody
        .Subject = SubjectLine
        .Send
    End With
    With MailItem
        .UnRead = False
        .Move RejectionFolder
    End With
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

' Trimite mesaj de respingere a inregistrarii
Sub SendConfirmationReplyHTML()
    Dim objItem As Object
    Dim objMail As MailItem
    Dim objReply As MailItem
    Dim objExplorer As Explorer
    Dim objSelection As Selection
    Dim isMailItemOpen As Boolean
    Dim response As VbMsgBoxResult
    
    isMailItemOpen = False
    
    
        ' Mesajul care trebuie trimis
    Dim reply_body4 As String
     reply_body4 = "<p>Buna ziua,</p>" & _
    "<p>Confirmam receptionarea mesajului dumneavoastra.</p>" & _
    "<p>Continutul mesajului urmeaza va fi tiparit, inregistrat si depus la dosar, dupa parcurgerea traseului administrativ (24h de la înregistrare).</p>" & _
    "<p>Nu este necesar sa mai trimiteti documentele si pe alta cale de comunicare (fax, posta, etc ).</p>" & _
    "<p>Va reamintim ca la adresa de email <a href=""mailto:tr-timis-reg@just.ro"">tr-timis-reg@just.ro</a> se pot trimite inscrisuri respectand urmatoarele reguli :</p>" & _
    "<ul>" & _
    "<li>a) Corpul emailului trebuie sa contina doar text, fara link-uri externe;</li>" & _
    "<li>b) Formatul admis al atasamentelor este .docx sau .pdf;</li>" & _
    "<li>c) Marimea maxima a tuturor atasamentelor este de 10 MB;</li>" & _
    "<li>d) Rezolutia maxima a atasamentelor este 200 dpi;</li>" & _
    "<li>e) Fundalul paginilor scanate trebuie sa fie alb;</li>" & _
    "<li>f) Atasamentele trebuie sa cuprinda doar text lizibil, fara elemente grafice mari (>20% din suprafata paginii).</li>" & _
    "</ul>" & _
    "<p>Plansele foto pot fi depuse doar in mod fizic la registratura sau prin posta/curier.</p>" & _
    "<p><strong>ATENTIE:</strong></p>" & _
    "<p>Trimiterea repetata a unor email-uri sau faxuri care nu indeplinesc aceste conditii poate rezulta in etichetarea de catre sistem ca posta electronica nedorita (SPAM) a adresei de email/numarului de telefon.</p>" & _
    "<p>Nedeschiderea link-urilor de comunicare emise de instanta poate duce la stergerea automata a email-ului din baza de date ca fiind gresit.</p>" & _
    "<p>Email-urile si faxurile referitoare la dosare aflate in curs de judecata trebuie trimise cu cel putin 24 de ore inaintea sedintei de judecata. In caz contrar exista riscul ca documentele sa fie atasate la dosar dupa terminarea sedintei.</p>" & _
    "<p>Pozele, filmarile, inscrisurile si celelalte mijloace de proba stocate pe suport electronic vor fi depuse la dosar pe suport CD sau DVD, nu pe stick-uri USB.</p>" & _
    "<p>Nu capsati cererile si documentele depuse la registratura sau trimise prin posta pentru a evita degradarea documentelor sau a echipamentelor de scanare si pentru a facilita procesul de scanare.</p>" & _
    "<p>Tribunalul Timis, Compartimentul Registratura," & _
    "Piata Tepes Voda nr. 2, Timisoara, Timis, 300055, " & _
    "Email: <a href=""mailto:tr-timis-reg@just.ro"">tr-timis-reg@just.ro</a></p>"

    
    
    
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
            Set objReply = objMail.Reply
            With objReply
                .HTMLBody = reply_body4 & .HTMLBody
                .Subject = "Confirmam primirea si inregistrarea documentelor " & objMail.Subject
                .Send
            End With
        Else
            ' User declined, do not send the reply
            ' MsgBox "Operatiunea de confirmare a fost anulata", vbInformation
        End If
    Else
        MsgBox "Deschideti emailul pentru a folosi aceasta operatiune.", vbExclamation
    End If
End Sub

'Trimite mesaj de respingere a inregistrarii
Sub SendRejectionReplyHTML()
    Dim objItem As Object
    Dim objMail As MailItem
    Dim objReply As MailItem
    Dim objExplorer As Explorer
    Dim objSelection As Selection
    Dim isMailItemOpen As Boolean
    Dim response As VbMsgBoxResult
    
    isMailItemOpen = False
    
        ' Mesajul de respingere de la inregistrare
    Dim reply_body5 As String
     reply_body5 = "<p>Buna ziua,</p>" & _
    "<p>Mesajul dumneavoastra a fost respins de la inregistrare din cauza neindeplinirii conditiilor tehnice pentru inregistrare.</p>" & _
    "<p>Au fost identificate una din urmatoarele nereguli: d) Rezolutia maxima a atasamentelor este 200 dpi; e) Fundalul paginilor scanate trebuie sa fie alb; f) Atasamentele trebuie sa cuprinda doar text lizibil, fara elemente grafice mari (>20% din suprafata paginii).</p>" & _
    "<p>La adresa de email <a href=""mailto:tr-timis-reg@just.ro"">tr-timis-reg@just.ro</a> se pot trimite inscrisuri respectand urmatoarele reguli :</p>" & _
    "<ul>" & _
    "<li>a) Corpul emailului trebuie sa contina doar text, fara link-uri externe;</li>" & _
    "<li>b) Formatul admis al atasamentelor este .docx sau .pdf;</li>" & _
    "<li>c) Marimea maxima a tuturor atasamentelor este de 10 MB;</li>" & _
    "<li>d) Rezolutia maxima a atasamentelor este 200 dpi;</li>" & _
    "<li>e) Fundalul paginilor scanate trebuie sa fie alb;</li>" & _
    "<li>f) Atasamentele trebuie sa cuprinda doar text lizibil, fara elemente grafice mari (>20% din suprafata paginii).</li>" & _
    "</ul>" & _
    "<p>Plansele foto pot fi depuse doar in mod fizic la registratura sau prin posta/curier.</p>" & _
    "<p><strong>ATENTIE:</strong></p>" & _
    "<p>Trimiterea repetata a unor email-uri sau faxuri care nu indeplinesc aceste conditii poate rezulta in etichetarea de catre sistem ca posta electronica nedorita (SPAM) a adresei de email/numarului de telefon.</p>" & _
    "<p>Nedeschiderea link-urilor de comunicare emise de instanta poate duce la stergerea automata a email-ului din baza de date ca fiind gresit.</p>" & _
    "<p>Email-urile si faxurile referitoare la dosare aflate in curs de judecata trebuie trimise cu cel putin 24 de ore inaintea sedintei de judecata. In caz contrar exista riscul ca documentele sa fie atasate la dosar dupa terminarea sedintei.</p>" & _
    "<p>Pozele, filmarile, inscrisurile si celelalte mijloace de proba stocate pe suport electronic vor fi depuse la dosar pe suport CD sau DVD, nu pe stick-uri USB.</p>" & _
    "<p>Nu capsati cererile si documentele depuse la registratura sau trimise prin posta pentru a evita degradarea documentelor sau a echipamentelor de scanare si pentru a facilita procesul de scanare.</p>" & _
    "<p>Tribunalul Timis, Compartimentul Registratura," & _
    "Piata Tepes Voda nr. 2, Timisoara, Timis, 300055, " & _
    "Email: <a href=""mailto:tr-timis-reg@just.ro"">tr-timis-reg@just.ro</a></p>"

    
    
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
            ' Create a reply email
            Set objReply = objMail.Reply
            With objReply
                .HTMLBody = reply_body5 & .HTMLBody
                .Subject = "Auto Reply: Email-ul dumneavoastra a fost respins de la inregistrare. Nu dati reply. " & objMail.Subject
                .Send
            End With
        Else
            ' User declined, do not send the reply
            ' MsgBox "Operatiunea de respingere a fost anulata", vbInformation
        End If
    Else
        MsgBox "Please select or open an email to use this feature.", vbExclamation
    End If
End Sub


