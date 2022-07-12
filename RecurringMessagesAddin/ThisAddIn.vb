Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Tools
Imports System.IO
Imports System.Text
Imports System.Configuration

Public Class ThisAddIn
    Public holidays = ConfigurationManager.AppSettings("holidays").Split(",")
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Dim startupMessage = MsgBox("Smart Outlook Assistant is not active. " & vbCr & "Do you want to turn it on ?", vbYesNo, "Managing parameters...")
        If startupMessage = vbYes Then
            Globals.Ribbons.Ribbon1.ModeOnOff.Checked = True
            MsgBox("Addin is on now.")
        Else
            MsgBox("Smart Outlook Assistant is not active. " & vbCr & "You can turn it on anytime from the ribbon.")
        End If
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
    End Sub

    Private Sub createTrace(dosya As String)
        Dim fs As Object
        Dim myFile = "C:\temp\" & Format(Now(), "yyMMdd") & "_" & Left(dosya, 40) & ".txt"
        fs = CreateObject("Scripting.FileSystemObject")
        Dim a As Object
        a = fs.CreateTextFile(myFile, True)
    End Sub
    Function isSentBefore(item As String) As Boolean
        Dim myFile = "C:\temp\" & Format(Now(), "yyMMdd") & "_" & Left(item, 40) & ".txt"
        If FileFolderExists(myFile) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Application_Reminder(Item As Object) Handles Application.Reminder
        Dim objMsg As MailItem
        Dim olMailItem As OlItemType = Nothing
        objMsg = Application.CreateItem(olMailItem)

        If Globals.Ribbons.Ribbon1.ModeOnOff.Checked = True Then
            If TypeOf Item Is AppointmentItem Then

                If InStr(1, Item.Categories, "SendEmailWithAttachment") <> 0 Then
                    If isSentBefore(Item.Subject) = False Then
                        Call Sendfile(Item.Location, Item.Subject, Item.Body)
                    End If
                End If

                If InStr(1, Item.Categories, "CheckProofAndInformOnlyAbsence") <> 0 Then
                    Call CheckProofAndInformOnlyAbsence(Item.Location, Item.Subject, Item.Body)
                Else : End If

                If InStr(1, Item.Categories, "CheckFileProof") <> 0 Then
                    If isSentBefore(Item.Subject) = False Then
                        Call CheckProofFile(Item.Location, Item.Subject, Item.Body)
                    End If
                Else : End If

                If InStr(1, Item.Categories, "SendRecurringEmail") <> 0 Then
                    If isSentBefore(Item.Subject) = False Then
                        Call Sendmessage(Item.Location, Item.Subject, Item.Body)
                    End If
                End If

                'If InStr(1, Item.Subject, "testing") <> 0 Then
                '    'If isSentBefore(Item.Subject) = False Then
                '    'Call Screenall(Item.Subject)
                '    'MsgBox(Reppdate())
                '    Call CheckProofAndInformOnlyAbsence(Item.Location, Item.Subject, Item.Body)

                '    'Call Sendmessage(Item.Location, Item.Subject, Item.Body)
                '    'End If
                'End If

            End If
        End If

    End Sub

    Function Reppdate() As Date
        Dim wd_correction
        Dim yest, checkDate As Date
        Dim yesterday As Date
        Dim tatiller = holidays
        Dim j = 0
        If Weekday(Now()) = 2 Then
            wd_correction = -3
        Else
            If Weekday(Now()) = 1 Then
                wd_correction = -2
            Else
                wd_correction = -1
            End If
        End If
        yesterday = DateAdd("d", wd_correction, Now())
        yest = Format(DateAdd("d", j, yesterday), "dd.MM.yyyy")
        For i = 0 To UBound(tatiller)
            j = 0
            checkDate = Format(CDate(tatiller(i)), "dd.MM.yyyy")
            If yest > checkDate Then
                GoTo son
            Else
                If yest = checkDate Then
                    If Weekday(yest) = 2 Then
                        j = j - 3
                    Else : j = j - 1
                    End If
                Else
                    If j < -1 Then
                        Exit For
                    Else : End If
                End If
                yest = Format(DateAdd("d", j, yest), "dd.MM.yyyy")
            End If
        Next i
son:
        Reppdate = yest
    End Function

    Public Function FileFolderExists(ByVal StrFullPath As String) As Boolean
        'Macro Purpose: Check if a file or folder exists
        On Error GoTo EarlyExit
        If Not Dir(StrFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
EarlyExit:
        On Error GoTo 0
    End Function
    Public Function FileFolderExists_Multi(ByVal ss As String) As Boolean
        'Macro Purpose: Check if a file or folder exists
        On Error GoTo EarlyExit
        ''Dim arr As Variant
        ''arr = Split(ss, "~")
        Dim arr = Split(ss, "~")
        Dim n = 0
        Do While n < UBound(arr)
            If Dir(arr(n), vbDirectory) = vbNullString Then
                FileFolderExists_Multi = False
                Exit Function
            Else : End If
            n = n + 1
        Loop
        FileFolderExists_Multi = True
EarlyExit:
        On Error GoTo 0
    End Function


    Sub CheckProofFile(filePath As String, file As String, recipients As String)
        ' CheckProofFile tries to send the requested file to recipients.
        ' This macro is intended to run multiple times in a day. Therefore if send operation is success, leaves a trace for future runs to avoid duplicate messages.
        ' If file does not exist, informs people that file not found and therefore related task may not be completed yet.
        file = EvaluateFolder(file)
        Dim objMsg As MailItem
        Dim olMailItem As OlItemType = Nothing
        objMsg = Application.CreateItem(olMailItem)
        recipients = Replace(recipients, "~", "@")
        Dim vamsg = "Attention, " & Left(file, Len(file) - 4) & " file not found, task may not be completed !"
        Dim vamsg2 = Left(file, Len(file) - 4) & " task is completed.  Click to see proof / screenshot."
        Dim fullFileName = EvaluateFolder(filePath) & file
        Dim myFile = "C:\temp\" & Format(Now(), "yyMMdd") & "\" & file & ".txt"
        If FileFolderExists(fullFileName) = False Then
            Dim msg = vamsg
        Else
            If FileFolderExists(myFile) = True Then
                Exit Sub
            Else : End If

            objMsg.To = recipients
            objMsg.Subject = vamsg2
            objMsg.Attachments.Add(fullFileName)
            objMsg.Body = "Please open screenshot attached and check it."
            objMsg.Send()
            objMsg = Nothing

            Call createTrace(file)
            Exit Sub
        End If
        objMsg.To = recipients
        objMsg.Subject = vamsg
        objMsg.Body = "File " & fullFileName & " not found. " & vbNewLine & vbNewLine & file & " task may not be completed !"
        objMsg.Send()
        objMsg = Nothing

    End Sub

    Sub Sendfile(fileName As String, subject As String, recipients As String)
        ' Sendfile tries to send the requested file to recipients.
        ' If file does not exist, sends an e-mail saying file not found.  
        subject = EvaluateFolder(subject)
        Dim objMsg As MailItem
        Dim link As Integer
        Dim olMailItem As OlItemType = Nothing
        Dim myFile = "C:\temp\" & Format(Now(), "yyMMdd") & "\" & subject & ".txt"
        objMsg = Application.CreateItem(olMailItem)

        If Left(fileName, 5) = "Link:" Then
            fileName = Right(fileName, Len(fileName) - 5)
            link = 1
        Else
            fileName = fileName
            link = 0
        End If

        Dim vamsg = "Attention, " & subject & " file not found and therefore could not be sent !"
        Dim vamsg2 = subject & " file is attached."
        Dim fileNameEvaluated = EvaluateFolder(fileName)
        Dim bodyLink = "<a href=""" & fileNameEvaluated & """ >" & vamsg2 & "</a>"
        recipients = Replace(recipients, "~", "@")
        If FileFolderExists(fileNameEvaluated) = False Then
        Else
            If FileFolderExists(myFile) = True Then
                Exit Sub
            Else : End If

            objMsg.To = recipients
            objMsg.Subject = vamsg2
            If link = 0 Then
                objMsg.Attachments.Add(fileNameEvaluated)
            Else
            End If
            'Giving link can be unnecessary in cases where recipients are outside of organization/network..
            objMsg.HTMLBody = bodyLink
            objMsg.Send()
            objMsg = Nothing
            Call createTrace(subject)
            Exit Sub
        End If
        objMsg.To = recipients
        objMsg.Subject = vamsg
        objMsg.Body = "File " & fileNameEvaluated & " not found and therefore could not be sent !"
        objMsg.Send()
        objMsg = Nothing
    End Sub

    Public Function EvaluateFolder(fileName As String)
        Dim lastWorkDay = Reppdate()
        Dim gib_date As Date

        If Weekday(Now()) = 2 Then
            gib_date = DateAdd("d", -2, Now())
        Else
            gib_date = Now()
        End If

        Dim fileName2 = Replace(fileName, "[yyyyMM]", Format(lastWorkDay, "yyyyMM"))
        Dim fileName3 = Replace(fileName2, "[yyMMdd]", Format(lastWorkDay, "yyMMdd"))
        Dim fileName4 = Replace(fileName3, "[yyyyMMdd_today]", Format(Now(), "yyyyMMdd"))
        Dim fileName5 = Replace(fileName4, "[dd]", Format(lastWorkDay, "dd"))
        Dim fileName6 = Replace(fileName5, "[suffix]", Suffix())
        Dim fileName7 = Replace(fileName6, "[dd_MM_yyyy]", Format(lastWorkDay, "dd_MM_yyyy"))
        Dim fileName8 = Replace(fileName7, "[yyyy_MM_dd_today]", Format(Now(), "yyyy_MM_dd"))
        Dim fileName9 = Replace(fileName8, "[yyyy_MM]", Format(lastWorkDay, "yyyy_MM"))
        Dim fileName10 = Replace(fileName9, "[yyyy_MM_dd_FTP]", Format(gib_date, "yyyy_MM_dd"))
        Dim fileName11 = Replace(fileName10, "[dd-MM-yyyy]", Format(lastWorkDay, "dd-MM-yyyy"))
        EvaluateFolder = fileName11
    End Function
    Public Function Suffix() As String
        Dim saat = CDbl(Format(Now(), "HHmm"))
        If saat > 1805 Then
            Suffix = "_1805"
        ElseIf saat > 1735 Then
            Suffix = "_1735"
        ElseIf saat > 1705 Then
            Suffix = "_1705"
        ElseIf saat > 1635 Then
            Suffix = "_1635"
        Else
            Suffix = "_1605"
        End If
    End Function
    Function YYYYAAGG(XX)
        YYYYAAGG = Format(XX, "YYYYmmdd")
    End Function

    Sub CheckProofAndInformOnlyAbsence(filePath As String, absenceMessageToGive As String, recipients As String)
        ' CheckProofAndInformOnlyAbsence looks for file and sends e-mail only if it fails to find. 
        Dim fileName = Path.GetFileName(filePath)
        fileName = EvaluateFolder(fileName)
        Dim objMsg As MailItem
        Dim olMailItem As OlItemType = Nothing
        objMsg = Application.CreateItem(olMailItem)
        recipients = Replace(recipients, "~", "@")
        Dim filePathEvaluated = EvaluateFolder(filePath)
        If FileFolderExists(filePathEvaluated) = False Then
        Else
            Exit Sub
        End If
        objMsg.To = recipients
        objMsg.Subject = absenceMessageToGive
        objMsg.Body = "File " & fileName & " not found." & vbNewLine & vbNewLine & absenceMessageToGive
        objMsg.Send()
        objMsg = Nothing

    End Sub


    Sub Sendmessage(bodyText As String, subject As String, recipients As String)
        ' Sendmessage just sends plain message without attachments.  Created for messages need to be sent periodically.
        subject = EvaluateFolder(subject)
        bodyText = EvaluateFolder(bodyText)
        Dim lastWorkDay = Reppdate()
        Dim objMsg As MailItem
        Dim olMailItem As OlItemType = Nothing
        Dim myFile = "C:\temp\" & Format(Now(), "yyMMdd") & "\" & subject & ".txt"
        If FileFolderExists(myFile) = True Then
            Exit Sub
        Else : End If

        objMsg = Application.CreateItem(olMailItem)

        recipients = Replace(recipients, "~", "@")
        Dim newSubject = Replace(subject, "[REPORTING_MONTH]", Format(lastWorkDay, "yyyy-MM"))
        Dim newBody = Replace(bodyText, "[REPORTING_MONTH]", Format(lastWorkDay, "yyyy-MM"))

        objMsg.To = recipients
        objMsg.Subject = newSubject
        objMsg.Body = "Hello, " & vbNewLine & newBody & vbNewLine & "Thanks."
        objMsg.Send()
        objMsg = Nothing
        Call createTrace(subject)

    End Sub



    Function EdevletDailyCheck(Liste As String) As Boolean
        Dim myNamespace As Microsoft.Office.Interop.Outlook.NameSpace
        Dim olFolderInbox As OlDefaultFolders = Nothing
        Dim reportDate, reportDate_f As Date
        reportDate = Reppdate()

        Dim dosyaadi1_ok = "C:\Temp\" & Format(reportDate, "yyyyMM") & "\Daily\" & Format(reportDate, "dd") & "\" & Liste & "_ok.txt"
        If FileFolderExists(dosyaadi1_ok) = True Then
            EdevletDailyCheck = True
            Exit Function
        Else : End If

        If Microsoft.VisualBasic.DateAndTime.Day(reportDate) < 10 Then
            reportDate_f = Right(Format(reportDate, "dd.MM.yyyy"), Len(Format(reportDate, "dd.MM.yyyy")) - 1)
        Else
            reportDate_f = Format(reportDate, "dd.MM.yyyy")
        End If
        Dim arananbody As String
        arananbody = Format(reportDate_f, "D.MM.yyyy") & " 00:00:00 Dosya Türü: Ticari " & Liste
        Dim ProfileName = CurrentUserEmailAddress()
        Dim MailBoxName = ProfileName

        myNamespace = Application.GetNamespace("MAPI")
        'myContacts = myNamespace.GetDefaultFolder(olFolderInbox).Items

        Dim outApp As Object
        'Create an Outlook session
        outApp = CreateObject("Outlook.Application")

        'Dim FmtToday = Format(DateValue(Now()), "ddddd h:nn AMPM")
        'Dim FmtToday = Format(DateValue(Now()), "dd.MM.yyyy hh:mm tt")
        Dim FmtToday = Format(DateValue(Now()), "dd.MM.yyyy hh:mm")
        Dim outSubFolder = myNamespace.Session.Folders(MailBoxName).Folders("Edevlet")
        Dim myItems = outSubFolder.Items.Restrict("[SenderName]='edevlet@edevlet.org' and [ReceivedTime] > '" & FmtToday & "'")
        If myItems.Count > 0 Then
            Dim i = 1
            While i <= myItems.Count
                'DoEvents
                Dim outItem = myItems(i)
                'Does the findText occur in this email's body text?
                Dim outMail = outItem
                If InStr(1, outMail.Body, Format(reportDate_f, "D.MM.yyyy") & " 00:00:00 Dosya Türü: Ticari " & Liste, vbTextCompare) > 0 Then
                    EdevletDailyCheck = True
                    Dim fs = CreateObject("Scripting.FileSystemObject")
                    Dim a = fs.CreateTextFile(dosyaadi1_ok, True)
                    '                    MsgBox reppdate() & " kapanan liste " & outSubFolder.Name & " klasöründe bulundu"
                    Exit Function
                Else
                End If
                i = i + 1
            End While
        Else : End If
        EdevletDailyCheck = False

        If Globals.Ribbons.Ribbon1.Mode_debug.Checked = True Then
            Call Debuglog("EdevletCheck", "arananbody:" & arananbody & vbNewLine & "raportarihi_f:" & reportDate_f & vbNewLine & "MailBoxName: " & MailBoxName & vbNewLine & "ProfileName: " & ProfileName & vbNewLine & "FmtToday: " & FmtToday & vbNewLine & "ProfileName: " & ProfileName)

            'Dim objMsg As MailItem
            'Dim olMailItem As OlItemType = Nothing
            'objMsg = Application.CreateItem(olMailItem)
            'objMsg.To = "Ozgun.Senyuva@rabobank.com"
            'objMsg.Subject = "EDEVLET FONK DEBUGGING"
            'objMsg.Body = "arananbody:" & arananbody & vbNewLine & "raportarihi_f:" & raportarihi_f & vbNewLine & "MailBoxName: " & MailBoxName & vbNewLine & "ProfileName: " & ProfileName & vbNewLine & "FmtToday: " & FmtToday & vbNewLine & "ProfileName: " & ProfileName
            ''& "outSubFolder: " & outSubFolder & vbNewLine 
            'objMsg.Send()
            'objMsg = Nothing
        End If


    End Function

    Function CurrentUserEmailAddress() As String

        Dim outApp As Object, outSession As Object
        'Create an Outlook session
        outApp = CreateObject("Outlook.Application")

        'Check if session is created
        If outApp Is Nothing Then
            CurrentUserEmailAddress = "Cannot create Microsoft Outlook session."
            CurrentUserEmailAddress = "Not found"
            Exit Function
        End If

        'Set a NameSpace object variable with .Session property (same as .GetNamespace("MAPI"), to access existing Outlook items, and get
        'current user name
        outSession = outApp.Session.CurrentUser

        'Get current user email address
        CurrentUserEmailAddress = outSession.AddressEntry.GetExchangeUser().PrimarySmtpAddress

        outApp = Nothing
    End Function

    Private Sub Application_NewMailEx(EntryIDCollection As String) Handles Application.NewMailEx
        Dim fileName, keysubject, keybody, keysender As String
        Dim raportarihi As Date
        raportarihi = Reppdate()

        keybody = Format(raportarihi, "d.MM.yyyy") & " 00:00:00 File Type: List of additions"
        keysender = "edevlet@edevlet.org"
        keysubject = "File Upload Result"
        fileName = "C:\Temp\" & Format(raportarihi, "yyyyMM") & "\Daily\" & Format(raportarihi, "dd") & "\" & "ListAdditions_ok.txt"
        Call Keymailcheck(EntryIDCollection, keysender, keysubject, keybody, fileName)

        keybody = ""
        keysender = "ozgun.senyuva@gmail.com"
        keysubject = "Daily Position Followup - Date: " & Format(raportarihi, "dd.MM.yyyy")
        fileName = "C:\Temp\" & Format(raportarihi, "yyyyMM") & "\Daily\" & Format(raportarihi, "dd") & "\" & "FXPosFollowup_ok.txt"
        Call Keymailcheck(EntryIDCollection, keysender, keysubject, keybody, fileName)

    End Sub
    Sub Keymailcheck(keyid As String, keysender As String, keysubject As String, keybody As String, fileName As String)
        Dim mai As Object
        Dim intInitial As Integer
        Dim intFinal As Integer
        Dim strEntryId As String
        Dim intLength As Integer

        Call Debuglog("keymailcheck", "keysender:" & keysender & "keysubject:" & keysubject & vbNewLine & "keybody:" & keybody & vbNewLine & "filename: " & fileName & vbNewLine & "keyid: " & keyid & vbNewLine)

        intInitial = 1
        intLength = Len(keyid)
        intFinal = InStr(intInitial, keyid, ",")

        Do While intFinal <> 0
            strEntryId = Strings.Mid(keyid, intInitial, (intLength - intInitial))
            mai = Application.Session.GetItemFromID(strEntryId)

            If keybody <> "" Then
                If mai.Subject = keysubject And InStr(1, mai.Body, keybody) <> 0 And mai.SenderEmailAddress = keysender Then
                    Dim fs As Object
                    fs = CreateObject("Scripting.FileSystemObject")
                    Dim a As Object
                    a = fs.CreateTextFile(fileName, True)
                End If
            Else
                If mai.Subject = keysubject And mai.SenderEmailAddress = keysender Then
                    Dim fs As Object
                    fs = CreateObject("Scripting.FileSystemObject")
                    Dim a As Object
                    a = fs.CreateTextFile(fileName, True)
                End If
            End If
        Loop
        strEntryId = Strings.Mid(keyid, intInitial, (intLength - intInitial) + 1)
        mai = Application.Session.GetItemFromID(strEntryId)
        If keybody <> "" Then
            If mai.Subject = keysubject And InStr(1, mai.Body, keybody) <> 0 And mai.SenderEmailAddress = keysender Then
                Dim fs As Object
                fs = CreateObject("Scripting.FileSystemObject")
                Dim a As Object
                a = fs.CreateTextFile(fileName, True)
            End If
        Else
            If mai.Subject = keysubject And mai.SenderEmailAddress = keysender Then
                Dim fs As Object
                fs = CreateObject("Scripting.FileSystemObject")
                Dim a As Object
                a = fs.CreateTextFile(fileName, True)
            End If
        End If
    End Sub
    Sub Debuglog(job As String, log As String)
        Dim fileName As String
        If Globals.Ribbons.Ribbon1.Mode_debug.Checked = True Then
            fileName = "c:\temp\" & job & "_" & Format(Now(), "yyMMdd") & ".txt"
            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter(fileName, True)
            file.WriteLine(Format(Now(), "hhmmss") & "-" & log)
            file.Close()
        End If
    End Sub

End Class
