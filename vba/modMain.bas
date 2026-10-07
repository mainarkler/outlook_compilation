Attribute VB_Name = "modMain"
Option Explicit

Public Sub CompileQuestions()
    On Error GoTo EH
    Dim targetDate As Date
    If Not AskTargetDate(targetDate) Then Exit Sub
    Dim olFolder As Object
    Set olFolder = GetQuestionsFolder()
    If olFolder Is Nothing Then
        MsgBox "Не найдена папка Outlook: " & QUESTIONS_FOLDER_NAME, vbExclamation
        Exit Sub
    End If
    Dim messages As Collection
    Set messages = GetQuestionMessages(olFolder, targetDate)
    If messages.Count = 0 Then
        MsgBox "В папке """ & QUESTIONS_FOLDER_NAME & """ не найдено вопросов на " & Format$(targetDate, "dd/mm/yyyy") & ".", vbInformation
        Exit Sub
    End If
    EnsureFolderExists OUTPUT_FOLDER_PATH
    Dim outputPath As String
    outputPath = BuildOutputPath(targetDate)
    FileCopy WORD_TEMPLATE_PATH, outputPath
    Dim wdApp As Object, wdDoc As Object
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(outputPath)
    InsertQuestions wdDoc, messages
    wdDoc.Save
    wdDoc.Close
    wdApp.Quit
    MsgBox "Документ сформирован:" & vbCrLf & outputPath, vbInformation
    Exit Sub
EH:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=True
    If Not wdApp Is Nothing Then wdApp.Quit
    On Error GoTo 0
    MsgBox "Ошибка: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Function AskTargetDate(ByRef resultDate As Date) As Boolean
    Dim s As String
    s = InputBox("Введите дату вопросов в формате ДД/ММ/ГГГГ:", "Формирование вопросов", Format$(Date, "dd/mm/yyyy"))
    If Len(Trim$(s)) = 0 Then Exit Function
    If Not IsDate(s) Then
        MsgBox "Не удалось распознать дату: " & s, vbExclamation
        Exit Function
    End If
    resultDate = DateValue(s)
    AskTargetDate = True
End Function

Private Function GetQuestionsFolder() As Object
    On Error GoTo EH
    Dim ns As Object, root As Object, candidate As Object
    Set ns = Application.GetNamespace("MAPI")
    For Each root In ns.Folders
        Set candidate = FindFolderRecursive(root, QUESTIONS_FOLDER_NAME)
        If Not candidate Is Nothing Then
            Set GetQuestionsFolder = candidate
            Exit Function
        End If
    Next root
EH:
End Function

Private Function FindFolderRecursive(ByVal parentFolder As Object, ByVal wantedName As String) As Object
    On Error GoTo EH
    If StrComp(parentFolder.Name, wantedName, vbTextCompare) = 0 Then
        Set FindFolderRecursive = parentFolder
        Exit Function
    End If
    Dim subFolder As Object
    For Each subFolder In parentFolder.Folders
        Set FindFolderRecursive = FindFolderRecursive(subFolder, wantedName)
        If Not FindFolderRecursive Is Nothing Then Exit Function
    Next subFolder
EH:
End Function

Private Function GetQuestionMessages(ByVal folder As Object, ByVal targetDate As Date) As Collection
    Dim result As New Collection
    Dim items As Object, item As Object
    Set items = folder.Items
    On Error Resume Next
    items.Sort "[ReceivedTime]", False
    On Error GoTo 0
    For Each item In items
        If IsQuestionMail(item, targetDate) Then result.Add item
    Next item
    Set GetQuestionMessages = result
End Function

Private Function IsQuestionMail(ByVal item As Object, ByVal targetDate As Date) As Boolean
    On Error GoTo EH
    If item.Class <> 43 Then Exit Function
    Dim subject As String, expectedDate As String
    subject = CStr(item.Subject)
    expectedDate = Format$(targetDate, "dd/mm/yyyy")
    If InStr(1, subject, SUBJECT_MARKER, vbTextCompare) = 0 Then Exit Function
    If InStr(1, subject, expectedDate, vbTextCompare) = 0 Then Exit Function
    IsQuestionMail = True
    Exit Function
EH:
    IsQuestionMail = False
End Function

Private Sub InsertQuestions(ByVal wdDoc As Object, ByVal messages As Collection)
    Dim insertRange As Object
    Set insertRange = wdDoc.Paragraphs(TEMPLATE_KEEP_PARAGRAPHS).Range
    insertRange.Collapse 0
    Dim appendixTables As Collection
    Set appendixTables = New Collection
    Dim questionNumber As Long, mail As Object
    questionNumber = 0
    For Each mail In messages
        questionNumber = questionNumber + 1
        InsertQuestionHeader insertRange, questionNumber, GetSenderName(mail)
        InsertMailBodyWithoutTables insertRange, mail
        CollectMailTables appendixTables, mail, questionNumber
        insertRange.InsertAfter vbCrLf
        insertRange.Collapse 0
    Next mail
    If appendixTables.Count > 0 Then InsertAppendices wdDoc, appendixTables
End Sub

Private Sub InsertQuestionHeader(ByVal targetRange As Object, ByVal questionNumber As Long, ByVal senderName As String)
    Dim text As String
    text = QUESTION_LABEL & " " & CStr(questionNumber) & vbCrLf & RAISED_BY_LABEL & " " & senderName & vbCrLf
    targetRange.InsertAfter text
    targetRange.Collapse 0
End Sub

Private Sub InsertMailBodyWithoutTables(ByVal targetRange As Object, ByVal mail As Object)
    On Error GoTo Fallback
    Dim editor As Object, p As Object, r As Object
    Set editor = mail.GetInspector.WordEditor
    For Each p In editor.Paragraphs
        Set r = p.Range
        If Not IsRangeInsideTable(r) Then
            targetRange.FormattedText = r.FormattedText
            targetRange.Collapse 0
        End If
    Next p
    Exit Sub
Fallback:
    targetRange.InsertAfter mail.Body & vbCrLf
    targetRange.Collapse 0
End Sub

Private Function IsRangeInsideTable(ByVal r As Object) As Boolean
    On Error GoTo NotInTable
    IsRangeInsideTable = CBool(r.Information(12))
    Exit Function
NotInTable:
    IsRangeInsideTable = False
End Function

Private Sub CollectMailTables(ByVal tables As Collection, ByVal mail As Object, ByVal questionNumber As Long)
    On Error GoTo EH
    Dim editor As Object, t As Object, entry As Collection
    Set editor = mail.GetInspector.WordEditor
    For Each t In editor.Tables
        Set entry = New Collection
        entry.Add questionNumber
        entry.Add t.Range.FormattedText
        tables.Add entry
    Next t
EH:
End Sub

Private Sub InsertAppendices(ByVal wdDoc As Object, ByVal appendixTables As Collection)
    Dim r As Object, entry As Variant, questionNumber As Long
    Set r = wdDoc.Content
    r.Collapse 0
    r.InsertAfter vbCrLf & vbCrLf & APPENDICES_LABEL & vbCrLf & vbCrLf
    r.Collapse 0
    For Each entry In appendixTables
        questionNumber = CLng(entry(1))
        r.InsertAfter APPENDIX_FOR_QUESTION_LABEL & " " & CStr(questionNumber) & vbCrLf
        r.Collapse 0
        r.FormattedText = entry(2)
        r.Collapse 0
        r.InsertAfter vbCrLf
        r.Collapse 0
    Next entry
End Sub

Private Function GetSenderName(ByVal mail As Object) As String
    On Error GoTo EH
    If Len(Trim$(CStr(mail.SenderName))) > 0 Then
        GetSenderName = CStr(mail.SenderName)
    Else
        GetSenderName = CStr(mail.SenderEmailAddress)
    End If
    Exit Function
EH:
    GetSenderName = "(неизвестный отправитель)"
End Function

Private Function BuildOutputPath(ByVal targetDate As Date) As String
    Dim baseName As String, n As Long
    baseName = "Вопросы_" & Format$(targetDate, "dd-mm-yyyy")
    BuildOutputPath = OUTPUT_FOLDER_PATH & "\" & baseName & ".docx"
    n = 1
    Do While Len(Dir$(BuildOutputPath)) > 0
        BuildOutputPath = OUTPUT_FOLDER_PATH & "\" & baseName & "_" & CStr(n) & ".docx"
        n = n + 1
    Loop
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub