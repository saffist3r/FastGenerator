Option Explicit

Dim fichier As String
Dim longstatus As Long, longwarnings As Long
Dim swApp As Object
Dim Part As Object
Dim swCustPrpMgr As SldWorks.CustomPropertyManager
Dim fs, f As Object
Private m_sDefaultFolder As String

Private Sub CheckBox1_Click()
    Dim i As Integer

    If CheckBox1.Value = True Then
        For i = 0 To ListBox1.ListCount - 1
            ListBox1.Selected(i) = True
        Next
    Else
        For i = 0 To ListBox1.ListCount - 1
            ListBox1.Selected(i) = False
        Next
    End If
End Sub

Private Function NormalizePath(ByVal pathValue As String) As String
    NormalizePath = Replace(Trim(pathValue), "/", "\")
End Function

Private Function IsValidDirectory(ByVal pathValue As String) As Boolean
    Dim normalized As String
    normalized = NormalizePath(pathValue)

    If Len(normalized) = 0 Then
        IsValidDirectory = False
        Exit Function
    End If

    IsValidDirectory = (Dir(normalized, vbDirectory) <> "")
End Function

Private Function HasSelectedDrawings() As Boolean
    Dim i As Integer

    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            HasSelectedDrawings = True
            Exit Function
        End If
    Next

    HasSelectedDrawings = False
End Function

Private Function HasSelectedFormats() As Boolean
    Dim j As Integer

    For j = 0 To ComboBox1.ListCount - 1
        If ComboBox1.Selected(j) = True Then
            HasSelectedFormats = True
            Exit Function
        End If
    Next

    HasSelectedFormats = False
End Function

Private Function BuildLogPath(ByVal destinationFolder As String) As String
    BuildLogPath = destinationFolder & "\ExportLog_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
End Function

Private Sub WriteLogFile(ByVal logPath As String, ByVal logLines As Collection)
    Dim fileNo As Integer
    Dim i As Integer

    fileNo = FreeFile
    Open logPath For Output As #fileNo
    Print #fileNo, "Timestamp,File,Format,Result,StatusCode,WarningCode,Message"

    For i = 1 To logLines.Count
        Print #fileNo, logLines(i)
    Next

    Close #fileNo
End Sub

Private Sub AddLogLine(ByRef logLines As Collection, ByVal fileName As String, ByVal exportFormat As String, ByVal result As String, ByVal statusCode As Long, ByVal warningCode As Long, ByVal message As String)
    Dim safeMessage As String
    safeMessage = Replace(message, ",", ";")

    logLines.Add Format(Now, "yyyy-mm-dd hh:nn:ss") & "," & fileName & "," & exportFormat & "," & result & "," & CStr(statusCode) & "," & CStr(warningCode) & "," & safeMessage
End Sub

Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim j As Integer
    Dim sourceFolder As String
    Dim destinationFolder As String
    Dim nom As String
    Dim NomFichierSansExtension As String
    Dim exportPath As String
    Dim exportOk As Boolean
    Dim openStatus As Long
    Dim openWarnings As Long
    Dim saveStatus As Long
    Dim saveWarnings As Long
    Dim logLines As Collection
    Dim logPath As String
    Dim successCount As Long
    Dim failedCount As Long

    Set swApp = Application.SldWorks
    Label5.Caption = ""

    sourceFolder = NormalizePath(TextBox1.Text)
    destinationFolder = NormalizePath(TextBox2.Text)

    If Not IsValidDirectory(sourceFolder) Then
        Label5.Caption = "Source folder is invalid or does not exist."
        Exit Sub
    End If

    If Not IsValidDirectory(destinationFolder) Then
        Label5.Caption = "Destination folder is invalid or does not exist."
        Exit Sub
    End If

    If Not HasSelectedDrawings() Then
        Label5.Caption = "Select at least one drawing file."
        Exit Sub
    End If

    If Not HasSelectedFormats() Then
        Label5.Caption = "Select at least one export format."
        Exit Sub
    End If

    Set logLines = New Collection
    successCount = 0
    failedCount = 0

    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            nom = ListBox1.List(i)
            NomFichierSansExtension = Left(nom, Len(nom) - 7)

            openStatus = 0
            openWarnings = 0
            Set Part = swApp.OpenDoc6(sourceFolder & "\" & nom, 3, 0, "", openStatus, openWarnings)

            If Part Is Nothing Then
                failedCount = failedCount + 1
                Call AddLogLine(logLines, nom, "ALL", "FAILED", openStatus, openWarnings, "OpenDoc6 returned Nothing")
                GoTo ContinueNextDrawing
            End If

            Set Part = swApp.ActivateDoc2(nom, False, longstatus)
            If Part Is Nothing Then
                failedCount = failedCount + 1
                Call AddLogLine(logLines, nom, "ALL", "FAILED", longstatus, openWarnings, "ActivateDoc2 returned Nothing")
                swApp.CloseDoc nom
                GoTo ContinueNextDrawing
            End If

            Set swCustPrpMgr = Part.Extension.CustomPropertyManager("")

            For j = 0 To ComboBox1.ListCount - 1
                If ComboBox1.Selected(j) = True Then
                    saveStatus = 0
                    saveWarnings = 0
                    exportPath = destinationFolder & "\" & NomFichierSansExtension & ComboBox1.List(j)

                    exportOk = Part.Extension.SaveAs(exportPath, 0, 0, Nothing, saveStatus, saveWarnings)

                    If exportOk Then
                        successCount = successCount + 1
                        Call AddLogLine(logLines, nom, ComboBox1.List(j), "SUCCESS", saveStatus, saveWarnings, "Export completed")
                    Else
                        failedCount = failedCount + 1
                        Call AddLogLine(logLines, nom, ComboBox1.List(j), "FAILED", saveStatus, saveWarnings, "SaveAs returned False")
                    End If
                End If
            Next

            swApp.CloseDoc nom
            Set Part = Nothing
        End If
ContinueNextDrawing:
    Next

    logPath = BuildLogPath(destinationFolder)
    Call WriteLogFile(logPath, logLines)

    Label5.Caption = "Done. Success: " & CStr(successCount) & " | Failed: " & CStr(failedCount) & " | Log: " & logPath
End Sub

Private Sub CommandButton2_Click()
    Dim selectedFolder As String

    ListBox1.Clear
    selectedFolder = BrowseForFolder()

    If Len(selectedFolder) > 0 Then
        m_sDefaultFolder = selectedFolder
    End If

    TextBox1.Text = selectedFolder

    fichier = Dir(NormalizePath(TextBox1.Text) & "\*.slddrw")
    Do While fichier <> ""
        ListBox1.AddItem (fichier)
        UserForm1.Repaint
        fichier = Dir
    Loop

    ' Also include uppercase extension variant for compatibility.
    fichier = Dir(NormalizePath(TextBox1.Text) & "\*.SLDDRW")
    Do While fichier <> ""
        If ListBox1.ListCount = 0 Then
            ListBox1.AddItem (fichier)
        Else
            If ListBox1.List(ListBox1.ListCount - 1) <> fichier Then
                ListBox1.AddItem (fichier)
            End If
        End If
        UserForm1.Repaint
        fichier = Dir
    Loop
End Sub

Private Sub CommandButton3_Click()
    Dim selectedFolder As String

    selectedFolder = BrowseForFolder()

    If Len(selectedFolder) > 0 Then
        m_sDefaultFolder = selectedFolder
    End If

    TextBox2.Text = selectedFolder
End Sub

Private Sub UserForm_Initialize()
    UserForm1.ComboBox1.AddItem (".dxf")
    UserForm1.ComboBox1.AddItem (".dwg")
    UserForm1.ComboBox1.AddItem (".PDF")
    UserForm1.ComboBox1.AddItem (".JPG")
    UserForm1.ComboBox1.AddItem (".TIF")
    UserForm1.ComboBox1.AddItem (".edrw")
End Sub