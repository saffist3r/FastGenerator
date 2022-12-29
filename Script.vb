Option Explicit
Dim fichier As String
Dim longstatus As Long, longwarnings As Long
Dim swApp As Object
Dim Part As Object
Dim swCustPrpMgr As SldWorks.CustomPropertyManager
Dim fs, f As Object
Private m_sDefaultFolder As String


Private Sub CheckBox1_Click()
Dim compt As Integer
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

Private Sub CommandButton1_Click()
Dim i As Integer
Dim j As Integer
Dim selCount As Integer
Dim selCunt2 As Integer
Dim dossier As String
Dim nom As String
Dim NomFichierSansExtension As String
Set swApp = Application.SldWorks
    Label5.Caption = ""
    If (TextBox1.Text = "" Or TextBox2.Text = "") Then
    Label5.Caption = "Please verify destination/Source folder."
    Else
    dossier = TextBox1.Text & "\"
    selCount = -1
    For i = 0 To ListBox1.ListCount - 1
     If ListBox1.Selected(i) = True Then
     nom = ListBox1.List(i)
     NomFichierSansExtension = Left(nom, Len(nom) - 7)
     Set Part = swApp.OpenDoc6(dossier & nom, 3, 0, "", longstatus, longwarnings)
     swApp.OpenDoc6 dossier & nom, 3, 0, "", longstatus, longwarnings
     Set Part = swApp.ActivateDoc2(nom, False, longstatus)
     Set swCustPrpMgr = Part.Extension.CustomPropertyManager("")
     
    For j = 0 To ComboBox1.ListCount - 1
     If (ComboBox1.Selected(j) = True) Then
     Part.Extension.SaveAs TextBox2.Text & "/" & NomFichierSansExtension & ComboBox1.List(j), 0, 0, Nothing, longstatus, longwarnings
     End If
    Next
    'Fermeture du plan
    Set Part = Nothing
    swApp.CloseDoc nom
    
     End If
    Next
    End If
End Sub

Private Sub CommandButton2_Click()
    ListBox1.Clear
    Dim selectedFolder As String
    selectedFolder = BrowseForFolder()
    
    If Len(selectedFolder) > 0 Then
        m_sDefaultFolder = selectedFolder
    End If
    TextBox1.Text = selectedFolder
    fichier = Dir(TextBox1.Text & "/*.SLDDRW")
    Do While fichier <> ""
    ListBox1.AddItem (fichier)
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
