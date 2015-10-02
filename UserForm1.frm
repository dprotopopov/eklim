VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Макрос генерации листов для печати на чехлах телефонов"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12270
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    TextBox1.Text = CorelScriptTools.GetFileBox("Excel Files (*.xls)|*.xls", "Select a file", 0, TextBox1.Text)
End Sub

Private Sub CommandButton2_Click()
    TextBox2.Text = FOLDER_SELECTION(TextBox2.Text, 4, "Make your selection.")
End Sub

Private Sub CommandButton3_Click()
    TextBox3.Text = FOLDER_SELECTION(TextBox3.Text, 4, "Make your selection.")
End Sub

Private Sub CommandButton4_Click()
    TextBox4.Text = CorelScriptTools.GetFileBox("Excel Template Files (*.xlt)|*.xlt", "Select a file", 0, TextBox4.Text)
End Sub

Private Sub CommandButton5_Click()
    If TextBox1.Text = "" _
    Or TextBox2.Text = "" _
    Or TextBox3.Text = "" _
    Or TextBox4.Text = "" _
    Or TextBox5.Text = "" _
    Then
        MsgBox "Пустое поле"
        Exit Sub
    End If

    myFileName = TextBox1.Text
    mySaveFolderName = TextBox2.Text
    myExcelFolderName = TextBox3.Text
    myExcelFileName = TextBox4.Text
    myImageFolderName = TextBox5.Text
    myExportWhite = CheckBox1.Value
    myExportThruAI = CheckBox2.Value
    myWhiteAsCyan = CheckBox3.Value

    Hide
End Sub

Private Sub CommandButton6_Click()
    Hide
    Err.Raise vbObjectError, "Форма ввода", "Отмена операции"
End Sub

Private Sub CommandButton7_Click()
    TextBox5.Text = FOLDER_SELECTION(TextBox5.Text, 4, "Make your selection.")
End Sub

Private Sub UserForm_Initialize()
    TextBox1.Text = myFileName
    TextBox2.Text = mySaveFolderName
    TextBox3.Text = myExcelFolderName
    TextBox4.Text = myExcelFileName
    TextBox5.Text = myImageFolderName
    CheckBox1.Value = myExportWhite
    CheckBox2.Value = myExportThruAI
    CheckBox3.Value = myWhiteAsCyan
End Sub
