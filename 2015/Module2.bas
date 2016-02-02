Attribute VB_Name = "Module2"
' ������� ����������
' dmitry@protopopov.ru

Option Explicit

Public myFileName As String '��� ����� ������
Public mySaveFolderName As String '��� ����� ����������
Public myExcelFolderName As String '��� ����� ���������� �������
Public myExcelFileName As String '��� ����� - ������� ������
Public myImageFolderName As String '��� ����� � �������������
Public myExportWhite As Boolean '�������������� ����� ����
Public myExportThruAI As Boolean '�������������� ����� AI
Public myWhiteAsCyan As Boolean '�������� ����� CMYK �� Cyan
Public myCyan As New Color '�������� ����� CMYK �� Cyan
Public myRgbWhite As Color ' �����
Public myCmykWhite As Color ' �����

Function myFONT(title As String) As String
    Dim s As String
    s = UCase(Trim(title))
    
    Dim latin As String
    latin = " 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_+-:/\\.,;"
    Dim i As Integer
    For i = 1 To Len(latin)
        s = Trim(Replace(s, Mid(latin, i, 1), ""))
    Next i
    If Len(s) = 0 Then
        myFONT = "BigNoodleTitling"
    Else
        myFONT = "AA Higherup"
    End If
End Function

Function myCOLORCODE(title As String) As String
    Dim c As Color
    Set c = myRGB(title)
    If c.IsSame(myRgbWhite) Then
        myCOLORCODE = "�"
    ElseIf c.IsSame(myRGB("������")) Then
        myCOLORCODE = "�"
    ElseIf c.IsSame(myRGB("�����")) Then
        myCOLORCODE = "�"
    ElseIf c.IsSame(myRGB("������")) Then
        myCOLORCODE = "�"
    ElseIf c.IsSame(myRGB("�������")) Then
        myCOLORCODE = "�"
    ElseIf c.IsSame(myRGB("����������")) Then
        myCOLORCODE = "�"
    ElseIf c.IsSame(myRGB("���������")) Then
        myCOLORCODE = "�"
    ElseIf c.IsSame(myRGB("�������")) Then
        myCOLORCODE = "�"
    ElseIf c.IsSame(myRGB("�������")) Then
        myCOLORCODE = "�"
    ElseIf c.IsSame(myRGB("�������")) Then
        myCOLORCODE = "�"
    ElseIf c.IsSame(myRGB("�������")) Then
        myCOLORCODE = "���"
    ElseIf c.IsSame(myRGB("����-�������")) Then
        myCOLORCODE = "�-�"
    ElseIf c.IsSame(myRGB("������� ����������")) Then
        myCOLORCODE = "����"
    ElseIf c.IsSame(myRGB("��������������")) Then
        myCOLORCODE = "��"
    Else
        Err.Raise vbObjectError, "myCOLORCODE", "��� ���� ��� ����� " + title
    End If
End Function

Function myIMAGE(model As String, title As String) As String
    Dim s1 As String, s2 As String
    s1 = Trim(Replace(model, "/", " "))
    s2 = Trim(Replace(title, "/", " "))
    myIMAGE = myImageFolderName + "\" + s1 + "\" + s2
End Function

Function mySYMBOL(model As String, title As String) As String
    Dim s1 As String, s2 As String
    s1 = UCase(Trim(Replace(model, "/", "_")))
    s2 = UCase(Trim(Replace(title, "/", "_")))
    mySYMBOL = s1 + "-" + s2
End Function

Function myCMYK(title As String) As Color
    Dim s As String
    s = UCase(Trim(title))
    
    Dim c As New Color
    Select Case s
    Case "�����"
        If myWhiteAsCyan Then
            Set c = myCyan
        Else
            c.CMYKAssign 0, 0, 0, 0
        End If
    Case "������"
        c.CMYKAssign 0, 0, 0, 100
    Case "�����"
        c.CMYKAssign 100, 100, 0, 0
    Case "������"
        c.CMYKAssign 0, 0, 100, 0
    Case "�������"
        c.CMYKAssign 0, 100, 100, 0
    Case Else
        s = Replace(s, ",", " ")
        Dim temp As String
        Do
          temp = s
          s = Replace(s, "  ", " ") 'remove multiple white spaces
        Loop Until temp = s
        'http://www.exceltrick.com/formulas_macros/vba-split-function/
        Dim WrdArray() As String
        WrdArray() = Split(s, " ")
        If UBound(WrdArray) <> 3 Then Err.Raise vbObjectError, "myCMYK", "��� ����� " + title
        If CInt(WrdArray(0)) = 0 _
        And CInt(WrdArray(1)) = 0 _
        And CInt(WrdArray(2)) = 0 _
        And CInt(WrdArray(3)) = 0 _
        And myWhiteAsCyan Then
            Set c = myCyan
        Else
            c.CMYKAssign CInt(WrdArray(0)), CInt(WrdArray(1)), CInt(WrdArray(2)), CInt(WrdArray(3))
        End If
    End Select
    Set myCMYK = c
End Function

Function myRGB(title As String) As Color
    Dim s As String
    s = UCase(Trim(title))
    
    Dim c As New Color
    Select Case s
    Case "�����"
        c.RGBAssign 255, 255, 255
    Case "������"
        c.RGBAssign 0, 0, 0
    Case "�����"
        c.RGBAssign 0, 0, 255
    Case "������"
        c.RGBAssign 255, 255, 0
    Case "�������"
        c.RGBAssign 255, 0, 0
    Case "����������"
        c.RGBAssign 139, 0, 255
    Case "���������"
        c.RGBAssign 255, 165, 0
    Case "�������"
        c.RGBAssign 66, 170, 255
    Case "�������"
        c.RGBAssign 0, 128, 0
    Case "�������"
        c.RGBAssign 153, 255, 153
    Case "�������"
        c.RGBAssign 255, 192, 203
    Case "����-�������"
        c.RGBAssign 252, 15, 192
    Case "������� ����������"
        c.RGBAssign 153, 254, 153
    Case "��������������"
        c.RGBAssign 153, 224, 153
    Case Else
        s = Replace(s, ",", " ")
        Dim temp As String
        Do
          temp = s
          s = Replace(s, "  ", " ") 'remove multiple white spaces
        Loop Until temp = s
        'http://www.exceltrick.com/formulas_macros/vba-split-function/
        Dim WrdArray() As String
        WrdArray() = Split(s, " ")
        Select Case UBound(WrdArray)
        Case 2
            c.RGBAssign CInt(WrdArray(0)), CInt(WrdArray(1)), CInt(WrdArray(2))
        Case 3
            If CInt(WrdArray(0)) = 0 _
            And CInt(WrdArray(1)) = 0 _
            And CInt(WrdArray(2)) = 0 _
            And CInt(WrdArray(3)) = 0 Then
                c.RGBAssign 255, 255, 255
            Else
                Set c = myCMYK(s)
            End If
        Case Else
            Err.Raise vbObjectError, "myRGB", "��� ����� " + title
        End Select
    End Select
    Set myRGB = c
End Function

Function GetFileName(prj As String, Suffix As String)
    GetFileName = mySaveFolderName + "\" + prj + Suffix
End Function

Function GetProjectName() As String
    GetProjectName = Format(Now, "YYYYmmDD-HHMMSS")
End Function


