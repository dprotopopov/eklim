Attribute VB_Name = "Module2"
' Дмитрий Протопопов
' dmitry@protopopov.ru

Option Explicit

Public myFileName As String 'Имя файла данных
Public mySaveFolderName As String 'Имя папки сохранения
Public myExcelFolderName As String 'Имя папки сохранения заказов
Public myExcelFileName As String 'Имя файла - шаблона заказа
Public myImageFolderName As String 'Имя папки с изображениями
Public myExportWhite As Boolean 'Экспортировать белый слой
Public myExportThruAI As Boolean 'Экспортировать через AI
Public myWhiteAsCyan As Boolean 'Заменять белый CMYK на Cyan
Public myCyan As New Color 'Заменять белый CMYK на Cyan
Public myRgbWhite As Color ' белый
Public myCmykWhite As Color ' белый

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
        myCOLORCODE = "Б"
    ElseIf c.IsSame(myRGB("ЧЕРНЫЙ")) Then
        myCOLORCODE = "Ч"
    ElseIf c.IsSame(myRGB("СИНИЙ")) Then
        myCOLORCODE = "С"
    ElseIf c.IsSame(myRGB("ЖЕЛТЫЙ")) Then
        myCOLORCODE = "Ж"
    ElseIf c.IsSame(myRGB("КРАСНЫЙ")) Then
        myCOLORCODE = "К"
    ElseIf c.IsSame(myRGB("ФИОЛЕТОВЫЙ")) Then
        myCOLORCODE = "Ф"
    ElseIf c.IsSame(myRGB("ОРАНЖЕВЫЙ")) Then
        myCOLORCODE = "О"
    ElseIf c.IsSame(myRGB("РОЗОВЫЙ")) Then
        myCOLORCODE = "Р"
    ElseIf c.IsSame(myRGB("ГОЛУБОЙ")) Then
        myCOLORCODE = "Г"
    ElseIf c.IsSame(myRGB("ЗЕЛЕНЫЙ")) Then
        myCOLORCODE = "З"
    ElseIf c.IsSame(myRGB("СИЛИКОН")) Then
        myCOLORCODE = "Сил"
    ElseIf c.IsSame(myRGB("ЯРКО-РОЗОВЫЙ")) Then
        myCOLORCODE = "Я-Р"
    ElseIf c.IsSame(myRGB("СИЛИКОН ПРОЗРАЧНЫЙ")) Then
        myCOLORCODE = "СилП"
    ElseIf c.IsSame(myRGB("ПОЛУПРОЗРАЧНЫЙ")) Then
        myCOLORCODE = "ПП"
    Else
        Err.Raise vbObjectError, "myCOLORCODE", "Нет кода для цвета " + title
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
    Case "БЕЛЫЙ"
        If myWhiteAsCyan Then
            Set c = myCyan
        Else
            c.CMYKAssign 0, 0, 0, 0
        End If
    Case "ЧЕРНЫЙ"
        c.CMYKAssign 0, 0, 0, 100
    Case "СИНИЙ"
        c.CMYKAssign 100, 100, 0, 0
    Case "ЖЕЛТЫЙ"
        c.CMYKAssign 0, 0, 100, 0
    Case "КРАСНЫЙ"
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
        If UBound(WrdArray) <> 3 Then Err.Raise vbObjectError, "myCMYK", "Нет цвета " + title
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
    Case "БЕЛЫЙ"
        c.RGBAssign 255, 255, 255
    Case "ЧЕРНЫЙ"
        c.RGBAssign 0, 0, 0
    Case "СИНИЙ"
        c.RGBAssign 0, 0, 255
    Case "ЖЕЛТЫЙ"
        c.RGBAssign 255, 255, 0
    Case "КРАСНЫЙ"
        c.RGBAssign 255, 0, 0
    Case "ФИОЛЕТОВЫЙ"
        c.RGBAssign 139, 0, 255
    Case "ОРАНЖЕВЫЙ"
        c.RGBAssign 255, 165, 0
    Case "ГОЛУБОЙ"
        c.RGBAssign 66, 170, 255
    Case "ЗЕЛЕНЫЙ"
        c.RGBAssign 0, 128, 0
    Case "СИЛИКОН"
        c.RGBAssign 153, 255, 153
    Case "РОЗОВЫЙ"
        c.RGBAssign 255, 192, 203
    Case "ЯРКО-РОЗОВЫЙ"
        c.RGBAssign 252, 15, 192
    Case "СИЛИКОН ПРОЗРАЧНЫЙ"
        c.RGBAssign 153, 254, 153
    Case "ПОЛУПРОЗРАЧНЫЙ"
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
            Err.Raise vbObjectError, "myRGB", "Нет цвета " + title
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


