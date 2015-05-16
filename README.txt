Открыть документ eklim.cdr
Выполнить макрос eklim
Разрешить макросы


In the VBA Editor under Tools then select References, scroll to the reference you want and click the small square to its left. Select OK and now this reference will be moved to near the top of all the references used
http://corel-vba.awardspace.com/From_Excel.htm

В референсы, в режиме редактора макросов требуется
Excel Type Library
Illustrator Type Library


    myFileName = "D:\Projects\eklim\èìåíà.xls"
    myFolderName = "D:\Projects\eklim\Íîâàÿ ïàïêà"
    myExcelFolderName = "D:\Projects\eklim\Íîâàÿ ïàïêà"
    myExcelFileName = "D:\Projects\eklim\çàêàç.xlt"
    myExportWhite = False
    myExportThruAI = False
    myWhiteAsCyan = True
    myCyan.CMYKAssign 100, 0, 0, 0
    

Function myFONT(title As String) As String
    Dim s As String
    s = Trim(UCase(title))
    
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
Function myCMYK(title As String) As Color
Function myRGB(title As String) As Color

