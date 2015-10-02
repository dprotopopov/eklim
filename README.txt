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

У Вас со всеми Galaxy одна и та же проблема
Хотя каждый контур является одной кривой, но на самом деле состоит из отдельных не связанных между собой сегментов – поэтому такая кривая просто даже не может иметь заливки, т.к. не имеет замкнутых кривых.
Попробуйте – вручную сами закрасить не сможете.
Попробуйте – выделите кривую – пункт меню Break Curve – увидите что сегменты не соединены.

Способы решения:
1/ правильно экспортировать-импортировать если верстали в другой программе - рекомендую импортировать закрашенные фигуры, а не контуры - если фигура закрашена - значит кривые замкнуты.
2/ вручную соединять отрезки – т.е. сперва разбить Break Curve – увеличить изображение – выделять соприкасающиеся концы 2шт и пункт меню Join – полученные замкнутые кривые опять в одну кривую – меню Combine

v3.
Во входящем файле теперь 2 строки с именами - Надпись и Надпись 2. А также строка Шаблон.
Шаблон указывает на то, как размещать строки на чехле.
Если в поле шаблон стоит 1, то чехол обычный (как сейчас) и имя берется из ячейки Надпись. Если шаблон = 2, то расположение в 2 строки параллельные друг другу. Если шаблон =3, то расположение, когда строки перпендикулярны.
Надпись - это первая строка, Надпись 2 - это вторая.
Все это должно также корректно работать с заполнением букв фонами. Логика в этом плане не меняется.
