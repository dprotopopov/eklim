Attribute VB_Name = "MainModule"
' Дмитрий Протопопов
' dmitry@protopopov.ru

Option Explicit

Sub eklim()


On Error GoTo errMyErrorHandler

    
    Dim i  As Integer, j  As Integer, m As Integer, n As Integer, k As Integer, t As Integer
    Dim Width As Double, Height As Double
    Dim total As Integer
    Dim sr As ShapeRange
    Dim sr0 As ShapeRange
    Dim sr1 As ShapeRange
    Dim sr2 As ShapeRange
    Dim sr3 As ShapeRange
    Dim sr4 As ShapeRange
    Dim s As Shape
    Dim lyr As Layer
    Dim lyr0 As Layer ' MESSAGE
    Dim lyr1 As Layer ' SUBMESSAGE
    Dim lyr2 As Layer ' LOGOTYPE
    Dim lyr3 As Layer ' COLORCODE
    Dim lyr4 As Layer ' CONTOUR
    Dim lyr5 As Layer ' WHITE
    Dim lyr6 As Layer ' RGB
    Dim rgb As Color, rgb1 As Color
    Dim cmyk As Color
    Dim export As StructExportOptions
    Dim expflt As ExportFilter
    Dim impopt As StructImportOptions
    Dim impflt As ImportFilter
    Dim syms As SymbolDefinitions
    Dim file As String
    Dim title As String
    Dim symname As String
    Dim x0 As Double, y0 As Double, w0 As Double, h0 As Double
    Dim x1 As Double, y1 As Double, w1 As Double, h1 As Double
    Dim scaleX As Double, scaleY As Double, scaleXY  As Double
                
    
'1.       Отобразить экранную форму
'2.       Имя папки охранения eps = имя папки охранения
'3.       Имя папки охранения jpg = имя папки охранения
'4.       Очищаем папку охранения eps
'5.       Очищаем папку охранения jpg
'6.       Очищаем папку охранения заказов
'7.       Index=0
'8.       Составляем список используемых моделей чехлов
'9.       Для каждой модели чехла
'a.       Делаем выборку из общего списка
'b.      Берём лист с совпадающим именем
'c.       Пока есть строки в выборке
'i.      Создаём временный лист с двумя слоями
'ii.      Набиваем временный лист под завязку, при этом в верхний слой копия объекта Фигурный текст с изменёнными параметрами, в нижний слой копия объекта Фигурный текст с изменёнными параметрами и белым цветом и выплёвываем
'1.       В папку охранения Index .eps
'2.       В папку охранения Index .jpg
'                                                           iii.Index = Index + 1
'10.   Index=0
'11.   Составляем список номеров заказов
'12.   Для каждого номера заказа
'a.       Делаем выборку из общего списка
'b.      Создаём копию файла - шаблона заказа Index .xls
'c.       Наполняем эту копию данными из выборки
'd.Index = Index + 1
 

    
'1.       Отобразить экранную форму
    
    myFileName = "D:\Projects\eklim\имена.v3.xls"
    mySaveFolderName = "D:\Projects\eklim\makets"
    myExcelFolderName = "D:\Projects\eklim\makets"
    myExcelFileName = "D:\Projects\eklim\заказ.xlt"
    myImageFolderName = "D:\Projects\eklim\backgrounds"
    
'    myFileName = "E:\ipapai\скрипт имена\Под печать\имена.xls"
'    mySaveFolderName = "E:\ipapai\скрипт имена\Под печать\makets"
'    myExcelFolderName = "E:\ipapai\скрипт имена\Под печать\makets"
'    myExcelFileName = "E:\ipapai\скрипт имена\Под печать\заказ.xlt"
'    myImageFolderName = "E:\ipapai\скрипт имена\Под печать\backgrounds"
    
    myExportWhite = False
    myExportThruAI = False
    myWhiteAsCyan = True
    myCyan.CMYKAssign 100, 0, 0, 0
    
    UserForm1.Show vbModal
    
    Set myRgbWhite = myRGB("БЕЛЫЙ")
    Set myCmykWhite = myCMYK("БЕЛЫЙ")
    
' http://www.taltech.com/support/entry/opening_and_closing_an_application_from_vba
If myExportThruAI Then
On Error Resume Next
Dim x
x = Shell("X:\Program Files\Adobe\Adobe Illustrator CS6 (64 Bit)\Support Files\Contents\Windows\Illustrator.exe", vbNormalFocus)
x = Nothing
On Error GoTo errMyErrorHandler
End If

'4.       Очищаем папку охранения eps
'5.       Очищаем папку охранения jpg
'6.       Очищаем папку охранения заказов
    
    STATUS.BeginProgress "Удаляем файлы", True
    
'http://www.jpsoftwaretech.com/vba/filesystemobject-vba-examples/
    Dim fs, f, fl, fc
    Set fs = CreateObject("Scripting.FileSystemObject")
    
        Set f = fs.GetFolder(mySaveFolderName)
        Set fc = f.Files
        For Each fl In fc
            Select Case UCase(fs.GetExtensionName(fl))
            Case "AI"
                fl.Delete
            Case "EPS"
                fl.Delete
            Case "JPG"
                fl.Delete
            End Select
        Next
     
        Set f = fs.GetFolder(myExcelFolderName)
        Set fc = f.Files
        For Each fl In fc
            Select Case UCase(fs.GetExtensionName(fl))
            Case "XLS"
                fl.Delete
            End Select
        Next
   
    STATUS.EndProgress
    
'8.       Составляем список используемых моделей чехлов
' http://corel-vba.awardspace.com/From_Excel.htm

Dim xls(10000, 13) As String


Dim ORDERDATE(10000) As String
Dim ORDER(10000) As String
Dim model(10000) As String
Dim MODELCOLOR(10000) As String
Dim MODELCOLOR1(10000) As String
Dim MESSAGE(10000) As String
Dim SUBMESSAGE(10000) As String
Dim MESSAGETEMPLATE(10000) As String
Dim LOGOTYPE(10000) As String
Dim LABELCOLOR(10000) As String
Dim MESSAGEFONT(10000) As String
Dim LOGOTYPEFONT(10000) As String
Dim AMOUNT(10000) As String
Dim LOGOTYPETEMPLATE(10000) As String

Dim MODEL_COUNTER As Integer
Dim MODELCOLOR_COUNTER As Integer
Dim ORDERDATE_COUNTER As Integer
Dim ORDER_COUNTER As Integer
Dim COUNTER As Integer

'The data array here allows 10,000 sets of data.
'If more data sets are supplied in the data file, increase the 10000 to a larger number.
'The second variable in the array ie 2 is the number of variables in each data set less 1.
'ie I have used 3 columns of data in the Excel Worksheet, so Dim DATA(10000, 2).
'If you have only 1 variable in a data set then Dim DATA(10000,0).
'If you have 10 variables in a data set then Dim DATA(10000,9).

Dim EXCELAPP As Object
Dim ROW_COUNTER As Integer
Dim row As Integer

    STATUS.BeginProgress "Чтение данных", True

Set EXCELAPP = CreateObject("excel.application")
'Make Excel invisible.
EXCELAPP.visible = False

'Now open the Excel file that is located in the same folder as this CorelDraw fie.
Workbooks.Open (myFileName)

'The following line can be used if the Excel file is on Fred's Desktop whilst the CorelDraw file can be anywhere on the same computer.
'Workbooks.Open ("C:\Documents and Settings\Fred\Desktop\" & EXCEL_DATA_FILE)

'Now assign the spreadsheet data to the array.
ROW_COUNTER = 0
'To ensure we read all the data in all 3 spreadsheet columns the while statement must check each column for data.
'If there was only one column of data then it would be;
'While cells(ROW_COUNTER + 1,......
'Note that we use cells(ROW_COUNTER + 1) in this While loop.
'This is because the first row of the spreadsheet is a header & so is not read.
'If you do not have a header use cells(ROW_COUNTER,...... in place of cells(ROW_COUNTER,..... ie Do not use + 1
Dim dateString As String
While Trim(CStr(cells(ROW_COUNTER + 2, 1))) <> "" _
Or Trim(CStr(cells(ROW_COUNTER + 2, 2))) <> "" _
Or Trim(CStr(cells(ROW_COUNTER + 2, 3))) <> ""
    If Trim(CStr(cells(ROW_COUNTER + 2, 1))) <> "" Then
        dateString = Trim(CStr(cells(ROW_COUNTER + 2, 1)))
    End If
    xls(ROW_COUNTER, 0) = dateString
    For i = 1 To 10
        xls(ROW_COUNTER, i) = Trim(CStr(cells(ROW_COUNTER + 2, i + 1)))
    Next i
    ROW_COUNTER = ROW_COUNTER + 1
Wend

'Close excel.
Excel.Application.Workbooks.Close

    STATUS.EndProgress
    

    STATUS.BeginProgress "Генерация отчётов", True
    
ORDERDATE_COUNTER = 0

For row = 0 To ROW_COUNTER - 1
    If xls(row, 0) <> "" Then
        ORDERDATE(ORDERDATE_COUNTER) = CStr(xls(row, 0))
        ORDERDATE_COUNTER = ORDERDATE_COUNTER + 1
    End If
Next row

If ORDERDATE_COUNTER = 0 Then Err.Raise vbObjectError, "myFileName", "Нет списка заказов"

'http://www.xtremevbtalk.com/showthread.php?t=150521
If ORDERDATE_COUNTER > 1 Then Call QuickSort(ORDERDATE, 0, ORDERDATE_COUNTER - 1)

'удаление дипликатов из ORDERDATE
For k = ORDERDATE_COUNTER - 2 To 0 Step -1
    If UCase(ORDERDATE(k)) = UCase(ORDERDATE(k + 1)) Then
        ORDERDATE(k + 1) = ORDERDATE(ORDERDATE_COUNTER - 1)
        ORDERDATE_COUNTER = ORDERDATE_COUNTER - 1
    End If
Next k

If ORDERDATE_COUNTER > 1 Then Call QuickSort(ORDERDATE, 0, ORDERDATE_COUNTER - 1)

For m = 0 To ORDERDATE_COUNTER - 1
        
    COUNTER = 0

    For row = 0 To ROW_COUNTER - 1
        If UCase(xls(row, 0)) = UCase(ORDERDATE(m)) _
        And CStr(xls(row, 7)) <> "" Then
            ORDER(COUNTER) = CStr(xls(row, 1))
            model(COUNTER) = CStr(xls(row, 2))
            MODELCOLOR(COUNTER) = CStr(xls(row, 3))
            MESSAGE(COUNTER) = CStr(xls(row, 4))
            SUBMESSAGE(COUNTER) = CStr(xls(row, 5))
            MESSAGETEMPLATE(COUNTER) = CStr(xls(row, 6))
            LABELCOLOR(COUNTER) = CStr(xls(row, 7))
            LOGOTYPE(COUNTER) = CStr(xls(row, 8))
            AMOUNT(COUNTER) = CStr(xls(row, 9))
            LOGOTYPETEMPLATE(COUNTER) = CStr(xls(row, 10))
            MESSAGEFONT(COUNTER) = myFONT(MESSAGE(COUNTER) + "/" + SUBMESSAGE(COUNTER))
            LOGOTYPEFONT(COUNTER) = myFONT(LOGOTYPE(COUNTER))
            COUNTER = COUNTER + 1
        End If
    Next row
    
    
    total = 0
    For k = 0 To COUNTER - 1
        total = total + CInt(AMOUNT(k))
    Next k
        
    If total > 0 Then
        Workbooks.Open (myExcelFileName)
        cells(5, 5) = mySaveFolderName
        cells(6, 5) = CStr(total)
        cells(7, 5) = ORDERDATE(m)
            
        For k = 0 To COUNTER - 1
            cells(11 + k, 2) = CStr(k + 1)
            cells(11 + k, 4) = MESSAGE(k) + "/" + SUBMESSAGE(k)
            cells(11 + k, 6) = model(k)
            cells(11 + k, 7) = MODELCOLOR(k)
            cells(11 + k, 12) = AMOUNT(k)
        Next k
    
        Workbooks(1).SaveAs (myExcelFolderName + "\" + ORDERDATE(m) + ".xls")
        Excel.Application.Workbooks.Close
    End If
Next m

EXCELAPP.visible = True
EXCELAPP.Quit

    STATUS.EndProgress

    STATUS.BeginProgress "Генерация столов", True

MODEL_COUNTER = 0

For row = 0 To ROW_COUNTER - 1
    If xls(row, 2) <> "" Then
        model(MODEL_COUNTER) = CStr(xls(row, 2))
        MODEL_COUNTER = MODEL_COUNTER + 1
    End If
Next row

If MODEL_COUNTER = 0 Then Err.Raise vbObjectError, "myFileName", "Нет списка моделей"

'http://www.xtremevbtalk.com/showthread.php?t=150521
If MODEL_COUNTER > 1 Then Call QuickSort(model, 0, MODEL_COUNTER - 1)

'удаление дипликатов из MODEL
For k = MODEL_COUNTER - 2 To 0 Step -1
    If UCase(model(k)) = UCase(model(k + 1)) Then
        model(k + 1) = model(MODEL_COUNTER - 1)
        MODEL_COUNTER = MODEL_COUNTER - 1
    End If
Next k

If MODEL_COUNTER > 1 Then Call QuickSort(model, 0, MODEL_COUNTER - 1)

Dim templateDoc As Document
Set templateDoc = ThisDocument
templateDoc.Dirty = False

Dim newDoc As Document
Set newDoc = CreateDocument
newDoc.Activate
    
newDoc.Unit = templateDoc.Unit
newDoc.ReferencePoint = cdrCenter
    
n = 1
    

'9.       Для каждой модели чехла
For m = 0 To MODEL_COUNTER - 1
    
    MODELCOLOR_COUNTER = 0

    For row = 0 To ROW_COUNTER - 1
        If UCase(xls(row, 2)) = UCase(model(m)) Then
            MODELCOLOR1(MODELCOLOR_COUNTER) = CStr(xls(row, 3))
            MODELCOLOR_COUNTER = MODELCOLOR_COUNTER + 1
        End If
    Next row

    'http://www.xtremevbtalk.com/showthread.php?t=150521
    If MODELCOLOR_COUNTER > 1 Then Call QuickSort(MODELCOLOR1, 0, MODELCOLOR_COUNTER - 1)
    
    'удаление дипликатов из MODELCOLOR
    For k = MODELCOLOR_COUNTER - 2 To 0 Step -1
        If UCase(MODELCOLOR1(k)) = UCase(MODELCOLOR1(k + 1)) Then
            MODELCOLOR1(k + 1) = MODELCOLOR1(MODELCOLOR_COUNTER - 1)
            MODELCOLOR_COUNTER = MODELCOLOR_COUNTER - 1
        End If
    Next k

    If MODELCOLOR_COUNTER > 1 Then Call QuickSort(MODELCOLOR1, 0, MODELCOLOR_COUNTER - 1)
    
'a.       Делаем выборку из общего списка
'b.      Берём лист с совпадающим именем
    
    Dim id As Integer
    id = FindPage(templateDoc, model(m))
    If id = 0 Then Err.Raise vbObjectError, "Документ шаблонов", "Нет модели " + model(m)
    
    Dim objPage As Page
    Set objPage = templateDoc.Pages(id)
        
        COUNTER = 0
        
    For i = 0 To MODELCOLOR_COUNTER - 1
    
        For row = 0 To ROW_COUNTER - 1
            If UCase(xls(row, 2)) = UCase(model(m)) _
            And UCase(xls(row, 3)) = UCase(MODELCOLOR1(i)) Then
                MODELCOLOR(COUNTER) = CStr(xls(row, 3))
                MESSAGE(COUNTER) = CStr(xls(row, 4))
                SUBMESSAGE(COUNTER) = CStr(xls(row, 5))
                MESSAGETEMPLATE(COUNTER) = CStr(xls(row, 6))
                LABELCOLOR(COUNTER) = CStr(xls(row, 7))
                LOGOTYPE(COUNTER) = CStr(xls(row, 8))
                AMOUNT(COUNTER) = CStr(xls(row, 9))
                LOGOTYPETEMPLATE(COUNTER) = CStr(xls(row, 10))
                MESSAGEFONT(COUNTER) = myFONT(MESSAGE(COUNTER))
                LOGOTYPEFONT(COUNTER) = myFONT(LOGOTYPE(COUNTER))
                
                If MESSAGE(COUNTER) + SUBMESSAGE(COUNTER) = "" Then Err.Raise vbObjectError, model(m), "Не задана MESSAGE в строке " + CStr(row + 1)
                If LABELCOLOR(COUNTER) = "" Then Err.Raise vbObjectError, model(m), "Не задана LABELCOLOR в строке " + CStr(row + 1)
                If MESSAGEFONT(COUNTER) = "" Then Err.Raise vbObjectError, model(m), "Не задана MESSAGEFONT в строке " + CStr(row + 1)
                If LOGOTYPEFONT(COUNTER) = "" Then Err.Raise vbObjectError, model(m), "Не задана LOGOTYPEFONT в строке " + CStr(row + 1)
    
                COUNTER = COUNTER + 1
            End If
        Next row
        
    Next i
        
        k = 0
        AMOUNT(COUNTER) = "0"
        
        Do
            Dim newPage As Page
            If n = 1 Then
                Set newPage = newDoc.Pages(1)
            Else
                Set newPage = newDoc.AddPages(1)
            End If
        
            objPage.GetSize Width, Height
            With newPage
                .SetSize Width, Height
                .name = model(m) + "-" + CStr(n)
                .PrintExportBackground = True
                .Background = cdrPageBackgroundNone
                Set syms = .Parent.Parent.SymbolLibrary.Symbols
            End With
            
            
            For Each lyr In newPage.Layers
                lyr.Printable = False
                lyr.editable = True
                lyr.visible = True
            Next lyr
            
            Set lyr6 = FindOrCreateLayer(newPage, "RGB")
            Set lyr5 = FindOrCreateLayer(newPage, "WHITE")
            Set lyr4 = FindOrCreateLayer(newPage, "CONTOUR")
            Set lyr3 = FindOrCreateLayer(newPage, "COLORCODE")
            Set lyr2 = FindOrCreateLayer(newPage, "LOGOTYPE")
            Set lyr1 = FindOrCreateLayer(newPage, "SUBMESSAGE")
            Set lyr0 = FindOrCreateLayer(newPage, "MESSAGE")
            
            CopyPage2 objPage, newPage
                        
            Set lyr = FindOrCreateLayer(newPage, "IMPORT")
            lyr.Delete ' IMPORT
            
            For Each lyr In newPage.Layers
                With lyr
                    .Printable = False
                    .editable = True
                End With
            Next lyr
                                
            Set sr0 = lyr0.FindShapes(Type:=cdrTextShape) ' MESSAGE
            Set sr1 = lyr1.FindShapes(Type:=cdrTextShape) ' SUBMESSAGE
            Set sr2 = lyr2.FindShapes(Type:=cdrTextShape) ' LOGOTYPE
            Set sr3 = lyr3.FindShapes(Type:=cdrTextShape) ' COLORCODE
            Set sr4 = lyr4.FindShapes(Type:=cdrCurveShape) ' CONTOUR

            t = 0
            If t < sr0.Count Then t = sr0.Count ' MESSAGE
            If t < sr1.Count Then t = sr1.Count ' SUBMESSAGE
            If t < sr2.Count Then t = sr2.Count ' LOGOTYPE
            If t < sr3.Count Then t = sr3.Count ' COLORCODE
            If t < sr4.Count Then t = sr4.Count ' CONTOUR
            
            If t > sr0.Count And sr0.Count > 0 Then t = sr0.Count ' MESSAGE
            If t > sr1.Count And sr1.Count > 0 Then t = sr1.Count ' SUBMESSAGE
            If t > sr2.Count And sr2.Count > 0 Then t = sr2.Count ' LOGOTYPE
            If t > sr3.Count And sr3.Count > 0 Then t = sr3.Count ' COLORCODE
            If t > sr4.Count And sr4.Count > 0 Then t = sr4.Count ' CONTOUR
            
            If t = 0 Then Err.Raise vbObjectError, model(m), "Нет позиций на листе"
            
            For j = 1 To t
                If k < COUNTER And CInt(AMOUNT(k)) > 0 Then
                    
                    Set rgb1 = myRGB(MODELCOLOR(k))
                    
                    If j <= sr4.Count Then
                        With sr4(j) ' CONTOUR
                            .GetPosition x1, y1
                            .GetSize w1, h1
                            .Fill.ApplyUniformFill rgb1
                            .Copy
                            .Fill.ApplyNoFill
                        End With
                        lyr6.Paste
                    End If
                    
                    If j <= sr3.Count Then
                        With sr3(j) ' COLORCODE
                            .GetPosition x0, y0
                            .GetSize w0, h0
                            .Outline.SetNoOutline
                            .Text.Story.Replace myCOLORCODE(MODELCOLOR(k))
                        End With
                    End If
                        
                    file = myIMAGE(model(m), LABELCOLOR(k))
                    
                    If fs.FileExists(file) Then
                    
                        symname = mySYMBOL(model(m), LABELCOLOR(k))
                        
                        If FindSymbol(syms, symname) Is Nothing Then
                            Set lyr = FindOrCreateLayer(newPage, "IMPORT")
                            ' регистрируем символ
                            Set impopt = CreateStructImportOptions
                            With impopt
                                .Mode = cdrImportFull
                                .MaintainLayers = True
                                With .ColorConversionOptions
                                    .SourceColorProfileList = "sRGB IEC61966-2.1,ISO Coated v2 (ECI),Dot Gain 15%"
                                    .TargetColorProfileList = "sRGB IEC61966-2.1,ISO Coated v2 (ECI),Dot Gain 15%"
                                End With
                            End With
                            Set impflt = lyr.ImportEx(file, cdrPNG, impopt)
                            With impflt
                                .Finish
                            End With
                            lyr.Shapes.All.ConvertToSymbol symname
                            lyr.Shapes.All.Delete
                            lyr.Delete ' IMPORT
                        End If
                             
                        title = MESSAGE(k)
                        If title <> "" Then
                            Set lyr = FindOrCreateLayer(newPage, "MESSAGE" + MESSAGETEMPLATE(k))
                            If lyr.name = lyr0.name Then
                                Set sr = sr0
                            Else
                                Set sr = lyr.FindShapes(Type:=cdrTextShape) ' MESSAGETEMPLATE<n>
                            End If
                            If j <= sr.Count Then
                                 With sr(j) ' MESSAGE
                                     .GetPosition x0, y0
                                     .GetSize w0, h0
                                     .Text.Story.Replace title
                                     .Text.Story.Font = MESSAGEFONT(k)
                                     .SetPosition x0, y0
                                     .SetSize w0, h0
                                    If Not rgb1.IsSame(myRgbWhite) Then
                                        .Copy
                                        lyr5.Paste ' WHITE
                                    End If
                                     .Outline.SetNoOutline
                                     .Fill.ApplyNoFill
                                    .Copy
                                    lyr0.Paste ' MESSAGE
                                 End With
                                 
                                 Set s = lyr0.FindShape(Type:=cdrTextShape) ' MESSAGE
                                 
                                 With lyr0.CreateSymbol(x0, y0, symname)
                                     .GetPosition x0, y0
                                     .GetSize w0, h0
                                     scaleX = w1 / w0
                                     scaleY = h1 / h0
                                     scaleXY = IIf(scaleX > scaleY, scaleX, scaleY)
                                     .SetPosition x1, y1
                                     .SetSize scaleXY * w0, scaleXY * h0
                                     .AddToPowerClip s, cdrFalse
                                 End With
                                 s.Copy
                                 lyr6.Paste ' RGB
                             End If
                        End If
                        
                        title = SUBMESSAGE(k)
                        If title <> "" Then
                            Set lyr = FindOrCreateLayer(newPage, "SUBMESSAGE" + MESSAGETEMPLATE(k))
                            If lyr.name = lyr1.name Then
                                Set sr = sr1
                            Else
                                Set sr = lyr.FindShapes(Type:=cdrTextShape) ' MESSAGETEMPLATE<n>
                            End If
                            If j <= sr.Count Then
                                 With sr(j) ' SUBMESSAGE
                                     .GetPosition x0, y0
                                     .GetSize w0, h0
                                     .Text.Story.Replace title
                                     .Text.Story.Font = MESSAGEFONT(k)
                                     .SetPosition x0, y0
                                     .SetSize w0, h0
                                    If Not rgb1.IsSame(myRgbWhite) Then
                                        .Copy
                                        lyr5.Paste ' WHITE
                                    End If
                                     .Outline.SetNoOutline
                                     .Fill.ApplyNoFill
                                    .Copy
                                    lyr1.Paste ' MESSAGE
                                 End With
                                 
                                 Set s = lyr1.FindShape(Type:=cdrTextShape) ' SUBMESSAGE
                                 
                                 With lyr1.CreateSymbol(x0, y0, symname)
                                     .GetPosition x0, y0
                                     .GetSize w0, h0
                                     scaleX = w1 / w0
                                     scaleY = h1 / h0
                                     scaleXY = IIf(scaleX > scaleY, scaleX, scaleY)
                                     .SetPosition x1, y1
                                     .SetSize scaleXY * w0, scaleXY * h0
                                     .AddToPowerClip s, cdrFalse
                                 End With
                                 s.Copy
                                 lyr6.Paste ' RGB
                             End If
                        End If
                        
                        title = LOGOTYPE(k)
                        If title <> "" Then
                            Set lyr = FindOrCreateLayer(newPage, "LOGOTYPE" + LOGOTYPETEMPLATE(k))
                            If lyr.name = lyr2.name Then
                                Set sr = sr2
                            Else
                                Set sr = lyr.FindShapes(Type:=cdrTextShape) ' LOGOTYPE<n>
                            End If
                            If j <= sr.Count Then
                                With sr(j) ' LOGOTYPE
                                    .GetPosition x0, y0
                                    .GetSize w0, h0
                                    .Text.Story.Replace title
                                    '.Text.Story.Font = LOGOTYPEFONT(k)
                                    .SetPosition x0, y0
                                    If Not rgb1.IsSame(myRgbWhite) Then
                                        .Copy
                                        lyr5.Paste ' WHITE
                                    End If
                                    .Outline.SetNoOutline
                                    .Fill.ApplyNoFill
                                    .Copy
                                    lyr2.Paste ' LOGOTYPE
                                End With
                                
                                Set s = lyr2.FindShape(Type:=cdrTextShape) ' LOGOTYPE
                                
                                With lyr2.CreateSymbol(x0, y0, symname)
                                    .GetPosition x0, y0
                                    .GetSize w0, h0
                                    scaleX = w1 / w0
                                    scaleY = h1 / h0
                                    scaleXY = IIf(scaleX > scaleY, scaleX, scaleY)
                                    .SetPosition x1, y1
                                    .SetSize scaleXY * w0, scaleXY * h0
                                    .AddToPowerClip s, cdrFalse
                                End With
                                s.Copy
                                lyr6.Paste ' RGB
                            End If
                        End If
                        
                        If j <= sr0.Count Then sr0(j).Delete
                        If j <= sr1.Count Then sr1(j).Delete
                        If j <= sr2.Count Then sr2(j).Delete
                    Else
                        Set rgb = myRGB(LABELCOLOR(k))
                        Set cmyk = myCMYK(LABELCOLOR(k))
                        
                        title = MESSAGE(k)
                        If title <> "" Then
                            Set lyr = FindOrCreateLayer(newPage, "MESSAGE" + MESSAGETEMPLATE(k))
                            If lyr.name = lyr0.name Then
                                Set sr = sr0
                            Else
                                Set sr = lyr.FindShapes(Type:=cdrTextShape) ' MESSAGETEMPLATE<n>
                            End If
                            If j <= sr.Count Then
                                With sr(j) ' MESSAGE
                                    .GetPosition x0, y0
                                    .GetSize w0, h0
                                    .Text.Story.Replace title
                                    .Text.Story.Font = MESSAGEFONT(k)
                                    .SetPosition x0, y0
                                    .SetSize w0, h0
                                    '.Outline.SetNoOutline
                                    .Outline.Color.CopyAssign rgb
                                    .Fill.ApplyUniformFill rgb
                                    .Copy
                                    lyr6.Paste ' RGB
                                    .Outline.Color.CopyAssign cmyk
                                    .Fill.ApplyUniformFill cmyk
                                    .Copy
                                    lyr0.Paste ' MESSAGE
                                End With
                                If Not rgb1.IsSame(myRgbWhite) Then lyr5.Paste ' WHITE
                            End If
                        End If
                        
                        title = SUBMESSAGE(k)
                        If title <> "" Then
                            Set lyr = FindOrCreateLayer(newPage, "SUBMESSAGE" + MESSAGETEMPLATE(k))
                            If lyr.name = lyr1.name Then
                                Set sr = sr1
                            Else
                                Set sr = lyr.FindShapes(Type:=cdrTextShape) ' MESSAGETEMPLATE<n>
                            End If
                            If j <= sr.Count Then
                                With sr(j) ' SUBMESSAGE
                                    .GetPosition x0, y0
                                    .GetSize w0, h0
                                    .Text.Story.Replace title
                                    .Text.Story.Font = MESSAGEFONT(k)
                                    .SetPosition x0, y0
                                    .SetSize w0, h0
                                    '.Outline.SetNoOutline
                                    .Outline.Color.CopyAssign rgb
                                    .Fill.ApplyUniformFill rgb
                                    .Copy
                                    lyr6.Paste ' RGB
                                    .Outline.Color.CopyAssign cmyk
                                    .Fill.ApplyUniformFill cmyk
                                    .Copy
                                    lyr1.Paste ' SUBMESSAGE
                                End With
                                If Not rgb1.IsSame(myRgbWhite) Then lyr5.Paste ' WHITE
                            End If
                        End If
                        
                        title = LOGOTYPE(k)
                        If title <> "" Then
                            Set lyr = FindOrCreateLayer(newPage, "LOGOTYPE" + LOGOTYPETEMPLATE(k))
                            If lyr.name = lyr2.name Then
                                Set sr = sr2
                            Else
                                Set sr = lyr.FindShapes(Type:=cdrTextShape) ' LOGOTYPE<n>
                            End If
                            If j <= sr.Count Then
                                With sr(j) ' LOGOTYPE
                                    .GetPosition x0, y0
                                    .GetSize w0, h0
                                    .Text.Story.Replace title
                                    '.Text.Story.Font = LOGOTYPEFONT(k)
                                    .SetPosition x0, y0
                                    '.Outline.SetNoOutline
                                    .Outline.Color.CopyAssign rgb
                                    .Fill.ApplyUniformFill rgb
                                    .Copy
                                    lyr6.Paste ' RGB
                                    .Outline.Color.CopyAssign cmyk
                                    .Fill.ApplyUniformFill cmyk
                                    .Copy
                                    lyr2.Paste ' LOGOTYPE
                                End With
                                If Not rgb1.IsSame(myRgbWhite) Then lyr5.Paste ' WHITE
                            End If
                        End If
                        
                        If j <= sr0.Count Then sr0(j).Delete
                        If j <= sr1.Count Then sr1(j).Delete
                        If j <= sr2.Count Then sr2(j).Delete
                    End If
                    
                    AMOUNT(k) = CInt(AMOUNT(k)) - 1
                    If CInt(AMOUNT(k)) = 0 Then
                        While k < COUNTER And CInt(AMOUNT(k)) = 0
                            k = k + 1
                        Wend
                    End If
                Else
                    If j <= sr0.Count Then sr0(j).Delete
                    If j <= sr1.Count Then sr1(j).Delete
                    If j <= sr2.Count Then sr2(j).Delete
                    If j <= sr3.Count Then sr3(j).Delete
                    With sr4(j)
                        .Fill.ApplyNoFill
                        .Copy
                    End With
                    lyr6.Paste ' RGB
                End If
                
            Next j
            
On Error Resume Next
            For Each s In lyr5.Shapes.All  ' WHITE
                With s
                    .Outline.Color.CopyAssign myCmykWhite
                    .Fill.ApplyUniformFill myCmykWhite
                End With
            Next s
On Error GoTo errMyErrorHandler
                        
            Set sr0 = Nothing
            Set sr1 = Nothing
            Set sr2 = Nothing
            Set sr3 = Nothing
            Set sr4 = Nothing
            
            n = n + 1
            
        Loop While k < COUNTER
Next m

    STATUS.EndProgress
    
    STATUS.BeginProgress "Сохранение и экспорт", True
    
    For Each newPage In newDoc.Pages
        For Each lyr In newPage.Layers
            lyr.editable = False
            lyr.Printable = False
            lyr.visible = True
            If lyr.name = "WHITE" Then lyr.visible = myExportWhite
            If lyr.name = "MESSAGE" Then lyr.visible = True
            If lyr.name = "SUBMESSAGE" Then lyr.visible = True
            If lyr.name = "LOGOTYPE" Then lyr.visible = True
            If lyr.name = "COLORCODE" Then lyr.visible = True
            If lyr.name = "CONTOUR" Then lyr.visible = True
            If lyr.name = "WHITE" Then lyr.Printable = myExportWhite
            If lyr.name = "MESSAGE" Then lyr.Printable = True
            If lyr.name = "SUBMESSAGE" Then lyr.Printable = True
            If lyr.name = "LOGOTYPE" Then lyr.Printable = True
            If lyr.name = "COLORCODE" Then lyr.Printable = True
            If lyr.name = "CONTOUR" Then lyr.Printable = True
        Next lyr
    Next newPage
    
'Сохранение копии документа
    Dim prj As String
    
    prj = GetProjectName
    newDoc.Activate
    newDoc.SaveAs GetFileName(prj, ".cdr")
    newDoc.Dirty = False
    
    For Each newPage In newDoc.Pages
        newPage.Activate
        file = newPage.name + "-" + prj
        file = Replace(file, "/", "_")
        file = Trim(file)
        
        For Each lyr In newPage.Layers
            lyr.editable = False
            lyr.Printable = False
            lyr.visible = True
            If lyr.name = "MESSAGE" Then lyr.visible = True
            If lyr.name = "SUBMESSAGE" Then lyr.visible = True
            If lyr.name = "LOGOTYPE" Then lyr.visible = True
            If lyr.name = "COLORCODE" Then lyr.visible = True
            If lyr.name = "CONTOUR" Then lyr.visible = True
            If lyr.name = "MESSAGE" Then lyr.Printable = True
            If lyr.name = "SUBMESSAGE" Then lyr.Printable = True
            If lyr.name = "LOGOTYPE" Then lyr.Printable = True
            If lyr.name = "COLORCODE" Then lyr.Printable = True
            If lyr.name = "CONTOUR" Then lyr.Printable = True
        Next lyr
        
        Set export = CreateStructExportOptions
            
        With export
            .AntiAliasingType = cdrNoAntiAliasing
            .ImageType = cdrPalettedImage
            .Overwrite = True
            .SizeX = Width
            .SizeY = Height
            .UseColorProfile = True
            .MaintainAspect = True
            .MaintainLayers = True
        End With
                
If myExportThruAI Then
        Set expflt = newDoc.ExportEx(GetFileName(file, ".ai"), cdrAI, cdrCurrentPage, export)
        With expflt
            .TextAsCurves = True
            .Finish
        End With
Else
        Set expflt = newDoc.ExportEx(GetFileName(file, ".eps"), cdrEPS, cdrCurrentPage, export)
        With expflt
            .TextAsCurves = True
            .BoundingBox = 0 ' FilterEPSLib.epsObjects
            .Finish
        End With
End If

        For Each lyr In newPage.Layers
            lyr.editable = False
            lyr.Printable = False
            lyr.visible = True
            If lyr.name = "RGB" Then lyr.visible = True
            If lyr.name = "RGB" Then lyr.Printable = True
        Next lyr
        
        newDoc.export GetFileName(file, ".jpg"), cdrJPEG, cdrCurrentPage
               
        For Each lyr In newPage.Layers
            lyr.editable = False
            lyr.Printable = False
            lyr.visible = False
            If lyr.name = "MESSAGE" Then lyr.visible = True
            If lyr.name = "SUBMESSAGE" Then lyr.visible = True
            If lyr.name = "LOGOTYPE" Then lyr.visible = True
            If lyr.name = "COLORCODE" Then lyr.visible = True
            If lyr.name = "CONTOUR" Then lyr.visible = True
            If lyr.name = "MESSAGE" Then lyr.Printable = True
            If lyr.name = "SUBMESSAGE" Then lyr.Printable = True
            If lyr.name = "LOGOTYPE" Then lyr.Printable = True
            If lyr.name = "COLORCODE" Then lyr.Printable = True
            If lyr.name = "CONTOUR" Then lyr.Printable = True
            If lyr.name = "RGB" Then lyr.visible = True
        Next lyr
    
    Next newPage
            
    templateDoc.Dirty = False
    newDoc.Dirty = False

If myExportThruAI Then
' illustrator_scripting_reference_vbscript_cs5.pdf
Dim appRef, docRef, epsSaveOptions
Set appRef = CreateObject("Illustrator.Application")
Set epsSaveOptions = CreateObject("Illustrator.EPSSaveOptions")

        Set f = fs.GetFolder(mySaveFolderName)
        Set fc = f.Files
        For Each fl In fc
            Select Case UCase(fs.GetExtensionName(fl))
            Case "AI"
                Set docRef = appRef.Open(fs.GetAbsolutePathName(fl))
                docRef.SaveAs GetFileName(fs.GetBaseName(fl), ".eps"), epsSaveOptions
            End Select
        Next
End If

    STATUS.EndProgress

    Beep
    Exit Sub
    
errMyErrorHandler:
  MsgBox Err.Description, _
    vbExclamation + vbOKCancel, _
    "Error: " & CStr(Err.Number)
  Err.Clear
End Sub
