Attribute VB_Name = "Module5"
' ƒмитрий ѕротопопов
' dmitry@protopopov.ru

' ѕоиск сло€ по имени или создание сло€ с данным именем, если не удалось найти слой по имени
Function FindPage(doc As Document, name As String) As Integer
    Dim i As Integer
    On Error Resume Next
    For i = 1 To doc.Pages.Count
        If UCase(doc.Pages(i).name) = UCase(name) Then
            FindPage = i
            Exit Function
        End If
    Next
    FindPage = 0
End Function

' ѕоиск сло€ по имени или создание сло€ с данным именем, если не удалось найти слой по имени
Function FindOrCreateLayer(pg As Page, name As String) As Layer
    If pg.Layers.Find(name) Is Nothing Then pg.CreateLayer (name)
    Set FindOrCreateLayer = pg.Layers.Find(name)
End Function

' ѕоиск символа по имени
Function FindSymbol(syms As SymbolDefinitions, name As String) As SymbolDefinition
    On Error Resume Next
    Set FindSymbol = syms(name)
End Function

'  опирование сло€ со страницы на страницу в пределах одного документа
Sub CopyLayer(name As String, pg1 As Page, pg2 As Page)
    Dim lyr1 As Layer, lyr2 As Layer
    Dim editable As Boolean
    Dim visible As Boolean
    On Error Resume Next
    Set lyr1 = pg1.Layers(name)
    editable = lyr1.editable
    visible = lyr1.visible
    lyr1.editable = True
    lyr1.visible = True
    Set lyr2 = FindOrCreateLayer(pg2, name)
    lyr2.editable = True
    lyr2.visible = True
    lyr1.Shapes.All.CopyToLayer lyr2
    lyr1.editable = editable
    lyr1.visible = visible
End Sub

'  опирование сло€ со страницы одного документа на страницу другого документа
Sub CopyLayer2(name As String, pg1 As Page, pg2 As Page)
    Dim lyr1 As Layer, lyr2 As Layer
    On Error Resume Next
    Set lyr1 = pg1.Layers(name)
    Set lyr2 = FindOrCreateLayer(pg2, name)
    lyr1.Shapes.All.Copy
    lyr2.Paste
    ' Clipboard.Clear
End Sub

'  опирование страницы в пределах одного документа
Sub CopyPage(pg1 As Page, pg2 As Page)
    Dim i As Integer
    Dim lyr As Layer
    Dim editable As Boolean
    Dim visible As Boolean
    For i = pg1.Layers.Count To 1 Step -1
        Set lyr = pg1.Layers.Item(i)
        If (Not lyr.IsSpecialLayer) _
        Then
            editable = lyr.editable
            visible = lyr.visible
            lyr.editable = True
            lyr.visible = True
            CopyLayer lyr.name, pg1, pg2
            lyr.editable = editable
            lyr.visible = visible
        End If
    Next
End Sub

'  опирование страницы одного документа на страницу другого документа
Sub CopyPage2(pg1 As Page, pg2 As Page)
    Dim i As Integer
    Dim lyr As Layer
    Dim editable As Boolean
    Dim visible As Boolean
    For i = pg1.Layers.Count To 1 Step -1
        Set lyr = pg1.Layers.Item(i)
        If (Not lyr.IsSpecialLayer) _
        Then
            editable = lyr.editable
            visible = lyr.visible
            lyr.editable = True
            lyr.visible = True
            CopyLayer2 lyr.name, pg1, pg2
            lyr.editable = editable
            lyr.visible = visible
        End If
    Next
End Sub

