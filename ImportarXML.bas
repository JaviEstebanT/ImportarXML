Attribute VB_Name = "Módulo2"
Sub ImportarXML()
Attribute ImportarXML.VB_ProcData.VB_Invoke_Func = " \n14"

'Creo las variables de ruta para los archivos de origen y temporales

        Dim Fecha As String
        Fecha = Format(Date, "yyyy-mm-dd")
        Dim Ruta As String
        Ruta = "la ruta que toque"
        Dim ArchOrigen As String
        ArchOrigen = Ruta & Fecha & "-data_export.xml"
        Dim ArchTemp As String
        ArchTemp = Ruta & "temp.xlsx"
        Dim NombreLibro As String
        NombreLibro = InputBox("Introduzca el nombre que desa dar al archivo")

'Importo el archivo xml
         
        Application.DisplayAlerts = False
        Workbooks.OpenXML Filename:=ArchOrigen, _
        LoadOption:=xlXmlLoadImportToList
        ActiveWorkbook.SaveAs (ArchTemp)
        Application.DisplayAlerts = True
        Sheets("hoja1").Select
        Range("A2").Select
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Selection.Copy
        Workbooks("archivo.xlsm").Sheets("hoja1").Activate
        Range("A2").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False

'Limpio los residuos generados
        
        Workbooks("temp.xlsx").Close
        Kill ArchOrigen
        Kill ArchTemp
        Kill Ruta & Fecha & "-data_export.xml.zip"

'Ordeno las filas por fecha de más antiguo a más nuevo

    ActiveWorkbook.Worksheets("hoja1").ListObjects("Tabla13").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("hoja1").ListObjects("Tabla13").Sort.SortFields.Add _
        Key:=Range("Tabla13[startTime]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("hoja1").ListObjects("Tabla13").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Filtro las columnas que no me aportan información

    Columns("N:N").ColumnWidth = 16.71
    Columns("A:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:I").Select
    Selection.EntireColumn.Hidden = True
    Columns("K:M").Select
    Selection.EntireColumn.Hidden = True
    Columns("O:AM").Select
    Selection.EntireColumn.Hidden = True
    Columns("AO:AS").Select
    Selection.EntireColumn.Hidden = True
    Columns("AU:BB").Select
    Selection.EntireColumn.Hidden = True
    Columns("BD:BL").Select
    Selection.EntireColumn.Hidden = True
    Columns("BN:CQ").Select
    Selection.EntireColumn.Hidden = True

'Guardamos libro

    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\ruta que toque\" & NombreLibro & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        

End Sub
