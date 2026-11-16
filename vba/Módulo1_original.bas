Attribute VB_Name = "Módulo1"
Option Explicit

'========================================================
' CORRE 2 VECES con distintos MAPA y distintos TEMPLATES
'========================================================
Public Sub GenerarPPT_DosVeces()

    Dim outFolder As String
    Dim templateTottus As String, templatePrecioUno As String

    outFolder = "D:\OneDrive - Falabella\Archivos DMML\Proyectos\1. Automatización Presentaciones\Presentación\"

    templateTottus = outFolder & "PPT_MERCADO_ACTUALIZADO.pptx"
    templatePrecioUno = outFolder & "PPT_MERCADO_PRECIO_UNO.pptx"

    '1) TOTTUS
    GenerarPPT _
        "MAPA", _
        templateTottus, _
        outFolder & "PPT_OUTPUT_TOTTUS_" & Format(Now, "yyyymmdd_hhnn") & ".pptx"

    '2) PRECIO UNO
    GenerarPPT _
        "MAPA_PRECIOUNO", _
        templatePrecioUno, _
        outFolder & "PPT_OUTPUT_PRECIOUNO_" & Format(Now, "yyyymmdd_hhnn") & ".pptx"

End Sub


'========================================================
' MOTOR: 1 corrida (MAPA + TEMPLATE + OUTPUT)
'========================================================
Public Sub GenerarPPT(ByVal mapSheet As String, ByVal pptTemplate As String, ByVal outputFullPath As String)

    Dim wsMap As Worksheet
    Dim lastRow As Long, r As Long

    Dim pptApp As Object, pptPres As Object

    'Registro de Excels embebidos (solo los que abre ChartData)
    Dim dictChartExcelApps As Object
    Set dictChartExcelApps = CreateObject("Scripting.Dictionary")

    On Error GoTo ErrHandler

    Set wsMap = ThisWorkbook.Worksheets(mapSheet)
    lastRow = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 101, , "El MAPA '" & mapSheet & "' no tiene filas."

    Set pptApp = GetPowerPointApp()
    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Open(pptTemplate, ReadOnly:=False, WithWindow:=True)

    ExcelFastMode True

    '---- Loop MAPA ----
    For r = 2 To lastRow
        ProcesarFilaMapa wsMap, r, pptPres, dictChartExcelApps
    Next r

    '---- Redraw fuerte (sin tocar ChartData masivamente) ----
    ForceRedrawPresentation pptPres

    '---- Guardar ----
    
    pptPres.Save
    Application.Wait Now + TimeSerial(0, 0, 10)

CleanExit:
    On Error Resume Next

    'Cerrar SOLO Excels embebidos que se hayan abierto por ChartData
    CloseChartExcelApps_Safe dictChartExcelApps, 1

    ExcelFastMode False
    If Not pptPres Is Nothing Then pptPres.Close

    Set pptPres = Nothing
    Set pptApp = Nothing
    Set wsMap = Nothing
    Set dictChartExcelApps = Nothing

    On Error GoTo 0
    Exit Sub

ErrHandler:
    Debug.Print "ERROR (GenerarPPT - " & mapSheet & "): " & Err.Description
    Resume CleanExit
End Sub


'========================================================
' Procesa 1 fila del MAPA
' A Hoja | B TIPO_VENTA | C ESTADO_TIENDA | D INICIO | E FIN | F GRAFICO | G SLIDE
'========================================================
Private Sub ProcesarFilaMapa(ByVal wsMap As Worksheet, ByVal r As Long, ByVal pptPres As Object, ByVal dictChartExcelApps As Object)

    Dim sheetName As String, tipoVenta As String, estadoTienda As String
    Dim inicioAddr As String, finAddr As String, grafName As String
    Dim slideNum As Long

    Dim wsData As Worksheet
    Dim rngSource As Range

    sheetName = Trim$(CStr(wsMap.Cells(r, "A").Value))
    tipoVenta = CStr(wsMap.Cells(r, "B").Value)
    estadoTienda = CStr(wsMap.Cells(r, "C").Value)
    inicioAddr = Trim$(CStr(wsMap.Cells(r, "D").Value))
    finAddr = Trim$(CStr(wsMap.Cells(r, "E").Value))
    grafName = Trim$(CStr(wsMap.Cells(r, "F").Value))
    slideNum = CLng(Val(wsMap.Cells(r, "G").Value))

    If sheetName = "" Or inicioAddr = "" Or finAddr = "" Or grafName = "" Or slideNum <= 0 Then
        Debug.Print "Fila " & r & " (" & wsMap.Name & "): incompleta, se salta."
        Exit Sub
    End If

    Set wsData = ThisWorkbook.Worksheets(sheetName)

    'Parámetros
    wsData.Range("F3").Value = tipoVenta
    wsData.Range("I3").Value = estadoTienda

    wsData.Calculate
    DoEvents

    Set rngSource = wsData.Range(inicioAddr & ":" & finAddr)

    'Actualizar gráfico (fix fuerte: SetSourceData + series)
    ActualizarChartEnSlide_ForceRender pptPres, slideNum, grafName, rngSource, dictChartExcelApps

End Sub


'========================================================
' Actualiza 1 chart y FUERZA el render (equivale a entrar a “Modificar datos”)
' FIX CLAVE:
'   - SetSourceData para reenganchar el rango
'   - Reasignación de Series (XValues/Values/Name) si aplica
'========================================================
Private Sub ActualizarChartEnSlide_ForceRender(ByVal pptPres As Object, ByVal slideNum As Long, _
                                              ByVal chartShapeName As String, ByVal rngSource As Range, _
                                              ByVal dictChartExcelApps As Object)

    Dim sld As Object, shp As Object
    Dim cdWb As Object, cdWs As Object, xlApp As Object
    Dim rr As Long, cc As Long
    Dim src As Object 'Range

    On Error GoTo ErrHandler

    Set sld = pptPres.Slides(slideNum)

    Set shp = Nothing
    On Error Resume Next
    Set shp = sld.Shapes(chartShapeName)
    On Error GoTo ErrHandler

    If shp Is Nothing Then
        Debug.Print "No se encontró '" & chartShapeName & "' en slide " & slideNum
        Exit Sub
    End If

    If shp.HasChart <> True Then
        Debug.Print "El shape '" & chartShapeName & "' no es chart (slide " & slideNum & ")"
        Exit Sub
    End If

    '----- 0) Ir a la slide (ayuda a renderizar en algunas PCs) -----
    On Error Resume Next
    If pptPres.Application.Windows.Count > 0 Then
        pptPres.Application.ActiveWindow.View.GotoSlide slideNum
    End If
    On Error GoTo ErrHandler
    DoEvents

    '----- 1) “Entrar” a Modificar datos -----
    ActivateChartDataSafe shp

    '----- 2) Tomar workbook embebido -----
    Set cdWb = shp.Chart.ChartData.Workbook
    Set xlApp = cdWb.Application
    TrackExcelApp xlApp, dictChartExcelApps
    Set cdWs = cdWb.Worksheets(1)

    'Ocultar y evitar prompts
    On Error Resume Next
    xlApp.Visible = False
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    xlApp.EnableEvents = False
    On Error GoTo ErrHandler

    rr = rngSource.Rows.Count
    cc = rngSource.Columns.Count

    '----- 3) Pegar data SIN destruir todo a lo bruto (evita que PPT pierda “rastro”) -----
    'Primero limpia un área razonable y luego escribe el tamaño exacto
    cdWs.Cells.Clear
    cdWs.Range("A1").Resize(rr, cc).Value = rngSource.Value

    Set src = cdWs.Range("A1").Resize(rr, cc)

    '----- 4) FIX CLAVE: reenganchar el origen del gráfico -----
    'Esto suele ser lo que se “reactiva” cuando tú haces Modificar datos manual
    On Error Resume Next
    shp.Chart.SetSourceData Source:=src
    On Error GoTo ErrHandler

    '----- 5) (Muy recomendado) Reasignar series: XValues / Values / Name -----
    'Funciona para la mayoría de gráficos de línea/columna/área y también ayuda a XY.
    ForceSeriesBindings shp, cdWs, rr, cc

    '----- 6) Calcular + refresh fuerte -----
    xlApp.CalculateFullRebuild
    DoEvents

    shp.Chart.Refresh
    DoEvents

    'Truco suave para forzar repintado sin cambiar nada visible
    Dim oldLegend As Boolean
    oldLegend = shp.Chart.HasLegend
    shp.Chart.HasLegend = oldLegend

    shp.Chart.Refresh
    DoEvents

    '----- 7) Cerrar workbook embebido -----
    cdWb.Close SaveChanges:=True

    Exit Sub

ErrHandler:
    Debug.Print "Error actualizando '" & chartShapeName & "' (slide " & slideNum & "): " & Err.Description
    On Error Resume Next
    If Not cdWb Is Nothing Then cdWb.Close SaveChanges:=False
    On Error GoTo 0
End Sub


'========================================================
' Fuerza el binding de series contra la tabla del Excel embebido
' Asume formato típico:
'   A = categorías (X) (desde fila 2)
'   fila 1 = headers (series names)
'   B..N = valores (Y)
'========================================================
Private Sub ForceSeriesBindings(ByVal shp As Object, ByVal cdWs As Object, ByVal rr As Long, ByVal cc As Long)

    On Error Resume Next

    If rr < 2 Or cc < 2 Then Exit Sub

    Dim lastRow As Long, lastCol As Long, i As Long
    lastRow = rr
    lastCol = cc

    'SerieCount esperado = columnas - 1 (si A son categorías)
    Dim sc As Long
    sc = shp.Chart.SeriesCollection.Count

    'Si el chart ya tiene series, rebind
    Dim seriesIndex As Long
    seriesIndex = 1

    For i = 2 To lastCol
        If seriesIndex > sc Then Exit For

        With shp.Chart.SeriesCollection(seriesIndex)
            'Nombre de serie (header)
            .Name = cdWs.Cells(1, i).Value

            'XValues = A2:A(lastRow)
            .XValues = cdWs.Range(cdWs.Cells(2, 1), cdWs.Cells(lastRow, 1))

            'Values = Col i (fila 2..lastRow)
            .Values = cdWs.Range(cdWs.Cells(2, i), cdWs.Cells(lastRow, i))
        End With

        seriesIndex = seriesIndex + 1
    Next i

    On Error GoTo 0
End Sub


'========================================================
' Activación segura del ChartData (prefiere ActivateChartDataWindow)
'========================================================
Private Sub ActivateChartDataSafe(ByVal shp As Object)
    On Error Resume Next
    CallByName shp.Chart.ChartData, "ActivateChartDataWindow", VbMethod
    If Err.Number <> 0 Then
        Err.Clear
        shp.Chart.ChartData.Activate
    End If
    On Error GoTo 0
    DoEvents
End Sub


'========================================================
' Redraw fuerte de la presentación (sin abrir ChartData)
'========================================================
Private Sub ForceRedrawPresentation(ByVal pptPres As Object)
    On Error Resume Next

    pptPres.UpdateLinks

    If pptPres.Application.Windows.Count > 0 Then
        'Moverse por slides fuerza repintado
        Dim i As Long
        For i = 1 To pptPres.Slides.Count
            pptPres.Application.ActiveWindow.View.GotoSlide i
            DoEvents
        Next i
        pptPres.Application.ActiveWindow.View.GotoSlide 1
    End If

    On Error GoTo 0
End Sub


'========================================================
' Track + close SOLO Excels embebidos (sin tocar tu Excel principal)
'========================================================
Private Sub TrackExcelApp(ByVal xlApp As Object, ByVal dict As Object)
    If xlApp Is Nothing Then Exit Sub

    On Error Resume Next
    If xlApp.Hwnd = Application.Hwnd Then Exit Sub
    On Error GoTo 0

    Dim k As String
    k = CStr(ObjPtr(xlApp))
    If Not dict.Exists(k) Then dict.Add k, xlApp
End Sub

Private Sub CloseChartExcelApps_Safe(ByVal dict As Object, Optional ByVal waitSeconds As Long = 0)
    If dict Is Nothing Then Exit Sub

    Dim k As Variant, xlApp As Object

    If waitSeconds > 0 Then
        Application.Wait Now + TimeSerial(0, 0, waitSeconds)
    End If

    For Each k In dict.Keys
        Set xlApp = dict(k)

        On Error Resume Next
        If xlApp.Hwnd <> Application.Hwnd Then
            xlApp.Visible = False
            xlApp.DisplayAlerts = False
            xlApp.EnableEvents = False

            'Cerrar solo si no tiene libros
            If xlApp.Workbooks.Count = 0 Then
                xlApp.Quit
            End If
        End If
        On Error GoTo 0
    Next k
End Sub


'========================================================
' PowerPoint App
'========================================================
Private Function GetPowerPointApp() As Object
    Dim pptApp As Object
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    Set GetPowerPointApp = pptApp
End Function


'========================================================
' Excel Fast Mode
'========================================================
Private Sub ExcelFastMode(ByVal enable As Boolean)
    On Error Resume Next
    If enable Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub


'========================================================
' CORRE 2 VECES: TOTTUS y PRECIO UNO (solo METRICAS, sin guardar)
'========================================================
Public Sub Test_ProcesarTodoMetricas_Ambos()

    Dim outFolder As String
    outFolder = "D:\OneDrive - Falabella\Archivos DMML\Proyectos\1. Automatización Presentaciones\Presentación\"

    '1) TOTTUS
    Test_ProcesarMetricas _
        "METRICAS", _
        outFolder & "PPT_MERCADO_ACTUALIZADO.pptx"

    '2) PRECIO UNO
    Test_ProcesarMetricas _
        "METRICAS_PRECIOUNO", _
        outFolder & "PPT_MERCADO_PRECIO_UNO.pptx"

End Sub


'========================================================
' MOTOR DE PRUEBA: 1 corrida (HOJA_METRICAS + TEMPLATE), sin guardar
'========================================================
Private Sub Test_ProcesarMetricas(ByVal metSheet As String, ByVal pptTemplate As String)

    Dim wsMet As Worksheet
    Dim pptApp As Object, pptPres As Object
    Dim lastRowMet As Long, rMet As Long

    On Error GoTo ErrHandler

    Set wsMet = ThisWorkbook.Worksheets(metSheet)
    lastRowMet = wsMet.Cells(wsMet.Rows.Count, "A").End(xlUp).Row

    If lastRowMet < 2 Then
        MsgBox "La hoja " & metSheet & " no tiene filas de datos."
        Exit Sub
    End If

    Set pptApp = GetPowerPointApp()
    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Open(pptTemplate, ReadOnly:=False, WithWindow:=True)

    For rMet = 2 To lastRowMet
        ProcesarFilaMetricas wsMet, rMet, pptPres
    Next rMet

    MsgBox "Prueba OK (" & metSheet & "). " & (lastRowMet - 1) & " filas procesadas." & vbCrLf & _
           "Revisa el PPT abierto — NO se guardó nada."

CleanExit:
    'No cerramos pptPres a propósito: se queda abierto para que revises
    Set pptPres = Nothing
    Set pptApp = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error en prueba (" & metSheet & ", fila " & rMet & "): " & Err.Description
    Resume CleanExit
End Sub


'========================================================
' Procesa 1 fila de METRICAS
' A Hoja | B TIPO_VENTA | C ESTADO_TIENDA | D METRICA (celda) | E CUADRO | F SLIDE
'========================================================
Private Sub ProcesarFilaMetricas(ByVal wsMet As Worksheet, ByVal r As Long, ByVal pptPres As Object)

    Dim sheetName As String, tipoVenta As String, estadoTienda As String
    Dim celdaAddr As String, cuadroName As String
    Dim slideNum As Long
    Dim wsData As Worksheet
    Dim valorTexto As String

    sheetName = Trim$(CStr(wsMet.Cells(r, "A").Value))
    tipoVenta = CStr(wsMet.Cells(r, "B").Value)
    estadoTienda = CStr(wsMet.Cells(r, "C").Value)
    celdaAddr = Trim$(CStr(wsMet.Cells(r, "D").Value))
    cuadroName = Trim$(CStr(wsMet.Cells(r, "E").Value))
    slideNum = CLng(Val(wsMet.Cells(r, "F").Value))

    If sheetName = "" Or celdaAddr = "" Or cuadroName = "" Or slideNum <= 0 Then
        Debug.Print "Fila " & r & " (" & wsMet.Name & "): incompleta, se salta."
        Exit Sub
    End If

    Set wsData = ThisWorkbook.Worksheets(sheetName)

    wsData.Range("F3").Value = tipoVenta
    wsData.Range("I3").Value = estadoTienda

    wsData.Calculate
    DoEvents

    'Uso .Text (no .Value) para respetar el formato numérico de origen (%, decimales, etc.)
    valorTexto = wsData.Range(celdaAddr).Text

    ActualizarTextoEnSlide_ForceRender pptPres, slideNum, cuadroName, valorTexto

End Sub


'========================================================
' Actualiza el TEXTO de un shape (cuadro/textbox) en una slide,
' respetando el formato de destino (fuente, tamaño, color, negrita)
'========================================================
Private Sub ActualizarTextoEnSlide_ForceRender(ByVal pptPres As Object, ByVal slideNum As Long, _
                                              ByVal shapeName As String, ByVal nuevoValor As String)

    Dim sld As Object, shp As Object

    On Error GoTo ErrHandler

    Set sld = pptPres.Slides(slideNum)

    Set shp = Nothing
    On Error Resume Next
    Set shp = sld.Shapes(shapeName)
    On Error GoTo ErrHandler

    If shp Is Nothing Then
        Debug.Print "No se encontró '" & shapeName & "' en slide " & slideNum
        Exit Sub
    End If

    If shp.HasTextFrame <> True Then
        Debug.Print "El shape '" & shapeName & "' no tiene TextFrame (slide " & slideNum & ")"
        Exit Sub
    End If

    'Solo cambia el texto del run, no toca fuente/tamaño/negrita
    shp.TextFrame.TextRange.Text = nuevoValor

    '----- Color condicional SOLO si hay '+' o '-' literal en el texto -----
    If InStr(1, nuevoValor, "+", vbBinaryCompare) > 0 Then
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(4, 95, 52)   '#045F34 verde
    ElseIf InStr(1, nuevoValor, "-", vbBinaryCompare) > 0 Then
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(192, 0, 0)   '#C00000 rojo
    End If
    'Si no tiene ni '+' ni '-', se deja el color que ya traía el shape en el template

    Exit Sub

ErrHandler:
    Debug.Print "Error actualizando texto '" & shapeName & "' (slide " & slideNum & "): " & Err.Description
End Sub
