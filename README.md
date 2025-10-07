Option Explicit

'======================
' CONFIGURACIÓN
'======================
Private Const NOMBRE_HOJA_CONSOLIDADO As String = "Consolidado"
' Patrones de nombre de archivo a buscar (en el nombre del archivo)
Private FILE_PATTERNS As Variant

'======================
' ENTRADA PRINCIPAL
'======================
Public Sub Consolidar_TRF_VNT_CHQ()
    Dim carpeta As String
    Dim archivos As Collection
    Dim wbDestino As Workbook, wsDest As Worksheet
    Dim encabezadoCopiado As Boolean
    Dim ruta As Variant
    Dim rutaGuardado As String, timestamp As String
    
    FILE_PATTERNS = Array("trf", "vnt", "chq") ' patrones de nombre

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    carpeta = ElegirCarpeta()
    If Len(carpeta) = 0 Then
        MsgBox "No se seleccionó carpeta.", vbInformation
        GoTo Salir
    End If

    Set archivos = BuscarArchivosPorPatron(carpeta, FILE_PATTERNS)
    If archivos Is Nothing Or archivos.Count = 0 Then
        MsgBox "No se encontraron archivos que coincidan con los patrones (trf, vnt, chq).", vbExclamation
        GoTo Salir
    End If

    ' Crear libro destino y hoja consolidada
    Set wbDestino = Workbooks.Add
    On Error Resume Next
    Set wsDest = wbDestino.Worksheets(NOMBRE_HOJA_CONSOLIDADO)
    On Error GoTo 0
    If wsDest Is Nothing Then
        Set wsDest = wbDestino.Worksheets(1)
        wsDest.Name = NOMBRE_HOJA_CONSOLIDADO
    Else
        wsDest.Cells.Clear
    End If

    encabezadoCopiado = False

    ' Procesar cada archivo encontrado
    For Each ruta In archivos
        ProcesarYVolcar ruta, wsDest, encabezadoCopiado
        encabezadoCopiado = True ' a partir del primero, ya no copiamos encabezado
    Next ruta

    ' Ajustes finales
    AutoFitDestino wsDest

    ' === Guardar automáticamente en la misma carpeta con timestamp ===
    timestamp = Format(Now, "yyyy-mm-dd_hh-nn")
    rutaGuardado = AgregarSlash(carpeta) & "Consolidado_" & timestamp & ".xlsx"
    
    Application.DisplayAlerts = False
    wbDestino.SaveAs Filename:=rutaGuardado, FileFormat:=xlOpenXMLWorkbook ' XLSX
    Application.DisplayAlerts = True

    MsgBox "✅ Consolidado creado y guardado en:" & vbCrLf & rutaGuardado, vbInformation

Salir:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume Salir
End Sub

'======================
' PROCESO POR ARCHIVO
'======================
Private Sub ProcesarYVolcar(ByVal ruta As String, ByRef wsDestino As Worksheet, ByVal encabezadoYaCopiado As Boolean)
    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim origen As Range, datos As Range
    Dim filaIni As Long, filas As Long, cols As Long
    Dim destFila As Long

    On Error GoTo LocalErr
    Set wb = Workbooks.Open(Filename:=ruta, ReadOnly:=True)
    Set ws = wb.Worksheets(1) ' primera hoja; ajusta si necesitas una hoja por nombre

    ' ====== FORMATEOS SOLICITADOS ======
    ' 1) Eliminar fila 1 y 2
    If ws.UsedRange.Rows.Count >= 2 Then ws.Rows("1:2").Delete

    ' 2) Eliminar SOLO la celda A1 y correr a la izquierda (alinear títulos)
    ws.Range("A1").Delete Shift:=xlToLeft

    ' 3) Borrar filas según columna C (desde la fila 2 hacia abajo)
    lastRow = UltimaFila(ws)
    If lastRow >= 2 Then
        Dim r As Long, valC As String
        For r = lastRow To 2 Step -1
            valC = LCase$(Trim$(CStr(ws.Cells(r, 3).Value))) ' col C
            If (valC = "") Or (InStr(valC, "reporte") > 0) Or (InStr(valC, "beneficiario") > 0) Then
                ws.Rows(r).Delete
            End If
        Next r
    End If

    ' Detectar rango final con datos
    lastRow = UltimaFila(ws)
    lastCol = UltimaColumna(ws)
    If lastRow = 0 Or lastCol = 0 Then GoTo Cerrar

    ' Determinar qué copiar (con o sin encabezado)
    If encabezadoYaCopiado Then
        If lastRow < 2 Then GoTo Cerrar ' no hay datos
        filaIni = 2
    Else
        filaIni = 1
    End If

    filas = lastRow - filaIni + 1
    cols = lastCol
    Set datos = ws.Range(ws.Cells(filaIni, 1), ws.Cells(lastRow, cols))

    ' Volcar valores al destino (solo valores)
    destFila = SiguienteFilaLibre(wsDestino)
    wsDestino.Cells(destFila, 1).Resize(filas, cols).Value = datos.Value

Cerrar:
    wb.Close SaveChanges:=False
    Exit Sub

LocalErr:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise Err.Number, , Err.Description
End Sub

'======================
' BUSCAR ARCHIVOS
'======================
Private Function BuscarArchivosPorPatron(ByVal carpeta As String, ByVal patrones As Variant) As Collection
    Dim col As New Collection
    Dim f As String, ext As Variant, ok As Boolean, nm As String
    Dim extensiones As Variant
    extensiones = Array("*.xlsx", "*.xlsm", "*.xlsb", "*.xls")

    Dim base As String
    base = AgregarSlash(carpeta)

    For Each ext In extensiones
        f = Dir(base & ext, vbNormal)
        Do While Len(f) > 0
            nm = LCase$(f)
            ok = False
            Dim p As Variant
            For Each p In patrones
                If InStr(nm, LCase$(CStr(p))) > 0 Then
                    ok = True: Exit For
                End If
            Next p

            If ok Then col.Add base & f
            f = Dir
        Loop
    Next ext

    If col.Count > 0 Then
        Set BuscarArchivosPorPatron = col
    Else
        Set BuscarArchivosPorPatron = Nothing
    End If
End Function

Private Function ElegirCarpeta() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Selecciona la carpeta que contiene TRF, VNT y CHQ"
        If .Show = -1 Then
            ElegirCarpeta = .SelectedItems(1)
        Else
            ElegirCarpeta = vbNullString
        End If
    End With
End Function

'======================
' DESTINO / UTILIDADES
'======================
Private Sub AutoFitDestino(ByVal ws As Worksheet)
    Dim lc As Long
    lc = UltimaColumna(ws)
    If lc > 0 Then ws.Columns("A:" & ColLetter(lc)).AutoFit
End Sub

Private Function SiguienteFilaLibre(ByVal ws As Worksheet) As Long
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr = 1 And WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
        SiguienteFilaLibre = 1
    Else
        SiguienteFilaLibre = lr + 1
    End If
End Function

Private Function UltimaFila(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
        If WorksheetFunction.CountA(ws.Cells) = 0 Then UltimaFila = 0 Else UltimaFila = r
    Else
        UltimaFila = r
    End If
End Function

Private Function UltimaColumna(ByVal ws As Worksheet) As Long
    Dim c As Long
    c = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If c < 1 Then UltimaColumna = 0 Else UltimaColumna = c
End Function

Private Function ColLetter(ByVal colNum As Long) As String
    ColLetter = Split(Cells(1, colNum).Address(True, False), "$")(0)
End Function

Private Function AgregarSlash(ByVal path As String) As String
    If Right$(path, 1) = "\" Or Right$(path, 1) = "/" Then
        AgregarSlash = path
    Else
        AgregarSlash = path & "\"
    End If
End Function
