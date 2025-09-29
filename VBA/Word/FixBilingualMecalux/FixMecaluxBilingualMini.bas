Attribute VB_Name = "FixMecaluxBilingualMini"
Option Explicit

' ====== ENTRADA PRINCIPAL ======
Sub ProcessDocs_SegmentAndTransUnitID_Strict()
    Dim rootPath As String
    Dim recursive As VbMsgBoxResult
    Dim files As Collection
    Dim reportData As Collection
    Dim appAlerts As WdAlertLevel
    
    rootPath = PickFolder()
    If Len(rootPath) = 0 Then
        Application.StatusBar = False
        MsgBox "Operación cancelada.", vbInformation
        Exit Sub
    End If
    
    recursive = MsgBox("¿Procesar subcarpetas también (recursivo)?", vbYesNo + vbQuestion, "Procesamiento recursivo")
    
    Set files = New Collection
    EnumerateWordFiles rootPath, (recursive = vbYes), files
    If files.count = 0 Then
        Application.StatusBar = False
        MsgBox "No se han encontrado archivos .doc/.docx/.docm en la carpeta.", vbExclamation
        Exit Sub
    End If
    
    Set reportData = New Collection
    appAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = wdAlertsNone
    Application.ScreenUpdating = False
    
    ' Iniciar barra de estado
    Application.StatusBar = "Iniciando procesamiento..."
    
    Dim i As Long, fpath As String
    For i = 1 To files.count
        fpath = CStr(files(i))
        
        ' Actualizar StatusBar
        Application.StatusBar = "Procesando documento " & i & " de " & files.count & ": " & fpath
        
        On Error GoTo EH_FILE
        Dim statusText As String, commentsRemoved As Boolean
        statusText = ProcessOneDocument_Strict(fpath, commentsRemoved)
        reportData.Add MakeReportRow3(fpath, statusText, IIf(commentsRemoved, "Yes", "No"))
        GoTo CONT_NEXT
EH_FILE:
        reportData.Add MakeReportRow3(fpath, "Error: " & Err.Description, "No")
        Err.Clear
CONT_NEXT:
        DoEvents
    Next i
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = appAlerts

    ' === Normalización opcional de nombres para SDLXLIFF Convertor ===
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Normalizar nombres de archivo para SDLXLIFF Convertor?", vbYesNo + vbQuestion, "SDLXLIFF Convertor")
    If resp = vbYes Then
        Set reportData = NormalizeFilenamesForSDLXLIFF(reportData)
    End If
    
    ' Crear y GUARDAR el informe
    CreateReportDoc3 reportData, rootPath
    
    ' Restaurar StatusBar
    Application.StatusBar = False
    
    MsgBox "Procesamiento finalizado. Se generó y guardó el informe en:" & vbCrLf & rootPath, vbInformation
End Sub

' ====== PROCESO POR DOCUMENTO ======
' Criterio: Fila r (>=2) debe empezar por CStr(r-1). Prefijo -> SegmentID (visible), resto -> TransUnitID (oculto).
' Si alguna fila no empieza por el esperado, se devuelve "Bad SegmentID at row: N" y NO se modifica el documento.
Private Function ProcessOneDocument_Strict(ByVal fpath As String, ByRef commentsRemoved As Boolean) As String
    commentsRemoved = False
    
    Dim doc As Document
    Set doc = Documents.Open(FileName:=fpath, ReadOnly:=False, Visible:=False, AddToRecentFiles:=False)
    
    ' Estilos requeridos
    Dim stySeg As Style, styTrans As Style
    If Not TryGetCharStyle(doc, "SegmentID", stySeg) Then
        doc.Close SaveChanges:=wdDoNotSaveChanges
        ProcessOneDocument_Strict = "Style Missing"
        Exit Function
    End If
    If Not TryGetCharStyle(doc, "TransUnitID", styTrans) Then
        doc.Close SaveChanges:=wdDoNotSaveChanges
        ProcessOneDocument_Strict = "Style Missing"
        Exit Function
    End If
    
    ' Una sola tabla
    If doc.Tables.count <> 1 Then
        doc.Close SaveChanges:=wdDoNotSaveChanges
        ProcessOneDocument_Strict = "Not Single Table"
        Exit Function
    End If
    
    Dim t As Table: Set t = doc.Tables(1)
    
    ' Contenido previo a la tabla
    Dim hasPreContent As Boolean: hasPreContent = (t.Range.Start > doc.Range.Start)
    
    ' Validación: comprobar prefijo esperado y estilos/hidden sin modificar
    Dim needsFix As Boolean: needsFix = False
    Dim totalRows As Long: totalRows = t.Rows.count
    Dim r As Long
    
    If totalRows >= 2 Then
        For r = 2 To totalRows
            Dim val As Variant
            val = ValidateCellRow1_ExpectedPrefixAndStyles(t.cell(r, 1), r, stySeg.NameLocal, styTrans.NameLocal)
            If val(1) = False Then
                doc.Close SaveChanges:=wdDoNotSaveChanges
                ProcessOneDocument_Strict = "Bad SegmentID at row: " & CStr(r)
                Exit Function
            Else
                If val(2) = False Then needsFix = True
            End If
        Next r
    End If
    
    ' Not Changed (todo correcto y nada delante de la tabla)
    If (hasPreContent = False) And (needsFix = False) Then
        If doc.Comments.count > 0 Then
            Do While doc.Comments.count > 0
                doc.Comments(1).Delete
            Loop
            commentsRemoved = True
        End If
        doc.Save
        doc.Close SaveChanges:=wdDoNotSaveChanges
        ProcessOneDocument_Strict = "Not Changed"
        Exit Function
    End If
    
    ' Fixed (borrar contenido previo si lo hay + aplicar estilos correctos)
    If hasPreContent Then
        Dim pre As Range
        Set pre = doc.Range(Start:=doc.Range.Start, End:=t.Range.Start)
        pre.Delete
    End If
    
    If totalRows >= 2 Then
        For r = 2 To totalRows
            ApplyStrictStylesInCell t.cell(r, 1), stySeg, styTrans, r
        Next r
    End If
    
    ' Anchos fijos de columnas (si hay >=4 columnas)
    t.AllowAutoFit = False
    t.PreferredWidthType = wdPreferredWidthPoints
    If t.Columns.count >= 1 Then t.Columns(1).Width = 55.15
    If t.Columns.count >= 2 Then t.Columns(2).Width = 76.35
    If t.Columns.count >= 3 Then t.Columns(3).Width = 305.65
    If t.Columns.count >= 4 Then t.Columns(4).Width = 293.65
    
    ' Fuente Aptos (Body)
    ApplyAptosBody doc
    
    ' Eliminar comentarios
    If doc.Comments.count > 0 Then
        Do While doc.Comments.count > 0
            doc.Comments(1).Delete
        Loop
        commentsRemoved = True
    End If
    
    ' Keywords
    doc.BuiltInDocumentProperties("Keywords").Value = "sidebyside"
    
    doc.Save
    doc.Close SaveChanges:=wdDoNotSaveChanges
    ProcessOneDocument_Strict = "Fixed"
End Function

' ====== VALIDACIÓN (prefijo esperado y estilos/hidden) ======
' Devuelve Variant(1..2): (startsWithExpected As Boolean, stylesOk As Boolean)
' - startsWithExpected: True si la celda empieza por CStr(rowNumber-1)
' - stylesOk: True si el prefijo tiene SegmentID + Hidden=False, y el resto TransUnitID + Hidden=True
Public Function ValidateCellRow1_ExpectedPrefixAndStyles(ByVal cell As cell, ByVal rowNumber As Long, _
                                                         ByVal segStyleName As String, ByVal transStyleName As String) As Variant
    Dim result(1 To 2) As Variant
    result(1) = False: result(2) = False
    
    Dim rngCell As Range: Set rngCell = cell.Range
    If rngCell.End - rngCell.Start <= 1 Then
        ValidateCellRow1_ExpectedPrefixAndStyles = result
        Exit Function
    End If
    rngCell.End = rngCell.End - 1 ' quitar marcador de fin de celda
    
    Dim txt As String: txt = rngCell.text
    Dim expected As String: expected = CStr(rowNumber - 1)
    If Not StartsWith(txt, expected) Then
        ValidateCellRow1_ExpectedPrefixAndStyles = result
        Exit Function
    End If
    
    result(1) = True ' empieza por expected
    
    ' Rangos prefijo/resto
    Dim rngNum As Range, rngRest As Range
    Set rngNum = rngCell.Duplicate
    rngNum.End = rngNum.Start + Len(expected)
    
    Set rngRest = rngCell.Duplicate
    rngRest.Start = rngNum.End
    
    ' Prefijo: estilo SegmentID y Hidden=False
    Dim numOK As Boolean
    numOK = RangeHasUniformCharStyleHidden(rngNum, segStyleName, False)
    
    ' Resto: estilo TransUnitID y Hidden=True
    Dim restOK As Boolean
    If rngRest.Start < rngRest.End Then
        restOK = RangeHasUniformCharStyleHidden(rngRest, transStyleName, True)
    Else
        restOK = True
    End If
    
    result(2) = (numOK And restOK)
    ValidateCellRow1_ExpectedPrefixAndStyles = result
End Function

' ====== APLICACIÓN (aplica prefijo esperado y visibilidad) ======
' Aplica SegmentID al prefijo CStr(row-1) (Hidden=False) y TransUnitID + Hidden=True al resto.
Public Sub ApplyStrictStylesInCell(ByVal cell As cell, ByVal stySeg As Style, ByVal styTrans As Style, ByVal rowNumber As Long)
    On Error GoTo SafeExit
    Dim rngCell As Range: Set rngCell = cell.Range
    If rngCell.End - rngCell.Start <= 1 Then Exit Sub
    rngCell.End = rngCell.End - 1
    
    Dim txt As String: txt = rngCell.text
    Dim expected As String: expected = CStr(rowNumber - 1)
    If Not StartsWith(txt, expected) Then Exit Sub
    
    ' Prefijo esperado -> SegmentID (visible)
    Dim rngNum As Range: Set rngNum = rngCell.Duplicate
    rngNum.End = rngNum.Start + Len(expected)
    rngNum.Style = stySeg
    rngNum.Font.Hidden = False
    
    ' Resto -> TransUnitID (oculto)
    Dim rngRest As Range: Set rngRest = rngCell.Duplicate
    rngRest.Start = rngNum.End
    If rngRest.Start < rngRest.End Then
        rngRest.Style = styTrans
        rngRest.Font.Hidden = True
    End If
SafeExit:
End Sub

' ====== AUX ======
Private Function StartsWith(ByVal s As String, ByVal prefix As String) As Boolean
    If Len(prefix) = 0 Then
        StartsWith = True
    ElseIf Len(s) < Len(prefix) Then
        StartsWith = False
    Else
        StartsWith = (StrComp(Left$(s, Len(prefix)), prefix, vbBinaryCompare) = 0)
    End If
End Function

' True si TODO el rango usa el estilo de carácter indicado
' y el atributo Hidden coincide con mustBeHidden (True/False).
Public Function RangeHasUniformCharStyleHidden(ByVal r As Range, ByVal charStyleName As String, ByVal mustBeHidden As Boolean) As Boolean
    On Error Resume Next
    Dim ch As Range, sty As Variant, s As Style, isHidden As Boolean
    For Each ch In r.Characters
        ' Estilo de carácter
        sty = ch.Style
        If VarType(sty) = vbObject Then
            If sty.Type <> wdStyleTypeCharacter Or sty.NameLocal <> charStyleName Then _
                RangeHasUniformCharStyleHidden = False: Exit Function
        Else
            Set s = Nothing
            Set s = r.Document.Styles(CStr(sty))
            If s Is Nothing Then RangeHasUniformCharStyleHidden = False: Exit Function
            If s.Type <> wdStyleTypeCharacter Or s.NameLocal <> charStyleName Then _
                RangeHasUniformCharStyleHidden = False: Exit Function
        End If
        ' Hidden (en Word: -1 = True, 0 = False)
        isHidden = (ch.Font.Hidden <> 0)
        If isHidden <> mustBeHidden Then RangeHasUniformCharStyleHidden = False: Exit Function
    Next ch
    RangeHasUniformCharStyleHidden = True
End Function

Private Function TryGetCharStyle(ByVal doc As Document, ByVal styleName As String, ByRef outStyle As Style) As Boolean
    On Error Resume Next
    Set outStyle = doc.Styles(styleName)
    If Not outStyle Is Nothing Then
        If outStyle.Type = wdStyleTypeCharacter Then
            TryGetCharStyle = True
        Else
            TryGetCharStyle = False
            Set outStyle = Nothing
        End If
    Else
        TryGetCharStyle = False
    End If
    On Error GoTo 0
End Function

' ====== NORMALIZAR NOMBRES (opcional tras el procesado) ======
Private Function NormalizeFilenamesForSDLXLIFF(ByVal reportData As Collection) As Collection
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim newData As New Collection
    Dim i As Long, rowArr As Variant
    Dim statusText As String, oldPath As String, newPath As String

    For i = 1 To reportData.count
        rowArr = reportData(i)
        statusText = CStr(rowArr(2))
        oldPath = CStr(rowArr(1))
        
        If statusText = "Fixed" Or statusText = "Not Changed" Then
            newPath = BuildRenamedPath(oldPath, fso)
            On Error Resume Next
            Name oldPath As newPath               ' renombrar en disco
            If Err.Number <> 0 Then
                Err.Clear
                newPath = oldPath                 ' si falla, dejamos el original
            End If
            On Error GoTo 0
            rowArr(1) = newPath                   ' actualizar path para el informe
        End If
        
        newData.Add rowArr
    Next i
    
    Set NormalizeFilenamesForSDLXLIFF = newData
End Function

' Reemplaza ".review" final (antes de la extensión) por ".Preview"; si no existe, añade ".Preview".
' Siempre antepone "Generated_". Evita colisiones con sufijos _1, _2...
Private Function BuildRenamedPath(ByVal oldPath As String, ByVal fso As Object) As String
    Dim folder As String, base As String, ext As String
    Dim candidate As String, k As Long
    Dim newBase As String

    On Error Resume Next
    folder = fso.GetParentFolderName(oldPath)
    base = fso.GetBaseName(oldPath)         ' nombre sin extensión
    ext = fso.GetExtensionName(oldPath)     ' extensión sin el punto
    On Error GoTo 0

    If Len(base) >= 7 And LCase$(Right$(base, 7)) = ".review" Then
        newBase = Left$(base, Len(base) - 7) & ".Preview"
    Else
        newBase = base & ".Preview"
    End If

    candidate = fso.BuildPath(folder, "Generated_" & newBase & IIf(ext <> "", "." & ext, ""))

    k = 1
    Do While fso.FileExists(candidate)
        candidate = fso.BuildPath(folder, "Generated_" & newBase & "_" & k & IIf(ext <> "", "." & ext, ""))
        k = k + 1
    Loop

    BuildRenamedPath = candidate
End Function

' ====== APTOS BODY ======
Private Sub ApplyAptosBody(ByVal doc As Document)
    On Error Resume Next
    With doc.Range.Font
        .Name = "Aptos (Body)"
        .NameAscii = "Aptos (Body)"
        .NameOther = "Aptos (Body)"
        .NameFarEast = "Aptos (Body)"
        .NameBi = "Aptos (Body)"
    End With
    Dim st As Style
    Set st = Nothing
    Set st = doc.Styles(wdStyleNormal)
    If Not st Is Nothing Then
        st.Font.Name = "Aptos (Body)"
        st.Font.NameAscii = "Aptos (Body)"
        st.Font.NameOther = "Aptos (Body)"
        st.Font.NameFarEast = "Aptos (Body)"
        st.Font.NameBi = "Aptos (Body)"
    End If
    Dim s As Style
    For Each s In doc.Styles
        Select Case s.Type
            Case wdStyleTypeParagraph, wdStyleTypeCharacter, wdStyleTypeTable
                s.Font.Name = "Aptos (Body)"
                s.Font.NameAscii = "Aptos (Body)"
                s.Font.NameOther = "Aptos (Body)"
                s.Font.NameFarEast = "Aptos (Body)"
                s.Font.NameBi = "Aptos (Body)"
        End Select
    Next s
    On Error GoTo 0
End Sub

' ====== SOPORTE ======
Private Function PickFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Selecciona la carpeta con documentos Word"
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = ""
        End If
    End With
End Function

Private Sub EnumerateWordFiles(ByVal rootPath As String, ByVal recursive As Boolean, ByRef outFiles As Collection)
    Dim fso As Object, folder As Object, file As Object, subf As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(rootPath) Then Exit Sub
    Set folder = fso.GetFolder(rootPath)
    Dim ext As String
    For Each file In folder.files
        ext = LCase$(fso.GetExtensionName(file.Name))
        If ext = "docx" Or ext = "docm" Or ext = "doc" Then
            outFiles.Add file.Path
        End If
    Next file
    If recursive Then
        For Each subf In folder.SubFolders
            EnumerateWordFiles subf.Path, True, outFiles
        Next subf
    End If
End Sub

Private Function MakeReportRow3(ByVal fpath As String, ByVal statusText As String, ByVal commentsFlag As String) As Variant
    Dim arr(1 To 3) As String
    arr(1) = fpath
    arr(2) = statusText
    arr(3) = commentsFlag
    MakeReportRow3 = arr
End Function

' ====== INFORME ======
Private Sub CreateReportDoc3(ByVal reportData As Collection, ByVal saveDir As String)
    Dim rep As Document: Set rep = Documents.Add
    rep.Content.Font.Size = 11
    rep.Content.Font.Name = "Aptos (Body)"
    
    ' Resumen inicial
    Dim totalX As Long, fixedY As Long, notChangedZ As Long, errorsA As Long
    Dim i As Long, rowArr As Variant, statusText As String
    totalX = reportData.count
    For i = 1 To reportData.count
        rowArr = reportData(i)
        statusText = CStr(rowArr(2))
        If statusText = "Fixed" Then fixedY = fixedY + 1
        If statusText = "Not Changed" Then notChangedZ = notChangedZ + 1
    Next i
    errorsA = totalX - fixedY - notChangedZ
    
    Dim sumTbl As Table
    Set sumTbl = rep.Tables.Add(rep.Range(0, 0), 2, 4)
    With sumTbl
        .AllowAutoFit = False
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        With .Borders
            .Enable = True
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = wdLineWidth075pt
            .InsideLineStyle = wdLineStyleSingle
            .InsideLineWidth = wdLineWidth050pt
        End With
        .Range.Font.Size = 11
        .Range.Font.Name = "Aptos (Body)"
        .cell(1, 1).Range.text = "Files processed"
        .cell(1, 2).Range.text = "Files Fixed"
        .cell(1, 3).Range.text = "Not Changed"
        .cell(1, 4).Range.text = "Error in Files"
        .Rows(1).Range.Bold = True
        .cell(2, 1).Range.text = CStr(totalX)
        .cell(2, 2).Range.text = CStr(fixedY)
        .cell(2, 3).Range.text = CStr(notChangedZ)
        With .cell(2, 4).Range
            .text = CStr(errorsA)
            If errorsA > 0 Then .Font.Color = wdColorRed
        End With
    End With
    
    ' Separador
    rep.Range(rep.Content.End - 1, rep.Content.End - 1).InsertParagraphAfter
    
    ' Tabla detalle
    Dim tbl As Table
    Set tbl = rep.Tables.Add(rep.Paragraphs.Last.Range, reportData.count + 1, 3)
    With tbl
        .cell(1, 1).Range.text = "File"
        .cell(1, 2).Range.text = "Status"
        .cell(1, 3).Range.text = "Comments Removed"
        .Rows(1).Range.Bold = True
        .AllowAutoFit = False
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        .Columns(1).PreferredWidth = 60
        .Columns(2).PreferredWidth = 25
        .Columns(3).PreferredWidth = 15
        With .Borders
            .Enable = True
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = wdLineWidth075pt
            .InsideLineStyle = wdLineStyleSingle
            .InsideLineWidth = wdLineWidth050pt
        End With
        .Range.Font.Size = 11
        .Range.Font.Name = "Aptos (Body)"
    End With
    
    ' Rellenar filas
    Dim cellRng As Range
    For i = 1 To reportData.count
        rowArr = reportData(i)
        
        ' (1) File con hipervínculo
        Set cellRng = tbl.cell(i + 1, 1).Range
        cellRng.End = cellRng.End - 1
        cellRng.text = CStr(rowArr(1))
        rep.Hyperlinks.Add Anchor:=cellRng, Address:=CStr(rowArr(1)), _
                            SubAddress:="", ScreenTip:="Open document", _
                            TextToDisplay:=CStr(rowArr(1))
        
        ' (2) Status (verde si Fixed/Not Changed; rojo si error)
        statusText = CStr(rowArr(2))
        With tbl.cell(i + 1, 2).Range
            .text = statusText
            If (statusText = "Fixed") Or (statusText = "Not Changed") Then
                .Font.Color = wdColorGreen
            Else
                .Font.Color = wdColorRed
            End If
        End With
        
        ' (3) Comments Removed
        tbl.cell(i + 1, 3).Range.text = CStr(rowArr(3))
    Next i
    
    ' Guardar automáticamente en la carpeta seleccionada
    Dim sep As String
    sep = IIf(Right$(saveDir, 1) = "\" Or Right$(saveDir, 1) = "/", "", Application.PathSeparator)
    Dim outPath As String
    outPath = saveDir & sep & "BatchReport_SegmentID_" & Format(Now, "yyyymmdd_hhnnss") & ".docx"
    
    On Error Resume Next
    rep.SaveAs2 FileName:=outPath, FileFormat:=wdFormatXMLDocument
    On Error GoTo 0
    
    rep.Activate
End Sub


