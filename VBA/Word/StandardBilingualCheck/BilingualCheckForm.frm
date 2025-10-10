VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BilingualCheckForm 
   Caption         =   "Trados Bilingual Check"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   OleObjectBlob   =   "BilingualCheckForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BilingualCheckForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ArrayCellsIncorrectTagStyle() As Integer 'Esta será global numero de celdas con estilo fake, Array
Dim NumberCellIncorrectTagStyle As Long 'Contiene el numero de celdas que tienen caracteres Tag y no son tags
Dim NumberCellsStrikeUnder As Long 'numero celdas con strike o under encontrados
Dim ArrayCellsStrikeUnder() As Integer 'Esta será global numero de celdas con estilo fake, Array
Dim NumberCellsIncorrectStyle As Integer 'Esta será global numero de celdas con tags sin estilo, Array
Dim WrongTagsList() As Variant 'Esta sera global array con posicion de celda en columna 4 y listado de tags que no tienen estilo
Dim contaTagsCell As Long 'Contador del navegador de celda tags; se inicializa a -1 por la entrada en primera vez
Dim contaTagsCellStyle As Long 'Contador del navegador de celda tags STYLE; se inicializa a -1 por la entrada en primera vez
Dim contaStrikeUnd As Long 'Contador del navegador de celda Strike-Underline; se inicializa a -1 por la entrada en primera vez
Dim contaWrongTagStyle As Long 'Contador del navegador de celda Wrong Tag Style; se inicializa a -1 por la entrada en primera vez
Dim celdasConRetorno() As cell
Dim Col3Col4() As Integer
Dim CeldaCurr As Integer
Dim resultsArray() As Variant
Dim numRowsInTemp As Long
Dim TrackChangesActive As Boolean
Dim ShowRevisionActive As Boolean
Dim SalvadoDoc As Boolean
Dim CrAtBeg As Boolean
Dim CrAtEnd As Boolean
Dim CrAtMid As Boolean
Dim startPos As Integer '
Dim endPos As Integer '


Private Sub btnCheckStart_Click()
    Dim tabla As Table
    Dim numTablas As Integer
    Dim numColumnas As Integer
    Dim celda As cell
    Dim tieneRetorno As Boolean
    Dim numCeldasConRetorno As Integer
    Dim cellcontent As String
    Dim newcontent As String
    Dim i As Integer
    Dim TempArray() As Integer
    Dim rngBefore As Range
    
    
    'Inicializa Variables
    ReDim Col3Col4(1 To 1, 1 To 2) As Integer
    CeldaCurr = 0
    
    'Inicializa valores por si es la segunda vez que lo corremos
    lblCurrCell.Caption = ""
    lblCurrCellTag.Caption = ""
    lblCurrStyleTagCell.Caption = ""
    lblCurrentStrikeUnder.Caption = ""
    lblCurrCellStyleOutside.Caption = ""
    txtNumTagStyleOutsideTag.Caption = ""
    btnCopyColumn.Enabled = False
    
    txtTickNumTables = "?"
    txtTickColNumTables = "?"
    txtTickCarrRet = "?"
    txtSTickSrikeUnder = "?"
    txtTickIncTagStyle = "?"
    txtTickMissExtraTags = "?"
    txtTickTagStyleOutside = "?"
    
    
    'Tamaño Form
    BilingualCheckForm.Width = 346
    'Botones
    btnShowMissExtraTags.Enabled = False
    btnFixSingle.Enabled = False
    btnShowTagStyle.Enabled = False
    btnFixTagStyleOut.Enabled = False
    btnFixGlobalStyles.Enabled = False
    btnTagStyleFord.Enabled = False
    'Radio
    optMLB.Enabled = False
    OptSpace.Enabled = False
  
  
   'Verificar que no hay contenido antes de la tabla.
    Dim answer1 As Integer
    Set rngBefore = ActiveDocument.Range(Start:=0, End:=ActiveDocument.Tables(1).Range.Start)
        If rngBefore.text > "" Then
            answer1 = MsgBox("There is content before the table. It should not exist." & vbCrLf & vbCrLf & "You should exit, fix manually and rerun macro." & vbCrLf & vbCrLf & "Do you want to continue running macro?", vbQuestion + vbYesNo)
            If answer1 = vbNo Then End
        End If
  
  
    'Track Changes off
    ActiveDocument.TrackRevisions = False
        lblTrackChanges.Caption = "Track Changes OFF"
        lblTrackChanges.ForeColor = RGB(0, 128, 0)
    
    'Marca el boton como inactivo
    btnCheckStart.Enabled = False
    
    ' Verificar si el documento tiene solo una tabla
    numTablas = ActiveDocument.Tables.count
    txtNumTables.Caption = numTablas
    If numTablas <> 1 Then
        txtTickNumTables = "X"
        txtTickNumTables.ForeColor = RGB(255, 0, 0)
        txtTickNumTables.Font.Bold = True
        MsgBox "This document should contain one single table."
        Exit Sub
    Else
        txtTickNumTables = ChrW(&H2713)
        txtTickNumTables.ForeColor = RGB(0, 255, 0)
        txtTickNumTables.Font.Bold = True
    End If
    
    ' Verificar si la tabla tiene exactamente cuatro columnas
    Set tabla = ActiveDocument.Tables(1)
    numColumnas = tabla.Columns.count
    txtColNum.Caption = numColumnas
    If numColumnas <> 4 Then
        txtColNum = numColumnas
        txtTickColNumTables = "X"
        txtTickColNumTables.ForeColor = RGB(255, 0, 0)
        txtTickColNumTables.Font.Bold = True
        MsgBox "Table does not contains 4 columns."
        Exit Sub
    Else
        txtTickColNumTables = ChrW(&H2713)
        txtTickColNumTables.ForeColor = RGB(0, 255, 0)
        txtTickColNumTables.Font.Bold = True
    End If
    DoEvents
    
     
    'LLAMA A INSERTIONS AND DELETIONS
     CheckInsertionsDeletions
        
    
' VERIFICAR CELDAS CON RETORNO SI ACTIVADO
If chkCRonCELL.Value = True Then
    
        ReDim celdasConRetorno(1 To 1) As cell ' Inicializar el array con un elemento
        ReDim Colum3Ret(1 To 1) As Integer
        ReDim Colum4Ret(1 To 1) As Integer
        
        numCeldasConRetorno = 0 ' Contador de celdas con retorno
        
        For Each celda In tabla.Columns(3).Cells
            'Quita los dos ultimos caracteres para verificar que tiene retornos de carro Chap GP no se entera
            cellcontent = celda.Range.text
            If Len(cellcontent) > 2 Then
                newcontent = Left(cellcontent, Len(cellcontent) - 2)
            Else
                newcontent = ""
            End If
            
            tieneRetorno = InStr(newcontent, vbCr) > 0
            If tieneRetorno Then
                numCeldasConRetorno = numCeldasConRetorno + 1
                
                ' Crear una nueva matriz temporal con el tamaño actualizado
                ReDim TempArray(1 To numCeldasConRetorno, 1 To 2)
                
                ' Copiar los valores existentes a la matriz temporal
                For i = 1 To UBound(Col3Col4, 1)
                    For j = 1 To 2
                        TempArray(i, j) = Col3Col4(i, j)
                    Next j
                Next i
                
                ' Asignar la nueva matriz temporal a miArray
                Col3Col4 = TempArray
                
                Col3Col4(numCeldasConRetorno, 1) = 3
                Col3Col4(numCeldasConRetorno, 2) = celda.rowIndex
                
                
                ReDim Preserve celdasConRetorno(1 To numCeldasConRetorno) As cell ' Redimensionar el array para agregar una nueva celda
                Set celdasConRetorno(numCeldasConRetorno) = celda ' Almacenar la celda con retorno
                
                
            End If
        Next celda
        
        For Each celda In tabla.Columns(4).Cells
            'Quita los dos ultimos caracteres para verificar que tiene retornos de carro Chap GP no se entera
            cellcontent = celda.Range.text
            If Len(cellcontent) > 2 Then
                newcontent = Left(cellcontent, Len(cellcontent) - 2)
            Else
                newcontent = ""
            End If
            
            tieneRetorno = InStr(newcontent, vbCr) > 0
            If tieneRetorno Then
               
                numCeldasConRetorno = numCeldasConRetorno + 1
                
                ' Crear una nueva matriz temporal con el tamaño actualizado
                ReDim TempArray(1 To numCeldasConRetorno, 1 To 2)
                
                ' Copiar los valores existentes a la matriz temporal
                For i = 1 To UBound(Col3Col4, 1)
                    For j = 1 To 2
                        TempArray(i, j) = Col3Col4(i, j)
                    Next j
                Next i
                
                ' Asignar la nueva matriz temporal a miArray
                Col3Col4 = TempArray
                
                'Mete un nuevo valor
                Col3Col4(numCeldasConRetorno, 1) = 4
                Col3Col4(numCeldasConRetorno, 2) = celda.rowIndex
                
                
                
                ReDim Preserve celdasConRetorno(1 To numCeldasConRetorno) As cell ' Redimensionar el array para agregar una nueva celda
                Set celdasConRetorno(numCeldasConRetorno) = celda ' Almacenar la celda con retorno
                
            End If
        Next celda
        

        
        ' Mostrar resultados
        If numCeldasConRetorno > 0 Then
            txtCarrret = numCeldasConRetorno
            txtTickCarrRet = "X"
            txtTickCarrRet.ForeColor = RGB(255, 0, 0)
            txtTickCarrRet.Font.Bold = True
        Else
            txtCarrret = numCeldasConRetorno
            txtTickCarrRet = ChrW(&H2713)
            txtTickCarrRet.ForeColor = RGB(0, 255, 0)
            txtTickCarrRet.Font.Bold = True
            
        End If
                
Else
        
    txtTickCarrRet = "N/A"
    txtTickCarrRet.Font.Bold = True
    txtTickCarrRet.ForeColor = RGB(255, 0, 0)
    
End If
    
    
    DoEvents
    
    'Enable navegador
    If numCeldasConRetorno > 0 Then
        CarrBack.Enabled = False
        CarrFord.Enabled = True
    End If
    
    
    
    
'LLAMA A COMPROBAR UNDERLINE Y STRIKETROUGH  SI ACTIVADO
If chkUnderStrikeT.Value = True Then
        
        StrikeUnderline
        
        If NumberCellsStrikeUnder > 0 Then
            txtNumStriUnder.Caption = NumberCellsStrikeUnder
            txtSTickSrikeUnder = "X"
            txtSTickSrikeUnder.ForeColor = RGB(255, 0, 0)
            txtSTickSrikeUnder.Font.Bold = True
            
            btnStrikeUnderFord.Enabled = True
            
            'Contador de posicion filas a cero
            contaStrikeUnd = 0
            
        Else
            txtSTickSrikeUnder.Caption = NumberCellsStrikeUnder
            txtSTickSrikeUnder = ChrW(&H2713)
            txtSTickSrikeUnder.ForeColor = RGB(0, 255, 0)
            txtSTickSrikeUnder.Font.Bold = True
        End If
    
Else
       txtSTickSrikeUnder = "N/A"
       txtSTickSrikeUnder.Font.Bold = True
       txtSTickSrikeUnder.ForeColor = RGB(255, 0, 0)
End If
    DoEvents
    
    
    
'LLAMA A COMPROBAR LOS ESTILOS de LAS TAGS
    
If chkTagsIncorrStyle.Value = True Then
    
    CheckTagStyleInColumn4
    
    If NumberCellsIncorrectStyle > 0 Then
        txtNumIncTagsStyle.Caption = NumberCellsIncorrectStyle
        txtTickIncTagStyle = "X"
        txtTickIncTagStyle.ForeColor = RGB(255, 0, 0)
        txtTickIncTagStyle.Font.Bold = True
        
        btnFordStyleTag.Enabled = True
        
        'Contador de posicion filas a cero
        contaTagsCellStyle = -1
        
    Else
        txtTickIncTagStyle.Caption = NumberCellsIncorrectStyle
        txtTickIncTagStyle = ChrW(&H2713)
        txtTickIncTagStyle.ForeColor = RGB(0, 255, 0)
        txtTickIncTagStyle.Font.Bold = True
    End If
    DoEvents
Else
    txtTickIncTagStyle = "N/A"
    txtTickIncTagStyle.Font.Bold = True
    txtTickIncTagStyle.ForeColor = RGB(255, 0, 0)
End If


'LLAMA A COMPROBAR LAS TAGS MISSING EXTRA TAGS

If chkExtraMissTags.Value = True Then
    CheckTagsInWordTable
    
    
    If resultsArray(0, 1) > "" Then
        txtNumExtraTags.Caption = UBound(resultsArray) + 1
        txtTickMissExtraTags = "X"
        txtTickMissExtraTags.ForeColor = RGB(255, 0, 0)
        txtTickMissExtraTags.Font.Bold = True
        
        btnTagFord.Enabled = True
        'Contador de posicion filas a cero
        contaTagsCell = -1
        
    Else
        txtNumExtraTags.Caption = UBound(resultsArray)
        txtTickMissExtraTags = ChrW(&H2713)
        txtTickMissExtraTags.ForeColor = RGB(0, 255, 0)
        txtTickMissExtraTags.Font.Bold = True
    End If
    DoEvents
Else
    txtTickMissExtraTags = "N/A"
    txtTickMissExtraTags.Font.Bold = True
    txtTickMissExtraTags.ForeColor = RGB(255, 0, 0)
End If
    
    
    
'LLAMA A COMPROBAR TAG STYLE FUERA DE TAGS
    
If chkTagStyleOutTag.Value = True Then
        MarkTagStyleOutAsyellow
        TagStyleOutAsyellowDetect
        
        txtNumTagStyleOutsideTag.Caption = NumberCellIncorrectTagStyle
    
        If NumberCellIncorrectTagStyle > 0 Then
            txtTickTagStyleOutside = "X"
            txtTickTagStyleOutside.ForeColor = RGB(255, 0, 0)
            txtTickTagStyleOutside.Font.Bold = True
            
            btnTagStyleFord.Enabled = True
            
            
            'Contador de posicion filas a cero
            contaWrongTagStyle = -1
            
        Else
            txtTickTagStyleOutside = ChrW(&H2713)
            txtTickTagStyleOutside.ForeColor = RGB(0, 255, 0)
            txtTickTagStyleOutside.Font.Bold = True
        End If
    Else
       txtTickTagStyleOutside = "N/A"
       txtTickTagStyleOutside.Font.Bold = True
       txtTickTagStyleOutside.ForeColor = RGB(255, 0, 0)
End If
    DoEvents
    
    
    
    'Marca el boton como activo
    btnCheckStart.Enabled = True
    'HABILITA BOTON PARAC COPIAR COLUMNA 4
    btnCopyColumn.Enabled = True
   
    
    
End Sub




Private Sub BtnFixGlobalStyles_Click()
'Resetea los estilos de todas las celdas en las columna 4 a Normal o Tag
Dim response As VbMsgBoxResult
Dim ListSep As String
Dim FindT1 As String
Dim FindT2 As String
Dim FindT3 As String

'Tiene en cuenta cual es el separador de listas si es coma o punto y coma
ListSep = Application.International(wdListSeparator)
FindT1 = "\<([0-9]{1" & ListSep & "})\>"
FindT2 = "\<\/([0-9]{1" & ListSep & "})\>"
FindT3 = "\<([0-9]{1" & ListSep & "})\/\>"
    
    response = MsgBox("This will clean all styles in column 4 and refresh it." & vbCrLf & "Do you want to proceed?", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbNo Then
       Exit Sub
    End If
        

 'Selecciona columna 4 desde la celda 2 a final
    Dim tbl As Table
    Dim lastRow As Long
    Dim col As Long
    Dim cellRange As Range

    ' Set the table number
    Set tbl = ActiveDocument.Tables(1)

    ' Find the last row in the table
    lastRow = tbl.Rows.count

    ' Set the column number (change 4 to the desired column number)
    col = 4

    ' Check if the table has at least 2 rows (header + data)
    If lastRow >= 2 Then
        ' Select the cell range in column 4 from row 2 to the end of the table
        Set cellRange = tbl.cell(2, col).Range
        cellRange.End = tbl.cell(lastRow, col).Range.End
        cellRange.Select
    Else
        MsgBox "Table has fewer than 2 rows.", vbExclamation
    End If

'CAMBIA TODO A DEFAULT PARAGRAPH FONT

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles( _
        "Default Paragraph Font")
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .text = "?"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

'CAMBIA TAGS COMO ESTILO TAG

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT1
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT2
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT3
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    btnFixAutoStyle.Enabled = False
    btnFixMissingStyleOnCell.Enabled = False
    Selection.Collapse Direction:=wdCollapseEnd
    
End Sub

Private Sub btnFixMissingStyleOnCell_Click()

    Dim oCell As cell
    Dim oRange As Range
    Dim restRange As Range
    Dim SearchText As String
    Dim ReplaceText As String
    Dim cellText As String
    Dim newText As String
    Dim i As Integer
    Dim CrAtEnd As Boolean
    Dim CrAtMid As Boolean
    
    Dim ListSep As String
    Dim FindT1 As String
    Dim FindT2 As String
    Dim FindT3 As String
    
    'Crea marcadores si lleva retorno de carro al final y/o en medio
    cellText = Selection.Cells(1).Range.text
    
    
     
    'Inicializa las variables
    ' Set a reference to the selected cell.
    Set oCell = Selection.Cells(1)

    ' Set a reference to the range inside the cell.
    Set oRange = oCell.Range

    
    'Tiene en cuenta cual es el separador de listas si es coma o punto y coma
    ListSep = Application.International(wdListSeparator)
    FindT1 = "\<([0-9]{1" & ListSep & "})\>"
    FindT2 = "\<\/([0-9]{1" & ListSep & "})\>"
    FindT3 = "\<([0-9]{1" & ListSep & "})\/\>"

'CAMBIA TODO A DEFAULT PARAGRAPH FONT

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles( _
        "Default Paragraph Font")
    With Selection.Find
        .text = "?"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

'CAMBIA TAGS COMO ESTILO TAG

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT1
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT2
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT3
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    btnFixAutoStyle.Enabled = False
    btnFixMissingStyleOnCell.Enabled = False
    Selection.Collapse Direction:=wdCollapseEnd
        
         
   

End Sub

Private Sub btnFixTagStyleOut_Click()

    Dim oCell As cell
    Dim oRange As Range
    Dim restRange As Range
    Dim SearchText As String
    Dim ReplaceText As String
    Dim cellText As String
    Dim newText As String
    Dim i As Integer
    Dim CrAtEnd As Boolean
    Dim CrAtMid As Boolean
    
    Dim ListSep As String
    Dim FindT1 As String
    Dim FindT2 As String
    Dim FindT3 As String
    
    'Crea marcadores si lleva retorno de carro al final y/o en medio
    cellText = Selection.Cells(1).Range.text
    
    
     
    'Inicializa las variables
    ' Set a reference to the selected cell.
    Set oCell = Selection.Cells(1)

    ' Set a reference to the range inside the cell.
    Set oRange = oCell.Range

    
    'Tiene en cuenta cual es el separador de listas si es coma o punto y coma
    ListSep = Application.International(wdListSeparator)
    FindT1 = "\<([0-9]{1" & ListSep & "})\>"
    FindT2 = "\<\/([0-9]{1" & ListSep & "})\>"
    FindT3 = "\<([0-9]{1" & ListSep & "})\/\>"
    
'QUITA HL
'Selection.Find.ClearFormatting
'    Selection.Find.Highlight = True
'    Selection.Find.Replacement.ClearFormatting
'    Selection.Find.Replacement.Highlight = False
'    With Selection.Find
'        .text = ""
'        .Replacement.text = ""
'        .Forward = True
'        .Wrap = wdFindAsk
'        .Format = True
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchKashida = False
'        .MatchDiacritics = False
'        .MatchAlefHamza = False
'        .MatchControl = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'    End With
'    Selection.Find.Execute Replace:=wdReplaceAll
'
    

'CAMBIA TODO A DEFAULT PARAGRAPH FONT

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles( _
        "Default Paragraph Font")
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .text = "?"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

'CAMBIA TAGS COMO ESTILO TAG

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT1
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT2
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT3
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    btnFixAutoStyle.Enabled = False
    btnFixMissingStyleOnCell.Enabled = False
    Selection.Collapse Direction:=wdCollapseEnd
        
         
   
End Sub

Private Sub btnShowTagStyle_Click()
'Quita del formulario lo que no interesa enseñar
txtExtraTag.Visible = False
txtxRepeatedTags.Visible = False
lblMissTags.Visible = False
lblRepTags.Visible = False
'Cambia cosas de los otro botones que no corresponden a esta tarea
BilingualCheckForm.Width = 450
btnShowMissExtraTags.Enabled = False
btnShowMissExtraTags.Caption = "Show"
'Cambia la label de la columna
lblTagsExtra.Caption = "Tag:"
'Cambia el tamaño del frame
FmeTags.Width = 61


    If btnShowTagStyle.Caption = "Show" Then
        BilingualCheckForm.Width = 414
        btnShowTagStyle.Caption = "Hide"
    Exit Sub
End If

If btnShowTagStyle.Caption = "Hide" Then
    BilingualCheckForm.Width = 346
    btnShowTagStyle.Caption = "Show"
    Exit Sub
End If
End Sub

Private Sub btnFordStyleTag_Click()
    'Mira si estabamos en otro check antes
    If btnShowMissExtraTags.Caption = "Hide" Then
        'Pone el formaulario corto
        BilingualCheckForm.Width = 346
        'Cambia el caption de los otros show
        btnShowMissExtraTags.Caption = "Show"
    End If
    'Deactivate Show for other tasks
    btnShowMissExtraTags.Enabled = False
    btnFixSingle.Enabled = False
    optMLB.Enabled = False
    OptSpace.Enabled = False
    'Activa boton Fix
    btnFixAutoStyle.Enabled = True
    btnFixMissingStyleOnCell.Enabled = True
   

    
    'Specific job tasks
    If UBound(WrongTagsList) > contaTagsCellStyle Then
        contaTagsCellStyle = contaTagsCellStyle + 1
    End If
    lblCurrStyleTagCell.Caption = contaTagsCellStyle + 1
    ActiveDocument.Tables(1).cell(WrongTagsList(contaTagsCellStyle, 1), 4).Range.Select
    'Centra la pantalla
    centrapantalla
    'BilingualCheckForm.Width = 472
    If UBound(WrongTagsList) <= contaTagsCellStyle Then
        btnFordStyleTag.Enabled = False
    End If
    
    btnShowTagStyle.Enabled = True
    'If NumberCellsIncorrectStyle = 1 Then
    '    btnBackStyleTag.Enabled = False
    'Else
        btnBackStyleTag.Enabled = True
    'End If
    
    txtMissTag.Value = WrongTagsList(contaTagsCellStyle, 2)
    txtExtraTag.Enabled = False
    txtxRepeatedTags.Enabled = False
    
    lblTagsExtra.Caption = "Tags:"
    
End Sub

Private Sub btnBackStyleTag_Click()
    'Mira si estabamos en otro check antes
    If btnShowMissExtraTags.Caption = "Hide" Then
        'Pone el formaulario corto
        BilingualCheckForm.Width = 346
        'Cambia el caption de los otros show
        btnShowMissExtraTags.Caption = "Show"
    End If
    'Deactivate Show for other tasks
    btnShowMissExtraTags.Enabled = False
    btnFixSingle.Enabled = False
    optMLB.Enabled = False
    OptSpace.Enabled = False
    'Activa boton Fix
    btnFixAutoStyle.Enabled = True
    btnFixMissingStyleOnCell.Enabled = False
   

    
    'Specific job tasks
    If contaTagsCellStyle >= 1 Then
        contaTagsCellStyle = contaTagsCellStyle - 1
    End If
    lblCurrStyleTagCell.Caption = contaTagsCellStyle + 1
    ActiveDocument.Tables(1).cell(WrongTagsList(contaTagsCellStyle, 1), 4).Range.Select
    'Centra la pantalla
    centrapantalla
    'BilingualCheckForm.Width = 472
    If contaTagsCellStyle < 1 Then
        btnBackStyleTag.Enabled = False
    End If
    
    btnShowTagStyle.Enabled = True
    btnFordStyleTag.Enabled = True
    
    txtMissTag.Value = WrongTagsList(contaTagsCellStyle, 2)
    txtExtraTag.Enabled = False
    txtxRepeatedTags.Enabled = False
    
End Sub


Sub CheckTagStyleInColumn4()
    'Dim NumberCellsIncorrectStyle As Integer 'Esta será global numero de celdas con tags sin estilo
    'Dim WrongTagsList() As Variant 'Esta sera global array con posicion de celda en columna 4 y listado de tags que no tienen estilo
    Dim TempArray() As Variant 'Array Temporal
    Dim TagsProbIncell As String
    Dim tbl As Table
    Dim row As Integer
    Dim tagPattern As String
    Dim regEx As Object
    Dim matches As Object
    Dim cellText As String
    Dim firstTagOnly As Boolean
    
    'Setting variables
    NumberCellsIncorrectStyle = 0
    TagsProbIncell = ""
    
    ReDim WrongTagsList(0 To 0, 1 To 2) 'Redimensiona a cero rows y 2 columns
    
    ' Define the regex pattern to match "<n/>" where n can be any number
    tagPattern = "</?(\d+)/?>"
    
    On Error Resume Next
    'Set tbl = Selection.Tables(1)  'Cambiado porque daba error.
    Set tbl = ActiveDocument.Tables(1)
    On Error GoTo 0
          
    
    ' Initialize the regular expression object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.Pattern = tagPattern

    
    ' Loop through all rows in column 4 and check the tag and style
    For row = 1 To tbl.Rows.count
        cellText = tbl.cell(row, 4).Range.text
        
        ' Check if the tag pattern is present in the cell text
        If regEx.Test(cellText) Then
            ' Find all matches of the tag pattern in the cell text
            Set matches = regEx.Execute(cellText)
            For Each match In matches
                ' Get the tag from the match
                Dim tag As String
                tag = match.Value
                
                ' Create a temporary range for the matched tag
                Dim tempRange As Range
                Set tempRange = tbl.cell(row, 4).Range.Duplicate
                
                ' Find the tag in the cell text and select it
                With tempRange.Find
                    .text = tag
                    .MatchCase = False
                    .MatchWholeWord = False
                    .Forward = True
                    .Wrap = wdFindStop
                    .Execute
                End With
                
                ' Check if the style is different from "Tag" (Replace "Tag" with the actual style name)
                If tempRange.Style <> "Tag" Then
                    ' Select the cell with the incorrect tag
                    'tbl.cell(row, 4).Select
                    
                    'Suma las tags incorrecta de la celda específica
                    TagsProbIncell = TagsProbIncell & tag
                End If
                

            Next match
            
            If TagsProbIncell <> "" Then
                    'MsgBox "Celda: " & row & " Contiene: " & TagsProbIncell
                    NumberCellsIncorrectStyle = NumberCellsIncorrectStyle + 1
                
                
                    'Metemos en el array temporal
                    If WrongTagsList(0, 1) = "" Then
                        WrongTagsList(0, 1) = row
                        WrongTagsList(0, 2) = TagsProbIncell
                    Else
                        ReDim TempArray(0 To NumberCellsIncorrectStyle - 1, 1 To 2)
                    For z = 0 To NumberCellsIncorrectStyle - 2
                        TempArray(z, 1) = WrongTagsList(z, 1)
                        TempArray(z, 2) = WrongTagsList(z, 2)
                    Next z
                    
                    'Metemos el ultimo valor, Celda y tags
                    TempArray(NumberCellsIncorrectStyle - 1, 1) = row
                    TempArray(NumberCellsIncorrectStyle - 1, 2) = TagsProbIncell
                    
                    WrongTagsList = TempArray
                    
                    Erase TempArray
                    
                    End If
                    
            
            End If
            
            TagsProbIncell = ""
            
        End If
        

    Next row
    
    
    txtNumIncTagsStyle.Caption = NumberCellsIncorrectStyle
    'MsgBox "Numero de Celdas con tags incorrectas: " & NumberCellsIncorrectStyle
    
End Sub

Private Sub btnShowMissExtraTags_Click()
txtExtraTag.Visible = True
txtxRepeatedTags.Visible = True
lblMissTags.Visible = True
lblRepTags.Visible = True
'Cambia la label de la columna
lblTagsExtra.Caption = "Missing"
'Cambia el tamaño del frame
FmeTags.Width = 168.7
'Desactiva otros botones
btnShowTagStyle.Enabled = False
btnShowTagStyle.Caption = "Show"


If btnShowMissExtraTags.Caption = "Show" Then
    BilingualCheckForm.Width = 522
    btnShowMissExtraTags.Caption = "Hide"
    Exit Sub
End If

If btnShowMissExtraTags.Caption = "Hide" Then
    BilingualCheckForm.Width = 346
    btnShowMissExtraTags.Caption = "Show"
    Exit Sub
End If
    
End Sub




Private Sub btnTagFord_Click()
    'Mira si estabamos en otro check antes
    If btnShowTagStyle.Caption = "Hide" Then
        'Pone el formaulario corto
        BilingualCheckForm.Width = 346
        'Cambia el caption de los otros show
        btnShowTagStyle.Caption = "Show"
    End If
    'Deactivate Show for other tasks
    btnShowTagStyle.Enabled = False
    btnFixSingle.Enabled = False
    optMLB.Enabled = False
    OptSpace.Enabled = False
    btnFixAutoStyle.Enabled = False
    btnFixMissingStyleOnCell.Enabled = False
    
    
    'Job Specific Tasks
    If UBound(resultsArray) > contaTagsCell Then
        contaTagsCell = contaTagsCell + 1
    End If
    lblCurrCellTag.Caption = contaTagsCell + 1
    ActiveDocument.Tables(1).cell(resultsArray(contaTagsCell, 1), 4).Range.Select
    'Centra la pantalla
    centrapantalla
    'BilingualCheckForm.Width = 472
    If UBound(resultsArray) <= contaTagsCell Then
        btnTagFord.Enabled = False
    End If
    
    btnShowMissExtraTags.Enabled = True
    btnTagBack.Enabled = True
    
    txtMissTag.Value = resultsArray(contaTagsCell, 2)
    txtExtraTag.Value = resultsArray(contaTagsCell, 3)
    txtxRepeatedTags.Value = resultsArray(contaTagsCell, 4)
End Sub
Private Sub btnTagBack_Click()
    'Mira si estabamos en otro check antes
    If btnShowTagStyle.Caption = "Hide" Then
        'Pone el formaulario corto
        BilingualCheckForm.Width = 346
        'Cambia el caption de los otros show
        btnShowTagStyle.Caption = "Show"
    End If
    'Deactivate Show for other tasks
    btnShowTagStyle.Enabled = False
    btnFixSingle.Enabled = False
    optMLB.Enabled = False
    OptSpace.Enabled = False
    btnFixAutoStyle.Enabled = False
    btnFixMissingStyleOnCell.Enabled = False
    
    
    
    'Job Specific Tasks
    If contaTagsCell >= 1 Then
        contaTagsCell = contaTagsCell - 1
    End If
    lblCurrCellTag.Caption = contaTagsCell + 1
    ActiveDocument.Tables(1).cell(resultsArray(contaTagsCell, 1), 4).Range.Select
    'Centra la pantalla
    centrapantalla
    'BilingualCheckForm.Width = 472
    If contaTagsCell < 1 Then
        btnTagBack.Enabled = False
    End If
    
    btnShowMissExtraTags.Enabled = True
    btnTagFord.Enabled = True
    
    txtMissTag.Value = resultsArray(contaTagsCell, 2)
    txtExtraTag.Value = resultsArray(contaTagsCell, 3)
    txtxRepeatedTags.Value = resultsArray(contaTagsCell, 4)
    
End Sub

Private Sub btnFixSingle_Click()
Dim oCell As cell
Dim oRange As Range
Dim ReplaceTextCR As String


'Texto para replacement
ReplaceTextCR = Chr(11)



' Set a reference to the selected cell.
Set oCell = Selection.Cells(1)
Set oRange = oCell.Range

'If there are CR at extremes run this
If CrAtBeg = True Or CrAtEnd = True Then

    oRange.Start = oRange.Start + (startPos - 1)
    oRange.End = oRange.Start + (endPos - startPos) + 1
    
End If

'if CR at mid run this and make replacement
If CrAtMid = True Then

    'If space or MLB
    If optMLB.Value = True Then
            ReplaceTextCR = Chr(11)
    Else
            ReplaceTextCR = " "
    End If

    'Make replacement
    oRange.Find.Execute findText:=Chr(13), ReplaceWith:=ReplaceTextCR, Replace:=wdReplaceAll

End If

'Copy fixed text to cell
oRange.Copy
DoEvents
oCell.Range.PasteSpecial DataType:=wdPasteHTML

'Make form assignements
btnFixSingle.Enabled = False
optMLB.Enabled = False
OptSpace.Enabled = False

End Sub



Private Sub btnTagStyleFord_Click()
'Mira si estabamos en otro check antes
    If btnShowTagStyle.Caption = "Hide" Then
        'Pone el formaulario corto
        BilingualCheckForm.Width = 346
        'Cambia el caption de los otros show
        btnShowTagStyle.Caption = "Show"
    End If
    'Deactivate Show for other tasks
    btnShowTagStyle.Enabled = False
    btnFixSingle.Enabled = False
    optMLB.Enabled = False
    OptSpace.Enabled = False
    btnFixAutoStyle.Enabled = False
    btnFixMissingStyleOnCell.Enabled = False
    
    
    'Job Specific Tasks
    
    'MsgBox UBound(ArrayCellsIncorrectTagStyle)
    
    If UBound(ArrayCellsIncorrectTagStyle) > contaWrongTagStyle Then
        contaWrongTagStyle = contaWrongTagStyle + 1
    End If
    lblCurrCellStyleOutside.Caption = contaWrongTagStyle + 1
    ActiveDocument.Tables(1).cell(ArrayCellsIncorrectTagStyle(contaWrongTagStyle), 4).Range.Select
    'Centra la pantalla
    centrapantalla
    'BilingualCheckForm.Width = 472
    If UBound(ArrayCellsIncorrectTagStyle) <= contaWrongTagStyle Then
        btnTagStyleFord.Enabled = False
    End If
    
    
    
    btnFixTagStyleOut.Enabled = True
    btnFixGlobalStyles.Enabled = True
    btnTagStyleBack.Enabled = True
    
    If contaWrongTagStyle = 0 Then
        btnTagStyleBack.Enabled = False
    End If
    
   
End Sub

Private Sub btnTagStyleBack_Click()
    'Mira si estabamos en otro check antes
    If btnShowTagStyle.Caption = "Hide" Then
        'Pone el formaulario corto
        BilingualCheckForm.Width = 346
        'Cambia el caption de los otros show
        btnShowTagStyle.Caption = "Show"
    End If
    'Deactivate Show for other tasks
    btnShowTagStyle.Enabled = False
    btnFixSingle.Enabled = False
    optMLB.Enabled = False
    OptSpace.Enabled = False
    btnFixAutoStyle.Enabled = False
    btnFixMissingStyleOnCell.Enabled = False
    
        
    'Job Specific Tasks
    If contaWrongTagStyle >= 1 Then
        contaWrongTagStyle = contaWrongTagStyle - 1
    End If
    lblCurrCellStyleOutside.Caption = contaWrongTagStyle + 1
    ActiveDocument.Tables(1).cell(ArrayCellsIncorrectTagStyle(contaWrongTagStyle), 4).Range.Select
    'Centra la pantalla
    centrapantalla
    'BilingualCheckForm.Width = 472
    
    'Maneja el navegador
    If contaWrongTagStyle < 1 Then
        btnTagStyleBack.Enabled = False
    End If
    
    If UBound(ArrayCellsIncorrectTagStyle) <= contaWrongTagStyle Then
        btnTagStyleFord.Enabled = False
    Else
        btnTagStyleFord.Enabled = True
    End If
    
    If contaWrongTagStyle = 0 Then
        btnTagStyleBack.Enabled = False
    End If
    
    btnFixTagStyleOut.Enabled = True
    
End Sub

Private Sub CarrBack_Click()
Dim cellText As String

'Mira si estabamos en otro check antes
 If btnShowTagStyle.Caption = "Hide" Or btnShowMissExtraTags.Caption = "Hide" Then
        'Pone el formaulario corto
        BilingualCheckForm.Width = 346
        'Cambia el caption de los otros show
        btnShowTagStyle.Caption = "Show"
        btnShowMissExtraTags.Caption = "Show"
 End If

'Deactivate Show for other tasks
btnShowTagStyle.Enabled = False
btnShowMissExtraTags.Enabled = False

optMLB.Enabled = False
OptSpace.Enabled = False

'Codigo de este sub
CeldaCurr = CeldaCurr - 1
lblCurrCell.Caption = CeldaCurr
    
       
    ActiveDocument.Tables(1).cell(Col3Col4(CeldaCurr, 2), Col3Col4(CeldaCurr, 1)).Select
    
    
    If CeldaCurr <= LBound(Col3Col4, 1) Then
        CarrBack.Enabled = False
        CarrFord.Enabled = True
        Exit Sub
    End If
    
    
     'Llama a sub que mira a ver si se debe activar los botones de CR y Space
    CheckCRorSPC
    
    
    If CrAtMid = True Then
        optMLB.Enabled = True
        OptSpace.Enabled = True
    Else
        optMLB.Enabled = False
        OptSpace.Enabled = False
    End If
  
centrapantalla
Application.ScreenRefresh
End Sub

Private Sub CarrFord_Click()
Dim cellIndex As Integer
Dim targetCell As cell
Dim cellText As String
Dim CarrRetOnStr As Long
  
'Mira si estabamos en otro check antes
If btnShowTagStyle.Caption = "Hide" Or btnShowMissExtraTags.Caption = "Hide" Then
        'Pone el formaulario corto
        BilingualCheckForm.Width = 346
        'Cambia el caption de los otros show
        btnShowTagStyle.Caption = "Show"
        btnShowMissExtraTags.Caption = "Show"
End If
'Deactivate Show for other tasks
btnShowTagStyle.Enabled = False
btnShowMissExtraTags.Enabled = False
btnFixAutoStyle.Enabled = False
btnFixMissingStyleOnCell.Enabled = False

optMLB.Enabled = False
OptSpace.Enabled = False
    
'Codigo de este sub
CeldaCurr = CeldaCurr + 1
lblCurrCell.Caption = CeldaCurr
   
    If CeldaCurr <= 1 Then
        CarrBack.Enabled = False
    Else
        CarrBack.Enabled = True
    End If
    
    If CeldaCurr = UBound(Col3Col4, 1) Then
        CarrFord.Enabled = False
        'Exit Sub
    End If
    
    
    ActiveDocument.Tables(1).cell(Col3Col4(CeldaCurr, 2), Col3Col4(CeldaCurr, 1)).Select
  centrapantalla
  Application.ScreenRefresh
    
  btnFixSingle.Enabled = True
  
    cellText = Selection.Cells(1).Range.text
    
    'Habilitamos las opciones de reemplazo
    'Numero de veces que aparece el retorno de carro
    For z = 1 To Len(cellText) - 2
        If Mid(cellText, z, 1) = Chr(13) Then
            CarrRetOnStr = CarrRetOnStr + 1
        End If
    Next z
    
    
    'Llama a sub que mira a ver si se debe activar
    
    CheckCRorSPC
    
    
    If CrAtMid = True Then
        optMLB.Enabled = True
        OptSpace.Enabled = True
    Else
        optMLB.Enabled = False
        OptSpace.Enabled = False
    End If
    
  
End Sub

Private Sub CheckCRorSPC()
    Dim cellText As String '
    
    CrAtBeg = False
    CrAtEnd = False
    CrAtMid = False
    
    cellText = Selection.Cells(1).Range.text
    startPos = 1

    'en endPos quita los dos ultimos caracteres que marcan la celda
    endPos = Len(cellText) - 2

    ' Determina donde empieza el texto de inicio quitando los CR y se guarda en startPos
    While startPos <= endPos And (Mid(cellText, startPos, 1) = Chr(10) Or Mid(cellText, startPos, 1) = Chr(13))
        startPos = startPos + 1
        CrAtBeg = True
    Wend

    ' Determina donde empieza el texto de final quitando los CR y se guarda en endPos
    While endPos >= startPos And (Mid(cellText, endPos, 1) = Chr(10) Or Mid(cellText, endPos, 1) = Chr(13))
        endPos = endPos - 1
        CrAtEnd = True
    Wend

    'Determina si hay CR entre texto pilla el texto sin los extremos y evalua
    MidText = Mid(cellText, startPos, endPos - startPos + 1)
    
    If InStr(1, MidText, Chr(13)) > 0 Then
        CrAtMid = True
    End If

End Sub
Function centrapantalla()
'Centra la pantalla en la selección
 Dim pLeft As Long
  Dim pTop As Long, lTop As Long, wTop As Long
  Dim pWidth As Long
  Dim pHeight As Long, wHeight As Long
  Dim Direction As Integer

  wHeight = PixelsToPoints(ActiveWindow.Height, True)
  ActiveWindow.GetPoint pLeft, wTop, pWidth, pHeight, ActiveWindow
  ActiveWindow.ScrollIntoView Selection.Range, True 'Este lo meto porque falla cuando hay un page break si se quita funciona igual
  ActiveWindow.GetPoint pLeft, pTop, pWidth, pHeight, Selection.Range

  Direction = Sgn((pTop + pHeight / 2) - (wTop + wHeight / 2))
  Do While Sgn((pTop + pHeight / 2) - (wTop + wHeight / 2)) = Direction And (lTop <> pTop)
    ActiveWindow.SmallScroll Direction
    lTop = pTop
    ActiveWindow.GetPoint pLeft, pTop, pWidth, pHeight, Selection.Range
  Loop

End Function

Function GetCellIndex(cell As cell) As Integer
    Dim i As Integer
    For i = LBound(celdasConRetorno) To UBound(celdasConRetorno)
        If celdasConRetorno(i).rowIndex = cell.rowIndex And celdasConRetorno(i).columnIndex = cell.columnIndex Then
            GetCellIndex = i
            Exit Function
        End If
    Next i
    GetCellIndex = -1 ' Cell not found in the array
End Function

Private Sub CheckTagsInWordTable()

    Dim doc As Document
    Dim tbl As Table
    Dim col3Text As String
    Dim col4Text As String
    Dim tagsCol3() As String
    Dim tagsCol4() As String
    Dim i As Long
    Dim RowCurrent As Long
    
    
    'Inicializa
    ReDim resultsArray(0 To 0, 1 To 4) ' Inicializar con una única fila y 4 columnas
    
    ' Set the active Word document.
    Set doc = ActiveDocument
    
    ' Assuming the table is the first table in the document. Adjust the index if needed.
    Set tbl = doc.Tables(1)
    
    ' Loop through each row in the table.
    For Each row In tbl.Rows
        ' Check if the row has enough cells to access columns 3 and 4 (avoid index out of range error).
        If row.Cells.count >= 4 Then
            ' Access the cells in column 3 and column 4 for each row.
            col3Text = row.Cells(3).Range.text
            col4Text = row.Cells(4).Range.text
            
            'Celda en la que nos encontramos
            RowCurrent = row.Index
            
            ' Extract tags from column 3 and column 4.
            tagsCol3 = GetTagsFromString(col3Text)
            tagsCol4 = GetTagsFromString(col4Text)
            
            
            'Busca la diferencia solo si hay valores
            If tagsCol3(0) > "" Or tagsCol4(0) > "" Then
                'Llama a funcion busca valores únicos
                FindUniqueValues RowCurrent, tagsCol3, tagsCol4
            End If
            
            
        End If
    Next row
    

End Sub

Function GetTagsFromString(inputText As String) As String()
    Dim regEx As Object
    Dim matches As Object
    Dim tagPattern As String
    Dim tags() As String
    Dim i As Long
    
    tagPattern = "</?\d+/?>"
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = True
        .IgnoreCase = True
        .Pattern = tagPattern
        
        If .Test(inputText) Then
            Set matches = .Execute(inputText)
            ReDim tags(matches.count - 1)
            
            For i = 0 To matches.count - 1
                tags(i) = matches.Item(i).Value
            Next i
        Else
            ' If no matches are found, return an empty array.
            ReDim tags(0)
            tags(0) = ""
        End If
    End With
    
    GetTagsFromString = tags
End Function

Function IsTagInArray(tag As String, arr() As String) As Boolean
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) = tag Then
            IsTagInArray = True
            Exit Function
        End If
    Next i
    
    IsTagInArray = False
End Function


Private Sub FindUniqueValues(RowCurrent As Long, tagsCol3() As String, tagsCol4() As String)
    Dim j As Long
    Dim foundMatch As Boolean ' Declare foundMatch only once here
    Dim numRows As Long 'Almacena el numero de filas del temporal
    
    Dim MissT As String
    Dim ExtraT As String
    Dim RepeatT As String
    
    Dim numRowsInTemp As Long
        
    
    ' Encontrar valores únicos de tagsCol3 que no están en tagsCol4
    For i = LBound(tagsCol3) To UBound(tagsCol3)
        foundMatch = False ' Initialize foundMatch inside the loop

        For j = LBound(tagsCol4) To UBound(tagsCol4)
            If tagsCol3(i) = tagsCol4(j) Then
                foundMatch = True
                Exit For
            End If
        Next j

        If Not foundMatch Then
            MissT = MissT & tagsCol3(i)
        End If
    Next i

    ' Encontrar valores únicos de tagsCol4 que no están en tagsCol3
    For i = LBound(tagsCol4) To UBound(tagsCol4)
        foundMatch = False ' Initialize foundMatch inside the loop

        For j = LBound(tagsCol3) To UBound(tagsCol3)
            If tagsCol4(i) = tagsCol3(j) Then
                foundMatch = True
                Exit For
            End If
        Next j

        If Not foundMatch Then
             ExtraT = ExtraT & tagsCol4(i)
        End If
    Next i
    
    'Encontar valores repetidos en column 4
    For i = LBound(tagsCol4) To UBound(tagsCol4)
    
        ' Check against the remaining elements in the array
        For j = LBound(tagsCol4) To UBound(tagsCol4)
            If tagsCol4(i) = tagsCol4(j) And i <> j Then
                ' If a match is found, mark it as a repeated value
                RepeatT = RepeatT & tagsCol4(i)
                Exit For ' No need to check further for this element
            End If
        Next j
    Next i
    
    
    'mete valores en el array final solo si hay algo que meter
    If ExtraT <> "" Or MissT <> "" Or RepeatT <> "" Then
        
        If resultsArray(0, 1) = "" Then 'Esto es para el valor inicial
            resultsArray(0, 1) = RowCurrent
            resultsArray(0, 2) = MissT
            resultsArray(0, 3) = ExtraT
            resultsArray(0, 4) = RepeatT
        Else
           numRowsInTemp = UBound(resultsArray, 1) + 1
           Dim TempArray() As Variant
           ReDim TempArray(0 To numRowsInTemp, 1 To 4)
            
            ' Copy the values from the original array to the temporary array
            Dim z As Long
            For z = 0 To numRowsInTemp - 1
                TempArray(z, 1) = resultsArray(z, 1)
                TempArray(z, 2) = resultsArray(z, 2)
                TempArray(z, 3) = resultsArray(z, 3)
                TempArray(z, 4) = resultsArray(z, 4)
            Next z
            
            ' Fill the new row with values
            TempArray(numRowsInTemp, 1) = RowCurrent
            TempArray(numRowsInTemp, 2) = MissT
            TempArray(numRowsInTemp, 3) = ExtraT
            TempArray(numRowsInTemp, 4) = RepeatT
            
            
           ' Reassign the original array to the temporary array
            resultsArray = TempArray
    
            Erase TempArray
           
        End If
    End If
End Sub
Private Sub btnFixAutoStyle_Click()
'Resetea los estilos de todas las celdas en las columna 4 a Normal o Tag
Dim response As VbMsgBoxResult
Dim ListSep As String
Dim FindT1 As String
Dim FindT2 As String
Dim FindT3 As String

'Tiene en cuenta cual es el separador de listas si es coma o punto y coma
ListSep = Application.International(wdListSeparator)
FindT1 = "\<([0-9]{1" & ListSep & "})\>"
FindT2 = "\<\/([0-9]{1" & ListSep & "})\>"
FindT3 = "\<([0-9]{1" & ListSep & "})\/\>"
    
    response = MsgBox("This will clean all styles in column 4 and refresh it." & vbCrLf & "Do you want to proceed?", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbNo Then
       Exit Sub
    End If
        

 'Selecciona columna 4 desde la celda 2 a final
    Dim tbl As Table
    Dim lastRow As Long
    Dim col As Long
    Dim cellRange As Range

    ' Set the table number
    Set tbl = ActiveDocument.Tables(1)

    ' Find the last row in the table
    lastRow = tbl.Rows.count

    ' Set the column number (change 4 to the desired column number)
    col = 4

    ' Check if the table has at least 2 rows (header + data)
    If lastRow >= 2 Then
        ' Select the cell range in column 4 from row 2 to the end of the table
        Set cellRange = tbl.cell(2, col).Range
        cellRange.End = tbl.cell(lastRow, col).Range.End
        cellRange.Select
    Else
        MsgBox "Table has fewer than 2 rows.", vbExclamation
    End If

'CAMBIA TODO A DEFAULT PARAGRAPH FONT

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles( _
        "Default Paragraph Font")
    With Selection.Find
        .text = "?"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

'CAMBIA TAGS COMO ESTILO TAG

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT1
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT2
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Tag")
    With Selection.Find
        .text = FindT3
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    btnFixAutoStyle.Enabled = False
    btnFixMissingStyleOnCell.Enabled = False
    Selection.Collapse Direction:=wdCollapseEnd
    
    
End Sub

Sub StrikeUnderline()
    Dim doc As Document
    Dim tbl As Table
    Dim col As Column
    Dim cell As cell
    Dim count As Integer
    Dim i As Integer
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Set the table (assuming it's the first table in the document)
    Set tbl = doc.Tables(1)
    
    ' Set the column (assuming it's the fourth column)
    Set col = tbl.Columns(4)
    
    ' Initialize array and count
    count = 0
    
    ' Loop through each cell in the column
    For Each cell In col.Cells
        If CellHasUnderlineOrStrikethrough(cell) Then
            count = count + 1
            ReDim Preserve ArrayCellsStrikeUnder(1 To count)
            ArrayCellsStrikeUnder(count) = cell.rowIndex
        End If
    Next cell
    
    NumberCellsStrikeUnder = count
    txtNumStriUnder = count
    

End Sub
Function CellHasUnderlineOrStrikethrough(cell As cell) As Boolean
'Funcion usada por Sub StrikeUnderline para saber si llevan strike o underline
    Dim rng As Range
    'For Each rng In cell.Range.Words
    For Each rng In cell.Range.Characters
        If rng.Font.Underline <> wdUnderlineNone Or rng.Font.StrikeThrough = True Then
            CellHasUnderlineOrStrikethrough = True
            Exit Function
        End If
    Next rng
    CellHasUnderlineOrStrikethrough = False
End Function
Private Sub btnStrikeUnderFord_Click()
'Mira si estabamos en otro check antes
    If btnShowTagStyle.Caption = "Hide" Or btnShowMissExtraTags.Caption = "Hide" Then
        'Pone el formaulario corto
        BilingualCheckForm.Width = 346
        'Cambia el caption de los otros show
        btnShowTagStyle.Caption = "Show"
        btnShowMissExtraTags.Caption = "Show"
    End If
    'Deactivate Show for other tasks
    btnShowTagStyle.Enabled = False
    btnFixSingle.Enabled = False
    optMLB.Enabled = False
    OptSpace.Enabled = False
    btnFixAutoStyle.Enabled = False
    btnFixMissingStyleOnCell.Enabled = False
    btnShowMissExtraTags.Enabled = False
    
    
    'Job Specific Tasks
    If UBound(ArrayCellsStrikeUnder) > contaStrikeUnd Then
        contaStrikeUnd = contaStrikeUnd + 1
    End If
    lblCurrentStrikeUnder.Caption = contaStrikeUnd
    ActiveDocument.Tables(1).cell(ArrayCellsStrikeUnder(contaStrikeUnd), 4).Range.Select
    'Centra la pantalla
    centrapantalla
    'Desactiba Fordward si llegamos al tope
    If UBound(ArrayCellsStrikeUnder) <= contaStrikeUnd Then
        btnStrikeUnderFord.Enabled = False
    End If
    
    
    btnStrikeUnderBack.Enabled = True
    

End Sub

Private Sub btnStrikeUnderBack_Click()
 'Mira si estabamos en otro check antes
    If btnShowTagStyle.Caption = "Hide" Then
        'Pone el formaulario corto
        BilingualCheckForm.Width = 346
        'Cambia el caption de los otros show
        btnShowTagStyle.Caption = "Show"
        btnShowMissExtraTags.Caption = "Show"
    End If
    'Deactivate Show for other tasks
    btnShowTagStyle.Enabled = False
    btnFixSingle.Enabled = False
    optMLB.Enabled = False
    OptSpace.Enabled = False
    btnFixAutoStyle.Enabled = False
    btnFixMissingStyleOnCell.Enabled = False
    btnShowMissExtraTags.Enabled = False
    
    
    
    'Job Specific Tasks
    If contaStrikeUnd > 1 Then
        contaStrikeUnd = contaStrikeUnd - 1
    End If
    lblCurrentStrikeUnder.Caption = contaStrikeUnd
    ActiveDocument.Tables(1).cell(ArrayCellsStrikeUnder(contaStrikeUnd), 4).Range.Select
    'Centra la pantalla
    centrapantalla
    'BilingualCheckForm.Width = 472
    If contaStrikeUnd < 2 Then
        btnStrikeUnderBack.Enabled = False
    End If
    
    
    btnStrikeUnderFord.Enabled = True
    
   
End Sub
Sub CheckInsertionsDeletions()
Dim lInsertsWords As Long
Dim lInsertsChar As Long
Dim lDeletesWords As Long
Dim lDeletesChar As Long
Dim oRevision As Revision
    
lInsertsWords = 0
lInsertsChar = 0
lDeletesWords = 0
lDeletesChar = 0
    
    
' Felipe. There are problems with the Macro when there are fields in the document, converting them to text firts.
'De momento lo quitamos
'FieldsToText
    
  
For Each oRevision In ActiveDocument.Revisions
        
        Select Case oRevision.Type
            Case wdRevisionInsert
                lInsertsChar = lInsertsChar + Len(oRevision.Range.text)
                lInsertsWords = lInsertsWords + oRevision.Range.Words.count
            Case wdRevisionDelete
                lDeletesChar = lDeletesChar + Len(oRevision.Range.text)
                lDeletesWords = lDeletesWords + oRevision.Range.Words.count
        End Select
Next oRevision
    
'Copia a formulario
txtInsertions.Caption = lInsertsWords & "--" & lInsertsChar
txtDeletions.Caption = lDeletesWords & "--" & lDeletesChar

   
End Sub

Private Sub btnCopyColumn_Click()
    Dim areColumnsIdentical As Boolean
    Dim result As Integer
    Dim dlg As FileDialog
    Dim cell1 As cell
    Dim cell2 As cell
    Dim TGDoc As String
    
    Dim sourceDoc As Document
    Dim targetDoc As Document
    Dim sourceTable As Table
    Dim targetTable As Table
    Dim iRow As Long
    Dim sourceCellRange As Range
    Dim targetCellRange As Range

'Comprueba que solo hay un doc abierto si no sale
    If Documents.count > 1 Then
        MsgBox "Only the Bilingual document should be open in word", vbCritical
        Exit Sub
    End If

    'ABRE SEGUNDO DOCUMENTO
    ' Initialize the flag for column comparison
    areColumnsIdentical = True
  
    
    ' Create a FileDialog object as a File Picker dialog box
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Set title and filters for the dialog
    dlg.Title = "Select Bilingual document to be pasted Column 4"
    dlg.Filters.Clear
    dlg.Filters.Add "Word Documents", "*.docx"
    dlg.Filters.Add "All Files", "*.*"
    
    ' Show the dialog and get the user's selection
    result = dlg.Show
    
    ' If the user selects a file, open it
    If result <> 0 Then
        filePath = dlg.SelectedItems(1)
        Documents.Open filePath
    End If
    
    ' Release the FileDialog object
    Set dlg = Nothing

'Modificado
   Set targetDoc = Documents.Item(2)
   
'Original
   Set sourceDoc = Documents.Item(1)
   
'Tablas
   'Set targetTable = targetDoc.Tables(1)
   Set targetTable = sourceDoc.Tables(1)
   'Set sourceTable = sourceDoc.Tables(1)
   Set sourceTable = targetDoc.Tables(1)
  
  
  'Track changes off in both docs
   sourceDoc.TrackRevisions = False
   'doc2.TrackRevisions = False



'If they have same number of rows compare
  If sourceTable.Rows.count = targetTable.Rows.count Then
    
    ' Loop through the rows in the first column of both tables
    For i = 1 To sourceTable.Rows.count
        Set cell1 = sourceTable.cell(i, 1)
        Set cell2 = targetTable.cell(i, 1)
        
        ' Compare the text in the cells column 1
        If cell1.Range.text <> cell2.Range.text Then
            areColumnsIdentical = False
            MsgBox "Files at table row: " & i & " are not the same in ID colum"
            Exit For
        End If
    Next i
    
   Else
        MsgBox "Documents no have same number of rows!" & vbCrLf & vbCrLf & "Translated document rows number: " & _
        sourceTable.Rows.count & vbCrLf & "Original document rows number: " & targetTable.Rows.count & vbCrLf & vbCrLf & "Exit!", vbCritical
        Exit Sub
  End If
  

    ' Loop through each row of the source table
    For iRow = 2 To sourceTable.Rows.count
        If iRow <= targetTable.Rows.count Then
            ' Define the cell range, excluding the end-of-cell marker
            Set sourceCellRange = sourceTable.cell(iRow, 4).Range
            sourceCellRange.End = sourceCellRange.End - 1
            
            Set targetCellRange = targetTable.cell(iRow, 4).Range
            targetCellRange.End = targetCellRange.End - 1

            ' Copy formatted text from the 4th column of source table to target table
            targetCellRange.FormattedText = sourceCellRange.FormattedText
        End If
    Next iRow

   
    
    
    
    sourceDoc.TrackRevisions = True
    lblTrackChanges.Caption = "Track Changes ON"
    lblTrackChanges.ForeColor = RGB(0, 0, 0)
    
     'Save the fixed one in the same path but with name modified
    sourceDoc.SaveAs2 FileName:=targetDoc.Path & "\" & "Fixed_" & targetDoc.Name
    'sourceDoc.Saved = True
    
    
    'Save the one we fixed
    TGDoc = targetDoc.Path
    targetDoc.Save
    targetDoc.Saved = True
    targetDoc.Close
    
    
    MsgBox "Now This document is a copy of original one keeping the changes in column 4" & vbCr & vbCr & "Saved at path: " & TGDoc
    
    SalvadoDoc = True
    
End Sub






Private Sub UserForm_Initialize()

    
   'Comprueba si Track changes está activo
    If ActiveDocument.TrackRevisions = True Then
        TrackChangesActive = True
        lblTrackChanges.Caption = "Track Changes ON"
        lblTrackChanges.ForeColor = RGB(255, 0, 0)
    Else
        TrackChangesActive = False
        lblTrackChanges.Caption = "Track Changes OFF"
        lblTrackChanges.ForeColor = RGB(0, 128, 0)
    End If
    
    
    
   
End Sub

Private Sub UserForm_Terminate()

'Leave track changes as initial
If SalvadoDoc = False Then
    If TrackChangesActive = True Then
        ActiveDocument.TrackRevisions = True
    Else
        ActiveDocument.TrackRevisions = False
    End If
End If


'Track Changes on
'lblTrackChanges.ForeColor = RGB(255, 0, 0)
'   ActiveDocument.TrackRevisions = True
'   With ActiveDocument
'        .TrackRevisions = True
'        .ShowRevisions = True
'    End With
End Sub

Sub MarkTagStyleOutAsyellow()
    Dim oTable As Table
    Dim oCell As cell
    Dim oRange As Range
    Dim found As Boolean

    ' Set the first table
    Set oTable = ActiveDocument.Tables(1)
    
    'Clears HL in column 4
    oTable.Columns(4).Select
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    

    ' Iterate through each cell in the fourth column
    For Each oCell In oTable.Columns(4).Cells
        Set oRange = oCell.Range
        oRange.End = oRange.End - 1 ' Adjust to exclude end-of-cell marks

        With oRange.Find
            .ClearFormatting
            .Style = ActiveDocument.Styles("tag")
            .text = ""
            .Format = True
            .Forward = True
            .Wrap = wdFindStop
        End With

        ' Find the first instance
        found = oRange.Find.Execute

        While found
            'If Not (oRange.text Like "*[<>/" & Chr(48) & "-" & Chr(57) & "]*") Then
            If oRange.text Like "*[!<>/" & Chr(48) & "-" & Chr(57) & "]*" Then
                oRange.HighlightColorIndex = wdYellow ' Highlight the found text in yellow
            End If

            ' Collapse the range and find the next instance
            oRange.Collapse wdCollapseEnd
            found = oRange.Find.Execute
        Wend

        found = False ' Resetting the found variable for the next cell
    Next oCell
    
   
End Sub

Sub TagStyleOutAsyellowDetect()
    Dim tabla As Table
    Dim celda As cell
    'Dim ArrayCellsIncorrectTagStyle() As Integer
    Dim i As Integer
    Dim count As Integer
    
    ' Definir la tabla (ajusta el nombre de la tabla según sea necesario)
    Set tabla = ActiveDocument.Tables(1)
    
    
    'Inicializa contador por si hay segunda vuelta
    NumberCellIncorrectTagStyle = 0
    
    
    ' Iterar a través de cada celda de la columna 4
    For Each celda In tabla.Columns(4).Cells
        ' Realizar la búsqueda y reemplazo en la celda actual
        With celda.Range.Find
        
            .Highlight = True
          
            ' Verificar si se encontró resaltado amarillo
            Do While .Execute
                ' Obtener el número de fila
                Dim numeroFila As Integer
                numeroFila = celda.rowIndex
                
                ' Verificar si el número de fila ya está en el array
                Dim yaEnArray As Boolean
                yaEnArray = False
                
                ' Inicializar el array si aún no se ha hecho
                If count = 0 Then
                    ReDim ArrayCellsIncorrectTagStyle(0)
                    count = 1
                End If
                
                ' Buscar si el número de fila ya está en el array
                For i = LBound(ArrayCellsIncorrectTagStyle) To UBound(ArrayCellsIncorrectTagStyle)
                    If ArrayCellsIncorrectTagStyle(i) = numeroFila Then
                        yaEnArray = True
                        Exit For
                    End If
                Next i
                
                ' Si no está en el array, agregarlo
                If Not yaEnArray Then
                    ReDim Preserve ArrayCellsIncorrectTagStyle(count - 1)
                    ArrayCellsIncorrectTagStyle(count - 1) = numeroFila
                    count = count + 1
                    'Añade 1 al contador general
                    NumberCellIncorrectTagStyle = NumberCellIncorrectTagStyle + 1
                    
                End If
            Loop
        End With
    Next celda
    
    
    ' Imprimir el contenido del array en la ventana inmediata (Ctrl + G para ver la ventana inmediata)
    'For i = LBound(ArrayCellsIncorrectTagStyle) To UBound(ArrayCellsIncorrectTagStyle)
    '    Debug.Print ArrayCellsIncorrectTagStyle(i)
    'Next i
End Sub

Sub ClearClipboard()
    OpenClipboard 0
    EmptyClipboard
    CloseClipboard
End Sub
