Attribute VB_Name = "Módulo1"
Sub INICIO()
    'RESETAR PLANILHA
        Sheets("Resultado").Select
        Range("A2:U999999, W4:AD999999").ClearContents
        Sheets("Entradas").Select
        Cells.ClearContents
        Cells.ClearContents
        
'IMPORTACAO JETTAX
        Dim caminhoUso As String
        Dim wbUso As Workbook
        Dim wbSaida As Workbook
        
        MsgBox "Selecione a planilha de Entrada do G5", vbExclamation
        
    'Abertura de planilha
        caminhoUso = Application.GetOpenFilename(FileFilter:="Planilha do G5 (*.xlsx; *.xls), *.xlsx; *.xls", Title:="Selecione a Planilha")
        If caminhoUso <> "Falso" Then
            Set wbUso = Workbooks.Open(caminhoUso)
            wbUso.Activate
    

    
    'COPIA DA PLANILHA DO JETTAX
        wbUso.Activate
        Range("A4:AM4").Select
        Range(Selection, Selection.End(xlDown)).Copy
        Windows("ANALISE.xlsm").Activate
        Sheets("Entradas").Select
        Range("A1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        wbUso.Close SaveChanges:=False
        Range("T:T").Replace What:="-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    
    VALOR_CONTABIL


    MsgBox "Analise Completa", vbExclamation
End If
Sheets("MENU").Select
End Sub

Sub GIA()
Attribute GIA.VB_ProcData.VB_Invoke_Func = " \n14"
'
'BASE + ISENTO + OUTROS + ST + IPI
        Dim rng As Range
        Dim resultadoFiltro As Long
        Sheets("Entradas").Select
        ActiveSheet.ShowAllData
        
    'FORMULA
        Range("I2").FormulaR1C1 = "=ROUNDDOWN(RC[3]-RC[4]-RC[6]-RC[7]-RC[11],2)"
        Range("I2").Copy
        Range("J2").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(0, -1).Select
        Range(Selection, Selection.End(xlUp)).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False

'FILTRO VALOR CONTABIL
        Columns("I:I").Select
        Application.CutCopyMode = False
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        Range("$A$1:$AM$999999").AutoFilter Field:=9, Criteria1:="<>0", Operator:=xlOr, Criteria2:="<>0"
            
    'Defina o intervalo onde você aplicou o filtro
        Set rng = ThisWorkbook.Sheets("Entradas").Range("H2:H999999")
    
    'Verifique se há células visíveis após aplicar o filtro
        resultadoFiltro = Application.WorksheetFunction.Subtotal(103, rng)
    
    'Verifique se há informações no resultado do filtro
        If resultadoFiltro = 0 Then
            Else
                'Se houver informações
                    MsgBox "                    ATENÇÃO" & vbCrLf & vbCrLf & _
                    "Base de calculo + Isento + Outros + ST + IPI" & vbCrLf & _
                    "Não da o Valor Contabil", vbCritical
                    
                'Copiar e colar na aba Resultado
                    Range("H1:I1").Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy
                    Sheets("Resultado").Select
                    Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    Range("B2").FormulaR1C1 = "Valor Divergente"
                    
        End If
Sheets("Resultado").Select
JettaxVSG5
End Sub
Sub VALOR_CONTABIL()
Attribute VALOR_CONTABIL.VB_ProcData.VB_Invoke_Func = " \n14"

'VALOR CONTABIL ZERADO
    'FILTRO VALOR CONTABIL
        Sheets("Entradas").Select
        Range("$A$1:$AM$999999").AutoFilter Field:=12, Criteria1:="0,00", _
            Operator:=xlOr, Criteria2:="0"
    'Defina o intervalo onde você aplicou o filtro
        Set rng = ThisWorkbook.Sheets("Entradas").Range("H2:H999999")
    
    'Verifique se há células visíveis após aplicar o filtro
        resultadoFiltro = Application.WorksheetFunction.Subtotal(103, rng)
    
    'Verifique se há informações no resultado do filtro
        If resultadoFiltro = 0 Then
            BASE_DE_ICMS
            Else
                'Se houver informações
                    MsgBox "             ATENÇÃO" & vbCrLf & vbCrLf & _
                    "Notas Sem Valor contabil", vbCritical
                    
                'Copiar e colar na aba Resultado
                    Range("H1").Select
                    Range(Selection, Selection.End(xlDown)).Copy
                    Sheets("Resultado").Select
                    Range("D2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
            BASE_DE_ICMS
        End If
End Sub
Sub BASE_DE_ICMS()

'BASE DE ICMS SEM ICMS
        Sheets("Entradas").Select
        ActiveSheet.ShowAllData
        
    'FILTRO VALOR CONTABIL
        Range("$A$1:$AM$999999").AutoFilter Field:=14, Criteria1:="0,00", _
            Operator:=xlOr, Criteria2:="0"
        Range("$A$1:$AM$999999").AutoFilter Field:=13, Criteria1:=">0,00", _
            Operator:=xlOr, Criteria2:=">0"
            
    'Defina o intervalo onde você aplicou o filtro
        Set rng = ThisWorkbook.Sheets("Entradas").Range("H2:H999999")
    
    'Verifique se há células visíveis após aplicar o filtro
        resultadoFiltro = Application.WorksheetFunction.Subtotal(103, rng)
    
    'Verifique se há informações no resultado do filtro
        If resultadoFiltro = 0 Then
            ICMS_SEM_BASE
            Else
                'Se houver informações
                    MsgBox "             ATENÇÃO" & vbCrLf & vbCrLf & _
                    "BASE DE CALCULO SEM ICMS", vbCritical
                    
                'Copiar e colar na aba Resultado
                    Range("K1:P1").Select
                    Range(Selection, Selection.End(xlDown)).Copy
                    Sheets("Resultado").Select
                    Range("H2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
                    Sheets("Entradas").Select
                    Range("H1").Select
                    Range(Selection, Selection.End(xlDown)).Copy
                    Sheets("Resultado").Select
                    Range("G2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
                    ICMS_SEM_BASE
        End If
        
End Sub
Sub ICMS_SEM_BASE()
'ICMS SEM BASE
        Sheets("Entradas").Select
        ActiveSheet.ShowAllData
        
    'FILTRO VALOR CONTABIL
        Range("$A$1:$AM$999999").AutoFilter Field:=14, Criteria1:=">0,00", _
            Operator:=xlOr, Criteria2:=">0"
        Range("$A$1:$AM$999999").AutoFilter Field:=13, Criteria1:="0,00", _
            Operator:=xlOr, Criteria2:="0"
            
    'Defina o intervalo onde você aplicou o filtro
        Set rng = ThisWorkbook.Sheets("Entradas").Range("H2:H999999")
    
    'Verifique se há células visíveis após aplicar o filtro
        resultadoFiltro = Application.WorksheetFunction.Subtotal(103, rng)
    
    'Verifique se há informações no resultado do filtro
        If resultadoFiltro = 0 Then
            GIA
            Else
                'Se houver informações
                    MsgBox "             ATENÇÃO" & vbCrLf & vbCrLf & _
                    "ICMS SEM BASE DE CALCULO", vbCritical
                    
                'Copiar e colar na aba Resultado
                    Range("K1:P1").Select
                    Range(Selection, Selection.End(xlDown)).Copy
                    Sheets("Resultado").Select
                    Range("P2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
                    Sheets("Entradas").Select
                    Range("H1").Select
                    Range(Selection, Selection.End(xlDown)).Copy
                    Sheets("Resultado").Select
                    Range("O2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
                    GIA
        End If
End Sub

Sub JettaxVSG5()



'IMPORTACAO JETTAX
        Dim caminhoUso As String
        Dim wbJtx As Workbook
        
        MsgBox "Selecione a planilha do JETTAX", vbExclamation
        
    'Abertura de planilha
        caminhoUso = Application.GetOpenFilename(FileFilter:="Planilha do G5 (*.xlsx; *.xls), *.xlsx; *.xls", Title:="Selecione a Planilha")
        If caminhoUso <> "Falso" Then
            Set wbJtx = Workbooks.Open(caminhoUso)
        wbJtx.Activate
        Select Case ActiveSheet.Name
        Case "Relatório Detalhado por Nota", "Relatório Detalhado por Produto"
        
    
    'LIMPAR PLANILHA
        Windows("ANALISE.xlsm").Activate
        Sheets("JTXvsG5").Select
        Range("A3:N999999").ClearContents
        Range("A:E").ClearContents
        
    'COPIA DA PLANILHA JTX
        Set wbJtx = Workbooks.Open(caminhoUso)
            wbJtx.Activate
        Sheets("Relatório Detalhado por Nota").Select
        Range("O2").NumberFormat = "General"
        Range("O2").FormulaR1C1 = _
            "=VLOOKUP(RC[-13],'Relatório Detalhado por Produto'!C[-14]:C,14,0)"
        Range("O2").Select
        Selection.Copy
        Range(Selection, Selection.End(xlDown)).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        Range("H:H, N:N, O:O, P:P, Q:Q").Select
        Range("Q1").Activate
        Selection.Copy
        Windows("ANALISE.xlsm").Activate
        Sheets("JTXvsG5").Select
        Range("J1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.DisplayAlerts = False
        wbJtx.Close SaveChanges:=False
        Application.DisplayAlerts = True
        
    'TABELA DINAMICA
        Sheets("Entradas").Select
        Columns("H:N").Select
        Application.CutCopyMode = False
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Entradas!R1C8:R1048576C20", Version:=8).CreatePivotTable TableDestination _
            :="JTXvsG5!R1C1", TableName:="Tabela dinâmica3", DefaultVersion:=8
        Sheets("JTXvsG5").Select
        Cells(1, 1).Select
        With ActiveSheet.PivotTables("Tabela dinâmica3")
            .ColumnGrand = True
            .HasAutoFormat = True
            .DisplayErrorString = False
            .DisplayNullString = True
            .EnableDrilldown = True
            .ErrorString = ""
            .MergeLabels = False
            .NullString = ""
            .PageFieldOrder = 2
            .PageFieldWrapCount = 0
            .PreserveFormatting = True
            .RowGrand = True
            .SaveData = True
            .PrintTitles = False
            .RepeatItemsOnEachPrintedPage = True
            .TotalsAnnotation = False
            .CompactRowIndent = 1
            .InGridDropZones = False
            .DisplayFieldCaptions = True
            .DisplayMemberPropertyTooltips = False
            .DisplayContextTooltips = True
            .ShowDrillIndicators = True
            .PrintDrillIndicators = False
            .AllowMultipleFilters = False
            .SortUsingCustomLists = True
            .FieldListSortAscending = False
            .ShowValuesRow = False
            .CalculatedMembersInFilters = False
            .RowAxisLayout xlCompactRow
        End With
        With ActiveSheet.PivotTables("Tabela dinâmica3").PivotCache
            .RefreshOnFileOpen = False
            .MissingItemsLimit = xlMissingItemsDefault
        End With
        ActiveSheet.PivotTables("Tabela dinâmica3").RepeatAllLabels xlRepeatLabels
        ActiveWorkbook.ShowPivotTableFieldList = True
        With ActiveSheet.PivotTables("Tabela dinâmica3").PivotFields("Numero")
            .Orientation = xlRowField
            .Position = 1
        End With
        ActiveSheet.PivotTables("Tabela dinâmica3").AddDataField ActiveSheet. _
            PivotTables("Tabela dinâmica3").PivotFields("Valor Contabil"), _
            "Soma de Valor Contabil", xlSum
        ActiveSheet.PivotTables("Tabela dinâmica3").AddDataField ActiveSheet. _
            PivotTables("Tabela dinâmica3").PivotFields("Base Calculo ICMS"), _
            "Soma de Base Calculo ICMS", xlSum
        ActiveSheet.PivotTables("Tabela dinâmica3").AddDataField ActiveSheet. _
            PivotTables("Tabela dinâmica3").PivotFields("Valor ICMS"), "Soma de Valor ICMS" _
            , xlSum
        Columns("A:D").Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.Replace What:="Soma de ", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Application.CutCopyMode = False
        Columns("B:D").Select
        Selection.Style = "Currency"
        Cells.EntireColumn.AutoFit
        
    'COPIA E COLA DA SOMA
        Range("F2:H2").Copy
        Range("D2").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(0, 2).Select
        Range(Selection, Selection.End(xlUp)).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
    'COPIA PARA RESULTADOS
            Cells.Select
        Selection.AutoFilter
        ActiveWorkbook.Worksheets("JTXvsG5").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("JTXvsG5").AutoFilter.Sort.SortFields.Add2 Key:= _
            Range("F1:F83"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
            :=xlSortNormal
        With ActiveWorkbook.Worksheets("JTXvsG5").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        ActiveWorkbook.Worksheets("JTXvsG5").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("JTXvsG5").AutoFilter.Sort.SortFields.Add2 Key:= _
            Range("G1:G83"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
            :=xlSortNormal
        With ActiveWorkbook.Worksheets("JTXvsG5").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        ActiveWorkbook.Worksheets("JTXvsG5").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("JTXvsG5").AutoFilter.Sort.SortFields.Add2 Key:= _
            Range("H1:H83"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
            :=xlSortNormal
        With ActiveWorkbook.Worksheets("JTXvsG5").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        Selection.AutoFilter
        Range("A2:H16").Copy
        Sheets("Resultado").Select
        Range("W4").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    End Select
    End If
    
End Sub
