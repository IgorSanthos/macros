Attribute VB_Name = "Módulo1"
Sub ENTRADAS()
'
' ENTRADAS
'

'IMPORTACAO JETTAX
        Dim caminhoUso As String
        Dim wbUso As Workbook
        Dim wbSaida As Workbook
        
        MsgBox "SELECIONE A PLANILHA DO JETTAX"
        
    'Abertura de planilha
        
        caminhoUso = Application.GetOpenFilename(FileFilter:="Arquivos do JETTAX (*.xlsx; *.xls), *.xlsx; *.xls", Title:="Selecione a Planilha")
        If caminhoUso <> "Falso" Then
            Set wbUso = Workbooks.Open(caminhoUso)
            wbUso.Activate
    
    'Validação
        Select Case ActiveSheet.Name
        Case "Relatório Detalhado por Nota", "Relatório Detalhado por Produto"
    
    'LIMPAR PLANILHA
        Windows("ANALISE_DE_ENTRADAS.xlsm").Activate
        Sheets("Itens").Select
        Rows("1:" & Selection.End(xlDown).Row).ClearContents
        
        Sheets("IMPORITEM").Select
        Rows("3:" & Selection.End(xlDown).Row).ClearContents
        
        Sheets("Identificação NFE").Select
        Rows("1:" & Selection.End(xlDown).Row).ClearContents
        
        Sheets("JETTAX").Select
        Rows("3:" & Selection.End(xlDown).Row).ClearContents
        
        Sheets("MENU").Select
        
    'COPIA DA PLANILHA DO JETTAX
        wbUso.Activate
        Sheets("Relatório Detalhado por Produto").Select
        Range("A:A,K:K,N:N").Copy
        Windows("ANALISE_DE_ENTRADAS.xlsm").Activate
        Sheets("JETTAX").Select
        Range("B1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        Range("A2").Copy
        Range("B1").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(0, -1).Select
        Range(Selection, Selection.End(xlUp)).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
    'FECHANDO PLANILHA
        Set wbUso = Workbooks.Open(caminhoUso)
            wbUso.Activate
            wbUso.Close SaveChanges:=False
        Case Else
            wbUso.Close SaveChanges:=False
            MsgBox "PLANILHA DO JETTAX INCORRETA"
            Exit Sub
        End Select
        End If

'IMPORTAÇÃO PLANILHA DO SYSCONV
        MsgBox "SELECIONE A PLANILHA DE XML DO SYSCONV"
        
        Dim caminhoArquivo As String
        Dim wbOrigem As Workbook
            'Dim wbDestino As Workbook
    'Importar planilha
        caminhoArquivo = Application.GetOpenFilename(FileFilter:="Arquivos do SYSCONV (*.xlsx; *.xls), *.xlsx; *.xls", Title:="Selecione a Planilha")
    
    'Verifica se o usuário selecionou o arquivo e abrir
        If caminhoArquivo <> "Falso" Then
            Set wbOrigem = Workbooks.Open(caminhoArquivo)
            wbOrigem.Activate
    
    'Validar planilha
        Select Case ActiveSheet.Name
            Case "Identificação NFE", "Emitente", "Destinatário", "Entrega", "Autorizadas", _
                 "Itens", "Total", "Transportadora", "Cobrança", "Pagamento", _
                 "Inf. Adicional", "Assinatura", "Protocolo"
    
    'COPIAR DA PLANILHA SYSCONV - Itens
        wbOrigem.Activate
        Sheets("Identificação NFE").Select
        Range("A:A,B:B,C:C").Copy
        Windows("ANALISE_DE_ENTRADAS.xlsm").Activate
        Sheets("Identificação NFE").Select
        Range("A1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
    'COPIAR DA PLANILHA SYSCONV - NFE
        wbOrigem.Activate
        Sheets("Itens").Select
        Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Copy
        Windows("ANALISE_DE_ENTRADAS.xlsm").Activate
        Sheets("Itens").Select
        Range("B1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
    
'PROCV DAS CATEGORIAS
        Range("A2").Select
        ActiveCell.FormulaR1C1 = _
            "=VLOOKUP(VLOOKUP(RC[1],'Identificação NFE'!C:C[2],3,0)&RC[4],JETTAX!C:C[52],4,0)"
        Range("A2").Copy
        Range("B2").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(0, -1).Select
        Range(Selection, Selection.End(xlUp)).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        Range("A1").Select
        Selection.AutoFilter
        Columns("DT:DT").Select
        Selection.Replace What:="1", Replacement:="2", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "General"
        
'ERRO DE CLASSIFICAÇÃO
    'FILTRO VALOR CONTABIL
        Range("$A$1:$AM$999999").AutoFilter Field:=1, Criteria1:="#N/D", _
            Operator:=xlAnd, Criteria2:="#N/D"
            
    'Defina o intervalo onde você aplicou o filtro
        Set Rng = ThisWorkbook.Sheets("Itens").Range("B2:B999999")
    
    'Verifique se há células visíveis após aplicar o filtro
        resultadoFiltro = Application.WorksheetFunction.Subtotal(103, Rng)
    
    'Verifique se há informações no resultado do filtro
        If resultadoFiltro = 0 Then
            Else
                'Se houver informações
                    MsgBox "             ATENÇÃO" & vbCrLf & vbCrLf & _
                    "Notas com erro de classificação verificar MENU", vbCritical
                    
                'Limpar MENU
                    Sheets("MENU").Select
                    Range("G3:L999999").ClearContents

                'Copiar e colar na aba Resultado
                    Sheets("Itens").Select
                    Range("B1:C1").Select
                    Range(Selection, Selection.End(xlDown)).Copy
                    Sheets("MENU").Select
                    Range("K2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
                    Sheets("Itens").Select
                    ActiveSheet.ShowAllData
        End If


CST_ST
CFOP_CONSUMOST
CFOP_CONSUMO
CST_CONSUMO
INDUSTRIALIZACAO
CSOSN
ATIVO
BONIFICAÇAO
BENEFICIAMENTO
REMESSA
Selection.AutoFilter
CADITEN_TIPOITEM
FINAL

    'Final da planilha do sysconv
        Case Else
            wbOrigem.Close SaveChanges:=False
            MsgBox "PLANILHA SYSCONV INCORRETA"
        End Select
    'Copiando informações para o sysconv
        Windows("ANALISE_DE_ENTRADAS.xlsm").Activate
        Sheets("Itens").Select
        Range("B1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Copy
        wbOrigem.Activate
        Sheets("Itens").Select
        Range("A1").Select
        ActiveSheet.Paste
        wbOrigem.Close SaveChanges:=True
        MsgBox "Planilha Sysconv Salvo"
    End If

    
End Sub
Sub REMESSA()
    'FILTRO
        ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=16, Criteria1:="*916", Operator:=xlOr, Criteria2:="*915"
    'CST
        Columns("DX:DX").Select
        Selection.Replace What:="", Replacement:="41", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="**", Replacement:="41", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "@"
        Range("DX1").FormulaR1C1 = "ICMS_CST"
        ActiveSheet.ShowAllData

End Sub
Sub BENEFICIAMENTO()
    'FILTRO
        ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=1, Criteria1:="Beneficiamento"
    'CST
        Columns("DX:DX").Select
        Selection.Replace What:="", Replacement:="90", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="**", Replacement:="90", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "@"
        Range("DX1").FormulaR1C1 = "ICMS_CST"
        ActiveSheet.ShowAllData
    'LIMPANDO ICMS
        ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=128, Criteria1:="90"
        Range("EF:EF, EG:EG, EK:EK").Select
        Selection.Replace What:="*", Replacement:="0", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
        Selection.NumberFormat = "@"
        Range("EF1").FormulaR1C1 = "ICMS_vBC"
        Range("EG1").FormulaR1C1 = "ICMS_pICMS"
        Range("EK1").FormulaR1C1 = "ICMS_vICMS"
        ActiveSheet.ShowAllData

End Sub
Sub INDUSTRIALIZACAO()

    ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=1, Criteria1:="Industrialização", Operator:=xlOr, Criteria2:="Matéria Prima"
        
    Columns("P:P").Select
    
    Selection.Replace What:="5102", Replacement:="5101", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.Replace What:="6102", Replacement:="6101", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.Replace What:="6404", Replacement:="6401", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.Replace What:="6403", Replacement:="6401", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.Replace What:="5403", Replacement:="5401", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.Replace What:="5405", Replacement:="5401", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.NumberFormat = "General"
    ActiveSheet.ShowAllData

End Sub
Sub ATIVO()

'TROCANDO CFOP  ATIVO COM ST

With Selection
        ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=1, Criteria1:="Ativo Imobilizado"
        ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=128, Criteria1:="60"
        'TROCAR CFOP PARA 5557
        Columns("P:P").Select
        Selection.Replace What:="5405", Replacement:="5406", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5403", Replacement:="5406", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5401", Replacement:="5406", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5929", Replacement:="5406", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6401", Replacement:="6406", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6403", Replacement:="6406", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6404", Replacement:="6406", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6929", Replacement:="6406", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6108", Replacement:="6406", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

'TROCANDO CFOP  DE Ativo Imobilizado TRIBUTADO
        ActiveSheet.ShowAllData
        ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=1, Criteria1:="Ativo Imobilizado"
        'MUDAR CST PARA 90
        Columns("DX:DX").Select
        Selection.Replace What:="", Replacement:="90", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="**", Replacement:="90", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "@"
        Range("DX1").FormulaR1C1 = "ICMS_CST"
        
        ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=128, Criteria1:="90"
        'TROCAR CFOP PARA 5556
        Columns("P:P").Select
        Selection.Replace What:="5101", Replacement:="5991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5102", Replacement:="5991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5103", Replacement:="5991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5120", Replacement:="5991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5105", Replacement:="5991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5106", Replacement:="5991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5929", Replacement:="5991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6101", Replacement:="6991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6120", Replacement:="6991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6102", Replacement:="6991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6103", Replacement:="6991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6105", Replacement:="6991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6106", Replacement:="6991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6929", Replacement:="6991", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "General"
        ActiveSheet.ShowAllData
    End With



    
End Sub
Sub CFOP_CONSUMOST()



'TROCANDO CFOP  DE USO E CONSUMO COM ST
    With Selection
        ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=1, Criteria1:="Uso e Consumo"
        ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=128, Criteria1:="60"
        'TROCAR CFOP PARA 5557
        Columns("P:P").Select
        Selection.Replace What:="5405", Replacement:="5407", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5403", Replacement:="5407", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5401", Replacement:="5407", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5929", Replacement:="5407", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6401", Replacement:="6407", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6403", Replacement:="6407", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6404", Replacement:="6407", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6929", Replacement:="6407", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6108", Replacement:="6407", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "General"
        ActiveSheet.ShowAllData
    End With
    
End Sub
Sub CFOP_CONSUMO()

'GASOLINA

    ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=9, Criteria1:="2710*"
    Columns("P:P").Select
    Selection.Replace What:="5929", Replacement:="5653", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="6929", Replacement:="5653", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    Selection.NumberFormat = "General"
    ActiveSheet.ShowAllData
    
'TROCANDO CFOP  DE USO E CONSUMO TRIBUTADO
        ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=1, Criteria1:="Uso e Consumo"
        ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=128, Criteria1:="<>60"
        'TROCAR CFOP PARA 5556
        Columns("P:P").Select
        Selection.Replace What:="5101", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5102", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5103", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5105", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5120", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5106", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="5929", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6101", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6102", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6103", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6105", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6106", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6120", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6108", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="6929", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "General"
        ActiveSheet.ShowAllData
        

    
End Sub
Sub BONIFICAÇAO()
    ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=1, Criteria1:="Bonificação, Doação ou Brinde"
        
    'MUDAR CST PARA 90
    Columns("DX:DX").Select
    Selection.Replace What:="", Replacement:="90", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="**", Replacement:="90", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.NumberFormat = "@"
    Range("DX1").FormulaR1C1 = "ICMS_CST"
    Dim filterCriteria As Variant
    filterCriteria = Array("5102", "5103", "5101", "5105", "5106", "6102", "6103", "6101", "6105", "6106")
    ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=16, Criteria1:=filterCriteria, Operator:=xlFilterValues
    
    Columns("P:P").Select
    Selection.Replace What:="5101", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="5102", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="5103", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="5105", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="5106", Replacement:="5999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="6101", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="6102", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="6103", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="6105", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="6106", Replacement:="6999", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.NumberFormat = "General"
    ActiveSheet.ShowAllData

End Sub

Sub CST_ST()

    'FILTRO TUDO QUE TEM ST
    ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=1, Criteria1:="Revenda"
    ActiveSheet.Range("$A$1:$HB$999999").AutoFilter Field:=164, Criteria1:=">", Operator:=xlOr, Criteria2:=">0"

    'MUDAR CST PARA 60
    Columns("DX:DX").Select
    Selection.Replace What:="**", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.NumberFormat = "General"
    ActiveSheet.ShowAllData
    
    '5405
    ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=16, Criteria1:="5405*"
    Selection.Replace What:="**", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveSheet.ShowAllData
    '5403
    ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=1, Criteria1:="Revenda"
    ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=16, Criteria1:="5403*"
    Selection.Replace What:="**", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveSheet.ShowAllData
    '5401
    ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=1, Criteria1:="Revenda"
    ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=16, Criteria1:="5401*"
    Selection.Replace What:="**", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveSheet.ShowAllData
    '6401
    ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=1, Criteria1:="Revenda"
    ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=16, Criteria1:="6401"
    Selection.Replace What:="**", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveSheet.ShowAllData
    '6403
    ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=16, Criteria1:="6403"
    Selection.Replace What:="**", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveSheet.ShowAllData
    '6404
    ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=1, Criteria1:="Revenda"
    ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=16, Criteria1:="6404"
    Selection.Replace What:="**", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveSheet.ShowAllData
    Selection.NumberFormat = "General"
    Range("DX1").FormulaR1C1 = "ICMS_CST"
    Range("DY2:DZ999999").ClearContents
    
End Sub
Sub CST_CONSUMO()

'TRIBUTADO
    With Selection
        'FILTRO
                
        ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=1, Criteria1:="Uso e Consumo", Operator:=xlOr, Criteria2:="Ativo Imobilizado"
        ActiveSheet.Range("$A$1:$HC$999999").AutoFilter Field:=128, Criteria1:="<>60"
        'MUDAR CST PARA 90
        Columns("DX:DX").Select
        Selection.Replace What:="", Replacement:="90", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.Replace What:="**", Replacement:="90", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "@"
        Range("DX1").FormulaR1C1 = "ICMS_CST"
        ActiveSheet.ShowAllData
    End With
    
End Sub
Sub CSOSN()

'CSOSN 101
    'FILTRO
        ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=125, Criteria1:="101"
        Dim filterCriteria As Variant
            filterCriteria = Array("Revenda", "Industrialização", "Bonificação, Doação ou Brinde", "Remessa", "Matéria Prima", "Beneficiamento")
            ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=1, Criteria1:=filterCriteria, Operator:=xlFilterValues
        Columns("DX:DX").Select
        Selection.Replace What:="", Replacement:="00", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "00"
        Range("DX1").FormulaR1C1 = "ICMS_CST"
        ActiveSheet.ShowAllData
    

'CSOSN 102
    'FILTRO
        ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=125, Criteria1:="102"
            filterCriteria = Array("Revenda", "Industrialização", "Bonificação, Doação ou Brinde", "Remessa", "Matéria Prima", "Beneficiamento")
            ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=1, Criteria1:=filterCriteria, Operator:=xlFilterValues
        Columns("DX:DX").Select
        Selection.Replace What:="", Replacement:="41", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "General"
        Range("DX1").FormulaR1C1 = "ICMS_CST"
        ActiveSheet.ShowAllData

'CSOSN 103
    'FILTRO
        ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=125, Criteria1:="102"
            filterCriteria = Array("Revenda", "Industrialização", "Bonificação, Doação ou Brinde", "Remessa", "Matéria Prima", "Beneficiamento")
            ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=1, Criteria1:=filterCriteria, Operator:=xlFilterValues
        Columns("DX:DX").Select
        Selection.Replace What:="", Replacement:="40", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "General"
        Range("DX1").FormulaR1C1 = "ICMS_CST"
        ActiveSheet.ShowAllData
    
'CSOSN 300
    'FILTRO
        ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=125, Criteria1:="300"
            filterCriteria = Array("Revenda", "Industrialização", "Bonificação, Doação ou Brinde", "Remessa", "Matéria Prima", "Beneficiamento")
            ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=1, Criteria1:=filterCriteria, Operator:=xlFilterValues
        Columns("DX:DX").Select
        Selection.Replace What:="", Replacement:="41", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "General"
        Range("DX1").FormulaR1C1 = "ICMS_CST"
        ActiveSheet.ShowAllData
    
'CSOSN 400
    'FILTRO
        ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=125, Criteria1:="400"
            filterCriteria = Array("Revenda", "Industrialização", "Bonificação, Doação ou Brinde", "Remessa", "Retorno", "Matéria Prima")
            ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=1, Criteria1:=filterCriteria, Operator:=xlFilterValues
        Columns("DX:DX").Select
        Selection.Replace What:="", Replacement:="41", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "General"
        Range("DX1").FormulaR1C1 = "ICMS_CST"
        ActiveSheet.ShowAllData
    
'CSOSN 500
    'FILTRO
        ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=125, Criteria1:="500"
            filterCriteria = Array("Revenda", "Industrialização", "Bonificação, Doação ou Brinde", "Remessa", "Retorno", "Matéria Prima")
            ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=1, Criteria1:=filterCriteria, Operator:=xlFilterValues
        Columns("DX:DX").Select
        Selection.Replace What:="", Replacement:="60", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Selection.NumberFormat = "General"
        Range("DX1").FormulaR1C1 = "ICMS_CST"
        ActiveSheet.ShowAllData
    
'CSOSN 900
        'FILTRO
            ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=125, Criteria1:="900"
                filterCriteria = Array("Revenda", "Industrialização", "Bonificação, Doação ou Brinde", "Remessa", "Retorno", "Matéria Prima")
                ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=1, Criteria1:=filterCriteria, Operator:=xlFilterValues
            Columns("DX:DX").Select
            Selection.Replace What:="", Replacement:="90", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.NumberFormat = "General"
            Range("DX1").FormulaR1C1 = "ICMS_CST"
            ActiveSheet.ShowAllData

' QUANDO NAO TEM CST E NEM CSOSN
            ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=125, Criteria1:="="
            ActiveSheet.Range("$A$1:$IM$999999").AutoFilter Field:=128, Criteria1:="="
            Columns("DX:DX").Select
            Selection.Replace What:="", Replacement:="90", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Selection.NumberFormat = "General"
            ActiveSheet.ShowAllData
        'LIMPANDO CSOSN
            Range("DU2:DU999999").ClearContents
    
'MUDAR TIPO DO CSOSN
            Range("DS:DS").Select
            Selection.Replace What:="", Replacement:="90", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            Application.CutCopyMode = False
            Range("DS2").Select
            Range("DS2").FormulaR1C1 = "=RC[5]"
            Range("DS2").Copy
            Range(Selection, Selection.End(xlDown)).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            Selection.NumberFormat = "General"
    
'CST 00 e 'Tipo_ICMS
            Range("DX:DX, DS:DS").Select
            ActiveSheet.Range("$B$1:$IM$999999").AutoFilter Field:=127, Criteria1:="=0", _
                Operator:=xlAnd, Criteria2:="=0"
            ActiveSheet.ShowAllData
            Selection.NumberFormat = "00"
            

End Sub
Sub PISCOFINS()

Range("HK2:HM999999").ClearContents
Range("HR2:HT999999").ClearContents

'CST PIS
Range("HJ2:HJ999999").Select
Selection.Replace What:="**", Replacement:="73", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
Selection.NumberFormat = "General"

'CST COFINS
Range("HQ2:HQ999999").Select
Selection.Replace What:="**", Replacement:="73", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
Selection.NumberFormat = "General"


End Sub

Sub CADITEN_TIPOITEM()

    
'_TIPOITEM
' CST TRIBUTADO
'

'
    ActiveSheet.Range("$B$1:$IM$999999").AutoFilter Field:=127, Criteria1:="00"
    ActiveWorkbook.Worksheets("Itens").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Itens").AutoFilter.Sort.SortFields.Add2 Key:=Range _
        ("DX1:DX284"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Itens").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("DX1").Select
    Selection.End(xlDown).Select
    Selection.Copy
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlUp).Select
    ActiveCell.FormulaR1C1 = "ICMS_CST"
    Range("DY1").Select
    Selection.AutoFilter
    
'_TIPOITEM
    Sheets("IMPORITEM").Select
    Sheets("Itens").Select
    Columns("E:E").Copy
    Sheets("IMPORITEM").Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").FormulaR1C1 = "COD_ITEM"
    Range("B2:AO2").Copy
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
End Sub

Sub FINAL()

    Dim abaSelecionada As Worksheet
    Dim caminhoDestino As Variant

    ' Verifica se há pelo menos uma aba ativa
    If ActiveSheet Is Nothing Then
        MsgBox "Nenhuma aba ativa selecionada.", vbExclamation
        Exit Sub
    End If

    ' Obtém a aba ativa
    Set abaSelecionada = ThisWorkbook.Sheets("IMPORITEM")

    ' Abre o diálogo de salvamento
    MsgBox "Escolha onde salvar a planilha de Itens", vbExclamation
    caminhoDestino = Application.GetSaveAsFilename(FileFilter:="Arquivos do Excel de Itens(*.xlsx), *.xlsx")

    ' Verifica se o usuário clicou em Cancelar
    If caminhoDestino = "Falso" Then
        Exit Sub
    End If
    ' Salva a aba ativa como um novo arquivo
    Windows("ANALISE_DE_ENTRADAS.xlsm").Activate
    Sheets("IMPORITEM").Select
    abaSelecionada.Copy
    ActiveWorkbook.SaveAs caminhoDestino
    ActiveWorkbook.Close SaveChanges:=True
    MsgBox "Planilha de ITENS salva"
End Sub





