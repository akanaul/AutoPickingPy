Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub AjustarDados()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim dataAmanha As Date

    Set ws = ThisWorkbook.Sheets("Ajuste")

    ' --- validação básica ---
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ultimaLinha < 2 Then
        MsgBox "Não há dados abaixo do cabeçalho.", vbExclamation
        Exit Sub
    End If

    ' calcula data de amanhã
    dataAmanha = Date + 1

    ' garante que exista AutoFilter na faixa (se já existir, mantém)
    If Not ws.AutoFilterMode Then
        ws.Range("A1:K" & ultimaLinha).AutoFilter
    End If

    ' --- 1) FILTRAR COLUNA E: mostrar tudo que é <> "Fábrica ITU" e excluir essas linhas visíveis ---
    ws.Range("A1:K" & ultimaLinha).AutoFilter Field:=5, Criteria1:="<>FABRICA ITU"

    On Error Resume Next
    ' seleciona só as células visíveis (linha de cabeçalho fica fora porque começamos em A2)
    ws.Range("A2:K" & ultimaLinha).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    On Error GoTo 0

    ' limpa filtros
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' atualiza última linha após exclusões
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ultimaLinha < 2 Then
        MsgBox "Não restaram dados após a exclusão por fábrica.", vbInformation
        Exit Sub
    End If

    ' --- 2) FILTRAR COLUNA H: excluir tudo que NÃO for a data de amanhã ---
    ' Para evitar problemas de formatação de texto/locale, comparamos pelo número serial da data:
    ' mostramos as linhas com data < dataAmanha  OU  >= (dataAmanha + 1) — essas são as que queremos apagar.
    ws.Range("A1:K" & ultimaLinha).AutoFilter Field:=8, _
        Criteria1:="<" & CLng(dataAmanha), Operator:=xlOr, _
        Criteria2:=">=" & CLng(dataAmanha + 1)

    On Error Resume Next
    ws.Range("A2:K" & ultimaLinha).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    On Error GoTo 0

    ' limpa filtro final
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    ' autoajusta todas as colunas da tabela após concluir
ws.Columns("A:K").AutoFit


End Sub

Sub TransferirParaCabreuva()

    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinha As Long

    Set wsOrigem = ThisWorkbook.Sheets("Ajuste")
    Set wsDestino = ThisWorkbook.Sheets("CABREUVA")

    ' Descobre última linha com dados na aba Ajuste
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row

    If ultimaLinha < 2 Then
        MsgBox "Não há dados na aba Ajuste para transferir.", vbExclamation
        Exit Sub
    End If

    ' -----------------------------
    ' 1) Coluna A ? CABREUVA!A3
    ' -----------------------------
    wsOrigem.Range("A2:A" & ultimaLinha).Copy
    wsDestino.Range("A3").PasteSpecial xlPasteValues

    ' -----------------------------
    ' 2) Coluna C ? CABREUVA!B3
    ' -----------------------------
    wsOrigem.Range("C2:C" & ultimaLinha).Copy
    wsDestino.Range("B3").PasteSpecial xlPasteValues

    ' -----------------------------------------
    ' 3) Colunas G, H, I ? CABREUVA!C3, D3, E3
    ' -----------------------------------------
    wsOrigem.Range("G2:I" & ultimaLinha).Copy
    wsDestino.Range("C3").PasteSpecial xlPasteValues

    ' -----------------------------
    ' 4) Coluna K ? CABREUVA!F3
    ' -----------------------------
    wsOrigem.Range("K2:K" & ultimaLinha).Copy
    wsDestino.Range("F3").PasteSpecial xlPasteValues
    
    ' ----------------------------------------------------
    ' AJUSTES FINAIS (centralizar, formatar, inserir títulos)
    ' ----------------------------------------------------

    ' Centraliza toda a planilha CABREUVA
    wsDestino.Cells.HorizontalAlignment = xlCenter

    ' Formata coluna A como DATA
    wsDestino.Range("A3:A" & (ultimaLinha + 1)).NumberFormat = "dd/mm/yyyy"

    ' Formata colunas E e F como HORA
    wsDestino.Range("E3:F" & (ultimaLinha + 1)).NumberFormat = "hh:mm"

    ' Insere títulos na linha 2
    wsDestino.Range("A2").Value = "DT"
    wsDestino.Range("B2").Value = "STO"
    wsDestino.Range("C2").Value = "DESTINO"
    wsDestino.Range("D2").Value = "DATA SAÍDA"
    wsDestino.Range("E2").Value = "HORA SUB PICKING"
    wsDestino.Range("F2").Value = "HORA SAÍDA CARGA"

End Sub
Sub LimparPlanilhas()

    Dim ws As Worksheet
    Dim nomes As Variant
    Dim i As Long
    
    'Nomes das abas que serão limpas
    nomes = Array("CABREUVA", "AJUSTE")
    
    For i = LBound(nomes) To UBound(nomes)
        
        Set ws = ThisWorkbook.Sheets(nomes(i))
        
        '--- LIMPAR FILTROS ---
        On Error Resume Next
        If ws.AutoFilterMode Then ws.AutoFilter.ShowAllData
        ws.AutoFilterMode = False
        On Error GoTo 0
        
        '--- APAGAR APENAS OS DADOS (SEM REMOVER FORMATAÇÃO) ---
        ws.Cells.ClearContents
        
    Next i
    
    MsgBox ("Limpeza realizada com sucesso")

End Sub
Sub AlinharConverterEFormatarCabreuva_vFinal()

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = Sheets("CABREUVA")

    '--- Alinhar A1:AF1 ---
    With ws.Range("A1:F1")
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With

    '--- Converter COLUNA A para formato GERAL ---
    With ws
        .Columns("A").NumberFormat = "General"
        .Columns("A").TextToColumns Destination:=.Range("A1"), DataType:=xlFixedWidth
    End With

    '==========================================================
    '     BORDAS APENAS NO BLOCO REAL A3:F(ÚLTIMA LINHA)
    '==========================================================

    ' Encontra a última linha REAL com dados entre A e F
    lastRow = ws.Range("A:F").Find("*", SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious).Row

    ' Se não houver dados, sair
    If lastRow < 3 Then Exit Sub

    ' Aplica bordas somente no bloco de dados
    With ws.Range("A3:F" & lastRow).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With

End Sub
Sub START()
    Call AjustarDados
    Call TransferirParaCabreuva
    Call AlinharConverterEFormatarCabreuva_vFinal
    Sheets("Orientação").Select
    
    MsgBox ("Sucesso!")
    
End Sub


