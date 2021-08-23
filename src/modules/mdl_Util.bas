Attribute VB_Name = "mdl_Util"
Option Explicit
Public Const cSenhaPlanilhas = 136479
Sub imprimirArrayNoConsole(pArray() As Variant)
    Dim y, z As Long
    Dim texto As String
    
    For y = 0 To UBound(pArray, 2)
        
        For z = 0 To UBound(pArray)
            texto = texto & "  |  " & pArray(z, y)
        Next z
            texto = texto & vbNewLine
    Next y
    
    Debug.Print texto
End Sub
Sub visualizacaoDePlanilhas(pMostrarPlanilhas As Boolean)

    Dim i As Integer
    
    'DESPROTEGE A PASTA INTEIRA
    ThisWorkbook.Unprotect cSenhaPlanilhas
    
    For i = 2 To ThisWorkbook.Sheets.Count
        ThisWorkbook.Worksheets(i).Visible = pMostrarPlanilhas
    Next i
    
    'PROTEGE TODA A PASTA DE TRABALHO
    ThisWorkbook.Protect cSenhaPlanilhas, Structure:=True, Windows:=True
    
End Sub

Sub protegeOuDesprotegePlanilhas(pProteger As Boolean)

    Dim i As Integer
    
    Application.ScreenUpdating = False

    For i = 1 To ThisWorkbook.Worksheets.Count
        If pProteger Then
            ThisWorkbook.Worksheets(i).Protect cSenhaPlanilhas, DrawingObjects:=True, Contents:=True, Scenarios:=True
        Else
            ThisWorkbook.Worksheets(i).Unprotect cSenhaPlanilhas
        End If
    Next i
    
    'ThisWorkbook.Save
    
    Application.ScreenUpdating = True
    
End Sub

'
'
'
'
'--------------------------------------------------
'esta função alimenta um array onde suas linhas podem ser
'redimensionadas/incrementadas preservando seus valores antigos
'
'
'Ex.: de redimencionamento
'ReDim Preserve arrayDados(UBound(arrayDados), linArray + 10)
'--------------------------------------------------
'
'
'
Function arrayComDadosDeTabela(codenamePlanilha As Worksheet) As Variant()

    Dim linha               As Long
    Dim coluna              As Long
    Dim ultimaLinha         As Long
    Dim ultimaColuna        As Long
    Dim linArray            As Long
    Dim colArray            As Long
    Dim arrayLocal() As Variant


    coluna = 1
    linha = 2
    ultimaLinha = codenamePlanilha.Cells(1, 1).End(xlDown).Row - 2 '-2 para excluir cabeçalho e matriz inicia com ZERO
    ultimaColuna = codenamePlanilha.Cells(1, coluna).End(xlToRight).Column - 1 'conta a partir dos cabeçalhos para evitar celulas em branco
    ReDim arrayLocal(ultimaColuna, ultimaLinha)
    
    'popula as LINHAS DO ARRAY
    For linArray = 0 To ultimaLinha
    
        'popula as COLUNAS DO ARRAY
        For colArray = 0 To ultimaColuna
        
            'popula array - percorre cada COLUNA populando E M   L I N H A S
            arrayLocal(colArray, linArray) = codenamePlanilha.Cells(linha, coluna).Value2
            
            'incrementa para a proxima coluna
            coluna = coluna + 1
        Next colArray
        
        'reseta a coluna para a proxima linha
        coluna = 1
        'incrementa a linha para reiniciar o loop e adicionar nova linha
        linha = linha + 1
    Next linArray
    
    arrayComDadosDeTabela = arrayLocal

End Function

Public Function somaItensListview(pNomeListview As ListView, pIndexColuna As Integer) As Double

    Dim soma As Double
    Dim i As Long
    Dim valor As Double

    For i = 1 To pNomeListview.ListItems.Count

        valor = pNomeListview.ListItems.Item(i).ListSubItems(pIndexColuna)

        soma = soma + valor

    Next i

    somaItensListview = soma

End Function

Public Sub populaCombos(nomeCombo As ComboBox, codenamePlanilha As Worksheet, pNumeroColunaComDados As Integer, linha As Integer, Optional pNumeroColunaId As Long = 0)
    
    nomeCombo.Clear
    
  
    Do While codenamePlanilha.Cells(linha, pNumeroColunaComDados) <> ""

        Dim itemLista As String
        
        If pNumeroColunaId > 0 Then
        
            itemLista = codenamePlanilha.Cells(linha, pNumeroColunaId) & "-" & codenamePlanilha.Cells(linha, pNumeroColunaComDados)
        
            nomeCombo.AddItem itemLista
        
        Else
            itemLista = codenamePlanilha.Cells(linha, pNumeroColunaComDados)
        
            nomeCombo.AddItem itemLista
        End If
       
    
        'redireciona o range para a proxima linha da coluna
        linha = linha + 1
    Loop



End Sub

Public Sub deletaRegistro(nomePlanilha As Worksheet, ByVal IdRegistroProcurado As Long, Optional pRegistroPossuiChavesEstrangeiras As Boolean = False, Optional pPlanilhaComRelacionamentos As Worksheet, Optional pNumeroColunaChaveEstrangeira As Long = 0)


    Dim verificaRegistroCadastrado As Range

    With nomePlanilha.Range("A:A")
        Set verificaRegistroCadastrado = .Find(IdRegistroProcurado, LookIn:=xlValues, LookAt:=xlWhole)
        If Not verificaRegistroCadastrado Is Nothing Then
            
            Dim confirma As String

            confirma = MsgBox("Deseja confirmar a exclusão de registro?", vbYesNo + vbExclamation, "Confirmação")
                
            If confirma = vbYes Then
                verificaRegistroCadastrado.EntireRow.Delete
                    
                '=====================================
                'deleta registros relacionados caso
                'haja relacionamento entre tabelas
                '=====================================
                If pRegistroPossuiChavesEstrangeiras And pNumeroColunaChaveEstrangeira > 0 Then
                        
                    Dim i As Long
                    Dim qtdRegistros As Long
                        
                    qtdRegistros = mdl_Util.ultimaLinhaEmBranco(pPlanilhaComRelacionamentos) - 1
                        
                    'loop AO CONTRARIO
                    For i = qtdRegistros To 1 Step -1
                            
                        If IdRegistroProcurado = pPlanilhaComRelacionamentos.Cells(i, pNumeroColunaChaveEstrangeira) Then
                            '=====================================
                            'deleta registro de baixo pra cima na tabela
                            '=====================================
                            pPlanilhaComRelacionamentos.Cells(i, pNumeroColunaChaveEstrangeira).EntireRow.Delete
                        End If
                            
                    Next i
                        
                End If
                    
                MsgBox "Registro excluído com sucesso...", vbInformation, "Exclusão de Registro"
            ElseIf confirma = vbNo Then
                Exit Sub
            End If
        Else: MsgBox "Registro nao encontrado.", vbCritical
        End If
    End With

End Sub
Sub deletaRegistrosRelacionados(pPlanilhaRelacionada As Worksheet, pChaveEstrangeira As Long, pLetraColunaChaveEstrangeira As String)
    '===================================
    ' fonte onde foi aprendido sobre uso do FINDNEXT:
    ' https://www.wallstreetmojo.com/vba-find-next/
    '===================================


    Dim incrementaArray As Integer
    Dim Rng As Range
    Dim FindRng As Range
    Dim primeiroRegistroEncontrado As String
    Dim linhasRegistrosEncontrados() As Long
    Dim i As Integer

    Set Rng = pPlanilhaRelacionada.Range(pLetraColunaChaveEstrangeira & ":" & pLetraColunaChaveEstrangeira)
    Set FindRng = Rng.Find(pChaveEstrangeira, LookIn:=xlValues, LookAt:=xlWhole)

    If Not FindRng Is Nothing Then
    
        primeiroRegistroEncontrado = FindRng.Address
        
        'armazena as linhas dos registros encontrados para posteriormente deleta-las
        Do
            ReDim Preserve linhasRegistrosEncontrados(incrementaArray)
            linhasRegistrosEncontrados(incrementaArray) = FindRng.Row
            incrementaArray = incrementaArray + 1
            Set FindRng = Rng.FindNext(FindRng)
            
        Loop While primeiroRegistroEncontrado <> FindRng.Address
        
        'deleta de baixo para cima
        For i = UBound(linhasRegistrosEncontrados) To 0 Step -1
            pPlanilhaRelacionada.Cells(linhasRegistrosEncontrados(i), 1).EntireRow.Delete
        Next i
        
    End If

End Sub
Function listaDadosRelacionadosViaChaveEstrangeira(pPlanilha As Worksheet, pChaveEstrangeira As Long, pLetraColunaChaveEstrangeira As String) As Variant()
    '===================================
    '
    ' UTIL PARA RELACIONAMENTO ENTRE TABELAS
    '
    ' TEM FUNÇÃO DE RETORNAR REGISTROS DE ACORDO
    ' COM A CHAVE ESTRANGEIRA PASSADA
    ' IDEAL PARA TRABALHAR COM RELACIONAMENTO ENTRE TABELA
    '
    '
    ' fonte onde foi aprendido sobre uso do FINDNEXT:
    ' https://www.wallstreetmojo.com/vba-find-next/
    '===================================
    
    Dim arrLocal() As Variant
    Dim linhaArrayIncrementada As Integer
    Dim qtdColunasNaTabela As Integer
    Dim Rng As Range
    Dim FindRng As Range
    Dim primeiroRegistroEncontrado As String
    Dim linhaRegistro As Long
    Dim coluna As Long
    
    'subtraido 1 devido index de array
    qtdColunasNaTabela = pPlanilha.Cells(1, 1).End(xlToRight).Column - 1

    Set Rng = pPlanilha.Range(pLetraColunaChaveEstrangeira & ":" & pLetraColunaChaveEstrangeira)
    Set FindRng = Rng.Find(pChaveEstrangeira, LookIn:=xlValues, LookAt:=xlWhole)

    If Not FindRng Is Nothing Then
    
        primeiroRegistroEncontrado = FindRng.Address
        
        'loop até retornar ao address inicial
        Do
            linhaRegistro = FindRng.Row
            ReDim Preserve arrLocal(qtdColunasNaTabela, linhaArrayIncrementada)

            For coluna = 0 To qtdColunasNaTabela
                
                arrLocal(coluna, linhaArrayIncrementada) = pPlanilha.Cells(linhaRegistro, coluna + 1).Value2
                
            Next coluna

            linhaArrayIncrementada = linhaArrayIncrementada + 1
            Set FindRng = Rng.FindNext(FindRng)

        Loop While primeiroRegistroEncontrado <> FindRng.Address
        
    End If

    listaDadosRelacionadosViaChaveEstrangeira = arrLocal

End Function
Function validaDelecaoDeRegistroComChavesEstrangeirasEmUso(pChaveId As Long, pNumeroColunaChaveId As Integer, pPlanilha As Worksheet, Optional pMensagem As String = "") As Boolean

    Dim i As Long
    Dim qtdRegistros As Long

    qtdRegistros = mdl_Util.ultimaLinhaEmBranco(pPlanilha) - 1

    For i = 2 To qtdRegistros
        If pChaveId = pPlanilha.Cells(i, pNumeroColunaChaveId) Then
            validaDelecaoDeRegistroComChavesEstrangeirasEmUso = True
            If pMensagem <> "" Then
                MsgBox pMensagem, vbCritical
            Else
                MsgBox "Não é possível excluir este registro. Outras tabelas ou formulários dependem das informações deste registro. Para excluir, é necessário antes excluir suas chaves estrangeiras/dependências.", vbCritical
            End If
            Exit For
        End If
    Next i

End Function

Function ultimaLinhaEmBranco(pPlanilha As Worksheet, Optional pIncluirPrimeiraLinha As Boolean = False) As Long
    
    If pIncluirPrimeiraLinha Then
        
        If pPlanilha.Cells(1, 1) = "" Then
            ultimaLinhaEmBranco = 1
        Else
            ultimaLinhaEmBranco = pPlanilha.Cells(1, 1).End(xlDown).Row + 1
        End If
        
    ElseIf pPlanilha.Cells(2, 1) = "" Then
        ultimaLinhaEmBranco = 2
    Else
        ultimaLinhaEmBranco = pPlanilha.Cells(1, 1).End(xlDown).Row + 1
    End If

End Function

Function extraiIdDaStringNaCombobox(pItemSelecionadoCombobox As String) As Long

    Dim posicaoCaractereSeparador As Integer

    posicaoCaractereSeparador = InStr(1, pItemSelecionadoCombobox, "-", vbTextCompare) - 1
    extraiIdDaStringNaCombobox = Left(pItemSelecionadoCombobox, posicaoCaractereSeparador)


End Function

'=======================================
'inserir no evento KeyPress do textBox
'=======================================
Sub textBoxSomenteNumerosOptionalMoeda(ByVal pKeyAscii As MSForms.ReturnInteger, Optional pMoeda As Boolean = False)

    Dim strValid As String
    
    If pMoeda Then
        'moeda permite virgula para casas decimais de centavos
        strValid = "0123456789,"
    Else
        strValid = "0123456789"
    End If
    
    If InStr(strValid, Chr(pKeyAscii)) = 0 Then
        pKeyAscii = 0
    End If

End Sub

'=======================================
'inserir no evento KeyDown do textBox
'
'FONTE: https://www.clubedohardware.com.br/topic/1158866-formatar-textbox-para-moeda/
'
'alterações feitas por: Emerson
'data: 04/10/2020
'=======================================
Sub formatoMoeda(ByVal KeyCode As MSForms.ReturnInteger, pTextBox As MSForms.textBox)

    'adaptação feita por Emerson 04/10/2020
    'para funcionar com o teclado numerico lateral
    'converte o keycode teclado numerico lateral para o teclado numerico superior
    Select Case KeyCode
    Case 96: KeyCode = 48
    Case 97: KeyCode = 49
    Case 98: KeyCode = 50
    Case 99: KeyCode = 51
    Case 100: KeyCode = 52
    Case 101: KeyCode = 53
    Case 102: KeyCode = 54
    Case 103: KeyCode = 55
    Case 104: KeyCode = 56
    Case 105: KeyCode = 57
    End Select


    Dim zTemp As String
    pTextBox.TextAlign = fmTextAlignRight
    If IsNumeric(Chr(KeyCode)) Or KeyCode = 8 Then
        If pTextBox.text <> "" Then
            zTemp = pTextBox.text & IIf(KeyCode <> 8, Chr(KeyCode), "")
            zTemp = Right(zTemp, Len(zTemp) - 2)
            zTemp = Replace(zTemp, ".", "")
            zTemp = Replace(zTemp, ",", "")
            If KeyCode = 8 Then
                If Len(zTemp) > 3 Then
                    zTemp = Left(zTemp, Len(zTemp) - 1)
                Else
                    zTemp = "0" & Left(zTemp, Len(zTemp) - 1)
                End If
            End If
            zTemp = Left(zTemp, Len(zTemp) - 2) & "." & Right(zTemp, 2)
        Else
            zTemp = "0.0" & IIf(KeyCode <> 8, Chr(KeyCode), "0")
        End If
        pTextBox.text = Format(Val(zTemp), "R$ ##,##0.00")
        KeyCode = 0
    Else
        If KeyCode <> 13 And KeyCode <> 9 And KeyCode <> 40 And KeyCode <> 38 Then KeyCode = 0
    End If

End Sub

Function geraId(pPlanilha As Worksheet) As Long
    
    Dim ultimaLinha As Long

    ultimaLinha = ultimaLinhaEmBranco(pPlanilha)
    
    If pPlanilha.Cells(2, 1) = "" Then
        geraId = 1
    Else
        geraId = pPlanilha.Cells(ultimaLinha - 1, 1) + 1
    End If
    
    
End Function

Function indexLinhaRegistroPorId(ByVal pId As Long, pPlanilha As Worksheet) As Long

    Dim registro As Range

    Set registro = pPlanilha.Range("A:A").Find(What:=pId, LookIn:=xlValues, LookAt:=xlWhole)

    If Not registro Is Nothing Then
        indexLinhaRegistroPorId = registro.Row
    Else: MsgBox "Registro nao encontrado.", vbCritical
        indexLinhaRegistroPorId = 0
    End If


End Function

'=========================================
'FUNÇÃO QUE VERIFICA SE UM ARRAY FOI
'INICIALIZADO, COM DADOS OU SE ESTA VAZIO
'=========================================
Function arrayIniciado(ByRef arr() As Variant, Optional pMostrarMensagem As Boolean = False) As Boolean
    On Error Resume Next
    arrayIniciado = IsNumeric(UBound(arr))
    If arrayIniciado = False Then
        If pMostrarMensagem Then
            MsgBox "Um array de dados vazio/não inicializado está sendo utilizado para coleta de dados." _
                 + vbNewLine + vbNewLine + "Provavelmente alguma pesquisa retornou uma lista vazia ou algum ID/registro não foi encontrado para que fosse possível popular tal array.", vbCritical
        End If
    End If
    On Error GoTo 0
End Function

Function validaSelecaoDeItemListview(pListview As ListView) As Boolean

    If pListview.ListItems.Count <= 0 Then
        MsgBox "Impossivel prosseguir. Não existem itens na lista.", vbCritical
        validaSelecaoDeItemListview = False
        Exit Function
    ElseIf pListview.SelectedItem Is Nothing Then
        MsgBox "Selecione um item para continuar a operação", vbExclamation
        validaSelecaoDeItemListview = False
        Exit Function
    Else
        validaSelecaoDeItemListview = True
    End If


End Function

'*****************************************************************************
'este procedimento depende da função "InvNumber" para funcionar
'-----------------------------------------------------------------------------
Public Sub classificaColunasListview(nomeListView As ListView, ByVal nomeColunaClicada As MSComctlLib.ColumnHeader)
    
    On Error Resume Next
       
    '    Começa ordenar o listview pela coluna clicada
    Dim vbHourglass
    
    With nomeListView
    
        ' Mostrar o cursor ampulheta enquanto faz o filtro
        
        Dim lngCursor As Long
        lngCursor = .MousePointer
        .MousePointer = vbHourglass
        
        'A rotina impede que o controle ListView faça atualização na tela
        'Isto é para esconder as mudanças que estão sendo feitas aos listitems
        'E também para acelerar o código
        
        'Verifique o tipo de dados da coluna de ser classificada,
        'para nomeá-la em conformidade
        
        Dim l As Long
        Dim strFormat As String
        Dim strData() As String
        
        Dim lngIndex As Long
        lngIndex = nomeColunaClicada.Index - 1
    
        '***************************************************************************
        ' Ordenar por data.
        
        Select Case UCase$(nomeColunaClicada.Tag)
        Case "DATE"
        
            
            
            strFormat = "YYYYMMDDHhNnSs"
        
            'O Loop através dos valores desta coluna organizam
            'As datas de modo que eles possam ser classificados em ordem alfabética,
        
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .text & Chr$(0) & .Tag
                            If IsDate(.text) Then
                                .text = Format(CDate(.text), _
                                               strFormat)
                            Else
                                .text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .text & Chr$(0) & .Tag
                            If IsDate(.text) Then
                                .text = Format(CDate(.text), _
                                               strFormat)
                            Else
                                .text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            ' Ordenar a lista em ordem alfabética por esta coluna
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = nomeColunaClicada.Index - 1
            .Sorted = True
            
            ' Restaura os valores anteriores das "células" nesta
            ' Coluna da lista das tags, e também restaura
            ' as tags para os valores originais
            
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, Chr$(0))
                            .text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            strData = Split(.Tag, Chr$(0))
                            .text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
            
            '***************************************************************************
            'Ordenar Numericamente
        
        Case "NUMBER"
        
       
            strFormat = String(30, "0") & "." & String(30, "0")
        
            ' Loop através dos valores desta coluna. Ordena os valores de modo que eles
            ' Podem ser classificados em ordem
        
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .text & Chr$(0) & .Tag
                            If IsNumeric(.text) Then
                                If CDbl(.text) >= 0 Then
                                    .text = Format(CDbl(.text), _
                                                   strFormat)
                                Else
                                    .text = "&" & InvNumber( _
                                            Format(0 - CDbl(.text), _
                                                   strFormat))
                                End If
                            Else
                                .text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .text & Chr$(0) & .Tag
                            If IsNumeric(.text) Then
                                If CDbl(.text) >= 0 Then
                                    .text = Format(CDbl(.text), _
                                                   strFormat)
                                Else
                                    .text = "&" & InvNumber( _
                                            Format(0 - CDbl(.text), _
                                                   strFormat))
                                End If
                            Else
                                .text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            ' Ordenar a lista em ordem alfabética por esta coluna
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = nomeColunaClicada.Index - 1
            .Sorted = True
            
                      
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, Chr$(0))
                            .text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            strData = Split(.Tag, Chr$(0))
                            .text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
        
        Case Else                                ' Assume ordenação como string
            
            
        
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = nomeColunaClicada.Index - 1
            .Sorted = True
            
        End Select
   
        .MousePointer = lngCursor
    
    End With
    
End Sub

'*****************************************************************************
'InvNumber
'Função usada para permitir que os números negativos possam ser classificados
'-----------------------------------------------------------------------------
Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function


