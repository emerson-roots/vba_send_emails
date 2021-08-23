VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTitulosVencidos 
   Caption         =   "SELECIONAR CLIENTES PARA ENVIO DE EMAIL"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15840
   OleObjectBlob   =   "frmTitulosVencidos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTitulosVencidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cabecalhoListview()

    With lstTitulosVencidos
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
            
        ' Cabeçalhos
        .ColumnHeaders.Clear
        .ColumnHeaders.Add(, , "#", width:=20, Alignment:=0).Tag = ""
        .ColumnHeaders.Add(, , "CLIENTE", width:=200, Alignment:=0).Tag = ""
        .ColumnHeaders.Add(, , "VALOR", width:=70, Alignment:=0).Tag = "NUMBER"
        .ColumnHeaders.Add(, , "DATA VENCIMENTO", width:=80, Alignment:=0).Tag = "DATE"
        .ColumnHeaders.Add(, , "N° FATURA", width:=80, Alignment:=0).Tag = ""
        .ColumnHeaders.Add(, , "ADMINISTRAÇÃO", width:=145, Alignment:=0).Tag = ""
        .ColumnHeaders.Add(, , "E-MAILS", width:=170, Alignment:=0).Tag = ""
            
    End With

End Sub

Sub preencheListview(pDados() As Variant)
    
    Dim itemDaLista As ListItem
    Dim i As Long
    
    lstTitulosVencidos.ListItems.Clear
    For i = 1 To UBound(pDados, 2)
        Set itemDaLista = lstTitulosVencidos.ListItems.Add(text:="")
        itemDaLista.SubItems(1) = pDados(0, i)
        itemDaLista.SubItems(2) = Format(pDados(1, i), "R$ #,##0.00")
        itemDaLista.SubItems(3) = pDados(2, i)
        itemDaLista.SubItems(4) = pDados(3, i)
        itemDaLista.SubItems(5) = pDados(4, i)
        itemDaLista.SubItems(6) = pDados(5, i)
    Next i
    
    lstTitulosVencidos.SelectedItem = Nothing
    
End Sub

Private Sub btnEnviarEmails_Click()
    
    
    Dim i As Long
    Dim y As Long
    Dim linhaArr As Long
    Dim arrSelecionadosNoListview() As Variant
    Dim arrAdmsWithEmails() As Variant
    Dim arrDadosSeparadosPorAdms() As Variant
    Dim adm As String
    
    ' armazena itens selecionados no listview
    For i = 1 To lstTitulosVencidos.ListItems.Count
        If lstTitulosVencidos.ListItems(i).Checked Then
            ReDim Preserve arrSelecionadosNoListview(5, linhaArr)
            
            arrSelecionadosNoListview(0, linhaArr) = lstTitulosVencidos.ListItems(i).ListSubItems(1)
            arrSelecionadosNoListview(1, linhaArr) = lstTitulosVencidos.ListItems(i).ListSubItems(2)
            arrSelecionadosNoListview(2, linhaArr) = lstTitulosVencidos.ListItems(i).ListSubItems(3)
            arrSelecionadosNoListview(3, linhaArr) = lstTitulosVencidos.ListItems(i).ListSubItems(4)
            arrSelecionadosNoListview(4, linhaArr) = lstTitulosVencidos.ListItems(i).ListSubItems(5)
            arrSelecionadosNoListview(5, linhaArr) = lstTitulosVencidos.ListItems(i).ListSubItems(6)
            
            linhaArr = linhaArr + 1
        End If
    
    Next i
    
    ' filtra os selecionados, retirando as adms duplicadas
    arrAdmsWithEmails = listaAdministradorasComEmailsExcluindoDuplicados(arrSelecionadosNoListview)
    
    ' separa os titulos por administradora e envia um e-mail com uma lista personalizada para cada adm
    For i = LBound(arrAdmsWithEmails, 2) To UBound(arrAdmsWithEmails, 2)
        
        linhaArr = 0
        adm = arrAdmsWithEmails(0, i)
        
        For y = LBound(arrSelecionadosNoListview, 2) To UBound(arrSelecionadosNoListview, 2)
            If adm = arrSelecionadosNoListview(4, y) Then
                ReDim Preserve arrDadosSeparadosPorAdms(5, linhaArr)
                
                arrDadosSeparadosPorAdms(0, linhaArr) = arrSelecionadosNoListview(0, y)
                arrDadosSeparadosPorAdms(1, linhaArr) = arrSelecionadosNoListview(1, y)
                arrDadosSeparadosPorAdms(2, linhaArr) = arrSelecionadosNoListview(2, y)
                arrDadosSeparadosPorAdms(3, linhaArr) = arrSelecionadosNoListview(3, y)
                arrDadosSeparadosPorAdms(4, linhaArr) = arrSelecionadosNoListview(4, y)
                arrDadosSeparadosPorAdms(5, linhaArr) = arrSelecionadosNoListview(5, y)
                
                linhaArr = linhaArr + 1
            End If
            
        Next y
        
        Call enviarEmail(arrDadosSeparadosPorAdms)

        ' apaga array para popular novos dados e novo e-mail
        Erase arrDadosSeparadosPorAdms
        
    Next i
    
    MsgBox "Emails enviado com sucesso...", vbInformation
    
End Sub

'
' adicionar referencia Microsoft Scripting Runtime
'
' ou setar uma variavel do tipo Object com set = CreateObject("Scripting.Dictionary")
'
Private Function listaAdministradorasComEmailsExcluindoDuplicados(vArray() As Variant) As Variant

    Dim dicionarioAdms As Dictionary
    'Dim oDict As Object
    Dim i As Long
    Dim y As Long
    Dim arrLocal() As Variant
    
    Set dicionarioAdms = New Dictionary
    
    ' percore array de dados utilizando a bibioteca Dictionary para armazenar somente
    ' indexes não repetidos
    For i = LBound(vArray, 2) To UBound(vArray, 2)
        dicionarioAdms(vArray(4, i)) = True
    Next
    
    ' redimensiona array de acordo com o tamanho do dicionario
    ReDim Preserve arrLocal(1, UBound(dicionarioAdms.Keys()))
    
    ' popula um array bidirecional com as administradoras e seus e-mails correspondentes
    For i = LBound(dicionarioAdms.Keys()) To UBound(dicionarioAdms.Keys())
        arrLocal(0, i) = dicionarioAdms.Keys(i)
        
        For y = LBound(vArray, 2) To UBound(vArray, 2)
            If vArray(4, y) = arrLocal(0, i) Then
                arrLocal(1, i) = vArray(5, y)
                Exit For
            End If
            
        Next y
    Next i
    
    listaAdministradorasComEmailsExcluindoDuplicados = arrLocal
    
End Function
Private Sub btnImportarDados_Click()
    Dim arrDados() As Variant
    
    arrDados = importaDadosDeArquivoExcelParaArray
    
    If arrayIniciado(arrDados) Then
        Call preencheListview(arrDados)
    Else
        lstTitulosVencidos.ListItems.Clear
    End If
End Sub

Private Sub btnInserirEmail_Click()
    
    Dim email As String
    
    If lstTitulosVencidos.ListItems.Count <= 0 Then
        MsgBox "Impossivel prosseguir. Não existem itens na lista. Importe dados para a lista para poder manipular e-mails.", vbCritical
        Exit Sub
    ElseIf lstTitulosVencidos.SelectedItem Is Nothing Then
        MsgBox "Selecione um item para inserir/alterar e-mail", vbExclamation
        Exit Sub
    Else
        Dim adm As String
        email = InputBox("Digite um e-mail válido...")
        adm = lstTitulosVencidos.ListItems.Item(lstTitulosVencidos.SelectedItem.Index).ListSubItems(5)
        
        If emailIsValid(email) Then
            Call alteraEmailPorAdministradora(email, adm)
        Else
            MsgBox "O e-mail não é válido.", vbCritical
        End If
    End If
    
    
End Sub

Function emailIsValid(pEmail As String) As Boolean

    If InStr(1, pEmail, "@", 1) = 0 Or InStr(1, pEmail, ".", 1) = 0 Then
        emailIsValid = False
    Else
        emailIsValid = True
    End If
    
End Function

Function admIsValid(pAdm As String) As Boolean

    If InStr(1, pAdm, "null -", 1) <> 0 Then
        admIsValid = False
    Else
        admIsValid = True
    End If
    
End Function

Sub alteraEmailPorAdministradora(pEmail As String, pAdm As String)

    Dim indexItemSelecionado As Integer
    Dim qtdItensNaLista As Integer
    Dim i As Integer
    Dim adm As String
    
    
    qtdItensNaLista = lstTitulosVencidos.ListItems.Count
    indexItemSelecionado = lstTitulosVencidos.SelectedItem.Index
    
    
    For i = 1 To qtdItensNaLista
        
        adm = lstTitulosVencidos.ListItems.Item(i).ListSubItems(5)
        
        If adm = pAdm Then
            lstTitulosVencidos.ListItems.Item(i).SubItems(6) = pEmail
        End If
        
    Next i
    



End Sub

Private Sub lstTitulosVencidos_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lstTitulosVencidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call classificaColunasListview(lstTitulosVencidos, ColumnHeader)
End Sub

Private Sub lstTitulosVencidos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    Dim i As Integer
    Dim marcado As Boolean
    
    If itemMarcadoListviewIsValid(Item) = False Then Exit Sub
    
    'verifica se existem itens marcados no listview
    For i = 1 To lstTitulosVencidos.ListItems.Count
        If lstTitulosVencidos.ListItems.Item(i).Checked Then
            marcado = True
            Exit For
        End If
    Next i
    
    If marcado = True Then
        btnEnviarEmails.Enabled = True
    Else
        btnEnviarEmails.Enabled = False
    End If

End Sub
Function itemMarcadoListviewIsValid(pItemMarcado As MSComctlLib.ListItem) As Boolean
    
    Dim email As String
    Dim adm As String
    
    email = pItemMarcado.ListSubItems(6)
    adm = pItemMarcado.ListSubItems(5)
    
    If InStr(1, email, "@", 1) = 0 And InStr(1, email, ".com", 1) = 0 Then
        pItemMarcado.Checked = False
        MsgBox "A linha selecionada não possui um e-mail válido. Não é possível selecionar.", vbCritical
        itemMarcadoListviewIsValid = False
    ElseIf InStr(1, adm, "null", 1) <> 0 Then
        pItemMarcado.Checked = False
        MsgBox "O cliente não tem vínculos com administradora. É possível que não esteja na carteira de clientes. Não é possível selecionar", vbCritical
        itemMarcadoListviewIsValid = False
    Else
        itemMarcadoListviewIsValid = True
    End If
    
    
End Function
Function importaDadosDeArquivoExcelParaArray() As Variant()


Dim wkbOrigem As Workbook
Dim caminho As String
Dim dados() As Variant
Dim ultimaColuna As Long
Dim ultimaLinha As Long
Dim i, y As Long

caminho = caixaDialogoEscolherArquivoExcel

    'caminho em branco aborta a rotina ou arquivo nao selecionado
    If caminho = "" Then
        Erase importaDadosDeArquivoExcelParaArray
        Exit Function
    End If
    
Set wkbOrigem = Workbooks.Open(caminho)
ActiveWindow.Visible = False

'ultima linha -4 devido cabeçalhos gerados pelo site Itau
ultimaLinha = wkbOrigem.Worksheets(1).Cells(1, 1).End(xlDown).Row - 1
ultimaColuna = 5

ReDim dados(ultimaColuna, ultimaLinha)

    'popula array com os dados da planilha externa
    For i = 0 To ultimaColuna
        For y = 0 To ultimaLinha
            dados(i, y) = wkbOrigem.Worksheets(1).Cells(y + 1, i + 1)
        Next y
    Next i
'copia array para a função
importaDadosDeArquivoExcelParaArray = dados


    
    'testa algumas coluinas do array para verificar se armazenou os dados de acordo com o layout de colunas necessário
    If wkbOrigem.Worksheets(1).Cells(1, 1) <> "CLIENTE" And wkbOrigem.Worksheets(1).Cells(1, 2) <> "VALOR" And wkbOrigem.Worksheets(1).Cells(1, 3) <> "DATA VENCIMENTO" Then
        MsgBox "O arquivo selecionado não possui o layout padrao necessário. Não será possivel efetuar as baixas de faturas.", vbCritical
        'fecha a planilha
        Erase dados
        wkbOrigem.Close False
        Exit Function
    End If
    
    'fecha a planilha
    wkbOrigem.Close False
    

End Function
Function caixaDialogoEscolherArquivoExcel() As String

   Dim caixaDialogo As FileDialog
   Dim verificaBotaoApertadoCaixaDialogo As Integer
   Dim arquivoSelecionado As String
   
   Set caixaDialogo = Application.FileDialog(msoFileDialogFilePicker)
     
     With caixaDialogo
        .Filters.Clear
        .Filters.Add "Planilha títulos vencidos", "*.xlsm"
        .ButtonName = "Selecionar Planilha"
        .InitialFileName = ""
        .Title = "Selecione a planilha com relatório de títulos vencidos"
        .AllowMultiSelect = False
        verificaBotaoApertadoCaixaDialogo = .Show

        'se o usuario clicar em cancelar ou nao salvar com outro nome, fecha o arquivo modelo
        If verificaBotaoApertadoCaixaDialogo > -1 Then
            
            MsgBox "Nenhum arquivo foi selecionado.", vbCritical
            
        Else
            'armazena o diretorio+arquivo do item selecionado na caixa de dialogo
            arquivoSelecionado = caixaDialogo.SelectedItems.Item(1)
        End If
        
        caixaDialogoEscolherArquivoExcel = arquivoSelecionado
        
     End With
End Function

Private Sub lstTitulosVencidos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim indexItemSelecionado As Integer
    Dim adm As String
    Dim email As String

    indexItemSelecionado = lstTitulosVencidos.SelectedItem.Index
    
    email = lstTitulosVencidos.ListItems.Item(indexItemSelecionado).ListSubItems(6)
    adm = lstTitulosVencidos.ListItems.Item(indexItemSelecionado).ListSubItems(5)
    
    If emailIsValid(email) = False Then
        btnInserirEmail.Enabled = False
    Else
        btnInserirEmail.Enabled = True
    End If
    
    If admIsValid(adm) = False Then
        btnInserirEmail.Enabled = False
    Else
        btnInserirEmail.Enabled = True
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    
        
    Call mdl_Util.protegeOuDesprotegePlanilhas(False)
        
    Call cabecalhoListview
    
    'simula 2 clicks no cabecalho para iniciar ordenando de A a Z
    Call lstTitulosVencidos_ColumnClick(lstTitulosVencidos.ColumnHeaders.Item(2))
    Call lstTitulosVencidos_ColumnClick(lstTitulosVencidos.ColumnHeaders.Item(2))
    
    lstTitulosVencidos.SelectedItem = Nothing
    
    Application.ScreenUpdating = False
    
End Sub

Private Sub UserForm_Terminate()
    Call mdl_Util.protegeOuDesprotegePlanilhas(True)
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    
End Sub
