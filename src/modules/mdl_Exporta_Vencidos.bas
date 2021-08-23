Attribute VB_Name = "mdl_Exporta_Vencidos"
Option Explicit

Public objExcel As Excel.Application
Public objWb As Excel.Workbook
Public objWs As Excel.Worksheet
Sub gerarPlanilhaComDadosDeFaturasVencidas()
    
    Dim arrFaturasManutencao() As Variant
    Dim arFaturasVenda() As Variant
    
    arrFaturasManutencao = listaDeFaturasComStatusVencido(wsFaturamentoManutencao, 2, 6, 7, 8, 11)
    arFaturasVenda = listaDeFaturasComStatusVencido(wsVendasFaturamento, 2, 11, 9, 10, 18)
    
    Call gerarPlanilhaVencidos(arrFaturasManutencao, True, False, True)
    Call gerarPlanilhaVencidos(arFaturasVenda, False, True, False)

End Sub

Function listaDeFaturasComStatusVencido(pWorkSheet As Worksheet, _
                        pClienteIndexColuna As Integer, _
                        pValorFaturaIndexColuna As Integer, _
                        pDataVencimentoIndexColuna As Integer, _
                        pNumeroFaturaIndexColuna As Integer, _
                        pStatusVencidoIndexColuna As Integer) As Variant()


    Dim qtdRegistros As Long
    Dim i As Long
    Dim arrLocal() As Variant
    Dim incrementaLinhaArray As Integer
    Dim cliente As String
    Dim adm As String
    Dim emails As String

    qtdRegistros = mdl_Util.ultimaLinhaEmBranco(pWorkSheet) - 1

    For i = 2 To qtdRegistros
        If pWorkSheet.Cells(i, pStatusVencidoIndexColuna) = "VENCIDO" Then
            ReDim Preserve arrLocal(5, incrementaLinhaArray)
            
            cliente = pWorkSheet.Cells(i, pClienteIndexColuna)
            adm = buscaAdministradoraCliente(cliente)
            emails = buscaEmailAdministradora(adm)
            
            arrLocal(0, incrementaLinhaArray) = cliente
            arrLocal(1, incrementaLinhaArray) = pWorkSheet.Cells(i, pValorFaturaIndexColuna)  ' valor
            arrLocal(2, incrementaLinhaArray) = pWorkSheet.Cells(i, pDataVencimentoIndexColuna)  ' vencimento
            arrLocal(3, incrementaLinhaArray) = pWorkSheet.Cells(i, pNumeroFaturaIndexColuna)  ' numero fatura
            arrLocal(4, incrementaLinhaArray) = adm
            arrLocal(5, incrementaLinhaArray) = emails
            
            incrementaLinhaArray = incrementaLinhaArray + 1
        End If
        
    Next i
    
    
    listaDeFaturasComStatusVencido = arrLocal
    
    'Call imprimirArrayNoConsole(arrLocal)

End Function
Function buscaAdministradoraCliente(pCliente As String) As String

    Dim registro As Range
    Dim adm As String

    Set registro = wsCadastroClientes.Range("B:B").Find(What:=pCliente, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not registro Is Nothing Then
        adm = wsCadastroClientes.Cells(registro.Row, 4)
        buscaAdministradoraCliente = adm
    Else: buscaAdministradoraCliente = "null - Cliente não encontrado. Não esta na carteira de clientes!"
       
    End If
    

End Function
Function buscaEmailAdministradora(pAdm As String) As String

    Dim registro As Range
    Dim emails As String
    
    'somente tras e-mail para cliente que nao é AUTO GESTAO
    If pAdm <> "Auto Gestao" Then
        Set registro = wsAdministracoes.Range("B:B").Find(What:=pAdm, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not registro Is Nothing Then
            emails = wsAdministracoes.Cells(registro.Row, 8)
            
            'verifica se possui e-mail cadastrado
            If emails = "" Then
                emails = "Administradora não possui e-mails cadastrados."
            End If
            
            buscaEmailAdministradora = emails
        Else: buscaEmailAdministradora = "Não esta na carteira de clientes!"
           
        End If
    Else
        buscaEmailAdministradora = "Auto Gestao - Verifique o email manualmente."
    End If
    

End Function

Sub gerarPlanilhaVencidos(arrayDadosDeTitulos() As Variant, pSetarNovoWorkbook As Boolean, pLimparWorkbookDaMemoria As Boolean, pGerarCabecalhos As Boolean)
    
'    Dim objExcel As Excel.Application
'    Dim objWb As Excel.Workbook
'    Dim objWs As Excel.Worksheet
    Dim i As Integer
    Dim ultimaLinha As Long

    
    'excel
    If pSetarNovoWorkbook Then
        Set objExcel = New Excel.Application
        Set objWb = objExcel.Workbooks.Add
        Set objWs = objWb.Sheets(1)
    End If
    
    If pGerarCabecalhos Then
        Call criaCabecalhos(objWs)
    End If
    
    
    
    'arrays
    With objWs
        Call insereDados(objWs, arrayDadosDeTitulos)
        
        .Columns.AutoFit
    End With
    

    objExcel.Visible = True
    objExcel.DisplayAlerts = False
    
    If pLimparWorkbookDaMemoria Then
        'objExcel.Quit
        Set objWs = Nothing
        Set objWb = Nothing
        Set objExcel = Nothing
    End If
    

End Sub

Sub insereDados(pWorkSheet As Excel.Worksheet, pDados() As Variant)
    
    Dim ultimaLinha As Long
    Dim i As Long
    Dim y As Long
    Dim qtdLinhasArray As Long
    Dim qtdColunasArray As Long
    
    qtdLinhasArray = UBound(pDados, 2)
    qtdColunasArray = 5 'de acordo com index 0

    With pWorkSheet
        ultimaLinha = mdl_Util.ultimaLinhaEmBranco(pWorkSheet)
        ultimaLinha = ultimaLinha
        For i = 0 To qtdLinhasArray
            For y = 0 To qtdColunasArray
                .Cells(ultimaLinha, y + 1) = pDados(y, i)
            Next y
            ultimaLinha = ultimaLinha + 1
        Next i
    End With
    
End Sub

Sub criaCabecalhos(pWorkSheet As Excel.Worksheet)
    
    Dim i As Integer
    
    With pWorkSheet
    
        .Cells(1, 1) = "CLIENTE"
        .Cells(1, 2) = "VALOR"
        .Cells(1, 3) = "DATA VENCIMENTO"
        .Cells(1, 4) = "N° FATURA"
        .Cells(1, 5) = "ADMINISTRAÇÃO"
        .Cells(1, 6) = "EMAIL"
        
        
        For i = 1 To 6
            .Cells(1, i).Interior.color = Cores.VERDE_CLARO
            .Cells(1, i).Font.Bold = True
            .Cells(1, i).HorizontalAlignment = xlCenter
            .Cells(1, i).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            .Cells(1, i).Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            .Cells(1, i).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            .Cells(1, i).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

        Next i
        
    End With
    
End Sub

