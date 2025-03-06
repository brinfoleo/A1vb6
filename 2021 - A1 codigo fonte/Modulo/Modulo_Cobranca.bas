Attribute VB_Name = "Modulo_Cobranca"
Option Explicit
Type CliPosicaoFinanceira
    Pagar      As String
    Pago       As String
End Type
Public Sub AtualizarBoleto(idBol As Long, novaData As String)
    '#################################################################################################
    '### 04/01/2012
    '### Funcao que atualiza os dados na tabela imprime o boleto e volta ao valor anterior
    '#################################################################################################
    Dim dVencNominal       As String
    Dim ValorNominal        As String
    
    
    
    Dim dtCalc              As String
    Dim MultaMora           As String
    Dim vCob                As String
    
    Dim vReg(10)    As Variant
    Dim cReg        As Integer
    
    dtCalc = novaData
    dVencNominal = PgDadosFinanceiroFatura(idBol).Vencimento
    ValorNominal = PgDadosFinanceiroFatura(idBol).vlDuplicata
    
    'Modifica os valores do boleto
    MultaMora = Val(ChkVal(AtualizaCobranca(idBol, dtCalc).vMora, 0, cDecMoeda)) + Val(ChkVal(AtualizaCobranca(idBol, dtCalc).vMulta, 0, cDecMoeda))
    MultaMora = ChkVal(MultaMora, 0, cDecMoeda)
    vCob = Val(ChkVal(ValorNominal, 0, cDecMoeda)) + Val(ChkVal(MultaMora, 0, cDecMoeda)) + Val(ChkVal(PgDadosFinanceiroFatura(idBol).Acrescimo, 0, cDecMoeda))
    vCob = Val(ChkVal(vCob, 0, cDecMoeda)) - (Val(ChkVal(PgDadosFinanceiroFatura(idBol).Deducoes, 0, cDecMoeda)) + Val(ChkVal(PgDadosFinanceiroFatura(idBol).Abatimento, 0, cDecMoeda)))
    vCob = ChkVal(vCob, 0, cDecMoeda)
    'Grava os Novos Valores
    cReg = 0
    vReg(cReg) = Array("Vencimento", dtCalc, "D"): cReg = cReg + 1
    vReg(cReg) = Array("MultaMora", MultaMora, "S"): cReg = cReg + 1
    vReg(cReg) = Array("vlCobrado", vCob, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ObsBol3", "Vencimento / Valor original: " & dVencNominal & " / " & ConvMoeda(ValorNominal), "S"): cReg = cReg + 1
    cReg = cReg - 1
    RegistroAlterar "FinanceiroContasPRCadastro", vReg, cReg, "id=" & idBol
    
    RegLogDataBase 0, "0", "0", "Atualizacao: Duplicata " & PgDadosFinanceiroFatura(idBol).NumDuplicata & " (id:" & idBol & ") atualizado para " & dtCalc & "."
    
    'Imprime o boleto atualizado
    BoletoBancario (idBol)
    
    'Volta com os valores nominais da Fatura
    vReg(cReg) = Array("Vencimento", dVencNominal, "D"): cReg = cReg + 1
    vReg(cReg) = Array("MultaMora", "0", "S"): cReg = cReg + 1
    vReg(cReg) = Array("vlCobrado", "0", "S"): cReg = cReg + 1
    vReg(cReg) = Array("ObsBol3", "", "S"): cReg = cReg + 1
    cReg = cReg - 1
    RegistroAlterar "FinanceiroContasPRCadastro", vReg, cReg, "id=" & idBol
    
End Sub
Private Function BoletoBancario_001_NossoNumero(Id As Long) As String
    '########################################################################################################
    '### Montagem de NOSSO NUMERO
    '### 28.02.25
    '########################################################################################################
    '# Multiplicador base 9
    Dim NN1, NN2 As String
    NN1 = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).convenio
    NN1 = Left(String(7, "0"), 7 - Len(Trim(NN1))) & Trim(NN1)
    If Len(Trim(Id)) > 10 Then
            NN2 = Right(Id, 10)
        Else
            NN2 = Left(String(10, "0"), 10 - Len(Trim(Id))) & Trim(Id)
    End If
    BoletoBancario_001_NossoNumero = "000" & NN1 & NN2
End Function

Function Calculo_DV10(strNumero As String) As String
    'declara As variáveis
    Dim intContador As Integer
    Dim intNumero As Integer
    Dim intTotalNumero As Integer
    Dim intMultiplicador As Integer
    Dim intResto As Integer

    ' se nao for um valor numerico sai da função
    If Not IsNumeric(strNumero) Then
        Calculo_DV10 = ""
        Exit Function
    End If

    'inicia o multiplicador
    intMultiplicador = 2

    'pega cada caracter do numero a partir da direita
    For intContador = Len(strNumero) To 1 Step -1

        'extrai o caracter e multiplica pelo multiplicador
        intNumero = Val(Mid(strNumero, intContador, 1)) * intMultiplicador

        ' se o resultado for maior que nove soma os algarismos do resultado
        If intNumero > 9 Then
            Do Until intNumero < 10 '<= 10
                intNumero = Val(Left(intNumero, 1)) + Val(Right(intNumero, 1))
            Loop
        End If

        'soma o resultado para totalização
        intTotalNumero = intTotalNumero + intNumero

        'se o multiplicador for igual a 2 atribuir valor 1 se for 1 atribui 2
        intMultiplicador = IIf(intMultiplicador = 2, 1, 2)

    Next

    Dim DezenaSuperior As Integer
    Dim dv As Integer
    If intTotalNumero < 10 Then
            DezenaSuperior = 10
        Else
            DezenaSuperior = 10 * (Val(Left(CStr(intTotalNumero), 1)) + 1)
    End If
    intResto = intTotalNumero Mod 10 'DezenaSuperior - intTotalNumero
    dv = Right(DezenaSuperior - intResto, 1)
    
    'verifica as exceções ( 0 -> DV=0 )
    Select Case intResto
        Case 0
            Calculo_DV10 = "0"
        Case Else
            Calculo_DV10 = str(dv)
    End Select

End Function
Private Function cnab240LoteHeader(contaId As Integer, lote As String) As String
'Layout CNAB 240 - v10.1
    'Dim  'file Destination
    Dim line As String
    'fd As String
    'fd = arqDestino
   
   Dim dvAgCc As String
   Dim faturaId As Long
   faturaId = 0
   dvAgCc = ""
    
    '***********************************************************
    '*** Registro Header de Lote ***
    '***********************************************************
     
    line = ""
    line = line & F("n", 3, pgDadosBanco(PgDadosFinanceiroFatura(faturaId).IdBanco).Numero)
    line = line & F("n", 4, lote) 'Lote
    line = line & "1"
     
    line = line & "R" '04.1 - operacao R arquivo remessa
    line = line & "01" '05.1 - servico
    line = line & "  " '06.1 - cnab
    line = line & "060" '07.1- layout do arquivo
     
    line = line & " "  'CNAB
    
    line = line & "0" '09.1 - Tipo de inscricao
    line = line & F("n", 15, PgDadosEmpresa(ID_Empresa).CNPJ)
    line = line & F("a", 20, pgDadosConta(contaId).convenio)
    line = line & F("n", 5, pgDadosConta(contaId).agencia)
    line = line & F("n", 1, pgDadosConta(contaId).AgenciaDV)
    line = line & F("n", 12, pgDadosConta(contaId).conta)
    line = line & F("n", 1, pgDadosConta(contaId).ContaDV)
    line = line & F("n", 1, dvAgCc)
    line = line & F("a", 30, PgDadosEmpresa(ID_Empresa).Nome) '17.1 - nome da empresa
    
    line = line & String(40, " ") '18.1 - Informação 1
     line = line & String(40, " ") '19.1 - Informação 2
    
'    line = line & F("a", 30, PgDadosEmpresa(ID_Empresa).Lgr)
'    line = line & F("n", 5, PgDadosEmpresa(ID_Empresa).Nro)
'    line = line & F("a", 15, PgDadosEmpresa(ID_Empresa).Cpl)
'    line = line & F("a", 20, PgDadosEmpresa(ID_Empresa).Mun)
'    Dim CEP As String
'    CEP = F("n", 8, PgDadosEmpresa(ID_Empresa).CEP)
'    line = line & Mid(CEP, 1, 5)
'    line = line & Mid(CEP, 6, 3)
'    line = line & F("a", 2, PgDadosEmpresa(ID_Empresa).UF)
'
    line = line & F("n", 8, "0") '20.1
    line = line & F("n", 8, "0") '21.1
    
    line = line & F("n", 8, "0") '22.1 - dt credito
    line = line & String(33, " ") 'CNAB
    
    
    'grvFile fd, line
    cnab240LoteHeader = line

End Function

Private Function cnab240Q(faturaId As Long, lote As String) As String
    '***********************************************************
    '*** Registro Detalhe - Segmento Q (Obrigatorio Remessa) ***
    '***********************************************************
    Dim line As String
    line = ""
    line = line & F("n", 3, pgDadosBanco(PgDadosFinanceiroFatura(faturaId).IdBanco).Numero)
    line = line & F("n", 4, lote)
    line = line & "3"
    line = line & F("n", 5, "1")
    line = line & "Q"
    line = line & F("n", 1, " ")
    line = line & F("n", 2, "0") '07.3Q - Cod movimento
    line = line & "0" '08.3Q - tipo inscricao
    line = line & F("n", 15, "0") '09.3Q - numero inscricao
    line = line & F("a", 40, PgDadosFinanceiroFatura(faturaId).Sacado)
    line = line & F("a", 40, PgDadosCliente(PgDadosFinanceiroFatura(faturaId).IDSacado).Lgr & _
                    PgDadosCliente(PgDadosFinanceiroFatura(faturaId).IDSacado).Nro)
    line = line & F("a", 15, PgDadosCliente(PgDadosFinanceiroFatura(faturaId).IDSacado).Bairro)
    line = line & F("n", 5, Mid(PgDadosCliente(PgDadosFinanceiroFatura(faturaId).IDSacado).CEP, 1, 5))
    line = line & F("n", 3, Mid(PgDadosCliente(PgDadosFinanceiroFatura(faturaId).IDSacado).CEP, 6, 3))
    
    line = line & F("a", 15, PgDadosCliente(PgDadosFinanceiroFatura(faturaId).IDSacado).Mun)
    line = line & F("a", 2, PgDadosCliente(PgDadosFinanceiroFatura(faturaId).IDSacado).uf)
    'Sac/Avalista
    line = line & "0"
    line = line & F("n", 15, "0")
    line = line & F("n", 40, " ")
    '---
    line = line & F("n", 3, pgDadosBanco(PgDadosFinanceiroFatura(faturaId).IdBanco).Numero)
    line = line & F("a", 20, PgDadosFinanceiroFatura(faturaId).NossoNumero)
    line = line & F("a", 8, " ")
    'grvFile fd, line
    '*** FIM Seguimento Q ***
    cnab240Q = line
    
End Function

Public Function cobCalcMora(sValor As String, iDiasVencidos As Integer, pMoraDia As String, sRetorno As String) As String
    'sRetorno = T = total da Mora
    '           D = Mora do dia
    
    Dim vMora As String
    
    'Achar o valor da mora ao dia
    vMora = Val(ChkVal(sValor, 0, cDecMoeda)) * Val(ChkVal(pMoraDia, 0, 3)) / 100
    vMora = ChkVal(vMora, 0, cDecMoeda)
    
    Select Case UCase(sRetorno)
        Case "T"
            vMora = Val(vMora) * Val(iDiasVencidos)
        Case "D"
            'Mantem o valor achado
        Case Else
            vMora = 0
    End Select
    cobCalcMora = ChkVal(vMora, 0, cDecMoeda)
        
    
End Function

Public Function cobCalcMulta(sValor As String, pMultaMes As String, Optional iMeses = 1) As String
    Dim vMulta As String
    vMulta = Val(ChkVal(Val(ChkVal(sValor, 0, cDecMoeda)) * Val(ChkVal(pMultaMes, 0, 3)), 0, cDecMoeda)) / 100
    vMulta = Val(ChkVal(vMulta, 0, cDecMoeda)) * Val(iMeses)
    vMulta = ChkVal(vMulta, 0, cDecMoeda)
    cobCalcMulta = vMulta
End Function

Private Function F(Tipo As String, tam As Integer, txtInformativo As Variant) As String
    ' Funcao que retorna as string formatada conforme layout p/CNAB240
    On Error GoTo trtErroF
    Dim nTexto As String
    Dim texto As String
    texto = Trim(txtInformativo)
    
    If tam < Len(texto) And Tipo <> "d" Then
        'Debug.Print " *** string maior que o campo ***"
        texto = Mid(texto, 1, tam)
    End If
    
    
    If Len(texto) > tam And LCase(Tipo) = "a" Then
            nTexto = Mid(texto, 1, tam)
            MsgBox "String maior que o valor (" & tam & ") passado:" & vbCrLf & _
            "texto ant.: [" & texto & "]" & vbCrLf & _
            "novo texto: [" & nTexto & "]"
            
        Else
    
            Select Case LCase(Tipo)
                Case "a" 'Alfanumerico
                    nTexto = texto & String(tam - Len(texto), " ")
                Case "n" 'Numerico
                    texto = Replace(texto, ".", "")
                    nTexto = String(tam - Len(texto), "0") & texto
                Case "d"
                    If tam >= 8 Then
                            nTexto = Format(texto, "ddmmYYYY")
                        Else
                            nTexto = Format(texto, "ddmmYY")
                    End If
                Case Else
                    
                    MsgBox "Tipo de dado (" & Tipo & ") não localizado"
                    nTexto = ""
            End Select
    End If
    F = nTexto
    'Debug.Print Len(F)
    Exit Function
trtErroF:
    Debug.Print "*************************************************************"
    Debug.Print "Função F (modulo_Cobranca)"
    Debug.Print "   String: " & txtInformativo
    Debug.Print "   Descrição: " & Err.Number & " - " & Err.Description
    Debug.Print "*************************************************************"
    Resume Next
    
End Function

Public Function Monta_CodBarras(banco As String, _
                                Moeda As String, _
                                Valor As Single, _
                                Vencimento As Date, _
                                agencia As String, _
                                conta As String, _
                                NossoNumero As String, _
                                dvLinhaDig As Integer)

    Dim codigo_sequencia As String
    Dim database As Date
    Dim fator As Integer
    Dim intDac As Integer

    'database para calculo do fator
    database = CDate("03/07/2000")
    fator = DateDiff("d", database, Format(Vencimento, "dd/mm/yyyy"))
    fator = fator + 1000
    Valor = Int(Valor * 100)
    'Livre = Format(Livre, "0000000000000000000000000")

    ' sequencia sem o DV
    codigo_sequencia = banco & Moeda & fator & Format(Valor, "0000000000") & agencia & conta & dvLinhaDig & Left(String(13, "0"), 13 - Len(NossoNumero)) & NossoNumero

    ' calculo do DV
    intDac = calcula_DV_CodBarras(codigo_sequencia)

    ' monta a sequencia para o codigo de barras com o DV
    Monta_CodBarras = Left(codigo_sequencia, 4) & intDac & Right(codigo_sequencia, 39)

End Function
Private Function calcula_DV_CodBarras(sequencia As String) As Integer

    Dim intContador, intNumero, intTotalNumero As Integer
    Dim intMultiplicador, intResto, intresultado As Integer
    Dim Caracter As String

    intMultiplicador = 2

    For intContador = 1 To 43
        Caracter = Mid(Right(sequencia, intContador), 1, 1)
        If intMultiplicador > 9 Then
            intMultiplicador = 2
            intNumero = 0
        End If
        intNumero = Caracter * intMultiplicador
        intTotalNumero = intTotalNumero + intNumero
        intMultiplicador = intMultiplicador + 1
    Next

    intResto = intTotalNumero Mod 11

    intresultado = 11 - intResto

    If intresultado = 10 Or intresultado = 11 Then
            calcula_DV_CodBarras = 1
        Else
            calcula_DV_CodBarras = intresultado
    End If

End Function
Private Function CalculoFator(Vencimento As Date) As Integer
    Dim fator As Integer
    Dim database As Date
    'database para calculo do fator
    database = CDate("03/07/2000")
    fator = DateDiff("d", database, Format(Vencimento, "dd/mm/yyyy"))
    fator = fator + 1000
    CalculoFator = fator
    'valor = Int(valor * 100)
    'Livre = Format(Livre, "0000000000000000000000000")
End Function
Function Calculo_NossoNumero(sequencia As String) As String
    'montamos o nosso numero com o numero do convenio ( 6 posicoes)
    Dim dv As Integer
    sequencia = IIf(Trim(sequencia) = "", "0", sequencia)
    dv = Calculo_DV11(sequencia)
    Calculo_NossoNumero = Format(sequencia & dv, "0000000000000")

End Function

Function Calculo_DV11(strNumero As String) As String
    'declara as variáveis
    Dim intContador As Integer
    Dim intNumero As Integer

    Dim intTotalNumero As Integer

    Dim intMultiplicador As Integer

    Dim intResto As Integer

    ' se nao for um valor numerico sai da função
    If Not IsNumeric(strNumero) Then
        Calculo_DV11 = ""
        Exit Function
    End If

    'inicia o multiplicador
    'intMultiplicador = 9

    'pega cada caracter do numero a partir da direita
    For intContador = Len(strNumero) To 1 Step -1

        'extrai o caracter e multiplica prlo multiplicador
        intNumero = Val(Mid(strNumero, intContador, 1)) * intMultiplicador

        'soma o resultado para totalização
        intTotalNumero = intTotalNumero + intNumero
    
        'se o multiplicador for maior que 2 decrementa-o caso contrario atribuir valor padrao original
        intMultiplicador = IIf(intMultiplicador > 2, intMultiplicador - 1, 9)

    Next

    'calcula o resto da divisao do total por 11
    intResto = intTotalNumero Mod 11

    'verifica as exceções ( 0 -> DV=0    10 -> DV=X (para o BB) e retorna o DV
    Select Case intResto
        Case 0
            Calculo_DV11 = "0"
        Case 10
            Calculo_DV11 = "X"
        Case Else
            Calculo_DV11 = str(intResto)
    End Select

End Function

Function calculo_dv11base9(strNumero As String) As String
    'declara as variáveis
    Dim intContador As Integer
    Dim intNumero As Integer

    Dim intTotalNumero As Integer

    Dim intMultiplicador As Integer

    Dim intResto As Integer
    Dim intMultBase As Integer
    Dim intDV As Integer

    ' se nao for um valor numerico sai da função
    If Not IsNumeric(strNumero) Then
        calculo_dv11base9 = ""
        Exit Function
    End If

    'inicia o multiplicador
    intMultBase = 2
    intMultiplicador = 9
    
    Dim a As Integer

    'pega cada caracter do numero a partir da direita
    For intContador = Len(strNumero) To 1 Step -1

        'extrai o caracter e multiplica prlo multiplicador
        a = Mid(strNumero, intContador, 1)
        intNumero = Val(a) * intMultBase

        'soma o resultado para totalização
        intTotalNumero = intTotalNumero + intNumero
    
        'se o multiplicador for maior que 2 decrementa-o caso contrario atribuir valor padrao original
        intMultBase = IIf(intMultBase < 9, intMultBase + 1, 2)

    Next

    'calcula o resto da divisao do total por 11
    intResto = intTotalNumero Mod 11

    'verifica as exceções ( 0 -> DV=0    10 -> DV=X (para o BB) e retorna o DV
    intDV = 11 - intResto
    
    Select Case intDV
        Case 0
            calculo_dv11base9 = "1"
        Case Is > 9
            calculo_dv11base9 = "1"
        Case Else
            calculo_dv11base9 = str(intDV)
    End Select

End Function

Function calculo_dv11base7(strNumero As String) As String
    'declara as variáveis
    Dim intContador As Integer
    Dim intNumero As Integer

    Dim intTotalNumero As Integer

    Dim intMultiplicador As Integer

    Dim intResto As Integer
    Dim intMultBase As Integer
    Dim intDV As Integer

    ' se nao for um valor numerico sai da função
    If Not IsNumeric(strNumero) Then
        calculo_dv11base7 = ""
        Exit Function
    End If

    'inicia o multiplicador
    intMultBase = 2
    intMultiplicador = 7
    
    Dim a As Integer

    'pega cada caracter do numero a partir da direita
    For intContador = Len(strNumero) To 1 Step -1

        'extrai o caracter e multiplica prlo multiplicador
        a = Mid(strNumero, intContador, 1)
        intNumero = Val(a) * intMultBase

        'soma o resultado para totalização
        intTotalNumero = intTotalNumero + intNumero
    
        'se o multiplicador for maior que 2 decrementa-o caso contrario atribuir valor padrao original
        intMultBase = IIf(intMultBase < 7, intMultBase + 1, 2)

    Next

    'calcula o resto da divisao do total por 11
    intResto = intTotalNumero Mod 11

    'verifica as exceções ( 0 -> DV=0    10 -> DV=X (para o BB) e retorna o DV
    intDV = 11 - intResto
    
    Select Case intDV
        Case 11
            calculo_dv11base7 = "0"
        Case 10
            calculo_dv11base7 = "P"
        Case Else
            calculo_dv11base7 = str(intDV)
    End Select
    End Function
Public Sub BoletoBancario(Id As Long, Optional Visualizar = True)
    Dim cBanco As String
    cBanco = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero
    Select Case cBanco
        Case "001"
            BoletoBancario_001 (Id)
        Case "237"
            BoletoBancario_237 (Id)
        Case "356"
            BoletoBancario_356 (Id)
        Case Else
            MsgBox "Codificação para o banco " & cBanco & " não encontrada. Favor avisar ao suporte!", vbInformation, "Aviso"
            Exit Sub
    End Select
    'Comando de impressao do boleto
    ImprBoletoBancario Id, Visualizar
End Sub
'
Public Sub BoletoBancario_001(Id As Long)


'Incusao da API 26/02/25
API_BBCobranca Id
'===========================

    '#######################################################################################
    '### Banco do Brasil
    '#######################################################################################
    Dim NossoNumero     As String
    Dim LinhaDigitavel  As String
    Dim CodigoBarras    As String
    
    Dim agencia         As String
    Dim conta           As String
    Dim fator           As Integer
    Dim Valor           As String
    
    Dim seqI            As String
    Dim seqII           As String
    Dim seqIII          As String
    Dim seqIV           As String
    Dim sequencia       As String
    
    Dim dvLinhaDig      As String
    Dim dvCB            As String
    Dim dvNN            As String
    
    Dim dv1, dv2, dv3   As Integer
    '########################################################################################################
    '### Montagem de NOSSO NUMERO
    '########################################################################################################
    '# Multiplicador base 9
    Dim NN1, NN2 As String
    NN1 = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).convenio
    If Len(Trim(Id)) > 5 Then
            NN2 = Right(Id, 5)
        Else
            NN2 = Left(String(5, "0"), 5 - Len(Trim(Id))) & Trim(Id)
    End If
    NossoNumero = NN1 & NN2
    'dvNN = Trim(Calculo_DV11(NossoNumero))
    dvNN = Trim(calculo_dv11base9(NossoNumero))
    '########################################################################################################
    
    agencia = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).agencia
    agencia = Left("0000", 4 - Len(agencia)) & agencia
    
    conta = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).conta
    conta = Mid(String(8, "0"), 1, Len(Trim(conta))) & conta
    fator = CalculoFator(PgDadosFinanceiroFatura(Id).Vencimento)
    
    
    Valor = IIf(Trim(PgDadosFinanceiroFatura(Id).vlCobrado) <> 0, PgDadosFinanceiroFatura(Id).vlCobrado, PgDadosFinanceiroFatura(Id).vlDuplicata)
    Valor = RS(ChkVal(Valor, 0, cDecMoeda))
    Valor = Left(String(10, "0"), 10 - Len(Valor)) & Valor
    
    '########################################################################################################
    '###  CODIGO DE BARRAS
    '########################################################################################################
    CodigoBarras = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & _
                   "9" & _
                   fator & _
                   Left(String(10, "0"), 10 - Len(Valor)) & Valor & _
                   NossoNumero & _
                   agencia & _
                   conta & _
                   pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).carteira
    
    dvCB = Trim(calculo_dv11base9(CodigoBarras))
    Select Case dvCB
        Case 0
            dvCB = 1
        Case 10
            dvCB = 1
        Case 11
            dvCB = 1
        Case "X"
            dvCB = 1
    End Select
    
    CodigoBarras = Left(CodigoBarras, 4) & dvCB & Mid(CodigoBarras, 5, Len(CodigoBarras))
    '########################################################################################################
    
    
    '******************************** Bloco I *************************************************************
    'seqI = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & _
                        "9" & _
                         Mid(NossoNumero, 1, 5)
                         
    seqI = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & _
                        "9" & _
                         Mid(CodigoBarras, 20, 5)
    dv1 = Trim(Calculo_DV10(seqI))
    
    '******************************** Bloco II *************************************************************
    'seqII = Mid(NossoNumero, 6, Len(NossoNumero)) & Agencia
    seqII = Mid(CodigoBarras, 25, 10)
    dv2 = Trim(Calculo_DV10(seqII))
    '******************************** Bloco III *************************************************************
    'seqIII = Left(String(8, "0"), 8 - Len(Conta)) & Conta & pgDadosConta(PgDadosFinanceiroFatura(Id).IdConta).Carteira
    seqIII = Mid(CodigoBarras, 35, 10)
    
    dv3 = Trim(Calculo_DV10(seqIII))
    
    '******************************** Bloco IV *************************************************************
    seqIV = fator & Left(String(10, "0"), 10 - Len(Valor)) & Valor
    
    '*******************************************************************************************************
    
    sequencia = seqI & seqII & seqIII
            
    
    
    
    sequencia = seqI & dv1 & seqII & dv2 & seqIII & dv3 & dvCB & seqIV
    LinhaDigitavel = sequencia 'Formatar_Linha_Digitavel(sequencia)
    
    
    NossoNumero = NossoNumero & dvNN
    
    grvDadosBoleto Id, NossoNumero, LinhaDigitavel, CodigoBarras
    
    
    'ImprBoletoBancario Id ', NossoNumero, LinhaDigitavel, CodigoBarras
    
End Sub
Public Sub API_BBCobranca(faturaId As Long)
    Dim Id As Long
    Dim strJSON As String
    Dim producao As Boolean
   
    
    Dim convenio As String
    Dim carteira As String
    Dim carteiraVariacao As String
    Dim tipoConta As String
    
    
    
    
     
    Id = faturaId
    producao = False
    
    '-------------------------------------------------DADOS PARA API --------------------
     ' 0 = num // 1= str // 2= empty
     

    strJSON = "{" & vbCrLf
    
    If producao = True Then
            'Modulo Producao
            convenio = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).convenio
            carteira = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).carteira
            carteiraVariacao = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Variacao
            tipoConta = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Tipo
        Else
            
            'Modulo Homologacao
            convenio = "3128557"
            carteira = "17"
            carteiraVariacao = "35"
            tipoConta = "4"
    
    End If
    
    strJSON = strJSON & mJ("numeroConvenio", convenio, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroCarteira", carteira, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroVariacaoCarteira", CInt(carteiraVariacao), 0) & "," & vbCrLf
    strJSON = strJSON & mJ("codigoModalidade", CInt(tipoConta), 0) & "," & vbCrLf
    
    
    strJSON = strJSON & mJ("dataEmissao", Replace(PgDadosFinanceiroFatura(Id).emissao, "/", "."), 1) & "," & vbCrLf
    strJSON = strJSON & mJ("dataVencimento", Replace(PgDadosFinanceiroFatura(Id).Vencimento, "/", "."), 1) & "," & vbCrLf
    strJSON = strJSON & mJ("valorOriginal", PgDadosFinanceiroFatura(Id).vlCobrado, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("valorAbatimento", PgDadosFinanceiroFatura(Id).Deducoes, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("quantidadeDiasProtesto", PgDadosFinanceiroFatura(Id).DiasProtesto, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("quantidadeDiasNegativacao", "0", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("orgaoNegativador", "10", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("indicadorAceiteTituloVencido", "S", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroDiasLimiteRecebimento", PgDadosFinanceiroFatura(Id).DiasProtesto, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("codigoAceite", "A", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("codigoTipoTitulo", "2", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("descricaoTipoTitulo", "DM", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("indicadorPermissaoRecebimentoParcial", "N", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroTituloBeneficiario", PgDadosFinanceiroFatura(Id).NumFatura, 1) & "," & vbCrLf
    strJSON = strJSON & mJ("campoUtilizacaoBeneficiario", PgDadosFinanceiroFatura(Id).NumDuplicata, 1) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroTituloCliente", BoletoBancario_001_NossoNumero(Id), 1) & "," & vbCrLf
    strJSON = strJSON & mJ("mensagemBloquetoOcorrencia", PgDadosFinanceiroFatura(Id).Obs, 1) & ","
    
    'DESCONTO
    strJSON = strJSON & vbCrLf & _
    mJ("desconto", "", 2) & "{" & vbCrLf & _
    mJ("tipo", "0", 0) & "," & vbCrLf & _
    mJ("dataExpiracao", Replace(PgDadosFinanceiroFatura(Id).Vencimento, "/", "."), 1) & "," & vbCrLf & _
    mJ("porcentagem", "0", 0) & "," & vbCrLf & _
    mJ("valor", PgDadosFinanceiroFatura(Id).Deducoes, 0) & _
    vbCrLf & "},"
    
    strJSON = strJSON & vbCrLf & _
     mJ("segundoDesconto", "", 2) & "{" & vbCrLf & _
    mJ("dataExpiracao", "", 1) & "," & vbCrLf & _
    mJ("porcentagem", "0", 0) & "," & vbCrLf & _
    mJ("valor", "0", 0) & _
    vbCrLf & "},"
   
    strJSON = strJSON & vbCrLf & _
     mJ("terceiroDesconto", "", 2) & "{" & vbCrLf & _
    mJ("dataExpiracao", "", 1) & "," & vbCrLf & _
    mJ("porcentagem", "0", 0) & "," & vbCrLf & _
    mJ("valor", "0", 0) & _
    vbCrLf & "},"
    
    'Mora Diaria
    Dim vMD As String
    Dim vMulta As String
    vMD = ConvMoeda(cobCalcMora(PgDadosFinanceiroFatura(Id).vlDuplicata, 1, PgDadosFinanceiroFatura(Id).Juros, "D"))
    vMulta = PgDadosFinanceiroFatura(Id).Multa & "% ou " & ConvMoeda(cobCalcMulta(PgDadosFinanceiroFatura(Id).vlDuplicata, PgDadosFinanceiroFatura(Id).Multa, 1))
            
    
    'JUROS MORA
    strJSON = strJSON & vbCrLf & _
    strJSON = strJSON & mJ("jurosMora", "", 2) & "{" & vbCrLf
    strJSON = strJSON & mJ("tipo", "1", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("porcentagem", PgDadosFinanceiroFatura(Id).Juros, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("valor", vMD, 0) & vbCrLf
    strJSON = strJSON & "},"
    
    'MULTA
    strJSON = strJSON & vbCrLf & _
    mJ("multa", "", 2) & "{" & vbCrLf & _
    mJ("tipo", "1", 0) & "," & vbCrLf & _
    mJ("data", Replace(PgDadosFinanceiroFatura(Id).Vencimento, "/", "."), 1) & "," & vbCrLf & _
    mJ("porcentagem", PgDadosFinanceiroFatura(Id).Multa, 0) & "," & vbCrLf & _
    mJ("valor", vMulta, 0) & _
    vbCrLf & "},"
        
   'PAGADOR
    strJSON = strJSON & vbCrLf & _
     mJ("pagador", "", 2) & "{" & vbCrLf & _
    mJ("tipoInscricao", IIf(LCase(PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Pessoa) = "juridica", 2, 1), 0) & "," & vbCrLf & _
    mJ("numeroInscricao", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Doc, 1) & "," & vbCrLf & _
    mJ("nome", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Nome, 1) & "," & vbCrLf & _
    mJ("endereco", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Lgr, 1) & "," & vbCrLf & _
    mJ("cep", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).CEP, 1) & "," & vbCrLf & _
    mJ("cidade", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Mun, 1) & "," & vbCrLf & _
    mJ("bairro", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Bairro, 1) & "," & vbCrLf & _
    mJ("uf", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).uf, 1) & "," & vbCrLf & _
    mJ("telefone", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Fone, 1) & "," & vbCrLf & _
    mJ("email", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).emailfin, 1) & vbCrLf & _
    " },"
    
    'BENEFICIARIO
    strJSON = strJSON & vbCrLf & _
    mJ("beneficiarioFinal", "", 2) & "{" & vbCrLf & _
    mJ("tipoInscricao", 2, 0) & "," & vbCrLf & _
    mJ("numeroInscricao", PgDadosEmpresa(ID_Empresa).CNPJ, 0) & "," & vbCrLf & _
    mJ("nome", PgDadosEmpresa(ID_Empresa).Nome, 1) & vbCrLf & _
    "}," & vbCrLf & _
    mJ("indicadorPix", "N", 1)
    strJSON = strJSON & vbCrLf & "}"
'----------------------------------------------------------------------
End Sub
Private Function mJ(sField As String, sData As Variant, iType As Integer) As String
    'Monta uma linha para JSON
    '
    'iType
    '0 = number
    '1 = string
    If iType = 0 Then
            'Number
            mJ = """" & sField & """: " & sData
        ElseIf iType = 1 Then
            'String
            '
            mJ = """" & sField & """" & ": " & """" & sData & """"
            
        Else
            mJ = """" & sField & """" & ": "
    End If
End Function
Public Sub BoletoBancario_237(Id As Long)
    '#######################################################################################
    '### Banco do Brasdesco - 237
    '### Maio/2015
    '#######################################################################################
    Dim NossoNumero     As String
    Dim LinhaDigitavel  As String
    Dim CodigoBarras    As String
    
    Dim agencia         As String
    Dim conta           As String
    Dim fator           As Integer
    Dim Valor           As String
    Dim carteira        As String
    
    Dim seqI            As String
    Dim seqII           As String
    Dim seqIII          As String
    Dim seqIV           As String
    Dim sequencia       As String
    
    Dim dvLinhaDig      As String
    Dim dvCB            As String
    Dim dvNN            As String
    
    Dim dv1, dv2, dv3   As Integer
    '########################################################################################################
    '### Montagem de NOSSO NUMERO
    '########################################################################################################
    Dim NN1, NN2 As String
    NN1 = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).convenio
    carteira = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).carteira
    If Len(Trim(Id)) > 5 Then
            NN2 = Right(Id, 5)
        Else
            NN2 = Left(String(5, "0"), 5 - Len(Trim(Id))) & Trim(Id)
    End If
    NossoNumero = NN1 & NN2
    NossoNumero = Mid(String(11, "0"), 1, Len(Trim(NossoNumero)) + 1) & NossoNumero
    dvNN = Trim(calculo_dv11base7(carteira & NossoNumero))
    '########################################################################################################
    
    agencia = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).agencia
    agencia = Left("0000", 4 - Len(agencia)) & agencia
    
    conta = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).conta
    conta = Mid(String(8, "0"), 1, Len(Trim(conta))) & conta
    fator = CalculoFator(PgDadosFinanceiroFatura(Id).Vencimento)
    
    
    Valor = IIf(Trim(PgDadosFinanceiroFatura(Id).vlCobrado) <> 0, PgDadosFinanceiroFatura(Id).vlCobrado, PgDadosFinanceiroFatura(Id).vlDuplicata)
    Valor = RS(ChkVal(Valor, 0, cDecMoeda))
    Valor = Left(String(10, "0"), 10 - Len(Valor)) & Valor
    
    '########################################################################################################
    '###  CODIGO DE BARRAS
    '########################################################################################################
    Dim cd1 As String
    Dim campoLivre As String
    cd1 = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & _
                   "9" & _
                   "" & _
                   fator & _
                   Left(String(10, "0"), 10 - Len(Valor)) & Valor
    campoLivre = agencia & _
                carteira & _
                NossoNumero & _
                Right(conta, 7) & _
                "0"
    CodigoBarras = cd1 & campoLivre
    dvCB = Trim(calculo_dv11base9(CodigoBarras))
    
    CodigoBarras = Left(CodigoBarras, 4) & dvCB & Mid(CodigoBarras, 5, Len(CodigoBarras))
    '##############################################################################################
    '### LINHA DIGITAVEL
    
    '******************************** Bloco I *************************************************************
    'seqI = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & _
                        "9" & _
                         Mid(NossoNumero, 1, 5)
                         
    seqI = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & _
                        "9" & _
                         Mid(campoLivre, 1, 5)
    dv1 = Trim(Calculo_DV10(seqI))
    
    '******************************** Bloco II *************************************************************
    'seqII = Mid(NossoNumero, 6, Len(NossoNumero)) & _
            Agencia
    seqII = Mid(campoLivre, 6, 10)
    dv2 = Trim(Calculo_DV10(seqII))
    '******************************** Bloco III *************************************************************
    'seqIII = Left(String(8, "0"), 8 - Len(Conta)) & Conta & pgDadosConta(PgDadosFinanceiroFatura(Id).IdConta).Carteira
    seqIII = Mid(campoLivre, 16, 10)
    
    dv3 = Trim(Calculo_DV10(seqIII))
    
    '******************************** Bloco IV *************************************************************
    seqIV = fator & Left(String(10, "0"), 10 - Len(Valor)) & Valor
    
    '*******************************************************************************************************
    
    'sequencia = seqI & seqII & seqIII & seqIV
            
    
    
    
    sequencia = seqI & dv1 & seqII & dv2 & seqIII & dv3 & dvCB & seqIV
    LinhaDigitavel = sequencia 'Formatar_Linha_Digitavel(sequencia)
    
    
    NossoNumero = NossoNumero & dvNN
    
    grvDadosBoleto Id, NossoNumero, LinhaDigitavel, CodigoBarras
    
    
    'ImprBoletoBancario Id ', NossoNumero, LinhaDigitavel, CodigoBarras
    
End Sub


Private Sub BoletoBancario_356(Id As Long)
    MsgBox "Revisar pois os modulos mudarao!", vbCritical, "Aviso"
    '#######################################################################################
    '### Banco Real
    '#######################################################################################
    '
    Exit Sub
    
    Dim NossoNumero     As String
    Dim LinhaDigitavel  As String
    Dim CodigoBarras    As String
    
    Dim agencia         As String
    Dim conta           As String
    
    Dim seqI            As String
    Dim seqII           As String
    Dim seqIII          As String
    Dim sequencia       As String
    Dim dvLinhaDig      As String
    Dim dvCob           As Integer
                    
    NossoNumero = Calculo_NossoNumero(IIf(Trim(PgDadosFinanceiroFatura(Id).NossoNumero) = "", RS(PgDadosFinanceiroFatura(Id).NumDuplicata), PgDadosFinanceiroFatura(Id).NossoNumero))
            
    agencia = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).agencia
    agencia = Left("0000", 4 - Len(agencia)) & agencia
    conta = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).conta
    conta = Left("0000000", 7 - Len(conta)) & conta
            
    seqI = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & _
                        "9" & _
                        agencia & _
                        Left(conta, 1)

    dvCob = Trim(Calculo_DV10(Left(NossoNumero, Len(NossoNumero) - 1) & pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).agencia & pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).conta))
    seqII = Mid(conta, 2, Len(conta)) & _
            dvCob & _
            Left(NossoNumero, 3)
        
    seqIII = Mid(Left(NossoNumero, Len(NossoNumero) - 1), 3, Len(NossoNumero))
            
    sequencia = seqI & seqII & seqIII
            
    dvLinhaDig = Trim(Calculo_DV11(sequencia))
            
    'LinhaDigitavel = Formatar_Linha_Digitavel(sequencia, dvLinhaDig, CStr(PgDadosFinanceiroFatura(Id).Vencimento), CSng(PgDadosFinanceiroFatura(Id).vlDuplicata))
            
    CodigoBarras = Monta_CodBarras(pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero, _
                                                                    "9", _
                                                                    CSng(PgDadosFinanceiroFatura(Id).vlDuplicata), _
                                                                    PgDadosFinanceiroFatura(Id).Vencimento, _
                                                                    agencia, _
                                                                    conta, _
                                                                    Left(NossoNumero, Len(NossoNumero) - 1), _
                                                                    dvCob)
    
    grvDadosBoleto Id, NossoNumero, LinhaDigitavel, CodigoBarras
    
    'ImprBoletoBancario Id ', NossoNumero, LinhaDigitavel, CodigoBarras
    
End Sub

Private Sub grvDadosBoleto(idBol As Long, sNN As String, sLD As String, sCB As String)
    Dim vReg(10)    As Variant
    Dim cReg        As Integer
    
    cReg = 0
    vReg(cReg) = Array("NossoNumero", sNN, "S"): cReg = cReg + 1
    vReg(cReg) = Array("LinhaDigitavel", sLD, "S"): cReg = cReg + 1
    vReg(cReg) = Array("CodigoBarras", sCB, "S"): cReg = cReg + 1
    cReg = cReg - 1
    RegistroAlterar "FinanceiroContasPRCadastro", vReg, cReg, "id = " & idBol
End Sub
Public Function ClientePosicaoFinanceira(idCliente As Integer) As CliPosicaoFinanceira
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim Pagar   As String
    Dim Pago    As String
    
    Pagar = 0
    Pago = 0
    
    sSQL = "SELECT * FROM FinanceiroContasPRCadastro " & _
           "WHERE Tabela = 'Clientes' AND idSacado = " & idCliente
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                If IsNull(Rst.Fields("DataQuitacao")) = True Then
                        Pagar = Val(ChkVal(Pagar, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("VlCobrado"), 0, cDecMoeda))
                    Else
                        Pago = Val(ChkVal(Pago, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("VlCobrado"), 0, cDecMoeda))
                End If
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    
    ClientePosicaoFinanceira.Pagar = ChkVal(Pagar, 0, cDecMoeda)
    ClientePosicaoFinanceira.Pago = ChkVal(Pago, 0, cDecMoeda)
End Function
Public Function AtualizaCobranca(idDupl As Long, DtCalculo As String) As CalcTitulo
'#####################################################################
'### Atualiza a Duplicata para a Dt solicitada
'#####################################################################
    Dim DiasVencidos    As Integer
    Dim vMulta          As String
    Dim vMora           As String
    Dim vTotal          As String
    
    Dim pMulta          As String
    Dim pMora           As String
    Dim dVenc           As String
    Dim vDupl           As String
    Dim vAcresc         As String
    Dim vDesc           As String
    Dim vDed            As String
    Dim vCalcFin        As String
    If idDupl = 0 Then Exit Function
    
    vAcresc = PgDadosFinanceiroFatura(idDupl).Acrescimo
    vDesc = PgDadosFinanceiroFatura(idDupl).Abatimento
    vDed = PgDadosFinanceiroFatura(idDupl).Deducoes
    
    pMulta = PgDadosFinanceiroFatura(idDupl).Multa
    pMora = PgDadosFinanceiroFatura(idDupl).Juros
    dVenc = PgDadosFinanceiroFatura(idDupl).Vencimento
    vDupl = PgDadosFinanceiroFatura(idDupl).vlDuplicata
    
    '07.06.2017
    'verifica se a data cai no sabado ou domingo
    Dim dnVenc As Date
    
    dnVenc = dVenc
    
    Select Case DatePart("w", dVenc)
        Case 1 'Domingo
            'DOMINGO"
            dnVenc = dnVenc + 1
        Case 7 'Sabado
            'SABADO"
            dnVenc = dnVenc + 2
        Case Else
            'Me.Caption = ""
    End Select
    
    If DtCalculo > dnVenc Then
            'dVenc = DtCalculo
        Else
            dVenc = dnVenc
    End If
    
    
    
    
    
    
    '************** Dias Vencidos
    DiasVencidos = CDate(DtCalculo) - CDate(dVenc)
    DiasVencidos = IIf(DiasVencidos < 0, 0, DiasVencidos)
    
    If DiasVencidos <> "0" Then
            '************** Valor Multa
            'vMulta = Val(ChkVal(Val(ChkVal(vDupl, 0, cDecMoeda)) * Val(ChkVal(pMulta, 0, 3)), 0, cDecMoeda)) / 100
            'vMulta = ChkVal(vMulta, 0, cDecMoeda)
            vMulta = cobCalcMulta(vDupl, pMulta)
            '************** Valor Mora
            'vMora = Val(DiasVencidos) * Val(ChkVal(pMora, 0, 3))
            'vMora = Val(ChkVal(vMora, 0, 3)) * Val(ChkVal(vDupl, 0, cDecMoeda)) / 100
            'vMora = ChkVal(vMora, 0, cDecMoeda)
            vMora = cobCalcMora(vDupl, DiasVencidos, pMora, "T")
        Else
            vMulta = 0
            vMora = 0
    End If
    'Modificado em 16.10.13
    'vTotal = Val(ChkVal(vDupl, 0, cDecMoeda)) + Val(vMulta) + Val(vMora)
    vCalcFin = Val(vMulta) + Val(vMora)
    vCalcFin = Val(ChkVal(vCalcFin, 0, cDecMoeda)) + Val(ChkVal(vAcresc, 0, cDecMoeda))
    vCalcFin = Val(ChkVal(vCalcFin, 0, cDecMoeda)) - (Val(ChkVal(vDesc, 0, cDecMoeda)) + Val(ChkVal(vDed, 0, cDecMoeda)))
    
    vTotal = Val(ChkVal(vDupl, 0, cDecMoeda)) + Val(vCalcFin)
    vTotal = ChkVal(vTotal, 0, cDecMoeda)
    
    
    
    
    AtualizaCobranca.DiasVencidos = DiasVencidos
    AtualizaCobranca.vMora = vMora
    AtualizaCobranca.vMulta = vMulta
    AtualizaCobranca.vTotal = vTotal
    AtualizaCobranca.vCalcFin = vCalcFin
End Function
Public Sub cobrLancamentoAutomaticoContasFixas()
    '###################################################################################################
    '### 26/03/2012
    '### Objetivo: Lancar automaticamente mes a mes todas as despesas fixas
    '###################################################################################################
    On Error GoTo TrtErroCtaFixas
    Dim sSQL        As String
    Dim RstF        As Recordset 'Recordset das despesas Fixas
    Dim Rst         As Recordset 'Recordset das despesas
    Dim MesAtual    As String
    Dim AnoAtual    As String
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    Dim dVenc       As Date
    
    Dim vIni        As String
    Dim vFin        As String
    
    
    MesAtual = Format(Date, "MM")
    AnoAtual = Format(Date, "YYYY")
    
    vIni = "01/" & MesAtual & "/" & AnoAtual
    vFin = CalcData("01", 3, CDate(vIni))
    
    'Removido em 25/04/2012 - para efetuar o lancamento no 1o. dia do mes
    'sSQL = "SELECT * FROM FinanceiroContasPRFixa" & _
         " WHERE ID_Empresa=" & ID_Empresa & _
         " AND VencInicial <='" & Format(Date, "YYYY-MM-DD") & "'" & _
         " AND VencFinal >='" & Format(Date, "YYYY-MM-DD") & "'" & _
         " AND Meses LIKE '%" & MesAtual & "%'"
    
    sSQL = "SELECT * FROM FinanceiroContasPRFixa" & _
         " WHERE ID_Empresa=" & ID_Empresa & _
         " AND VencInicial <='" & Format(vIni, "YYYY-MM-DD") & "'" & _
         " AND VencFinal >='" & Format(vFin, "YYYY-MM-DD") & "'" & _
         " AND Meses LIKE '%" & MesAtual & "%'"
    Set RstF = RegistroBuscar(sSQL)
    If RstF Is Nothing Then
        'RstF.Close
        Exit Sub
    End If
    
     If RstF.BOF And RstF.EOF Then
            RstF.Close
            Exit Sub
        Else
            RstF.MoveFirst
    End If
    Do Until RstF.EOF
        'Calcular dt vencimento
        dVenc = CalcData(cNull(RstF.Fields("vencDia")), cNull(RstF.Fields("AntSabDom")), Date)
        'Verifica se a duplicata ja foi inclusa
        'sSQL = "SELECT * FROM FinanceiroContasPRCadastro" & _
            " WHERE FixoVariavel='F'" & _
            " AND Vencimento ='" & Format(dVenc, "YYYY-MM-DD") & "'" & _
            " AND idFixa = " & RstF.Fields("ID")
            sSQL = "SELECT * FROM FinanceiroContasPRCadastro" & _
            " WHERE FixoVariavel='F'" & _
            " AND Vencimento BETWEEN '" & Format(vIni, "YYYY-MM-DD") & "'" & _
            " AND '" & Format(vFin, "YYYY-MM-DD") & "'" & _
            " AND idFixa = " & RstF.Fields("ID")
        Set Rst = RegistroBuscar(sSQL)
        If Rst.BOF And Rst.EOF Then
                'Duplicata nao inclusa
                MovimentarContasPagarReceber cNull(RstF.Fields("ContaPR")), _
                                            CDate(dVenc), _
                                            cNull(RstF.Fields("nFatura")), _
                                            cNull(RstF.Fields("vFatura")), _
                                            cNull(RstF.Fields("Tabela")), _
                                            IIf(cNull(Trim(RstF.Fields("idSacado"))) = "", "0", cNull(RstF.Fields("idSacado"))), _
                                            cNull(RstF.Fields("Nome")), _
                                            cNull(RstF.Fields("CNPJ")), _
                                            "0", _
                                            IIf(Trim(cNull(RstF.Fields("CentroCusto"))) = "", "0", cNull(RstF.Fields("CentroCusto"))), _
                                            cNull(RstF.Fields("TpDocumento")), _
                                            cNull(RstF.Fields("PlanoContas")), _
                                            "", _
                                            "", _
                                            CDate(dVenc), _
                                            cNull(RstF.Fields("nFatura")), _
                                            "0", "0", "0", "0", _
                                            "0", "0", "0", _
                                            cNull(RstF.Fields("vFatura")), _
                                            cNull(RstF.Fields("Obs")), _
                                            "", _
                                            "F", _
                                            cNull(RstF.Fields("ID"))

                'cReg = 0
                'vReg(cReg) = Array("idFixa", , "N"): cReg = cReg + 1
                'vReg(cReg) = Array("ContaPR", , "S"): cReg = cReg + 1
                'vReg(cReg) = Array("NumFatura", , "S"): cReg = cReg + 1
                'vReg(cReg) = Array("vlFatura", , "S"): cReg = cReg + 1
                'vReg(cReg) = Array("CentroCusto", , "N"): cReg = cReg + 1
                'vReg(cReg) = Array("TpDocumento", , "N"): cReg = cReg + 1
                'vReg(cReg) = Array("PlanoContas", , "N"): cReg = cReg + 1
                'vReg(cReg) = Array("Tabela", , "S"): cReg = cReg + 1
                'vReg(cReg) = Array("idSacado", , "N"): cReg = cReg + 1
                'vReg(cReg) = Array("CNPJ", , "S"): cReg = cReg + 1
                'vReg(cReg) = Array("Nome", , "S"): cReg = cReg + 1
                'vReg(cReg) = Array("NumDuplicata", , "S"): cReg = cReg + 1
                'vReg(cReg) = Array("vlDuplicata", , "S"): cReg = cReg + 1
                'vReg(cReg) = Array("FixoVariavel", "F", "S"): cReg = cReg + 1
                'vReg(cReg) = Array("vlCobrado", cNull(RstF.Fields("vFatura")), "S"): cReg = cReg + 1
                'vReg(cReg) = Array("Emissao", , "D"): cReg = cReg + 1
                'vReg(cReg) = Array("Vencimento", , "D"): cReg = cReg + 1
                
                'vReg(cReg) = Array("Obs", , "S"): cReg = cReg + 1
                
                'cReg = cReg - 1
                'RegistroIncluir "FinanceiroContasPRCadastro", vReg, cReg
                
            Else
                'Duplicata ja inclusa
        End If
        Rst.Close
        RstF.MoveNext
    Loop
    RstF.Close
    Exit Sub
TrtErroCtaFixas:
    'MsgBox Err.Description, vbCritical, Err.Number
    RegLogDataBase 0, "Error n." & Err.Number, 0, "(cobrLancamentoAutomaticoContasFixas) " & Err.Description
End Sub
Public Sub cnab240(DtIni As Date, DtFin As Date, _
                    contaId As Integer, Optional lote As String)

    Dim fd              As String
    Dim line            As String
    Dim rgP(9000)       As String
    Dim rgQ(9000)       As String
    Dim c               As Integer
    Dim sQL             As String
    
    Dim Rst As Recordset
    If Len(Trim(lote)) = 0 Then
        lote = "100"
    End If
    lote = F("n", 4, lote)

    
    'Monta o nome do arquivo
    fd = App.Path & "\log\cnab240_" & lote & ".REM"
    If Dir(fd) <> "" Then
        'Exclui o arquivo caso exista
        ExcluirFile fd
    End If
    
    Dim tpPeriodo As String
    
    tpPeriodo = tpPeriodo & "Emissao >= '" & Format(DtIni, "YYYY-MM-DD") & _
            "' AND " & tpPeriodo & " Emissao <= '" & Format(DtFin, "YYYY-MM-DD") & "'"
    
    
    sQL = "SELECT * FROM financeirocontasprcadastro"
    sQL = sQL & " WHERE ID_Empresa = " & ID_Empresa & " AND " & tpPeriodo
    'sQL = sQL & " WHERE gerarcnab240 = " & lote 'lote deve estar vazio pela 1 vez
    sQL = sQL & " AND conta = " & contaId
    sQL = sQL & " AND nossonumero Is NOT  Null"
    Set Rst = RegistroBuscar(sQL)
    
    If Rst.BOF And Rst.EOF Then
        Rst.Close
        MsgBox "Nenhum registro encontrado"
        Exit Sub
    End If
    
    Rst.MoveFirst
   
    
     
    '*** ROTEIRO DE CRIACAO DO ARQUIVO ***
    '
    ' 1 - Registro Header do Arquivo
    ' 2 - Registro Header do Lote
    ' 3 -
    ' 4 - Registro Trailer do Lote
    ' 5 - Registro Trailer do Arquivo
    
    
    'Montando os registros do arquivo
    c = 0
    
    Do Until Rst.EOF
    If Len(cNull(Rst.Fields("id"))) = 0 Then
        MsgBox "ops!"
    End If
        rgP(c) = cnab240P(Rst.Fields("id"), lote)
        rgQ(c) = cnab240Q(Rst.Fields("id"), lote): c = c + 1
'       line = cnab240P(Rst.Fields("id"))
'       grvFile fd, line
'        'Desmarcar linha selecionada para gerar arquivo
'        vDados(cReg) = Array("gerarcnab240", 0, "N"): cReg = cReg + 1
'        cReg = cReg - 1
'        criterio = "id=" & rst.Fields("id")
'        RegistroAlterar "financeirocontasprcadastro", vDados, cReg, criterio
    
    
        Rst.MoveNext
    Loop
    Rst.Close
    
    
    '### Header do Arquivo ###
    grvFile fd, cnab240ArquivoHeader(contaId, lote)
    
    '### Header do Lote ###
    grvFile fd, cnab240LoteHeader(contaId, lote)
    
    Dim i As Integer
    '### Registro P ####
    For i = 0 To c - 1
        grvFile fd, rgP(i)
    Next
    
    '### Registro Q ###
    For i = 0 To c
        grvFile fd, rgQ(i)
    Next
    
    'Trailer do Lote
    grvFile fd, cnab240LoteTrailer(contaId, lote, c)
    
    'Trailer do Arquivo
    grvFile fd, cnab240ArquivoTrailer(contaId, lote, c)
    
    
    
    MsgBox "Arquivo: " & fd & " criado com sucesso!"
End Sub

Private Function cnab240ArquivoHeader(contaId As Integer, lote As String) As String
    'Layout CNAB 240 - v10.1
    'Dim fd As String 'file Destination
    Dim line As String
    
    'fd = arqDestino
   
    '***********************************************************
    '********* Registro Header de Arquivo              *********
    '***********************************************************
    line = ""
    line = line & F("n", 3, pgDadosBanco(pgDadosConta(contaId).banco).Numero)
    line = line & F("n", 4, lote) '02.0 - Lote
    line = line & "0"
    '----------------
    line = line & String(9, " ")
    '----------------
    line = line & "0" '05.0 - Tipo de inscricao
    line = line & F("n", 14, PgDadosEmpresa(ID_Empresa).CNPJ) '06.0 - num insc empresa
    line = line & F("a", 20, pgDadosConta(contaId).convenio) '07.0- convenio
    line = line & F("n", 5, pgDadosConta(contaId).agencia)
    line = line & F("n", 1, pgDadosConta(contaId).AgenciaDV)
    line = line & F("n", 12, pgDadosConta(contaId).conta)
    line = line & F("n", 1, pgDadosConta(contaId).ContaDV)
    Dim dvAgCc As String
    dvAgCc = Calculo_DV11(pgDadosConta(contaId).agencia & _
                          pgDadosConta(contaId).conta)
    line = line & F("n", 1, dvAgCc)
    line = line & F("a", 30, PgDadosEmpresa(ID_Empresa).Nome) '13.0 - nome empresa
    line = line & F("a", 30, pgDadosBanco(pgDadosConta(contaId).banco).Nome)
    '----------------
    line = line & String(10, " ") '15.0 - cnab
    
    
    line = line & "1" '16.0 - Codigo remessa arquivo
    line = line & Format(Date, "ddmmYYYY")
    line = line & Format(Time, "hhmmss")
    line = line & F("n", 6, "0") '19.0 - NSA numero de sequencia do arquivo
    
    line = line & "089" '20.0 - layout do arquivo
    line = line & F("n", 5, "00000") '21.0 - Densidade da gravacao do arquivo
    
    line = line & String(20, " ") '22.0 - Reservado ao banco
    line = line & String(20, " ") '23.0 - Reservado a Empresa
    line = line & String(29, " ") '24.0 - Reservado ao FEBRABAN
    
    'grvFile fd, line
    cnab240ArquivoHeader = line
    

End Function
Private Function cnab240P(faturaId As Long, lote As String) As String
    ' RJ, 30.08.2016
    ' Autor: Leonardo Aquino
    '
    ' Funcao para criar arquivo exportação
    ' cnab240 em conformidade com Layout
    ' padrão Febraban 240 posicoes V10.1
    '
    Dim sQL As String
    
    Dim Rst As Recordset
    
    Dim vDados(10) As Variant
    Dim cReg As Integer
    Dim i As Integer
    Dim criterio As String
    Dim fd As String
    'Dim lote As Integer
    Dim arqDestino As String
    
    'Seleciona as boletas que vao gerar arquivo de remessa
   
        
    
    'Dim fd As String 'file Destination
    Dim line As String
    
   
    
   'fd = arqDestino
   
    '***********************************************************
    '*** Registro Detalhe - Segmento P (Obrigatorio Remessa) ***
    '***********************************************************
    line = ""
    line = line & F("n", 3, pgDadosBanco(PgDadosFinanceiroFatura(faturaId).IdBanco).Numero)
    line = line & F("n", 4, lote)
    line = line & "3"
    line = line & F("n", 5, "1")
    line = line & "P"
    line = line & F("a", 1, " ")
    line = line & F("n", 2, "0")
    line = line & F("n", 5, pgDadosConta(PgDadosFinanceiroFatura(faturaId).idConta).agencia)
    line = line & F("a", 1, pgDadosConta(PgDadosFinanceiroFatura(faturaId).idConta).AgenciaDV)
    line = line & F("n", 12, pgDadosConta(PgDadosFinanceiroFatura(faturaId).idConta).conta)
    line = line & F("a", 1, pgDadosConta(PgDadosFinanceiroFatura(faturaId).idConta).ContaDV)
    line = line & F("a", 1, " ") '12.3P - Digito verificador agencia/conta
    
    
    '
    'BANCO DO BRASIL - Nesse campo pede o id do titulo sem o convenio que compoe o N/Numero
    '
    'line = line & F("a", 20, PgDadosFinanceiroFatura(faturaID).NossoNumero) '13.3 - id titulo no banco
    line = line & F("a", 20, Mid(PgDadosFinanceiroFatura(faturaId).NossoNumero, 1, Len(PgDadosFinanceiroFatura(faturaId).NossoNumero) - 1)) '13.3 - id titulo no banco
    
    
    'Se o banco for BB informar codigo de acordo com as particularidades do banco
    Dim carteira As Integer
    If pgDadosBanco(PgDadosFinanceiroFatura(faturaId).IdBanco).Numero = "001" Then
            carteira = 7
        Else
            carteira = pgDadosConta(PgDadosFinanceiroFatura(faturaId).idConta).carteira
    End If
    line = line & F("n", 1, carteira)
    line = line & F("n", 1, "0") '15.3P - Cadastramento
    line = line & F("a", 1, " ") '16.3P - Tipo documento (DM)
    line = line & F("n", 1, "0") '17.3P - Emissao boleto de pagamento
    line = line & F("a", 1, " ") '18.3P - identificacao distribuicao
    line = line & F("a", 15, PgDadosFinanceiroFatura(faturaId).NumDuplicata)
    line = line & F("d", 8, PgDadosFinanceiroFatura(faturaId).Vencimento)
    line = line & F("n", 15, ChkVal(PgDadosFinanceiroFatura(faturaId).vlCobrado, 0, 2)) '21.3P - vl tit
    line = line & F("n", 5, pgDadosConta(PgDadosFinanceiroFatura(faturaId).idConta).agencia)
    line = line & F("a", 1, pgDadosConta(PgDadosFinanceiroFatura(faturaId).idConta).AgenciaDV)
    line = line & F("n", 2, "0") '24.3P - Especie
    line = line & "N" '25.3 - aceite
    line = line & F("d", 8, PgDadosFinanceiroFatura(faturaId).emissao)
    line = line & F("n", 1, "0") '27.3P -Cod juros mora
    line = line & F("d", 8, PgDadosFinanceiroFatura(faturaId).Vencimento)
    line = line & F("n", 15, ChkVal(PgDadosFinanceiroFatura(faturaId).Juros, 0, 2)) '29.3
    line = line & F("n", 1, "0") '30.3P -Cod desconto
    line = line & F("n", 8, "0") '31.3p- data descontoPgDadosFinanceiroFatura(faturaID).Vencimento
    line = line & F("n", 15, ChkVal(PgDadosFinanceiroFatura(faturaId).Deducoes, 0, 2)) '32.3
    line = line & F("n", 15, "0.00") '33.3P - vl IOF
    line = line & F("n", 15, ChkVal(PgDadosFinanceiroFatura(faturaId).Abatimento, 0, 2)) '34.3
    line = line & F("a", 25, PgDadosFinanceiroFatura(faturaId).NumFatura) '35.3P - Uso empresa beneficiaria
    line = line & F("n", 1, "0") '36.3P - Cod protesto
    line = line & F("n", 2, PgDadosFinanceiroFatura(faturaId).DiasProtesto)
    line = line & F("n", 1, "0") '38.3P - Cod baixa devolucao
    line = line & F("a", 3, " ") '39.3P - num dias baixa devolucao
    line = line & F("n", 2, "01") '40.3P - Cod moeda
    line = line & F("n", 10, pgDadosConta(PgDadosFinanceiroFatura(faturaId).idConta).Contrato)
    line = line & F("a", 1, " ") '42.3P - uso livre
    
    cnab240P = line
    'grvFile fd, line
    '*** FIM Seguimento P ***
    
   
    
    
    
    '***********************************************************
    '*** Registro Detalhe - Segmento R (Obrigatorio Remessa) ***
    '***********************************************************
    

End Function
Private Function cnab240LoteTrailer(contaId As Integer, lote As String, tReg As Integer) As String
'#### REGISTRO TRAILER DO LOTE ####
    Dim line    As String
    Dim vTotal  As String
    
    line = ""
    line = line & F("n", 3, pgDadosBanco(pgDadosConta(contaId).banco).Numero)
    line = line & F("n", 4, lote) 'Lote
    line = line & "5" '03.5 - tipo de registro
    line = line & String(9, " ")
    
    line = line & F("n", 6, tReg)
    
    vTotal = 0
    line = line & F("n", 16, ChkVal(vTotal, 0, 2)) '06.5 - valor total
    line = line & F("n", 13, ChkVal(vTotal, 0, 5)) '07.5 - quant de moeda
    line = line & String(6, "0")
    line = line & String(165, " ")
    line = line & String(10, " ")
    
    cnab240LoteTrailer = line


End Function
Private Function cnab240ArquivoTrailer(contaId As Integer, lote As String, tReg As Integer) As String
'#### REGISTRO TRAILER DO ARQUIVO ####
    Dim line    As String
    Dim vTotal As String
    
    line = ""
    line = line & F("n", 3, pgDadosBanco(pgDadosConta(contaId).banco).Numero)
    line = line & F("n", 4, lote) 'Lote
    line = line & "9"
    line = line & String(9, " ") '04.9 - cnab
    
    line = line & F("n", 6, "1")
    
    vTotal = 0
    line = line & F("n", 6, "1") '05.9 - quantidade de lotes no arquivo
    line = line & F("n", 6, tReg) '06.9 - qtd de reg do arq
    line = line & String(6, "1") '07.9 - qtd contas p/ conc
    line = line & String(205, " ")
    
    
    cnab240ArquivoTrailer = line

End Function
