VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BBCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Function GerarBoletoBB(Convenio As String, _
                        carteira As String, _
                        carteiraVariacao As String, _
                        tipoConta As String, _
                        dataEmissao As String, _
                        DataVencimento As String, _
                        nFatura As String, _
                        nDuplicata As String, _
                        Valor As String, _
                        vDeducao As String, _
                        vMulta As String, _
                        vJuros As String, _
                        DiasProtesto As String, _
                        Sacado As String, _
                        cnpjBeneficiario As String, _
                        nomeBeneficiario As String, _
                        NossoNumero As String, _
                        smsg As String) As String
    'Dim Id As Long
    Dim strJSON As String
  
    '-------------------------------------------------DADOS PARA API --------------------
     ' mJ(string , 0 = num // 1= str // 2= empty)
 
            
            
    
    strJSON = "{" & vbCrLf
    strJSON = strJSON & mJ("numeroConvenio", Convenio, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroCarteira", carteira, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroVariacaoCarteira", CInt(carteiraVariacao), 0) & "," & vbCrLf
    strJSON = strJSON & mJ("codigoModalidade", CInt(tipoConta), 0) & "," & vbCrLf
    
    
    strJSON = strJSON & mJ("dataEmissao", Replace(dataEmissao, "/", "."), 1) & "," & vbCrLf
    strJSON = strJSON & mJ("dataVencimento", Replace(DataVencimento, "/", "."), 1) & "," & vbCrLf
    strJSON = strJSON & mJ("valorOriginal", Valor, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("valorAbatimento", vDeducao, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("quantidadeDiasProtesto", DiasProtesto, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("quantidadeDiasNegativacao", "0", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("orgaoNegativador", "10", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("indicadorAceiteTituloVencido", "S", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroDiasLimiteRecebimento", DiasProtesto, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("codigoAceite", "A", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("codigoTipoTitulo", "2", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("descricaoTipoTitulo", "DM", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("indicadorPermissaoRecebimentoParcial", "N", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroTituloBeneficiario", nFatura, 1) & "," & vbCrLf
    strJSON = strJSON & mJ("campoUtilizacaoBeneficiario", Replace(nDuplicata, "/", "-"), 1) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroTituloCliente", NossoNumero, 1) & "," & vbCrLf
    strJSON = strJSON & mJ("mensagemBloquetoOcorrencia", smsg, 1) & ","
    
    'DESCONTO
    Dim tmpData As String
    Dim tmpValor As String
    tmpData = Replace(CDate(DataVencimento) + 1, "/", ".")
    tmpValor = vDeducao
    If tmpValor = 0 Then tmpData = ""
    
    strJSON = strJSON & vbCrLf
    strJSON = strJSON & mJ("desconto", "", 2) & "{" & vbCrLf
    strJSON = strJSON & mJ("tipo", "0", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("dataExpiracao", tmpData, 1) & "," & vbCrLf
    strJSON = strJSON & mJ("porcentagem", "0", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("valor", tmpValor, 0)
    strJSON = strJSON & vbCrLf & "},"
    
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
   
    
    'JUROS MORA
    strJSON = strJSON & vbCrLf
    strJSON = strJSON & mJ("jurosMora", "", 2) & "{" & vbCrLf
    strJSON = strJSON & mJ("tipo", "1", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("porcentagem", 0, 0) & "," & vbCrLf
    'strJSON = strJSON & mJ("porcentagem", PgDadosFinanceiroFatura(Id).Juros, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("valor", vJuros, 0) & vbCrLf
    strJSON = strJSON & vbCrLf & "},"
    
    'MULTA
    'tmpData = Replace(CDate(PgDadosFinanceiroFatura(Id).Vencimento) + 1, "/", ".")
    'tmpValor = vMulta
    'If tmpValor = 0 Then tmpData = ""
    
    strJSON = strJSON & vbCrLf
    strJSON = strJSON & mJ("multa", "", 2) & "{" & vbCrLf
    strJSON = strJSON & mJ("tipo", IIf(vMulta = 0, 0, "1"), 0) & "," & vbCrLf
    strJSON = strJSON & mJ("data", tmpData, 1) & "," & vbCrLf
    strJSON = strJSON & mJ("porcentagem", 0, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("valor", vMulta, 0)
    strJSON = strJSON & vbCrLf & "},"
        
   'PAGADOR
   'Dim Sacado As String
   'Sacado = pgDadosSacado(IDSacado)
   strJSON = strJSON & vbCrLf & Sacado
    
    'BENEFICIARIO
    strJSON = strJSON & vbCrLf & _
    mJ("beneficiarioFinal", "", 2) & "{" & vbCrLf & _
    mJ("tipoInscricao", 2, 0) & "," & vbCrLf & _
    mJ("numeroInscricao", cnpjBeneficiario, 0) & "," & vbCrLf & _
    mJ("nome", nomeBeneficiario, 1) & vbCrLf & _
    "}," & vbCrLf & _
    mJ("indicadorPix", "N", 1)
    strJSON = strJSON & vbCrLf & "}"
    
    GerarBoletoBB = strJSON
'----------------------------------------------------------------------
    'Debug.Print "Boleto JSON: " & strJSON
    
End Function

Public Function GerarNossoNumero(Convenio As String, Id As Long) As String

'########################################################################################################
    '### Montagem de NOSSO NUMERO 20 CARACTERES
    '### 28.02.25
    '########################################################################################################
    
    
    Dim NN1, NN2 As String
    'Convenio = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Convenio
    NN1 = Convenio
    Select Case Len(Trim(Convenio))
        Case Is <= 6
            'Dim NN1, NN2 As String
            
            If Len(Trim(Id)) > 5 Then
                    NN2 = Right(Id, 5)
                Else
                    NN2 = Left(String(5, "0"), 5 - Len(Trim(Id))) & Trim(Id)
            End If
            GerarNossoNumero = NN1 & NN2 & Trim(calculo_dv11base9(NN1 & NN2))
            
    
        Case 7
            NN1 = Left(String(7, "0"), 7 - Len(Trim(NN1))) & Trim(NN1)
            If Len(Trim(Id)) > 10 Then
                    NN2 = Right(Id, 10)
                Else
                    NN2 = Left(String(10, "0"), 10 - Len(Trim(Id))) & Trim(Id)
            End If
            GerarNossoNumero = "000" & NN1 & NN2
    End Select
    
 'Debug.Print "GerarNossoNumero: " & GerarNossoNumero
    

End Function

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


Public Function jsonSacado(scTpInsc As String, scCNPJ As String, scNome As String, scLgr As String, scCEP As String, scMun As String, scBairro As String, scUF As String, scFone As String, scEmail As String) As String
   
    Dim strPagador As String
   'PAGADOR
    strPagador = vbCrLf & _
    mJ("pagador", "", 2) & "{" & vbCrLf & _
    mJ("tipoInscricao", scTpInsc, 0) & "," & vbCrLf & _
    mJ("numeroInscricao", scCNPJ, 1) & "," & vbCrLf & _
    mJ("nome", scNome, 1) & "," & vbCrLf & _
    mJ("endereco", scLgr, 1) & "," & vbCrLf & _
    mJ("cep", scCEP, 1) & "," & vbCrLf & _
    mJ("cidade", scMun, 1) & "," & vbCrLf & _
    mJ("bairro", scBairro, 1) & "," & vbCrLf & _
    mJ("uf", scUF, 1) & "," & vbCrLf & _
    mJ("telefone", scFone, 1) & "," & vbCrLf & _
    mJ("email", cemail, 1) & vbCrLf & _
    " },"
    
    jsonSacado = strPagador
End Function


Public Function GerarCodigoBarrasBB( _
    ByVal banco As String, _
    ByVal moeda As String, _
    ByVal NossoNumero As String, _
    ByVal Valor As String, _
    ByVal Vencimento As String, _
    ByVal carteira As String _
) As String

    Dim CodigoBarras As String
    Dim FatorVencimento As String
    Dim DV As Integer
    
    FatorVencimento = CalcularFatorVencimento(CDate(Vencimento))
    Valor = Replace(Valor, ".", "")
    Valor = Left(String(10, "0"), 10 - Len(Trim(Valor))) & Valor

    ' 1. Monta o código de barras (sem o dígito verificador)
    CodigoBarras = banco & moeda & FatorVencimento & Valor & "000" & NossoNumero & carteira

    ' 2. Calcula o dígito verificador (Módulo 11)
    DV = Modulo11(CodigoBarras)

    ' 3. Insere o dígito verificador no código de barras
    CodigoBarras = banco & moeda & DV & FatorVencimento & Valor & "000" & NossoNumero & carteira

    ' 4. Retorna o resultado
    GerarCodigoBarrasBB = CodigoBarras

End Function

Private Function CalcularFatorVencimento(DataVencimento As Date) As Long
    Dim DataBase As Date
    Dim fator As Long

    
    DataBase = DateSerial(1997, 10, 7)

    
    fator = DateDiff("d", DataBase, DataVencimento)
     If fator > 9999 Then fator = fator - 9000
    ' Retorna o fator de vencimento
    CalcularFatorVencimento = fator
End Function
Public Function GerarLinhaDigitavelBB( _
    ByVal banco As String, _
    ByVal moeda As String, _
    ByVal Convenio As String, _
    ByVal carteira As String, _
    ByVal NossoNumero As String, _
    ByVal Valor As String, _
    ByVal Vencimento As String _
) As String

    Dim FatorVencimento As String
    Dim CodigoBarras As String
    Dim Campo1 As String
    Dim Campo2 As String
    Dim Campo3 As String
    Dim Campo4 As String
    Dim Campo5 As String
    Dim DV As Integer
    Dim DV1 As Integer
    Dim DV2 As Integer
    Dim DV3 As Integer
    Dim DV4 As Integer
    
    FatorVencimento = CalcularFatorVencimento(CDate(Vencimento))
    Valor = Replace(Valor, ".", "")
    Valor = Left(String(10, "0"), 10 - Len(Trim(Valor))) & Valor

    ' 1. Monta o código de barras
    CodigoBarras = GerarCodigoBarrasBB( _
                                        banco:=banco, _
                                        moeda:=moeda, _
                                        carteira:=carteira, _
                                        NossoNumero:=NossoNumero, _
                                        Valor:=Valor, _
                                        Vencimento:=Vencimento _
                                    )
                                    
    ' 4. Divide o código de barras em campos
    Campo1 = banco & moeda & Mid(CodigoBarras, 20, 5)
    Campo2 = Mid(CodigoBarras, 25, 10)
    Campo3 = Mid(CodigoBarras, 35, 10)
    Campo4 = Mid(CodigoBarras, 1, 4) & Mid(CodigoBarras, 6, 44) ' & Mid(CodigoBarras, 6, 3) & Mid(CodigoBarras, 7, 1)
    Campo5 = FatorVencimento & Valor '(CodigoBarras, 8, 10)

    ' 5. Calcula os dígitos verificadores de cada campo
    DV1 = Modulo10(Campo1)
    DV2 = Modulo10(Campo2)
    DV3 = Modulo10(Campo3)
    DV4 = Modulo11(Campo4)

    ' 6. Monta a linha digitável
    GerarLinhaDigitavelBB = Campo1 & DV1 & Campo2 & DV2 & Campo3 & DV3 & DV4 & Campo5
    'GerarLinhaDigitavelBB = Campo1 & DV1 & "" & Campo2 & DV2 & "" & Campo3 & DV3 & "" & Campo4 & DV4 & "" & Campo5

End Function
Private Function Modulo11(ByVal Codigo As String) As Integer
    Dim Soma As Integer
    Dim Peso As Integer
    Dim Digito As Integer
    Dim i As Integer

    Peso = 2
    Soma = 0

    For i = Len(Codigo) To 1 Step -1
        Soma = Soma + (Mid(Codigo, i, 1) * Peso)
        Peso = Peso + 1
        If Peso > 9 Then Peso = 2
    Next i

    Digito = 11 - (Soma Mod 11)

    If Digito = 0 Or Digito = 10 Or Digito = 11 Then
        Digito = 1
    End If

    Modulo11 = Digito
End Function



 Private Function Modulo10(ByVal Numero As String) As Integer

  Dim Multiplicador As Integer
    Dim Soma As Integer
    Dim Produto As Integer
    Dim DigitoVerificador As Integer
    Dim i As Integer

    Multiplicador = 2
    Soma = 0

    ' b) Multiplicadores começam com 2, alternando 1 e 2 (da direita para a esquerda)
    For i = Len(Numero) To 1 Step -1
        ' c) Multiplicar cada algarismo pelo seu respectivo peso
        Produto = Val(Mid(Numero, i, 1)) * Multiplicador

        ' d) Se o resultado for maior que 9, somar os algarismos do produto
        If Produto > 9 Then
            Produto = (Produto Mod 10) + (Produto \ 10)
        End If

        Soma = Soma + Produto

        ' Alternar multiplicador
        If Multiplicador = 2 Then
            Multiplicador = 1
        Else
            Multiplicador = 2
        End If
    Next i

    ' e) Subtrair o total da dezena imediatamente superior
    DigitoVerificador = ((Soma \ 10) + 1) * 10 - Soma

    ' g) Se o resultado for 10, o dígito verificador é 0
    If DigitoVerificador = 10 Then
        DigitoVerificador = 0
    End If

    ' f) Retorna o dígito verificador
    Modulo10 = DigitoVerificador
    
End Function
