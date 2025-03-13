Attribute VB_Name = "Modulo_CobrancaBB"
Public Sub mockTestBBCob()
   'Modulo Homologacao
   
            Convenio = "3128557"
            Carteira = "17"
            carteiraVariacao = "35"
            tipoConta = "4"
            cnpjPagador = "74910037000193"
            cnpjBeneficiario = "98959112000179"
            nomeBeneficiario = "LIVRARIA CUNHA DA CUNHA"
            
            'Convenio = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Convenio
            'Carteira = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Carteira
            'carteiraVariacao = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Variacao
            'tipoConta = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Tipo
            
            'cnpjPagador = PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Doc
            'cnpjBeneficiario = PgDadosEmpresa(ID_Empresa).CNPJ
            'nomeBeneficiario = PgDadosEmpresa(ID_Empresa).Nome
            'dataEmissao=PgDadosFinanceiroFatura(Id).emissao
            'dataVencimento = PgDadosFinanceiroFatura(Id).Vencimento
            'valor=PgDadosFinanceiroFatura(Id).vlCobrado
            'vDeducoes=PgDadosFinanceiroFatura(Id).Deducoes
            'diasProtesto=PgDadosFinanceiroFatura(Id).DiasProtesto
            'numFatura=PgDadosFinanceiroFatura(Id).NumFatura
            'numDuplicata= PgDadosFinanceiroFatura(Id).NumDuplicata
            'sMsg = PgDadosFinanceiroFatura(Id).Obs
            'pJuros=PgDadosFinanceiroFatura(Id).Juros
            'vMulta=PgDadosFinanceiroFatura(Id).Multa,
            'tpPessoa= IIf(LCase(PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Pessoa) = "fisica", 1, 2)
            
            
            
            
'--------JSON
'                    {
'        "numeroConvenio": 3128557,
'        "numeroCarteira": 17,
'        "numeroVariacaoCarteira": 35,
'        "codigoModalidade": 4,
'        "dataEmissao": "13.03.2025",
'        "dataVencimento": "20.03.2025",
'        "valorOriginal": 100.00,
'        "valorAbatimento": 0.00,
'        "quantidadeDiasProtesto": 5,
'        "quantidadeDiasNegativacao": 0,
'        "orgaoNegativador": 10,
'        "indicadorAceiteTituloVencido": "S",
'        "numeroDiasLimiteRecebimento": 5,
'        "codigoAceite": "A",
'        "codigoTipoTitulo": 2,
'        "descricaoTipoTitulo": "DM",
'        "indicadorPermissaoRecebimentoParcial": "N",
'        "numeroTituloBeneficiario": "FAT-2223",
'        "campoUtilizacaoBeneficiario": "FAT-2223-1-1",
'        "numeroTituloCliente": "00031285570000043986",
'        "mensagemBloquetoOcorrencia": "",
'        "desconto": {
'        "tipo": 0,
'        "dataExpiracao": "",
'        "porcentagem": 0,
'        "valor": 0.00
'        },
'        "segundoDesconto": {
'        "dataExpiracao": "",
'        "porcentagem": 0,
'        "valor": 0
'        },
'        "terceiroDesconto": {
'        "dataExpiracao": "",
'        "porcentagem": 0,
'        "valor": 0
'        },
'        "jurosMora": {
'        "tipo": 1,
'        "porcentagem": 0,
'        "valor": 0.20
'
'        },
'        "multa": {
'        "tipo": 0,
'        "data": "",
'        "porcentagem": 0,
'        "valor": 0.00
'        },
'        "pagador": {
'        "tipoInscricao": 2,
'        "numeroInscricao": "74910037000193",
'        "nome": "187 CENTRAL CARIOCA DE PECAS LTDA-EPP",
'        "endereco": "RUA DE SANTANA",
'        "cep": "20230260",
'        "cidade": "RIO DE JANEIRO",
'        "bairro": "CENTRO",
'        "uf": "RJ",
'        "telefone": "22219755",
'        "email": ""
'         },
'        "beneficiarioFinal": {
'        "tipoInscricao": 2,
'        "numeroInscricao": 98959112000179,
'        "nome": "LIVRARIA CUNHA DA CUNHA"
'        },
'        "indicadorPix": "N"
'        }
'
        
'--------------RETORNO

'        {
'        "beneficiario": {
'            "agencia": 452,
'            "contaCorrente": 123873,
'            "tipoEndereco": 1,
'            "logradouro": "ST AUXILIAR DE GARAGENS RUA 9 LOTE 10",
'            "bairro": "TAGUATINGA NORTE",
'            "cidade": "BRASILIA",
'            "codigoCidade": 2000,
'            "uf": "DF",
'            "cep": 72145760,
'            "indicadorComprovacao": "0"
'        },
'        "qrCode": {
'            "url": "",
'            "txId": "",
'            "emv": ""
'        },
'        "numero": "00031285570000043986",
'        "numeroCarteira": 17,
'        "numeroVariacaoCarteira": 35,
'        "codigoCliente": 704950857,
'        "linhaDigitavel": "00190000090312855700000043986173310260000010000",
'        "codigoBarraNumerico": "00193102600000100000000003128557000004398617",
'        "numeroContratoCobranca": 19581316,
'        "urlImagemBoleto": "https://boleto.apps.bb.com.br/segunda-via/00190000090312855700000043986173310260000010000/74910037000193/0",
'        "observacao": ""
'        }
            '-------------------
            

End Sub


Public Sub API_BBCobranca(Convenio As String, _
                        Carteira As String, _
                        carteiraVariacao As String, _
                        tipoConta As String, _
                        dataEmissao As String, _
                        dataVencimento As String, _
                        nFatura As String, _
                        nDuplicata As String, _
                        valor As String, _
                        vDeducao As String, _
                        vMulta As String, _
                        vJuros As String, _
                        diasProtesto As String, _
                        tpPessoa As Integer, _
                        cnpjPagador As String, _
                        cnpjBeneficiario As String, _
                        nomeBeneficiario As String, _
                        sMsg As String)
    Dim Id As Long
    Dim strJSON As String
  
    '-------------------------------------------------DADOS PARA API --------------------
     ' mJ(string , 0 = num // 1= str // 2= empty)
 
            
            

    
    strJSON = strJSON & mJ("numeroConvenio", Convenio, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroCarteira", Carteira, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroVariacaoCarteira", CInt(carteiraVariacao), 0) & "," & vbCrLf
    strJSON = strJSON & mJ("codigoModalidade", CInt(tipoConta), 0) & "," & vbCrLf
    
    
    strJSON = strJSON & mJ("dataEmissao", Replace(dataEmissao, "/", "."), 1) & "," & vbCrLf
    strJSON = strJSON & mJ("dataVencimento", Replace(dataVencimento, "/", "."), 1) & "," & vbCrLf
    strJSON = strJSON & mJ("valorOriginal", valor, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("valorAbatimento", vDeducoes, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("quantidadeDiasProtesto", diasProtesto, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("quantidadeDiasNegativacao", "0", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("orgaoNegativador", "10", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("indicadorAceiteTituloVencido", "S", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroDiasLimiteRecebimento", diasProtesto, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("codigoAceite", "A", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("codigoTipoTitulo", "2", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("descricaoTipoTitulo", "DM", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("indicadorPermissaoRecebimentoParcial", "N", 1) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroTituloBeneficiario", NumFatura, 1) & "," & vbCrLf
    strJSON = strJSON & mJ("campoUtilizacaoBeneficiario", Replace(NumDuplicata, "/", "-"), 1) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroTituloCliente", API_Calculo_NossoNumero(Id), 1) & "," & vbCrLf
    strJSON = strJSON & mJ("mensagemBloquetoOcorrencia", sMsg, 1) & ","
    
    'DESCONTO
    Dim tmpData As String
    Dim tmpValor As String
    tmpData = Replace(CDate(dataVencimento) + 1, "/", ".")
    tmpValor = vDeducoes
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
    Dim vMD As String
    Dim vMulta As String
    vMD = cobCalcMora(vDuplicata, 1, pJuros, "D")
    vMulta = cobCalcMulta(vDuplicata, vMulta, 1)
            
    
    'JUROS MORA
    strJSON = strJSON & vbCrLf
    strJSON = strJSON & mJ("jurosMora", "", 2) & "{" & vbCrLf
    strJSON = strJSON & mJ("tipo", "1", 0) & "," & vbCrLf
    strJSON = strJSON & mJ("porcentagem", 0, 0) & "," & vbCrLf
    'strJSON = strJSON & mJ("porcentagem", PgDadosFinanceiroFatura(Id).Juros, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("valor", vMD, 0) & vbCrLf
    strJSON = strJSON & vbCrLf & "},"
    
    'MULTA
    tmpData = Replace(CDate(PgDadosFinanceiroFatura(Id).Vencimento) + 1, "/", ".")
    tmpValor = vMulta
    If tmpValor = 0 Then tmpData = ""
    
    strJSON = strJSON & vbCrLf
    strJSON = strJSON & mJ("multa", "", 2) & "{" & vbCrLf
    strJSON = strJSON & mJ("tipo", IIf(vMulta = 0, 0, "1"), 0) & "," & vbCrLf
    strJSON = strJSON & mJ("data", tmpData, 1) & "," & vbCrLf
    strJSON = strJSON & mJ("porcentagem", 0, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("valor", vMulta, 0)
    strJSON = strJSON & vbCrLf & "},"
        
   'PAGADOR
    strJSON = strJSON & vbCrLf & _
     mJ("pagador", "", 2) & "{" & vbCrLf & _
    mJ("tipoInscricao", tpPessoa, 0) & "," & vbCrLf & _
    mJ("numeroInscricao", cnpjPagador, 1) & "," & vbCrLf & _
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
    mJ("numeroInscricao", cnpjBeneficiario, 0) & "," & vbCrLf & _
    mJ("nome", nomeBeneficiario, 1) & vbCrLf & _
    "}," & vbCrLf & _
    mJ("indicadorPix", "N", 1)
    strJSON = strJSON & vbCrLf & "}"
'----------------------------------------------------------------------
    Debug.Print "Boleto JSON: " & strJSON
End Sub

Private Function API_Calculo_NossoNumero(Id As Long) As String

'########################################################################################################
    '### Montagem de NOSSO NUMERO
    '### 28.02.25
    '########################################################################################################
    Dim Convenio As String
    
    Dim NN1, NN2 As String
    Convenio = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Convenio
    NN1 = Convenio
    Select Case Len(Trim(Convenio))
        Case Is <= 6
            'Dim NN1, NN2 As String
            
            If Len(Trim(Id)) > 5 Then
                    NN2 = Right(Id, 5)
                Else
                    NN2 = Left(String(5, "0"), 5 - Len(Trim(Id))) & Trim(Id)
            End If
            API_Calculo_NossoNumero = NN1 & NN2 & Trim(calculo_dv11base9(NN1 & NN2))
            
    
        Case 7
            NN1 = Left(String(7, "0"), 7 - Len(Trim(NN1))) & Trim(NN1)
            If Len(Trim(Id)) > 10 Then
                    NN2 = Right(Id, 10)
                Else
                    NN2 = Left(String(10, "0"), 10 - Len(Trim(Id))) & Trim(Id)
            End If
            API_Calculo_NossoNumero = "000" & NN1 & NN2
    End Select
    
 Debug.Print "API_Calculo_NossoNumero: " & API_Calculo_NossoNumero
    

End Function
Private Function API_CodigoBarras(Id As Long) As String
'pagina-4
'01-03 - Código do Banco na Câmara de Compensação = '001'
'04-04 - Código da Moeda = 9 (Real)
'05-05 - Digito Verificador (DV) do código de Barras*
'06-09 - Fator de Vencimento **
'10-19 - Valor
'20-25 - Zeros
'26-42 - Nosso Numero sem DV
'43-44 - Tipo de Carteira/Modalidade de Cobrança
'

    Dim NossoNumero As String
    Dim TipoCarteira As String
    Dim fator As String
    Dim valor As String
    Dim CodigoBarras As String
    Dim dvCB As String


    
    fator = CalculoFator(PgDadosFinanceiroFatura(Id).Vencimento)
    
    
    valor = IIf(Trim(PgDadosFinanceiroFatura(Id).vlCobrado) <> 0, PgDadosFinanceiroFatura(Id).vlCobrado, PgDadosFinanceiroFatura(Id).vlDuplicata)
    valor = RS(ChkVal(valor, 0, cDecMoeda))
    valor = Left(String(10, "0"), 10 - Len(valor)) & valor
    
    NossoNumero = API_Calculo_NossoNumero(Id)
    TipoCarteira = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Variacao
    '########################################################################################################
    '###  CODIGO DE BARRAS
    '########################################################################################################
    CodigoBarras = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & _
                   "9" & _
                   fator & _
                   Left(String(10, "0"), 10 - Len(valor)) & valor & _
                   valor & _
                   "000000" & _
                   NossoNumero & _
                   TipoCarteira
    
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
 
 Debug.Print "API_CodigoBarras: " & API_CodigoBarras
 
    API_CodigoBarras = CodigoBarras


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



