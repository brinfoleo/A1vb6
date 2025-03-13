Attribute VB_Name = "Modulo_CobrancaBB"

Public Sub mockLinhaDigitavel()
    Dim LinhaDigitavel As String
    
    LinhaDigitavel = GerarLinhaDigitavelBB( _
        banco:="001", _
        Moeda:="9", _
        Convenio:="3128557", _
        Carteira:="17", _
        NossoNumero:="00031285570000043986", _
        Valor:="0000010000", _
        FatorVencimento:=CalcularFatorVencimento("20/03/2025") _
    )
    
    Debug.Print LinhaDigitavel
End Sub
Public Sub mockGerarCodigoBarraBB()
Dim CodigoBarras As String

CodigoBarras = GerarCodigoBarrasBB( _
    banco:="001", _
    Moeda:="9", _
    Carteira:="17", _
    NossoNumero:="00031285570000043986", _
    Valor:="0000010000", _
    FatorVencimento:=CalcularFatorVencimento("20/03/2025") _
)

Debug.Print CodigoBarras
End Sub
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
             
               
            '--MULTA
             'Dim vMD As String
            'Dim vMulta As String
            'vMD = cobCalcMora(vDuplicata, 1, pJuros, "D")
            'vMulta = cobCalcMulta(vDuplicata, vMulta, 1)
                    
            
            
            
            'tpPessoa= IIf(LCase(PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Pessoa) = "fisica", 1, 2)
            
            
          
            
            Dim Sacado As String
            Dim vJurosMora As String
            vJurosMora = cobCalcMora("100", 1, 2, "D")
            
            Dim vMulta As String
            vMulta = cobCalcMulta("100", 0, 1)
            
            Sacado = pgDadosSacado(2, "74910037000193", "187 CENTRAL CARIOCA DE PECAS LTDA-EPP", "RUA DE SANTANA", "20230260", "RIO DE JANEIRO", "CENTRO", "RJ", "22219755", "email@email.com")
 
            
            API_BBCobranca "3128557", "17", "35", "4", "13/03/2025", "20/03/2025", "FAT123", "DUP123", "100.00", "0.00", vMulta, vJurosMora, "5", "2", "74910037000193", "98959112000179", "LIVRARIA CUNHA DA CUNHA", "00031285570000043986", Sacado, "MENSAGEM"
            
            
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


Private Sub API_BBCobranca(Convenio As String, _
                        Carteira As String, _
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
                        tpPessoa As Integer, _
                        cnpjPagador As String, _
                        cnpjBeneficiario As String, _
                        nomeBeneficiario As String, _
                        NossoNumero As String, _
                        Sacado As String, _
                        smsg As String)
    'Dim Id As Long
    Dim strJSON As String
  
    '-------------------------------------------------DADOS PARA API --------------------
     ' mJ(string , 0 = num // 1= str // 2= empty)
 
            
            

    
    strJSON = strJSON & mJ("numeroConvenio", Convenio, 0) & "," & vbCrLf
    strJSON = strJSON & mJ("numeroCarteira", Carteira, 0) & "," & vbCrLf
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
   strJSON = strJSON & vbCrLf & Sacado
    'strJSON = strJSON & vbCrLf & _
    ' mJ("pagador", "", 2) & "{" & vbCrLf & _
    'mJ("tipoInscricao", tpPessoa, 0) & "," & vbCrLf & _
    'mJ("numeroInscricao", cnpjPagador, 1) & "," & vbCrLf & _
    'mJ("nome", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Nome, 1) & "," & vbCrLf & _
    'mJ("endereco", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Lgr, 1) & "," & vbCrLf & _
    'mJ("cep", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).CEP, 1) & "," & vbCrLf & _
    'mJ("cidade", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Mun, 1) & "," & vbCrLf & _
    'mJ("bairro", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Bairro, 1) & "," & vbCrLf & _
    'mJ("uf", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).uf, 1) & "," & vbCrLf & _
    'mJ("telefone", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).Fone, 1) & "," & vbCrLf & _
    'mJ("email", PgDadosCliente(PgDadosFinanceiroFatura(Id).IDSacado).emailfin, 1) & vbCrLf & _
    '" },"
    
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

Private Function Calculo_NossoNumero(Id As Long) As String

'########################################################################################################
    '### Montagem de NOSSO NUMERO 20 CARACTERES
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
            Calculo_NossoNumero = NN1 & NN2 & Trim(calculo_dv11base9(NN1 & NN2))
            
    
        Case 7
            NN1 = Left(String(7, "0"), 7 - Len(Trim(NN1))) & Trim(NN1)
            If Len(Trim(Id)) > 10 Then
                    NN2 = Right(Id, 10)
                Else
                    NN2 = Left(String(10, "0"), 10 - Len(Trim(Id))) & Trim(Id)
            End If
            Calculo_NossoNumero = "000" & NN1 & NN2
    End Select
    
 Debug.Print "Calculo_NossoNumero: " & Calculo_NossoNumero
    

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
    Dim Valor As String
    Dim CodigoBarras As String
    Dim dvCB As String


    
    fator = CalculoFator(PgDadosFinanceiroFatura(Id).Vencimento)
    
    
    Valor = IIf(Trim(PgDadosFinanceiroFatura(Id).vlCobrado) <> 0, PgDadosFinanceiroFatura(Id).vlCobrado, PgDadosFinanceiroFatura(Id).vlDuplicata)
    Valor = RS(ChkVal(Valor, 0, cDecMoeda))
    Valor = Left(String(10, "0"), 10 - Len(Valor)) & Valor
    
    NossoNumero = Calculo_NossoNumero(Id)
    TipoCarteira = pgDadosConta(PgDadosFinanceiroFatura(Id).idConta).Variacao
    '########################################################################################################
    '###  CODIGO DE BARRAS
    '########################################################################################################
    CodigoBarras = pgDadosBanco(PgDadosFinanceiroFatura(Id).IdBanco).Numero & _
                   "9" & _
                   fator & _
                   Left(String(10, "0"), 10 - Len(Valor)) & Valor & _
                   Valor & _
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


Private Function pgDadosSacado(scTpInsc As String, scCNPJ As String, scNome As String, scLgr As String, scCEP As String, scMun As String, scBairro As String, scUF As String, scFone As String, scEmail As String) As String
   
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
    
    pgDadosSacado = strPagador
End Function


Private Function GerarCodigoBarrasBB( _
    ByVal banco As String, _
    ByVal Moeda As String, _
    ByVal NossoNumero As String, _
    ByVal Valor As String, _
    ByVal FatorVencimento As String, _
    ByVal Carteira As String _
) As String

    Dim CodigoBarras As String
    Dim DV As Integer

    ' 1. Monta o código de barras (sem o dígito verificador)
    CodigoBarras = banco & Moeda & FatorVencimento & Valor & "000" & NossoNumero & Carteira

    ' 2. Calcula o dígito verificador (Módulo 11)
    DV = Modulo11(CodigoBarras)

    ' 3. Insere o dígito verificador no código de barras
    CodigoBarras = banco & Moeda & DV & FatorVencimento & Valor & "000" & NossoNumero & Carteira

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
Private Function GerarLinhaDigitavelBB( _
    ByVal banco As String, _
    ByVal Moeda As String, _
    ByVal Convenio As String, _
    ByVal Carteira As String, _
    ByVal NossoNumero As String, _
    ByVal Valor As String, _
    ByVal FatorVencimento As String _
) As String

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

    ' 1. Monta o código de barras
    CodigoBarras = GerarCodigoBarrasBB( _
                                        banco:=banco, _
                                        Moeda:=Moeda, _
                                        Carteira:=Carteira, _
                                        NossoNumero:=NossoNumero, _
                                        Valor:=Valor, _
                                        FatorVencimento:=FatorVencimento _
                                    )
                                    
    ' 4. Divide o código de barras em campos
    Campo1 = banco & Moeda & Mid(CodigoBarras, 20, 5)
    Campo2 = Mid(CodigoBarras, 25, 10)
    Campo3 = Mid(CodigoBarras, 35, 10)
    Campo4 = "" 'Mid(CodigoBarras, 4, 1) & Mid(CodigoBarras, 5, 1) & Mid(CodigoBarras, 6, 3) & Mid(CodigoBarras, 7, 1)
    Campo5 = FatorVencimento & Valor '(CodigoBarras, 8, 10)

    ' 5. Calcula os dígitos verificadores de cada campo
    DV1 = Modulo10(Campo1)
    DV2 = Modulo10(Campo2)
    DV3 = Modulo10(Campo3)
    DV4 = Modulo10(Campo4)

    ' 6. Monta a linha digitável
    GerarLinhaDigitavelBB = Campo1 & DV1 & "" & Campo2 & DV2 & "" & Campo3 & DV3 & "" & Campo4 & DV4 & "" & Campo5

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



 Function Modulo10(ByVal Numero As String) As Integer

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

