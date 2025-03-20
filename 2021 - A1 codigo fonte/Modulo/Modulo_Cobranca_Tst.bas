Attribute VB_Name = "Modulo_Cobranca_Tst"
Private tstConvenio As String
Private tstIdFatura As Long
Private tstCarteira As String
Private tstcarteiraVariacao As String
Private tstTipoConta As String
Private tstValor As String
Private tstEmissao As String
Private tstVencimento As String

Private cnpjBeneficiario As String
Private nomeBeneficiario As String


Public Sub InicializarVariaveis()
    
    
    
    tstIdFatura = 44070
    tstConvenio = "3741178"
    tstCarteira = "17"
    tstcarteiraVariacao = "27"
    tstTipoConta = "4"
    tstValor = "3358.98"
    
    tstEmissao = "11/03/2025"
    tstVencimento = "10/05/2025"
    
    cnpjBeneficiario = "98959112000179"
    nomeBeneficiario = "LIVRARIA CUNHA DA CUNHA"
    '***********************************************************
    'CAMPOS OBRIGATORIOS PARA TESTE DE HOMOLOGACAO
    'tstConvenio = "3128557"
    'tstCarteira = "17"
    'tstcarteiraVariacao = "35"
    'cnpjBeneficiario = "98959112000179"
    'nomeBeneficiario = "LIVRARIA CUNHA DA CUNHA"
    '***********************************************************
End Sub

Public Function mockGerarNossoNumero() As String
    InicializarVariaveis
    Dim bbCob As New BBCobranca
    Dim NossoNumero As String
   
    
    NossoNumero = bbCob.GerarNossoNumero(tstConvenio, tstIdFatura)
    mockGerarNossoNumero = NossoNumero
End Function
Public Sub mockLinhaDigitavel()
    InicializarVariaveis
    Dim bbCob As New BBCobranca
    Dim LinhaDigitavel As String
    
    LinhaDigitavel = bbCob.GerarLinhaDigitavelBB( _
        banco:="001", _
        moeda:="9", _
        Convenio:=tstConvenio, _
        carteira:=tstCarteira, _
        NossoNumero:=bbCob.GerarNossoNumero(tstConvenio, tstIdFatura), _
        Valor:=tstValor, _
        Vencimento:=tstVencimento _
    )
    
    Debug.Print LinhaDigitavel
End Sub
    Public Sub mockGerarCodigoBarraBB()
    InicializarVariaveis
    Dim bbCob As New BBCobranca
    Dim CodigoBarras As String
    
    CodigoBarras = bbCob.GerarCodigoBarrasBB( _
        banco:="001", _
        moeda:="9", _
        carteira:=tstCarteira, _
        NossoNumero:=bbCob.GerarNossoNumero(tstConvenio, tstIdFatura), _
        Valor:=tstValor, _
        Vencimento:=tstVencimento _
    )
    
    Debug.Print CodigoBarras
    
End Sub
Public Function mockGerarBoleto()
    InicializarVariaveis
    Dim bbCob As New BBCobranca
    'Modulo Homologacao
    Dim tstBoleto As String
    
    Dim vJurosMora As String
     vJurosMora = cobCalcMora(tstValor, 1, 2, "D")
     
     Dim vMulta As String
     vMulta = cobCalcMulta(tstValor, 0, 1)
    
        Dim Sacado As String
     Sacado = bbCob.jsonSacado(2, "74910037000193", "187 CENTRAL CARIOCA DE PECAS LTDA-EPP", "RUA DE SANTANA", "20230260", "RIO DE JANEIRO", "CENTRO", "RJ", "22219755", "email@email.com")

     
     tstBoleto = bbCob.GerarBoletoBB(Convenio:=tstConvenio, _
                                    carteira:=tstCarteira, _
                                    carteiraVariacao:=tstcarteiraVariacao, _
                                    tipoConta:=tstTipoConta, _
                                    dataEmissao:=tstEmissao, _
                                    DataVencimento:=tstVencimento, _
                                    nFatura:="FAT" & tstIdFatura, _
                                    nDuplicata:="DUP" & tstIdFatura, _
                                    Valor:=tstValor, _
                                    vDeducao:="0.00", _
                                    vMulta:=vMulta, _
                                    vJuros:=vJurosMora, _
                                    DiasProtesto:="5", _
                                    Sacado:=Sacado, _
                                    cnpjBeneficiario:=cnpjBeneficiario, _
                                    nomeBeneficiario:=nomeBeneficiario, _
                                    NossoNumero:=bbCob.GerarNossoNumero(tstConvenio, tstIdFatura), _
                                    smsg:="MENSAGEM")
            
    Debug.Print tstBoleto
    mockGerarBoleto = tstBoleto
    
End Function


