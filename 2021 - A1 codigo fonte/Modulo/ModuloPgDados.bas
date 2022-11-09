Attribute VB_Name = "ModuloPgDados"
Option Explicit
Type Dados_Empresa
    Nome            As String
    Fant            As String
    Lgr             As String
    Nro             As String
    Cpl             As String
    Bairro          As String
    Mun             As String
    uf              As String
    CEP             As String
    pais            As String
    Fone            As String
    Mail            As String
    CNPJ            As String
    IE              As String
    iest            As String
    im              As String
    cnae            As String
    RegimeTrib      As String
    PISCST          As String
    PISAliquota     As String
    COFINSCST       As String
    COFINSAliquota  As String
    Suframa         As String
    TipoAtividade   As Integer
    Logotipo        As String
    cNome           As String
    cCNPJ           As String
    cIE             As String
    cIM             As String
    cEndereco       As String
    cNumero         As String
    cCompl          As String
    cBairro         As String
    cUF             As String
    cMunicipio      As String
    cCEP            As String
    cFone1          As String
    cFone2          As String
    cMail           As String
    cCodID          As String
    crNome          As String
    crCPF           As String
    crCRC           As String
    crFone          As String
    crMail          As String
End Type

Type Dados_Cliente
    status              As String
    Pessoa              As String
    Fant                As String
    Doc                 As String
    Nome                As String
    Obs                 As String
    IE                  As String
    iest                As String
    im                  As String
    cnae                As String
    emailnfe            As String
    emailfin            As String
    emailcom            As String
    website             As String
    LimiteCredito       As String
    TipoDocumento       As String
    'localcobranca       As String
    condicoespagamento  As String
    CentroCustos         As String
    PlanoContas         As Integer
    Transportadora      As String
    'cobrancalgr         As String
    'cobrancanro         As String
    'cobrancacpl         As String
    'cobrancauf          As String
    'cobrancamun         As String
    'cobrancacep         As String
    entregaDoc           As String
    entregalgr          As String
    entreganro          As String
    entregacpl          As String
    entregabairro       As String
    entregauf           As String
    entregamun          As String
    entregacep          As String
    cobranca            As String
    entrega             As String
    Lgr                 As String
    Nro                 As String
    Cpl                 As String
    Bairro              As String
    uf                  As String
    Mun                 As String
    CEP                 As String
    Mail                As String
    Fone                As String
    Suframa             As String
    Vendedor            As Integer
    ObsCobNfe           As String
    ObsCobBoleto        As String
End Type
Type Dados_Fornecedor
    status              As String
    Pessoa              As String
    Fant                As String
    Doc                 As String
    Nome                As String
    Obs                 As String
    IE                  As String
    iest                As String
    im                  As String
    cnae                As String
    emailnfe            As String
    emailfin            As String
    emailcom            As String
    website             As String
    LimiteCredito       As String
    TipoDocumento       As String
    localcobranca       As String
    condicoespagamento  As String
    PlanoContas         As String
    Transportadora      As String
    cobrancalgr         As String
    cobrancanro         As String
    cobrancacpl         As String
    cobrancauf          As String
    cobrancamun         As String
    cobrancacep         As String
    entregalgr          As String
    entreganro          As String
    entregacpl          As String
    entregabairro       As String
    entregauf           As String
    entregamun          As String
    entregacep          As String
    cobranca            As String
    entrega             As String
    Lgr                 As String
    Nro                 As String
    Cpl                 As String
    Bairro              As String
    uf                  As String
    Mun                 As String
    CEP                 As String
    Mail                As String
    Fone                As String
End Type
Type Dados_Transportadora
    Pessoa  As String
    Nome    As String
    Fant    As String
    Lgr     As String
    Bairro  As String
    Mun     As String
    uf      As String
    CEP     As String
    Fone    As String
    Mail    As String
    CNPJ    As String
    IE      As String
End Type

Type Dados_EstoqueProduto
    Id                  As Integer
    IdDeposito          As Integer
    Referencia          As String
    status              As String
    CodBarras           As String
    Descricao           As String
    Grupo               As String
    subGrupo            As String
    NCM                 As String
    MVA                 As String
    ICMSOrigem          As Integer
    ICMSCST             As String
    IPIAliquota         As String
    IPICST              As String
    Enquadramento       As Integer
    VlCusto             As String
    Unidade             As String
    VlIPI               As String
    VlOutros            As String
    MarkUp              As String
    VlTabela            As String
    InfCompl            As String
    Saldo               As String
End Type
Type Dados_RHFuncionario
    CPF                 As String
    RG                  As String
    Nome                As String
    Cargo               As String
    Endereco            As String
    Num                 As String
    Compl               As String
    Bairro              As String
    Municipio           As String
    uf                  As String
    CEP                 As String
    Mail                As String
    Tel                 As String
    Salario             As String
    Comissao            As String
    Assinatura          As String
End Type
'Type Dados_UF
'    Id          As String
'    UF          As String
'    Descricao   As String
'End Type
Type Dados_Municipio
    Id          As String
    uf          As String
    codUF       As String
    codMun      As String
    Descricao   As String
End Type

Type Dados_MovEstoque
    Id          As String
    Descricao   As String
    Sigla       As String
    acao        As String
    AcaoDescr   As String
End Type
Type Dados_Banco
    Id          As String
    Nome        As String
    Numero      As String
End Type
Type Dados_Conta
    Id              As String
    banco           As String
    agencia         As String
    AgenciaDV       As String
    conta           As String
    ContaDV         As String
    Multa           As String
    Juros           As String
    DiasProtesto    As String
    
    Contrato        As String
    carteira        As String
    Variacao        As String
    Convenio        As String
    ConvenioLider   As String
    Tipo            As String
    Saldo           As String
End Type
Type Dados_CentroCustos
    Id          As String
    Descricao   As String
    Sigla       As String
End Type
Type Dados_TipoDocumento
    Id          As String
    Tipo        As String
    Descricao   As String
    Sigla       As String
    Impressao   As String
    formaPgto   As String
End Type
Type Dados_ICMS
    Id          As String
    Descricao   As String
    Sigla       As String
    ICMS        As String
    ICMSInt     As String
    ICMSFECP    As String
    codUF       As String
End Type

Type Dados_TipoNotaFiscal
    Descricao           As String
    TipoNota            As String
    TipoNotaDescr       As String
    Serie               As String
    Modelo              As String
    NumInicial          As Integer
    Natureza            As String
    Finalidade          As String
    FinalidadeDescr     As String
    EnvioRF             As Integer
    ChaveAcessoRef      As Integer
    MovEstoque          As Integer
    MovFisco            As Integer
    MovContasPR         As Integer
    MovComissao         As Integer
    ModBC               As Integer
    ModBCST             As Integer
    ImpCmpFatura        As Integer
    ImpDtSaida          As Integer
    ImpInfCompl         As Integer
    Obs                 As String
    conta               As Integer
    CentroCusto         As Integer
    PlanoContas         As Integer
    TipoDoc             As Integer
    CSTPIS              As String
    CSTCOFINS           As String
    ImpBCICMS           As Integer
    ImpvICMS            As Integer
    ImpBCICMSST         As Integer
    ImpvICMSST          As Integer
    ImpvTotalProduto    As Integer
    ImpvFrete           As Integer
    ImpvSeguro          As Integer
    ImpvDesconto        As Integer
    ImpvOutrasDesp      As Integer
    ImpvIPI             As Integer
    ImpvTotalNota       As Integer
End Type
Type Dados_Configuracoes
    DtUltMov                As Date
    
    cDecMoeda               As Integer
    cDecQtd                 As Integer
    EmissaoNFesPV           As Integer
    VisualizarOutrosFunc    As Integer 'Visualizar os dados de outros vendedores
    
    ImpBoleto               As Integer
    TranspVolumes           As Integer 'Checa se os dados da NF forao preenchidos
    
    CodProdImpresso         As Integer 'Codigo que sera usado na NFE 1 - Interno / 2 - Referencia
    
    EstoqueAtualizarCusto   As Integer 'Atualiza o CUSTO na entrada da NFe
    EstoqueSUverDepositos   As Integer 'Permite ou nao o super Usuario a ver outros depositos
    EstoqueDepositoPadrao   As Integer
    
    NFDevolucaoCompra       As Integer 'Verifica se nos relatorios as entradas por devolucao entao como compra
    EntradaNFSemAutorSEFAZ  As Integer 'Aceitar NFE de entrada sem autorizacao da SEFAZ
    
    ClienteAplLimiteCredito As Integer
    
    fusoHorario             As String  'Pega o fuso horario local
    
    MenuManutencaoTabelas   As Integer
    pXMLFornecedor          As String
    pFileArmazenamento      As String
    pUniDANFe               As String
    DANFenCopias            As Integer
    'DANFeEnviarMail         As Integer
    'DANFEEnviarMailCC       As Integer
    'DANFEeMailCC            As String
    DANFeVisualizar         As Integer
    DANFEPreview            As Integer
    InserirNomeVendXML      As Integer
    
    BloqueionNFManual       As Integer
    
    MailSMTPPorta           As String
    MailSMTP                As String
    MailEndereco            As String
    MailLogin               As String
    MailSenha               As String
    MailAutenticacao        As Integer
    MailRecCopia            As Integer
    
    RHConta                 As Integer
    RHCentroCustos          As Integer
    RHDocumento             As Integer
    RHPlanoContas           As Integer
    
    FornecedorCC            As Integer
    FornecedorTpDoc         As Integer
    FornecedorPlanoContas   As Integer
    
    uf                      As Integer
    Ambiente                As Integer
    
    TpEmissao               As Integer
    ContingenciaDt          As String
    ContingenciaHr          As String
    ContingenciaMotivo      As String
    
    FormatoPasta            As String
    DiasXMLTemp             As Integer
    GravarRetornoTXT        As Integer
    NFePrazoCancelamento    As Integer
    
    pEnviados               As String
    pEnviadosLote           As String
    pRetorno                As String
    pEnvio                  As String
    pErro                   As String
    pBackup                 As String
    pValidar                As String
    
    IniValCertDigital       As String
    FinValCertDigital       As String
End Type

Type Dados_FinanceiroFatura
    Id                      As Integer
    ContaPR                 As String
    TpConta                 As String
    emissao                 As Date
    NumFatura               As String
    vlFatura                As String
    idConta                 As Integer
    idCentroCustos          As Integer
    idPlanoContas           As Integer
    idTpDoc                 As Integer
    Tabela                  As String
    IDSacado                As Integer
    CNPJSacado              As String
    Sacado                  As String
    CodigoBarras            As String
    LinhaDigitavel          As String
    NossoNumero             As String
    Vencimento              As Date
    NumDuplicata            As String
    vlDuplicata             As String
    Multa                   As String
    Juros                   As String
    DiasProtesto            As Integer
    IdBanco                 As String ' 18.08.17
    Acrescimo               As String
    Abatimento              As String
    Deducoes                As String
    MultaMora               As String
    vlCobrado               As String
    DataQuitacao            As Date
    Obs                     As String
    ObsBol1                 As String
    ObsBol2                 As String
    ObsBol3                 As String
End Type

Type Dados_Usuario
    Id                      As Integer
    Nome                    As String
    Login                   As String
    senha                   As String
    idFunc                  As Integer
    Grupo                   As Integer
    SenhaNuncaExp           As Integer
    TrocarSenha             As Integer
    SuperUsuario            As Integer
    Menus                   As String
End Type
Type Dados_CST
    Id                      As Integer
    Descricao               As String
    cst                     As String
    Tabela                  As String
End Type
Type Dados_CFOP
    Situacao                As Integer
    cst                     As String
    CFOP                    As String
    ICMS                    As Integer
    ICMSST                  As Integer
End Type
Type Dados_UsuGrupo
    Nome        As String
    Descricao   As String
End Type

Type Dados_NotaFiscal
    'Outros
    EnvioRF                 As Integer
    MovFinanceiro           As Integer
    MovFisco                As Integer
    ImpFatura               As Integer
    '*********************************************************************************
    'NFe Autorizada
     Id                     As Integer
     nProt                  As String
     dhProt                 As String
     lote                   As String
     nRecibo                As String
     cStat                  As String
     xMotivo                As String
     StatusNFe              As String
    '*********************************************************************************
    'NFe Cancelada
     canc_nProt             As String
     canc_dhRecbto          As String
     canc_xJust             As String
     canc_Status            As String
    '*********************************************************************************
    'Numero de NFe Inutilizado
     inut_nProt             As String
     inut_dhRecbto          As String
     inut_xJust             As String
     inut_Status As String
    '*********************************************************************************
    'cabecario do Pedido (ide)
     Versao                 As String
     idNFe                  As String
     ide_cUF                As String
     ide_cNF                As String
     ide_natOp              As String
     ide_indPag             As String
     ide_mod                As String
     ide_serie              As String
     ide_nNF                As String
     ide_dEmi               As String
     ide_dSaiEnt            As String
     ide_hSaiEnt            As String
     ide_tpNF               As String
     ide_cMunFG             As String
     ide_refNFe             As String
     ide_tpImp              As String
     ide_tpEmis             As String
     ide_cDV                As String
     ide_tpAmb              As String
     ide_finNFe             As String
     ide_procEmi            As String
     ide_verProc            As String
    'Emitente
     emit_CNPJ              As String
     emit_xNome             As String
     emit_xFant             As String
     emit_xLgr              As String
     emit_nro               As String
     emit_xCpl              As String
     emit_Bairro            As String
     emit_cMun              As String
     emit_xMun              As String
     emit_UF                As String
     emit_CEP               As String
     emit_cPais             As String
     emit_xPais             As String
     emit_fone              As String
     emit_IE                As String
    
     emit_IEST              As String
     emit_IM                As String
     emit_CNAE              As String
    
     emit_CRT               As String
    
    'Destinatario
     dest_idDest            As String
     dest_pessoa            As String
     dest_CNPJ              As String
     dest_xNome             As String
     dest_xFant             As String
     dest_xLgr              As String
     dest_nro               As String
     dest_xCpl              As String
     dest_Bairro            As String
     dest_cMun              As String
     dest_xMun              As String
     dest_UF                As String
     dest_CEP               As String
     dest_cPais             As String
     dest_xPais             As String
     dest_fone              As String
     dest_IE                As String
     dest_ISUF              As String
     dest_email             As String
     infAdic_infCpl         As String
    'Transporte
     transp_modFrete        As String
     transp_Pessoa          As String
     transp_CNPJ            As String
     transp_xNome           As String
     transp_IE              As String
     transp_xEnder          As String
     transp_xMun            As String
     transp_UF              As String
     transp_qVol            As String
     transp_esp             As String
     transp_marca           As String
     transp_nVol            As String
     transp_pesoL           As String
     transp_pesoB           As String
        
    'TOTAIS
     total_vBC              As String
     total_vICMS            As String
     total_vBCST            As String
     total_vICMSST          As String
     total_vProd            As String
     total_vFrete           As String
     total_vSeg             As String
     total_vDesc            As String
     total_vIPI             As String
     total_vPIS             As String
     total_vCOFINS          As String
     total_vOutro           As String
     total_vNF              As String
     ger_Vendedor           As String
     ger_idPV               As String
    
    
    'Produto******************************************************************************
    
    ' IdNFe                  As String
    ' det_IdProduto          As String
    ' det_cProd              As String
    ' det_cEAN               As String
    ' det_xProd              As String
    ' det_InfAdProd          As String
    ' det_NCM                As String
    ' det_EXTIPI             As String
    ' det_CFOP               As String
    ' det_uCom               As String
    ' det_qCom               As String
    ' det_vUnCom             As String
    ' det_vProd              As String
    
    ' det_cEANTrib           As String
    ' det_uTrib              As String
    ' det_qTrib              As String
    ' det_vUnTrib            As String
    ' det_vFrete             As String
    ' det_vSeg               As String
    ' det_vDesc              As String
    ' det_vOutro             As String
    ' det_indTot             As String
    ' det_xPed               As String
    ' det_nItemPed           As String
    
    'IMPOSTOS
    'ICMS - 'Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST|vICMSST
    ' ICMS_origem            As String
    ' ICMS_CST               As String
    ' ICMS_modBC             As String
    ' ICMS_pRedBC            As String
    ' ICMS_vBC               As String
    ' ICMS_pICMS             As String
    ' ICMS_vICMS             As String
    ' ICMS_modBCST           As String
    ' ICMS_pMVAST            As String
    ' ICMS_pRedBCST          As String
    ' ICMS_vBCST             As String
    ' ICMS_pICMSST           As String
    ' ICMS_vICMSST           As String
    ' ICMS_MotDesICMS        As String
    ''IPI
    ' IPI_cEnq               As String
    ' IPI_CST                As String
    ' IPI_vBC                As String
    ' IPI_pIPI               As String
    ' IPI_vIPI               As String
    ''PIS
    ' PIS_CST                As String
    ' PIS_vBC                As String
    ' PIS_pPIS               As String
     'PIS_vPIS               As String
    ''COFINS
    ' COFINS_CST             As String
    ' COFINS_vBC             As String
    ' COFINS_pCOFINS         As String
    ' COFINS_vCOFINS         As String
    ''Informacoes Gerenciais
    ' estoque_Unid           As String
    ' estoque_Qtd            As String
    ' estoque_vUnit          As String
    ' comissao_pComissao     As String
    ' comissao_vComissao     As String
   
  
    
    ''COBRANCA
    ' IdNFe                  As String
    ' cobr_TpDoc             As String
    ' cobr_nFat              As String
    ' cobr_vOrig             As String
    ' cobr_vDesc             As String
    ' cobr_vLiq              As String
    ' cobr_nDup              As String
    ' cobr_dVenc             As String
    ' cobr_vDup              As String
    ' cobr_Emissao           As String
    ' cobr_Multa             As String
    ' cobr_Mora              As String
    ' cobr_Protesto As String
    'cobr_idCliente As String
    
    
    
    'Email
    'email_IdNFe As String
    'Email_Status As String
    
End Type
Type Dados_NCM
    Id          As Integer
    Descricao   As String
    pIPI        As String
    NCM         As String
    cest        As String
End Type
Type Dados_PlanoContas
    Id          As Integer
    Codigo      As String
    Descricao   As String
    cd          As String
    totalizador As Integer
End Type
Type Dados_CEST
    Id          As Integer
    Descricao   As String
    NCM         As String
    cest        As String
End Type


Public Function PgDadosNCM(sCampo As String, sBusca As String, SN As String) As Dados_NCM
    Dim Rst     As Recordset
    Dim sSQL    As String
    If SN = "N" Then
            sSQL = "SELECT * FROM tributacaoncm WHERE ID_Empresa = " & ID_Empresa & " AND " & sCampo & " = " & sBusca
        Else
            sSQL = "SELECT * FROM tributacaoncm WHERE ID_Empresa = " & ID_Empresa & " AND " & sCampo & " = '" & sBusca & "'"
    End If
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'MsgBox "Erro ao localizar dados NCM", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            PgDadosNCM.Id = cNull(Rst.Fields("Id"))
            PgDadosNCM.Descricao = cNull(Rst.Fields("Descricao"))
            PgDadosNCM.NCM = cNull(Rst.Fields("NCM"))
            PgDadosNCM.pIPI = cNull(Rst.Fields("IPI"))
            PgDadosNCM.cest = cNull(Rst.Fields("cest"))
    End If
    Rst.Close
End Function
Public Function PgDadosUsuGrupo(idGrupo As Integer) As Dados_UsuGrupo
    Dim Rst     As Recordset
    Dim sSQL    As String
    sSQL = "SELECT * FROM UsuGerenciadorGrupo WHERE ID_Empresa = " & ID_Empresa & " AND id = " & idGrupo
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            PgDadosUsuGrupo.Nome = IIf(IsNull(Rst.Fields("grupo_Nome")), "", Rst.Fields("grupo_Nome"))
            PgDadosUsuGrupo.Descricao = IIf(IsNull(Rst.Fields("grupo_Descricao")), "", Rst.Fields("grupo_Descricao"))
    End If
    Rst.Close
End Function

Public Function PgDadosConfig() As Dados_Configuracoes
    On Error GoTo TrtErro
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    sSQL = "SELECT * FROM configuracoes WHERE ID_Empresa = " & ID_Empresa
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'MsgBox "Favor verificar as configurações do sistema"
            RegLog "", "", "Erro ao consultar na base de dados as configurações do sistema"
            Exit Function
        Else
            Rst.MoveFirst
    End If
    
    PgDadosConfig.DtUltMov = IIf(IsNull(Rst.Fields("DtUltMov")), Date, Rst.Fields("DtUltMov"))
    
    PgDadosConfig.cDecMoeda = IIf(IsNull(Rst.Fields("cDecMoeda")), "2", Rst.Fields("cDecMoeda"))
    PgDadosConfig.cDecQtd = IIf(IsNull(Rst.Fields("cDecQtd")), "2", Rst.Fields("cDecQtd"))
    PgDadosConfig.EmissaoNFesPV = IIf(IsNull(Rst.Fields("EmissaoNFesPV")), "0", Rst.Fields("EmissaoNFesPV"))
    
    PgDadosConfig.ClienteAplLimiteCredito = IIf(IsNull(Rst.Fields("ClienteLimiteCredito")), "0", Rst.Fields("ClienteLimiteCredito"))
    
    PgDadosConfig.fusoHorario = IIf(IsNull(Rst.Fields("fusoHorario")), "-03:00", Rst.Fields("fusoHorario"))
    
    'Estoque
    PgDadosConfig.EstoqueAtualizarCusto = IIf(IsNull(Rst.Fields("EstoqueAtualizarCusto")), "0", Rst.Fields("EstoqueAtualizarCusto"))
    PgDadosConfig.EstoqueSUverDepositos = IIf(IsNull(Rst.Fields("EstoqueSUverDepositos")), "0", Rst.Fields("EstoqueSUverDepositos"))
    ID_Deposito = IIf(IsNull(Rst.Fields("EstoqueDepositoPadrao")), "0", Rst.Fields("EstoqueDepositoPadrao"))
    PgDadosConfig.EstoqueDepositoPadrao = ID_Deposito
    
    'CodProdImpresso
    PgDadosConfig.CodProdImpresso = IIf(IsNull(Rst.Fields("CodProdImpresso")), 1, CInt(Rst.Fields("CodProdImpresso")))
    
    PgDadosConfig.DANFEPreview = IIf(IsNull(Rst.Fields("PreviewDanfe")), "0", Rst.Fields("PreviewDanfe"))
    
    PgDadosConfig.BloqueionNFManual = IIf(IsNull(Rst.Fields("BloqueionNFManual")), "0", Rst.Fields("BloqueionNFManual"))
    
    PgDadosConfig.TranspVolumes = IIf(IsNull(Rst.Fields("TranspVolumes")), "0", Rst.Fields("TranspVolumes"))
    
    PgDadosConfig.VisualizarOutrosFunc = IIf(IsNull(Rst.Fields("GerClientesVisualizarOutrosFunc")), "0", Rst.Fields("GerClientesVisualizarOutrosFunc"))
    
    PgDadosConfig.ImpBoleto = IIf(IsNull(Rst.Fields("Boleto")), "2", Rst.Fields("Boleto"))
    
    PgDadosConfig.pXMLFornecedor = IIf(IsNull(Rst.Fields("pXMLFornecedor")), 0, Rst.Fields("pXMLFornecedor"))
    PgDadosConfig.pFileArmazenamento = IIf(IsNull(Rst.Fields("pFileArmazenamento")), "", Rst.Fields("pFileArmazenamento"))
    
    PgDadosConfig.pUniDANFe = IIf(IsNull(Rst.Fields("pUniDANFe")), "", Rst.Fields("pUniDANFe"))
    
    'RH
    PgDadosConfig.RHCentroCustos = IIf(IsNull(Rst.Fields("RHCentroCustos")), "0", Rst.Fields("RHCentroCustos"))
    PgDadosConfig.RHConta = IIf(IsNull(Rst.Fields("RHConta")), "0", Rst.Fields("RHConta"))
    PgDadosConfig.RHDocumento = IIf(IsNull(Rst.Fields("RHDocumento")), "0", Rst.Fields("RHDocumento"))
    PgDadosConfig.RHPlanoContas = IIf(IsNull(Rst.Fields("RHPlanoContas")), "0", Rst.Fields("RHPlanoContas"))
    
    
    'Fornecedor
    PgDadosConfig.FornecedorPlanoContas = IIf(IsNull(Rst.Fields("FornecedorPlanoContas")), "0", Rst.Fields("FornecedorPlanoContas"))
    PgDadosConfig.FornecedorTpDoc = IIf(IsNull(Rst.Fields("FornecedorTpDoc")), "0", Rst.Fields("FornecedorTpDoc"))
    PgDadosConfig.FornecedorCC = IIf(IsNull(Rst.Fields("FornecedorCC")), "0", Rst.Fields("FornecedorCC"))
    PgDadosConfig.NFDevolucaoCompra = IIf(IsNull(Rst.Fields("NFDevolucaoCompra")), 0, Rst.Fields("NFDevolucaoCompra"))
    PgDadosConfig.EntradaNFSemAutorSEFAZ = IIf(IsNull(Rst.Fields("AceitarEntradaNFSemAutorizacaoSEFAZ")), 0, Rst.Fields("AceitarEntradaNFSemAutorizacaoSEFAZ"))
    
    
    PgDadosConfig.DANFenCopias = IIf(IsNull(Rst.Fields("DANFenCopias")), "0", Rst.Fields("DANFenCopias"))
    
    'PgDadosConfig.DANFeEnviarMail = IIf(IsNull(Rst.Fields("DANFEEnviarMail")), "0", Rst.Fields("DANFEEnviarMail"))
    'PgDadosConfig.DANFEEnviarMailCC = IIf(IsNull(Rst.Fields("DANFEEnviarMailCC")), "0", Rst.Fields("DANFEEnviarMailCC"))
    'PgDadosConfig.DANFEeMailCC = IIf(IsNull(Rst.Fields("eMailCC")), "", Rst.Fields("eMailCC"))
    
    'E-Mail
    
    PgDadosConfig.MailSMTP = IIf(IsNull(Rst.Fields("MailSMTP")), "", Rst.Fields("MailSMTP"))
    PgDadosConfig.MailSMTPPorta = IIf(IsNull(Rst.Fields("MailSMTPPorta")), "25", Rst.Fields("MailSMTPPorta"))
    PgDadosConfig.MailEndereco = IIf(IsNull(Rst.Fields("MailEndereco")), "", Rst.Fields("MailEndereco"))
    PgDadosConfig.MailLogin = IIf(IsNull(Rst.Fields("MailLogin")), "", Rst.Fields("MailLogin"))
    PgDadosConfig.MailSenha = IIf(IsNull(Rst.Fields("MailSenha")), "", Rst.Fields("MailSenha"))
    PgDadosConfig.MailAutenticacao = IIf(IsNull(Rst.Fields("MailAutenticacao")), "0", Rst.Fields("MailAutenticacao"))
    PgDadosConfig.MailRecCopia = IIf(IsNull(Rst.Fields("MailRecCopia")), "0", Rst.Fields("MailRecCopia"))

    
    PgDadosConfig.DANFeVisualizar = IIf(IsNull(Rst.Fields("DANFEVisualizar")), "0", Rst.Fields("DANFEVisualizar"))

    PgDadosConfig.MenuManutencaoTabelas = IIf(IsNull(Rst.Fields("MenuManutencaoTabelas")), 0, Rst.Fields("MenuManutencaoTabelas"))
    PgDadosConfig.uf = IIf(IsNull(Rst.Fields("EstadoUF")), 0, Rst.Fields("EstadoUF"))
    PgDadosConfig.Ambiente = IIf(IsNull(Rst.Fields("Ambiente")), 0, Rst.Fields("Ambiente"))
    
    PgDadosConfig.TpEmissao = IIf(IsNull(Rst.Fields("TipoEmissao")), 0, Rst.Fields("TipoEmissao"))
    
    PgDadosConfig.ContingenciaDt = cNull(Rst.Fields("DataContigencia"))
    PgDadosConfig.ContingenciaHr = cNull(Rst.Fields("HoraContigencia"))
    PgDadosConfig.ContingenciaMotivo = cNull(Rst.Fields("MotivoContigencia"))
    
    PgDadosConfig.FormatoPasta = IIf(IsNull(Rst.Fields("FormatoPasta")), 0, Rst.Fields("FormatoPasta"))
    PgDadosConfig.DiasXMLTemp = IIf(IsNull(Rst.Fields("DiasXMLTemp")), 0, Rst.Fields("DiasXMLTemp"))
    PgDadosConfig.GravarRetornoTXT = IIf(IsNull(Rst.Fields("RetornoTXT")), 0, Rst.Fields("RetornoTXT"))
    PgDadosConfig.InserirNomeVendXML = IIf(IsNull(Rst.Fields("InserirNomeVendXML")), 0, Rst.Fields("InserirNomeVendXML"))
    'NFePrazoCancelamento
    PgDadosConfig.NFePrazoCancelamento = IIf(IsNull(Rst.Fields("NFePrazoCancelamento")), 0, Rst.Fields("NFePrazoCancelamento"))
    
    
    PgDadosConfig.pEnviados = IIf(IsNull(Rst.Fields("pEnviados")), 0, Rst.Fields("pEnviados"))
    PgDadosConfig.pEnviadosLote = IIf(IsNull(Rst.Fields("pEnviadosLote")), 0, Rst.Fields("pEnviadosLote"))
    PgDadosConfig.pRetorno = IIf(IsNull(Rst.Fields("pRetorno")), 0, Rst.Fields("pRetorno"))
    PgDadosConfig.pEnvio = IIf(IsNull(Rst.Fields("pEnvio")), 0, Rst.Fields("pEnvio"))
    PgDadosConfig.pErro = IIf(IsNull(Rst.Fields("pErro")), 0, Rst.Fields("pErro"))
    PgDadosConfig.pBackup = IIf(IsNull(Rst.Fields("pBackup")), 0, Rst.Fields("pBackup"))
    PgDadosConfig.pValidar = IIf(IsNull(Rst.Fields("pValidar")), 0, Rst.Fields("pValidar"))
    
    PgDadosConfig.IniValCertDigital = IIf(IsNull(Rst.Fields("InicioValidadeCertDigital")), "", Rst.Fields("InicioValidadeCertDigital"))
    PgDadosConfig.FinValCertDigital = IIf(IsNull(Rst.Fields("FinalValidadeCertDigital")), "", Rst.Fields("FinalValidadeCertDigital"))
    
    ' IniValCertDigital       As String
    'FinValCertDigital       As String
    Rst.Close
    Exit Function
TrtErro:
   ' MsgBox "Erro na Configuração do sistema!" & vbCrLf & _
            "Favor verificar", vbCritical, "Configurações"
     RegLog "", Err.Number, Err.Description
    Resume Next
End Function

Public Function PgDadosEmpresa(Id As Integer) As Dados_Empresa
    On Error Resume Next
    Dim Rst     As Recordset
    Dim strSQL  As String
    '15.05.2017
    'strSQL = "SELECT * FROM Empresas WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & Id
    
    strSQL = "SELECT * FROM Empresas WHERE ID = " & Id
    Set Rst = RegistroBuscar(strSQL)
    If Rst.BOF And Rst.EOF Then
            
        Else
            Rst.MoveFirst
            PgDadosEmpresa.Nome = Rst.Fields("Nome")
            PgDadosEmpresa.Fant = IIf(IsNull(Rst.Fields("fant")), "", Rst.Fields("fant"))
            PgDadosEmpresa.Lgr = IIf(IsNull(Rst.Fields("lgr")), "", Rst.Fields("lgr"))
            PgDadosEmpresa.Nro = IIf(IsNull(Rst.Fields("nro")), "", Rst.Fields("nro"))
            PgDadosEmpresa.Cpl = IIf(IsNull(Rst.Fields("cpl")), "", Rst.Fields("cpl"))
            PgDadosEmpresa.Bairro = IIf(IsNull(Rst.Fields("bairro")), "", Rst.Fields("bairro"))
            'PgDadosEmpresa.cmun = IIf(IsNull(Rst.Fields("cmun")), "", Rst.Fields("cmun"))
            PgDadosEmpresa.Mun = IIf(IsNull(Rst.Fields("mun")), "", Rst.Fields("mun"))
            PgDadosEmpresa.uf = IIf(IsNull(Rst.Fields("uf")), "", Rst.Fields("uf"))
            PgDadosEmpresa.CEP = IIf(IsNull(Rst.Fields("cep")), "", Rst.Fields("cep"))
            'PgDadosEmpresa.cpais = IIf(IsNull(Rst.Fields("cpais")), "", Rst.Fields("cpais"))
            PgDadosEmpresa.pais = IIf(IsNull(Rst.Fields("pais")), "", Rst.Fields("pais"))
            PgDadosEmpresa.Fone = IIf(IsNull(Rst.Fields("fone")), "", Rst.Fields("fone"))
            PgDadosEmpresa.Mail = IIf(IsNull(Rst.Fields("mail")), "", Rst.Fields("mail"))
            PgDadosEmpresa.CNPJ = IIf(IsNull(Rst.Fields("cnpj")), "", Rst.Fields("cnpj"))
            PgDadosEmpresa.IE = IIf(IsNull(Rst.Fields("ie")), "", Rst.Fields("ie"))
            PgDadosEmpresa.iest = IIf(IsNull(Rst.Fields("iest")), "", Rst.Fields("iest"))
            PgDadosEmpresa.im = IIf(IsNull(Rst.Fields("im")), "", Rst.Fields("im"))
            PgDadosEmpresa.cnae = IIf(IsNull(Rst.Fields("cnae")), "", Rst.Fields("cnae"))
            PgDadosEmpresa.RegimeTrib = IIf(IsNull(Rst.Fields("RegimeTrib")), "", Rst.Fields("RegimeTrib"))
            
            PgDadosEmpresa.Suframa = IIf(IsNull(Rst.Fields("Suframa")), "", Rst.Fields("Suframa"))
            PgDadosEmpresa.TipoAtividade = IIf(IsNull(Rst.Fields("TipoAtividade")), "", Left(Rst.Fields("TipoAtividade"), 1))
            
            PgDadosEmpresa.PISCST = IIf(IsNull(Rst.Fields("PISCST")), "0", Rst.Fields("PISCST"))
            PgDadosEmpresa.PISAliquota = IIf(IsNull(Rst.Fields("PISAliquota")), "0", Rst.Fields("PISAliquota"))
            PgDadosEmpresa.COFINSCST = IIf(IsNull(Rst.Fields("COFINSCST")), "0", Rst.Fields("COFINSCST"))
            PgDadosEmpresa.COFINSAliquota = IIf(IsNull(Rst.Fields("COFINSAliquota")), "0", Rst.Fields("COFINSAliquota"))
            
            PgDadosEmpresa.Logotipo = IIf(IsNull(Rst.Fields("fLogotipo")), "", Rst.Fields("fLogotipo"))
                   
            'Dados do Contador
            PgDadosEmpresa.cNome = cNull(Rst.Fields("ContNome"))
            PgDadosEmpresa.cCNPJ = cNull(Rst.Fields("ContCNPJ"))
            PgDadosEmpresa.cIE = cNull(Rst.Fields("ContIE"))
            PgDadosEmpresa.cIM = cNull(Rst.Fields("ContIM"))
            PgDadosEmpresa.cEndereco = cNull(Rst.Fields("ContEndereco"))
            PgDadosEmpresa.cNumero = cNull(Rst.Fields("ContNumero"))
            PgDadosEmpresa.cCompl = cNull(Rst.Fields("ContComplemento"))
            PgDadosEmpresa.cBairro = cNull(Rst.Fields("ContBairro"))
            PgDadosEmpresa.cUF = cNull(Rst.Fields("ContUF"))
            PgDadosEmpresa.cMunicipio = cNull(Rst.Fields("ContMunicipio"))
            PgDadosEmpresa.cCEP = cNull(Rst.Fields("ContCEP"))
            PgDadosEmpresa.cFone1 = cNull(Rst.Fields("ContFone1"))
            PgDadosEmpresa.cFone2 = cNull(Rst.Fields("ContFone2"))
            PgDadosEmpresa.cMail = cNull(Rst.Fields("ContMail"))
            PgDadosEmpresa.cCodID = cNull(Rst.Fields("ContCodigoID"))
            PgDadosEmpresa.crNome = cNull(Rst.Fields("ContRNome"))
            PgDadosEmpresa.crCPF = cNull(Rst.Fields("ContRCPF"))
            PgDadosEmpresa.crCRC = cNull(Rst.Fields("ContRCRC"))
            PgDadosEmpresa.crFone = cNull(Rst.Fields("ContRFone"))
            PgDadosEmpresa.crMail = cNull(Rst.Fields("ContRMail"))
    
    End If
    Rst.Close
    
End Function
Public Function PgDadosFornecedor(Id As Integer) As Dados_Fornecedor
    Dim Rst     As Recordset
    Dim strSQL  As String
    
    strSQL = "SELECT * FROM Fornecedores WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & Id
    Set Rst = RegistroBuscar(strSQL)
    If Rst.BOF And Rst.EOF Then
            
        Else
            Rst.MoveFirst
            PgDadosFornecedor.Nome = IIf(IsNull(Rst.Fields("xNome")), "", Rst.Fields("xNome"))
            PgDadosFornecedor.status = IIf(IsNull(Rst.Fields("status")), "", Rst.Fields("status"))
            PgDadosFornecedor.Pessoa = IIf(IsNull(Rst.Fields("pessoa")), "", Rst.Fields("pessoa"))
            PgDadosFornecedor.Fant = IIf(IsNull(Rst.Fields("fant")), "", Rst.Fields("fant"))
            PgDadosFornecedor.Doc = IIf(IsNull(Rst.Fields("doc")), "", Rst.Fields("doc"))
            PgDadosFornecedor.Obs = IIf(IsNull(Rst.Fields("Obs")), "", Rst.Fields("Obs"))
            PgDadosFornecedor.IE = IIf(IsNull(Rst.Fields("ie")), "", Rst.Fields("ie"))
            PgDadosFornecedor.iest = IIf(IsNull(Rst.Fields("iest")), "", Rst.Fields("iest"))
            PgDadosFornecedor.im = IIf(IsNull(Rst.Fields("im")), "", Rst.Fields("iM"))
            PgDadosFornecedor.cnae = IIf(IsNull(Rst.Fields("cnae")), "", Rst.Fields("cnae"))
            'PgDadosFornecedor.emailnfe = IIf(IsNull(Rst.Fields("emainfe")), "", Rst.Fields("emailnfe"))
            'PgDadosFornecedor.emalfin = IIf(IsNull(Rst.Fields("emaifin")), "", Rst.Fields("emailfin"))
            'PgDadosFornecedor.emailcom = IIf(IsNull(Rst.Fields("emaicom")), "", Rst.Fields("emailcom"))
            PgDadosFornecedor.website = IIf(IsNull(Rst.Fields("website")), "", Rst.Fields("website"))
            PgDadosFornecedor.LimiteCredito = IIf(IsNull(Rst.Fields("limitecredito")), "", Rst.Fields("limitecredito"))
            'PgDadosFornecedor.tipodocumento = IIf(IsNull(Rst.Fields("tipodoumento")), "", Rst.Fields("tipodocumento"))
            'PgDadosFornecedor.localcobranca = IIf(IsNull(Rst.Fields("localcobranca")), "", Rst.Fields("localcobranca"))
            PgDadosFornecedor.condicoespagamento = IIf(IsNull(Rst.Fields("condicoespagamento")), "", Rst.Fields("condicoespagamento"))
            'PgDadosFornecedor.emailn = IIf(IsNull(Rst.Fields("planocontas")), "", Rst.Fields("planocontas"))
            PgDadosFornecedor.Transportadora = IIf(IsNull(Rst.Fields("transportadora")), "", Rst.Fields("transportadora"))
            PgDadosFornecedor.cobrancalgr = IIf(IsNull(Rst.Fields("cobrancalgr")), "", Rst.Fields("cobrancalgr"))
            PgDadosFornecedor.cobrancanro = IIf(IsNull(Rst.Fields("cobrancanro")), "", Rst.Fields("cobrancanro"))
            PgDadosFornecedor.cobrancacpl = IIf(IsNull(Rst.Fields("cobrancacpl")), "", Rst.Fields("cobrancacpl"))
            PgDadosFornecedor.cobrancauf = IIf(IsNull(Rst.Fields("cobrancauf")), "", Rst.Fields("cobrancauf"))
            PgDadosFornecedor.cobrancamun = IIf(IsNull(Rst.Fields("cobrancamun")), "", Rst.Fields("cobrancamun"))
            PgDadosFornecedor.cobrancacep = IIf(IsNull(Rst.Fields("cobrancacep")), "", Rst.Fields("cobrancacep"))
            PgDadosFornecedor.entregalgr = IIf(IsNull(Rst.Fields("entregalgr")), "", Rst.Fields("entregalgr"))
            PgDadosFornecedor.entreganro = IIf(IsNull(Rst.Fields("entreganro")), "", Rst.Fields("entreganro"))
            PgDadosFornecedor.entregacpl = IIf(IsNull(Rst.Fields("entregacpl")), "", Rst.Fields("entregacpl"))
            PgDadosFornecedor.entregabairro = IIf(IsNull(Rst.Fields("entregabairro")), "", Rst.Fields("entregabairro"))
            PgDadosFornecedor.entregauf = IIf(IsNull(Rst.Fields("entregauf")), "", Rst.Fields("entregauf"))
            PgDadosFornecedor.entregamun = IIf(IsNull(Rst.Fields("entregamun")), "", Rst.Fields("entregamun"))
            PgDadosFornecedor.entregacep = IIf(IsNull(Rst.Fields("entregacep")), "", Rst.Fields("entregacep"))
            PgDadosFornecedor.entrega = IIf(IsNull(Rst.Fields("entrega")), "", Rst.Fields("entrega"))
            PgDadosFornecedor.cobranca = IIf(IsNull(Rst.Fields("cobranca")), "", Rst.Fields("cobranca"))
            PgDadosFornecedor.Lgr = IIf(IsNull(Rst.Fields("lgr")), "", Rst.Fields("lgr"))
            PgDadosFornecedor.Nro = IIf(IsNull(Rst.Fields("nro")), "", Rst.Fields("nro"))
            PgDadosFornecedor.Cpl = IIf(IsNull(Rst.Fields("cpl")), "", Rst.Fields("cpl"))
            PgDadosFornecedor.Bairro = IIf(IsNull(Rst.Fields("bairro")), "", Rst.Fields("bairro"))
            PgDadosFornecedor.uf = IIf(IsNull(Rst.Fields("uf")), "", Rst.Fields("uf"))
            PgDadosFornecedor.Mun = IIf(IsNull(Rst.Fields("mun")), "", Rst.Fields("mun"))
            PgDadosFornecedor.CEP = IIf(IsNull(Rst.Fields("cep")), "", Rst.Fields("cep"))
            PgDadosFornecedor.Mail = IIf(IsNull(Rst.Fields("mail")), "", Rst.Fields("mail"))
            PgDadosFornecedor.Fone = IIf(IsNull(Rst.Fields("fone")), "", Rst.Fields("fone"))
            
    End If
    Rst.Close
    
End Function
Public Function PgDadosRhFuncionario(Id As Integer) As Dados_RHFuncionario
    Dim Rst     As Recordset
    Dim strSQL  As String
    
    strSQL = "SELECT * FROM RhFuncionarioCadastro WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & Id
    Set Rst = RegistroBuscar(strSQL)
    If Rst.BOF And Rst.EOF Then
            
        Else
            Rst.MoveFirst
            PgDadosRhFuncionario.Nome = IIf(IsNull(Rst.Fields("xNome")), "", Rst.Fields("xNome"))
            PgDadosRhFuncionario.CPF = IIf(IsNull(Rst.Fields("cpf")), "", Rst.Fields("CPF"))
            PgDadosRhFuncionario.RG = IIf(IsNull(Rst.Fields("RG")), "", Rst.Fields("RG"))
            PgDadosRhFuncionario.Cargo = IIf(IsNull(Rst.Fields("Cargo")), "", Rst.Fields("Cargo"))
            PgDadosRhFuncionario.Endereco = IIf(IsNull(Rst.Fields("lgr")), "", Rst.Fields("lgr"))
            PgDadosRhFuncionario.Num = IIf(IsNull(Rst.Fields("nro")), "", Rst.Fields("nro"))
            PgDadosRhFuncionario.Compl = IIf(IsNull(Rst.Fields("Cpl")), "", Rst.Fields("Cpl"))
            PgDadosRhFuncionario.Bairro = IIf(IsNull(Rst.Fields("Bairro")), "", Rst.Fields("Bairro"))
            PgDadosRhFuncionario.uf = IIf(IsNull(Rst.Fields("UF")), "", Rst.Fields("UF"))
            PgDadosRhFuncionario.Municipio = IIf(IsNull(Rst.Fields("Mun")), "", Rst.Fields("Mun"))
            PgDadosRhFuncionario.CEP = IIf(IsNull(Rst.Fields("CEP")), "", Rst.Fields("CEP"))
            PgDadosRhFuncionario.Mail = IIf(IsNull(Rst.Fields("mail")), "", Rst.Fields("mail"))
            PgDadosRhFuncionario.Tel = IIf(IsNull(Rst.Fields("fone")), "", Rst.Fields("fone"))
            PgDadosRhFuncionario.Salario = IIf(IsNull(Rst.Fields("Salario")), "0", Rst.Fields("Salario"))
            PgDadosRhFuncionario.Comissao = IIf(IsNull(Rst.Fields("Comissao")), "0", Rst.Fields("Comissao"))
            PgDadosRhFuncionario.Assinatura = IIf(IsNull(Rst.Fields("Assinatura")), cNull(Rst.Fields("xNome")), Rst.Fields("Assinatura"))
            
    End If
    Rst.Close
    
End Function
Public Function PgDadosUsuario(Id As Integer) As Dados_Usuario
    On Error GoTo TrtErroUsu
    Dim Rst     As Recordset
    Dim strSQL  As String
    
    strSQL = "SELECT * FROM UsuGerenciador WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & Id
    Set Rst = RegistroBuscar(strSQL)

    If Rst.BOF And Rst.EOF Then
            
        Else
            Rst.MoveFirst
            PgDadosUsuario.Id = IIf(IsNull(Rst.Fields("id")), "0", Rst.Fields("id"))
            PgDadosUsuario.Login = IIf(IsNull(Rst.Fields("usu_Login")), "", Rst.Fields("usu_Login"))
            PgDadosUsuario.Nome = IIf(IsNull(Rst.Fields("usu_Nome")), "", Rst.Fields("usu_Nome"))
            PgDadosUsuario.senha = IIf(IsNull(Rst.Fields("usu_Senha")), "", Rst.Fields("usu_Senha"))
            
            PgDadosUsuario.idFunc = IIf(IsNull(Rst.Fields("usu_Nome")), "0", Left(Trim(Rst.Fields("usu_Nome")), 3))
            
            PgDadosUsuario.Grupo = IIf(IsNull(Rst.Fields("usu_grupo")), "0", Rst.Fields("usu_grupo"))
            PgDadosUsuario.SenhaNuncaExp = IIf(IsNull(Rst.Fields("usu_SenhaNuncaExpira")), "0", Rst.Fields("usu_SenhaNuncaExpira"))
            PgDadosUsuario.TrocarSenha = IIf(IsNull(Rst.Fields("usu_TrocaSenha")), "0", Rst.Fields("usu_TrocaSenha"))
            PgDadosUsuario.SuperUsuario = IIf(IsNull(Rst.Fields("usu_SuperUsuario")), "0", Rst.Fields("usu_SuperUsuario"))
            PgDadosUsuario.Menus = IIf(IsNull(Rst.Fields("usu_Menus")), "", Rst.Fields("usu_Menus"))
            
    End If
    Rst.Close
    Exit Function
TrtErroUsu:
    'MsgBox "Erro pgDadosUsuario." & vbCrLf & Err.Description, vbCritical, Err.Number
    RegLog "", Err.Number, Err.Description
    Resume Next
End Function

Public Function PgDadosCliente(Id As Integer) As Dados_Cliente
    On Error GoTo TrtErroCli
    Dim Rst     As Recordset
    Dim strSQL  As String
    
    strSQL = "SELECT * FROM Clientes WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & Id
    Set Rst = RegistroBuscar(strSQL)
    If Rst.BOF And Rst.EOF Then
            
        Else
            Rst.MoveFirst
            PgDadosCliente.Nome = IIf(IsNull(Rst.Fields("xnome")), "", Rst.Fields("xnome"))
            PgDadosCliente.status = IIf(IsNull(Rst.Fields("status")), "", Rst.Fields("status"))
            PgDadosCliente.Pessoa = IIf(IsNull(Rst.Fields("pessoa")), "", Rst.Fields("pessoa"))
            PgDadosCliente.Fant = IIf(IsNull(Rst.Fields("fant")), "", Rst.Fields("fant"))
            PgDadosCliente.Doc = IIf(IsNull(Rst.Fields("doc")), "", Rst.Fields("doc"))
            PgDadosCliente.Obs = IIf(IsNull(Rst.Fields("Obs")), "", Rst.Fields("Obs"))
            PgDadosCliente.IE = IIf(IsNull(Rst.Fields("ie")), "", Rst.Fields("ie"))
            PgDadosCliente.iest = IIf(IsNull(Rst.Fields("iest")), "", Rst.Fields("iest"))
            PgDadosCliente.im = IIf(IsNull(Rst.Fields("im")), "", Rst.Fields("iM"))
            PgDadosCliente.cnae = IIf(IsNull(Rst.Fields("cnae")), "", Rst.Fields("cnae"))
            PgDadosCliente.emailnfe = IIf(IsNull(Rst.Fields("emailnfe")), "", Rst.Fields("emailnfe"))
            PgDadosCliente.emailfin = IIf(IsNull(Rst.Fields("emailfin")), "", Rst.Fields("emailfin"))
            PgDadosCliente.emailcom = IIf(IsNull(Rst.Fields("emailcom")), "", Rst.Fields("emailcom"))
            PgDadosCliente.website = IIf(IsNull(Rst.Fields("website")), "", Rst.Fields("website"))
            PgDadosCliente.LimiteCredito = IIf(IsNull(Rst.Fields("limitecredito")), "0", Rst.Fields("limitecredito"))
            PgDadosCliente.TipoDocumento = IIf(IsNull(Rst.Fields("tipodocumento")), "0", Rst.Fields("tipodocumento"))
            'PgDadosCliente.localcobranca = IIf(IsNull(Rst.Fields("cobranca")), "", Rst.Fields("cobranca"))
            PgDadosCliente.condicoespagamento = IIf(IsNull(Rst.Fields("condicoespagamento")), "0", Rst.Fields("condicoespagamento"))
            PgDadosCliente.CentroCustos = IIf(IsNull(Rst.Fields("CentroCustos")), "0", Rst.Fields("CentroCustos"))
            
            PgDadosCliente.PlanoContas = IIf(IsNull(Rst.Fields("PlanoContas")), "0", Rst.Fields("PlanoContas"))
            
            PgDadosCliente.Transportadora = IIf(IsNull(Rst.Fields("transportadora")), "0", Rst.Fields("transportadora"))
            'PgDadosCliente.cobrancalgr = IIf(IsNull(Rst.Fields("cobrancalgr")), "", Rst.Fields("cobrancalgr"))
            'PgDadosCliente.cobrancanro = IIf(IsNull(Rst.Fields("cobrancanro")), "", Rst.Fields("cobrancanro"))
            'PgDadosCliente.cobrancacpl = IIf(IsNull(Rst.Fields("cobrancacpl")), "", Rst.Fields("cobrancacpl"))
            'PgDadosCliente.cobrancauf = IIf(IsNull(Rst.Fields("cobrancauf")), "", Rst.Fields("cobrancauf"))
            'PgDadosCliente.cobrancamun = IIf(IsNull(Rst.Fields("cobrancamun")), "", Rst.Fields("cobrancamun"))
            'PgDadosCliente.cobrancacep = IIf(IsNull(Rst.Fields("cobrancacep")), "", Rst.Fields("cobrancacep"))
            PgDadosCliente.ObsCobNfe = IIf(IsNull(Rst.Fields("ObsNFe")), "", Rst.Fields("ObsNFe"))
            PgDadosCliente.ObsCobBoleto = IIf(IsNull(Rst.Fields("ObsBoleto")), "", Rst.Fields("ObsBoleto"))
            
            PgDadosCliente.entregaDoc = IIf(IsNull(Rst.Fields("entregaDoc")), "", Rst.Fields("entregaDoc"))
            PgDadosCliente.entregalgr = IIf(IsNull(Rst.Fields("entregalgr")), "", Rst.Fields("entregalgr"))
            PgDadosCliente.entreganro = IIf(IsNull(Rst.Fields("entreganro")), "", Rst.Fields("entreganro"))
            PgDadosCliente.entregacpl = IIf(IsNull(Rst.Fields("entregacpl")), "", Rst.Fields("entregacpl"))
            PgDadosCliente.entregabairro = IIf(IsNull(Rst.Fields("entregabairro")), "", Rst.Fields("entregabairro"))
            PgDadosCliente.entregauf = IIf(IsNull(Rst.Fields("entregauf")), "", Rst.Fields("entregauf"))
            PgDadosCliente.entregamun = IIf(IsNull(Rst.Fields("entregamun")), "", Rst.Fields("entregamun"))
            PgDadosCliente.entregacep = IIf(IsNull(Rst.Fields("entregacep")), "", Rst.Fields("entregacep"))
            PgDadosCliente.entrega = IIf(IsNull(Rst.Fields("entrega")), "", Rst.Fields("entrega"))
            
            'PgDadosCliente.cobranca = IIf(IsNull(Rst.Fields("cobranca")), "", Rst.Fields("cobranca"))
            PgDadosCliente.Lgr = IIf(IsNull(Rst.Fields("xlgr")), "", Rst.Fields("xlgr"))
            PgDadosCliente.Nro = IIf(IsNull(Rst.Fields("nro")), "", Rst.Fields("nro"))
            PgDadosCliente.Cpl = IIf(IsNull(Rst.Fields("xcpl")), "", Rst.Fields("xcpl"))
            PgDadosCliente.Bairro = IIf(IsNull(Rst.Fields("xbairro")), "", Rst.Fields("xbairro"))
            PgDadosCliente.uf = IIf(IsNull(Rst.Fields("uf")), "", Rst.Fields("uf"))
            PgDadosCliente.Mun = IIf(IsNull(Rst.Fields("xmun")), "", Rst.Fields("xmun"))
            PgDadosCliente.CEP = IIf(IsNull(Rst.Fields("cep")), "", Rst.Fields("cep"))
            PgDadosCliente.Mail = IIf(IsNull(Rst.Fields("email")), "", Rst.Fields("email"))
            PgDadosCliente.Fone = IIf(IsNull(Rst.Fields("fone")), "", Rst.Fields("fone"))
            PgDadosCliente.Suframa = IIf(IsNull(Rst.Fields("SUFRAMA")), "", Rst.Fields("SUFRAMA"))
            PgDadosCliente.Vendedor = IIf(IsNull(Rst.Fields("Vendedor")), 0, Rst.Fields("Vendedor"))
            
    End If
    Rst.Close
    Exit Function
TrtErroCli:
    MsgBox "pgDadosCliente: " & Err.Description, vbCritical, "Num. " & Err.Number
    Resume Next
    
End Function

Public Function PgDadosMunicipio(sUF As String, sMun As String) As Dados_Municipio
    Dim Rst As Recordset
    Dim sSQL As String
    If Trim(sUF) = "" Then Exit Function
    
    sSQL = "SELECT * FROM TributacaoMunicipio WHERE codUF = " & pgDadosICMS(sUF, 0).codUF & " AND Descricao = '" & sMun & "' ORDER BY Descricao"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            PgDadosMunicipio.Id = Rst.Fields("Id")
            PgDadosMunicipio.uf = Rst.Fields("UF")
            PgDadosMunicipio.Descricao = Rst.Fields("Descricao")
            PgDadosMunicipio.codUF = Rst.Fields("coduf")
            PgDadosMunicipio.codMun = Rst.Fields("CodMun")
    End If
    Rst.Close
End Function

Public Sub ExibirDados(Formulario As Form, sSQL As String)
    On Error Resume Next
    Dim Rst     As Recordset
    Dim i As Integer
    Dim Controle As Control
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            
        Else
            Rst.MoveFirst

'**************************************
            For i = 0 To Formulario.Controls.Count - 1
                Set Controle = Formulario.Controls(i)
                
                If TypeOf Controle Is TextBox Then
                    Controle.Text = IIf(IsNull(Rst.Fields(Mid(Controle.Name, 4, Len(Controle.Name)))), "", Rst.Fields(Mid(Controle.Name, 4, Len(Controle.Name))))
                End If
                If TypeOf Controle Is ComboBox Then
                    Controle.Clear
                    Controle.AddItem IIf(IsNull(Rst.Fields(Mid(Controle.Name, 4, Len(Controle.Name)))), " ", Rst.Fields(Mid(Controle.Name, 4, Len(Controle.Name))))
                    Controle.Text = Controle.List(0)
                End If
                If TypeOf Controle Is CheckBox Then
                    Controle.Value = IIf(IsNull(Rst.Fields(Mid(Controle.Name, 4, Len(Controle.Name)))), 0, Rst.Fields(Mid(Controle.Name, 4, Len(Controle.Name))))
                End If
                If TypeOf Controle Is DTPicker Then
                    Controle.Value = IIf(IsNull(Rst.Fields(Mid(Controle.Name, 4, Len(Controle.Name)))), 0, Rst.Fields(Mid(Controle.Name, 4, Len(Controle.Name))))
                End If
            Next
    End If


End Sub
Public Function pgDescrFabricante(idFabricante As Integer) As String
    On Error Resume Next
    If Trim(idFabricante) = "" Then Exit Function
    Dim Rst As Recordset
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueFabricante WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & idFabricante)
    If Rst.BOF And Rst.EOF Then
            pgDescrFabricante = ""
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgDescrFabricante = Rst.Fields("Descricao")
            Rst.Close
            Exit Function
    End If

End Function
Public Function pgDescrGrupo(idGrupo As String) As String
    On Error Resume Next
    If Trim(idGrupo) = "" Then Exit Function
    Dim Rst As Recordset
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueGrupos WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & idGrupo)
    If Rst.BOF And Rst.EOF Then
            pgDescrGrupo = "<< Grupo não encontrado >>"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgDescrGrupo = Rst.Fields("Descricao")
            Rst.Close
            Exit Function
    End If
End Function
Public Function pgIdPais(strPais As String) As Integer
    Dim Rst As Recordset
    Set Rst = RegistroBuscar("SELECT * FROM TributacaoPais WHERE pais = '" & strPais & "'")
    If Rst.BOF And Rst.EOF Then
            pgIdPais = "0"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgIdPais = Rst.Fields("codigo")
            Rst.Close
            Exit Function
    End If
End Function
Public Function pgDescrSubGrupo(IdSubGrupo As String) As String
    On Error Resume Next
    Dim Rst As Recordset
    If Trim(IdSubGrupo) = "" Then Exit Function
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueSubGrupo WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdSubGrupo)
    If Rst.BOF And Rst.EOF Then
            pgDescrSubGrupo = "<< SubGrupo não encontrado >>"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgDescrSubGrupo = Rst.Fields("Descricao")
            Rst.Close
            Exit Function
    End If
End Function
'Public Function pgDescrCST(IdCST As String, Tabela As String) As String
'    'tabela - A - Origem
'    '            B - CST do ICMS
'    '            I - CST do IPI
'    If Trim(IdCST) = "" Then Exit Function
'    Dim Rst As Recordset
'    Set Rst = RegistroBuscar("SELECT * FROM TributacaoCST WHERE tabela = '" & Tabela & "' AND cst = " & IdCST)
'    If Rst.BOF And Rst.EOF Then
'            pgDescrCST = "<< ICMSOrigem não encontrado >>"
'            Rst.Close
'            Exit Function
'        Else
'            Rst.MoveFirst
'            pgDescrCST = Rst.Fields("descricao")
'            Rst.Close
'            Exit Function
'    End If
'End Function
Public Function PgDadosCST(idCST As String, Tabela As String) As Dados_CST
    'Tabela
    '    Origem = A
    '    COFINS = P
    '    PIS = P
    '    IPI = I
    '   ICMS = B
    Dim sTab As String
    
    
    Select Case UCase(Tabela)
        Case "ORIGEM"
            sTab = "A"
            If Len(idCST) > 1 Then
                MsgBox "Codigo de origem invalido!", vbInformation, "Aviso"
                Exit Function
            End If
        Case "ICMS"
            If PgDadosEmpresa(ID_Empresa).RegimeTrib = 1 Then
                    sTab = "C"
                Else
                    sTab = "B"
            End If
            'idCST = Left("00", 2 - Len(idCST)) & idCST
        Case "COFINS", "PIS"
            sTab = "P"
            idCST = Left("00", 2 - Len(idCST)) & idCST
        Case "IPI"
            sTab = "I"
            idCST = Left("00", 2 - Len(idCST)) & idCST
        Case Else
            'MsgBox "Erro ao localizar tabela CST"
            PgDadosCST.Id = 0
            PgDadosCST.Tabela = ""
            PgDadosCST.cst = ""
            PgDadosCST.Descricao = ""
            Exit Function
    End Select
    Dim sSQL    As String
    Dim Rst     As Recordset
    'cboCSTPIS.Clear
    'sSQL = "SELECT * FROM TributacaoCST WHERE ID_Empresa = " & ID_Empresa & " AND CST = '" & idCST & "' AND Tabela = '" & sTab & "'"
    sSQL = "SELECT * FROM TributacaoCST WHERE CST = '" & idCST & "' AND Tabela = '" & sTab & "'"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            PgDadosCST.Id = Rst.Fields("id")
            PgDadosCST.Tabela = Rst.Fields("Tabela")
            PgDadosCST.cst = Rst.Fields("CST")
            PgDadosCST.Descricao = Rst.Fields("Descricao")
    End If
    Rst.Close
End Function


Public Function pgDescrDeposito(IdDep As Integer) As String
    On Error Resume Next
    Dim Rst As Recordset
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueDeposito WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdDep)
    If Rst.BOF And Rst.EOF Then
            pgDescrDeposito = "<< Deposito não encontrado >>"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgDescrDeposito = Rst.Fields("Descricao")
            Rst.Close
            Exit Function
    End If
End Function
Public Function pgDescrCondPag(IdCondPag As String) As String
    Dim Rst As Recordset
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCondicoesPagamento WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdCondPag)
    If Rst.BOF And Rst.EOF Then
            pgDescrCondPag = "<< CondPag não encontrado >>"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgDescrCondPag = Rst.Fields("Descricao")
            Rst.Close
            Exit Function
    End If
End Function
Public Function pgDescrTipoDoc(IdTipoDoc As String) As String
    Dim Rst As Recordset
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdTipoDoc)
    If Rst.BOF And Rst.EOF Then
            pgDescrTipoDoc = "<< TipoDoc não encontrado >>"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgDescrTipoDoc = Rst.Fields("Descricao")
            Rst.Close
            Exit Function
    End If
End Function
Public Function pgIdTipoDoc(descDoc As String) As Integer
    Dim Rst As Recordset
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento WHERE ID_Empresa = " & ID_Empresa & " AND Descricao = '" & descDoc & "'")
    If Rst.BOF And Rst.EOF Then
            pgIdTipoDoc = 0
        Else
            Rst.MoveFirst
            pgIdTipoDoc = Rst.Fields("Id")
    End If
    Rst.Close
End Function

Public Function pgIdcCusto(descCC As String) As Integer
    Dim Rst As Recordset
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCentroCustos WHERE ID_Empresa = " & ID_Empresa & " AND Descricao = '" & descCC & "'")
    If Rst.BOF And Rst.EOF Then
            pgIdcCusto = 0
        Else
            Rst.MoveFirst
            pgIdcCusto = Rst.Fields("Id")
    End If
    Rst.Close
End Function

Public Function pgMovEst(IdMov As Integer) As Dados_MovEstoque
    Dim Rst As Recordset
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueMovimento WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdMov)
    If Rst.BOF And Rst.EOF Then
            pgMovEst.Descricao = "<< Movimento não encontrado >>"
            MsgBox "Erro ao localizar dados do Movimento"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            
            pgMovEst.Id = IdMov
            pgMovEst.Descricao = Rst.Fields("Descricao")
            pgMovEst.Sigla = Rst.Fields("Sigla")
            Select Case Trim(Rst.Fields("Acao"))
                Case "SOMAR (+)"
                    pgMovEst.acao = "+"
                    pgMovEst.AcaoDescr = "SOMAR (+)"
                Case "SUBTRAIR (-)"
                    pgMovEst.acao = "-"
                    pgMovEst.AcaoDescr = "SUBTRAIR (-)"
                Case "NENHUM"
                    'Fara somente registro
                    pgMovEst.acao = "N"
                    pgMovEst.AcaoDescr = "NENHUM"
                Case Else
                    MsgBox "Erro ao localizar movimento Estoque"
            End Select
            Rst.Close
            Exit Function
    End If
End Function
Public Function pgDadosEstoqueProduto(IdEstoqueProduto As Long) As Dados_EstoqueProduto
    On Error GoTo NotificarErro:
    Dim Rst As Recordset
    Dim sQL As String
    '06.02.2017
    'sql = "SELECT * FROM EstoqueProduto WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND Id = " & IdEstoqueProduto
    sQL = "SELECT * FROM EstoqueProduto WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdEstoqueProduto
    Set Rst = RegistroBuscar(sQL)
    If Rst.BOF And Rst.EOF Then
            'pgMovEst.Descricao = "<< Movimento não encontrado >>"
            'MsgBox "Erro ao localizar Dados: pgDadosEstoqueProduto!" & vbCrLf & " Deposito: " & ID_Deposito & vbCrLf & " ID Produto: " & IdEstoqueProduto
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            
            pgDadosEstoqueProduto.Id = IdEstoqueProduto 'Left(String(3, "0"), 3 - Len(Trim(IdEstoqueProduto))) & IdEstoqueProduto
            
            pgDadosEstoqueProduto.IdDeposito = Rst.Fields("Deposito")
            pgDadosEstoqueProduto.Referencia = IIf(IsNull(Rst.Fields("Referencia")), "", Rst.Fields("Referencia"))
            pgDadosEstoqueProduto.status = IIf(IsNull(Rst.Fields("Status")), "", Rst.Fields("Status"))
            pgDadosEstoqueProduto.CodBarras = IIf(IsNull(Rst.Fields("CodigoBarras")), "", Rst.Fields("CodigoBarras"))
            pgDadosEstoqueProduto.Descricao = IIf(IsNull(Rst.Fields("Descricao")), "", Rst.Fields("Descricao"))
            pgDadosEstoqueProduto.Grupo = IIf(IsNull(Rst.Fields("Grupo")), "", Rst.Fields("Grupo"))
            pgDadosEstoqueProduto.subGrupo = IIf(IsNull(Rst.Fields("subGrupo")), "", Rst.Fields("subGrupo"))
            pgDadosEstoqueProduto.NCM = IIf(IsNull(Rst.Fields("NCM")), "", Rst.Fields("NCM"))
            pgDadosEstoqueProduto.MVA = IIf(IsNull(Rst.Fields("MVA")), "", Rst.Fields("MVA"))
            pgDadosEstoqueProduto.ICMSOrigem = IIf(IsNull(Rst.Fields("ICMSOrigem")), "", Rst.Fields("ICMSOrigem"))
            pgDadosEstoqueProduto.ICMSCST = IIf(IsNull(Rst.Fields("ICMSCST")), "", Rst.Fields("ICMSCST"))
            pgDadosEstoqueProduto.IPIAliquota = IIf(IsNull(Rst.Fields("IPIAliquota")), "", Rst.Fields("IPIAliquota"))
            pgDadosEstoqueProduto.IPICST = IIf(IsNull(Rst.Fields("IPICST")), "", Rst.Fields("IPICST"))
            pgDadosEstoqueProduto.Enquadramento = IIf(IsNull(Rst.Fields("IPICodEnquadramento")), "999", Rst.Fields("IPICodEnquadramento"))
            pgDadosEstoqueProduto.Unidade = IIf(IsNull(Rst.Fields("Unidade")), "", Rst.Fields("Unidade"))
            
            pgDadosEstoqueProduto.Saldo = IIf(IsNull(Rst.Fields("saldo")), "0", Rst.Fields("Saldo"))
            
            pgDadosEstoqueProduto.VlCusto = IIf(IsNull(Rst.Fields("Custo")), "", Rst.Fields("Custo"))
            pgDadosEstoqueProduto.VlIPI = IIf(IsNull(Rst.Fields("VlIPI")), "", Rst.Fields("VlIPI"))
            pgDadosEstoqueProduto.VlOutros = IIf(IsNull(Rst.Fields("Outros")), "", Rst.Fields("Outros"))
            pgDadosEstoqueProduto.MarkUp = IIf(IsNull(Rst.Fields("Markup")), "", Rst.Fields("Markup"))
            pgDadosEstoqueProduto.VlTabela = IIf(IsNull(Rst.Fields("Preco")), "", Rst.Fields("Preco"))
            pgDadosEstoqueProduto.InfCompl = IIf(IsNull(Rst.Fields("InformacoesComplementares")), "", Rst.Fields("InformacoesComplementares"))
            
            Rst.Close
            Exit Function
    End If
    Exit Function
NotificarErro:
    RegLog "0", "0", "[pgDadosEstoqueProduto] - " & Err.Number & " - " & Err.Description
    Resume Next
End Function
Public Function pgDadosBanco(IdBanco As String) As Dados_Banco '18.08.17
    Dim Rst As Recordset
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroBancoCadastro WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdBanco)
    If Rst.BOF And Rst.EOF Then
            'pgMovEst.Descricao = "<< Movimento não encontrado >>"
            MsgBox "Erro ao localizar Dados: pgDadosBanco idBanco: " & IdBanco
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgDadosBanco.Id = Left(String(3, "0"), 3 - Len(Trim(IdBanco))) & IdBanco
            pgDadosBanco.Nome = Rst.Fields("Nome")
            pgDadosBanco.Numero = IIf(IsNull(Rst.Fields("Numero")) = True, "0", Rst.Fields("Numero"))
            Rst.Close
            Exit Function
    End If
End Function
Public Function pgDadosConta(idConta As Integer) As Dados_Conta
    On Error GoTo regErro
    Dim Rst As Recordset
    If idConta = 0 Then Exit Function
        
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroConta WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & idConta)
    If Rst.BOF And Rst.EOF Then
            'pgMovEst.Descricao = "<< Movimento não encontrado >>"
            MsgBox "Erro ao localizar Dados: pgDadosConta"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgDadosConta.Id = ZE(idConta, 3)
            pgDadosConta.banco = Rst.Fields("Banco")
            pgDadosConta.agencia = IIf(IsNull(Rst.Fields("Agencia")), "", Rst.Fields("Agencia"))
            pgDadosConta.AgenciaDV = IIf(IsNull(Rst.Fields("AgenciaDV")), "", Rst.Fields("AgenciaDV"))
            pgDadosConta.conta = IIf(IsNull(Rst.Fields("Conta")), "", Rst.Fields("Conta"))
            pgDadosConta.ContaDV = IIf(IsNull(Rst.Fields("ContaDV")), "", Rst.Fields("ContaDV"))
            pgDadosConta.Multa = IIf(IsNull(Rst.Fields("Multa")), "0", Rst.Fields("Multa"))
            pgDadosConta.Juros = IIf(IsNull(Rst.Fields("Juros")), "0", Rst.Fields("Juros"))
            pgDadosConta.DiasProtesto = IIf(IsNull(Rst.Fields("DiasProtesto")), "0", Rst.Fields("DiasProtesto"))
            
            pgDadosConta.Contrato = IIf(IsNull(Rst.Fields("Contrato")), "", Rst.Fields("Contrato"))
            pgDadosConta.carteira = IIf(IsNull(Rst.Fields("Carteira")), "0", Rst.Fields("Carteira"))
            pgDadosConta.Variacao = IIf(IsNull(Rst.Fields("Variacao")), "", Rst.Fields("Variacao"))
            pgDadosConta.Convenio = IIf(IsNull(Rst.Fields("Convenio")), "", Rst.Fields("Convenio"))
            pgDadosConta.ConvenioLider = IIf(IsNull(Rst.Fields("ConvenioLider")), "", Rst.Fields("ConvenioLider"))
            pgDadosConta.Tipo = IIf(IsNull(Rst.Fields("Tipo")), "", Rst.Fields("Tipo"))
            
            pgDadosConta.Saldo = ChkVal(IIf(IsNull(Rst.Fields("Saldo")), "0", Rst.Fields("Saldo")), 0, cDecMoeda)
            Rst.Close
            Exit Function
    End If
    Exit Function
regErro:
    RegLog "0", "0", "PgDadosConta: " & Err.Number & " - " & Err.Description
    Resume Next
End Function
Public Function pgDadosTransportadora(IdTransportadora As Integer) As Dados_Transportadora
    Dim Rst As Recordset
    If IdTransportadora = 0 Then Exit Function
    Set Rst = RegistroBuscar("SELECT * FROM Transportadoras WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdTransportadora)
    If Rst.BOF And Rst.EOF Then
            'pgMovEst.Descricao = "<< Movimento não encontrado >>"
            'MsgBox "Erro ao localizar Dados: pgDadosTransportadora"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            'pgDadosTransportadora.id = Left(String(3, "0"), 3 - Len(Trim(IdTransportadora))) & IdTransportadora
            pgDadosTransportadora.Pessoa = IIf(IsNull(Rst.Fields("Pessoa")), "", Rst.Fields("Pessoa"))
            pgDadosTransportadora.Nome = Rst.Fields("xNome")
            pgDadosTransportadora.Fant = IIf(IsNull(Rst.Fields("Fant")), "", Rst.Fields("Fant"))
            pgDadosTransportadora.CNPJ = IIf(IsNull(Rst.Fields("CNPJ")), "", Rst.Fields("CNPJ"))
            pgDadosTransportadora.IE = IIf(IsNull(Rst.Fields("IE")), "", Rst.Fields("IE"))
            pgDadosTransportadora.Lgr = IIf(IsNull(Rst.Fields("xEnder")), "", Rst.Fields("xEnder"))
            'pgDadosTransportadora.Nro = IIf(IsNull(Rst.Fields("nro")), "", Rst.Fields("nro"))
            'pgDadosTransportadora.Cpl = IIf(IsNull(Rst.Fields("cpl")), "", Rst.Fields("cpl"))
            pgDadosTransportadora.Bairro = IIf(IsNull(Rst.Fields("Bairro")), "", Rst.Fields("Bairro"))
            pgDadosTransportadora.Mun = IIf(IsNull(Rst.Fields("Mun")), "", Rst.Fields("Mun"))
            pgDadosTransportadora.uf = IIf(IsNull(Rst.Fields("UF")), "", Rst.Fields("UF"))
            pgDadosTransportadora.CEP = IIf(IsNull(Rst.Fields("CEP")), "", Rst.Fields("CEP"))
            pgDadosTransportadora.Mail = IIf(IsNull(Rst.Fields("mail")), "", Rst.Fields("mail"))
            pgDadosTransportadora.Fone = IIf(IsNull(Rst.Fields("fone")), "", Rst.Fields("fone"))
            Rst.Close
            Exit Function
    End If
End Function

Public Function pgDadosCentroCustos(idCentroCustos As Integer) As Dados_CentroCustos
    Dim Rst As Recordset
    If idCentroCustos = 0 Then Exit Function
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCentroCustos WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & idCentroCustos)
    If Rst.BOF And Rst.EOF Then
            'pgMovEst.Descricao = "<< Movimento não encontrado >>"
            pgDadosCentroCustos.Id = 0
            MsgBox "Erro ao localizar Dados:pgDadosCentroCustos"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgDadosCentroCustos.Id = Left(String(3, "0"), 3 - Len(Trim(idCentroCustos))) & idCentroCustos
            pgDadosCentroCustos.Descricao = Rst.Fields("descricao")
            pgDadosCentroCustos.Sigla = Rst.Fields("Sigla")
            Rst.Close
            Exit Function
    End If
End Function
Public Function pgDadosTipoDocumento(IdTipoDocumento As Integer) As Dados_TipoDocumento
   On Error GoTo TrtErro
    Dim Rst As Recordset
    If IdTipoDocumento = 0 Then Exit Function
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdTipoDocumento)
    If Rst.BOF And Rst.EOF Then
            'pgMovEst.Descricao = "<< Movimento não encontrado >>"
            pgDadosTipoDocumento.Id = 0
            MsgBox "Modulo: pgDadosTipoDocumento" & vbCrLf & " (IdTipoDocumento=" & IdTipoDocumento & ")"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgDadosTipoDocumento.Id = Left(String(3, "0"), 3 - Len(Trim(IdTipoDocumento))) & IdTipoDocumento
            pgDadosTipoDocumento.Descricao = Rst.Fields("descricao")
            pgDadosTipoDocumento.Sigla = Rst.Fields("Sigla")
            pgDadosTipoDocumento.Impressao = Rst.Fields("Impressao")
            pgDadosTipoDocumento.Tipo = IIf(IsNull(Rst.Fields("Tipo")), "0", Rst.Fields("Tipo"))
            pgDadosTipoDocumento.formaPgto = IIf(Len(Trim(cNull(Rst.Fields("formapgto")))) = 0, "99  - Outros", Rst.Fields("formapgto"))
            Rst.Close
            Exit Function
    End If
    Exit Function
TrtErro:
pgDadosTipoDocumento.Id = 0
    RegLog "pgDadosTipoDocumento", "", Err.Number & " - " & Err.Description
End Function

Public Function pgDadosICMS(sBusca As String, tpBusca As Integer) As Dados_ICMS
    On Error Resume Next
    Dim Rst     As Recordset
    Dim sSQL    As String
    Select Case tpBusca
        Case 0 'Busca por sigla
            sSQL = "SELECT * FROM TributacaoUF WHERE sigla = '" & sBusca & "' ORDER BY sigla"
        Case 1 'Busca por CodUF
            sSQL = "SELECT * FROM TributacaoUF WHERE codUF = " & sBusca & " ORDER BY sigla"
        Case 2 'Busca po ID
            sSQL = "SELECT * FROM TributacaoUF WHERE Id = " & sBusca
    End Select
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'pgMovEst.Descricao = "<< Movimento não encontrado >>"
            MsgBox "Erro ao localizar Dados: pgDadosICMS"
        Else
            Rst.MoveFirst
            pgDadosICMS.Id = Rst.Fields("id") 'Left(String(3, "0"), 3 - Len(Trim(IdICMS))) & IdICMS
            pgDadosICMS.Descricao = IIf(IsNull(Rst.Fields("descricao")), "", Rst.Fields("descricao"))
            pgDadosICMS.Sigla = IIf(IsNull(Rst.Fields("Sigla")), "", Rst.Fields("Sigla"))
            pgDadosICMS.ICMS = IIf(IsNull(Rst.Fields("ICMS")), 0, Rst.Fields("ICMS"))
            pgDadosICMS.ICMSInt = IIf(IsNull(Rst.Fields("ICMSInt")), 0, Rst.Fields("ICMSInt"))
            
            pgDadosICMS.ICMSFECP = IIf(IsNull(Rst.Fields("ICMSFECP")), 0, Rst.Fields("ICMSFECP"))
            
            
            pgDadosICMS.codUF = IIf(IsNull(Rst.Fields("codUF")), 0, Rst.Fields("codUF"))
    End If
    Rst.Close
End Function
'Public Function PgDadosUF(sUF As String, Optional TpBusca As Integer) As Dados_UF
'    'tpBusca - Tipo de Busca a ser efetuada Default 0
'    '   0 - Busca pela Sigla
'    '   1 - Busca pelo CodUF
'    Dim Rst As Recordset
'    Dim sSQL As String
'    If TpBusca = 0 Then
'            sSQL = "SELECT * FROM TributacaoUF WHERE ID_Empresa = " & ID_Empresa & " AND sigla = '" & sUF & "' ORDER BY sigla"
'        Else
'            sSQL = "SELECT * FROM TributacaoUF WHERE ID_Empresa = " & ID_Empresa & " AND codUF = " & sUF & " ORDER BY sigla"
'    End If
'    Set Rst = RegistroBuscar(sSQL)
'    If Rst.BOF And Rst.EOF Then
'        Else
'            Rst.MoveFirst
'            PgDadosUF.Id = Rst.Fields("codUF")
'            PgDadosUF.Descricao = Rst.Fields("Descricao")
'            PgDadosUF.UF = Rst.Fields("Sigla")
'    End If
'    Rst.Close
'End Function


Public Function PgDadosFinanceiroFatura(Id As Long) As Dados_FinanceiroFatura
    'Pega as fatura pertinentes a emissao de documento fiscal
    'Nao ha vinculo com FaturamentoNFE
    On Error GoTo regErro
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    sSQL = "SELECT * FROM FinanceiroContasPRCadastro WHERE ID_Empresa = " & ID_Empresa & " AND id = " & Id
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar fatura n." & Id, vbInformation, "Aviso"
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
        
    End If
    
    PgDadosFinanceiroFatura.Id = Rst.Fields("Id")
    PgDadosFinanceiroFatura.ContaPR = IIf(IsNull(Rst.Fields("ContaPR")), "", Rst.Fields("ContaPR"))
    PgDadosFinanceiroFatura.TpConta = IIf(IsNull(Rst.Fields("TpDocumento")), "0", Rst.Fields("TpDocumento"))
    PgDadosFinanceiroFatura.emissao = IIf(IsNull(Rst.Fields("Emissao")), Empty, Rst.Fields("Emissao"))
    PgDadosFinanceiroFatura.NumFatura = IIf(IsNull(Rst.Fields("NumFatura")), "", Rst.Fields("NumFatura"))
    PgDadosFinanceiroFatura.vlFatura = ChkVal(IIf(IsNull(Rst.Fields("vlFatura")), "0", Rst.Fields("vlFatura")), 0, cDecMoeda)
    PgDadosFinanceiroFatura.idConta = IIf(IsNull(Rst.Fields("Conta")), "0", Rst.Fields("Conta"))
    
    PgDadosFinanceiroFatura.idCentroCustos = IIf(IsNull(Rst.Fields("CentroCusto")), "0", Rst.Fields("CentroCusto"))
    PgDadosFinanceiroFatura.idPlanoContas = IIf(IsNull(Rst.Fields("PlanoContas")), "0", Rst.Fields("PlanoContas"))
    
    
    PgDadosFinanceiroFatura.idTpDoc = IIf(IsNull(Rst.Fields("TpDocumento")), "0", Rst.Fields("TpDocumento"))
    PgDadosFinanceiroFatura.Tabela = IIf(IsNull(Rst.Fields("Tabela")), "", Rst.Fields("Tabela"))
    PgDadosFinanceiroFatura.IDSacado = IIf(IsNull(Rst.Fields("idSacado")), "", Rst.Fields("idSacado"))
    PgDadosFinanceiroFatura.CNPJSacado = IIf(IsNull(Rst.Fields("CNPJ")), "", Rst.Fields("CNPJ"))
    PgDadosFinanceiroFatura.Sacado = IIf(IsNull(Rst.Fields("Nome")), "", Rst.Fields("Nome"))
    PgDadosFinanceiroFatura.CodigoBarras = IIf(IsNull(Rst.Fields("CodigoBarras")), "", Rst.Fields("CodigoBarras"))
    PgDadosFinanceiroFatura.LinhaDigitavel = IIf(IsNull(Rst.Fields("LinhaDigitavel")), "", Rst.Fields("LinhaDigitavel"))
    PgDadosFinanceiroFatura.NossoNumero = IIf(IsNull(Rst.Fields("NossoNumero")), "", Rst.Fields("NossoNumero"))
    PgDadosFinanceiroFatura.Vencimento = IIf(IsNull(Rst.Fields("Vencimento")), Empty, Rst.Fields("Vencimento"))
    PgDadosFinanceiroFatura.NumDuplicata = IIf(IsNull(Rst.Fields("NumDuplicata")), "", Rst.Fields("NumDuplicata"))
    PgDadosFinanceiroFatura.vlDuplicata = ChkVal(IIf(IsNull(Rst.Fields("vlDuplicata")), "0", Rst.Fields("vlDuplicata")), 0, cDecMoeda)
    PgDadosFinanceiroFatura.Multa = IIf(IsNull(Rst.Fields("Multa")), "", Rst.Fields("Multa"))
    PgDadosFinanceiroFatura.MultaMora = IIf(IsNull(Rst.Fields("MultaMora")), "", Rst.Fields("MultaMora"))
    PgDadosFinanceiroFatura.Juros = IIf(IsNull(Rst.Fields("Juros")), "", Rst.Fields("Juros"))
    PgDadosFinanceiroFatura.IdBanco = IIf(IsNull(Rst.Fields("idBanco")), 0, Rst.Fields("idBanco"))
    PgDadosFinanceiroFatura.Acrescimo = IIf(IsNull(Rst.Fields("Acrescimo")), "", Rst.Fields("Acrescimo"))
    PgDadosFinanceiroFatura.Abatimento = IIf(IsNull(Rst.Fields("Abatimento")), "0", Rst.Fields("Abatimento"))
    PgDadosFinanceiroFatura.Deducoes = IIf(IsNull(Rst.Fields("Deducoes")), "", Rst.Fields("Deducoes"))
    PgDadosFinanceiroFatura.DiasProtesto = IIf(IsNull(Rst.Fields("DiasProtesto")), "0", Rst.Fields("DiasProtesto"))
    PgDadosFinanceiroFatura.vlCobrado = IIf(IsNull(Rst.Fields("vlCobrado")), "", Rst.Fields("vlCobrado"))
    PgDadosFinanceiroFatura.DataQuitacao = IIf(IsNull(Rst.Fields("DataQuitacao")), Empty, Rst.Fields("DataQuitacao"))
    PgDadosFinanceiroFatura.Obs = IIf(IsNull(Rst.Fields("Obs")), "", Rst.Fields("Obs"))
    
    PgDadosFinanceiroFatura.ObsBol1 = cNull(Rst.Fields("ObsBol1"))
    PgDadosFinanceiroFatura.ObsBol2 = cNull(Rst.Fields("ObsBol2"))
    PgDadosFinanceiroFatura.ObsBol3 = cNull(Rst.Fields("ObsBol3"))
    

    Rst.Close
    Exit Function
regErro:
    RegLog "0", "0", "PgDadosFinanceiroFatura: " & Err.Number & "-" & Err.Description & "(" & Id & ")"
    Resume Next
End Function
Public Function PgDadosTpNotaFiscal(Id As Integer) As Dados_TipoNotaFiscal
    'On Error GoTo TrtErroTpNF
    Dim Rst As Recordset
    If Id = 0 Then Exit Function
    Set Rst = RegistroBuscar("SELECT * FROM FaturamentoTipoNotaFiscal WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & Id)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar Tipo de NOTA Fiscal n. " & Id
        Else
            Rst.MoveFirst
            PgDadosTpNotaFiscal.Descricao = IIf(IsNull(Rst.Fields("Descricao")), "", Rst.Fields("Descricao"))
            PgDadosTpNotaFiscal.TipoNota = IIf(IsNull(Rst.Fields("TipoNota")), "", Rst.Fields("TipoNota"))
            PgDadosTpNotaFiscal.TipoNotaDescr = IIf(Rst.Fields("TipoNota") = 0, "Entrada", "Saída")
            
            PgDadosTpNotaFiscal.Serie = IIf(IsNull(Rst.Fields("Serie")), "", Rst.Fields("Serie"))
            PgDadosTpNotaFiscal.Modelo = IIf(IsNull(Rst.Fields("Modelo")), "", Rst.Fields("Modelo"))
            PgDadosTpNotaFiscal.NumInicial = IIf(IsNull(Rst.Fields("NumInicial")), "0", Rst.Fields("numInicial"))
            PgDadosTpNotaFiscal.Natureza = IIf(IsNull(Rst.Fields("Naturezaoperacao")), "", Rst.Fields("Naturezaoperacao"))
            PgDadosTpNotaFiscal.Finalidade = IIf(IsNull(Rst.Fields("Finalidade")), "", Rst.Fields("Finalidade"))
            
            Select Case IIf(IsNull(Rst.Fields("Finalidade")), "", Rst.Fields("Finalidade"))
                Case 1
                    PgDadosTpNotaFiscal.FinalidadeDescr = "NF-e normal"
                Case 2
                    PgDadosTpNotaFiscal.FinalidadeDescr = "NF-e Complementar"
                Case 3
                    PgDadosTpNotaFiscal.FinalidadeDescr = "NF-e de ajuste"
                Case 4
                    PgDadosTpNotaFiscal.FinalidadeDescr = "NF-e de devolução"
                Case Else
                    PgDadosTpNotaFiscal.FinalidadeDescr = "IDENTIFICAÇÃO NÃO ENCONTRADA"
            End Select

            
            
            PgDadosTpNotaFiscal.EnvioRF = IIf(IsNull(Rst.Fields("EnvioRF")), "0", Rst.Fields("EnvioRF"))
            PgDadosTpNotaFiscal.ChaveAcessoRef = IIf(IsNull(Rst.Fields("ChaveAcessoRef")), "0", Rst.Fields("ChaveAcessoRef"))
            
            'PgDadosTpNotaFiscal.TipoEmissaoDescr = PgDescrTipoEmissao(IIf(IsNull(Rst.Fields("tpEmis")), "0", Rst.Fields("tpEmis")))
            PgDadosTpNotaFiscal.MovFisco = IIf(IsNull(Rst.Fields("MovFisco")), "0", Rst.Fields("MovFisco"))
            PgDadosTpNotaFiscal.MovComissao = IIf(IsNull(Rst.Fields("MovComissao")), "0", Rst.Fields("MovComissao"))
            PgDadosTpNotaFiscal.MovEstoque = IIf(IsNull(Rst.Fields("MovEstoque")), "0", Rst.Fields("MovEstoque"))
            'PgDadosTpNotaFiscal.MovFisco = IIf(IsNull(Rst.Fields("MovFisco")), "0", Rst.Fields("MovFisco"))
            PgDadosTpNotaFiscal.MovContasPR = IIf(IsNull(Rst.Fields("MovContasPR")), "0", Rst.Fields("MovContasPR"))
            PgDadosTpNotaFiscal.ModBC = IIf(IsNull(Rst.Fields("ModBC")), "0", Rst.Fields("ModBC"))
            PgDadosTpNotaFiscal.ModBCST = IIf(IsNull(Rst.Fields("ModBCST")), "0", Rst.Fields("ModBCST"))
            PgDadosTpNotaFiscal.ImpCmpFatura = IIf(IsNull(Rst.Fields("ImprCampoFatura")), "0", Rst.Fields("ImprCampoFatura"))
            PgDadosTpNotaFiscal.ImpDtSaida = IIf(IsNull(Rst.Fields("ImprDataSaida")), "0", Rst.Fields("ImprDataSaida"))
            PgDadosTpNotaFiscal.ImpInfCompl = IIf(IsNull(Rst.Fields("ImprCampoObs")), "0", Rst.Fields("ImprCampoObs"))
            PgDadosTpNotaFiscal.Obs = IIf(IsNull(Rst.Fields("Obs")), "", Rst.Fields("Obs"))
            PgDadosTpNotaFiscal.conta = IIf(IsNull(Rst.Fields("Conta")), "0", Rst.Fields("Conta"))
            PgDadosTpNotaFiscal.TipoDoc = IIf(IsNull(Rst.Fields("TpDocumento")), "0", Rst.Fields("TpDocumento"))
            PgDadosTpNotaFiscal.CentroCusto = IIf(IsNull(Rst.Fields("CentroCusto")), "0", Rst.Fields("CentroCusto"))
            PgDadosTpNotaFiscal.PlanoContas = IIf(IsNull(Rst.Fields("PlanoContas")), "0", Rst.Fields("PlanoContas"))
            PgDadosTpNotaFiscal.CSTPIS = IIf(IsNull(Rst.Fields("CSTPIS")), "0", Rst.Fields("CSTPIS"))
            PgDadosTpNotaFiscal.CSTCOFINS = IIf(IsNull(Rst.Fields("CSTCOFINS")), "0", Rst.Fields("CSTCOFINS"))
            PgDadosTpNotaFiscal.ImpBCICMS = IIf(IsNull(Rst.Fields("ImpBCICMS")), "0", Rst.Fields("ImpBCICMS"))
            PgDadosTpNotaFiscal.ImpvICMS = IIf(IsNull(Rst.Fields("ImpvICMS")), "0", Rst.Fields("ImpvICMS"))
            PgDadosTpNotaFiscal.ImpBCICMSST = IIf(IsNull(Rst.Fields("ImpBCICMSST")), "0", Rst.Fields("ImpBCICMSST"))
            PgDadosTpNotaFiscal.ImpvICMSST = IIf(IsNull(Rst.Fields("ImpvICMSST")), "0", Rst.Fields("ImpvICMSST"))
            PgDadosTpNotaFiscal.ImpvTotalProduto = IIf(IsNull(Rst.Fields("ImpvTotalProduto")), "0", Rst.Fields("ImpvTotalProduto"))
            PgDadosTpNotaFiscal.ImpvFrete = IIf(IsNull(Rst.Fields("ImpvFrete")), "0", Rst.Fields("ImpvFrete"))
            PgDadosTpNotaFiscal.ImpvSeguro = IIf(IsNull(Rst.Fields("ImpvSeguro")), "0", Rst.Fields("ImpvSeguro"))
            PgDadosTpNotaFiscal.ImpvDesconto = IIf(IsNull(Rst.Fields("ImpvDesconto")), "0", Rst.Fields("ImpvDesconto"))
            PgDadosTpNotaFiscal.ImpvOutrasDesp = IIf(IsNull(Rst.Fields("ImpvOutrasDesp")), "0", Rst.Fields("ImpvOutrasDesp"))
            PgDadosTpNotaFiscal.ImpvIPI = IIf(IsNull(Rst.Fields("ImpvIPI")), "0", Rst.Fields("ImpvIPI"))
            PgDadosTpNotaFiscal.ImpvTotalNota = IIf(IsNull(Rst.Fields("ImpvTotalNota")), "0", Rst.Fields("ImpvTotalNota"))
    End If
    Rst.Close
    Exit Function
TrtErroTpNF:
        
        MsgBox "PgDadosTpNotaFiscal: " & Err.Description, vbInformation, Err.Number
        Resume Next
End Function
Public Function PgIdICMS(strSigla As String) As Integer
    Dim Rst     As Recordset
    Dim sSQL    As String
    sSQL = "SELECT * FROM TributacaoUF WHERE sigla = '" & strSigla & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar o ID do Registto"
            PgIdICMS = "0"
        Else
            Rst.MoveFirst
            PgIdICMS = Rst.Fields("id")
    End If
    Rst.Close
End Function
Public Function PgDadosCFOP(tpNF As Integer, cst As String, UFDest As String) As Dados_CFOP
    'tpNF - Tipo de Nota Fiscal Emitida
    'CST - Codigo do CST do Material
    'UFOrigem - Unidade federativa de origem
    'UFDest - Unidade federativa de destino
    '=========================================================
   
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim Situacao    As Integer
    
    Situacao = IIf(UFDest = PgDadosEmpresa(ID_Empresa).uf, 0, 1)
        
    sSQL = "SELECT * FROM FaturamentoTipoNotaFiscalcfop WHERE ID_Empresa = " & ID_Empresa & _
           " AND idTipoNotaFiscal = " & tpNF & " AND CST = " & cst & " AND situacao = " & Situacao '& "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar CFOP." & vbCrLf & vbCrLf & _
                   "ID Tipo de Nota Fiscal: " & tpNF & vbCrLf & _
                   "Situação: " & Situacao & " - " & IIf(Situacao = 0, "Venda dentro do Estado", "Venda fora do Estado") & vbCrLf & _
                   "CST: " & cst, vbInformation, "Aviso"
            Exit Function
        Else
            Rst.MoveFirst
            PgDadosCFOP.Situacao = IIf(IsNull(Rst.Fields("Situacao")), "0", Rst.Fields("Situacao"))
            PgDadosCFOP.cst = IIf(IsNull(Rst.Fields("CST")), "", Rst.Fields("CST"))
            PgDadosCFOP.CFOP = IIf(IsNull(Rst.Fields("CFOP")), "0", Rst.Fields("CFOP"))
            PgDadosCFOP.ICMS = IIf(IsNull(Rst.Fields("ICMS")), "0", Rst.Fields("ICMS"))
            PgDadosCFOP.ICMSST = IIf(IsNull(Rst.Fields("ICMSST")), "0", Rst.Fields("ICMSST"))
    End If
    Rst.Close
End Function

Public Function PgDescrTipoEmissao(Id As Integer) As String
    Select Case Id
        Case 1
            PgDescrTipoEmissao = "Normal"
        Case 2
            PgDescrTipoEmissao = "Contigencia com formulario de segurança (FS)"
        Case 3
            PgDescrTipoEmissao = "Contigencia com SCAN do ambiente nacional"
        Case 4
            PgDescrTipoEmissao = "Contigencia com DPEC"
        Case 5
            PgDescrTipoEmissao = "FS-DA"
        Case Else
            PgDescrTipoEmissao = ""
    End Select
End Function
Public Function PgDescrRegTrib(Id As Integer) As String
    Select Case Id
        Case 1
            PgDescrRegTrib = "Simples Nacional"
        Case 2
            PgDescrRegTrib = "Simples Nacional - exc. sublimite de rec. bruta"
        Case 3
            PgDescrRegTrib = "Regime Normal"
        'Case 4
        '    PgDescrRegTrib = "DPEC"
        'Case 5
        '    PgDescrRegTrib = "FS-DA"
        Case Else
            PgDescrRegTrib = "Tipo nao encontrado"
    End Select
End Function
Public Function pgAliqDifICMS(sNCM As String, sUFDest As String) As String
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim idNCM   As Integer
    sSQL = "SELECT * FROM TributacaoNCM WHERE ID_Empresa = " & ID_Empresa & " AND NCM='" & sNCM & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            pgAliqDifICMS = ""
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            idNCM = Rst.Fields("Id")
            Rst.Close
    End If
    sSQL = "SELECT * FROM TributacaoNCMICMS WHERE ID_Empresa = " & ID_Empresa & " AND idNCM = " & idNCM & " AND UF = '" & sUFDest & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            pgAliqDifICMS = ""
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            pgAliqDifICMS = Rst.Fields("ICMS")
            Rst.Close
    End If
    
End Function
Public Function pgDescrModBCST(idModBCST As Integer) As String
    Select Case idModBCST
        Case 0
            pgDescrModBCST = "Preço tabelado ou máximo sugerido"
        Case 1
            pgDescrModBCST = "Lista Negativa(valor)"
        Case 2
            pgDescrModBCST = "Lista Positiva(valor)"
        Case 3
            pgDescrModBCST = "Lista Neutra(valor)"
        Case 4
            pgDescrModBCST = "Margem Valor Agregado (%)"
        Case 5
            pgDescrModBCST = "Pauta (valor)"
        Case Else
            MsgBox "Modalidade de Base de Calculo não Cadastrada!", vbInformation, "Aviso"
    End Select

End Function

Public Function pgDescrModBC(idModBC As Integer) As String

    Select Case idModBC
        Case 0
            pgDescrModBC = "Margem de Valor Agregado(%)"
        Case 1
            pgDescrModBC = "Pauta (Valor)"
        Case 2
            pgDescrModBC = "Preço Tabelado Maximo(valor)"
        Case 3
            pgDescrModBC = "Valor da operação"
        Case Else
            MsgBox "Modalidade de Base de Calculo não Cadastrada!", vbInformation, "Aviso"
    End Select

End Function

Public Function PgDadosNotaFiscal(chv As String) As Dados_NotaFiscal
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE idNFe = '" & chv & "'"
    Set Rst = RegistroBuscar(sSQL)
    
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar Nota Fiscal!", vbInformation, "Aviso"
            Rst.Close
            Exit Function

        Else
            Rst.MoveFirst
            '* OUTROS
            PgDadosNotaFiscal.EnvioRF = cNull(Rst.Fields("enviorf"))
            PgDadosNotaFiscal.MovFinanceiro = cNull(Rst.Fields("movfinanceiro"))
            PgDadosNotaFiscal.MovFisco = cNull(Rst.Fields("movfisco"))
            PgDadosNotaFiscal.ImpFatura = cNull(Rst.Fields("impfatura"))
            '********************************************************
            'Do Until Rst.EOF
                PgDadosNotaFiscal.Id = cNull(Rst.Fields("id"))
                PgDadosNotaFiscal.nProt = cNull(Rst.Fields("nProt"))
                PgDadosNotaFiscal.dhProt = cNull(Rst.Fields("dhProt"))
                PgDadosNotaFiscal.lote = cNull(Rst.Fields("Lote"))
                PgDadosNotaFiscal.nRecibo = cNull(Rst.Fields("nRecibo"))
                PgDadosNotaFiscal.cStat = cNull(Rst.Fields("cStat"))
                PgDadosNotaFiscal.xMotivo = cNull(Rst.Fields("xMotivo"))
                PgDadosNotaFiscal.StatusNFe = cNull(Rst.Fields("StatusNFe"))
                '*********************************************************************************
                'NFe Cancelada
     PgDadosNotaFiscal.canc_nProt = cNull(Rst.Fields("canc_nProt"))
     PgDadosNotaFiscal.canc_dhRecbto = cNull(Rst.Fields("dhRecbto"))
     PgDadosNotaFiscal.canc_xJust = cNull(Rst.Fields("canc_xJust"))
     PgDadosNotaFiscal.canc_Status = cNull(Rst.Fields("canc_Status"))
    '*********************************************************************************
    'Numero de NFe Inutilizado
     PgDadosNotaFiscal.inut_nProt = cNull(Rst.Fields("inut_nProt"))
     PgDadosNotaFiscal.inut_dhRecbto = cNull(Rst.Fields("inut_dhRecbto"))
     PgDadosNotaFiscal.inut_xJust = cNull(Rst.Fields("inut_xJust"))
     PgDadosNotaFiscal.inut_Status = cNull(Rst.Fields("inut_Status"))
    '*********************************************************************************
    'cabecario do Pedido (ide)
     PgDadosNotaFiscal.Versao = cNull(Rst.Fields("versao"))
     PgDadosNotaFiscal.idNFe = cNull(Rst.Fields("idNFe"))
     PgDadosNotaFiscal.ide_cUF = cNull(Rst.Fields("ide_cUF"))
     PgDadosNotaFiscal.ide_cNF = cNull(Rst.Fields("ide_cNF"))
     PgDadosNotaFiscal.ide_natOp = cNull(Rst.Fields("ide_natOp"))
     PgDadosNotaFiscal.ide_indPag = cNull(Rst.Fields("ide_indPag"))
     PgDadosNotaFiscal.ide_mod = cNull(Rst.Fields("ide_mod"))
     PgDadosNotaFiscal.ide_serie = cNull(Rst.Fields("ide_serie"))
     PgDadosNotaFiscal.ide_nNF = cNull(Rst.Fields("ide_nNF"))
     PgDadosNotaFiscal.ide_dEmi = cNull(Rst.Fields("ide_dEmi"))
     PgDadosNotaFiscal.ide_dSaiEnt = cNull(Rst.Fields("ide_dSaiEnt"))
     PgDadosNotaFiscal.ide_hSaiEnt = cNull(Rst.Fields("ide_hSaiEnt"))
     PgDadosNotaFiscal.ide_tpNF = cNull(Rst.Fields("ide_tpNF"))
     PgDadosNotaFiscal.ide_cMunFG = cNull(Rst.Fields("ide_cMunFG"))
     PgDadosNotaFiscal.ide_refNFe = cNull(Rst.Fields("ide_refNFe"))
     PgDadosNotaFiscal.ide_tpImp = cNull(Rst.Fields("ide_tpImp"))
     PgDadosNotaFiscal.ide_tpEmis = cNull(Rst.Fields("ide_tpEmis"))
     PgDadosNotaFiscal.ide_cDV = cNull(Rst.Fields("ide_cDV"))
     PgDadosNotaFiscal.ide_tpAmb = cNull(Rst.Fields("ide_tpAmb"))
     PgDadosNotaFiscal.ide_finNFe = cNull(Rst.Fields("ide_finNFe"))
     PgDadosNotaFiscal.ide_procEmi = cNull(Rst.Fields("ide_procEmi"))
     PgDadosNotaFiscal.ide_verProc = cNull(Rst.Fields("ide_verProc"))
    'Emitente
     PgDadosNotaFiscal.emit_CNPJ = cNull(Rst.Fields("emit_CNPJ"))
     PgDadosNotaFiscal.emit_xNome = cNull(Rst.Fields("emit_xNome"))
     PgDadosNotaFiscal.emit_xFant = cNull(Rst.Fields("emit_xFant"))
     PgDadosNotaFiscal.emit_xLgr = cNull(Rst.Fields("emit_xLgr"))
     PgDadosNotaFiscal.emit_nro = cNull(Rst.Fields("emit_nro"))
     PgDadosNotaFiscal.emit_xCpl = cNull(Rst.Fields("emit_xCpl"))
     PgDadosNotaFiscal.emit_Bairro = cNull(Rst.Fields("emit_Bairro"))
     PgDadosNotaFiscal.emit_cMun = cNull(Rst.Fields("emit_cMun"))
     PgDadosNotaFiscal.emit_xMun = cNull(Rst.Fields("emit_xMun"))
     PgDadosNotaFiscal.emit_UF = cNull(Rst.Fields("emit_UF"))
     PgDadosNotaFiscal.emit_CEP = cNull(Rst.Fields("emit_CEP"))
     PgDadosNotaFiscal.emit_cPais = cNull(Rst.Fields("emit_cPais"))
     PgDadosNotaFiscal.emit_xPais = cNull(Rst.Fields("emit_xPais"))
     PgDadosNotaFiscal.emit_fone = cNull(Rst.Fields("emit_fone"))
     PgDadosNotaFiscal.emit_IE = cNull(Rst.Fields("emit_IE"))
    
     PgDadosNotaFiscal.emit_IEST = cNull(Rst.Fields("emit_IEST"))
     PgDadosNotaFiscal.emit_IM = cNull(Rst.Fields("emit_IM"))
     PgDadosNotaFiscal.emit_CNAE = cNull(Rst.Fields("emit_CNAE"))
    
     PgDadosNotaFiscal.emit_CRT = cNull(Rst.Fields("emit_CRT"))
    
    'Destinatario
     PgDadosNotaFiscal.dest_idDest = cNull(Rst.Fields("dest_idDest"))
     PgDadosNotaFiscal.dest_pessoa = cNull(Rst.Fields("dest_pessoa"))
     PgDadosNotaFiscal.dest_CNPJ = cNull(Rst.Fields("dest_CNPJ"))
     PgDadosNotaFiscal.dest_xNome = cNull(Rst.Fields("dest_xNome"))
     PgDadosNotaFiscal.dest_xFant = cNull(Rst.Fields("dest_xFant"))
     PgDadosNotaFiscal.dest_xLgr = cNull(Rst.Fields("dest_xLgr"))
     PgDadosNotaFiscal.dest_nro = cNull(Rst.Fields("dest_nro"))
     PgDadosNotaFiscal.dest_xCpl = cNull(Rst.Fields("dest_xCpl"))
     PgDadosNotaFiscal.dest_Bairro = cNull(Rst.Fields("dest_Bairro"))
     PgDadosNotaFiscal.dest_cMun = cNull(Rst.Fields("dest_cMun"))
     PgDadosNotaFiscal.dest_xMun = cNull(Rst.Fields("dest_xMun"))
     PgDadosNotaFiscal.dest_UF = cNull(Rst.Fields("dest_UF"))
     PgDadosNotaFiscal.dest_CEP = cNull(Rst.Fields("dest_CEP"))
     PgDadosNotaFiscal.dest_cPais = cNull(Rst.Fields("dest_cPais"))
     PgDadosNotaFiscal.dest_xPais = cNull(Rst.Fields("dest_xPais"))
     PgDadosNotaFiscal.dest_fone = cNull(Rst.Fields("dest_fone"))
     PgDadosNotaFiscal.dest_IE = cNull(Rst.Fields("dest_IE"))
     PgDadosNotaFiscal.dest_ISUF = cNull(Rst.Fields("dest_ISUF"))
     PgDadosNotaFiscal.dest_email = cNull(Rst.Fields("dest_email"))
     PgDadosNotaFiscal.infAdic_infCpl = cNull(Rst.Fields("infAdic_infCpl"))
    'Transporte
     PgDadosNotaFiscal.transp_modFrete = cNull(Rst.Fields("transp_modFrete"))
     PgDadosNotaFiscal.transp_Pessoa = cNull(Rst.Fields("transp_Pessoa"))
     PgDadosNotaFiscal.transp_CNPJ = cNull(Rst.Fields("transp_CNPJ"))
     PgDadosNotaFiscal.transp_xNome = cNull(Rst.Fields("transp_xNome"))
     PgDadosNotaFiscal.transp_IE = cNull(Rst.Fields("transp_IE"))
     PgDadosNotaFiscal.transp_xEnder = cNull(Rst.Fields("transp_xEnder"))
     PgDadosNotaFiscal.transp_xMun = cNull(Rst.Fields("transp_xMun"))
     PgDadosNotaFiscal.transp_UF = cNull(Rst.Fields("transp_UF"))
     PgDadosNotaFiscal.transp_qVol = cNull(Rst.Fields("transp_qVol"))
     PgDadosNotaFiscal.transp_esp = cNull(Rst.Fields("transp_esp"))
     PgDadosNotaFiscal.transp_marca = cNull(Rst.Fields("transp_marca"))
     PgDadosNotaFiscal.transp_nVol = cNull(Rst.Fields("transp_nVol"))
     PgDadosNotaFiscal.transp_pesoL = cNull(Rst.Fields("transp_pesoL"))
     PgDadosNotaFiscal.transp_pesoB = cNull(Rst.Fields("transp_pesoB"))
        
    'TOTAIS
     PgDadosNotaFiscal.total_vBC = cNull(Rst.Fields("total_vBC"))
     PgDadosNotaFiscal.total_vICMS = cNull(Rst.Fields("total_vICMS"))
     PgDadosNotaFiscal.total_vBCST = cNull(Rst.Fields("total_vBCST"))
     PgDadosNotaFiscal.total_vICMSST = cNull(Rst.Fields("total_vICMSST"))
     PgDadosNotaFiscal.total_vProd = cNull(Rst.Fields("total_vProd"))
     PgDadosNotaFiscal.total_vFrete = cNull(Rst.Fields("total_vFrete"))
     PgDadosNotaFiscal.total_vSeg = cNull(Rst.Fields("total_vSeg"))
     PgDadosNotaFiscal.total_vDesc = cNull(Rst.Fields("total_vDesc"))
     PgDadosNotaFiscal.total_vIPI = cNull(Rst.Fields("total_vIPI"))
     PgDadosNotaFiscal.total_vPIS = cNull(Rst.Fields("total_vPIS"))
     PgDadosNotaFiscal.total_vCOFINS = cNull(Rst.Fields("total_vCOFINS"))
     PgDadosNotaFiscal.total_vOutro = cNull(Rst.Fields("total_vOutro"))
     PgDadosNotaFiscal.total_vNF = cNull(Rst.Fields("total_vNF"))
     PgDadosNotaFiscal.ger_Vendedor = cNull(Rst.Fields("ger_Vendedor"))
     PgDadosNotaFiscal.ger_idPV = cNull(Rst.Fields("ger_idPV"))
    
    
    'Produto******************************************************************************
    
    ' IdNFe
    ' det_IdProduto
    ' det_cProd
    ' det_cEAN
    ' det_xProd
    ' det_InfAdProd
    ' det_NCM
    ' det_EXTIPI
    ' det_CFOP
    ' det_uCom
    ' det_qCom
    ' det_vUnCom
    ' det_vProd
    
    ' det_cEANTrib
    ' det_uTrib
    ' det_qTrib
    ' det_vUnTrib
    ' det_vFrete
    ' det_vSeg
    ' det_vDesc
    ' det_vOutro
    ' det_indTot
    ' det_xPed
    ' det_nItemPed
    
    'IMPOSTOS
    'ICMS - 'Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST|vICMSST
    ' ICMS_origem
    ' ICMS_CST
    ' ICMS_modBC
    ' ICMS_pRedBC
    ' ICMS_vBC
    ' ICMS_pICMS
    ' ICMS_vICMS
    ' ICMS_modBCST
    ' ICMS_pMVAST
    ' ICMS_pRedBCST
    ' ICMS_vBCST
    ' ICMS_pICMSST
    ' ICMS_vICMSST
    ' ICMS_MotDesICMS
    ''IPI
    ' IPI_cEnq
    ' IPI_CST
    ' IPI_vBC
    ' IPI_pIPI
    ' IPI_vIPI
    ''PIS
    ' PIS_CST
    ' PIS_vBC
    ' PIS_pPIS
     'PIS_vPIS
    ''COFINS
    ' COFINS_CST
    ' COFINS_vBC
    ' COFINS_pCOFINS
    ' COFINS_vCOFINS
    ''Informacoes Gerenciais
    ' estoque_Unid
    ' estoque_Qtd
    ' estoque_vUnit
    ' comissao_pComissao
    ' comissao_vComissao
    ''COBRANCA
    ' IdNFe
    ' cobr_TpDoc
    ' cobr_nFat
    ' cobr_vOrig
    ' cobr_vDesc
    ' cobr_vLiq
    ' cobr_nDup
    ' cobr_dVenc
    ' cobr_vDup
    ' cobr_Emissao
    ' cobr_Multa
    ' cobr_Mora
    ' cobr_Protesto
    'cobr_idCliente
    'Email
    'email_IdNFe
    'Email_Status


    End If
    Rst.Close


End Function
Public Function PgDadosPlanoContas(Campo As String, Busca As String) As Dados_PlanoContas
    '#################################################################################
    '### 24/01/2012
    '### Busca o Plano de Contas de acordo com os criterios da busca
    
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    sSQL = "SELECT * FROM FinanceiroPlanoContas " & _
         "WHERE " & Campo & "=" & Busca
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            PgDadosPlanoContas.Id = cNull(Rst.Fields("Id"))
            PgDadosPlanoContas.Descricao = cNull(Rst.Fields("Descricao"))
            PgDadosPlanoContas.Codigo = cNull(Rst.Fields("codigo"))
            PgDadosPlanoContas.cd = cNull(Rst.Fields("CD"))
            If cNull(Rst.Fields("totalizador")) = "" Then
                    PgDadosPlanoContas.totalizador = 0
                Else
                    PgDadosPlanoContas.totalizador = 1
            End If
            'PgDadosPlanoContas.totalizador = cNull(Rst.Fields("totalizador"))
    End If
    
    Rst.Close
End Function
Public Function PgDadosCEST(sCampo As String, sBusca As String, SN As String) As Dados_CEST
    Dim Rst     As Recordset
    Dim sSQL    As String
    If SN = "N" Then
            sSQL = "SELECT * FROM tributacaocest WHERE " & sCampo & " = " & sBusca
        Else
            sSQL = "SELECT * FROM tributacaocest WHERE " & sCampo & " = '" & sBusca & "'"
    End If
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'MsgBox "Erro ao localizar dados NCM", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            PgDadosCEST.Id = cNull(Rst.Fields("Id"))
            PgDadosCEST.Descricao = cNull(Rst.Fields("Descricao"))
            PgDadosCEST.NCM = cNull(Rst.Fields("NCM"))
            'PgDadosCEST.pIPI = cNull(Rst.Fields("IPI"))
            PgDadosCEST.cest = cNull(Rst.Fields("cest"))
    End If
    Rst.Close
End Function

