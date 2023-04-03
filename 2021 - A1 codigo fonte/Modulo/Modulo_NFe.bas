Attribute VB_Name = "Modulo_NFe"
Option Explicit
Public VersaoNFe    As String
Type CalculoST
    vBCICMSST       As String
    vICMSST         As String
End Type
Private strMountTXT    As String
Public Function Calculo_ICMSST(UF_Origem As String, UF_Destino As String, sMVA As String, sValor As String, vICMSOrigem As String) As CalculoST
    
    
    Dim pICMS_Origem     As String ' % Icms Origem (valor que gerara credito de icms / aliq destacada na NF)
    Dim pICMS_Destino    As String ' % Icms Destino (Aliq interna no estado de destino)
    
    Dim vICMS_Origem     As String ' $ Icms Origem
    Dim vICMS_Destino    As String ' $ Icms Destino
    
    Dim vBCICMSST        As String ' Valor da base de Calculo ICMS ST
    Dim vICMSST          As String ' Valor do ICMS ST
    
    If Trim(sMVA) = "" Or Trim(sMVA) = "0" Then
        'MsgBox "MVA com aliquota ZERO! Favor verificar!", vbInformation, "Aviso"
        Calculo_ICMSST.vBCICMSST = ChkVal("0", 0, cDecMoeda)
        Calculo_ICMSST.vICMSST = ChkVal("0", 0, cDecMoeda)
        Exit Function
    End If
    
    pICMS_Origem = pgDadosICMS(UF_Destino, 0).ICMS
    pICMS_Destino = pgDadosICMS(UF_Destino, 0).ICMSInt
    
    'ICMS de Origem
    'vICMS_Origem = Val(ChkVal(pICMS_Origem, 0, cDecMoeda)) * Val(ChkVal(sValor, 0, cDecMoeda))
    vICMS_Origem = ChkVal(vICMSOrigem, 0, cDecMoeda)
    
    
    'Valor do BC ICMS ST
    vBCICMSST = Val(ChkVal(sMVA, 0, cDecMoeda)) * Val(ChkVal(sValor, 0, cDecMoeda))
    vBCICMSST = Val(ChkVal(vBCICMSST, 0, cDecMoeda)) / 100
    vBCICMSST = Val(ChkVal(vBCICMSST, 0, cDecMoeda)) + Val(ChkVal(sValor, 0, cDecMoeda))
    
    
    'Valor do ICMS Destino com BC ICMS ST
    vICMS_Destino = Val(ChkVal(pICMS_Destino, 0, cDecMoeda)) * Val(ChkVal(vBCICMSST, 0, cDecMoeda))
    vICMS_Destino = Val(ChkVal(vICMS_Destino, 0, cDecMoeda)) / 100
    
    
    
    'ICMS ST
    vICMSST = Val(ChkVal(vICMS_Destino, 0, cDecMoeda)) - Val(ChkVal(vICMS_Origem, 0, cDecMoeda))
    
    
    'Resultado
    Calculo_ICMSST.vBCICMSST = ChkVal(vBCICMSST, 0, cDecMoeda)
    Calculo_ICMSST.vICMSST = ChkVal(vICMSST, 0, cDecMoeda)
    
    
End Function
Public Sub Consultar_NFe(chvNFe As String)
    'Dim Rst     As Recordset
    'Dim sSQL     As String
    Dim nmArq   As String
    
    Dim nProt   As String
    Dim xJust   As String
    
    '##############################################################################################
    '### 07/05/2012
    '### Registra a consulta de situação
    '.TextMatrix(.Rows - 1, 6) = IIf(IsNull(Rst.Fields("StatusNFe")), "Aguardando...", Rst.Fields("StatusNFe")) 'pgStatusNFe(.TextMatrix(.Rows - 1, 7))
    Dim vReg(10)    As Variant
    Dim cReg        As Integer
    Dim idNFe       As Integer
    
    cReg = 0
    idNFe = PgDadosNotaFiscal(chvNFe).Id
    
    vReg(cReg) = Array("cStat", "", "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("StatusNFe", "Consultando a Situação da NFe junto a SEFAZ...", "S")
    
    RegistroAlterar "FaturamentoNFe", vReg, cReg, "id=" & idNFe
    '##############################################################################################
    
    nmArq = chvNFe & "-ped-sit.txt"
    
    ChecarArquivo nmArq
    'sSQL = "SELECT * FROM FaturamentoNFe WHERE idNFe ='" & chvnfe & "'"
    'Set Rst = RegistroBuscar(sSQL)
    'If Rst.BOF And Rst.EOF Then
    '        Exit Sub
    '    Else
    '        nProt = IIf(IsNull(Rst.Fields("nProt")), "", Rst.Fields("nProt")) 'Mid(Rst.Fields("nProt"), 1, InStr(Rst.Fields("nProt"), " "))
    '        xJust = IIf(IsNull(Rst.Fields("canc_xJust")), "", Rst.Fields("canc_xJust"))
    'End If
    'Rst.Close
    grvReg nmArq, "versao|" & VersaoNFe
    grvReg nmArq, "tpAmb|" & PgDadosConfig.Ambiente
    grvReg nmArq, "xServ|CONSULTAR"
    grvReg nmArq, "chNFe|" & chvNFe
    'grvReg nmArq, "nProt|" & nProt
    'grvReg nmArq, "xJust|" & xJust
    MoverPastaEnvio_UniNFe (nmArq)
End Sub

Public Sub Cancelar_NFe(chvNFe As String)
    'Alterado em 27/08/2013
    '
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim nmArq       As String
    
    Dim nProt       As String
    Dim xJust       As String
    Dim dhProt      As String
    Dim tpEvento    As String
    
    Dim hrEvent As String
    Dim dtEvent As String
       
    tpEvento = "110111"
    nmArq = chvNFe & "-env-canc.txt"
    
    ChecarArquivo (nmArq)
    
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND idNFe ='" & chvNFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            nProt = IIf(IsNull(Rst.Fields("nProt")), "", Rst.Fields("nProt")) 'Mid(Rst.Fields("nProt"), 1, InStr(Rst.Fields("nProt"), " "))
            xJust = IIf(IsNull(Rst.Fields("canc_xJust")), "", Rst.Fields("canc_xJust"))
            dhProt = IIf(IsNull(Rst.Fields("dhprot")), "", Format(Rst.Fields("dhprot"), "YYYY-MM-DDTHH:MM:SS"))
    
    
            'hrEvent = PgDadosNotaFiscal(chvnfe).dhProt
            dtEvent = Format(Trim(Mid(dhProt, 1, InStr(dhProt, " "))), "YYYY-MM-DD")
            hrEvent = Trim(Mid(dhProt, InStr(dhProt, " "), Len(dhProt)))
            
    
    End If
    Rst.Close
    grvReg nmArq, "versao|1.00"
    grvReg nmArq, "idLote|0000000001"
    grvReg nmArq, "evento|1"
    grvReg nmArq, "versao|1.00"
    grvReg nmArq, "Id|ID" & tpEvento & chvNFe & "01"
    grvReg nmArq, "cOrgao|" & PgDadosConfig.uf
    grvReg nmArq, "tpAmb|" & PgDadosConfig.Ambiente
    grvReg nmArq, "xServ|CANCELAR"
    grvReg nmArq, "CNPJ|" & PgDadosEmpresa(ID_Empresa).CNPJ
    grvReg nmArq, "chNFe|" & chvNFe
    '** ERRO - 27.04.2015
    '** Tem de colocar a data do evento no formato AAAA-MM-DDThh:mm:ss-03:00
    grvReg nmArq, "dhEvento|" & dtEvent & "T" & hrEvent 'dhProt & PgDadosConfig.fusoHorario  '2013-08-19T08:41:42-03:00"
    grvReg nmArq, "tpEvento|" & tpEvento
    grvReg nmArq, "nSeqEvento|1"
    grvReg nmArq, "verEvento|1.00"
    grvReg nmArq, "detEvento|CANCELAMENTO"
    grvReg nmArq, "descEvento|Cancelamento"
    grvReg nmArq, "nProt|" & nProt
    grvReg nmArq, "xJust|" & xJust

    'grvReg nmArq, "tpAmb|" & PgDadosConfig.Ambiente
    'grvReg nmArq, "xServ|CANCELAR"
    'grvReg nmArq, "chNFe|" & chvnfe
    'grvReg nmArq, "nProt|" & nProt
    'grvReg nmArq, "xJust|" & xJust
    MoverPastaEnvio_UniNFe (nmArq)
End Sub
Public Sub Inutilizar_NFe(chvNFe As String)
    Dim Rst     As Recordset
    Dim sSQL     As String
    Dim nmArq   As String
    
    Dim nProt   As String
    Dim xJust   As String
    
    nmArq = chvNFe & "-ped-inu.txt"
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND idNFe ='" & chvNFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        MsgBox "Erro ao localizar registros de inutilização!", vbInformation, "Aviso"
        Exit Sub
    End If
    
    grvReg nmArq, "ID|" & chvNFe
    grvReg nmArq, "versao|" & Rst.Fields("versao")
    grvReg nmArq, "tpAmb|" & PgDadosConfig.Ambiente
    grvReg nmArq, "cUF|" & Rst.Fields("ide_cUF")
    grvReg nmArq, "Ano|" & Format(Rst.Fields("ide_dEmi"), "YY")
    grvReg nmArq, "mod|" & Rst.Fields("ide_Mod")
    grvReg nmArq, "Serie|" & Rst.Fields("ide_Serie")
    grvReg nmArq, "CNPJ|" & Rst.Fields("emit_CNPJ")
    grvReg nmArq, "nNFIni|" & Mid(chvNFe, 24, 9)
    grvReg nmArq, "nNFFin|" & Mid(chvNFe, 33, 9)
    grvReg nmArq, "xJust|" & Rst.Fields("inut_xJust")
    
    Rst.Close
    
    MoverPastaEnvio_UniNFe (nmArq)
End Sub
Public Sub MoverPastaEnvio_UniNFe(nArq As String)
    On Error Resume Next
    FileCopy PgDadosConfig.pFileArmazenamento & "\" & nArq, PgDadosConfig.pEnvio & "\" & nArq
End Sub

Public Function ChaveAcesso(numNota As String, dt As Date, _
                            cUF As String, _
                            CNPJ As String, _
                            Modelo As String, _
                            Serie As String, _
                            tpEmis As String, _
                            cNF As String)
    
'***********************************************************************
'**    Manual de Integração -  Contribuinte Versão 4.0.1 2009.006     **
'***********************************************************************
    Dim nNF             As String
    Dim AAMM            As String
    Dim cDV             As String
    Dim strChvAcesso    As String
        
    AAMM = Format(dt, "YYMM")
    CNPJ = Left(String(14, "0"), 14 - Len(Trim(CNPJ))) & CNPJ
    Modelo = Left("00", 2 - Len(Trim(Modelo))) & Trim(Modelo)
    Serie = Left("000", 3 - Len(Trim(Serie))) & Trim(Serie)
    nNF = Mid(String(9, "0"), 1, 9 - Len(Trim(numNota))) & Trim(numNota) '"123456789"
    'tpEmis = 1 '1-Normal 2-Contigemcia
    
    
    'strChvAcesso = cUF & AAMM & CNPJ & Modelo & Serie & nNF & tpEmis & cNF
    strChvAcesso = cUF & AAMM & CNPJ & Modelo & Serie & nNF & tpEmis & cNF
    
    cDV = cDV11(strChvAcesso)
    ChaveAcesso = strChvAcesso & cDV
    
    'ChaveAcesso = Format(ChaveAcesso, "@@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@")
    'Frame10.Caption = "Tamanho da chave: " & Len(ChaveAcesso)
    
End Function
Public Function cDV11(strNumero As String) As String
    Dim i      As Integer
    Dim k      As Integer
    Dim Soma   As Integer

    Soma = 0
    k = 2
    For i = Len(strNumero) To 1 Step -1
        Soma = Soma + (Val(Mid(strNumero, i, 1)) * k)
        k = k + 1
        If k > 9 Then k = 2
    Next
    Soma = 11 - (Soma Mod 11)
    If Soma >= 10 Then Soma = 0
    
    cDV11 = Chr(Soma + Asc("0"))
  
End Function
Private Sub ChecarArquivo(nmArquivo)
    Dim caminho As String
    caminho = PgDadosConfig.pFileArmazenamento & "\" & nmArquivo
    If Dir(caminho) <> "" Then
        Kill caminho
    End If
End Sub

Public Function Exportar_CCe_v200_TXT(chvNFe As String) As String
    Dim Rst1        As Recordset
    Dim sSQL        As String
    Dim nmArq       As String
    Dim dtEvent     As String
    Dim hrEvent     As String
    
    
    sSQL = "SELECT * FROM FaturamentoNFeCartaCorrecao WHERE ID_Empresa = " & ID_Empresa & " AND chvNFe = '" & chvNFe & "'"
    Set Rst1 = RegistroBuscar(sSQL)
    If Rst1.BOF And Rst1.EOF Then
            MsgBox "Erro ao localizar Carta de Correção na chave " & chvNFe, vbCritical, "Aviso"
            Exportar_CCe_v200_TXT = ""
            Exit Function
        Else
            Rst1.MoveFirst
    End If
    
     nmArq = chvNFe & "-00-env-cce.txt"
     
     ChecarArquivo (nmArq)
'========================================================================
    
    grvReg nmArq, "idLote|" & ZE(Rst1.Fields("id"), 15)
    grvReg nmArq, "evento|1"
    grvReg nmArq, "cOrgao|" & PgDadosConfig.uf
    grvReg nmArq, "tpAmb|" & PgDadosConfig.Ambiente
    grvReg nmArq, "CNPJ|" & PgDadosEmpresa(ID_Empresa).CNPJ
    grvReg nmArq, "chNFe|" & chvNFe
    
    hrEvent = PgDadosNotaFiscal(chvNFe).dhProt
    dtEvent = Format(Trim(Mid(hrEvent, 1, InStr(hrEvent, " "))), "YYYY-MM-DD")
    hrEvent = Trim(Mid(hrEvent, InStr(hrEvent, " "), Len(hrEvent)))
    
    grvReg nmArq, "dhEvento|" & dtEvent & "T" & hrEvent '& PgDadosConfig.fusoHorario
    grvReg nmArq, "tpEvento|110110"
    grvReg nmArq, "nSeqEvento|1"
    grvReg nmArq, "verEvento|1.00"
    grvReg nmArq, "xCorrecao|" & Rst1.Fields("Correcao")
    
    Exportar_CCe_v200_TXT = nmArq
    Rst1.Close
    'Move o txt para validacao do uninfe
    MoverPastaEnvio_UniNFe (nmArq)
End Function
Public Function Exportar_NFe_v200_TXT(chvNFe As String) As String
    Dim Rst1    As Recordset 'Cabecalho
    Dim Rst2    As Recordset 'Produto
    Dim Rst3    As Recordset 'Cobanca
    Dim sSQL    As String
    Dim nmArq   As String
    Dim cItens  As Integer 'Conta os registros dos itens da Nota
    Dim cCob    As Integer 'Conta os registros da cobranca da Nota
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & chvNFe & "'"
    Set Rst1 = RegistroBuscar(sSQL)
    If Rst1.BOF And Rst1.EOF Then
            MsgBox "Etapa 1 - Erro ao localizar NF-e"
            Exportar_NFe_v200_TXT = ""
            Exit Function
        Else
            Rst1.MoveFirst
    End If
    sSQL = "SELECT * FROM FaturamentoNFeItens WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & chvNFe & "'"
    Set Rst2 = RegistroBuscar(sSQL)
    If Rst2.BOF And Rst2.EOF Then
            MsgBox "Etapa 2 - Erro ao localizar NF-e"
            Exportar_NFe_v200_TXT = ""
            Exit Function
        Else
            Rst2.MoveFirst
    End If
    sSQL = "SELECT * FROM FaturamentoNFeCobranca WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & chvNFe & "'"
    Set Rst3 = RegistroBuscar(sSQL)
    If Rst3.BOF And Rst3.EOF Then
            MsgBox "Etapa 3 - Erro ao localizar NF-e"
            Exportar_NFe_v200_TXT = ""
            Exit Function
        Else
            Rst3.MoveFirst
    End If
     nmArq = Rst1.Fields("ide_nNF") & "_" & Rst1.Fields("emit_CNPJ") & "_" & Format(Rst1.Fields("ide_dEmi"), "DD") & "_" & Format(Rst1.Fields("ide_dEmi"), "MM") & "_" & Format(Rst1.Fields("ide_dEmi"), "YYYY") & "-nfe.txt"
     
     ChecarArquivo (nmArq)
'========================================================================
    
    grvReg nmArq, "NOTAFISCAL|1"
    'A
    grvReg nmArq, "A|" & Rst1.Fields("Versao") & "|NFe" & chvNFe & "|"
    'B
    grvReg nmArq, "B|" & _
                    Rst1.Fields("ide_cUF") & "|" & _
                    Rst1.Fields("ide_cNF") & "|" & _
                    Rst1.Fields("ide_NatOP") & "|" & _
                    Rst1.Fields("ide_indPag") & "|" & _
                    Rst1.Fields("ide_Mod") & "|" & _
                    Rst1.Fields("ide_serie") & "|" & _
                    Rst1.Fields("ide_nNF") & "|" & _
                    Format(Rst1.Fields("ide_demi"), "YYYY-MM-DD") & "|" & _
                    Format(Rst1.Fields("ide_dSaiEnt"), "YYYY-MM-DD") & "|" & _
                    Format(Rst1.Fields("ide_hSaiEnt"), "HH:MM:SS") & "|" & _
                    Rst1.Fields("ide_tpNF") & "|" & _
                    Rst1.Fields("ide_cMunFG") & "|" & _
                    Rst1.Fields("ide_TpImp") & "|" & _
                    Rst1.Fields("ide_tpEmis") & "|" & _
                    Rst1.Fields("ide_cDV") & "|" & _
                    Rst1.Fields("ide_tpAmb") & "|" & _
                    Rst1.Fields("ide_finNFe") & "|" & _
                    Rst1.Fields("ide_procEmi") & "|" & _
                    Rst1.Fields("ide_VerProc") & "|" & _
                    IIf(PgDadosConfig.ContingenciaDt <> "", Format(PgDadosConfig.ContingenciaDt, "YYYY-MM-DD") & "T" & PgDadosConfig.ContingenciaHr, "") & "|" & _
                    PgDadosConfig.ContingenciaMotivo & "|"
'                    "2012-05-28T11:39:30|" & _
                    "ERRO NA SVRS AUTORIZADO PELO SITE DA NFE. 28/05/2012" & "|"
                    
    If Not IsNull(Rst1.Fields("ide_refNFe")) Then
        grvReg nmArq, "B13|" & Rst1.Fields("ide_refNFe") & "|"
    End If
    
    'C - dados EMITENTE
    grvReg nmArq, "C|" & _
                    Rst1.Fields("emit_xNome") & "|" & _
                    Rst1.Fields("emit_xFant") & "|" & _
                    Rst1.Fields("emit_IE") & "|" & _
                    Rst1.Fields("emit_IEST") & "|" & _
                    Rst1.Fields("emit_IM") & "|" & _
                    IIf(Trim(Rst1.Fields("emit_IM")) <> "", Rst1.Fields("emit_CNAE"), "") & "|" & _
                    Rst1.Fields("emit_CRT") & "|"
                    
    grvReg nmArq, "C02|" & _
                    Rst1.Fields("emit_CNPJ") & "|"
    grvReg nmArq, "C05|" & _
                    Rst1.Fields("emit_xLgr") & "|" & _
                    Rst1.Fields("emit_nro") & "|" & _
                    Rst1.Fields("emit_xcpl") & "|" & _
                    Rst1.Fields("emit_Bairro") & "|" & _
                    Rst1.Fields("emit_cMun") & "|" & _
                    Rst1.Fields("emit_xMun") & "|" & _
                    Rst1.Fields("emit_UF") & "|" & _
                    Rst1.Fields("emit_CEP") & "|" & _
                    Rst1.Fields("emit_cPais") & "|" & _
                    Rst1.Fields("emit_xPais") & "|" & _
                    Rst1.Fields("emit_fone") & "|"
    'E - dados DESTINATARIO
    grvReg nmArq, "E|" & _
                    Rst1.Fields("dest_xNome") & "|" & _
                    Rst1.Fields("dest_IE") & "|" & _
                    Rst1.Fields("dest_ISUF") & "|" & _
                    Rst1.Fields("dest_email") & "|"
                    
    grvReg nmArq, "E" & IIf(UCase(Rst1.Fields("dest_Pessoa")) = "FISICA", "03", "02") & "|" & _
                    Rst1.Fields("dest_CNPJ") & "|"
                    
                    
    grvReg nmArq, "E05|" & _
                    Rst1.Fields("dest_xLgr") & "|" & _
                    Rst1.Fields("dest_nro") & "|" & _
                    Rst1.Fields("dest_xCpl") & "|" & _
                    Rst1.Fields("dest_Bairro") & "|" & _
                    Rst1.Fields("dest_cMun") & "|" & _
                    Rst1.Fields("dest_xMun") & "|" & _
                    Rst1.Fields("dest_UF") & "|" & _
                    Rst1.Fields("dest_CEP") & "|" & _
                    Rst1.Fields("dest_cPais") & "|" & _
                    Rst1.Fields("dest_xPais") & "|" & _
                    Rst1.Fields("dest_fone") & "|"
    
    'G - dadosEntrega
    If Trim(Rst1.Fields("dest_cnpj")) <> Trim(Rst1.Fields("entr_CNPJ")) Then
        grvReg nmArq, "G|" & _
                      Rst1.Fields("entr_xLgr") & "|" & _
                      Rst1.Fields("entr_nro") & "|" & _
                      Rst1.Fields("entr_xCpl") & "|" & _
                      Rst1.Fields("entr_xBairro") & "|" & _
                      Rst1.Fields("entr_cMun") & "|" & _
                      Rst1.Fields("entr_xMun") & "|" & _
                      Rst1.Fields("entr_UF") & "|"
        grvReg nmArq, "G02|" & _
                      Rst1.Fields("entr_CNPJ") & "|"
    End If
    'H - dados DESCRICAO DOS ITENS
    Rst2.MoveFirst
    For cItens = 0 To Rst2.RecordCount - 1
        grvReg nmArq, "H|" & cItens + 1 & "|" & _
                       IIf(Trim(Rst2.Fields("det_InfAdProd")) = "", "", Rst2.Fields("det_InfAdProd") & "|")
                       
        grvReg nmArq, "I|" & Rst2.Fields("det_cProd") & "|" & _
                        Rst2.Fields("det_cEAN") & "|" & _
                        Rst2.Fields("det_xProd") & "|" & _
                        Rst2.Fields("det_NCM") & "|" & _
                        Rst2.Fields("det_EXTIPI") & "|" & _
                        Rst2.Fields("det_CFOP") & "|" & _
                        Rst2.Fields("det_uCom") & "|" & _
                        Rst2.Fields("det_qCom") & "|" & _
                        Rst2.Fields("det_vUnCom") & "|" & _
                        Rst2.Fields("det_vprod") & "|" & _
                        Rst2.Fields("det_cEANTrib") & "|" & _
                        Rst2.Fields("det_uTrib") & "|" & _
                        Rst2.Fields("det_qTrib") & "|" & _
                        Rst2.Fields("det_vUnTrib") & "|" & _
                        Rst2.Fields("det_vFrete") & "|" & _
                        Rst2.Fields("det_vSeg") & "|" & _
                        Rst2.Fields("det_vDesc") & "|" & _
                        Rst2.Fields("det_vOutro") & "|" & _
                        Rst2.Fields("det_indTot") & "|" & _
                        Rst2.Fields("det_xPed") & "|" & _
                        Rst2.Fields("det_nItemPed") & "|"
        
        grvReg nmArq, "M|"
        grvReg nmArq, "N|"
        '*****************************************************************
        'ICMS ************************************************************
        Select Case Rst2.Fields("ICMS_CST")
            Case "00" 'Tributacao Integral (N02)
                grvReg nmArq, "N02|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|"
            Case "10" 'Tributada com cobranca ICMS (N03)
                grvReg nmArq, "N03|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
            Case "20" 'Tributacao do ICMS com reducao da Base de Calculo (N04)
                 grvReg nmArq, "N04|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|"
                                
            Case "30" 'Tributacao Isenta com cobranca de ICMS por ST (N05)
                grvReg nmArq, "N05|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
                                
            Case "40" 'Tributacao do ICMS ISENTA (N06)
                grvReg nmArq, "N06|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_MotDesICMS") & "|"
                                
            'Case "41" 'Tributacao do ICMS NAO TRIBUTADA ()
             grvReg nmArq, "N06|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_MotDesICMS") & "|"
            
            Case "50" 'Tributacao do ICMS SUSPENSAO ()
                grvReg nmArq, "N06|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|"
            
            Case "51" 'Tributacao do ICMS POR DIFERIMENTO (N07)
                grvReg nmArq, "N07|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|"
            
            Case "60" 'ICMS cobrado anteriormente por ST (N08)
                grvReg nmArq, "N08" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
            Case "70" 'Tributacao do com reducao da base de calculo do ICMS ST (N09)
                grvReg nmArq, "N09" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
            Case "90" ' 'Tributacao OUTROS (N10)
                grvReg nmArq, "N10" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
            
            Case "500" 'Tributacao de ICMS pelo SIMPLES NACIONAL (N10g)
             grvReg nmArq, "N10g" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                "0.00" & "|" & _
                                "0.00" & "|"
        End Select
        '************************************************************
        'IPI ********************************************************
        grvReg nmArq, "O|" & _
                        "|" & _
                        "|" & _
                        "|" & _
                        "|" & _
                        Rst2.Fields("IPI_cEnq") & "|"
        grvReg nmArq, "O07|" & _
                        Rst2.Fields("IPI_CST") & "|" & _
                        Rst2.Fields("IPI_vIPI") & "|"
        grvReg nmArq, "O10|" & _
                        Rst2.Fields("IPI_vBC") & "|" & _
                        Rst2.Fields("IPI_pIPI") & "|"
                        
                        
                        
        'PIS ************************************************************
        grvReg nmArq, "Q|"
        Select Case Rst2.Fields("PIS_CST")
            Case "01", "02"  'Aliquota Normal/Aliquota Diferenciada (Q02)
                grvReg nmArq, "Q02|" & _
                                Rst2.Fields("PIS_CST") & "|" & _
                                Rst2.Fields("PIS_vBC") & "|" & _
                                Rst2.Fields("PIS_pPIS") & "|" & _
                                Rst2.Fields("PIS_vPIS") & "|"
            Case Else
                grvReg nmArq, "Q04|" & Rst2.Fields("PIS_CST") & "|"
                'MsgBox "Verificar o Codigo de exportacao do PIS da NFe - CODIGO DO CST DO PIS DESCONHECIDO"
        End Select
        
        
        
        'COFINS ************************************************************
        grvReg nmArq, "S|"
        Select Case Rst2.Fields("COFINS_CST")
            Case "01", "02"  'Aliquota Normal/Aliquota Diferenciada (Q02)
                grvReg nmArq, "S02|" & _
                                Rst2.Fields("COFINS_CST") & "|" & _
                                Rst2.Fields("COFINS_vBC") & "|" & _
                                Rst2.Fields("COFINS_pCOFINS") & "|" & _
                                Rst2.Fields("COFINS_vCOFINS") & "|"
            Case Else
                'MsgBox "Verificar o Codigo de exportacao do COFINS da NFe - CODIGO DO CST DO COFINS DESCONHECIDO"
                grvReg nmArq, "S04|" & Rst2.Fields("COFINS_CST") & "|"
        End Select
        
        Rst2.MoveNext
       
    Next

    '*********************************** TOTAIS DA NF-e *****************************************
    grvReg nmArq, "W|"
    grvReg nmArq, "W02|" & _
                    Rst1.Fields("total_vBC") & "|" & _
                    Rst1.Fields("total_vICMS") & "|" & _
                    Rst1.Fields("total_vBCST") & "|" & _
                    Rst1.Fields("total_vICMSST") & "|" & _
                    Rst1.Fields("total_vProd") & "|" & _
                    Rst1.Fields("total_vFrete") & "|" & _
                    Rst1.Fields("total_vSeg") & "|" & _
                    Rst1.Fields("total_vDesc") & "|" & _
                    "0.00" & "|" & _
                    Rst1.Fields("total_vIPI") & "|" & _
                    Rst1.Fields("total_vPIS") & "|" & _
                    Rst1.Fields("total_vCOFINS") & "|" & _
                    Rst1.Fields("total_vOutro") & "|" & _
                    Rst1.Fields("total_vNF") & "|"

    '*************************************** TRANSPORTE ********************************************
    grvReg nmArq, "X|" & _
                    Rst1.Fields("transp_ModFrete") & "|"
    
        grvReg nmArq, "X03|" & _
                    Rst1.Fields("transp_xNome") & "|" & _
                    Rst1.Fields("transp_IE") & "|" & _
                    Rst1.Fields("transp_xEnder") & "|" & _
                    Rst1.Fields("transp_UF") & "|" & _
                    Rst1.Fields("transp_xMun") & "|"
    
    grvReg nmArq, "X" & IIf(UCase(Rst1.Fields("transp_Pessoa")) = "FISICA", "05", "04") & "|" & _
                    Rst1.Fields("transp_CNPJ") & "|"
    If cNull(Rst1.Fields("transp_VeicPlaca")) <> "" Then
        grvReg nmArq, "X18|" & _
                    cNull(Rst1.Fields("transp_VeicPlaca")) & "|" & _
                    cNull(Rst1.Fields("transp_VeicUF")) & "|" & _
                    "" & "|"
    End If
    grvReg nmArq, "X26|" & _
                    Rst1.Fields("transp_qVol") & "|" & _
                    Rst1.Fields("transp_esp") & "|" & _
                    Rst1.Fields("transp_marca") & "|" & _
                    Rst1.Fields("transp_nVol") & "|" & _
                    Rst1.Fields("transp_PesoL") & "|" & _
                    Rst1.Fields("transp_PesoB") & "|"

    '***************************************** COBRANCA ********************************************
    If PgDadosNotaFiscal(chvNFe).ImpFatura = 1 Then
        grvReg nmArq, "Y|"
        If cNull(Rst3.Fields("cobr_nFat")) <> "" Then
            grvReg nmArq, "Y02|" & _
                        Rst3.Fields("cobr_nFat") & "|" & _
                        Rst3.Fields("cobr_vOrig") & "|" & _
                        Rst3.Fields("cobr_vDesc") & "|" & _
                        Rst3.Fields("cobr_vLiq") & "|"
        End If
        Rst3.MoveFirst
        For cCob = 0 To Rst3.RecordCount - 1
            If cNull(Rst3.Fields("cobr_nDup")) <> "" Then
                grvReg nmArq, "Y07|" & _
                            Rst3.Fields("cobr_nDup") & "|" & _
                            Format(Rst3.Fields("cobr_dVenc"), "YYYY-MM-DD") & "|" & _
                            Rst3.Fields("cobr_vDup") & "|"
            End If
            Rst3.MoveNext
        Next
    End If
        
'****************************************************************************************************
    grvReg nmArq, "Z||" & _
                    Rst1.Fields("InfAdic_InfCpl") & "|"



'========================================================================
    Exportar_NFe_v200_TXT = nmArq
    Rst1.Close
    Rst2.Close
    Rst3.Close
End Function

Private Sub MountTXT(Dados As String)
    If Len(Trim(strMountTXT)) = 0 Then
            strMountTXT = Dados
        Else
            strMountTXT = strMountTXT & vbCrLf & Dados
    End If
End Sub

Private Sub grvReg(nmArquivo As String, Dados As String)
    On Error GoTo TrtErro
    'define o ObjPreview filesystem e demais variaveis
    Dim fso As New FileSystemObject
    Dim Arquivo As File
    Dim arquivoLog As TextStream
    Dim msg As String
    Dim caminho As String
    
    
    'If Dir(SistemPath & "\Nfe", vbDirectory) = "" Then
    '    MkDir SistemPath & "\NFe"
    'End If

    caminho = PgDadosConfig.pFileArmazenamento & "\" & nmArquivo
    'se o arquivo não existir então cria
    If fso.FileExists(caminho) Then
            Set Arquivo = fso.GetFile(caminho)
        Else
            Set arquivoLog = fso.CreateTextFile(caminho)
            arquivoLog.Close
            'caminho = caminho & "\" & nmArquivo
            Set Arquivo = fso.GetFile(caminho)
    End If
    'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
    
    'monta informações para gerar a linha da mensagem
    msg = Dados

    'inclui linhas no arquivo texto
    arquivoLog.WriteLine msg
    
    
    'escreve uma linha em branco no arquivo - se voce quiser
    'arquivoLog.WriteBlankLines (1)
    'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing
    Exit Sub
TrtErro:
        MsgBox "Modulo: modulo_NFe" & vbCrLf & _
               "Erro ao gerar registro da NFe em Texto. " & _
               "Verifique o log!" & _
           vbCrLf & vbCrLf & _
           "Erro n.: " & Err.Number & _
           vbCrLf & vbCrLf & _
           "Descrição: " & Err.Description & _
           vbCrLf
    RegLogDataBase 0, "", "", "Gerar NFe.txt [" & Err.Description & "] - " & caminho
End Sub

Public Function ImprimirProtCanc(chvNFe As String, Optional ModalShow As Integer) As Boolean
    'True = OK
    'False = Erro
    
    Dim Rst         As Recordset
    Dim sSQL        As String
    'Dim pUniDANFe   As String
    Dim pXML        As String
    Dim sCMD        As String
    Dim idDest      As Integer
    
    ImprimirProtCanc = False
    
    If Trim(chvNFe) = "" Then
        MsgBox "Selecione uma Nf-e.", vbInformation, "Aviso"
        Exit Function
    End If
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND idNFE = '" & chvNFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "NFe não encontrada.", vbInformation, "Aviso"
            
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
    End If
    
    'pUniDANFe = PgDadosConfig.pUniDANFe & "\Unidanfe.exe"
    
    idDest = Rst.Fields("dest_idDest")
    Set rptNFeProtCanc.DataSource = Rst.DataSource
    rptNFeProtCanc.Show IIf(Trim(ModalShow) = "", 0, 1)
    Rst.Close
    'Rst1.Close
    
    'Localiza o XML da NF-e para impressao
    'If Trim(chvnfe) = "" Then Exit Function
    
    'pXML = PgDadosConfig.pBackup & "\Autorizados\" & Format(Rst.Fields("Ide_dEmi"), "YYYYMM") & "\" & chvnfe & "-procNFe.xml"
    
    'If Dir(pXML) = "" Then
    '        MsgBox "Arquivo XML não encontrado..."
    '        ImprimirDANFE = False
    '        Exit Function
    'End If
    
    'sCMD = pUniDANFe & " " & _
    "I=""" & "selecionar" & """ " & _
    "A=""" & pXML & """ " & _
    "CC=""" & IIf(IsNull(Rst.Fields("canc_nProt")), 0, 1) & """ " & _
    "V=""" & PgDadosConfig.DANFeVisualizar & """ " & _
    "L=""" & PgDadosEmpresa(ID_Empresa).Logotipo & """ " '& _
    IIf(PgDadosConfig.DANFeEnviarMail = 1, "E=""" & LCase(PgDadosCliente(idDest).emailnfe) & """ ", "") & _
    IIf(PgDadosConfig.DANFEeMailCC <> "", "EC=""" & LCase(PgDadosConfig.DANFEeMailCC) & """", "")
    'Shell sCMD, vbNormalFocus
'        "P=""" & IIf(PgDadosConfig.DANFenCopias = 0, 1, PgDadosConfig.DANFenCopias) & """ " &
    ImprimirProtCanc = True
End Function

Public Function ImprimirDANFE(chvNFe As String) As Boolean
    'True = OK
    'False = Erro
    On Error GoTo TrtErroImpDanfe
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim pUniDANFe   As String
    Dim pXML        As String
    Dim sCMD        As String
    Dim idDest      As Integer
    If Trim(chvNFe) = "" Then
        MsgBox "Selecione uma Nf-e.", vbInformation, "Aviso"
        ImprimirDANFE = False
        Exit Function
    End If
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND idNFE = '" & chvNFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "NFe não encontrada.", vbInformation, "Aviso"
            ImprimirDANFE = False
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
    End If
    
    pUniDANFe = PgDadosConfig.pUniDANFe & "\Unidanfe.exe"
    
    idDest = Rst.Fields("dest_idDest")
    
    'Localiza o XML da NF-e para impressao
    If Trim(chvNFe) = "" Then Exit Function
    
    pXML = PgDadosConfig.pBackup & "\Autorizados\" & Format(Rst.Fields("Ide_dEmi"), "YYYYMM") & "\" & chvNFe & "-procNFe.xml"
    
    If Dir(pXML) = "" Then
            MsgBox "Arquivo XML não encontrado..."
            ImprimirDANFE = False
            Exit Function
    End If
    
    'sCMD = pUniDANFe & " " & _
    "I=""" & "selecionar" & """ " & _
    "A=""" & pXML & """ " & _
    "CC=""" & IIf(IsNull(Rst.Fields("canc_nProt")), 0, 1) & """ " & _
    "V=""" & PgDadosConfig.DANFeVisualizar & """ " & _
    "L=""" & PgDadosEmpresa(ID_Empresa).Logotipo & """ " '& _
    IIf(PgDadosConfig.DANFeEnviarMail = 1, "E=""" & LCase(PgDadosCliente(idDest).emailnfe) & """ ", "") & _
    IIf(PgDadosConfig.DANFEeMailCC <> "", "EC=""" & LCase(PgDadosConfig.DANFEeMailCC) & """", "")
    
    sCMD = pUniDANFe & " " & _
    "I=""" & "selecionar" & """ " & _
    "A=""" & pXML & """ " & _
    "CC=""" & IIf(IsNull(Rst.Fields("canc_nProt")), 0, 1) & """ " & _
    "M=1 V=" & PgDadosConfig.DANFeVisualizar & " " & _
    "L=""" & PgDadosEmpresa(ID_Empresa).Logotipo & """ " '& _
    IIf(PgDadosConfig.DANFeEnviarMail = 1, "E=""" & LCase(PgDadosCliente(idDest).emailnfe) & """ ", "") & _
    IIf(PgDadosConfig.DANFEeMailCC <> "", "EC=""" & LCase(PgDadosConfig.DANFEeMailCC) & """", "")
    Shell sCMD, vbNormalFocus
'        "P=""" & IIf(PgDadosConfig.DANFenCopias = 0, 1, PgDadosConfig.DANFenCopias) & """ " &
    ImprimirDANFE = True
    Exit Function
TrtErroImpDanfe:
    RegLogDataBase 0, "0", "0", Err.Number & " - " & Err.Description
    ImprimirDANFE = False
End Function
Public Sub ImprimirDANFE2(chvNFe As String, Optional ModalShow As Integer)
    Dim Rst         As Recordset    'CAB
    Dim sSQL        As String       'CAB
    Dim Rst1        As Recordset    'CORPO
    Dim sSQL1       As String       'CORPO
    Dim Rst2        As Recordset    'FATURA
    Dim sSQL2       As String       'FATURA
    Dim pUniDANFe   As String
    Dim pXML        As String
    Dim sCMD        As String
    Dim idDest      As Integer
    Dim cCob        As Integer      'Conta o numero de duplicatas/faturas
    If Trim(chvNFe) = "" Then
        MsgBox "Selecione uma Nf-e.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    '****************************************************************************************
    'Cab
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND idNFE = '" & chvNFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "NFe não encontrada.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
    End If
    '******************************************************************************************
    'Corpo  coalesce
    'sSQL1 = "SELECT det_cProd,det_xProd,det_InfAdProd,det_NCM,ICMS_CST AS CST,det_CFOP,det_uCom,FORMAT(det_qCom," & cDecQtd & ") AS qCom,FORMAT(det_vUnCom," & cDecMoeda & ") AS vUnCom,det_vProd,ICMS_vBc,ICMS_vICMS,IPI_vIPI,ICMS_pICMS,IPI_pIPI,ID_Empresa" & _
            " FROM FaturamentoNFeItens WHERE ID_Empresa = " & ID_Empresa & " AND idNFE = '" & chvnfe & "'"
    sSQL1 = "SELECT det_cProd,CONCAT(det_xProd ,' ',COALESCE(det_InfAdProd,'')) AS dProd , det_NCM,CONCAT(CONVERT(ICMS_origem,CHAR),ICMS_CST) AS CST,det_CFOP,det_uCom,FORMAT(det_qCom," & cDecQtd & ") AS qCom,FORMAT(det_vUnCom," & cDecMoeda & ") AS vUnCom,det_vProd,ICMS_vBc,ICMS_vICMS,IPI_vIPI,ICMS_pICMS,IPI_pIPI,ID_Empresa" & _
            " FROM FaturamentoNFeItens WHERE ID_Empresa = " & ID_Empresa & " AND idNFE = '" & chvNFe & "'"
    Set Rst1 = RegistroBuscar(sSQL1)
    If Rst1.BOF And Rst1.EOF Then
            MsgBox "Campo Descrição da NFe não encontrada.", vbInformation, "Aviso"
            Exit Sub
        Else
            Rst1.MoveFirst
    End If
    '******************************************************************************************
    'Fatura
    sSQL2 = "SELECT cobr_nDup, cobr_dVenc, cobr_vDup,ID_Empresa " & _
            "FROM FaturamentoNFeCobranca " & _
            "WHERE ID_Empresa = " & ID_Empresa & " AND idNFE = '" & chvNFe & "'"
    Set Rst2 = RegistroBuscar(sSQL2)
    If Rst2.BOF And Rst2.EOF Then
            MsgBox "Campo Descrição da NFe não encontrada.", vbInformation, "Aviso"
            Exit Sub
        Else
            Rst2.MoveFirst
    End If
    '******************************************************************************************
    With rptDANFe.Sections("Section2").Controls
    
        If Not IsNull(Rst.Fields("canc_nProt")) Then
                .Item("imgCancelada").Visible = True
            Else
                .Item("imgCancelada").Visible = False
        End If
        
        .Item("lbltpNF").Caption = Rst.Fields("ide_tpNF")
        .Item("lblnNF").Caption = "N.° " & Rst.Fields("ide_nNF")
        .Item("lblSerie").Caption = "SERIE " & Rst.Fields("ide_Serie")
        
        .Item("lblCodBarras").Caption = "<" & Rst.Fields("idNFe") & ">"
        .Item("lblChaveAcesso").Caption = Format(Rst.Fields("idNFe"), "@@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@") 'Rst.Fields("idNFe")
        
        .Item("lblnatOp").Caption = Rst.Fields("ide_natOp")
        .Item("lblnProtocolo").Caption = IIf(IsNull(Rst.Fields("nProt")), "", Rst.Fields("nProt"))
        .Item("lbldhProtocolo").Caption = IIf(IsNull(Rst.Fields("dhProt")), "", Rst.Fields("dhProt"))
        
        
        .Item("lblIE").Caption = IIf(IsNull(Rst.Fields("emit_IE")), "", UCase(Rst.Fields("emit_IE")))
        .Item("lblIEST").Caption = IIf(IsNull(Rst.Fields("emit_IEST")), "", UCase(Rst.Fields("emit_IEST")))
        
        If Len(Rst.Fields("emit_CNPJ")) > 13 Then
                .Item("lblCNPJ").Caption = IIf(IsNull(Rst.Fields("emit_CNPJ")), "", UCase(Format(Rst.Fields("emit_CNPJ"), "@@.@@@.@@@/@@@@-@@")))
            Else
                .Item("lblCNPJ").Caption = IIf(IsNull(Rst.Fields("emit_CNPJ")), "", UCase(Rst.Fields("emit_CNPJ")))
        End If
        
        .Item("lblide_dEmi").Caption = IIf(IsNull(Rst.Fields("ide_dEmi")), "", UCase(Rst.Fields("ide_dEmi")))
        .Item("lblide_Saida").Caption = IIf(IsNull(Rst.Fields("ide_dSaiEnt")), "", UCase(Rst.Fields("ide_dSaiEnt")))
        .Item("lblide_hrSaida").Caption = IIf(IsNull(Rst.Fields("ide_hSaiEnt")), "", UCase(Rst.Fields("ide_hSaiEnt")))
        
        'Destinatario
        .Item("lbldest_Nome").Caption = IIf(IsNull(Rst.Fields("dest_xNome")), "", UCase(Rst.Fields("dest_xNome")))
        If Len(Rst.Fields("dest_CNPJ")) > 13 Then
                .Item("lbldest_CNPJ").Caption = IIf(IsNull(Rst.Fields("dest_CNPJ")), "", UCase(Format(Rst.Fields("dest_CNPJ"), "@@.@@@.@@@/@@@@-@@")))
            Else
                .Item("lbldest_CNPJ").Caption = IIf(IsNull(Rst.Fields("dest_CNPJ")), "", UCase(Rst.Fields("dest_CNPJ")))
        End If
        
        
        .Item("lbldest_Lgr").Caption = IIf(IsNull(Rst.Fields("dest_xLgr")), "", UCase(Rst.Fields("dest_xLgr"))) & " " & IIf(IsNull(Rst.Fields("dest_nro")), "", UCase(Rst.Fields("dest_nro"))) & " " & IIf(IsNull(Rst.Fields("dest_xCpl")), "", UCase(Rst.Fields("dest_xCpl")))
        .Item("lbldest_Bairro").Caption = IIf(IsNull(Rst.Fields("dest_Bairro")), "", UCase(Rst.Fields("dest_Bairro")))
        .Item("lbldest_CEP").Caption = IIf(IsNull(Rst.Fields("dest_CEP")), "", UCase(Rst.Fields("dest_CEP")))
        .Item("lbldest_Mun").Caption = IIf(IsNull(Rst.Fields("dest_xMun")), "", UCase(Rst.Fields("dest_xMun")))
        .Item("lbldest_UF").Caption = IIf(IsNull(Rst.Fields("dest_UF")), "", UCase(Rst.Fields("dest_UF")))
        .Item("lbldest_Fone").Caption = IIf(IsNull(Rst.Fields("dest_Fone")), "", UCase(Rst.Fields("dest_Fone")))
        .Item("lbldest_IE").Caption = IIf(IsNull(Rst.Fields("dest_IE")), "", UCase(Rst.Fields("dest_IE")))
        
        'Faturamento / Duplicatas
        If Rst.Fields("impfatura") = 1 Then
            For cCob = 1 To 9
                .Item("lblFat0" & cCob).Caption = ""
            Next
            If IsNull(Rst2.Fields("cobr_nDup")) = False Then
                cCob = 1
                Do Until Rst2.EOF
                    .Item("lblFat0" & cCob).Caption = IIf(IsNull(Rst2.Fields("cobr_nDup")), "0", Rst2.Fields("cobr_nDup")) & "  " & Rst2.Fields("cobr_dVenc") & "  " & ConvMoeda(IIf(IsNull(Rst2.Fields("cobr_vDup")), "0", Rst2.Fields("cobr_vDup")))
                    cCob = cCob + 1
                    Rst2.MoveNext
                Loop
            End If
        End If
        'Calculo Impostos
        .Item("lblvBC").Caption = IIf(IsNull(Rst.Fields("total_vBC")), ConvMoeda("0"), ConvMoeda(Rst.Fields("total_vBC")))
        .Item("lblvICMS").Caption = IIf(IsNull(Rst.Fields("total_vICMS")), ConvMoeda("0"), ConvMoeda(Rst.Fields("total_vICMS")))
        .Item("lblvBCST").Caption = ConvMoeda(IIf(IsNull(Rst.Fields("total_vBCST")), "0", Rst.Fields("total_vBCST")))
        .Item("lblvICMSST").Caption = IIf(IsNull(Rst.Fields("total_vICMSST")), ConvMoeda("0"), ConvMoeda(Rst.Fields("total_vICMSST")))
        .Item("lblvProd").Caption = IIf(IsNull(Rst.Fields("total_vProd")), ConvMoeda("0"), ConvMoeda(Rst.Fields("total_vProd")))
        
        .Item("lblvFrete").Caption = IIf(IsNull(Rst.Fields("total_vFrete")), ConvMoeda("0"), ConvMoeda(Rst.Fields("total_vFrete")))
        .Item("lblvSeguro").Caption = IIf(IsNull(Rst.Fields("total_vSeg")), ConvMoeda("0"), ConvMoeda(Rst.Fields("total_vSeg")))
        .Item("lblvDesconto").Caption = IIf(IsNull(Rst.Fields("total_vDesc")), ConvMoeda("0"), ConvMoeda(Rst.Fields("total_vDesc")))
        .Item("lblvOutras").Caption = IIf(IsNull(Rst.Fields("total_vOutro")), ConvMoeda("0"), ConvMoeda(Rst.Fields("total_vOutro")))
        .Item("lblvIPI").Caption = IIf(IsNull(Rst.Fields("total_vIPI")), ConvMoeda("0"), ConvMoeda(Rst.Fields("total_vIPI")))
        .Item("lblvNota").Caption = IIf(IsNull(Rst.Fields("total_vNF")), ConvMoeda("0"), ConvMoeda(Rst.Fields("total_vNF")))
        
        'Transportador
        .Item("lbltransp_nome").Caption = IIf(IsNull(Rst.Fields("transp_xNome")), "", UCase(Rst.Fields("transp_xNome")))
        .Item("lblFrete").Caption = IIf(Rst.Fields("transp_modFrete") = 0, "0-EMIT.", "1-DESTINAT.")
        .Item("lbltransp_CNPJ").Caption = IIf(IsNull(Rst.Fields("transp_CNPJ")), "", UCase(Rst.Fields("transp_CNPJ")))
        .Item("lbltransp_Endereco").Caption = IIf(IsNull(Rst.Fields("transp_xEnder")), "", UCase(Rst.Fields("transp_xEnder")))
        .Item("lbltransp_Municipio").Caption = IIf(IsNull(Rst.Fields("transp_xMun")), "", UCase(Rst.Fields("transp_xMun")))
        .Item("lbltransp_UF").Caption = IIf(IsNull(Rst.Fields("transp_UF")), "", UCase(Rst.Fields("transp_UF")))
        .Item("lbltransp_IE").Caption = IIf(IsNull(Rst.Fields("transp_IE")), "", UCase(Rst.Fields("transp_IE")))
        
        .Item("lbltransp_qVol").Caption = IIf(IsNull(Rst.Fields("transp_qVol")), "", UCase(Rst.Fields("transp_qVol")))
        .Item("lbltransp_esp").Caption = IIf(IsNull(Rst.Fields("transp_esp")), "", UCase(Rst.Fields("transp_esp")))
        
        .Item("lbltransp_Marca").Caption = IIf(IsNull(Rst.Fields("transp_Marca")), "", UCase(Rst.Fields("transp_Marca")))
        .Item("lbltransp_Num").Caption = IIf(IsNull(Rst.Fields("transp_nVol")), "", UCase(Rst.Fields("transp_nVol")))
        
        .Item("lbltransp_PesoB").Caption = IIf(IsNull(Rst.Fields("transp_PesoB")), "", UCase(Rst.Fields("transp_PesoB")))
        .Item("lbltransp_PesoL").Caption = IIf(IsNull(Rst.Fields("transp_PesoL")), "", UCase(Rst.Fields("transp_PesoL")))
        
        .Item("lbltransp_PLACACARRO").Caption = IIf(IsNull(Rst.Fields("transp_VeicPlaca")), "", UCase(Rst.Fields("transp_VeicPlaca")))
        .Item("lbltransp_UFCARRO").Caption = IIf(IsNull(Rst.Fields("transp_VeicUF")), "", UCase(Rst.Fields("transp_VeicUF")))
    End With
    
    With rptDANFe.Sections("Section3").Controls
        .Item("lblInfComp").Caption = IIf(IsNull(Rst.Fields("infAdic_InfCpl")), "", UCase(Rst.Fields("infAdic_InfCpl")))
        
        Dim textoc As String
        'lblCanhotoTexto=RECEBEMOS DE @EMPRESA OS PRODUTOS E/OU SERVIÇOS CONSTANTES DA NOTA FISCAL ELETRÔNICA
        '                INDICADA AO LADO. EMISSÃO: @EMISSAO VALOR TOTAL: @VALOR DESTINATARIO: @DESTINATARIO
        textoc = "RECEBEMOS DE @EMPRESA OS PRODUTOS E/OU SERVIÇOS CONSTANTES DA NOTA FISCAL ELETRÔNICA INDICADA AO LADO. EMISSÃO: @EMISSAO VALOR TOTAL: @VALOR DESTINATARIO: @DESTINATARIO"
        textoc = Replace(textoc, "@EMPRESA", PgDadosEmpresa(ID_Empresa).Nome)
        textoc = Replace(textoc, "@EMISSAO", Rst.Fields("ide_dEmi"))
        textoc = Replace(textoc, "@VALOR", Rst.Fields("total_vNF"))
        textoc = Replace(textoc, "@DESTINATARIO", Rst.Fields("dest_xNome") & "-" & Rst.Fields("dest_CNPJ"))
        .Item("lblCanhotoTexto").Caption = textoc
        
        .Item("lblCanhotonNF").Caption = IIf(IsNull(Rst.Fields("ide_nNF")), "", UCase("N.° " & Rst.Fields("ide_nNF")))
        .Item("lblCanhotoSerie").Caption = IIf(IsNull(Rst.Fields("ide_Serie")), "", UCase("SERIE " & Rst.Fields("ide_Serie")))
        
    End With
    
    Set rptDANFe.DataSource = Rst1.DataSource
    rptDANFe.Show IIf(Trim(ModalShow) = "", 0, 1)
    Rst.Close
    Rst1.Close
    Rst2.Close
End Sub


Public Sub ImprimirDANFEFornecedor(chvNFe As String)
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim pUniDANFe   As String
    Dim pXML        As String
    Dim sCMD        As String
    Dim idDest      As Integer
    If Trim(chvNFe) = "" Then
        MsgBox "Selecione uma Nf-e.", vbInformation, "Aviso"
        Exit Sub
    End If
    sSQL = "SELECT * FROM FaturamentoNFeEntrada WHERE ID_Empresa = " & ID_Empresa & " AND idNFE = '" & chvNFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "NFe não encontrada.", vbInformation, "Aviso"
            Exit Sub
        Else
            Rst.MoveFirst
    End If
    
    
    
    pUniDANFe = PgDadosConfig.pUniDANFe & "\Unidanfe.exe"
    
    'idDest = Rst.Fields("dest_idDest")
    
    'Localiza o XML da NF-e para impressao
    
    If Trim(chvNFe) = "" Then Exit Sub
    
    pXML = PgDadosConfig.pXMLFornecedor & "\" & Format(Rst.Fields("Ide_dEmi"), "YYYYMM") & "\NFe" & chvNFe & ".xml"
    
    If Dir(pXML) = "" Then
            MsgBox "Arquivo XML não encontrado..."
            Exit Sub
    End If
    
    sCMD = pUniDANFe & " " & _
    "I=""" & "selecionar" & """ " & _
    "A=""" & pXML & """ " & _
    "V=""1""" & _
    "L="""""
    Shell sCMD, vbNormalFocus

End Sub


Public Sub trocarChvAcesso(chvAntiga As String, novaChave As String)
    On Error GoTo trtErroLocal
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    '13.11.2014 - Verifica se a chave de acesso é valida
    sSQL = "SELECT * FROM faturamentonfe WHERE IdNFe='" & chvAntiga & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Chave " & chvAntiga & " não encontrada!", vbInformation, App.EXEName
            Rst.Close
            'Exit Sub
        Else
            Rst.Close
            
    End If
    
    Dim vDados(5)   As Variant
    'Dim cReg        As Integer
    
    vDados(0) = Array("IdNFe", novaChave, "S")
    
    RegistroAlterar "faturamentonfe", vDados, 0, "IdNFe='" & chvAntiga & "'"
    
    'sSQL = "UPDATE faturamentonfecobranca SET IdNFe = '" & chvnova & " WHERE IdNFe = '" & chvnova & "'"
    RegistroAlterar "faturamentonfecobranca", vDados, 0, "IdNFe='" & chvAntiga & "'"
    
    'sSQL = "UPDATE faturamentonfeitens SET IdNFe = '" & chvnova & " WHERE IdNFe = '" & chvnova & "'"
    RegistroAlterar "faturamentonfeitens", vDados, 0, "IdNFe='" & chvAntiga & "'"
    
    'sSQL = "UPDATE faturamentonfesendmail SET IdNFe = '" & chvnova & " WHERE IdNFe = '" & chvnova & "'"
    RegistroAlterar "faturamentonfesendmail", vDados, 0, "IdNFe='" & chvAntiga & "'"
    
    vDados(0) = Array("chvNFe", novaChave, "S")
    'sSQL = "UPDATE faturamentonfecartacorrecao SET chvNFe = '" & chvnova & " WHERE chvNFe = '" & chvnova & "'"
    RegistroAlterar "faturamentonfecartacorrecao", vDados, 0, "chvNFe='" & chvAntiga & "'"
    
    'sSQL = "UPDATE faturamentonfecartacorrecaoitens SET chvNFe = '" & chvnova & " WHERE chvNFe = '" & chvnova & "'"
    RegistroAlterar "faturamentonfecartacorrecaoitens", vDados, 0, "chvNFe='" & chvAntiga & "'"
    Exit Sub
trtErroLocal:
    MsgBox Err.Description, vbInformation, "Erro n." & Err.Number
    RegLogDataBase 0, Err.Number, "trocarChvAcesso", Err.Description
    
End Sub
Public Function Exportar_NFe_v310_TXT(chvNFe As String) As String
    Dim Rst1    As Recordset 'Cabecalho
    Dim Rst2    As Recordset 'Produto
    Dim Rst3    As Recordset 'Cobanca
    Dim sSQL    As String
    Dim nmArq   As String
    Dim cItens  As Integer 'Conta os registros dos itens da Nota
    Dim cCob    As Integer 'Conta os registros da cobranca da Nota
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & chvNFe & "'"
    Set Rst1 = RegistroBuscar(sSQL)
    If Rst1.BOF And Rst1.EOF Then
            MsgBox "Etapa 1 - Erro ao localizar NF-e"
            Exportar_NFe_v310_TXT = ""
            Exit Function
        Else
            Rst1.MoveFirst
    End If
    sSQL = "SELECT * FROM FaturamentoNFeItens WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & chvNFe & "'"
    Set Rst2 = RegistroBuscar(sSQL)
    If Rst2.BOF And Rst2.EOF Then
            MsgBox "Etapa 2 - Erro ao localizar NF-e"
            Exportar_NFe_v310_TXT = ""
            Exit Function
        Else
            Rst2.MoveFirst
    End If
    sSQL = "SELECT * FROM FaturamentoNFeCobranca WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & chvNFe & "'"
    Set Rst3 = RegistroBuscar(sSQL)
    If Rst3.BOF And Rst3.EOF Then
            MsgBox "Etapa 3 - Erro ao localizar NF-e"
            Exportar_NFe_v310_TXT = ""
            Exit Function
        Else
            Rst3.MoveFirst
    End If
     nmArq = Rst1.Fields("ide_nNF") & "_" & Rst1.Fields("emit_CNPJ") & "_" & Format(Rst1.Fields("ide_dEmi"), "DD") & "_" & Format(Rst1.Fields("ide_dEmi"), "MM") & "_" & Format(Rst1.Fields("ide_dEmi"), "YYYY") & "-nfe.txt"
     
     ChecarArquivo (nmArq)
'========================================================================
    
    grvReg nmArq, "NOTAFISCAL|1"
    'A
    grvReg nmArq, "A|" & Rst1.Fields("Versao") & _
                        "|NFe" & chvNFe & "|"
    'B
    grvReg nmArq, "B|" & _
                    Rst1.Fields("ide_cUF") & "|" & _
                    Rst1.Fields("ide_cNF") & "|" & _
                    Rst1.Fields("ide_NatOP") & "|" & _
                    Rst1.Fields("ide_indPag") & "|" & _
                    Rst1.Fields("ide_Mod") & "|" & _
                    Rst1.Fields("ide_serie") & "|" & _
                    CInt(Rst1.Fields("ide_nNF")) & "|" & _
                    Format(Rst1.Fields("ide_demi"), "YYYY-MM-DD") & "T" & Format(Rst1.Fields("ide_hemi"), "HH:MM:SS") & PgDadosConfig.fusoHorario & "|" & _
                    IIf(IsNull(Rst1.Fields("ide_dSaiEnt")), "", Format(Rst1.Fields("ide_dSaiEnt"), "YYYY-MM-DD") & "T" & Format(Rst1.Fields("ide_hSaiEnt"), "HH:MM:SS") & PgDadosConfig.fusoHorario) & "|" & _
                    Rst1.Fields("ide_tpNF") & "|" & _
                    Rst1.Fields("ide_idDest") & "|" & _
                    Rst1.Fields("ide_cMunFG") & "|" & _
                    Rst1.Fields("ide_tpImp") & "|" & _
                    Rst1.Fields("ide_tpEmis") & "|" & _
                    Rst1.Fields("ide_cDV") & "|" & _
                    Rst1.Fields("ide_tpAmb") & "|" & _
                    Rst1.Fields("ide_finNFe") & "|" & _
                    Rst1.Fields("ide_indFinal") & "|" & _
                    "3" & "|" & _
                    Rst1.Fields("ide_procEmi") & "|" & _
                    Rst1.Fields("ide_VerProc") & "|" & _
                    IIf(PgDadosConfig.ContingenciaDt <> "", Format(PgDadosConfig.ContingenciaDt, "YYYY-MM-DD") & "T" & PgDadosConfig.ContingenciaHr, "") & "|" & _
                    PgDadosConfig.ContingenciaMotivo & "|"
'                    "2012-05-28T11:39:30|" & _
                    "ERRO NA SVRS AUTORIZADO PELO SITE DA NFE. 28/05/2012" & "|"
                    
    If Not IsNull(Rst1.Fields("ide_refNFe")) Then
        grvReg nmArq, "B13|" & Rst1.Fields("ide_refNFe") & "|"
    End If
    
    'C - dados EMITENTE
    grvReg nmArq, "C|" & _
                    Rst1.Fields("emit_xNome") & "|" & _
                    Rst1.Fields("emit_xFant") & "|" & _
                    Rst1.Fields("emit_IE") & "|" & _
                    Rst1.Fields("emit_IEST") & "|" & _
                    Rst1.Fields("emit_IM") & "|" & _
                    IIf(Trim(Rst1.Fields("emit_IM")) <> "", Rst1.Fields("emit_CNAE"), "") & "|" & _
                    Rst1.Fields("emit_CRT") & "|"
                    
    grvReg nmArq, "C02|" & _
                    Rst1.Fields("emit_CNPJ") & "|"
    grvReg nmArq, "C05|" & _
                    Rst1.Fields("emit_xLgr") & "|" & _
                    Rst1.Fields("emit_nro") & "|" & _
                    Rst1.Fields("emit_xcpl") & "|" & _
                    Rst1.Fields("emit_Bairro") & "|" & _
                    Rst1.Fields("emit_cMun") & "|" & _
                    Rst1.Fields("emit_xMun") & "|" & _
                    Rst1.Fields("emit_UF") & "|" & _
                    Rst1.Fields("emit_CEP") & "|" & _
                    Rst1.Fields("emit_cPais") & "|" & _
                    Rst1.Fields("emit_xPais") & "|" & _
                    Rst1.Fields("emit_fone") & "|"
    'E - dados DESTINATARIO
                    
    grvReg nmArq, "E|" & _
                    Rst1.Fields("dest_xNome") & "|" & _
                    Rst1.Fields("dest_indIEDest") & "|" & _
                    IIf(IsNull(Rst1.Fields("dest_IE")), "", Rst1.Fields("dest_IE")) & "|" & _
                    Rst1.Fields("dest_ISUF") & "|" & _
                    "" & "|" & _
                    Rst1.Fields("dest_email") & "|"
                    
    grvReg nmArq, "E" & IIf(UCase(Rst1.Fields("dest_Pessoa")) = "FISICA", "03", "02") & "|" & _
                    Rst1.Fields("dest_CNPJ") & "|"
                    
    grvReg nmArq, "E05|" & _
                    Rst1.Fields("dest_xLgr") & "|" & _
                    Rst1.Fields("dest_nro") & "|" & _
                    Rst1.Fields("dest_xCpl") & "|" & _
                    Rst1.Fields("dest_Bairro") & "|" & _
                    Rst1.Fields("dest_cMun") & "|" & _
                    Rst1.Fields("dest_xMun") & "|" & _
                    Rst1.Fields("dest_UF") & "|" & _
                    Rst1.Fields("dest_CEP") & "|" & _
                    Rst1.Fields("dest_cPais") & "|" & _
                    Rst1.Fields("dest_xPais") & "|" & _
                    Rst1.Fields("dest_fone") & "|"
    
    'G - dadosEntrega
    If Trim(Rst1.Fields("dest_cnpj")) <> Trim(Rst1.Fields("entr_CNPJ")) Then
        grvReg nmArq, "G|" & _
                      Rst1.Fields("entr_xLgr") & "|" & _
                      Rst1.Fields("entr_nro") & "|" & _
                      Rst1.Fields("entr_xCpl") & "|" & _
                      Rst1.Fields("entr_xBairro") & "|" & _
                      Rst1.Fields("entr_cMun") & "|" & _
                      Rst1.Fields("entr_xMun") & "|" & _
                      Rst1.Fields("entr_UF") & "|"
        grvReg nmArq, "G02|" & _
                      Rst1.Fields("entr_CNPJ") & "|"
    End If
    'H/I - dados DESCRICAO DOS ITENS
    Rst2.MoveFirst
    For cItens = 0 To Rst2.RecordCount - 1
        grvReg nmArq, "H|" & cItens + 1 & "|" & _
                       IIf(Trim(Rst2.Fields("det_InfAdProd")) = "", "", Rst2.Fields("det_InfAdProd") & "|")
                       
                       '23.03.2015 - Campo vazio antes do EXTIPI ref NVE
        grvReg nmArq, "I|" & Rst2.Fields("det_cProd") & "|" & _
                        Rst2.Fields("det_cEAN") & "|" & _
                        Rst2.Fields("det_xProd") & "|" & _
                        Rst2.Fields("det_NCM") & "|" & _
                        Rst2.Fields("det_EXTIPI") & "|" & _
                        Rst2.Fields("det_CFOP") & "|" & _
                        Rst2.Fields("det_uCom") & "|" & _
                        Rst2.Fields("det_qCom") & "|" & _
                        Rst2.Fields("det_vUnCom") & "|" & _
                        Rst2.Fields("det_vprod") & "|" & _
                        Rst2.Fields("det_cEANTrib") & "|" & _
                        Rst2.Fields("det_uTrib") & "|" & _
                        Rst2.Fields("det_qTrib") & "|" & _
                        Rst2.Fields("det_vUnTrib") & "|" & _
                        Rst2.Fields("det_vFrete") & "|" & _
                        Rst2.Fields("det_vSeg") & "|" & _
                        Rst2.Fields("det_vDesc") & "|" & _
                        Rst2.Fields("det_vOutro") & "|" & _
                        Rst2.Fields("det_indTot") & "|" & _
                        Rst2.Fields("det_xPed") & "|" & _
                        Rst2.Fields("det_nItemPed") & "|"
        
        grvReg nmArq, "M|"
        grvReg nmArq, "N|"
        '*****************************************************************
        'ICMS ************************************************************
        Select Case Rst2.Fields("ICMS_CST")
            Case "00" 'Tributacao Integral (N02)
                grvReg nmArq, "N02|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|"
                                
                                
                
                
                '22.12.17 - Inclusao da tag para DIFAL
                If Rst2.Fields("ICMS_vBCUFDest") > 0 Then
                 grvReg nmArq, "NA|" & _
                                Rst2.Fields("ICMS_vBCUFDest") & "|" & _
                                Rst2.Fields("ICMS_pFCPUFDest") & "|" & _
                                Rst2.Fields("ICMS_pICMSUFDest") & "|" & _
                                Rst2.Fields("ICMS_pICMSInter") & "|" & _
                                Rst2.Fields("ICMS_pICMSInterPart") & "|" & _
                                Rst2.Fields("ICMS_vFCPUFDest") & "|" & _
                                Rst2.Fields("ICMS_vICMSUFDest") & "|" & _
                                Rst2.Fields("ICMS_vICMSUFRemet") & "|"
                End If
                
            Case "10" 'Tributada com cobranca ICMS (N03)
                grvReg nmArq, "N03|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
            Case "20" 'Tributacao do ICMS com reducao da Base de Calculo (N04)
                 grvReg nmArq, "N04|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|"
                                
            Case "30" 'Tributacao Isenta com cobranca de ICMS por ST (N05)
                grvReg nmArq, "N05|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
                                
            Case "40" 'Tributacao do ICMS ISENTA (N06)
                grvReg nmArq, "N06|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_MotDesICMS") & "|"
                                
            'Case "41" 'Tributacao do ICMS NAO TRIBUTADA ()
            
            Case "50" 'Tributacao do ICMS SUSPENSAO ()
                grvReg nmArq, "N06|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|"
            
            Case "51" 'Tributacao do ICMS POR DIFERIMENTO (N07)
                grvReg nmArq, "N07|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|"
            
            Case "60" 'ICMS cobrado anteriormente por ST (N08)
                grvReg nmArq, "N08" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
            Case "70" 'Tributacao do com reducao da base de calculo do ICMS ST (N09)
                grvReg nmArq, "N09" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
            Case "90" ' 'Tributacao OUTROS (N10)
                grvReg nmArq, "N10" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
            Case "101" 'Tributado pelo SN com permicao de credito (N10c)
             grvReg nmArq, "N10c" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_pCredSN") & "|" & _
                                Rst2.Fields("ICMS_vCredICMSSN") & "|"
            Case "102", "103" 'Tributado pelo SN sem permicao de credito (N10d)
             grvReg nmArq, "N10d" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                "0.00" & "|" & _
                                "0.00" & "|" & _
                                "0.00" & "|"
            
            Case "300", "400" 'Nao Tributado pelo SN (N10d)
             grvReg nmArq, "N10d" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                "0.00" & "|" & _
                                "0.00" & "|" & _
                                "0.00" & "|"
            
            Case "500" 'Tributacao de ICMS pelo SIMPLES NACIONAL (N10g)
             grvReg nmArq, "N10g" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                "0.00" & "|" & _
                                "0.00" & "|"
            Case "900" 'Outros
                'orig|CSOSN|modBC|vBC|pRedBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST
                'vICMSST|pCredSN|vCredICMSSN
                grvReg nmArq, "N10h|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                "0|0|0" & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"

        End Select
        '************************************************************
        'IPI ********************************************************
        grvReg nmArq, "O|" & _
                        "|" & _
                        "|" & _
                        "|" & _
                        "|" & _
                        Rst2.Fields("IPI_cEnq") & "|"
        grvReg nmArq, "O07|" & _
                        Rst2.Fields("IPI_CST") & "|" & _
                        Rst2.Fields("IPI_vIPI") & "|"
        grvReg nmArq, "O10|" & _
                        Rst2.Fields("IPI_vBC") & "|" & _
                        Rst2.Fields("IPI_pIPI") & "|"
                        
                        
                        
        'PIS ************************************************************
        grvReg nmArq, "Q|"
        Select Case Rst2.Fields("PIS_CST")
            Case "01", "02"  'Aliquota Normal/Aliquota Diferenciada (Q02)
                grvReg nmArq, "Q02|" & _
                                Rst2.Fields("PIS_CST") & "|" & _
                                Rst2.Fields("PIS_vBC") & "|" & _
                                Rst2.Fields("PIS_pPIS") & "|" & _
                                Rst2.Fields("PIS_vPIS") & "|"
            Case Else
                grvReg nmArq, "Q04|" & Rst2.Fields("PIS_CST") & "|"
                'MsgBox "Verificar o Codigo de exportacao do PIS da NFe - CODIGO DO CST DO PIS DESCONHECIDO"
        End Select
        
        
        
        'COFINS ************************************************************
        grvReg nmArq, "S|"
        Select Case Rst2.Fields("COFINS_CST")
            Case "01", "02"  'Aliquota Normal/Aliquota Diferenciada (Q02)
                grvReg nmArq, "S02|" & _
                                Rst2.Fields("COFINS_CST") & "|" & _
                                Rst2.Fields("COFINS_vBC") & "|" & _
                                Rst2.Fields("COFINS_pCOFINS") & "|" & _
                                Rst2.Fields("COFINS_vCOFINS") & "|"
            Case Else
                'MsgBox "Verificar o Codigo de exportacao do COFINS da NFe - CODIGO DO CST DO COFINS DESCONHECIDO"
                grvReg nmArq, "S04|" & Rst2.Fields("COFINS_CST") & "|"
        End Select
        
        Rst2.MoveNext
       
    Next

    '*********************************** TOTAIS DA NF-e *****************************************
    grvReg nmArq, "W|"
    grvReg nmArq, "W02|" & _
                    Rst1.Fields("total_vBC") & "|" & _
                    Rst1.Fields("total_vICMS") & "|" & _
                    "0.00" & "|" & _
                    IIf(IsNull(Rst1.Fields("total_vFCPUFDest")), "", Rst1.Fields("total_vFCPUFDest") & "|") & _
                    IIf(IsNull(Rst1.Fields("total_vICMSUFDest")), "", Rst1.Fields("total_vICMSUFDest") & "|") & _
                    IIf(IsNull(Rst1.Fields("total_vICMSUFRemet")), "", Rst1.Fields("total_vICMSUFRemet") & "|") & _
                    Rst1.Fields("total_vBCST") & "|" & _
                    Rst1.Fields("total_vICMSST") & "|" & _
                    Rst1.Fields("total_vProd") & "|" & _
                    Rst1.Fields("total_vFrete") & "|" & _
                    Rst1.Fields("total_vSeg") & "|" & _
                    Rst1.Fields("total_vDesc") & "|" & _
                    "0.00" & "|" & _
                    Rst1.Fields("total_vIPI") & "|" & _
                    Rst1.Fields("total_vPIS") & "|" & _
                    Rst1.Fields("total_vCOFINS") & "|" & _
                    Rst1.Fields("total_vOutro") & "|" & _
                    Rst1.Fields("total_vNF") & "|" & _
                    "0.00" & "|"
    
'    Dim msgCredICMSSN As String
'
'    If Rst1.Fields("total_vCredICMSSN") <> "0.00" Then
'
'        msgCredICMSSN = "PERMITE O APROVEITAMENTO DO CRÉDITO DE ICMS " & _
'                        "NO VALOR DE R$" & Rst1.Fields("total_vCredICMSSN") & " " & _
'                        "CORRESPONDENTE À ALÍQUOTA DE " & "%, " & _
'                        "NOS TERMOS DO ARTIGO 23 DA LC 123."
'
'    End If

    '*************************************** TRANSPORTE ********************************************
    grvReg nmArq, "X|" & _
                    Rst1.Fields("transp_ModFrete") & "|"
    
        grvReg nmArq, "X03|" & _
                    Rst1.Fields("transp_xNome") & "|" & _
                    Rst1.Fields("transp_IE") & "|" & _
                    Rst1.Fields("transp_xEnder") & "|" & _
                    Rst1.Fields("transp_xMun") & "|" & _
                    Rst1.Fields("transp_UF") & "|"
    
    grvReg nmArq, "X" & IIf(UCase(Rst1.Fields("transp_Pessoa")) = "FISICA", "05", "04") & "|" & _
                    Rst1.Fields("transp_CNPJ") & "|"
    If cNull(Rst1.Fields("transp_VeicPlaca")) <> "" Then
        grvReg nmArq, "X18|" & _
                    cNull(Rst1.Fields("transp_VeicPlaca")) & "|" & _
                    cNull(Rst1.Fields("transp_VeicUF")) & "|" & _
                    "" & "|"
    End If
    grvReg nmArq, "X26|" & _
                    Rst1.Fields("transp_qVol") & "|" & _
                    Rst1.Fields("transp_esp") & "|" & _
                    Rst1.Fields("transp_marca") & "|" & _
                    Rst1.Fields("transp_nVol") & "|" & _
                    Rst1.Fields("transp_PesoL") & "|" & _
                    Rst1.Fields("transp_PesoB") & "|"

    '***************************************** COBRANCA ********************************************
    If PgDadosNotaFiscal(chvNFe).ImpFatura = 1 Then
        grvReg nmArq, "Y|"
        If cNull(Rst3.Fields("cobr_nFat")) <> "" Then
            grvReg nmArq, "Y02|" & _
                        Rst3.Fields("cobr_nFat") & "|" & _
                        Rst3.Fields("cobr_vOrig") & "|" & _
                        Rst3.Fields("cobr_vDesc") & "|" & _
                        Rst3.Fields("cobr_vLiq") & "|"
        End If
        Rst3.MoveFirst
        For cCob = 0 To Rst3.RecordCount - 1
            If cNull(Rst3.Fields("cobr_nDup")) <> "" Then
                grvReg nmArq, "Y07|" & _
                            Rst3.Fields("cobr_nDup") & "|" & _
                            Format(Rst3.Fields("cobr_dVenc"), "YYYY-MM-DD") & "|" & _
                            Rst3.Fields("cobr_vDup") & "|"
            End If
            Rst3.MoveNext
        Next
    End If
        
'****************************************************************************************************
    grvReg nmArq, "Z||" & _
                    Rst1.Fields("InfAdic_InfCpl") & "|" ' & Trim(msgCredICMSSN) & "|"



'========================================================================
    Exportar_NFe_v310_TXT = nmArq
    Rst1.Close
    Rst2.Close
    Rst3.Close
End Function
Public Function Exportar_NFe_v400_TXT(chvNFe As String) As String
    Dim Rst1    As Recordset 'Cabecalho
    Dim Rst2    As Recordset 'Produto
    Dim Rst3    As Recordset 'Cobanca
    Dim sSQL    As String
    Dim nmArq   As String
    Dim cItens  As Integer 'Conta os registros dos itens da Nota
    Dim cCob    As Integer 'Conta os registros da cobranca da Nota
    Dim numNFe  As String ' Numero da nfe
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & chvNFe & "'"
    Set Rst1 = RegistroBuscar(sSQL)
    If Rst1.BOF And Rst1.EOF Then
            MsgBox "Etapa 1 - Erro ao localizar NF-e"
            Exportar_NFe_v400_TXT = ""
            Exit Function
        Else
            Rst1.MoveFirst
    End If
    sSQL = "SELECT * FROM FaturamentoNFeItens WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & chvNFe & "'"
    Set Rst2 = RegistroBuscar(sSQL)
    If Rst2.BOF And Rst2.EOF Then
            MsgBox "Etapa 2 - Erro ao localizar NF-e"
            Exportar_NFe_v400_TXT = ""
            Exit Function
        Else
            Rst2.MoveFirst
    End If
    sSQL = "SELECT * FROM FaturamentoNFeCobranca WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & chvNFe & "'"
    Set Rst3 = RegistroBuscar(sSQL)
    If Rst3.BOF And Rst3.EOF Then
            MsgBox "Etapa 3 - Erro ao localizar NF-e"
            Exportar_NFe_v400_TXT = ""
            Exit Function
        Else
            Rst3.MoveFirst
    End If
    
    numNFe = Rst1.Fields("ide_nNF")
    nmArq = Rst1.Fields("ide_nNF") & "_" & Rst1.Fields("emit_CNPJ") & "_" & Format(Rst1.Fields("ide_dEmi"), "DD") & "_" & Format(Rst1.Fields("ide_dEmi"), "MM") & "_" & Format(Rst1.Fields("ide_dEmi"), "YYYY") & "-nfe.txt"
     
    ChecarArquivo (nmArq)
'========================================================================
    strMountTXT = ""
    MountTXT "NOTAFISCAL|1"
    'A
    MountTXT "A|" & Rst1.Fields("Versao") & _
                        "|NFe" & chvNFe & "|"
    'B
    ' Rst1.Fields("ide_indPag") & "|"
    'Indice do intermediador
    Dim indIntermed As String
    indIntermed = 0
    
    MountTXT "B|" & _
                    Rst1.Fields("ide_cUF") & "|" & _
                    Rst1.Fields("ide_cNF") & "|" & _
                    Rst1.Fields("ide_NatOP") & "|" & _
                    Rst1.Fields("ide_Mod") & "|" & _
                    Rst1.Fields("ide_serie") & "|" & _
                    CInt(Rst1.Fields("ide_nNF")) & "|" & _
                    Format(Rst1.Fields("ide_demi"), "YYYY-MM-DD") & "T" & Format(Rst1.Fields("ide_hemi"), "HH:MM:SS") & PgDadosConfig.fusoHorario & "|" & _
                    IIf(IsNull(Rst1.Fields("ide_dSaiEnt")), "", Format(Rst1.Fields("ide_dSaiEnt"), "YYYY-MM-DD") & "T" & Format(Rst1.Fields("ide_hSaiEnt"), "HH:MM:SS") & PgDadosConfig.fusoHorario) & "|" & _
                    Rst1.Fields("ide_tpNF") & "|" & _
                    Rst1.Fields("ide_idDest") & "|" & _
                    Rst1.Fields("ide_cMunFG") & "|" & _
                    Rst1.Fields("ide_tpImp") & "|" & _
                    Rst1.Fields("ide_tpEmis") & "|" & _
                    Rst1.Fields("ide_cDV") & "|" & _
                    Rst1.Fields("ide_tpAmb") & "|" & _
                    Rst1.Fields("ide_finNFe") & "|" & _
                    Rst1.Fields("ide_indFinal") & "|" & _
                    "3" & "|" & indIntermed & "|" & _
                    Rst1.Fields("ide_procEmi") & "|" & _
                    Rst1.Fields("ide_VerProc") & "|" & _
                    IIf(PgDadosConfig.ContingenciaDt <> "", Format(PgDadosConfig.ContingenciaDt, "YYYY-MM-DD") & "T" & PgDadosConfig.ContingenciaHr, "") & "|" & _
                    PgDadosConfig.ContingenciaMotivo & "|"
'                    "2012-05-28T11:39:30|" & _
                    "ERRO NA SVRS AUTORIZADO PELO SITE DA NFE. 28/05/2012" & "|"
                    
    If Not IsNull(Rst1.Fields("ide_refNFe")) Then
        MountTXT "B13|" & Rst1.Fields("ide_refNFe") & "|"
    End If
    
    'C - dados EMITENTE
    MountTXT "C|" & _
                    Rst1.Fields("emit_xNome") & "|" & _
                    Rst1.Fields("emit_xFant") & "|" & _
                    Rst1.Fields("emit_IE") & "|" & _
                    Rst1.Fields("emit_IEST") & "|" & _
                    Rst1.Fields("emit_IM") & "|" & _
                    IIf(Trim(Rst1.Fields("emit_IM")) <> "", Rst1.Fields("emit_CNAE"), "") & "|" & _
                    Rst1.Fields("emit_CRT") & "|"
                    
    MountTXT "C02|" & _
                    Rst1.Fields("emit_CNPJ") & "|"
    MountTXT "C05|" & _
                    Rst1.Fields("emit_xLgr") & "|" & _
                    Rst1.Fields("emit_nro") & "|" & _
                    Rst1.Fields("emit_xcpl") & "|" & _
                    Rst1.Fields("emit_Bairro") & "|" & _
                    Rst1.Fields("emit_cMun") & "|" & _
                    Rst1.Fields("emit_xMun") & "|" & _
                    Rst1.Fields("emit_UF") & "|" & _
                    Rst1.Fields("emit_CEP") & "|" & _
                    Rst1.Fields("emit_cPais") & "|" & _
                    Rst1.Fields("emit_xPais") & "|" & _
                    Rst1.Fields("emit_fone") & "|"
    'E - dados DESTINATARIO
                    
    MountTXT "E|" & _
                    Rst1.Fields("dest_xNome") & "|" & _
                    Rst1.Fields("dest_indIEDest") & "|" & _
                    IIf(IsNull(Rst1.Fields("dest_IE")), "", Rst1.Fields("dest_IE")) & "|" & _
                    Rst1.Fields("dest_ISUF") & "|" & _
                    "" & "|" & _
                    Rst1.Fields("dest_email") & "|"
                    
    MountTXT "E" & IIf(UCase(Rst1.Fields("dest_Pessoa")) = "FISICA", "03", "02") & "|" & _
                    Rst1.Fields("dest_CNPJ") & "|"
                    
    MountTXT "E05|" & _
                    Rst1.Fields("dest_xLgr") & "|" & _
                    Rst1.Fields("dest_nro") & "|" & _
                    Rst1.Fields("dest_xCpl") & "|" & _
                    Rst1.Fields("dest_Bairro") & "|" & _
                    Rst1.Fields("dest_cMun") & "|" & _
                    Rst1.Fields("dest_xMun") & "|" & _
                    Rst1.Fields("dest_UF") & "|" & _
                    Rst1.Fields("dest_CEP") & "|" & _
                    Rst1.Fields("dest_cPais") & "|" & _
                    Rst1.Fields("dest_xPais") & "|" & _
                    Rst1.Fields("dest_fone") & "|"
    
    'G - dadosEntrega
    If Trim(Rst1.Fields("dest_cnpj")) <> Trim(Rst1.Fields("entr_CNPJ")) Then
        MountTXT "G|" & _
                      Rst1.Fields("entr_xLgr") & "|" & _
                      Rst1.Fields("entr_nro") & "|" & _
                      Rst1.Fields("entr_xCpl") & "|" & _
                      Rst1.Fields("entr_xBairro") & "|" & _
                      Rst1.Fields("entr_cMun") & "|" & _
                      Rst1.Fields("entr_xMun") & "|" & _
                      Rst1.Fields("entr_UF") & "|"
        MountTXT "G02|" & _
                      Rst1.Fields("entr_CNPJ") & "|"
    End If
    'H/I - dados DESCRICAO DOS ITENS
    Rst2.MoveFirst
    For cItens = 0 To Rst2.RecordCount - 1
        MountTXT "H|" & cItens + 1 & "|" & _
                       IIf(Trim(Rst2.Fields("det_InfAdProd")) = "", "", Rst2.Fields("det_InfAdProd") & "|")
                       
                       '23.03.2015 - Campo vazio antes do EXTIPI ref NVE
                       'NFe 4.0
                       'I|cProd|cEAN|XProd|NCM|NVE|CEST|indEscala|CNPJFab|cBenef|
                       'EXTIPI|CFOP|UCom|QCom|VUnCom|VProd|CEANTrib|UTrib|QTrib|
                       'VUnTrib|VFrete|VSeg|VDesc|vOutro|indTot|xPed|nItemPed|nFCI|
                        '
        MountTXT "I|" & Rst2.Fields("det_cProd") & "|" & _
                        Rst2.Fields("det_cEAN") & "|" & _
                        Rst2.Fields("det_xProd") & "|" & _
                        Rst2.Fields("det_NCM") & "|" & _
                        "|" & _
                        Rst2.Fields("det_cest") & "|" & _
                        "|||" & _
                        Rst2.Fields("det_EXTIPI") & "|" & _
                        Rst2.Fields("det_CFOP") & "|" & _
                        Rst2.Fields("det_uCom") & "|" & _
                        Rst2.Fields("det_qCom") & "|" & _
                        Rst2.Fields("det_vUnCom") & "|" & _
                        Rst2.Fields("det_vprod") & "|" & _
                        Rst2.Fields("det_cEANTrib") & "|" & _
                        Rst2.Fields("det_uTrib") & "|" & _
                        Rst2.Fields("det_qTrib") & "|" & _
                        Rst2.Fields("det_vUnTrib") & "|" & _
                        Rst2.Fields("det_vFrete") & "|" & _
                        Rst2.Fields("det_vSeg") & "|" & _
                        Rst2.Fields("det_vDesc") & "|" & _
                        Rst2.Fields("det_vOutro") & "|" & _
                        Rst2.Fields("det_indTot") & "|" & _
                        Rst2.Fields("det_xPed") & "|" & _
                        Rst2.Fields("det_nItemPed") & "|" & _
                        "|"

        
        MountTXT "M|"
        MountTXT "N|"
        '*****************************************************************
        'ICMS ************************************************************
        Select Case Rst2.Fields("ICMS_CST")
            Case "00" 'Tributacao Integral (N02)
                MountTXT "N02|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_pFCP") & "|" & _
                                Rst2.Fields("ICMS_vFCP") & "|"
                                
                                
                
                
                '22.12.17 - Inclusao da tag para DIFAL
                If Rst2.Fields("ICMS_vBCUFDest") > 0 Then
                 MountTXT "NA|" & _
                                Rst2.Fields("ICMS_vBCUFDest") & "|" & _
                                Rst2.Fields("ICMS_vBCUFDest") & "|" & _
                                Rst2.Fields("ICMS_pFCPUFDest") & "|" & _
                                Rst2.Fields("ICMS_pICMSUFDest") & "|" & _
                                Rst2.Fields("ICMS_pICMSInter") & "|" & _
                                Rst2.Fields("ICMS_pICMSInterPart") & "|" & _
                                Rst2.Fields("ICMS_vFCPUFDest") & "|" & _
                                Rst2.Fields("ICMS_vICMSUFDest") & "|" & _
                                Rst2.Fields("ICMS_vICMSUFRemet") & "|"
                End If
                
            Case "10" 'Tributada com cobranca ICMS (N03)
                MountTXT "N03|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
            Case "20" 'Tributacao do ICMS com reducao da Base de Calculo (N04)
                      'N04|orig|CST|modBC|pRedBC|vBC|pICMS|vICMS|vBCFCP|pFCP|vFCP|vICMSDeson|motDesICMS
                 MountTXT "N04|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pFCP") & "|" & _
                                Rst2.Fields("ICMS_vFCP") & "|||"
                                
            Case "30" 'Tributacao Isenta com cobranca de ICMS por ST (N05)
                MountTXT "N05|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
                                
            Case "40" 'Tributacao do ICMS ISENTA (N06)
                MountTXT "N06|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_MotDesICMS") & "|"
                                
            Case "41" 'Tributacao do ICMS NAO TRIBUTADA ()
             MountTXT "N06|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_MotDesICMS") & "|"
            
            Case "50" 'Tributacao do ICMS SUSPENSAO ()
                MountTXT "N06|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|"
            
            Case "51" 'Tributacao do ICMS POR DIFERIMENTO (N07)
                MountTXT "N07|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|"
            
            Case "60" 'ICMS cobrado anteriormente por ST (N08)
                MountTXT "N08" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|" & _
                                "0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|"
            Case "70" 'Tributacao do com reducao da base de calculo do ICMS ST (N09)
                MountTXT "N09" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
            Case "90" ' 'Tributacao OUTROS (N10)
                MountTXT "N10" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_pRedBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                Rst2.Fields("ICMS_ModBCST") & "|" & _
                                Rst2.Fields("ICMS_pMVAST") & "|" & _
                                Rst2.Fields("ICMS_pRedBCST") & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
            Case "101" 'Tributado pelo SN com permicao de credito (N10c)
             MountTXT "N10c" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_pCredSN") & "|" & _
                                Rst2.Fields("ICMS_vCredICMSSN") & "|"
            Case "102", "103" 'Tributado pelo SN sem permicao de credito (N10d)
             MountTXT "N10d" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                "0.00" & "|" & _
                                "0.00" & "|" & _
                                "0.00" & "|"
            
            Case "300", "400" 'Nao Tributado pelo SN (N10d)
             MountTXT "N10d" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                "0.00" & "|" & _
                                "0.00" & "|" & _
                                "0.00" & "|"
            
            Case "500" 'Tributacao de ICMS pelo SIMPLES NACIONAL (N10g)
                       'Formato valido em 01.06.21 'N10g|0|500|0|0.00|0.00|0|0.00|0.00|0.00|
             MountTXT "N10g" & "|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                "0.00|0.00|0|0.00|0.00|0.00" & "|"
                                
                                
            Case "900" 'Outros
                'orig|CSOSN|modBC|vBC|pRedBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST
                'vICMSST|pCredSN|vCredICMSSN
                MountTXT "N10h|" & _
                                Rst2.Fields("ICMS_Origem") & "|" & _
                                Rst2.Fields("ICMS_CST") & "|" & _
                                Rst2.Fields("ICMS_ModBC") & "|" & _
                                Rst2.Fields("ICMS_vBC") & "|" & _
                                "|" & _
                                Rst2.Fields("ICMS_pICMS") & "|" & _
                                Rst2.Fields("ICMS_vICMS") & "|" & _
                                "0|0|0" & "|" & _
                                Rst2.Fields("ICMS_vBCST") & "|" & _
                                Rst2.Fields("ICMS_pICMSST") & "|" & _
                                Rst2.Fields("ICMS_vICMSST") & "|"
                                
                                
                                

        End Select
        '************************************************************
        'IPI ********************************************************
        'O|CNPJProd|cSelo|qSelo|cEnq|
        MountTXT "O||||" & _
                        Rst2.Fields("IPI_cEnq") & "|"
                        
        MountTXT "O07|" & _
                        Rst2.Fields("IPI_CST") & "|" & _
                        Rst2.Fields("IPI_vIPI") & "|"
        MountTXT "O10|" & _
                        Rst2.Fields("IPI_vBC") & "|" & _
                        Rst2.Fields("IPI_pIPI") & "|"
                        
                        
                        
        'PIS ************************************************************
        MountTXT "Q|"
        Select Case Rst2.Fields("PIS_CST")
            Case "01", "02"  'Aliquota Normal/Aliquota Diferenciada (Q02)
                MountTXT "Q02|" & _
                                Rst2.Fields("PIS_CST") & "|" & _
                                Rst2.Fields("PIS_vBC") & "|" & _
                                Rst2.Fields("PIS_pPIS") & "|" & _
                                Rst2.Fields("PIS_vPIS") & "|"
            Case Else
                MountTXT "Q04|" & Rst2.Fields("PIS_CST") & "|"
                'MsgBox "Verificar o Codigo de exportacao do PIS da NFe - CODIGO DO CST DO PIS DESCONHECIDO"
        End Select
        
        
        
        'COFINS ************************************************************
        MountTXT "S|"
        Select Case Rst2.Fields("COFINS_CST")
            Case "01", "02"  'Aliquota Normal/Aliquota Diferenciada (Q02)
                MountTXT "S02|" & _
                                Rst2.Fields("COFINS_CST") & "|" & _
                                Rst2.Fields("COFINS_vBC") & "|" & _
                                Rst2.Fields("COFINS_pCOFINS") & "|" & _
                                Rst2.Fields("COFINS_vCOFINS") & "|"
            Case Else
                'MsgBox "Verificar o Codigo de exportacao do COFINS da NFe - CODIGO DO CST DO COFINS DESCONHECIDO"
                MountTXT "S04|" & Rst2.Fields("COFINS_CST") & "|"
        End Select
        
        Rst2.MoveNext
       
    Next

    '*********************************** TOTAIS DA NF-e *****************************************
    Dim w As String
    MountTXT "W|"
        w = "W02|"
        w = w & Rst1.Fields("total_vBC") & "|"
        w = w & Rst1.Fields("total_vICMS") & "|"
        w = w & "0.00" & "|" 'vICMSDeson
        w = w & Rst1.Fields("total_vFCP") & "|"
        w = w & IIf(IsNull(Rst1.Fields("total_vFCPUFDest")), "0.00", Rst1.Fields("total_vFCPUFDest") & "|")
        w = w & IIf(IsNull(Rst1.Fields("total_vICMSUFDest")), "0.00", Rst1.Fields("total_vICMSUFDest") & "|")
        w = w & IIf(IsNull(Rst1.Fields("total_vICMSUFRemet")), "0.00", Rst1.Fields("total_vICMSUFRemet") & "|")
        
        w = w & Rst1.Fields("total_vBCST") & "|"
        w = w & Rst1.Fields("total_vICMSST") & "|"
        w = w & "0.00" & "|" 'vFCPST
        w = w & "0.00" & "|" 'vFCPSTRet
        w = w & Rst1.Fields("total_vProd") & "|"
        w = w & Rst1.Fields("total_vFrete") & "|"
        w = w & Rst1.Fields("total_vSeg") & "|"
        w = w & Rst1.Fields("total_vDesc") & "|"
        w = w & "0.00" & "|" 'vII
        w = w & Rst1.Fields("total_vIPI") & "|"
        w = w & "0.00" & "|" 'vIPIDevol
        w = w & Rst1.Fields("total_vPIS") & "|"
        w = w & Rst1.Fields("total_vCOFINS") & "|"
        w = w & Rst1.Fields("total_vOutro") & "|"
        w = w & Rst1.Fields("total_vNF") & "|"
        w = w & "0.00" & "|" 'vTotTrib
    MountTXT w
'    Dim msgCredICMSSN As String
'
'    If Rst1.Fields("total_vCredICMSSN") <> "0.00" Then
'
'        msgCredICMSSN = "PERMITE O APROVEITAMENTO DO CRÉDITO DE ICMS " & _
'                        "NO VALOR DE R$" & Rst1.Fields("total_vCredICMSSN") & " " & _
'                        "CORRESPONDENTE À ALÍQUOTA DE " & "%, " & _
'                        "NOS TERMOS DO ARTIGO 23 DA LC 123"
'
'    End If

    '*************************************** TRANSPORTE ********************************************
    MountTXT "X|" & _
                    Rst1.Fields("transp_ModFrete") & "|"
    
        MountTXT "X03|" & _
                    Rst1.Fields("transp_xNome") & "|" & _
                    Rst1.Fields("transp_IE") & "|" & _
                    Rst1.Fields("transp_xEnder") & "|" & _
                    Rst1.Fields("transp_xMun") & "|" & _
                    Rst1.Fields("transp_UF") & "|"
    
    MountTXT "X" & IIf(UCase(Rst1.Fields("transp_Pessoa")) = "FISICA", "05", "04") & "|" & _
                    Rst1.Fields("transp_CNPJ") & "|"
    If cNull(Rst1.Fields("transp_VeicPlaca")) <> "" Then
        MountTXT "X18|" & _
                    cNull(Rst1.Fields("transp_VeicPlaca")) & "|" & _
                    cNull(Rst1.Fields("transp_VeicUF")) & "|" & _
                    "" & "|"
    End If
    MountTXT "X26|" & _
                    Rst1.Fields("transp_qVol") & "|" & _
                    Rst1.Fields("transp_esp") & "|" & _
                    Rst1.Fields("transp_marca") & "|" & _
                    Rst1.Fields("transp_nVol") & "|" & _
                    Rst1.Fields("transp_PesoL") & "|" & _
                    Rst1.Fields("transp_PesoB") & "|"

    '***************************************** COBRANCA ********************************************
    If PgDadosNotaFiscal(chvNFe).ImpFatura = 1 Then
        'Fatura
        Dim tpPag As String
        
        MountTXT "Y|"
        
        tpPag = Trim(Left(pgDadosTipoDocumento(Rst3.Fields("cobr_TpDoc")).formaPgto, 3))
        
        If Trim(tpPag) = "90" Then
                MountTXT "YA|" & ZE(Trim(Rst1.Fields("ide_indPag")), 2) & "|" & tpPag

            Else
        
                If cNull(Rst3.Fields("cobr_nFat")) <> "" Then
                    MountTXT "Y02|" & _
                                Rst3.Fields("cobr_nFat") & "|" & _
                                Rst3.Fields("cobr_vOrig") & "|" & _
                                IIf(cNull(Rst3.Fields("cobr_vDesc")) = "", "0.00", Rst3.Fields("cobr_vDesc")) & "|" & _
                                Rst3.Fields("cobr_vLiq") & "|"
                End If
                Rst3.MoveFirst
                'Parcelas
                Dim parcela As String
                For cCob = 0 To Rst3.RecordCount - 1
                    If cNull(Rst3.Fields("cobr_nDup")) <> "" Then
                                       
                        tpPag = Trim(Left(pgDadosTipoDocumento(Rst3.Fields("cobr_TpDoc")).formaPgto, 3))

                        parcela = "Y07|"
                        parcela = parcela & Left("000", 3 - Len(Trim(cCob + 1))) & cCob + 1 & "|"
                        'parcela = parcela & Trim(Rst3.Fields("cobr_nDup")) & "|"
                        parcela = parcela & Format(Rst3.Fields("cobr_dVenc"), "YYYY-MM-DD") & "|"
                        parcela = parcela & Rst3.Fields("cobr_vDup") & "|"
                        MountTXT parcela
                        parcela = ""
                
                        '13.07.2018 - Conferme orientacao do grupo UNINFe
                        'YA|indPag|tPag|vPag|CNPJ|tBand|cAut|tpIntegra|vTroco
                
                        MountTXT "YA|" & _
                                Trim(Rst1.Fields("ide_indPag")) & "|" & _
                                tpPag & "|" & _
                                IIf(tpPag = "90", "0.00", Rst3.Fields("cobr_vDup")) & "|" & "|||||"
            
                    End If
                    Rst3.MoveNext
                Next
        End If
    End If
        
'****************************************************************************************************
    MountTXT "Z||" & _
                    Rst1.Fields("InfAdic_InfCpl") & "|" ' & Trim(msgCredICMSSN) & "|"



'========================================================================
    grvReg nmArq, strMountTXT

    'data lake included
    dataLakeInputNFe numNFe, chvNFe, strMountTXT

    Exportar_NFe_v400_TXT = nmArq
    Rst1.Close
    Rst2.Close
    Rst3.Close
End Function


