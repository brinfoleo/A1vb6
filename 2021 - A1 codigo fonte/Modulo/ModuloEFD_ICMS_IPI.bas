Attribute VB_Name = "ModuloEFD_ICMS_IPI"
Option Explicit
'Type Reg0000
'    nReg As Integer
'End Type
'Private Function Reg0000() As Reg0000
'    Reg0000.nReg = 30
'End Function
Dim vReg(100)   As Variant 'Array bidimencional para acumular os registros com TipoReg e Qtd.
Dim cArray      As Integer ' conta as arrais montadas
Dim nmFile      As String 'Nome do arquivo
Dim sR          As String 'Variavel de armazena do reg a ser gravado
Dim dtInicial   As Date
Dim dtFinal     As Date

Dim infInvent   As Boolean
Private Sub bloco0(op As Boolean)
'###############################################################################################
'### BLOCO 0 - ABERTURA, IDENTIFICACAO E REFERENCIAS
'###############################################################################################
  
    Registro0000 'dtInicial, dtFinal
    
    ' 0001 -ABERTURA DO BLOCO 0
    Registro0001

    ' "0005" 'DADOS COMPLEMENTARES
    Registro0005
    
    '"0100" 'DADOS DO CONTABILISTA
    Registro0100
    
    'REGISTRO 150 : TABELA DO CADASTRO DO PARTICIPANTE
    Registro0150 True
    
    'REGISTRO 0190: IDENTIFICAÇÃO DAS UNIDADES DE MEDIDA
    Registro0190
    
    'REGISTRO 0200: TABELA DE IDENTIFICAÇÃO DO ITEM (PRODUTO E SERVICOS)
    Registro0200 infInvent ', CDate(dtInicial), CDate(dtFinal)
    
    'REGISTRO 0400: TABELA DE NATUREZA DA OPERACAO/PRESTACAO
    Registro0400 op
    
    'REGISTRO 0990: ENCERRAMENTO DO BLOCO 0
    Registro0990
End Sub

Private Sub bloco1()
'###############################################################################################
'### BLOCO 1: OUTRAS INFORMACOES
'###############################################################################################
    
    'REGISTRO 1001: ABERTURA DO BLOCO 1
    Registro1001
    
    'REGISTRO 1010: OBRIGATORIEDADE DE REGISTROS DO BLOCO 1
    Registro1010
    
    'REGISTRO 1990: ENCERRAMENTO DO BLOCO 1
    Registro1990
End Sub

Private Sub bloco9()

'###############################################################################################
'### BLOCO 9
'###############################################################################################
    
    'REGISTRO 9001: ABERTURA DO BLOCO 9
    Registro9001
    
    'REGISTRO 9900: REGISTROS DO ARQUIVO
    Registro9900
    
    'REGISTRO 9990: ENCERRAMENTO DO BLOCO 9
    Registro9990
    
    'REGISTRO 9999: ENCERRAMENTO DO ARQUIVO DIGITAL
    Registro9999
    
End Sub


Private Sub blocoC(op As Boolean)

    'REGISTRO C001: ABERTURA DO BLOCO C
    RegistroC001 op
    If op = True Then
        'REGISTRO C100: REGISTRO DE NOTAS FISCAIS
        RegistroC100 'CDate(dtInicial), CDate(dtFinal)
    End If
     
    'REGISTRO C990: ENCERRAMENTO DO BLOCO C
    RegistroC990
End Sub
Private Sub blocoD(op As Boolean)
'###############################################################################################
'### BLOCO D: DOCUMENTOS FISCAIS II - SERVICOS ICMS
'###############################################################################################
    
    'REGISTRO D001: ABERTURA DO BLOCO D
    RegistroD001
    
    'REGISTRO D990: ENCERRAMENTO DO BLOCO D
    RegistroD990

End Sub


Private Sub blocoE(op As Boolean)
'###############################################################################################
'### BLOCO E: APURACAO DO ICMS E DO IPI
'###############################################################################################
    
    'REGISTRO E001: ABERTURA DO BLOCO E
    RegistroE001
    
    'REGISTRO E100: PERIODO DA APURACAO DO ICMS
    RegistroE100 'dtInicial, dtFinal
       
    'REGISTRO E110: APURACAO DO ICMS OPERACOES PROPRIAS
    RegistroE110 op 'CDate(dtInicial), CDate(dtFinal)
    
    'REGISTRO E500: PERIODO DE APURACAO DO IPI
    RegistroE500 ' dtInicial, dtFinal
    
    If op = True Then
         
        
        'REGISTRO E510: CONSOLIDACAO DOS VALORES DO IPI
        RegistroE510  'CDate(dtInicial), CDate(dtFinal)
    
    End If
    'REGISTRO E520: APURACAO DO IPI
    RegistroE520 op
    
    'REGISTRO E990: ENCERRAMENTO DO BLOCO E
    RegistroE990

End Sub

Private Sub blocoG(op As Boolean)
'###############################################################################################
'### BLOCO G: CONTROLE DO CREDITO DE ICMS DO ATIVO PERMANENTE - CIAO - MODELO C E D
'###############################################################################################
    
    'REGISTRO G001: ABERTURA DO BLOCO G
    RegistroG001
    
    'REGISTRO G990: ENCERRAMENTO DO BLOCO G
    RegistroG990
End Sub
Private Sub blocoK() 'op As Boolean)
'###############################################################################################
'### BLOCO K: CONTROLE DE PRODUÇÃO E DE ESTOQUE
'###############################################################################################
    
    'REGISTRO K001: ABERTURA DO BLOCO K
    RegistroK001
    
    'REGISTRO K990: ENCERRAMENTO DO BLOCO K
    RegistroK990
End Sub

Private Sub blocoH()
'###############################################################################################
'### BLOCO H: INVENTÁRIO FÍSICO
'###############################################################################################
    
    'infInvent = informar inventario
    
    'REGISTRO H001: ABERTURA DO BLOCO H
    RegistroH001 infInvent
    
    'REGISTRO H005: TOTAIS DO INVENTÁRIO
    RegistroH005 infInvent, "01"   'dtFinal, motInventario
    
    'REGISTRO H010: INVENTÁRIO
    RegistroH010 infInvent
    
    'REGISTRO H990: ENCERRAMENTO DO BLOCO H
    RegistroH990
    
End Sub


'
Public Sub MnFiscal_EFD(codFinalidade As Integer, DtIni As String, DtFin As String, informarInventario As Boolean)
    
    Dim sSQL            As String
    Dim Rst             As Recordset
    
    If Dir(App.Path & "\efd", vbDirectory) = Empty Then
        MkDir App.Path & "\efd"
    End If
    

    nmFile = App.Path & "\efd\" & App.EXEName & "_EFD" & _
             "_" & Replace(DtIni, "/", "-") & _
             "_" & Replace(DtFin, "/", "-") & _
             "_" & Replace(Date, "/", "") & "" & Replace(Time, ":", "") & _
             ".txt"
    
    'Apaga arquivos anteriores
    ExcluirFile nmFile
    cArray = 0
    dtInicial = DtIni
    dtFinal = DtFin
    'codFinalidade - 0 - Remessa de arquivo original
    '                1 - Remessa de arquivo substituto
     infInvent = informarInventario
     
   

'###############################################################################################
'### BLOCO 0 - ABERTURA, IDENTIFICACAO E REFERENCIAS
'###############################################################################################
    bloco0 True
'###############################################################################################
'### BLOCO C: DOCUMENTOS FISCAIS i - ICMS / IPI
'###############################################################################################
    blocoC True
'###############################################################################################
'### BLOCO D: DOCUMENTOS FISCAIS II - SERVICOS ICMS
'###############################################################################################
    blocoD True
'###############################################################################################
'### BLOCO E: APURACAO DO ICMS E DO IPI
'###############################################################################################
    blocoE True
'###############################################################################################
'### BLOCO G: CONTROLE DO CREDITO DE ICMS DO ATIVO PERMANENTE - CIAO - MODELO C E D
'###############################################################################################
    blocoG True
'###############################################################################################
'### BLOCO H: INVENTÁRIO FÍSICO
'###############################################################################################
    blocoH
'###############################################################################################
'### BLOCO K: CONTROLE DE PRODUÇÃO E DE ESTOQUE
'###############################################################################################
    blocoK
'###############################################################################################
'### BLOCO 1: OUTRAS INFORMACOES
'###############################################################################################
    bloco1
'###############################################################################################
'### BLOCO 9
'###############################################################################################
    bloco9
    
    
    MsgBox "Arquivo " & nmFile & " gerado com sucesso!", vbInformation, App.EXEName
    
End Sub
Private Function QtdLinRegistro(tpReg As String) As Integer
    Dim i As Integer
    Dim qtd_lin As Integer
    '*
    '* Nasce com 1 pois trata-se do proprio
    '* bloco que sera incrementado no final
    '* da funcao
    '*
    qtd_lin = 1
    For i = 1 To cArray
        If UCase(Left(vReg(i)(0), 1)) = tpReg Then
            qtd_lin = qtd_lin + vReg(i)(1)
        End If
    Next
    QtdLinRegistro = qtd_lin
End Function
Public Function cEFD(sDados As Variant, sQtd As Integer, cDecimal As Integer, sTp As String) As String
    '####################################################################################
    '### Converte os dados para criacao dos arquivos  confomr ATO COTEPE 9 18/04/2008 ###
    '####################################################################################
    '# Alterado em 03.12.2012 conforme Guia Pratico EFD v.2.0.10
    'On Error GoTo TrtErro
    Dim sDado       As String
    
    sDado = cNull(sDados)
    
    Select Case UCase(sTp)
        
        Case "N" '### Campo numerico / GPEFD-v.2.0.10
            If cDecimal = 0 Then
                sDado = Replace(sDado, ".", "")
                sDado = Replace(sDado, ",", "")
                Else
                    sDado = ChkVal(sDado, 0, cDecimal)
            End If
            If sQtd <> 0 Then
                    cEFD = Left(String(sQtd, "0"), sQtd - Len(Trim(sDado))) & Trim(sDado)
                Else
                    If Len(Trim(sDado)) > sQtd Then
                        cEFD = sDado
                        'Exit Function
                    End If
                
                    
            End If
            
            cEFD = Replace(cEFD, ".", ",")
        Case "C" '### Campo Caracteres / GPEFD-v.2.0.10
            'sDado = Replace(sDado, "|", "")
            'If Len(Trim(sDado)) > sQtd Then
            '        cEFD = sDado
            '        Exit Function
            '    Else
            '        If Trim(sDado) = "" Then
            '                cEFD = ""
            '            Else
            '                cEFD = Trim(sDado) & Mid(String(sQtd, " "), 1, sQtd - Len(Trim(sDado)))
            '        End If
            '        Exit Function
            'End If
            
            cEFD = Trim(sDado)
            If Len(cEFD) > sQtd And sQtd <> 0 Then
                MsgBox "Conjunto de dados maior que o permitido (" & sQtd & "): " & vbCrLf & cEFD, vbInformation, App.EXEName
                sDado = Trim(sDado)
                cEFD = Trim(Mid(sDado, 1, sQtd))
            End If
            
        Case "D" '### Converte a data em DDMMAAAA
            If sQtd = 6 Then
                    cEFD = Format(sDado, "DDMMYY")
                ElseIf sQtd = 8 Then
                    cEFD = Format(sDado, "DDMMYYYY")
                Else
            End If
        'Case "H" '### Converte a data em HHMMSS
        '    cEFD = Format(sDado, "HHMMSS")
        Case Else
            MsgBox "cEFD: Formato de dado não encontrado!", vbInformation, App.EXEName
    End Select

    Exit Function
TrtErro:
    'cEFD = Trim(Num)
    MsgBox "Erro conv"
    cEFD = ""
End Function

Private Function Registro0000() As Integer  '(dtInicial As String, dtFinal As String) As Integer
    
    sR = "0000"
    'Versao do Layout
    'Codigo | Versao    | obrigatoriedade
    '010    |   109     |   01/01/2016
    '011    |   110     |   01/01/2017
    sR = sR & "|" & "010" 'Guia pratico EFD-ICMS/IPI v.2.0.15 de 12/12/2014
    sR = sR & "|" & "0" 'codFinalidade
    sR = sR & "|" & cEFD(dtInicial, 8, 0, "D")
    sR = sR & "|" & cEFD(dtFinal, 8, 0, "D")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).Nome, 100, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).CNPJ, 14, 0, "C")
    sR = sR & "|" & cEFD(" ", 11, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).UF, 2, 0, "C")
    
    sR = sR & "|" & cEFD(RS(PgDadosEmpresa(ID_Empresa).IE), 14, 0, "C")
    
    sR = sR & "|" & cEFD(PgDadosMunicipio(PgDadosEmpresa(ID_Empresa).UF, PgDadosEmpresa(ID_Empresa).Mun).codMun, 7, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).im, 14, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).Suframa, 9, 0, "C")
    sR = sR & "|" & "A"
    sR = sR & "|" & PgDadosEmpresa(ID_Empresa).TipoAtividade
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("0000", 1)
    grvFile nmFile, sR
End Function
Private Function Registro0001() As Integer
    'ABERTURA DO BLOCO 0
    Dim sR As String
    sR = "0001"
    sR = sR & "|" & "0" '0 = com dados / 1 = sem dados
    sR = "|" & sR & "|"
    
   cArray = cArray + 1: vReg(cArray) = Array("0001", 1)
    grvFile nmFile, sR
End Function
Private Function Registro0005() As Integer
    Dim sR As String
    sR = "0005" 'DADOS COMPLEMENTARES
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).Nome, 60, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).CEP, 8, 0, "N")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).Lgr, 60, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).Nro, 10, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).Cpl, 60, 0, "N")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).Bairro, 60, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).Fone, 11, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).Fone, 11, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).Mail, 0, 0, "C")
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("0005", 1)

    grvFile nmFile, sR
End Function
Private Function Registro0100() As Integer
    Dim sR As String
    sR = "0100" 'DADOS DO CONTABILISTA
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).crNome, 100, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).crCPF, 11, 0, "N")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).crCRC, 15, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).cCNPJ, 14, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).cCEP, 8, 0, "N")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).cEndereco, 60, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).cNumero, 10, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).cCompl, 60, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).cBairro, 60, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).cFone1, 11, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).cFone2, 11, 0, "C")
    sR = sR & "|" & cEFD(PgDadosEmpresa(ID_Empresa).cMail, 0, 0, "C")
    sR = sR & "|" & cEFD(PgDadosMunicipio(PgDadosEmpresa(ID_Empresa).cUF, PgDadosEmpresa(ID_Empresa).cMunicipio).codMun, 7, 0, "C")
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("0100", 1)
    grvFile nmFile, sR
End Function
Private Function Registro0190() As Integer
    Dim sSQL As String
    Dim Rst As Recordset
    Dim i As Integer

    sSQL = "SELECT DISTINCT UM.sigla, UM.Descricao, EP.ID_Empresa, EP.Deposito, EP.Status, EP.IncluirBalanco, EP.Unidade " & _
         "FROM EstoqueProduto as EP, EstoqueUnidadeMedida as UM " & _
         "WHERE EP.Unidade = UM.sigla " & _
         "AND EP.ID_Empresa = " & ID_Empresa & " AND EP.Deposito = " & ID_Deposito & " AND EP.Status = 'ATIVO'" & _
         " AND EP.IncluirBalanco = 1 "
         
         
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro - 0190: Erro ao ler Unidades de Medida!", vbInformation, App.EXEName
            Rst.Close
        Else
            Rst.MoveFirst
    End If
    i = 0
    Do Until Rst.EOF
        i = i + 1
        sR = "0190"
        sR = sR & "|" & cEFD(Rst.Fields("sigla"), 6, 0, "C")
        sR = sR & "|" & cEFD(Rst.Fields("Descricao"), 0, 0, "C")
        sR = "|" & sR & "|"
        grvFile nmFile, sR
        Rst.MoveNext
    Loop
    
    
    '##### GAMBIARRA ###############################################
'        i = i + 1
'        sR = "0190"
'        sR = sR & "|" & cEFD("M", 6, 0, "C")
'        sR = sR & "|" & cEFD("METRO", 0, 0, "C")
'        sR = "|" & sR & "|"
'        grvFile nmFile, sR
        
        'i = i + 1
        'sR = "0190"
        'sR = sR & "|" & cEFD("TON", 6, 0, "C")
        'sR = sR & "|" & cEFD("TONELADA", 0, 0, "C")
        'sR = "|" & sR & "|"
        'grvFile nmFile, sR

        
'        i = i + 1
'        sR = "0190"
'        sR = sR & "|" & cEFD("RLS", 6, 0, "C")
'        sR = sR & "|" & cEFD("ROLOS", 0, 0, "C")
'        sR = "|" & sR & "|"
'        grvFile nmFile, sR
        
'        i = i + 1
'        sR = "0190"
'        sR = sR & "|" & cEFD("BOB", 6, 0, "C")
'        sR = sR & "|" & cEFD("BOBINA", 0, 0, "C")
'        sR = "|" & sR & "|"
'        grvFile nmFile, sR
        
'        i = i + 1
'        sR = "0190"
'        sR = sR & "|" & cEFD("PCS", 6, 0, "C")
'        sR = sR & "|" & cEFD("PECAS", 0, 0, "C")
'        sR = "|" & sR & "|"
'        grvFile nmFile, sR
        
'        i = i + 1
'        sR = "0190"
'        sR = sR & "|" & cEFD("METRO", 6, 0, "C")
'        sR = sR & "|" & cEFD("METROS", 0, 0, "C")
'        sR = "|" & sR & "|"
'        grvFile nmFile, sR
        
'        i = i + 1
'        sR = "0190"
'        sR = sR & "|" & cEFD("UN", 6, 0, "C")
'        sR = sR & "|" & cEFD("UNITARIO", 0, 0, "C")
'        sR = "|" & sR & "|"
'        grvFile nmFile, sR

'        i = i + 1
'        sR = "0190"
'        sR = sR & "|" & cEFD("VR", 6, 0, "C")
'        sR = sR & "|" & cEFD("VARA", 0, 0, "C")
'        sR = "|" & sR & "|"
'        grvFile nmFile, sR

    '##################################################################
    
    
    cArray = cArray + 1: vReg(cArray) = Array("0190", i)

    
End Function
Private Function Registro0200(infInvent As Boolean) As Integer
', Optional dtIni As Date, Optional dtFin As Date) As Integer
    'REGISTRO 0200: TABELA DE IDENTIFICAÇÃO DO ITEM (PRODUTO E SERVICOS)
    
    Dim i       As Integer
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim Aliq    As String
    If infInvent = True Then
            sSQL = "SELECT " & _
            "estoqueproduto.id_Empresa, estoqueproduto.Deposito, estoqueproduto.Status, " & _
            "estoqueproduto.Descricao AS Descr, " & _
            "estoqueproduto.Unidade AS un, " & _
            "estoqueproduto.NCM AS ncm, " & _
            "estoqueproduto.CodigoBarras AS codBar, " & _
            "estoqueproduto.Id AS idProd " & _
            "FROM " & _
            "EstoqueProduto " & _
            "WHERE " & _
            "ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND Status = 'ATIVO' " & _
            "AND IncluirBalanco = 1 " & _
            "ORDER BY Descricao"
            
        Else
            'sSQL = "SELECT " & _
            "faturamentonfe.ide_dEmi AS dEmi, estoqueproduto.Descricao AS Descr , " & _
            "estoqueproduto.Unidade AS un, " & _
            "estoqueproduto.NCM AS ncm, " & _
            "estoqueproduto.CodigoBarras AS codBar, " & _
            "faturamentonfeitens.det_IdProduto AS idProd, " & _
            "faturamentonfe.Id_Empresa " & _
            "FROM " & _
            "estoqueproduto INNER JOIN " & _
            "faturamentonfeitens ON estoqueproduto.ID = " & _
            "faturamentonfeitens.det_IdProduto INNER JOIN " & _
            "faturamentonfe ON faturamentonfe.IdNFe = faturamentonfeitens.IdNFe " & _
            "WHERE " & _
            "faturamentonfe.ide_dEmi BETWEEN '" & Format(dtIni, "YYYY-MM-DD") & "' AND '" & Format(dtFin, "YYYY-MM-DD") & "' " & _
            "GROUP BY " & _
            "faturamentonfeitens.det_IdProduto"
            
            
            '#############################################################
            
           MontarTabelaTemporaria_Reg0200 dtInicial, dtFinal
           
            
            
            
            
    End If
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro200: Nenhum registro no periodo", vbInformation, App.EXEName
        Else
            Rst.MoveFirst
        End If
    i = 0
    Do Until Rst.EOF
        sR = "0200"
        sR = sR & "|" & cEFD(Rst.Fields("idProd"), 60, 0, "C")
        sR = sR & "|" & cEFD(cNull(Rst.Fields("Descr")), 0, 0, "C")
        sR = sR & "|" & cEFD(cNull(Rst.Fields("codBar")), 0, 0, "C")
        sR = sR & "|" & cEFD("", 60, 0, "C") 'COD_ANT_ITEM
        sR = sR & "|" & cEFD(Rst.Fields("Un"), 6, 0, "C")
        sR = sR & "|" & cEFD("00", 2, 0, "N") 'TIPO_ITEM
        sR = sR & "|" & cEFD(cNull(Rst.Fields("NCM")), 8, 0, "C")
        sR = sR & "|" & cEFD("", 3, 0, "C")
        sR = sR & "|" & cEFD(IIf(cNull(Rst.Fields("NCM")) <> "", Left(cNull(Rst.Fields("NCM")), 2), ""), 2, 0, "N") 'cEFD("96", 2, 0, "C")
        sR = sR & "|" & cEFD("", 5, 0, "C")
        Aliq = pgDadosICMS(PgDadosEmpresa(ID_Empresa).UF, 0).ICMS  'cNull(Rst.fields("Aliquota"))
        If Not IsNumeric(Aliq) Then
            Aliq = 0
        End If
        sR = sR & "|" & cEFD(Aliq, 6, 2, "N")
        
        'sR = sR & "|" & cEFD("0", 7, 0, "N") 'CEST
        
        sR = "|" & sR & "|"
    
        i = i + 1
        grvFile nmFile, sR
        Rst.MoveNext
    Loop
    Rst.Close
    cArray = cArray + 1: vReg(cArray) = Array("0200", i)
End Function
Private Function Registro0400(op As Boolean) As Integer
    'REGISTRO 0400: TABELA DE NATUREZA DA OPERACAO/PRESTACAO
    Dim cCFOP(10)       As Variant
    Dim c               As Integer
    Dim n               As Integer


    'Dim sSQL As String
    'ssql="SELECT * FROM
    If op = False Then Exit Function

    c = 0
    'cCFOP(c) = Array("5102", "Venda de mercadoria adquirida ou recebida de terceiros"): c = c + 1
    'cCFOP(c) = Array("6102", "Venda de mercadoria adquirida ou recebida de terceiros"): c = c + 1
    
    'cCFOP(c) = Array("1101", "Compra para industrialização ou produção rural"): c = c + 1
    cCFOP(c) = Array("2101", "Compra para industrialização ou produção rural"): c = c + 1
    cCFOP(c) = Array("1102", "Compra para comercialização"): c = c + 1
    cCFOP(c) = Array("2102", "Compra para comercialização"): c = c + 1
    
    'cCFOP(c) = Array("2202", "Devolucao de venda de mercadoria adquirida ou recebida de terceiros"): c = c + 1
    
    'cCFOP(c) = Array("1202", "Devolução de venda de mercadoria adquirida ou recebida de terceiros"): c = c + 1
    
    'cCFOP(c) = Array("5929", "Lançamento efetuado em decorrência de emissão de documento fiscal relativo a operação ou prestação também registrada em equipamento Emissor de Cupom Fiscal - ECF"): c = c + 1
    
    'cCFOP(c) = Array("1124", "Industrialização efetuada por outra empresa"): c = c + 1
    
    'cCFOP(c) = Array("1902", "Retorno de mercadoria remetida para industrialização por encomenda"): c = c + 1
    
    'cCFOP(c) = Array("5403", "Venda de mercadoria, adquirida ou recebida de terceiros, sujeita ao regime de substituição tributária, na condição de contribuinte-substituto"): c = c + 1
    'cCFOP(c) = Array("1403", "Compra para comercialização em operação com mercadoria sujeita ao regime de substituição tributária"): c = c + 1
    cCFOP(c) = Array("2403", "Compra para comercialização em operação com mercadoria sujeita ao regime de substituição tributária"): c = c + 1
    
    For n = 0 To c - 1
        sR = "0400"
        sR = sR & "|" & cEFD(cCFOP(n)(0), 10, 0, "C")
        sR = sR & "|" & cEFD(UCase(cCFOP(n)(1)), 0, 0, "C")
        sR = "|" & sR & "|"
        
        grvFile nmFile, sR
    Next
    cArray = cArray + 1: vReg(cArray) = Array("0400", c)
End Function
Private Function Registro0990() As Integer
    'REGISTRO 0990: ENCERRAMENTO DO BLOCO 0
    Dim qtd_lin As Integer
    
    qtd_lin = QtdLinRegistro("0")
        
    sR = "0990"
    sR = sR & "|" & qtd_lin
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("0990", "1") ' qtd_lin)
    grvFile nmFile, sR
End Function
Private Function Registro1010() As Integer
    sR = "1010"
    sR = sR & "|" & cEFD("N", 1, 0, "C")
    sR = sR & "|" & cEFD("N", 1, 0, "C")
    sR = sR & "|" & cEFD("N", 1, 0, "C")
    sR = sR & "|" & cEFD("N", 1, 0, "C")
    sR = sR & "|" & cEFD("N", 1, 0, "C")
    sR = sR & "|" & cEFD("N", 1, 0, "C")
    sR = sR & "|" & cEFD("N", 1, 0, "C")
    sR = sR & "|" & cEFD("N", 1, 0, "C")
    sR = sR & "|" & cEFD("N", 1, 0, "C")
    
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("1010", "1")
    grvFile nmFile, sR
End Function

Private Function RegistroC001(infDados As Boolean) As Integer
    'REGISTRO C001: ABERTURA DO BLOCO C
    sR = "C001"
    sR = sR & "|" & IIf(infDados = False, "1", "0")
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("C001", "1")
    grvFile nmFile, sR
End Function
Private Function RegistroC100() As Integer '(DtIni As Date, DtFin As Date) As Integer
    'REGISTRO C100: DOC FISCAL 55 - SAIDA
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim c           As Integer
    Dim cRegC170    As Integer
    Dim cRegC190    As Integer
    
    
    c = 0
    cRegC170 = 0
    '####################################
    '###  Registro de SAIDA (1)
    '####################################
    
    sSQL = "SELECT * FROM faturamentonfe " & _
            "WHERE ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
            "ORDER BY ide_dEmi, id"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'RegistroC100 = 0
            Rst.Close
            'Exit Function
        Else
            Rst.MoveFirst
            
            Do Until Rst.EOF
                c = c + 1
                sR = "C100"
                sR = sR & "|" & "1" 'ind_oper
                sR = sR & "|" & "0" 'ind_emit
                sR = sR & "|C" & cEFD(Rst.Fields("dest_idDest"), 60, 0, "C")
                sR = sR & "|" & "55" ' cod_mod tabela 4.1.1
                sR = sR & "|" & "00" ' cod_sit tabela 4.1.2
                sR = sR & "|" & cEFD(Rst.Fields("ide_serie"), 3, 0, "C")
                sR = sR & "|" & cEFD(Rst.Fields("ide_nnf"), 9, 0, "N")
                sR = sR & "|" & cEFD(Rst.Fields("idNFe"), 44, 0, "N")
                sR = sR & "|" & cEFD(Rst.Fields("ide_dEmi"), 8, 0, "D")
                sR = sR & "|" & cEFD(Rst.Fields("ide_dEmi"), 8, 0, "D") 'DT_E_S Entrada=O / Saida = OC
                sR = sR & "|" & cEFD(Rst.Fields("total_vNF"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("ide_indPag"), 1, 0, "C")
                sR = sR & "|" & cEFD(Rst.Fields("total_vDesc"), 0, 2, "N")
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_ABAT_NT
                sR = sR & "|" & cEFD(Rst.Fields("total_vProd"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("transp_modFrete"), 0, 1, "C")
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_FRT
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_SEG
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_OUT_DA
                sR = sR & "|" & cEFD(Rst.Fields("total_vBC"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vICMS"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vBCST"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vICMSST"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vIPI"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vPIS"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vCOFINS"), 0, 2, "N")
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_PIS_ST
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_COFINS_ST
                sR = "|" & sR & "|"
                
                grvFile nmFile, sR
                
                'cRegC170 = cRegC170 + RegistroC170(Rst.fields("idNFe"), "s")
                cRegC190 = cRegC190 + RegistroC190("s", Rst.Fields("idNFe"))
                
                Rst.MoveNext
            Loop
            Rst.Close
    End If
    
    '####################################
    '###  Registro de ENTRADA (0)
    '####################################
    sSQL = "SELECT * FROM faturamentonfeentrada " & _
            "WHERE ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
            "ORDER BY ide_dEmi, id"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'RegistroC100 = 0
            Rst.Close
            'Exit Function
        Else
            Rst.MoveFirst
            
            Do Until Rst.EOF
                c = c + 1
                                
                sR = "C100"
                sR = sR & "|" & "0" 'ind_oper
                sR = sR & "|" & "1" 'ind_emit
                sR = sR & "|F" & cEFD(Rst.Fields("emit_id"), 60, 0, "C")
                sR = sR & "|" & "55" ' cod_mod tabela 4.1.1
                sR = sR & "|" & "00" ' cod_sit tabela 4.1.2
                sR = sR & "|" & cEFD(Rst.Fields("ide_serie"), 3, 0, "C")
                sR = sR & "|" & cEFD(Rst.Fields("ide_nnf"), 9, 0, "N")
                sR = sR & "|" & cEFD(Rst.Fields("idNFe"), 44, 0, "N")
                sR = sR & "|" & cEFD(Rst.Fields("ide_dEmi"), 8, 0, "D")
                sR = sR & "|" & cEFD(Rst.Fields("ide_dEmi"), 8, 0, "D") 'DT_E_S Entrada=O / Saida = OC
                sR = sR & "|" & cEFD(Rst.Fields("total_vNF"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("ide_indPag"), 1, 0, "C")
                sR = sR & "|" & cEFD(Rst.Fields("total_vDesc"), 0, 2, "N")
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_ABAT_NT
                sR = sR & "|" & cEFD(Rst.Fields("total_vProd"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("transp_modFrete"), 0, 1, "C")
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_FRT
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_SEG
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_OUT_DA
                sR = sR & "|" & cEFD(Rst.Fields("total_vBC"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vICMS"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vBCST"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vICMSST"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vIPI"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vPIS"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("total_vCOFINS"), 0, 2, "N")
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_PIS_ST
                sR = sR & "|" & cEFD("0", 0, 2, "N") 'VL_COFINS_ST
                sR = "|" & sR & "|"
                
                grvFile nmFile, sR
                cRegC170 = cRegC170 + RegistroC170(Rst.Fields("idNFe"), "e")
                cRegC190 = cRegC190 + RegistroC190("e", Rst.Fields("idNFe"))
                Rst.MoveNext
            Loop
            Rst.Close
    End If
    
    
    cArray = cArray + 1: vReg(cArray) = Array("C100", c)
    cArray = cArray + 1: vReg(cArray) = Array("C170", cRegC170)
    cArray = cArray + 1: vReg(cArray) = Array("C190", cRegC190)
End Function
Private Function RegistroC170(chv As String, Mov As String) As Integer

  'REGISTRO C170: Itens do documento fiscal
  
  'mov - e: ENTRADA / s: SAIDA
  
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim c           As Integer
    Dim Tabela      As String
    Dim nItem       As Integer
    
    c = 0
    nItem = 0
    If LCase(Mov) = "e" Then
            'ENTRADA
            Tabela = "faturamentonfeentradaitens"
        Else
            'SAIDA
            Tabela = "faturamentonfeitens"
    End If
    
    sSQL = "SELECT * FROM " & Tabela & " " & _
            "WHERE idNFe='" & chv & "' " & _
            "ORDER BY id"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            
            'Rst.Close
            
        Else
            Rst.MoveFirst
            
            Do Until Rst.EOF
                c = c + 1
                nItem = nItem + 1
                sR = "C170"
                sR = sR & "|" & cEFD(CStr(nItem), 3, 0, "N")
                sR = sR & "|" & cEFD(Rst.Fields("det_idProduto"), 60, 0, "C")
                sR = sR & "|" & cEFD(Rst.Fields("det_xProd"), 0, 0, "C")
                sR = sR & "|" & cEFD(Rst.Fields("det_qCom"), 0, 5, "N")
                sR = sR & "|" & cEFD(Rst.Fields("det_uCom"), 6, 5, "C")
                sR = sR & "|" & cEFD(Rst.Fields("det_vProd"), 0, 2, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("det_vDesc")), 0, 2, "N")
                sR = sR & "|" & cEFD("0", 1, 0, "C") 'Idn_mov
                sR = sR & "|" & cEFD(cNull(Rst.Fields("ICMS_CST")), 3, 0, "N")
                
                If LCase(Mov) = "e" Then
                        'ENTRADA
                        sR = sR & "|" & cEFD(cCFOPEntrada(chv, Rst.Fields("det_CFOP")), 4, 0, "N") '11 - CFOP
                        sR = sR & "|" & cEFD(cCFOPEntrada(chv, Rst.Fields("det_CFOP")), 4, 0, "N") '12 - cod_NAT
                    Else
                        'SAIDA
                        sR = sR & "|" & cEFD(Rst.Fields("det_CFOP"), 4, 0, "N") '11 - CFOP
                        sR = sR & "|" & cEFD(Rst.Fields("det_CFOP"), 10, 0, "C") '12 - cod_NAT
                End If
                
                
                sR = sR & "|" & cEFD(cNull(Rst.Fields("ICMS_vBC")), 0, 2, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("ICMS_pICMS")), 6, 2, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("ICMS_vICMS")), 0, 2, "N")
                
                sR = sR & "|" & cEFD(cNull(Rst.Fields("ICMS_vBCST")), 0, 2, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("ICMS_pICMSST")), 6, 2, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("ICMS_vICMSST")), 0, 2, "N")
                
                sR = sR & "|" & cEFD("0", 1, 0, "C")
                
                 If LCase(Mov) = "e" Then
                        'ENTRADA
                        sR = sR & "|" & cEFD(cCstIpiEntrada(cNull(Rst.Fields("IPI_CST"))), 2, 0, "C")
                    Else
                        'SAIDA
                        sR = sR & "|" & cEFD(cNull(Rst.Fields("IPI_CST")), 2, 0, "C")
                End If
                
                
                
                sR = sR & "|" & cEFD("", 3, 0, "C")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("IPI_vBC")), 0, 2, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("IPI_pIPI")), 6, 2, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("IPI_vIPI")), 0, 2, "N")
                
                sR = sR & "|" & cEFD(Rst.Fields("PIS_CST"), 2, 0, "N")
                sR = sR & "|" & cEFD(Rst.Fields("PIS_vBC"), 0, 2, "N")
                sR = sR & "|" & cEFD(Rst.Fields("PIS_pPIS"), 8, 4, "N")
                sR = sR & "|" & cEFD("", 0, 0, "N") '28 - QUANT_BC_PIS
                sR = sR & "|" & cEFD("", 8, 4, "N") 'cEFD(Rst.fields("PIS_pPIS"), 8, 4, "N")
                sR = sR & "|" & cEFD("", 0, 2, "N") 'cEFD(Rst.fields("PIS_vPIS"), 0, 2, "N")
                
                sR = sR & "|" & cEFD(Rst.Fields("COFINS_CST"), 2, 0, "N")
                sR = sR & "|" & cEFD(Rst.Fields("COFINS_vBC"), 0, 2, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("COFINS_pCOFINS")), 8, 4, "N")
                sR = sR & "|" & cEFD("", 0, 0, "N")
                sR = sR & "|" & cEFD(Rst.Fields("COFINS_pCOFINS"), 8, 4, "N")
                sR = sR & "|" & cEFD(Rst.Fields("COFINS_vCOFINS"), 0, 2, "N")
                
                sR = sR & "|" & cEFD("", 0, 0, "C") '37 - COD_CTA
                
                sR = "|" & sR & "|"
                
                grvFile nmFile, sR
                
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    
    'If vReg(cArray)(0) = "C170" Then
    '        vReg(cArray) = vReg(cArray)(1) + c
    '
    '    Else
    '        cArray = cArray + 1: vReg(cArray) = Array("C170", c)
    'End If
    RegistroC170 = c

End Function
Private Function Registro0150(op As Boolean) As Integer
    'REGISTRO 0150: TABELA DE CADASTRO DE PARTICIPANTES
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim c           As Integer
    
    c = 0
    
    If op = False Then Exit Function
    
    '*** Registro de SAIDA ***
    
    sSQL = "SELECT * FROM faturamentonfe " & _
            "WHERE ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
            "GROUP BY dest_idDest " & _
            "ORDER BY ide_dEmi, id "
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Registro0150 = 0
            
            'Exit Function
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                c = c + 1
                sR = "0150"
                sR = sR & "|C" & cEFD(Rst.Fields("dest_idDest"), 60, 0, "C")
                sR = sR & "|" & cEFD(Rst.Fields("dest_xNome"), 100, 0, "C")
                sR = sR & "|" & cEFD(Rst.Fields("dest_cPais"), 5, 0, "N")
                
                If LCase(Rst.Fields("dest_pessoa")) = "juridica" Then
                        sR = sR & "|" & cEFD(Rst.Fields("dest_CNPJ"), 14, 0, "N")
                        sR = sR & "|" & "" 'cEFD(IIf(LCase(Rst.fields("dest_pessoa")) = "fisica", Rst.fields("dest_CNPJ"), ""), 11, 0, "N")
                    Else
                        sR = sR & "|" & "" 'cEFD(IIf(LCase(Rst.fields("dest_pessoa")) = "juridica", Rst.fields("dest_CNPJ"), ""), 14, 0, "N")
                        sR = sR & "|" & cEFD(Rst.Fields("dest_CNPJ"), 11, 0, "N")
                End If
                sR = sR & "|" & cEFD(IIf(Trim(Rst.Fields("dest_IE")) = "ISENTO", "", Rst.Fields("dest_IE")), 14, 0, "C")
                sR = sR & "|" & cEFD(Rst.Fields("dest_cMun"), 7, 0, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("dest_ISUF")), 9, 0, "C")
                sR = sR & "|" & cEFD(Rst.Fields("dest_xLgr"), 60, 0, "C")
                sR = sR & "|" & cEFD(Rst.Fields("dest_nro"), 10, 0, "C")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("dest_xCpl")), 60, 0, "C")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("dest_Bairro")), 60, 0, "C")
                sR = "|" & sR & "|"
                grvFile nmFile, sR
                Rst.MoveNext
            Loop
    End If
   
    Rst.Close
    
    '*** Registro de ENTRADA ***
    
    sSQL = "SELECT * FROM faturamentonfeentrada " & _
            "WHERE ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
            "GROUP BY emit_id " & _
            "ORDER BY ide_dEmi, id"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        'Registro0150 = 0
        Rst.Close
        Exit Function
    End If
    Rst.MoveFirst
    
    Do Until Rst.EOF
        c = c + 1
        sR = "0150"
        sR = sR & "|F" & cEFD(Rst.Fields("emit_id"), 60, 0, "C")
        sR = sR & "|" & cEFD(Rst.Fields("emit_xNome"), 100, 0, "C")
        sR = sR & "|" & cEFD(IIf(cNull(Rst.Fields("emit_cPais")) = "", "1058", cNull(Rst.Fields("emit_cPais"))), 5, 0, "N")
        sR = sR & "|" & cEFD(Rst.Fields("emit_CNPJ"), 14, 0, "N")
        sR = sR & "|" '& cEFD(IIf(LCase(Rst.fields("emit_pessoa")) = "fisica", Rst.fields("emit_CNPJ"), ""), 11, 0, "N")
        sR = sR & "|" & cEFD(Rst.Fields("emit_IE"), 14, 0, "C")
        sR = sR & "|" & cEFD(Rst.Fields("emit_cMun"), 7, 0, "N")
        sR = sR & "|" '& cEFD(Rst.fields("emit_ISUF"), 9, 0, "C")
        sR = sR & "|" & cEFD(Rst.Fields("emit_xLgr"), 60, 0, "C")
        sR = sR & "|" & cEFD(Rst.Fields("emit_nro"), 10, 0, "C")
        sR = sR & "|" & cEFD(cNull(Rst.Fields("emit_xCpl")), 60, 0, "C")
        sR = sR & "|" & cEFD(Rst.Fields("emit_Bairro"), 60, 0, "C")
        sR = "|" & sR & "|"
        grvFile nmFile, sR
        Rst.MoveNext
    Loop
    Registro0150 = 1
    cArray = cArray + 1: vReg(cArray) = Array("0150", c)
    
End Function
Private Function RegistroC190(sMov As String, chv As String) As Integer
    'REGISTRO C190: REGISTRO ANALITICO DE DOCUMENTO
    'totalizacao de CST, CFOP, %ICMS

    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim c           As Integer
    
    'sSQL = "SELECT " & _
         "faturamentonfeitens.ICMS_vBCSTDest AS vBCSTICMS, " & _
         "faturamentonfeitens.det_CFOP AS cfop, " & _
         "faturamentonfeitens.IPI_vIPI AS IPI_vIPI, " & _
         "faturamentonfeitens.ICMS_vBC AS ICMS_vBC, " & _
         "faturamentonfeitens.ICMS_CST AS cstICMS, " & _
         "faturamentonfeitens.ICMS_pICMS AS pICMS, " & _
         "Sum(faturamentonfeitens.IPI_vIPI + faturamentonfeitens.ICMS_vBC) AS vlTotProd, " & _
         "Sum(faturamentonfeitens.ICMS_vBCSTRet) AS vBCSTRetICMS, " & _
         "Sum(faturamentonfeitens.ICMS_ICMSST) AS icmsStICMS, " & _
         "Sum(faturamentonfeitens.ICMS_pRedBC) AS pRedBCICMS, " & _
         "Sum(faturamentonfeitens.ICMS_vICMS) AS vICMS " & _
         "FROM " & _
         "faturamentonfe INNER JOIN " & _
         "faturamentonfeitens ON faturamentonfe.IdNFe = " & _
         "faturamentonfeitens.IdNFe " & _
         "WHERE faturamentonfe.IdNFe= '" & chv & "' " & _
         "GROUP BY " & _
         "faturamentonfeitens.det_CFOP, faturamentonfeitens.ICMS_CST, " & _
         "faturamentonfeitens.ICMS_pICMS"
         
         
         If chv = "33140417469701016170550000000017881929414295" Then
            MsgBox "ops"
        End If
         
         
         
         
         
         
         
         
         
         
         
    If LCase(sMov) = "s" Then
            sSQL = "SELECT " & _
                 "faturamentonfeitens.ICMS_vBCSTDest AS vBCSTICMS, " & _
                 "faturamentonfeitens.det_CFOP AS cfop, " & _
                 "SUM(faturamentonfeitens.IPI_vIPI) AS IPI_vIPI, " & _
                 "SUM(faturamentonfeitens.ICMS_vBC) AS ICMS_vBC, " & _
                 "faturamentonfeitens.ICMS_CST AS cstICMS, " & _
                 "faturamentonfeitens.ICMS_pICMS AS pICMS, " & _
                 "Sum(faturamentonfeitens.det_vProd) AS vlTotProd, " & _
                 "Sum(faturamentonfeitens.ICMS_vBCSTRet) AS vBCSTRetICMS, " & _
                 "Sum(faturamentonfeitens.ICMS_ICMSST) AS icmsStICMS, " & _
                 "Sum(faturamentonfeitens.ICMS_pRedBC) AS pRedBCICMS, " & _
                 "Sum(faturamentonfeitens.ICMS_vICMS) AS vICMS " & _
                 "FROM " & _
                 "faturamentonfeitens " & _
                 "WHERE IdNFe= '" & chv & "' " & _
                 "GROUP BY " & _
                 "faturamentonfeitens.det_CFOP, faturamentonfeitens.ICMS_CST, " & _
                 "faturamentonfeitens.ICMS_pICMS"
        Else
        'Alterado devido erro na descr do st do material - 27/06/2014
            'sSQL = "SELECT " & _
                 "faturamentonfeentradaitens.ICMS_vBCSTDest AS vBCSTICMS, " & _
                 "faturamentonfeentradaitens.det_CFOP AS cfop, " & _
                 "SUM(faturamentonfeentradaitens.IPI_vIPI) AS IPI_vIPI, " & _
                 "SUM(faturamentonfeentradaitens.ICMS_vBC) AS ICMS_vBC, " & _
                 "faturamentonfeentradaitens.ICMS_CST AS cstICMS, " & _
                 "faturamentonfeentradaitens.ICMS_pICMS AS pICMS, " & _
                 "SUM(faturamentonfeentradaitens.det_vProd) AS vlTotProd, " & _
                 "Sum(faturamentonfeentradaitens.ICMS_vBCSTRet) AS vBCSTRetICMS, " & _
                 "Sum(faturamentonfeentradaitens.ICMS_ICMSST) AS icmsStICMS, " & _
                 "Sum(faturamentonfeentradaitens.ICMS_pRedBC) AS pRedBCICMS, " & _
                 "Sum(faturamentonfeentradaitens.ICMS_vICMS) AS vICMS " & _
                 "FROM " & _
                 "faturamentonfeentradaitens " & _
                 "WHERE IdNFe= '" & chv & "' " & _
                 "GROUP BY " & _
                 "faturamentonfeentradaitens.det_CFOP, faturamentonfeentradaitens.ICMS_CST, " & _
                 "faturamentonfeentradaitens.ICMS_pICMS"
            sSQL = "SELECT " & _
                 "faturamentonfeentradaitens.ICMS_vBCST AS vBCSTICMS, " & _
                 "faturamentonfeentradaitens.det_CFOP AS cfop, " & _
                 "SUM(faturamentonfeentradaitens.IPI_vIPI) AS IPI_vIPI, " & _
                 "SUM(faturamentonfeentradaitens.ICMS_vBC) AS ICMS_vBC, " & _
                 "faturamentonfeentradaitens.ICMS_CST AS cstICMS, " & _
                 "faturamentonfeentradaitens.ICMS_pICMS AS pICMS, " & _
                 "SUM(faturamentonfeentradaitens.det_vProd) AS vlTotProd, " & _
                 "Sum(faturamentonfeentradaitens.ICMS_vBCSTRet) AS vBCSTRetICMS, " & _
                 "Sum(faturamentonfeentradaitens.ICMS_vICMSST) AS icmsStICMS, " & _
                 "Sum(faturamentonfeentradaitens.ICMS_pRedBC) AS pRedBCICMS, " & _
                 "Sum(faturamentonfeentradaitens.ICMS_vICMS) AS vICMS " & _
                 "FROM " & _
                 "faturamentonfeentradaitens " & _
                 "WHERE IdNFe= '" & chv & "' " & _
                 "GROUP BY " & _
                 "faturamentonfeentradaitens.det_CFOP, faturamentonfeentradaitens.ICMS_CST, " & _
                 "faturamentonfeentradaitens.ICMS_pICMS"
                
                
    End If
    Dim vlTotProd   As String
    Dim vl_red_BC   As String ' 10
    Dim cstICMS     As String
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Rst.Close
        Else
            Rst.MoveFirst
            c = 0
            Do Until Rst.EOF
                c = c + 1
                sR = "C190" '1
                cstICMS = cNull(Rst.Fields("cstICMS"))
                sR = sR & "|" & cEFD(cstICMS, 3, 0, "N") '2
                If LCase(sMov) = "e" Then
                        'ENTRADA
                        sR = sR & "|" & cEFD(cCFOPEntrada(chv, Rst.Fields("CFOP")), 4, 0, "N") '3
                    Else
                        'SAIDA
                        sR = sR & "|" & cEFD(Rst.Fields("CFOP"), 4, 0, "N") '3
                End If
                
                vlTotProd = Val(ChkVal(cNull(Rst.Fields("vlTotProd")), 0, cDecMoeda)) + Val(ChkVal(cNull(Rst.Fields("IPI_vIPI")), 0, cDecMoeda))
                
                If cstICMS = "20" Or cstICMS = "70" Then
                        vl_red_BC = cNull(Rst.Fields("ICMS_vBC"))
                    Else
                        vl_red_BC = Val(ChkVal(vlTotProd, 0, cDecMoeda)) - Val(ChkVal(cNull(Rst.Fields("ICMS_vBC")), 0, cDecMoeda))
                End If
                '############################
                sR = sR & "|" & cEFD(cNull(Rst.Fields("pICMS")), 6, 2, "N") '4
                sR = sR & "|" & cEFD(vlTotProd, 0, 2, "N") '5
                sR = sR & "|" & cEFD(cNull(Rst.Fields("ICMS_vBC")), 0, 2, "N") '6
                sR = sR & "|" & cEFD(cNull(Rst.Fields("vICMS")), 0, 2, "N") '7
                
                'Modificado devido o erro acima 27/06/2014
                'sR = sR & "|" & cEFD(cNull(Rst.fields("vBCSTRetICMS")), 0, 2, "N") '8
                'sR = sR & "|" & cEFD(cNull(Rst.fields("icmsStICMS")), 0, 2, "N") '9
                
                sR = sR & "|" & cEFD(cNull(Rst.Fields("vBCSTICMS")), 0, 2, "N") '8
                sR = sR & "|" & cEFD(cNull(Rst.Fields("icmsStICMS")), 0, 2, "N") '9
                
                sR = sR & "|" & cEFD(vl_red_BC, 0, 2, "N") '10
                
                sR = sR & "|" & cEFD(cNull(Rst.Fields("IPI_vIPI")), 0, 2, "N") '11
                sR = sR & "|" & cEFD("", 6, 0, "C") '12
                
                sR = "|" & sR & "|"
                grvFile nmFile, sR
                
                Rst.MoveNext
            Loop
    End If
    
    'cArray = cArray + 1: vReg(cArray) = Array("C190", c)
    RegistroC190 = c
    
End Function

Private Function RegistroC990() As Integer
    'REGISTRO C990: ENCERRAMENTO DO BLOCO C
    Dim qtd_lin As Integer
    
    qtd_lin = QtdLinRegistro("C")
    
    sR = "C990"
    sR = sR & "|" & qtd_lin
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("C990", "1") 'Quantidade de reg C9900
    grvFile nmFile, sR
    
End Function
Private Function RegistroD001() As Integer
'REGISTRO D001: ABERTURA DO BLOCO D
    sR = "D001"
    sR = sR & "|" & "1"
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("D001", "1")
    grvFile nmFile, sR
End Function
Public Function RegistroD990() As Integer
'REGISTRO D990: ENCERRAMENTO DO BLOCO D
    Dim qtd_lin As Integer
    qtd_lin = QtdLinRegistro("D")
    sR = "D990"
    sR = sR & "|" & qtd_lin
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("D990", qtd_lin)
    grvFile nmFile, sR
    
End Function

Private Function RegistroE001() As Integer
'REGISTRO E001: ABERTURA DO BLOCO E
    sR = "E001"
    sR = sR & "|" & "0"
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("E001", "1")
    grvFile nmFile, sR
End Function
Private Function RegistroE100() As Integer '(dtI As String, dtF As String) As Integer
    'REGISTRO E100: PERIODO DA APURACAO DO ICMS
    sR = "E100"
    sR = sR & "|" & cEFD(dtInicial, 8, 0, "D")
    sR = sR & "|" & cEFD(dtFinal, 8, 0, "D")
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("E100", "1")
    grvFile nmFile, sR
End Function
Private Function RegistroE110(op As Boolean) As Integer '(dtIni As Date, dtFin As Date) As Integer
    'REGISTRO E110: APURACAO DO ICMS OPERACOES PROPRIAS
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    Dim vIcmsDeb    As String
    Dim vIcmsCre    As String
    Dim vIcmsSal    As String
    
    Dim vIpiCre     As String
    
    If op = True Then
            '### PEGA VALOR TOTAL DO DEBITO ICMS
            sSQL = "SELECT SUM(total_vICMS) AS vICMS " & _
                    "FROM faturamentonfe " & _
                    "WHERE ide_tpNF=1 AND ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
                    "ORDER BY ide_dEmi, id"
            
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    vIcmsDeb = 0
                Else
                    Rst.MoveFirst
                    vIcmsDeb = Rst.Fields("vICMS")
            End If
            Rst.Close
            
            '### PEGA VALOR TOTAL DO CREDITO ICMS
            sSQL = "SELECT SUM(total_vICMS) AS vICMS " & _
                    "FROM faturamentonfeentrada " & _
                    "WHERE ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
                    "ORDER BY ide_dEmi, id"
            
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    vIcmsCre = 0
                Else
                    Rst.MoveFirst
                    vIcmsCre = Rst.Fields("vICMS")
            End If
            Rst.Close
            
            'Verifica as notas fiscais emitida com cod de DEVOLUCAO
            Dim vIcmsCreDEV As String
            sSQL = "SELECT SUM(total_vICMS) AS vICMS " & _
                    "FROM faturamentonfe " & _
                    "WHERE ide_tpNF=0 AND ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
                    "ORDER BY ide_dEmi, id"
            
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    'vIcmsDeb = 0
                Else
                    Rst.MoveFirst
                    vIcmsCreDEV = cNull(Rst.Fields("vICMS"))
            End If
            Rst.Close
            
            vIcmsCre = Val(ChkVal(vIcmsCre, 0, cDecMoeda)) + Val(ChkVal(vIcmsCreDEV, 0, cDecMoeda))
            
            '### PEGA VALOR TOTAL DO CREDITO IPI
            sSQL = "SELECT SUM(total_vIPI) AS vIPI " & _
                    "FROM faturamentonfeentrada " & _
                    "WHERE ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
                    "ORDER BY ide_dEmi, id"
            
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    vIpiCre = 0
                Else
                    Rst.MoveFirst
                    vIpiCre = Rst.Fields("vIPI")
            End If
            Rst.Close
            
        Else
        
            vIcmsDeb = 0
            vIcmsCre = 0
            vIpiCre = 0
        
    End If
            
    
    '##########################################
    vIcmsSal = Val(ChkVal(vIcmsDeb, 0, 2)) - Val(ChkVal(vIcmsCre, 0, 2))
    vIcmsSal = ChkVal(vIcmsSal, 0, 2)
    '##########################################
    
    sR = "E110" '1
    sR = sR & "|" & cEFD(vIcmsDeb, 0, 2, "N") '2
    sR = sR & "|" & cEFD("0", 0, 2, "N") '3
    sR = sR & "|" & cEFD("0", 0, 2, "N") '4
    sR = sR & "|" & cEFD("0", 0, 2, "N") '5
    sR = sR & "|" & cEFD(vIcmsCre, 0, 2, "N") '6
    sR = sR & "|" & cEFD("0", 0, 2, "N") '7
    sR = sR & "|" & cEFD("0", 0, 2, "N") '8
    sR = sR & "|" & cEFD("0", 0, 2, "N") '9
    sR = sR & "|" & cEFD("0", 0, 2, "N") '10
    sR = sR & "|" & cEFD(vIcmsSal, 0, 2, "N") '11
    sR = sR & "|" & cEFD("0", 0, 2, "N") '12
    sR = sR & "|" & cEFD(vIcmsSal, 0, 2, "N") '13
    sR = sR & "|" & cEFD("0", 0, 2, "N") '14
    sR = sR & "|" & cEFD("0", 0, 2, "N") '15
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("E110", "1")
    grvFile nmFile, sR
    
    RegistroE116 vIcmsSal, Format(dtFinal, "MMYYYY")
End Function
Private Function RegistroE116(vRecolher As String, mesRef As String) As Integer
    Dim dtVenc As Date
    'ICMS todo dia 10
    dtVenc = "10/" & Left(mesRef, 2) + 1 & "/ " & Right(mesRef, 4)
    
    sR = "E116" '1
    sR = sR & "|" & cEFD("000", 3, 0, "C") '2
    sR = sR & "|" & cEFD(ChkVal(vRecolher, 0, 2), 0, 2, "N") '3
    sR = sR & "|" & cEFD(dtVenc, 8, 0, "D") '4
    sR = sR & "|" & cEFD("0", 0, 0, "C") '5
    sR = sR & "|" & cEFD("0", 15, 0, "C") '6
    sR = sR & "|" & cEFD("0", 1, 0, "C") '7
    sR = sR & "|" & cEFD("0", 0, 0, "C") '8
    sR = sR & "|" & cEFD("0", 0, 0, "C") '9
    sR = sR & "|" & cEFD(mesRef, 6, 0, "N") '10
    
    sR = "|" & sR & "|"
    grvFile nmFile, sR
    cArray = cArray + 1: vReg(cArray) = Array("E116", "1")
    
End Function


Private Function RegistroE500() As Integer
'(dtI As String, dtF As String) As Integer
    'REGISTRO E500: PERIODO DE APURACAO DO IPI
    sR = "E500"
    sR = sR & "|" & cEFD("0", 1, 0, "C")
    sR = sR & "|" & cEFD(dtInicial, 8, 0, "D")
    sR = sR & "|" & cEFD(dtFinal, 8, 0, "D")
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("E500", "1") 'Quantidade de reg E00
    grvFile nmFile, sR
    
End Function
Private Function RegistroE510() As Integer
    'REGISTRO E510: CONSOLIDACAO DOS VALORES DO IPI
    'totalizacao de CST, CFOP
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim c           As Integer
    c = 0
    
    '### SAIDA ###########################################
    sSQL = "SELECT " & _
        "faturamentonfe.ide_dEmi, " & _
        "faturamentonfeitens.IPI_CST, " & _
        "faturamentonfeitens.det_CFOP, " & _
        "Sum(faturamentonfeitens.IPI_vBC) AS IPI_vBC, " & _
        "sum(faturamentonfeitens.IPI_vIPI) AS IPI_vIPI " & _
        "FROM " & _
        "faturamentonfe INNER JOIN " & _
        "faturamentonfeitens ON faturamentonfe.IdNFe = " & _
        "faturamentonfeitens.IdNFe " & _
        "WHERE faturamentonfe.ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
        "GROUP BY " & _
        "faturamentonfeitens.IPI_CST, " & _
        "faturamentonfeitens.det_CFOP"
        '"faturamentonfe.ide_dEmi"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Rst.Close
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                c = c + 1
        
                sR = "E510"
                sR = sR & "|" & cEFD(cNull(Rst.Fields("det_CFOP")), 4, 0, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("IPI_CST")), 4, 0, "C")
                sR = sR & "|" & cEFD("0", 0, 2, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("IPI_vBC")), 0, 2, "N")
                sR = sR & "|" & cEFD(cNull(Rst.Fields("IPI_vIPI")), 0, 2, "N")
                sR = "|" & sR & "|"
                
                grvFile nmFile, sR
                Rst.MoveNext
            Loop
            Rst.Close
    End If
    '################################################################
    '################################################################
    '################################################################
    
    '### ENTRADA ####################################################
    Dim r               As Integer
    Dim rDados(10)      As Variant
    
    sSQL = "SELECT " & _
        "faturamentonfeentrada.IdNFe, " & _
        "faturamentonfeentrada.ide_dEmi, " & _
        "faturamentonfeentradaitens.IPI_CST, " & _
        "faturamentonfeentradaitens.det_CFOP, " & _
        "Sum(faturamentonfeentradaitens.IPI_vBC) AS IPI_vBC, " & _
        "sum(faturamentonfeentradaitens.IPI_vIPI) AS IPI_vIPI " & _
        "FROM " & _
        "faturamentonfeentrada INNER JOIN " & _
        "faturamentonfeentradaitens ON faturamentonfeentrada.IdNFe = " & _
        "faturamentonfeentradaitens.IdNFe " & _
        "WHERE faturamentonfeentrada.ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
        "GROUP BY " & _
        "faturamentonfeentradaitens.IPI_CST, " & _
        "faturamentonfeentradaitens.det_CFOP"
        '"faturamentonfeentrada.ide_dEmi"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Rst.Close
        Else
            Rst.MoveFirst
            
            MontarTabelaTemporaria_RegE510_ent
            Do Until Rst.EOF
                
                r = 0
                rDados(r) = Array("cfop", cCFOPEntrada(Rst.Fields("IdNFe"), cNull(Rst.Fields("det_CFOP"))), "S"): r = r + 1
                rDados(r) = Array("ipiCst", cCstIpiEntrada(cNull(Rst.Fields("IPI_CST"))), "S"): r = r + 1
                rDados(r) = Array("ipivBC", ChkVal(cNull(Rst.Fields("IPI_vBC")), 0, 2), "S"): r = r + 1
                rDados(r) = Array("vIPI", ChkVal(cNull(Rst.Fields("IPI_vIPI")), 0, 2), "S"): r = r + 1
                RegistroIncluir "tmp_reg520ent", rDados, r - 1
                
                Rst.MoveNext
            Loop
            Rst.Close
            
            sSQL = "SELECT cfop, ipicst, SUM(ipivbc) AS ipivbc, SUM(vipi) AS vipi " & _
                   "FROM tmp_reg520ent " & _
                   "GROUP BY cfop, ipicst"
                   
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    Rst.Close
                Else
                    Rst.MoveFirst
                    Do Until Rst.EOF
                        c = c + 1
                        sR = "E510"
                        sR = sR & "|" & cEFD(cNull(Rst.Fields("cfop")), 4, 0, "N")
                        sR = sR & "|" & cEFD(cNull(Rst.Fields("IPICST")), 4, 0, "C")
                        sR = sR & "|" & cEFD("0", 0, 2, "N")
                        sR = sR & "|" & cEFD(cNull(Rst.Fields("IPIvBC")), 0, 2, "N")
                        sR = sR & "|" & cEFD(cNull(Rst.Fields("vIPI")), 0, 2, "N")
                        sR = "|" & sR & "|"
                        grvFile nmFile, sR
                        Rst.MoveNext
                    Loop
                    Rst.Close
            End If
    End If
    '##################################################################################
    cArray = cArray + 1: vReg(cArray) = Array("E510", c)
    
    
End Function

Private Sub MontarTabelaTemporaria_RegE510_ent()

    BD.Execute "DROP TABLE IF EXISTS tmp_reg520ent"
    BD.Execute "CREATE TABLE IF NOT EXISTS tmp_reg520ent " & _
               "(Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "cfop VARCHAR(100) default Null," & _
               "ipicst VARCHAR(100) default Null," & _
               "ipivbc VARCHAR(100) default Null," & _
               "vipi VARCHAR(100) default Null," & _
               "PRIMARY KEY (Id))"
End Sub

Private Sub MontarTabelaTemporaria_Reg0200(DtIni As Date, DtFin As Date)
    Dim sSQL    As String
    Dim Rst     As Recordset
    '### Monta a tabela
    BD.Execute "DROP TABLE IF EXISTS tmp_reg0200"
    BD.Execute "CREATE TABLE IF NOT EXISTS tmp_reg0200 " & _
               "(Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "idprod VARCHAR(100) default Null," & _
               "descr VARCHAR(100) default Null," & _
               "codbar VARCHAR(100) default Null," & _
               "un VARCHAR(100) default Null," & _
               "ncm VARCHAR(100) default Null," & _
               "PRIMARY KEY (Id))"
               
            '### ENTRADA
    sSQL = "SELECT " & _
           "faturamentonfeentrada.ide_dEmi, estoqueproduto.Descricao, " & _
           "estoqueproduto.NCM, estoqueproduto.CodigoBarras, " & _
           "estoqueproduto.Unidade, faturamentonfeentradaitens.det_IdProduto " & _
           "FROM " & _
           "faturamentonfeentrada INNER JOIN " & _
           "faturamentonfeentradaitens ON faturamentonfeentrada.IdNFe = " & _
           "faturamentonfeentradaitens.IdNFe INNER JOIN " & _
           "estoqueproduto ON faturamentonfeentradaitens.det_IdProduto = " & _
           "estoqueproduto.ID " & _
           "WHERE " & _
           "faturamentonfeentrada.ide_dEmi BETWEEN '2013-11-01' AND '2013-11-31' " & _
           "GROUP BY " & _
           "faturamentonfeentradaitens.det_IdProduto"
End Sub


Private Function RegistroE520(op As Boolean) As Integer
    'REGISTRO E520: APURACAO DO IPI
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim vIpiDeb     As String
    Dim vIpiCre     As String
    Dim vIpiSal     As String
    If op = True Then
            '### PEGA VALOR TOTAL DO DEBITO
            sSQL = "SELECT SUM(total_vIPI) AS vIpi " & _
                    "FROM faturamentonfe " & _
                    "WHERE ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
                    "ORDER BY ide_dEmi, id"
            
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    vIpiDeb = 0
                Else
                    Rst.MoveFirst
                    vIpiDeb = Rst.Fields("vIpi")
            End If
            Rst.Close
            
            '### PEGA VALOR TOTAL DO CREDITO
            sSQL = "SELECT SUM(total_vIPI) AS vIpi " & _
                    "FROM faturamentonfeentrada " & _
                    "WHERE ide_dEmi BETWEEN '" & Format(dtInicial, "YYYY-MM-DD") & "' AND '" & Format(dtFinal, "YYYY-MM-DD") & "' " & _
                    "ORDER BY ide_dEmi, id"
            
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    vIpiCre = 0
                Else
                    Rst.MoveFirst
                    vIpiCre = Rst.Fields("vIpi")
            End If
            Rst.Close
        Else
            vIpiDeb = 0
            vIpiCre = 0
        
    End If
            
    '########### SALDO IPI ###########################
    vIpiSal = Val(ChkVal(vIpiCre, 0, 2)) - Val(ChkVal(vIpiDeb, 0, 2))
    vIpiSal = ChkVal(vIpiSal, 0, 2)
    '#################################################
    
    sR = "E520" '1
    sR = sR & "|" & cEFD("0", 0, 2, "N") '2'
    sR = sR & "|" & cEFD(vIpiDeb, 0, 2, "N") '3
    sR = sR & "|" & cEFD(vIpiCre, 0, 2, "N") '4
    sR = sR & "|" & cEFD("0", 0, 2, "N") '5
    sR = sR & "|" & cEFD("0", 0, 2, "N") '6
    
    If vIpiSal > 0 Then
            sR = sR & "|" & cEFD(vIpiSal, 0, 2, "N") '7 Saldo Credor
            sR = sR & "|" & cEFD("0", 0, 2, "N")  '8 Saldo Devedor
        Else
            sR = sR & "|" & cEFD("0", 0, 2, "N") '7 Saldo Credor
            sR = sR & "|" & cEFD(Val(vIpiSal) * -1, 0, 2, "N") '8 Saldo Devedor
    End If
    
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("E520", "1")
    grvFile nmFile, sR
End Function

Private Function RegistroE990() As Integer
    'REGISTRO E990: ENCERRAMENTO DO BLOCO E
    Dim qtd_lin As Integer
    
    qtd_lin = QtdLinRegistro("E")
    
    sR = "E990"
    sR = sR & "|" & qtd_lin
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("E990", "1")
    grvFile nmFile, sR
End Function
Private Function RegistroG001() As Integer
    'REGISTRO G001: ABERTURA DO BLOCO G
    
    sR = "G001"
    sR = sR & "|" & "1"
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("G001", "1")
    grvFile nmFile, sR
End Function
Private Function RegistroK001() As Integer
    'REGISTRO K001: ABERTURA DO BLOCO K
    
    sR = "K001"
    sR = sR & "|" & "1"
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("K001", "1")
    grvFile nmFile, sR
End Function
Private Function RegistroG990() As Integer
    'REGISTRO G990: ENCERRAMENTO DO BLOCO G
    Dim qtd_lin As Integer
    
    qtd_lin = QtdLinRegistro("G")
    
    sR = "G990"
    sR = sR & "|" & qtd_lin
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("G990", "1")
    grvFile nmFile, sR
End Function
Private Function RegistroK990() As Integer
    'REGISTRO K990: ENCERRAMENTO DO BLOCO K
    Dim qtd_lin As Integer
    
    qtd_lin = QtdLinRegistro("K")
    
    sR = "K990"
    sR = sR & "|" & qtd_lin
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("K990", "1")
    grvFile nmFile, sR
End Function

Private Function RegistroH001(infInvent As Boolean) As Integer
    'REGISTRO H001: ABERTURA DO BLOCO H
    'inf - Informar inventario
    sR = "H001"
    sR = sR & "|" & IIf(infInvent = True, "0", "1")
    sR = "|" & sR & "|"

    cArray = cArray + 1: vReg(cArray) = Array("H001", "1")
    grvFile nmFile, sR
End Function
Private Function RegistroH005(infInvent As Boolean, mot_inv As String) As Integer '(infInvent As Boolean, dtF As String, mot_inv As String) As Integer
    'REGISTRO H005: TOTAIS DO INVENTÁRIO
    Dim vl      As String
    Dim vlTMP   As String
    Dim vCusto  As String
    Dim vSaldo  As String
    Dim sSQL    As String
    Dim Rst     As Recordset
    'inf - Informar inventario
    If infInvent = False Then
        Exit Function
    End If
    sSQL = "SELECT Saldo, Custo FROM EstoqueProduto" & _
           " WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND Status = 'ATIVO'" & _
           " AND IncluirBalanco = 1 ORDER BY Descricao"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            vl = 0
        Else
            Rst.MoveFirst
            vl = 0
            Do Until Rst.EOF
                vCusto = ChkVal(cNull(Rst.Fields("Custo")), 0, cDecMoeda)
                vSaldo = ChkVal(cNull(Rst.Fields("Saldo")), 0, cDecQtd)
                
                'Incluido para zerar saldos negativos - Leonardo Aquino 19.02.2015
                vSaldo = IIf(Val(vSaldo) < 0, 0, vSaldo)
                
                vlTMP = Val(vSaldo) * Val(vCusto)
                vlTMP = ChkVal(vlTMP, 0, cDecMoeda)
                
                vl = Val(ChkVal(vl, 0, cDecMoeda)) + Val(vlTMP)

                Rst.MoveNext
            Loop
    End If
    Rst.Close
    sR = "H005"
    sR = sR & "|" & cEFD(dtFinal, 8, 0, "D")
    sR = sR & "|" & cEFD(vl, 0, 2, "N")
    sR = sR & "|" & cEFD(mot_inv, 2, 0, "C")
    sR = "|" & sR & "|"
    'Debug.Print "Total H005: " & vl
    cArray = cArray + 1: vReg(cArray) = Array("H005", "1")
    grvFile nmFile, sR
End Function
Private Function RegistroH010(infInvent As Boolean) As Integer
    'REGISTRO H010: INVENTÁRIO
    
    
    Dim vl_Item As String
    Dim qSaldo  As String
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim i       As Integer
    
    'inf - Informar inventario
    If infInvent = False Then
        Exit Function
    End If
    
    sSQL = "SELECT * FROM EstoqueProduto" & _
           " WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND Status = 'ATIVO'" & _
           " AND IncluirBalanco = 1 ORDER BY Descricao"
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            
        Else
            Rst.MoveFirst
            i = 0
            
            'Dim sItem As String
            'sItem = 0
            Do Until Rst.EOF
                
                qSaldo = ChkVal(cNull(Rst.Fields("Saldo")), 0, 3)
                qSaldo = IIf(Val(qSaldo) < 0, 0, qSaldo)
                
                sR = "H010"
                sR = sR & "|" & cEFD(Rst.Fields("id"), 60, 0, "C") 'Cod do item
                sR = sR & "|" & cEFD(Rst.Fields("Unidade"), 6, 2, "C") 'Unidade
                sR = sR & "|" & cEFD(qSaldo, 0, 3, "N") 'Saldo do item
                sR = sR & "|" & cEFD(ChkVal(cNull(Rst.Fields("Custo")), 0, cDecMoeda), 0, 6, "N") 'Valor Unitario do item
                vl_Item = Val(ChkVal(qSaldo, 0, cDecQtd)) * Val(ChkVal(cNull(Rst.Fields("Custo")), 0, cDecMoeda))
                'sItem = Val(ChkVal(sItem, 0, 2)) + Val(ChkVal(vl_Item, 0, 2))
                sR = sR & "|" & cEFD(vl_Item, 0, 2, "N") 'Valor total do item
                sR = sR & "|" & cEFD("0", 1, 0, "C") 'IND_PROP
                sR = sR & "|" & cEFD("", 60, 0, "C") 'COD_PART
                sR = sR & "|" & cEFD("", 0, 0, "C") 'txt_compl
                sR = sR & "|" & cEFD("1000", 0, 0, "C") 'COD_CTA
                sR = sR & "|" & cEFD(vl_Item, 0, 2, "N") 'VL_ITEM_IR - Valor total do item para IR
                sR = "|" & sR & "|"
                i = i + 1
                grvFile nmFile, sR
                Rst.MoveNext
                
            Loop
    End If
    'Debug.Print "Total H010: " & sItem
    cArray = cArray + 1: vReg(cArray) = Array("H010", i)
End Function
Private Function RegistroH990() As Integer
    'REGISTRO H990: ENCERRAMENTO DO BLOCO H
    Dim qtd_lin As Integer
    
    qtd_lin = QtdLinRegistro("H")
    
    sR = "H990"
    sR = sR & "|" & qtd_lin
    sR = "|" & sR & "|"
    
    cArray = cArray + 1: vReg(cArray) = Array("H990", "1")
    grvFile nmFile, sR
End Function
Private Function Registro1001() As Integer
    'REGISTRO 1001: ABERTURA DO BLOCO 1
    sR = "1001"
    sR = sR & "|" & "0"
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("1001", "1")
    grvFile nmFile, sR
End Function

Private Function Registro1990() As Integer
    'REGISTRO 1990: ENCERRAMENTO DO BLOCO 1
    Dim qtd_lin As Integer
    
    qtd_lin = QtdLinRegistro("1")
    
    sR = "1990"
    sR = sR & "|" & qtd_lin
    sR = "|" & sR & "|"
    cArray = cArray + 1: vReg(cArray) = Array("1990", "1")
    grvFile nmFile, sR
End Function
Private Function Registro9001() As Integer
    'REGISTRO 9001: ABERTURA DO BLOCO 9
    
    sR = "9001"
    sR = sR & "|" & "0"
    sR = "|" & sR & "|"
    grvFile nmFile, sR
    cArray = cArray + 1: vReg(cArray) = Array("9001", "1") 'Quantidade de reg 9001
End Function
Private Function Registro9900() As Integer
    'REGISTRO 9900: REGISTROS DO ARQUIVO
    Dim i As Integer
    
    Dim qtd_lin As Integer
    
    cArray = cArray + 1: vReg(cArray) = Array("9990", "1")
    cArray = cArray + 1: vReg(cArray) = Array("9999", "1")
    cArray = cArray + 1: vReg(cArray) = Array("9900", cArray)
    
    'For i = 1 To cArray
    '    Debug.Print vReg(i)(0) & " - " & vReg(i)(1)
    '    qtd_lin = qtd_lin + vReg(i)(1)
    'Next
    
    
    For i = 1 To cArray
        sR = "9900"
        sR = sR & "|" & vReg(i)(0)
        If vReg(i)(0) = "9900" Then
                sR = sR & "|" & vReg(i)(1)
            Else
                sR = sR & "|" & IIf(InStr(vReg(i)(0), "990") <> 0, 1, vReg(i)(1))
        End If
        sR = "|" & sR & "|"
        'Debug.Print vReg(i)(0) & " - " & vReg(i)(1)
        grvFile nmFile, sR
        'qtd_lin = qtd_lin + vReg(i)(1)
    Next
End Function
Private Function Registro9990() As Integer
    'REGISTRO 9990: ENCERRAMENTO DO BLOCO 9
    Dim qtd_lin As Integer
    qtd_lin = QtdLinRegistro("9") - 1
    
    sR = "9990"
    sR = sR & "|" & qtd_lin
    sR = "|" & sR & "|"
    grvFile nmFile, sR
    '*** Registro totalizado no 9900 'cArray=cArray+1: vReg(cArray) = Array("9990", "1")
End Function
Private Function Registro9999() As Integer
    'REGISTRO 9999: ENCERRAMENTO DO ARQUIVO DIGITAL
    Dim t As Integer
    Dim i As Integer
    t = -1
    
    For i = 1 To cArray
        t = t + vReg(i)(1)
        'Debug.Print vReg(i)(0) & " ==> " & vReg(i)(1)
    Next
    sR = "9999"
    sR = sR & "|" & t
    sR = "|" & sR & "|"
    grvFile nmFile, sR
    
End Function
Private Function cCFOPEntrada(sChv As String, sCFOP As String) As String
    '### 16.12.2013
    '### Converte o codigo do CFOP em caso de entrada
    'sChv - chave da nfe de entrada
    'sCFOP - CFOP da transacao
    Dim ufEmpresa   As String
    Dim ufOrig      As String
    Dim n1          As Integer
    Dim n2          As Integer
    
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    sSQL = "SELECT * FROM faturamentonfeentrada " & _
            "WHERE idNFe = '" & sChv & "' " & _
            "ORDER BY id"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nota Fiscal de ENTRADA não encontrada!", vbCritical, App.EXEName
            cCFOPEntrada = sCFOP
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            ufOrig = Rst.Fields("emit_uf")
            Rst.Close
    End If
    ufEmpresa = PgDadosEmpresa(ID_Empresa).UF
    sCFOP = rc(sCFOP)
    n1 = Left(sCFOP, 1)
    n2 = Mid(sCFOP, 2, Len(sCFOP))
    
    '#############################################################
    '# Em casos de lancamento de cupom fiscal cfop 5929
    '# escriturar com 1102.
    '# RJ,17.12.2013
    '# Orientacao: Glorinha - Argos
    '#############################################################
    If n2 = "929" Then
        n2 = "102"
    End If
    If sCFOP = "5405" Then
        cCFOPEntrada = "1403"
        Exit Function
    End If
    '#############################################################
    
            Select Case n1
                Case 5
                    n1 = 1
                Case 6
                    n1 = 2
                Case Else
                    n1 = n1
            End Select
            cCFOPEntrada = n1 & n2
End Function
Private Function cCstIpiEntrada(sCST As String) As String
    '### 17.12.2013
    '### Converte o codigo do CST em caso de entrada
    '###
    
    Dim n1      As Integer
    Dim n2      As Integer
    Dim nCST    As String
    If Trim(sCST) = "" Then
        cCstIpiEntrada = "49"
        Exit Function
    End If
    
    If Val(sCST) <= 49 Then
        cCstIpiEntrada = sCST
        Exit Function
    End If
    
    n1 = Left(sCST, 1)
    n2 = Right(sCST, 1)
    
    If n1 = 0 Then
            n1 = 5
        ElseIf n1 = 5 Then
            n1 = 0
        ElseIf n1 = 4 Then
            n1 = 9
        ElseIf n1 = 9 Then
            n1 = 4
    End If
    nCST = n1 & n2
    'If nCST = "09" Then
    '        nCST = "99"
    '    ElseIf nCST = "09" Then
    '        nCST = 49
    'End If
    cCstIpiEntrada = nCST
    
End Function
