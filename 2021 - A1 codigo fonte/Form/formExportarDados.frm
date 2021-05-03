VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formExportarDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Dados"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8880
   Begin VB.Frame frmNFe 
      Caption         =   "Exportar XML da NF-e"
      Height          =   1875
      Left            =   120
      TabIndex        =   14
      Top             =   3780
      Width           =   8295
      Begin MSComDlg.CommonDialog cd 
         Left            =   7740
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton btoDestinoXMLNFe 
         Caption         =   "..."
         Height          =   315
         Left            =   7080
         TabIndex        =   22
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox txtDestXML 
         Height          =   315
         Left            =   840
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1380
         Width           =   6135
      End
      Begin VB.CommandButton btoExportarXMLNFe 
         Caption         =   "Exportar"
         Height          =   435
         Left            =   6660
         TabIndex        =   19
         Top             =   180
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpPeriodo 
         Height          =   315
         Left            =   1380
         TabIndex        =   17
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   57671683
         CurrentDate     =   40989
      End
      Begin VB.CheckBox chkNFeSaida 
         Caption         =   "NF-e Saida"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1080
         Width           =   1755
      End
      Begin VB.CheckBox chkNFeEntrada 
         Caption         =   "NF-e Entrada"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   780
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Destino:"
         Height          =   255
         Left            =   60
         TabIndex        =   20
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Periodo da NF-e:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   3300
      Visible         =   0   'False
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton btoGerarArquivo 
      Caption         =   "&Gerar Arquivo"
      Height          =   435
      Left            =   6780
      TabIndex        =   12
      Top             =   3300
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo:"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   900
      Width           =   4695
      Begin MSComCtl2.DTPicker dtpDtIni 
         Height          =   315
         Left            =   720
         TabIndex        =   8
         Top             =   390
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57671681
         CurrentDate     =   40977
      End
      Begin MSComCtl2.DTPicker dtpDtFinal 
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Top             =   390
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57671681
         CurrentDate     =   40977
      End
      Begin VB.Label Label2 
         Caption         =   "De:"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "Até:"
         Height          =   195
         Left            =   2520
         TabIndex        =   10
         Top             =   450
         Width           =   315
      End
   End
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Width           =   8055
   End
   Begin VB.Frame frameFORTES 
      Caption         =   "Empresa: FORTES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1860
      Width           =   8295
      Begin VB.ComboBox cboAliqEsp 
         Height          =   315
         Left            =   2580
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox txtComentario 
         Height          =   285
         Left            =   180
         MaxLength       =   40
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   900
         Width           =   7935
      End
      Begin VB.Label Label5 
         Caption         =   "Regime de Aliquotas Especificas"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Comentarios"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Selecione a Empresa"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "formExportarDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function cvtFortes(sTexto As String, tpDado As String) As String
    Select Case UCase(tpDado)
        Case "D" 'Data
            cvtFortes = Format(sTexto, "YYYYMMDD")
        Case Else
            cvtFortes = sTexto
    End Select
End Function



Private Function GerarArquivoFortes() As Boolean
    '##############################################################
    '### 09/03/2012
    '### Layout de Importacao Fortes AC Fiscal versao 60
    '##############################################################
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim l       As Integer 'numero de linhas no arquivo
    Dim sTxt    As String
    Dim nmFile  As String
    
    nmFile = App.Path & "\Fiscal-" & Format(Date, "YYYYMMDD") & ".fs"
    If Dir(nmFile) <> "" Then
        Kill nmFile
    End If
    l = 0
    '##############################################################
    '### Registro Tipo CAB - Cabecalho
    '##############################################################
    
    sTxt = "CAB"
    sTxt = sTxt & "|" & "60"
    sTxt = sTxt & "|" & App.ProductName
    sTxt = sTxt & "|" & cvtFortes(Date, "D")
    sTxt = sTxt & "|" & PgDadosEmpresa(ID_Empresa).Nome
    sTxt = sTxt & "|" & cvtFortes(dtpDtIni.Value, "D")
    sTxt = sTxt & "|" & cvtFortes(dtpDtFinal.Value, "D")
    sTxt = sTxt & "|" & txtComentario.Text
    sTxt = sTxt & "|" & cboAliqEsp.Text
    l = l + 1
    grvFile nmFile, sTxt
    
    '##############################################################
    '### Registro Tipo PAR - Participantes dos Documentos Fiscais
    '##############################################################
    sSQL = "SELECT * FROM Clientes"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            'Rst.MoveLast
            
            Rst.MoveFirst
            Do Until Rst.EOF
                Status (Rst.RecordCount)
                sTxt = "PAR"
                sTxt = sTxt & "|" & Rst.Fields("ID")
                sTxt = sTxt & "|" & Rst.Fields("xNome")
                sTxt = sTxt & "|" & Rst.Fields("UF")
                sTxt = sTxt & "|" & Rst.Fields("Doc")
                sTxt = sTxt & "|" & Rst.Fields("IE")
                sTxt = sTxt & "|" & Rst.Fields("IM")
                sTxt = sTxt & "|" & "" '8 - ISS Digital
                sTxt = sTxt & "|" & "" '9 - DIEF
                sTxt = sTxt & "|" & "" '10 - DIC
                sTxt = sTxt & "|" & "" '11 - DEMMS
                sTxt = sTxt & "|" & "" '12 - Orgão Publico
                sTxt = sTxt & "|" & "" '13 - Livro Eletronico
                sTxt = sTxt & "|" & "" '14 - Fornecedor de Prod Primario
                sTxt = sTxt & "|" & "" '15 - Sociedade Simples
                sTxt = sTxt & "|" & "35" '16 - Tipo Logradouro
                sTxt = sTxt & "|" & Rst.Fields("xLgr")
                sTxt = sTxt & "|" & Rst.Fields("Nro")
                sTxt = sTxt & "|" & Rst.Fields("xCpl")
                sTxt = sTxt & "|" & "01" '20 - Tipo de Bairro
                sTxt = sTxt & "|" & Rst.Fields("xBairro")
                sTxt = sTxt & "|" & Rst.Fields("CEP")
                sTxt = sTxt & "|" & Mid(Rst.Fields("xMun"), 2, Len(Rst.Fields("xMun")))
                sTxt = sTxt & "|" & "" '24 - DDD
                sTxt = sTxt & "|" & Rst.Fields("Fone")
                sTxt = sTxt & "|" & Rst.Fields("Suframa")
                sTxt = sTxt & "|" & "" '27 - Substituto ISS
                sTxt = sTxt & "|" & "" '28 - Conta Remetente/Prestador
                sTxt = sTxt & "|" & "" '29 - Conta Dest/Tomador
                sTxt = sTxt & "|" & "1058"
                sTxt = sTxt & "|" & "N" '31 - Exterior
                sTxt = sTxt & "|" & "" '32 - Isento de ICMS
                sTxt = sTxt & "|" & Rst.Fields("Email")
                l = l + 1
                grvFile nmFile, sTxt
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    
    '##############################################################
    '### Registro Tipo PRO - Produtos
    '##############################################################
    sSQL = "SELECT * FROM EstoqueProduto"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            'Rst.MoveLast
            
            Rst.MoveFirst
            Do Until Rst.EOF
                Status (Rst.RecordCount)
                sTxt = "PRO"
                sTxt = sTxt & "|" & Rst.Fields("ID")
                sTxt = sTxt & "|" & Rst.Fields("Descricao")
                sTxt = sTxt & "|" & Rst.Fields("ID")
                sTxt = sTxt & "|" & Rst.Fields("NCM")
                sTxt = sTxt & "|" & Rst.Fields("Unidade")
                sTxt = sTxt & "|" & "" ' 07 - Unidade Medida DIEF
                sTxt = sTxt & "|" & "" ' 08 - Unidade Medida CENFOP
                sTxt = sTxt & "|" & Rst.Fields("NCM")  ' 09 - Classificacao Fiscal
                sTxt = sTxt & "|" & "" ' 09 - Grupo
                sTxt = sTxt & "|" & "" ' 10 - Genero
                sTxt = sTxt & "|" & cNull(Rst.Fields("CodigoBarras"))
                sTxt = sTxt & "|" & "" ' 13 - Reducao
                sTxt = sTxt & "|" & "" ' 14 - Codigo GAM57
                sTxt = sTxt & "|" & Rst.Fields("ICMSCST")
                sTxt = sTxt & "|" & Rst.Fields("IPICST")
                sTxt = sTxt & "|" & "" 'Rst.Fields("PISCST")
                sTxt = sTxt & "|" & "" 'Rst.Fields("COFINS_CST")
                sTxt = sTxt & "|" & "" ' 19 - Codigo ANP
                sTxt = sTxt & "|" & "" ' 20 CST ICMS Simples Nacional
                sTxt = sTxt & "|" & "" ' 21 - CSOSN
                sTxt = sTxt & "|" & "" ' 22 - Produto Especifico
                sTxt = sTxt & "|" & "" ' 23 - Tipo de Medicamento
                l = l + 1
                grvFile nmFile, sTxt
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    
    '##############################################################
    '### Registro Tipo NFM - Notas Fiscais de Mercadoria
    '##############################################################
    sSQL = "SELECT FaturamentoNFe.*, FaturamentoNFeItens.* " & _
           "FROM FaturamentoNFe, FaturamentoNFeItens " & _
           "WHERE FaturamentoNFe.ide_dEmi BETWEEN '" & Format(dtpDtIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & "' " & _
           "AND FaturamentoNFe.IdNFe= FaturamentoNFeItens.idNFe"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            'Rst.MoveLast
            
            Rst.MoveFirst
            Do Until Rst.EOF
                Status (Rst.RecordCount)
                sTxt = "NFM"
                sTxt = sTxt & "|" & "" ' 2 - Estabelecimento
                sTxt = sTxt & "|" & IIf(Rst.Fields("ide_tpNF") = 0, "E", "S")
                sTxt = sTxt & "|" & "NFE"
                sTxt = sTxt & "|" & "S"
                sTxt = sTxt & "|" & "" ' 6 - AIDF
                sTxt = sTxt & "|" & Rst.Fields("ide_Serie")
                sTxt = sTxt & "|" & "" '8 - Sub SerieRst.Fields("ide_SSerie")
                sTxt = sTxt & "|" & Rst.Fields("ide_nNF")
                sTxt = sTxt & "|" & "" '10 - Formulario Inicial
                sTxt = sTxt & "|" & "" '11 - Formulario Final
                sTxt = sTxt & "|" & Rst.Fields("ide_dEmi")
                
                If Not IsNull(Rst.Fields("canc_nProt")) Then
                        sTxt = sTxt & "|" & "1"
                    Else
                        sTxt = sTxt & "|" & "0"
                End If
                
                sTxt = sTxt & "|" & "" '14 - Dt. Entr/Saida
                sTxt = sTxt & "|" & Rst.Fields("dest_idDest")
                sTxt = sTxt & "|" & "N" '16 - Vinculo GNRE
                sTxt = sTxt & "|" & "" '17 - GNRE ICMS
                sTxt = sTxt & "|" & "" '18 - GNRE Mes/Ano
                sTxt = sTxt & "|" & "" '19 - GNRE Convenio
                sTxt = sTxt & "|" & "" '20 - GNRE Data Venc.
                sTxt = sTxt & "|" & "" '21 - GNRE Data Recolhimento
                sTxt = sTxt & "|" & "" '22 - GNRE Banco
                sTxt = sTxt & "|" & "" '23 - GNRE Agencia
                sTxt = sTxt & "|" & "" '24 - GNRE Agencia DV
                sTxt = sTxt & "|" & "" '25 - GNRE Autenticado
                sTxt = sTxt & "|" & Rst.Fields("Total_vProd")
                sTxt = sTxt & "|" & Rst.Fields("Total_vFrete")
                sTxt = sTxt & "|" & Rst.Fields("Total_vSeg")
                sTxt = sTxt & "|" & Rst.Fields("Total_vOutro")
                sTxt = sTxt & "|" & "" '30 - ICMS Importacao
                sTxt = sTxt & "|" & "" '31 - ICMS Importacao Diferimento
                sTxt = sTxt & "|" & Rst.Fields("Total_vIPI")
                sTxt = sTxt & "|" & "" '33 - Substituicao retido
                sTxt = sTxt & "|" & "" '34 - Servico ISS
                sTxt = sTxt & "|" & Rst.Fields("Total_vDesc")
                sTxt = sTxt & "|" & Rst.Fields("Total_vNF")
                sTxt = sTxt & "|" & "" '37 - Quantidade de Intes/Produtos
                sTxt = sTxt & "|" & "" '38 - ST Recolhes
                sTxt = sTxt & "|" & "" '39 - Antecipar Recolher
                sTxt = sTxt & "|" & "" '40 - Diferencial de Aliquota
                sTxt = sTxt & "|" & "" '41 - Valor Contabil ST
                sTxt = sTxt & "|" & Rst.Fields("ICMS_vBCST") '42 - BC ICMS ST
                sTxt = sTxt & "|" & "" '43 - Valor Contabil Antecipado
                sTxt = sTxt & "|" & "" '44 - ISS Retido
                sTxt = sTxt & "|" & "" '45 - Data de Retencao do ISS
                sTxt = sTxt & "|" & "" '46 - Servico
                sTxt = sTxt & "|" & "" '47 - Data Entrada no Estado
                sTxt = sTxt & "|" & IIf(Rst.Fields("transp_ModFrete") = 0, "R", "D") '48 - Frete por conta 'Tab 11
                sTxt = sTxt & "|" & "P" '49 - Fatura ' Tab 12
                sTxt = sTxt & "|" & "" '50 - Numero do EEC
                sTxt = sTxt & "|" & "" '51 - Numero do Cupom
                sTxt = sTxt & "|" & "" '52 - Receita tributavel COFINS
                sTxt = sTxt & "|" & "" '53 - Receita tributavel PIS
                sTxt = sTxt & "|" & "" '54 - Receita tributavel CSL 1
                sTxt = sTxt & "|" & "" '55 - Receita tributavel CSL 2
                sTxt = sTxt & "|" & "" '56 - Receita tributavel IRPJ 1
                sTxt = sTxt & "|" & "" '57 - Receita tributavel IRPJ 2
                sTxt = sTxt & "|" & "" '58 - Receita tributavel IRPJ 3
                sTxt = sTxt & "|" & "" '59 - Receita tributavel IRPJ 4
                sTxt = sTxt & "|" & "" '60 - COFINS Retido na fonte
                sTxt = sTxt & "|" & "" '61 - PIS Retido na fonte
                sTxt = sTxt & "|" & "" '62 - CSL Retido na fonte
                sTxt = sTxt & "|" & "" '63 - IRPJ Retido na fonte
                sTxt = sTxt & "|" & "" '64 - Gera transferencia
                sTxt = sTxt & "|" & "" '65 - Observacoes
                sTxt = sTxt & "|" & "" '66 - Aliquota ST
                sTxt = sTxt & "|" & Rst.Fields("idNFe") '67 - Chave Eletronica
                sTxt = sTxt & "|" & "" '68 - INSS Retido na Fonte
                sTxt = sTxt & "|" & "" '69 - BC COFINS / PIS nao cumulativo
                sTxt = sTxt & "|" & cNull(Rst.Fields("canc_xJust")) '70 - Motivo de cancelamento
                sTxt = sTxt & "|" & Rst.Fields("ide_NatOp") '71 - Natureza da Operacao
                sTxt = sTxt & "|" & "" '72 - Cod. informacao complementar
                sTxt = sTxt & "|" & "" '73 - Complemento das inf. complementares
                sTxt = sTxt & "|" & "" '74 - Hora da Saida
                sTxt = sTxt & "|" & Rst.Fields("emit_UF") '75 - UF de Embarque
                sTxt = sTxt & "|" & "" '76 - Local de embarque
                l = l + 1
                grvFile nmFile, sTxt
                Rst.MoveNext
            Loop
    'End If
    'rst.close
    '##############################################################
    '### Registro Tipo PNM - Produtos(Notas Fiscais de Mercadoria)
    '##############################################################
    'sSQL = "SELECT * FROM FaturamentoNFeItens"
    'Set Rst = RegistroBuscar(sSQL)
    'If Rst.BOF And Rst.EOF Then
    '    Else
            'Rst.MoveLast
            
            Rst.MoveFirst
            Do Until Rst.EOF
                Status (Rst.RecordCount)
                sTxt = "PNM"
                sTxt = sTxt & "|" & Rst.Fields("det_cProd")
                sTxt = sTxt & "|" & Rst.Fields("det_CFOP")
                sTxt = sTxt & "|" & "" '4 - CFOP transferencia
                sTxt = sTxt & "|" & "" '5 - CSTA
                sTxt = sTxt & "|" & "" '6 - CSTB
                sTxt = sTxt & "|" & Rst.Fields("det_uCom")
                sTxt = sTxt & "|" & Rst.Fields("det_qCom")
                sTxt = sTxt & "|" & Rst.Fields("det_vProd")
                sTxt = sTxt & "|" & Rst.Fields("IPI_vIPI")
                sTxt = sTxt & "|" & "3" '11 - Tipo Trib ICMS 'tab.13
                sTxt = sTxt & "|" & Rst.Fields("ICMS_vBC")
                sTxt = sTxt & "|" & Rst.Fields("ICMS_pICMS")
                sTxt = sTxt & "|" & Rst.Fields("ICMS_vBCST")
                sTxt = sTxt & "|" & Rst.Fields("ICMS_vICMSST")
                sTxt = sTxt & "|" & "" '16 - Tipo de recolhimento
                sTxt = sTxt & "|" & "" '17 - Tipo Substituicao
                sTxt = sTxt & "|" & "" '18 - Custo Aquisicao ST
                sTxt = sTxt & "|" & "" '19 - Perc. Agreg. Substituicao
                sTxt = sTxt & "|" & Rst.Fields("ICMS_vBCST")
                sTxt = sTxt & "|" & "" '21 - Aliq ST
                sTxt = sTxt & "|" & "" '22 - Credito Origem
                sTxt = sTxt & "|" & "" '23 - Subst ja recolhido
                sTxt = sTxt & "|" & "" '24 - Custo da Aquisicao Antecip.
                sTxt = sTxt & "|" & "" '25 - Perc. Agregacao antecipada
                sTxt = sTxt & "|" & "" '26 - Aliquota Interna
                sTxt = sTxt & "|" & "" '27 - Credito de Origem
                sTxt = sTxt & "|" & "" '28 - Antec. ja Recolhido
                sTxt = sTxt & "|" & "" '29 - Base de Calc. Dif. Aliquota
                sTxt = sTxt & "|" & "" '30 - Aliquota de Origem
                sTxt = sTxt & "|" & "" '31 - Aliquota Interna
                sTxt = sTxt & "|" & "" '32 - Tipo Trib. IPI 'tab.13
                sTxt = sTxt & "|" & Rst.Fields("IPI_vBC")
                sTxt = sTxt & "|" & Rst.Fields("IPI_pIPI")
                sTxt = sTxt & "|" & Rst.Fields("IPI_vIPI")
                sTxt = sTxt & "|" & Rst.Fields("IPI_CST") '36 - CST IPI 'tab.17
                sTxt = sTxt & "|" & Rst.Fields("COFINS_CST") '37 - CST COFINS 'tab.18
                sTxt = sTxt & "|" & Rst.Fields("PIS_CST") '38 - CST PIS 'tab.18
                sTxt = sTxt & "|" & Rst.Fields("COFINS_vBC")
                sTxt = sTxt & "|" & Rst.Fields("PIS_vBC")
                sTxt = sTxt & "|" & Rst.Fields("det_vFrete")
                sTxt = sTxt & "|" & Rst.Fields("det_vSeg")
                sTxt = sTxt & "|" & Rst.Fields("det_vDesc")
                
                
                Dim vTotalSemImp As String
                
                vTotalSemImp = (Val(ChkVal(Rst.Fields("det_vProd"), 0, cDecMoeda)) + Val(ChkVal(cNull(Rst.Fields("det_vFrete")), 0, cDecMoeda)) + Val(ChkVal(cNull(Rst.Fields("det_vSeg")), 0, cDecMoeda))) - Val(ChkVal(cNull(Rst.Fields("det_vDesc")), 0, cDecMoeda))
                vTotalSemImp = ChkVal(vTotalSemImp, 0, cDecMoeda)
                sTxt = sTxt & "|" & vTotalSemImp '44 - Valor Produto(Somatorio dos campos 9+41+42-43)
                
                sTxt = sTxt & "|" & "" '45 - Natureza da Receita COFINS
                sTxt = sTxt & "|" & "" '46 - Natureza da Receita PIS
                sTxt = sTxt & "|" & "" '47 - Indicador Especial - PRODEPE
                sTxt = sTxt & "|" & "" '48 - Codigo de Apuracao PRODEPE
                sTxt = sTxt & "|" & "" '49 - Cod. da ST do CSOSN
                sTxt = sTxt & "|" & "" '50 - CSOSN
                sTxt = sTxt & "|" & "" '51 - Tipo Calc COFINS
                sTxt = sTxt & "|" & "" '52 - Aliquota COFINS(%)
                sTxt = sTxt & "|" & "" '53 - Aliquota COFINS(R$)
                sTxt = sTxt & "|" & "" '54 - Valor COFINS
                sTxt = sTxt & "|" & "" '55 - Tipo Calc PIS
                sTxt = sTxt & "|" & "" '56 - Aliquota PIS(%)
                sTxt = sTxt & "|" & "" '57 - Aliquota PIS(R$)
                sTxt = sTxt & "|" & "" '58 - Valor PIS
                sTxt = sTxt & "|" & "" '59 - Codigo Ajuste Fiscal
                sTxt = sTxt & "|" & Rst.Fields("det_xPed")
                sTxt = sTxt & "|" & Rst.Fields("det_nItemPed")
                l = l + 1
                grvFile nmFile, sTxt
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    '##############################################################
    '### Registro Tipo TRA - Trailler
    '##############################################################
    
    sTxt = "TRA"
    sTxt = sTxt & "|" & l
    l = l + 1
    grvFile nmFile, sTxt
    Status (1)
    MsgBox "Arquivo gravado em " & nmFile & "."
End Function
Private Sub Status(Max As Long)
    pb.Min = 0
    pb.Max = Max
    DoEvents
    pb.Value = pb.Value + 1
    If pb.Value > 0 And pb.Value < Max Then
            pb.Visible = True
            Me.Enabled = False
        Else
            pb.Visible = False
            pb.Value = 0
            Me.Enabled = True
    End If
End Sub

Private Sub btoDestinoXMLNFe_Click()
    cd.ShowSave
    txtDestXML.Text = cd.filename
    
End Sub

Private Sub btoExportarXMLNFe_Click()
    ExportarXMLdaNFe dtpPeriodo.Value
End Sub

Private Sub btoGerarArquivo_Click()
    GerarArquivoFortes
End Sub

Private Sub cboAliqEsp_DropDown()
    With cboAliqEsp
        .Clear
        .AddItem "SIM"
        .AddItem "NÃO"
    End With
End Sub


Private Sub cboEmpresa_Click()
    Dim opcao As String
    opcao = Mid(cboEmpresa.Text, 1, 3)
    Select Case opcao
        Case "001"
        Case Else
    End Select
End Sub

Private Sub cboEmpresa_DropDown()
    With cboEmpresa
        .Clear
        .AddItem "001 - Fortes Informática"
    End With
End Sub

Private Sub cboFortesAliqEsp_DropDown()
   
End Sub


Private Sub Form_Load()
    LimpaFormulario Me
End Sub
Private Sub ExportarXMLdaNFe(periodo As String)

    Dim caminho             As String
    Dim arquivoCliente      As String
    Dim arquivoFornecedor   As String
    periodo = Format(periodo, "YYYYMM")
    
    '#######################################################################################
    '### Saida : Clientes
    If chkNFeSaida.Value = 1 Then
        caminho = PgDadosConfig.pBackup & "\Autorizados\" & periodo
        If Dir(caminho, vbDirectory) = "" Then
            MsgBox "Erro ao localizar a pasta com os dados"
            Exit Sub
        End If
        arquivoCliente = PgDadosConfig.pFileArmazenamento & "\Saida-" & RS(PgDadosEmpresa(ID_Empresa).CNPJ) & Format(periodo, "YYYYMM") & ".zip"
        Compacta arquivoCliente, caminho & "\*.*"
    End If
    '#######################################################################################
    
    '#######################################################################################
    '### Entrada : Fornecedores
    If chkNFeEntrada.Value = 1 Then
        caminho = PgDadosConfig.pXMLFornecedor & "\" & periodo
        If Dir(caminho, vbDirectory) = "" Then
            MsgBox "Erro ao localizar a pasta com os dados"
            Exit Sub
        End If
        arquivoFornecedor = PgDadosConfig.pFileArmazenamento & "\Entrada-" & RS(PgDadosEmpresa(ID_Empresa).CNPJ) & Format(periodo, "YYYYMM") & ".zip"
        Compacta arquivoFornecedor, caminho & "\*.*"
    End If
    '#######################################################################################
    
    Compacta txtDestXML.Text, PgDadosConfig.pFileArmazenamento & "\*-" & RS(PgDadosEmpresa(ID_Empresa).CNPJ) & Format(periodo, "YYYYMM") & ".zip"
    
    '### Exclui os Arquivos pre compactados
    If chkNFeEntrada.Value = 1 Then
        Kill arquivoFornecedor
    End If
    If chkNFeSaida.Value = 1 Then
        Kill arquivoCliente
    End If
    
    
    MsgBox "Arquivos gerados com sucesso!", vbInformation, App.EXEName
    
End Sub
