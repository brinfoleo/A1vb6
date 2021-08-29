VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formFaturamentoNFe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Emissão de NF-e"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   7560
   Begin VB.Frame frmAlteracoes 
      ForeColor       =   &H00E0E0E0&
      Height          =   1995
      Left            =   60
      TabIndex        =   23
      Top             =   6000
      Width           =   7395
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         Height          =   1635
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "formFaturamentoNFe.frx":0000
         Top             =   180
         Width           =   7275
      End
   End
   Begin VB.TextBox txtrefNFe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   44
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   5640
      Width           =   7275
   End
   Begin VB.Frame Frame2 
      Caption         =   "|  Tipo de Nota Fiscal  |"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   60
      TabIndex        =   7
      Top             =   480
      Width           =   7395
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   1875
         Left            =   1320
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Text            =   "formFaturamentoNFe.frx":0006
         Top             =   1740
         Width           =   5955
      End
      Begin MSComCtl2.DTPicker dtpSaida 
         Height          =   315
         Left            =   1320
         TabIndex        =   18
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   133169153
         CurrentDate     =   40561
      End
      Begin MSComCtl2.DTPicker dtpEmissao 
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   133169153
         CurrentDate     =   40561
      End
      Begin VB.TextBox txtNumNota 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox txtTpNF 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   300
         Width           =   795
      End
      Begin VB.CommandButton btoPesqTipoNF 
         Height          =   315
         Left            =   2160
         Picture         =   "formFaturamentoNFe.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Observações:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Saida:"
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   660
         TabIndex        =   15
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Num. Nota:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label lblTipoNF 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   300
         Width           =   4095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tp. Nota Fiscal:"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "|  Pedido  |"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   4320
      Width           =   7395
      Begin VB.CommandButton btoPesqPedido 
         Height          =   315
         Left            =   1620
         Picture         =   "formFaturamentoNFe.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtPedido 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   780
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblValor 
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4500
         TabIndex        =   6
         Top             =   660
         Width           =   2775
      End
      Begin VB.Label lblEmissao 
         Caption         =   "Emissao"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2100
         TabIndex        =   5
         Top             =   660
         Width           =   2295
      End
      Begin VB.Label lblNome 
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2100
         TabIndex        =   3
         Top             =   180
         Width           =   5175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pedido:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Vizualizar pré-venda"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Editar pré-venda"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Envia NF-e"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Manutenção da Tabela"
            ImageIndex      =   11
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5280
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":0720
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":0B72
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":0E8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":171E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":2970
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":324A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":3ADC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":436E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":55C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":58DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":5BF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":5FEB
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":68C5
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":759F
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFe.frx":7E79
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Chave da NF-e de Referencia:"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   5400
      Width           =   2235
   End
End
Attribute VB_Name = "formFaturamentoNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strTabela       As String

Dim strChaveAcesso  As String
Dim idTpNF          As Integer 'Tipo de Nota Fiscal
Dim idPedido        As Integer
'Dim numNota         As String
Dim idCliente       As Integer
Dim IdTransp        As Integer

Dim condPag         As Integer
Dim formPag         As Integer
Dim SRef            As String
'Dim Vendedor        As Integer
Dim FreteConta      As Integer
Dim TranspRetEnt    As Integer

Dim qVol            As String
Dim Esp             As String
Dim Marca           As String
Dim nVol            As String
Dim PesoB           As String
Dim PesoL           As String

Dim bcICMS          As Integer '0 - BC Mercadoria / 1 - BC Total da Nota


Dim msgRedICMS      As String
'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************
'Gerenciamento (ger)
Dim ger_Vendedor    As Integer
Dim ger_idPV        As Integer
'cabecario do Pedido (ide)
'Dim Versao          As String
Dim Id              As String
Dim ide_cUF         As String
Dim ide_cNF         As String
Dim ide_natOp       As String
Dim ide_indPag      As String
Dim ide_mod         As String
Dim ide_serie       As String
Dim ide_nNF         As String
Dim ide_dEmi        As String
Dim ide_hEmi        As String
Dim ide_dSaiEnt     As String
Dim ide_hSaiEnt     As String
Dim ide_tpNF        As String
Dim ide_idDest      As String
Dim ide_cMunFG      As String
Dim ide_refNFe      As String
Dim ide_tpImp       As String
Dim ide_tpEmis      As String
Dim ide_cDV         As String
Dim ide_tpAmb       As String
Dim ide_finNFe      As String
Dim ide_indFinal    As String
Dim ide_indPres     As String
Dim ide_procEmi     As String
Dim ide_verProc     As String
    'Emitente
Dim emit_CNPJ       As String
Dim emit_xNome      As String
Dim emit_xFant      As String
Dim emit_xLgr       As String
Dim emit_nro        As String
Dim emit_xCpl       As String
Dim emit_Bairro     As String
Dim emit_cMun       As String
Dim emit_xMun       As String
Dim emit_UF         As String
Dim emit_CEP        As String
Dim emit_cPais      As String
Dim emit_xPais      As String
Dim emit_fone       As String
Dim emit_IE         As String
Dim emit_IEST       As String
Dim emit_IM         As String
Dim emit_CNAE       As String
Dim emit_CRT        As String
    'Destinatario
Dim dest_idDest     As Integer
Dim dest_pessoa     As String 'Variavel particular para saber que tipo de Pessoa F/J
Dim dest_CNPJ       As String
Dim dest_xNome      As String
Dim dest_xFant      As String
Dim dest_xLgr       As String
Dim dest_nro        As String
Dim dest_xCpl       As String
Dim dest_Bairro     As String
Dim dest_cMun       As String
Dim dest_xMun       As String
Dim dest_UF         As String
Dim dest_CEP        As String
Dim dest_cPais      As String
Dim dest_xPais      As String
Dim dest_fone       As String
Dim dest_IE         As String
Dim dest_ISUF       As String
Dim dest_email      As String
Dim dest_indIEDest  As String
Dim infAdic_infCpl  As String

'Local de Entrega
Dim entr_CNPJ       As String
Dim entr_CPF        As String
Dim entr_xLgr       As String
Dim entr_nro        As String
Dim entr_xCpl       As String
Dim entr_xBairro    As String
Dim entr_cMun       As String
Dim entr_xMun       As String
Dim entr_UF         As String
    'Produtos
Dim aItem(1000)     As Variant
Dim aICMS(1000)     As Variant
Dim aIcmsDifal(1000)     As Variant
Dim aIPI(1000)      As Variant
Dim aPIS(1000)      As Variant
Dim aCOFINS(1000)   As Variant
Dim aEstoque(1000)  As Variant 'Variavel gerenciadora para estoque
Dim aComissao(1000) As Variant 'Variavel gerenciado Comissao por item
Dim cItens          As Integer 'Contador dos itens
    'Cobranca
Dim aCob(100)       As Variant
Dim cCob            As Integer 'Contador das cobrancas

    
    'Transporte
Dim transp_modFrete As String
Dim transp_Pessoa   As String
Dim transp_CNPJ     As String
Dim transp_xNome    As String
Dim transp_IE       As String
Dim transp_xEnder   As String
Dim transp_xMun     As String
Dim transp_UF       As String
Dim transp_qVol     As String
Dim transp_esp      As String
Dim transp_marca    As String
Dim transp_nVol     As String
Dim transp_pesoL    As String
Dim transp_pesoB    As String
Dim transp_VeicUF   As String
Dim transp_VeicPlaca As String
    'TOTAIS
Dim total_vBC       As String
Dim total_vICMS     As String
Dim total_vFCP      As String
Dim total_vBCST     As String
Dim total_vICMSST   As String
Dim total_vCredICMSSN     As String
Dim total_vProd     As String
Dim total_vFrete    As String
Dim total_vSeg      As String
Dim total_vDesc     As String
Dim total_vIPI      As String
Dim total_vPIS      As String
Dim total_vCOFINS   As String
Dim total_vOutro    As String
Dim total_vNF       As String

'22.12.17 - Totalizadores DIFAL
Dim total_vFCPUFDest    As String
Dim total_vICMSUFDest   As String
Dim total_vICMSUFRemet  As String
'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************
'01.08.2018 - Solucao encontrada para registrar o aproveitamento
'           de credito de ICMS CST 101
Dim pCredICMS   As String 'Armazena o valor do cred de icms
Private Function calcComissao(nItem As Integer) As String
    '29/10/2012 - Criado para desembutir o IPI da comissao
    Dim pIPI As String '% Comissao
    Dim iIPI As String 'Indice de calculo
    Dim vBCc As String 'Valor Base de Calculo comissao
    Dim vBT  As String 'Valor Bruto do item
    Dim pComV As String  '% comissao vendedor
    Dim vComi   As String
    
    pIPI = ChkVal(PgDadosNCM("NCM", CStr(aItem(nItem)(5)), "S").pIPI, 0, 2)
    
    '09.02.2017 - alterado para quando o item nao entra no total da nf
    If aItem(nItem)(19) = 1 Then
            pComV = ChkVal(PgDadosRhFuncionario(ger_Vendedor).Comissao, 0, 3)
        Else
            pComV = 0
    End If
    vBCc = ChkVal(CStr(aItem(nItem)(10)), 0, cDecMoeda)
    vBT = vBCc
    
    
    If PgDadosTpNotaFiscal(idTpNF).MovComissao = 1 Then
            If pIPI <> 0 And Trim(aIPI(nItem)(4)) = ChkVal(0, 0, cDecMoeda) Then
                    'IPI embutido
                    iIPI = ChkVal(Val(pIPI) / 100 + 1, 0, 2)
                    vBCc = Val(ChkVal(vBCc, 0, cDecMoeda)) / Val(ChkVal(iIPI, 0, 2))
                    vComi = Val(ChkVal(pComV, 0, 3)) * Val(ChkVal(vBCc, 0, cDecMoeda)) / 100
                    pComV = (Val(ChkVal(vComi, 0, cDecMoeda)) * 100) / Val(vBT)
                    
                Else
                    'IPI Destacado na NF
                    vComi = Val(pComV) * Val(vBCc) / 100
            End If
        Else
            vComi = 0
            pComV = 0
    End If
    
    
    vComi = ChkVal(vComi, 0, cDecMoeda)
    pComV = ChkVal(pComV, 0, 3)
    aComissao(cItens) = Array(pComV, vComi)
End Function

Private Function CalcFCP_Total(CampoArray As Integer) As String
    'Soma o Total do FCP do ICMS
    Dim Soma   As String
    Dim i      As Integer
    Soma = 0
    For i = 0 To cItens
        Soma = Val(ChkVal(Soma, 0, 2)) + Val(ChkVal(CStr(aICMS(i)(CampoArray)), 0, 2))
    Next
    CalcFCP_Total = ChkVal(Soma, 0, 2)
    
End Function

Private Function CalcICMS_Item() As Boolean


    'On Error GoTo TratarErroCalcICMS
    Dim i As Integer
    'Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST|vICMSST
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    For i = 0 To cItens
        Select Case aICMS(i)(1)
            Case "00", "900" 'Tributacao Integral
                Calculo_ICMS_CST_00 (i)
                Calculo_ICMS_DIFAL (i)
                
            Case "10"
                Calculo_ICMS_CST_10 (i)
            Case "20"
                Calculo_ICMS_CST_20 (i)
            Case "30"
                MsgBox "CST - 30 incompleto"
            Case "40"
                Calculo_ICMS_CST_40 (i)
            Case "41"
                Calculo_ICMS_CST_41 (i)
            Case "50"
                Calculo_ICMS_CST_50 (i)
            Case "51"
                Calculo_ICMS_CST_51 (i)
            Case "60" 'ICMS cobrado anteriormente por ST
                If Calculo_ICMS_CST_60(i) = False Then
                    CalcICMS_Item = False
                    Exit Function
                End If
            Case "70"
                MsgBox "CST - 70 incompleto"
            Case "90"
                MsgBox "CST - 90 incompleto"
            Case "101" 'Tributado pelo SN com permicao de credito (N10c)
                Calculo_ICMS_CSOSN_101 (i)
        End Select
    Next
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    
    CalcICMS_Item = True
    Exit Function
TratarErroCalcICMS:
    Resume Next
    
End Function

Private Sub CalcPIS_Item()


    On Error GoTo TratarErroCalcPISCOFINS
    Dim i As Integer
    'CST|vBC|pCOFINS|vCOFINS
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    For i = 0 To cItens
        Select Case aPIS(i)(0)
            Case "01"
                aPIS(i)(3) = (Val(ChkVal(CStr(aPIS(i)(1)), 0, cDecMoeda)) * Val(ChkVal(CStr(aPIS(i)(2)), 0, 3))) / 100
                aPIS(i)(3) = ChkVal(CStr(aPIS(i)(3)), 0, cDecMoeda)
             Case "02"
                aPIS(i)(3) = (Val(ChkVal(CStr(aPIS(i)(1)), 0, cDecMoeda)) * Val(ChkVal(CStr(aPIS(i)(2)), 0, 3))) / 100
                aPIS(i)(3) = ChkVal(CStr(aPIS(i)(3)), 0, cDecMoeda)
            Case "03"
                MsgBox "CST - 03 incompleto"
            Case "04"
                aPIS(i)(3) = ChkVal("0", 0, cDecMoeda)
            Case "06"
                aPIS(i)(3) = ChkVal("0", 0, cDecMoeda)
            Case "07"
                aPIS(i)(3) = ChkVal("0", 0, cDecMoeda)
            Case "08"
                aPIS(i)(3) = ChkVal("0", 0, cDecMoeda)
            Case "09"
                aPIS(i)(3) = ChkVal("0", 0, cDecMoeda)
            Case "99"
                aPIS(i)(3) = (Val(ChkVal(CStr(aPIS(i)(1)), 0, cDecMoeda)) * Val(ChkVal(CStr(aPIS(i)(2)), 0, 3))) / 100
                aPIS(i)(3) = ChkVal(CStr(aPIS(i)(3)), 0, cDecMoeda)
        End Select
    Next
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    
    
    Exit Sub
TratarErroCalcPISCOFINS:
    Resume Next
    
End Sub
Private Sub CalcCOFINS_Item()


    On Error GoTo TratarErroCalcCOFINS
    Dim i As Integer
    'CST|vBC|pCOFINS|vCOFINS
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    For i = 0 To cItens
        Select Case aCOFINS(i)(0)
            Case "01" 'Tributacao Integral
                aCOFINS(i)(3) = (Val(ChkVal(CStr(aCOFINS(i)(1)), 0, cDecMoeda)) * Val(ChkVal(CStr(aCOFINS(i)(2)), 0, 3))) / 100
                aCOFINS(i)(3) = ChkVal(CStr(aCOFINS(i)(3)), 0, cDecMoeda)
                
            Case "02"
                aCOFINS(i)(3) = (Val(ChkVal(CStr(aCOFINS(i)(1)), 0, cDecMoeda)) * Val(ChkVal(CStr(aCOFINS(i)(2)), 0, 3))) / 100
                aCOFINS(i)(3) = ChkVal(CStr(aCOFINS(i)(3)), 0, cDecMoeda)
            Case "03"
                'MsgBox "CST - 03 incompleto"
                aCOFINS(i)(3) = ChkVal("0", 0, cDecMoeda)
            Case "04"
                aCOFINS(i)(3) = ChkVal("0", 0, cDecMoeda)
            Case "06"
                aCOFINS(i)(3) = ChkVal("0", 0, cDecMoeda)
            Case "07"
                aCOFINS(i)(3) = ChkVal("0", 0, cDecMoeda)
            Case "08"
                aCOFINS(i)(3) = ChkVal("0", 0, cDecMoeda)
            Case "09"
                aCOFINS(i)(3) = ChkVal("0", 0, cDecMoeda)
            Case "99"
                'MsgBox "CST - 99 inconpleto"
                aCOFINS(i)(3) = ChkVal("0", 0, cDecMoeda)
        End Select
    Next
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    '**********************************************************************************************************
    
    
    Exit Sub
TratarErroCalcCOFINS:
    Resume Next
    
End Sub
Private Sub CalcDIFAL_Total()
    'Calcula o valor total do DIFAL
    
    Dim i      As Integer
    'Zerar os itens
    total_vFCPUFDest = "0"
    total_vICMSUFDest = "0"
    total_vICMSUFRemet = "0"
    
    
    For i = 0 To cItens
        'Soma = Val(ChkVal(Soma, 0, 2)) + Val(ChkVal(CStr(aICMS(i)(CampoArray)), 0, 2))
        'Soma os totalizadores
        total_vFCPUFDest = Val(ChkVal(total_vFCPUFDest, 0, cDecMoeda)) + Val(ChkVal(CStr(aIcmsDifal(i)(5)), 0, cDecMoeda))
        total_vFCPUFDest = ChkVal(total_vFCPUFDest, 0, cDecMoeda)
        
        total_vICMSUFDest = Val(ChkVal(total_vICMSUFDest, 0, cDecMoeda)) + Val(ChkVal(CStr(aIcmsDifal(i)(6)), 0, cDecMoeda))
        total_vICMSUFDest = ChkVal(total_vICMSUFDest, 0, cDecMoeda)
        
        total_vICMSUFRemet = Val(ChkVal(total_vICMSUFRemet, 0, cDecMoeda)) + Val(ChkVal(CStr(aIcmsDifal(i)(7)), 0, cDecMoeda))
        total_vICMSUFRemet = ChkVal(total_vICMSUFRemet, 0, cDecMoeda)

    Next
    'CalcICMS_Total = Val(ChkVal(Soma, 0, 2))
    
    
    '-----------------------------
'    aIcmsDifal(Item)(0) = ChkVal(vBCUFDest, 0, cDecMoeda)
'    aIcmsDifal(Item)(1) = ChkVal(pFCPUFDest, 0, cDecMoeda)
'    aIcmsDifal(Item)(2) = ChkVal(pICMSUFDest, 0, cDecMoeda)
'    aIcmsDifal(Item)(3) = ChkVal(pICMSInter, 0, cDecMoeda)
'    aIcmsDifal(Item)(4) = ChkVal(pICMSInterPart, 0, cDecMoeda)
'    aIcmsDifal(Item)(5) = ChkVal(vFCPUFDest, 0, cDecMoeda)
'    aIcmsDifal(Item)(6) = ChkVal(vICMSUFDest, 0, cDecMoeda)
'    aIcmsDifal(Item)(7) = ChkVal(vICMSUFRemet, 0, cDecMoeda)
'
    
    
    
    
    
End Sub
Private Function CalcICMS_Total(CampoArray As Integer) As String
    'Soma o Total do ICMS
    Dim Soma   As String
    Dim i      As Integer
    Soma = 0
    For i = 0 To cItens
        Soma = Val(ChkVal(Soma, 0, 2)) + Val(ChkVal(CStr(aICMS(i)(CampoArray)), 0, 2))
    Next
    CalcICMS_Total = Val(ChkVal(Soma, 0, 2))
End Function

Private Function CalcIPI_Total(CampoArray As Integer) As String
    Dim Soma   As String
    Dim i      As Integer
    Soma = 0
    For i = 0 To cItens
        If aItem(i)(19) = 1 Then
            Soma = Val(ChkVal(Soma, 0, 2)) + Val(ChkVal(CStr(aIPI(i)(CampoArray)), 0, 2))
        End If
    Next
    CalcIPI_Total = Val(ChkVal(Soma, 0, 2))
End Function
Private Function CalcPIS_Total(CampoArray As Integer) As String
    Dim Soma   As String
    Dim i      As Integer
    
    For i = 0 To cItens
        Soma = Val(ChkVal(Soma, 0, 2)) + Val(ChkVal(CStr(aPIS(i)(CampoArray)), 0, 2))
    Next
    CalcPIS_Total = Val(ChkVal(Soma, 0, 2))
End Function
Private Function CalcCOFINS_Total(CampoArray As Integer) As String
    Dim Soma   As String
    Dim i      As Integer
    
    For i = 0 To cItens
        Soma = Val(ChkVal(Soma, 0, 2)) + Val(ChkVal(CStr(aCOFINS(i)(CampoArray)), 0, 2))
    Next
    CalcCOFINS_Total = Val(ChkVal(Soma, 0, 2))
End Function



Private Sub cobrAcertarParcelas()
    'Acerta os centavos sobrantes nas parcelas
    'Observar a variavel total_vNF caso mudar a chamada da funcao
    Dim i           As Integer
    Dim smCobTot    As String
    smCobTot = 0
    
    For i = 0 To cCob
    
        smCobTot = Val(ChkVal(CStr(aCob(i)(6)), 0, cDecMoeda)) + Val(ChkVal(smCobTot, 0, cDecMoeda))
    Next
    If ChkVal(smCobTot, 0, cDecMoeda) > ChkVal(total_vNF, 0, cDecMoeda) Then
            aCob(cCob)(6) = Val(ChkVal(CStr(aCob(cCob)(6)), 0, cDecMoeda)) - (Val(ChkVal(smCobTot, 0, cDecMoeda)) - Val(ChkVal(total_vNF, 0, cDecMoeda)))
            aCob(cCob)(6) = ChkVal(CStr(aCob(cCob)(6)), 0, 2)
        ElseIf ChkVal(smCobTot, 0, cDecMoeda) < ChkVal(total_vNF, 0, cDecMoeda) Then
            aCob(0)(6) = Val(ChkVal(CStr(aCob(0)(6)), 0, cDecMoeda)) + (Val(ChkVal(total_vNF, 0, cDecMoeda)) - Val(ChkVal(smCobTot, 0, cDecMoeda)))
            aCob(0)(6) = ChkVal(CStr(aCob(0)(6)), 0, 2)
    End If
    
        
End Sub

Private Sub DistribuirValorFrete()
    'Apenas 1 item na nota
'    If cItens = 0 Then
'        Trim(aItem(cItens)(15)) = ChkVal(total_vFrete, 0, cDecMoeda)
'        Exit Sub
'    End If
'    'Mais de 1 item
    aItem(0)(15) = ChkVal(total_vFrete, 0, cDecMoeda)
End Sub

Private Sub msgValid(sTexto As String)
    txtMsg.Text = txtMsg.Text & IIf(Trim(txtMsg.Text) = "", "", vbCrLf) & sTexto
End Sub

Private Function pgValorDuplicataPV(vCalcSis As Variant, nPedido As Integer, nDuplicata As Integer) As Variant
    Dim Rst  As Recordset
    Dim sSQL As String
    
    sSQL = "SELECT * FROM FaturamentoPVCobranca WHERE IdPV=" & nPedido & " AND Parcela = " & nDuplicata
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            pgValorDuplicataPV = ChkVal(CStr(vCalcSis), 0, cDecMoeda)
        Else
            pgValorDuplicataPV = ChkVal(Rst.Fields("Valor"), 0, cDecMoeda)
    End If
End Function

Private Sub Validar()
    If MontarVariaveis = False Then
            tbMenu.Buttons(3).Enabled = False
            'txtChaveAcesso.ForeColor = vbRed
            'txtChaveAcesso.Text = "<<< NF-e NÃO AUTORIZADA >>>"
            msgValid "<<<<<<<<<<<<<<< NF-e NÃO AUTORIZADA >>>>>>>>>>>>>>>"
        Else
            tbMenu.Buttons(3).Enabled = True
            
            
            If ValidarVariaveis = False Then
                tbMenu.Buttons(3).Enabled = False
                'txtChaveAcesso.ForeColor = vbRed
                'txtChaveAcesso.Text = "<<< NF-e NÃO AUTORIZADA >>>"
                msgValid "<<<<<<<<<<<<<<< NF-e NÃO AUTORIZADA >>>>>>>>>>>>>>>"
                Exit Sub
            End If
                    
            'txtChaveAcesso.ForeColor = vbBlue
            'txtChaveAcesso.Text = Format(strChaveAcesso, "@@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@")
            msgValid "###################  NF-e AUTORIZADA  ###################"
            'MsgBox "Nota Fiscal Eletronica valida para envio a Receita Federal.", vbInformation, "Aviso"
    End If
End Sub

Private Sub btoPesqPedido_Click()
    PesquisarPedido
End Sub
Private Sub PesquisarPedido(Optional nmPedido As Integer)
    Dim Rst         As Recordset 'Armazena dados do cab do pedido
    Dim sSQL        As String
    Dim idTMP       As String
    Dim CliID       As Integer
    
    DesMontarVariaveisFiscal
    
    If nmPedido = 0 Then
        idTMP = formBuscar.IniciarBusca("FaturamentoPV")
        If idTMP = 0 Then
                Exit Sub
            Else
                nmPedido = idTMP
        End If
    End If
    sSQL = "SELECT * FROM FaturamentoPV WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & nmPedido
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            idPedido = 0
            CliID = 0
            MsgBox "Nenhum registro encontrado!", vbCritical, "Aviso"
            
        Else
            Rst.MoveFirst
            idPedido = Rst.Fields("ID")
            txtPedido.Text = idPedido
            CliID = IIf(IsNull(Rst.Fields("idCliente")), 0, Rst.Fields("idCliente"))
            
            
            If PgDadosConfig.Ambiente = 2 Then
                    'Homologacao
                    lblNome.Caption = "Dest.: " & "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
                Else
                    'Producao
                    lblNome.Caption = "Dest.: " & Rst.Fields("Cliente")
            End If

            
            
            lblEmissao.Caption = "Emissão: " & Rst.Fields("Emissao")
            lblValor.Caption = "Valor Total: " & ConvMoeda(Rst.Fields("VlTotalPV"))
            
            txtObs.Text = "[N/Ref.: " & nmPedido & "] / " & _
                          IIf(IsNull(Rst.Fields("RefCliente")), "", "[S/Ref.: " & Trim(Rst.Fields("RefCliente")) & "] / ") & _
                          IIf(PgDadosTpNotaFiscal(idTpNF).ImpInfCompl = 0, "", Trim(PgDadosTpNotaFiscal(idTpNF).Obs) & " / ") & _
                          IIf(Trim(PgDadosCliente(CliID).ObsCobNfe) = "", "", Trim(PgDadosCliente(CliID).ObsCobNfe) & " / ") & _
                          Trim(Rst.Fields("Obs"))
    End If
    Rst.Close
End Sub

Private Function PesquisarNFe(Optional nmNFe As String) As Boolean
    Dim Rst         As Recordset 'Armazena dados do cab do NFe
    Dim sSQL        As String
    Dim idNFe       As Integer
    
   
    If Trim(nmNFe) = "" Then
            idNFe = formBuscar.IniciarBusca("FaturamentoNFe")
            If idNFe = 0 Then Exit Function
            sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & idNFe
        Else
            sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & nmNFe & "'"
    End If
    
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'MsgBox "Nenhum Registro encontrado."
            PesquisarNFe = False
        Else
            PesquisarNFe = True
            Rst.MoveFirst
            txtrefNFe.Text = Rst.Fields("IdNFe")
    End If
    Rst.Close
End Function
Private Sub PesquisarTpNF(Optional idNF As Integer)
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim idTMP   As String
    DesMontarVariaveisFiscal
    idPedido = 0
    If idNF = 0 Then
        idTMP = formBuscar.IniciarBusca("FaturamentoTipoNotaFiscal")
        If idTMP = 0 Then
                Exit Sub
            Else
                idNF = idTMP
        End If
    End If
    
    sSQL = "SELECT * FROM FaturamentoTipoNotaFiscal " & _
           "WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & idNF
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            LimpForm
            DesMontarVariaveisFiscal
            idPedido = 0
            idTpNF = 0
            HDMenu Me, True
            HDForm Me, False
            txtTpNF.Enabled = True
            btoPesqTipoNF.Enabled = True
            tbMenu.Buttons(3).Enabled = False

            MsgBox "Nenhum Registro encontrado."
        Else
            HDForm Me, True
            Rst.MoveFirst
            idTpNF = Rst.Fields("ID")
            txtTpNF.Text = idTpNF
            lblTipoNF.Caption = Rst.Fields("Descricao")
            '########################################################################################################
            If PgDadosConfig.BloqueionNFManual <> 0 Then
                    txtNumNota.Enabled = False
                    txtNumNota.Text = "" 'PgPxNumNota(Rst.Fields("NumInicial"))
                    ide_nNF = ""
                    ide_cNF = ""
                Else
                    txtNumNota.Enabled = True
                    txtNumNota.Text = PgPxNumNota
            End If
            
            dtpSaida.Enabled = IIf(Rst.Fields("ImprDataSaida") = 1, True, False)
            'txtObs.Text = IIf(IsNull(Rst.Fields("Obs")), "", Trim(Rst.Fields("Obs")))
            txtObs.Text = IIf(PgDadosTpNotaFiscal(idTpNF).ImpInfCompl = 0, "", Trim(PgDadosTpNotaFiscal(idTpNF).Obs))

            txtrefNFe.Enabled = IIf(cNull(Rst.Fields("ChaveAcessoRef")) = "1", True, False)
            
            txtPedido.SetFocus
    End If
    Rst.Close
    
    
End Sub
Private Function PgPxNumNota() As String
    'Devolve o mun da prox NF com 9 posicoes
'###################################################################################################
'# Caso o sistema de nNF seja automatico colocar um loop para caso a tab esteja bloq so sair apos  #
'# o desbloqueio ou X segundos                                                                     #
'###################################################################################################
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim numIni  As String
    numIni = PgDadosTpNotaFiscal(idTpNF).NumInicial
    ide_serie = PgDadosTpNotaFiscal(idTpNF).Serie
    'sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_Serie = " & PgDadosTpNotaFiscal(idTpNF).Serie & _
           " ORDER BY ide_nNF"
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_serie = '" & ide_serie & "'" & _
           " ORDER BY ide_nNF"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            PgPxNumNota = Left(String(9, "0"), 9 - Len(numIni)) & numIni
        Else
            Rst.MoveLast
            numIni = Rst.Fields("ide_nNf")
            numIni = Val(numIni) + 1
            PgPxNumNota = Left(String(9, "0"), 9 - Len(numIni)) & numIni
    End If
    Rst.Close
End Function


Private Sub btoPesqTipoNF_Click()
    PesquisarTpNF
End Sub


Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    LimpForm
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    DesMontarVariaveisFiscal
    HDMenu Me, True
    HDForm Me, False
    txtTpNF.Enabled = True
    btoPesqTipoNF.Enabled = True
    tbMenu.Buttons(3).Enabled = False
    idPedido = 0
    idTpNF = 0
    'txtTpNF.SetFocus
End Sub
Private Sub DesMontarVariaveisFiscal()
    'idPedido = 0
    'idTpNF = 0
    tbMenu.Buttons(3).Enabled = False
    txtMsg.Text = ""
    'txtrefNFe.Text = ""
    'Me.Height = 5895
End Sub

Private Sub EnviarNFe()
    Dim nArqNfe As String
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    
    If MsgBox("Gravar Nota Fiscal na Base de Dados?", vbQuestion + vbYesNo, "Gravar de NF-e na Base de Daods") = vbYes Then
    
                If grvRegistro = False Then
                    MsgBox "Erro ao gravar NFe", vbInformation, "Aviso"
                    Exit Sub
                End If
                'Seleciona o tipo de NFe sera exportada
                If VersaoNFe = "3.10" Then
                        nArqNfe = Exportar_NFe_v310_TXT(strChaveAcesso)
                    ElseIf (VersaoNFe = "4.00") Then
                        nArqNfe = Exportar_NFe_v400_TXT(strChaveAcesso)
                    Else
                        MsgBox "Versão de nfe não encontrada", vbInformation, App.EXEName
                End If
                
                
                
                If Trim(nArqNfe) <> "" Then
                        MsgBox "NF-e gerada com sucesso.", vbInformation
                        'Checa se envia a NFe pafa SEFAZ
                        If PgDadosTpNotaFiscal(idTpNF).EnvioRF <> "0" Then
                            'Checa se apresenta o Preview ante de enviar para SEFAZ
                            If PgDadosConfig.DANFEPreview <> "0" Then
                                    ImprimirDANFE2 strChaveAcesso, 1
                                    If MsgBox("O processo de solicitação de autorização junto a SEFAZ não lhe " & vbCrLf & _
                                          "permitirá efetuar modificações nesta Nota Fiscal." & vbCrLf & vbCrLf & _
                                          "Deseja realmente enviar esta solicitação?", vbQuestion + vbYesNo, "Envio de NF-e para SEFAZ") = vbYes Then
                                                    
                                                    MoverPastaEnvio_UniNFe (nArqNfe)
                                    End If
                                Else
                                    MoverPastaEnvio_UniNFe (nArqNfe)
                            End If
                        End If
                        LimpForm
                        idPedido = 0
                        idTpNF = 0
                    Else
                        MsgBox "Erro ao gerar NF-e (TXT).", vbInformation
                End If
                
    End If
    tbMenu.Buttons(3).Enabled = False
End Sub
Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim nArqNfe As String
    
    DesMontarVariaveisFiscal
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Vizualizar pré-venda"
            If Trim(txtPedido.Text) <> "" Then
                ImpPV (txtPedido.Text)
            End If
        Case "Editar pré-venda"
            If idPedido = 0 Then
                MsgBox "Selecione uma Pré-venda!", vbInformation, "Aviso"
                Exit Sub
            End If
            formFaturamentoPV.PesquisarRegistro (idPedido)
            LimpForm
            formFaturamentoPV.Show
        Case "Envia NF-e"
            EnviarNFe
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select
End Sub




Private Sub txtNumNota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtNumNota_LostFocus()
    txtNumNota.Text = Left(String(9, "0"), 9 - Len(Trim(txtNumNota.Text))) & Trim(txtNumNota.Text)
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = 13, 0, KeyAscii)
End Sub

Private Sub txtPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarPedido
    End If
End Sub
Private Sub LimpForm()
    LimpaFormulario Me
    dtpEmissao.Value = Date
    dtpSaida.Value = Date
    lblTipoNF.Caption = ""
    
    lblNome.Caption = ""
    lblEmissao.Caption = ""
    lblValor.Caption = ""
    idPedido = 0
    idTpNF = 0
    txtrefNFe.Text = ""
    msgRedICMS = ""
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        PesquisarPedido (txtPedido.Text)
        If idPedido <> 0 Then
            Validar
        End If
    End If
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub



Private Sub txtrefNFe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then PesquisarNFe
        
End Sub

Private Sub txtrefNFe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtTpNF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarTpNF
    End If

End Sub

Private Sub txtTpNF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        If Trim(txtTpNF.Text) = "" Then Exit Sub
        
        PesquisarTpNF (txtTpNF.Text)
        'txtPedido.SetFocus
    End If
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub
Private Function MontarVariaveis() As Boolean
    'On Error GoTo TrtMVErro
    Dim Rst1        As Recordset 'Armazena dados do cab do pedido
    Dim Rst2        As Recordset 'Armazena os itens
    Dim Rst3        As Recordset 'Busca as parcelas da cobranca
    Dim sSQL        As String
    Dim idTMP       As String
    Dim idProduto   As Integer 'Armazena temporariamente o id do produto para montagem da array
    Dim cont        As Integer
    Dim cProd       As String 'Armazena o codigo do produto que saira na NFe
    
   
   
    pCredICMS = ""
    
    MontarVariaveis = False
    
    If ChkPvTemNFe(idPedido) = True Then
        MsgBox "Pré-venda já possui Nota Fiscal vinculada.", vbInformation, "Aviso"
        MontarVariaveis = False
        Exit Function
    End If
    
    txtMsg.Text = ""
    If idPedido = 0 Then
        MsgBox "Favor selecionar um Pedido..."
        MontarVariaveis = False
        Exit Function
    End If
      If idTpNF = 0 Then
        MsgBox "Favor selecionar um Tipo de Nota Fiscal...", vbInformation, App.EXEName
        MontarVariaveis = False
        Exit Function
    End If
    
    'cabecario do Pedido (ide)
    
    ide_cUF = pgDadosICMS(PgDadosEmpresa(ID_Empresa).uf, 0).codUF
    ide_natOp = rc(PgDadosTpNotaFiscal(idTpNF).Natureza)
    
    'v.3.10 - Indica Forma de pagamento: 0 - avista, 1 - prazo , 2 -  outros
    'v.4.00 - Indica Forma de pagamento: 0 - avista, 1 - prazo
    ide_indPag = "1"
    
    ide_mod = PgDadosTpNotaFiscal(idTpNF).Modelo
    'ide_serie = PgDadosTpNotaFiscal(idTpNF).Serie
    
'###################################################################################################
'# Colocar um IF usando criterios de informar num da nf sim ou nao em configuracoes                #
'# e passa a função de gerar nNF para gravar com o mesmo IF, assim o sistema pegara o num da NF    #
'# aqui ou na hora de gravar                                                                       #
'###################################################################################################
    
    'ide_nNF = Left(String(9, "0"), 9 - Len(Trim(txtNumNota.Text))) & Trim(txtNumNota.Text)
    'If NumNotaFiscalExiste(ide_nNF) = True Then 'Verifica se o num da nota ja esta cadastrado
    '    MontarVariaveis = False
    '    MsgBox "Numero de Nota Fiscal ja cadastrado.", vbInformation, "Aviso"
    '    Exit Function
    'End If
    'ide_cNF = Format(Now(), "DDHHMMSS"): ide_cNF = Mid(String(8, "0"), 1, 8 - Len(Trim(ide_cNF))) & Trim(ide_cNF)
    
    ide_dEmi = dtpEmissao.Value
    'ide_hEmi = Format(Now(), "hh:mm:ss")
    ide_dSaiEnt = IIf(PgDadosTpNotaFiscal(idTpNF).ImpDtSaida = 0, "", dtpSaida.Value)
    ide_hSaiEnt = IIf(PgDadosTpNotaFiscal(idTpNF).ImpDtSaida = 0, "", Format(Time, "HH:MM:SS"))
    ide_tpNF = PgDadosTpNotaFiscal(idTpNF).TipoNota
    ide_cMunFG = PgDadosMunicipio(PgDadosEmpresa(ID_Empresa).uf, PgDadosEmpresa(ID_Empresa).Mun).codMun
    ide_indFinal = "0"
    ide_indPres = "1"
    
    
    
    
    ide_tpImp = "1" 'Colocar no formFaturamentoTipoNotaFiscal essa opcao e configurar o Unidanfe automaticamente"
    ide_tpEmis = PgDadosConfig.TpEmissao
    'ide_cDV = Right(Id, 1) - Transferi para depois de validar a nf poi o cDV estava ficando em branco
    
    'Ambiente
    '1 - Producao
    '2 - Homologacao
    ide_tpAmb = PgDadosConfig.Ambiente
        
    ide_finNFe = PgDadosTpNotaFiscal(idTpNF).Finalidade
    ide_procEmi = "0"
    ide_verProc = sVersao & "." & cVersao
    sSQL = "SELECT * FROM FaturamentoPV WHERE Id_Empresa = " & ID_Empresa & " AND ID = " & idPedido
    
    Set Rst1 = RegistroBuscar(sSQL)
    If Rst1.BOF And Rst1.EOF Then
            MsgBox "Nenhum Registro encontrado."
        Else
            Rst1.MoveFirst
            
            
            
            idCliente = Rst1.Fields("IdCliente")
            If idCliente = 0 Then
                MsgBox "Cliente não cadastrado. Favor cadastrar e informar na pre-venda o ID do Cliente.", vbInformation, "Aviso"
                MontarVariaveis = False
                Exit Function
            End If
            
            '############################################################################
            '### Checa se o cliente possui e-mail de envio
            '############################################################################
            If Trim(PgDadosCliente(idCliente).emailnfe) = "" Then
                msgValid "E-mail para envio do XML não encontrado."
                If MsgBox("E-mail para envio do XML não encontrado. Deseja continuar assim mesmo?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                    MontarVariaveis = False
                    Exit Function
                End If
            End If
            total_vProd = IIf(IsNull(Rst1.Fields("VlMercadoria")), "0", Rst1.Fields("VlMercadoria"))
            total_vFrete = IIf(IsNull(Rst1.Fields("Frete")), "0", Rst1.Fields("Frete"))
            total_vSeg = IIf(IsNull(Rst1.Fields("Seguro")), "0", Rst1.Fields("Seguro"))
            total_vDesc = IIf(IsNull(Rst1.Fields("Desconto")), "0", Rst1.Fields("Desconto"))
            total_vOutro = IIf(IsNull(Rst1.Fields("Outros")), "0", Rst1.Fields("Outros"))
            formPag = Rst1.Fields("FormaPagamento")
            SRef = IIf(IsNull(Rst1.Fields("RefCliente")), "", Rst1.Fields("RefCliente"))
            TranspRetEnt = IIf(IsNull(Rst1.Fields("transp_RetEnt")), 0, Rst1.Fields("transp_RetEnt"))
            IdTransp = IIf(IsNull(Rst1.Fields("Transportadora")), 0, Rst1.Fields("Transportadora"))
            'Vendedor = Rst1.fields("Vendedor")
            FreteConta = Rst1.Fields("FreteConta")

            qVol = IIf(IsNull(Rst1.Fields("transp_qVol")), "", Rst1.Fields("transp_qVol"))
            Esp = IIf(IsNull(Rst1.Fields("transp_Esp")), "", Rst1.Fields("transp_Esp"))
            Marca = IIf(IsNull(Rst1.Fields("transp_Marca")), "", Rst1.Fields("transp_Marca"))
            nVol = IIf(IsNull(Rst1.Fields("transp_nVol")), "", Rst1.Fields("transp_nVol"))
            PesoB = IIf(IsNull(Rst1.Fields("transp_PesoB")), "", Rst1.Fields("transp_PesoB"))
            PesoL = IIf(IsNull(Rst1.Fields("transp_PesoL")), "", Rst1.Fields("transp_PesoL"))
            
            bcICMS = IIf(IsNull(Rst1.Fields("bcICMS")), 0, Rst1.Fields("bcICMS"))
            '03.03.19 - Campo incluso para mudar para consumidor final
            '0 - Normal
            '1 - 1 Consumidor final
            If bcICMS = 1 Then ide_indFinal = "1"
            
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'Emitente
    emit_CNPJ = RS(PgDadosEmpresa(ID_Empresa).CNPJ)
    emit_xNome = rc(PgDadosEmpresa(ID_Empresa).Nome)
    emit_xFant = rc(PgDadosEmpresa(ID_Empresa).Fant)
    emit_xLgr = rc(PgDadosEmpresa(ID_Empresa).Lgr)
    emit_nro = rc(PgDadosEmpresa(ID_Empresa).Nro)
    emit_xCpl = rc(PgDadosEmpresa(ID_Empresa).Cpl)
    emit_Bairro = rc(PgDadosEmpresa(ID_Empresa).Bairro)
    emit_cMun = PgDadosMunicipio(PgDadosEmpresa(ID_Empresa).uf, PgDadosEmpresa(ID_Empresa).Mun).codMun
    emit_xMun = rc(PgDadosEmpresa(ID_Empresa).Mun)
    emit_UF = PgDadosEmpresa(ID_Empresa).uf
    emit_CEP = PgDadosEmpresa(ID_Empresa).CEP
    emit_cPais = pgIdPais(PgDadosEmpresa(ID_Empresa).pais)
    emit_xPais = rc(PgDadosEmpresa(ID_Empresa).pais)
    emit_fone = RS(PgDadosEmpresa(ID_Empresa).Fone)
    emit_IE = RS(PgDadosEmpresa(ID_Empresa).IE)
    emit_IEST = RS(PgDadosEmpresa(ID_Empresa).iest)
    emit_IM = RS(PgDadosEmpresa(ID_Empresa).im)
    emit_CNAE = RS(PgDadosEmpresa(ID_Empresa).cnae)
    emit_CRT = PgDadosEmpresa(ID_Empresa).RegimeTrib
    
    
    'Destinatario
    dest_idDest = idCliente
    dest_pessoa = PgDadosCliente(idCliente).Pessoa 'NAO PODE SER DEIXADO EM  BRANCO
    dest_CNPJ = RS(PgDadosCliente(idCliente).Doc)
    
    If ide_tpAmb = 2 Then
            'Homologacao
            dest_xNome = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
        Else
            'Producao
            dest_xNome = rc(PgDadosCliente(idCliente).Nome)
    End If
    
    dest_xFant = rc(PgDadosCliente(idCliente).Fant)
    dest_xLgr = rc(PgDadosCliente(idCliente).Lgr)
    dest_nro = rc(PgDadosCliente(idCliente).Nro)
    dest_xCpl = rc(PgDadosCliente(idCliente).Cpl)
    dest_Bairro = rc(PgDadosCliente(idCliente).Bairro)
    dest_cMun = PgDadosMunicipio(PgDadosCliente(idCliente).uf, PgDadosCliente(idCliente).Mun).codMun
    dest_xMun = rc(PgDadosCliente(idCliente).Mun)
    dest_UF = PgDadosCliente(idCliente).uf
    dest_CEP = RS(PgDadosCliente(idCliente).CEP)
    dest_cPais = emit_cPais
    dest_xPais = rc(emit_xPais)
    dest_fone = RS(PgDadosCliente(idCliente).Fone)
    dest_IE = RS(PgDadosCliente(idCliente).IE)
    dest_ISUF = RS(PgDadosCliente(idCliente).Suframa)
    dest_email = rc(PgDadosCliente(idCliente).Mail)
    'Indicador da IE do Destinatario
    '1 - Contribuinte de ICMS
    '2 - Contribuinte isento de cadasto no ICMS
    '9 - Não contribuinte que pode ou nao ter IE
    dest_indIEDest = IIf(Len(Trim(dest_IE)) = 0, 2, 1)
    '03.03.19 - Altera caso a prevenda venha com instrucao de consumidor final
    'e o cliente nao tenha IE
    If ide_indFinal = "1" And Len(Trim(dest_IE)) = 0 Then
        dest_indIEDest = "9"
    End If
    infAdic_infCpl = rc(Trim(txtObs.Text))
    
    'Informa se a operacao e 1 interna ou 2 interestadual
    ide_idDest = IIf(emit_UF = dest_UF, "1", "2")
    
    'Vendedor
    '############################################################
    '### 26/10/2012
    '### Funcao alterada para que o sistema respeite
    '### o vendedor do cadastro do cliente.
    '### Caso o cliente seja Consumidor(cnpj: 99999999999999)
    '### O sistema pegara o vendedor da Pre-venda
    '###
    '### Original: ger_Vendedor = Rst1.fields("Vendedor")
    '############################################################
    If dest_CNPJ = String(14, "9") Then
            'Cliente CONSUMIDOR: pega vendedor PV
            ger_Vendedor = Rst1.Fields("Vendedor")
        Else
            'Cliente CADASTRADO: pega vendedor do cadastro
            ger_Vendedor = IIf(Trim(PgDadosCliente(dest_idDest).Vendedor) <> 0, Trim(PgDadosCliente(dest_idDest).Vendedor), Rst1.Fields("Vendedor"))
    End If
            
            
    ger_idPV = Rst1.Fields("id")
    
    'Local de Entrega
    If Trim(PgDadosCliente(idCliente).entrega) = "1" Then
        entr_CNPJ = PgDadosCliente(idCliente).entregaDoc
        entr_xLgr = PgDadosCliente(idCliente).entregalgr
        entr_nro = PgDadosCliente(idCliente).entreganro
        entr_xCpl = PgDadosCliente(idCliente).entregacpl
        entr_xBairro = PgDadosCliente(idCliente).entregabairro
        entr_cMun = PgDadosMunicipio(PgDadosCliente(idCliente).entregauf, PgDadosCliente(idCliente).entregamun).codMun   'PgDadosCliente(idCliente).entregamun
        entr_xMun = PgDadosCliente(idCliente).entregamun
        entr_UF = PgDadosCliente(idCliente).entregauf
    End If

            
            
            'Pegar os ITENS do pedido
            sSQL = "SELECT * FROM FaturamentoPVItens WHERE id_empresa = " & ID_Empresa & " AND idPV = " & Rst1.Fields("Id")
            Set Rst2 = RegistroBuscar(sSQL)
            If Rst2.BOF And Rst2.EOF Then
                    MsgBox "Erro ao localizar registro de Itens"
                Else
                    Rst2.MoveFirst
                    cItens = 0
                    Do Until Rst2.EOF
                        idProduto = IIf(IsNull(Rst2.Fields("idProduto")), 0, Rst2.Fields("idProduto"))
                        If idProduto = 0 Then
                            MsgBox "Item " & cItens + 1 & " da Pré-Venda sem cadastro no estoque! Favor verificar.", vbInformation, "Aviso"
                            MontarVariaveis = False
                            Exit Function
                        End If
                                        
                                         'det_indTot = 0 = O valor do item compoe a NF / 1  = O valor do item nao compoe a NF
                        If PgDadosConfig.CodProdImpresso = 1 Then '1 - codigo Interno
                                cProd = idProduto
                            Else ' 2 - Referencia
                                If IsNull(Rst2.Fields("Referencia")) = True Or Trim(Rst2.Fields("Referencia")) = "" Then
                                        MsgBox "Produto sem Referencia! Favor checar nos dados do produto.", vbInformation, "Aviso"
                                        MontarVariaveis = False
                                        Exit Function
                                    Else
                                        cProd = rc(Trim(Rst2.Fields("Referencia")))
                                End If
                        End If
                        'cProd = rc(IIf(IsNull(Rst2.Fields("Referencia")), idProduto, Rst2.Fields("Referencia")))
                        
                        '* 12.12.2012
                        '* Checa se o produto vendido é do mesmo deposito corrente
                        '*
                        If pgDadosEstoqueProduto(idProduto).IdDeposito <> ID_Deposito Then
                            msgValid "- Erro no DEPÓSITO do produto id: " & idProduto
                            MontarVariaveis = False
                            Exit Function
                        End If
                        
                        '*
                        '*
                        
                        
                        '### 06/12/2011
                        '### CFOP do produto respeita o que vem da PV ignorando o que consta na base de estoque
                                        'id_intProd|det_cProd|det_cEAN|det_xProd|EXTIPI|det_NCM|det_CFOP|det_uCom|
                                        'det_qCom|det_vUnCom|det_vProd|det_cEANTrib|det_uTrib|
                                        'det_qTrib|det_vUnTrib|
                                        'det_vFrete|det_vSeg|det_vDesc|det_vOutro|
                                        'det_indTot|xPed|nItemPed|ComplDescrNFe|
                                        'CodCEST
                        aItem(cItens) = Array(idProduto, _
                                            CStr(cProd), _
                                            "SEM GTIN", _
                                            CStr(rc(Rst2.Fields("Descricao"))), _
                                            "", _
                                            IIf(IsNull(Rst2.Fields("NCM")), pgDadosEstoqueProduto(idProduto).NCM, Trim(Rst2.Fields("NCM"))), _
                                            PgDadosCFOP(idTpNF, IIf(cNull(Rst2.Fields("CST")) = "", pgDadosEstoqueProduto(idProduto).ICMSCST, Rst2.Fields("CST")), PgDadosCliente(idCliente).uf).CFOP, _
                                            CStr(IIf(IsNull(Rst2.Fields("Unidade")), "", Rst2.Fields("Unidade"))), _
                                            CStr(ChkVal(Rst2.Fields("Quantidade"), 0, 4)), _
                                            CStr(ChkVal(Rst2.Fields("ValorUnitario"), 0, 10)), _
                                            CStr(ChkVal(IIf(IsNull(Rst2.Fields("vlItem")), 0, Rst2.Fields("vlItem")), 0, cDecMoeda)), _
                                            "SEM GTIN", _
                                            CStr(IIf(IsNull(Rst2.Fields("Unidade")), "", Rst2.Fields("Unidade"))), _
                                            CStr(ChkVal(Rst2.Fields("Quantidade"), 0, 4)), _
                                            CStr(ChkVal(Rst2.Fields("ValorUnitario"), 0, 10)), _
                                            "0.00", "", _
                                            IIf(Rst2.Fields("DescItem") = 0, "", CStr(ChkVal(Rst2.Fields("DescItem"), 0, 2))), _
                                            "", CStr(IIf(IsNull(Rst2.Fields("indTot")), "1", Rst2.Fields("indTot"))), _
                                            CStr(IIf(IsNull(Rst2.Fields("nPedido")), "", Rst2.Fields("nPedido"))), _
                                            CStr(IIf(IsNull(Rst2.Fields("iPedido")), "", Rst2.Fields("iPedido"))), _
                                            CStr(IIf(IsNull(Rst2.Fields("ComplDescricaoNFe")), "", Rst2.Fields("ComplDescricaoNFe"))), _
                                            "")
                        
                        
                        
                        'Checa se o CFOP veio zerado
                        If Trim(aItem(cItens)(6)) = "" Then
                            msgValid "Item " & ZE(cItens, 3) & ": Erro ao localizar CFOP. id_Produto: " & aItem(cItens)(0) & " CST: " & pgDadosEstoqueProduto(idProduto).ICMSCST
                            MontarVariaveis = False
                            Exit Function
                        End If
                            
                        aEstoque(cItens) = Array("0", aItem(cItens)(8), "0")
                        If LCase(aItem(cItens)(7)) <> LCase(pgDadosEstoqueProduto(idProduto).Unidade) Then
                                'Estoque_Unid|Estoque_Qtd|Estoque_vUnit
                                
                                aEstoque(cItens)(0) = pgDadosEstoqueProduto(idProduto).Unidade
                                aEstoque(cItens)(1) = InputBox("A UNIDADE vendida (" & UCase(aItem(cItens)(7)) & ") diverge da unidade de armazenamento" & vbCrLf & vbCrLf & _
                                                            "Informe a quantidade TOTAL em " & aEstoque(cItens)(0) & " do:" & vbCrLf & vbCrLf & _
                                                            "Item " & cItens + 1 & " - " & aItem(cItens)(3) & vbCrLf & " ", _
                                                            "Dados para Baixa de Estoque", aItem(cItens)(8))
                                If Trim(aEstoque(cItens)(1)) = 0 Then
                                    MsgBox "Não é possivel achar o coeficiente de uma divisão por ZERO!", vbInformation, "Aviso"
                                    MontarVariaveis = False
                                    Exit Function
                                End If
                                aEstoque(cItens)(1) = IIf(Trim(aEstoque(cItens)(1)) = "", aItem(cItens)(8), aEstoque(cItens)(1))
                                        
                                
                                aEstoque(cItens)(2) = Val(aItem(cItens)(10)) / Val(ChkVal(CStr(aEstoque(cItens)(1)), 0, cDecQtd))
                             
                                
                                msgValid "Item " & Left("000", 3 - Len(Trim(cItens))) & Trim(cItens) + 1 & ": Unidade: " & aEstoque(cItens)(0) & " - Quantidade: " & aEstoque(cItens)(1) & " - Valor Unitario: " & ConvMoeda(CStr(aEstoque(cItens)(2)))
                            Else
                                aEstoque(cItens)(0) = pgDadosEstoqueProduto(idProduto).Unidade
                                aEstoque(cItens)(1) = aItem(cItens)(8)
                                aEstoque(cItens)(2) = aItem(cItens)(9)
                        End If
                        
                        'aICMS
                        '    Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|
                        '    pMVAST|pRedBCST|vBCST|pICMSST|vICMSST|pCredSN|
                        '    vCredICMSSN|pFCP|vFCP
                        '
                        'CST = Prevalece o CST da PV
                        
                        aICMS(cItens) = Array(pgDadosEstoqueProduto(idProduto).ICMSOrigem, _
                                            CStr(cNull(Rst2.Fields("CST"))), _
                                            PgDadosTpNotaFiscal(idTpNF).ModBC, _
                                            "0.00", _
                                            "0.00", _
                                            CStr(cNull(Rst2.Fields("pICMS"))), _
                                            "0.00", _
                                            0, _
                                            0, _
                                            "0.00", _
                                            CStr(cNull(Rst2.Fields("vBCICMSST"))), _
                                            "0.00", _
                                            CStr(cNull(Rst2.Fields("vICMSST"))), _
                                            "0.00", "0.00", _
                                            cNull(Rst2.Fields("pICMSFCP")), "")  'pFCP e vFCP
                                            
                                            'Mudar a aliquota de icms conforme ncm
                                            'aICMS(cItens)(5) = IIf(pgAliqDifICMS(CStr(aItem(cItens)(5)), dest_UF) = "", aICMS(cItens)(5), pgAliqDifICMS(CStr(aItem(cItens)(5)), dest_UF))
                                            
                                            'Checa se o CST diverge com o que esta no Estoque
                                            
                                            If aICMS(cItens)(1) <> pgDadosEstoqueProduto(idProduto).ICMSCST Then
                                            'CStr(IIf(cNull(Rst2.Fields("CST")) <> "", Rst2.Fields("CST"), pgDadosEstoqueProduto(idProduto).ICMSCST))
                                                If MsgBox("O CST informado no ITEM " & cItens + 1 & " da proposta diverge do CST constante na base de dados do produto." & vbCrLf & vbCrLf & _
                                                           "CST do Produto: " & pgDadosEstoqueProduto(idProduto).ICMSCST & vbCrLf & _
                                                           "CST da Pre-Venda: " & aICMS(cItens)(1) & vbCrLf & vbCrLf & _
                                                            "Deseja modificar o CST pelo codigo registrado na base de dados do produto?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                                    msgValid "Item " & cItens + 1 & ": Alteração do CST de " & IIf(Trim(aICMS(cItens)(1)) = "", "<VAZIO>", aICMS(cItens)(1)) & " para " & pgDadosEstoqueProduto(idProduto).ICMSCST
                                                    aICMS(cItens)(1) = pgDadosEstoqueProduto(idProduto).ICMSCST
                                                End If
                                                
                                            End If
                        '22.12.17 - Informar caso haja DIFAL
                        'vBCUFDest|pFCPUFDest|pICMSUFDest|pICMSInter|pICMSInterPart|vFCPUFDest|vICMSUFDest|vICMSUFRemet|
                        aIcmsDifal(cItens) = Array(0, 0, 0, 0, 0, 0, 0, 0)
                        
                                            'cEnq|CST|vBC|pIPI|vIPI
                        aIPI(cItens) = Array(pgDadosEstoqueProduto(idProduto).Enquadramento, _
                                            pgDadosEstoqueProduto(idProduto).IPICST, _
                                            CStr(ChkVal(Rst2.Fields("SubTotal"), 0, 2)), _
                                            CStr(ChkVal(IIf(IsNull(Rst2.Fields("IPI")), "0", Rst2.Fields("IPI")), 0, 2)), _
                                            CStr(ChkVal(Rst2.Fields("VlIPI"), 0, 2)))
                                            aIPI(cItens)(4) = ChkVal(CStr(aIPI(cItens)(4)), 0, 2)
                                            'Alteracao incluida devido o material poder sair ou não com o IPI incluso
                                            'conforme manual pag.143
                                            'Caso haja NF de entrda essa alteracao devera ser modificada
                                            aIPI(cItens)(1) = pgCSTIPI(ide_tpNF, CStr(aIPI(cItens)(3)))
                                            
                                            
                                            'CST|vBC|pPIS|vPIS
                        aPIS(cItens) = Array(PgDadosTpNotaFiscal(idTpNF).CSTPIS, _
                                            CStr(ChkVal(Rst2.Fields("SubTotal"), 0, 2)), _
                                            CStr(ChkVal(PgDadosEmpresa(ID_Empresa).PISAliquota, 0, 2)), _
                                            (Val(ChkVal(Rst2.Fields("SubTotal"), 0, 2)) * Val(ChkVal(PgDadosEmpresa(ID_Empresa).PISAliquota, 0, 2))) / 100)
                                            
                                            'aPIS(cItens)(3) = ChkVal(CStr(aPIS(cItens)(3)), 0, 2)
                                            'CalcPIS_Item
                                            
                                            'CST|vBC|pCOFINS|vCOFINS
                        aCOFINS(cItens) = Array(PgDadosTpNotaFiscal(idTpNF).CSTCOFINS, _
                                            CStr(ChkVal(Rst2.Fields("SubTotal"), 0, 2)), _
                                            CStr(ChkVal(PgDadosEmpresa(ID_Empresa).COFINSAliquota, 0, 2)), _
                                            (Val(ChkVal(Rst2.Fields("SubTotal"), 0, 2)) * Val(ChkVal(PgDadosEmpresa(ID_Empresa).COFINSAliquota, 0, 2))) / 100)
                                            
                                            'aCOFINS(cItens)(3) = ChkVal(CStr(aCOFINS(cItens)(3)), 0, 2)
                                            
                                            'CalcCOFINS_Item
                                            
                                            'pComissao|vComissao
                                            'Array(IIf(PgDadosTpNotaFiscal(idTpNF).MovComissao = 1, PgDadosRhFuncionario(ger_Vendedor).Comissao, 0), _
                                            ChkVal(Val(ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).MovComissao = 1, PgDadosRhFuncionario(ger_Vendedor).Comissao, 0), 0, 3)) * Val(ChkVal(CStr(aItem(cItens)(10)), 0, cDecMoeda)) / 100, 0, cDecMoeda))
                        
                                '*** Movimenta Comissao ***
                                'aComissao(cItens) = Array(IIf(PgDadosTpNotaFiscal(idTpNF).MovComissao = 1, PgDadosRhFuncionario(ger_Vendedor).Comissao, 0), calcComissao(cItens))
                                '30.10.2012 - Funcao para deduzir o IPI do valo total do produto
                                calcComissao cItens
                        

                        Rst2.MoveNext
                        cItens = cItens + 1
                    Loop
                    cItens = cItens - 1
            End If
            'Calcula os Impostos
            If CalcICMS_Item = False Then
                MontarVariaveis = False
                Exit Function
            End If
            CalcPIS_Item
            CalcCOFINS_Item
            DistribuirValorFrete
            
            '*******************************************************
            '****** COBRANCA ***************************************
            '*******************************************************
            condPag = Rst1.Fields("CondicoesPagamento")
            sSQL = "SELECT * FROM FinanceiroCondicoesPagamentoParcelas WHERE IdCondicoes = " & condPag & " ORDER BY Parcela"
            Set Rst3 = RegistroBuscar(sSQL)
            If Rst3.BOF And Rst3.EOF Then
                    MsgBox "Erro ao encontrar as parcelas"
                Else
                    Rst3.MoveFirst
                    cCob = 0
                    Do Until Rst3.EOF
'############################################################################################
'# OBSERVACOES                                                                              #
'# nDup = ide_nNF & "-" & cCob + 1 & "/" & Rst3.RecordCount,                                #
'# o numero da duplicata deve ser montado na hora de gravar a NFe para que possa receber o  #
'# numero da NFe                                                                            #
'############################################################################################
                                    'nfat|vOrig|vDesc|vLiq|nDup|dVenc|vDup|TpDoc|idCliente|Emissao|Mora|Protesto
                        aCob(cCob) = Array(ide_nNF, _
                                            ChkVal(Rst1.Fields("VlTotalPV"), 0, 2), _
                                            "", _
                                            Val(ChkVal(Rst1.Fields("VlTotalPV"), 0, 2)) - Val(ChkVal("0.00", 0, 2)), _
                                            cCob + 1 & "/" & Rst3.RecordCount, _
                                            CDate(ide_dEmi) + IIf(IsNull(Rst3.Fields("DiasCorridos")), 0, Rst3.Fields("DiasCorridos")), _
                                            Val(ChkVal(IIf(IsNull(Rst3.Fields("Percentual")), 0, Rst3.Fields("Percentual")), 0, 3)) * Val(ChkVal(Rst1.Fields("VlTotalPV"), 0, 2)) / 100, _
                                            CStr(Rst1.Fields("FormaPagamento")), _
                                            dest_idDest, _
                                            ide_dEmi, _
                                            pgDadosConta(PgDadosTpNotaFiscal(idTpNF).conta).Juros, _
                                            pgDadosConta(PgDadosTpNotaFiscal(idTpNF).conta).DiasProtesto, _
                                            pgDadosConta(PgDadosTpNotaFiscal(idTpNF).conta).Multa)
                        aCob(cCob)(3) = ChkVal(CStr(aCob(cCob)(3)), 0, 2)
                        
                        '??? Analisar o objetivo dessa linah de codigo
                        'acredito que seja caso o valor tenha sido alterado na PV
                        aCob(cCob)(6) = pgValorDuplicataPV(aCob(cCob)(6), idPedido, cCob)
                        
                        Rst3.MoveNext
                        cCob = cCob + 1
                    Loop
                    cCob = cCob - 1
            End If
    End If
    
    Rst1.Close
    'cobrAcertarParcelas
    
    
    '*****************************************************************************************************
    '*****************************************************************************************************
    '*****************************************************************************************************

    'Transporte
    transp_modFrete = FreteConta
    
    If TranspRetEnt = 0 Then
            'Retira
            transp_Pessoa = UCase(dest_pessoa)
            transp_CNPJ = dest_CNPJ
            transp_xNome = dest_xNome
            transp_IE = dest_IE
            transp_xEnder = dest_xLgr & " " & dest_nro & " " & dest_xCpl
            transp_xMun = dest_xMun
            transp_UF = dest_UF
        Else
            If IdTransp = 0 Or RS(pgDadosTransportadora(IdTransp).CNPJ) = RS(PgDadosEmpresa(ID_Empresa).CNPJ) Then
                    'Entrega pelo carro da empresa
                    transp_Pessoa = UCase("JURIDICA")
                    transp_CNPJ = RS(PgDadosEmpresa(ID_Empresa).CNPJ)
                    transp_xNome = rc(PgDadosEmpresa(ID_Empresa).Nome)
                    transp_IE = RS(PgDadosEmpresa(ID_Empresa).IE)
                    transp_xEnder = rc(PgDadosEmpresa(ID_Empresa).Lgr & " " & PgDadosEmpresa(ID_Empresa).Nro)
                    transp_xMun = rc(PgDadosEmpresa(ID_Empresa).Mun)
                    transp_UF = PgDadosEmpresa(ID_Empresa).uf
                    '*******************************************************************
                    '* 27.12.2014
                    '* If incluso para somente solicitar esta informacao caso o sistema esteja
                    '* rodando na Metal Center
                    If RS(PgDadosEmpresa(ID_Empresa).CNPJ) = "40253676000101" Then
                            transp_VeicPlaca = Trim(UCase(InputBox("Informe a Placa do veículo ou: " & vbCrLf & _
                                                " 1 - Bongo (KYL-7016)" & vbCrLf & _
                                                " 2 - Caminhão (KNZ-3348)" & vbCrLf, "Veiculo")))
                        Else
                            transp_VeicPlaca = ""
                    End If
                    '*******************************************************************
                    If Trim(transp_VeicPlaca) = 1 Then
                            transp_VeicPlaca = "KYL7016"
                        ElseIf (Trim(transp_VeicPlaca) = 2) Then
                            transp_VeicPlaca = "KNZ3348"
                    End If
                    
                    If Trim(transp_VeicPlaca) = "" Then
                            transp_VeicUF = ""
                            msgValid "Transporte Placa/UF: <Não informado>"
                        Else
                            transp_VeicUF = UCase(InputBox("Informe a UF do veículo?", "Veiculo", PgDadosEmpresa(ID_Empresa).uf))
                            msgValid "Transporte Placa/UF: " & transp_VeicPlaca & "/" & transp_VeicUF
                    End If
                    
                Else
                    'Entrega por transportadora
                    transp_Pessoa = UCase(pgDadosTransportadora(IdTransp).Pessoa)
                    transp_CNPJ = RS(pgDadosTransportadora(IdTransp).CNPJ)
                    transp_xNome = rc(pgDadosTransportadora(IdTransp).Nome)
                    transp_IE = RS(pgDadosTransportadora(IdTransp).IE)
                    transp_xEnder = rc(pgDadosTransportadora(IdTransp).Lgr) '& " " & pgDadosTransportadora(IdTransp).Nro)
                    transp_xMun = rc(pgDadosTransportadora(IdTransp).Mun)
                    transp_UF = pgDadosTransportadora(IdTransp).uf
            End If
    End If
    transp_qVol = IIf(Trim(qVol) = "", "0", Trim(qVol))
    If Trim(transp_qVol) = "0" Then
        msgValid "Transporte/Quantidade esta vazio..."
    End If
     transp_esp = rc(Esp)
    If Trim(transp_esp) = "" Then
        msgValid "Transporte/Especie esta vazio..."
    End If
     transp_marca = IIf(Trim(Marca) = "", "S/M", rc(Marca))
     transp_nVol = IIf(Trim(nVol) = "", "S/N", rc(nVol))
     transp_pesoL = ChkVal(IIf(Trim(PesoL) = "", 0, PesoL), 0, 3)
     If Trim(PesoL) = "" Then
        msgValid "Transporte/Peso Liquido esta vazio..."
    End If
     
     transp_pesoB = ChkVal(IIf(Trim(PesoB) = "", 0, PesoB), 0, 3)
     If Trim(PesoB) = "" Then
        msgValid "Transporte/Peso Bruto esta vazio..."
    End If
    'TOTAIS
    CalcDIFAL_Total
    total_vBC = ChkVal(CalcICMS_Total(4), 0, 2)
    total_vICMS = ChkVal(CalcICMS_Total(6), 0, 2)
    total_vFCP = CalcFCP_Total(16)
    total_vBCST = ChkVal(CalcICMS_Total(10), 0, 2)
    total_vICMSST = ChkVal(CalcICMS_Total(12), 0, 2)
    total_vCredICMSSN = ChkVal(CalcICMS_Total(14), 0, 2)
    
    total_vProd = ChkVal(total_vProd, 0, 2)
    total_vFrete = ChkVal(total_vFrete, 0, 2)
    total_vSeg = ChkVal(total_vSeg, 0, 2)
    total_vDesc = ChkVal(total_vDesc, 0, 2)
    total_vOutro = ChkVal(total_vOutro, 0, 2)
    total_vIPI = ChkVal(CalcIPI_Total(4), 0, 2)
    total_vPIS = ChkVal(CalcPIS_Total(3), 0, 2)
    total_vCOFINS = ChkVal(CalcCOFINS_Total(3), 0, 2)
    total_vNF = (Val(total_vProd) + Val(total_vFrete) + Val(total_vIPI) + Val(total_vSeg) + Val(total_vOutro) + Val(total_vICMSST)) - Val(total_vDesc)
    total_vNF = ChkVal(total_vNF, 0, 2)
    
    
    
    
    ' 13/07/18 - Lanca nas obs o valor do Aproveitamento de ICMS
    '            para o % utiliza o valor do pICMS do 1 item
    If ChkVal(total_vCredICMSSN, 0, 2) <> "0.00" Then
        infAdic_infCpl = infAdic_infCpl & _
                        " [PERMITE O APROVEITAMENTO DO CRÉDITO DE ICMS " & _
                        "NO VALOR DE R$" & total_vCredICMSSN & " " & _
                        "CORRESPONDENTE À ALÍQUOTA DE " & ChkVal(CStr(pCredICMS), 0, 2) & " %, " & _
                        "NOS TERMOS DO ARTIGO 23 DA LC 123.]"
        txtObs.Text = Trim(rc(infAdic_infCpl))
    End If
    '*****************************************************************************************************
    
    '###########################################################
    '### 26/07/18 - Lanca nas obs o valor FCP ICMS
    '###########################################################
    
    If ChkVal(total_vFCP, 0, 2) <> "0.00" Then
        infAdic_infCpl = infAdic_infCpl & _
                        " [Total do vFCP: " & ConvMoeda(total_vFCP) & "]"
        txtObs.Text = Trim(rc(infAdic_infCpl))
    End If
    '*****************************************************************************************************
    
    '*****************************************************************************************************
    '*****************************************************************************************************
    MontarVariaveis = True
    'msgValid "************ NOTA FISCAL ELETRONICA VALIDADA ************"
    Exit Function
TrtMVErro:
    MsgBox "Numero: " & Err.Number & vbCrLf & vbCrLf & "Descricao: " & Err.Description, vbInformation, "Aviso"
    MontarVariaveis = False
    Exit Function
End Function
'
Private Function ValidarVariaveis() As Boolean
    Dim i           As Integer
    Dim nmVendedor  As String
    ValidarVariaveis = False
    
    
    
    
    If PgDadosTpNotaFiscal(idTpNF).ChaveAcessoRef = 1 Then
            If Trim(txtrefNFe.Text) = "" Then
                    MsgBox "O campo CHAVE DA  NF-e DE REFENENCIA não poder ser em branco.", vbInformation, App.EXEName
                    ValidarVariaveis = False 'MontarVariaveis = False
                    Exit Function
                Else
                    If PesquisarNFe(Trim(txtrefNFe.Text)) = True Then
                            ide_refNFe = txtrefNFe.Text 'Criar uma function para validar esta chave
                            txtObs.Text = txtObs.Text & "; Chave Ref.: " & ide_refNFe
                        Else
                            If MsgBox("Chave não cadastrada na base de dados. Deseja continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                                ValidarVariaveis = False 'MontarVariaveis = False
                                Exit Function
                            End If
                            ide_refNFe = txtrefNFe.Text
                    End If
            End If
        Else
            ide_refNFe = ""
    End If
    
    
    'cabecario do Pedido (ide)
    'Versao = "2.00"
    'ide_cUF = PgDadosUF(PgDadosEmpresa(ID_Empresa).UF).Id
    'ide_cNF = Format(Now(), "DDHHMMSS"): ide_cNF = Mid(String(8, "0"), 1, 8 - Len(Trim(ide_cNF))) & Trim(ide_cNF)
    'ide_natOp = RC(PgDadosTpNotaFiscal(idTpNF).Natureza)
    'ide_indPag = "2" 'Indica Forma de pagamento: 0 - avista, 1 - prazo , 2 -  outros
    'ide_mod = PgDadosTpNotaFiscal(idTpNF).Modelo
    'ide_serie = PgDadosTpNotaFiscal(idTpNF).Serie
    'ide_nNF = Left(String(9, "0"), 9 - Len(Trim(txtNumNota.Text))) & Trim(txtNumNota.Text)
    'ide_dEmi = dtpEmissao.Value
    'ide_dSaiEnt = IIf(PgDadosTpNotaFiscal(idTpNF).ImpDtSaida = 0, "", dtpSaida.Value)
    'ide_hSaiEnt = IIf(PgDadosTpNotaFiscal(idTpNF).ImpDtSaida = 0, "", Format(Time, "HH:MM:SS"))
    'ide_tpNF = PgDadosTpNotaFiscal(idTpNF).TipoNota
    'ide_cMunFG = PgDadosMunicipio(PgDadosEmpresa(ID_Empresa).UF, PgDadosEmpresa(ID_Empresa).Mun).codMun
    'ide_refNFe = ""
    'ide_tpImp = "1" 'Colocar no formFaturamentoTipoNotaFiscal essa opcao e configurar o Unidanfe automaticamente"
    'ide_tpEmis = PgDadosConfig.TpEmissao
    'ide_tpAmb = PgDadosConfig.Ambiente
    'ide_finNFe = PgDadosTpNotaFiscal(idTpNF).Finalidade
    'ide_procEmi = "0"
    'ide_verProc = sVersao
    'ide_cDV = Right(strChaveAcesso, 1)
    'Emitente
    'emit_CNPJ = RS(PgDadosEmpresa(ID_Empresa).CNPJ)
    'emit_xNome = RC(PgDadosEmpresa(ID_Empresa).Nome)
    'emit_xFant = RC(PgDadosEmpresa(ID_Empresa).Fant)
    'emit_xLgr = RC(PgDadosEmpresa(ID_Empresa).Lgr)
    'emit_nro = RC(PgDadosEmpresa(ID_Empresa).Nro)
    'emit_xCpl = RC(PgDadosEmpresa(ID_Empresa).Cpl)
    'emit_Bairro = RC(PgDadosEmpresa(ID_Empresa).Bairro)
    'emit_cMun = PgDadosMunicipio(PgDadosEmpresa(ID_Empresa).UF, PgDadosEmpresa(ID_Empresa).Mun).codMun
    'emit_xMun = RC(PgDadosEmpresa(ID_Empresa).Mun)
    'emit_UF = PgDadosEmpresa(ID_Empresa).UF
    'emit_CEP = PgDadosEmpresa(ID_Empresa).CEP
    'emit_cPais = pgIdPais(PgDadosEmpresa(ID_Empresa).pais)
    'emit_xPais = RC(PgDadosEmpresa(ID_Empresa).pais)
    'emit_fone = RS(PgDadosEmpresa(ID_Empresa).Fone)
    'emit_IE = RS(PgDadosEmpresa(ID_Empresa).IE)
    'emit_IEST = RS(PgDadosEmpresa(ID_Empresa).iest)
    'emit_IM = RS(PgDadosEmpresa(ID_Empresa).im)
    'emit_CNAE = RS(PgDadosEmpresa(ID_Empresa).cnae)
    'emit_CRT = PgDadosEmpresa(ID_Empresa).RegimeTrib
    'Destinatario
    'dest_idDest = IdCliente
    'dest_pessoa = PgDadosCliente(IdCliente).pessoa 'NAO PODE SER DEIXADO EM  BRANCO
    'dest_CNPJ = RS(PgDadosCliente(IdCliente).doc)
    'dest_xNome = RC(PgDadosCliente(IdCliente).Nome)
    'dest_xFant = RC(PgDadosCliente(IdCliente).Fant)
    'dest_xLgr = RC(PgDadosCliente(IdCliente).Lgr)
    'dest_nro = RC(PgDadosCliente(IdCliente).Nro)
    'dest_xCpl = RC(PgDadosCliente(IdCliente).Cpl)
    'dest_Bairro = RC(PgDadosCliente(IdCliente).Bairro)
    'dest_cMun = PgDadosMunicipio(PgDadosCliente(IdCliente).UF, PgDadosCliente(IdCliente).Mun).codMun
    'dest_xMun = RC(PgDadosCliente(IdCliente).Mun)
    'dest_UF = PgDadosCliente(IdCliente).UF
    'dest_CEP = RS(PgDadosCliente(IdCliente).CEP)
    'dest_cPais = emit_cPais
    'dest_xPais = RC(emit_xPais)
    'dest_fone = RS(PgDadosCliente(IdCliente).Fone)
    'dest_IE = RS(PgDadosCliente(IdCliente).IE)
    'dest_ISUF = RS(PgDadosCliente(IdCliente).SUFRAMA)
    'dest_email = RC(PgDadosCliente(IdCliente).Mail)
    If PgDadosConfig.InserirNomeVendXML = 1 Then
        nmVendedor = PgDadosRhFuncionario(ger_Vendedor).Nome & " "
        nmVendedor = Trim(Mid(nmVendedor, 1, InStr(nmVendedor, " ")))
        infAdic_infCpl = rc(Trim(txtObs.Text) & "[Vend.:" & nmVendedor & "]")
    End If
    '###########################################################
    '### 13/07/2012 - Incluido como recurso para msgRedICMS
    infAdic_infCpl = infAdic_infCpl & " " & msgRedICMS
    txtObs.Text = Trim(rc(infAdic_infCpl))
    '###########################################################
    'Pegar os ITENS do pedido
    'id_intProd|det_cProd|det_cEAN|det_xProd|EXTIPI|det_NCM|det_CFOP|det_uCom|
    'det_qCom|det_vUnCom|det_vProd|det_cEANTrib|det_uTrib|det_qTrib|
    'det_vUnTrib|
    'det_vFrete|det_vSeg|det_vDesc|det_vOutro|det_indTot|xPed|nItemPed
    'det_indTot = 0 = O valor do item compoe a NF / 1  = O valor do item nao compoe a NF
    
    '###########################################################################################
    '### Vai validar os campos da NFE
    '### futuramente transferir esta instrucao ´para a funcao a MontarVariaveis
    '###########################################################################################
    '#### ITEMS
    For i = 0 To cItens
        'Verifica se existe NCM no item
        If Trim(aItem(i)(5)) <> pgDadosEstoqueProduto(CInt(aItem(i)(0))).NCM Then
            If MsgBox("Item " & i + 1 & " - NCM diferente do cadastrado no produto!" & vbCrLf & vbCrLf & _
                   "NCM na PV: " & aItem(i)(5) & vbCrLf & _
                   "NCM do Produto: " & pgDadosEstoqueProduto(CInt(aItem(i)(0))).NCM & vbCrLf & vbCrLf & _
                   "Deseja substituir?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                        'Substituir NCM
                        ValidarVariaveis = True
                        msgValid "Item " & ZE(CInt(i) + 1, 3) & ": NCM (" & aItem(i)(5) & ") substituido por " & pgDadosEstoqueProduto(CInt(aItem(i)(0))).NCM & " do cadastro do produto."
                        aItem(i)(5) = pgDadosEstoqueProduto(CInt(aItem(i)(0))).NCM
                    Else
                        'ValidarVariaveis = False
                        msgValid "Item " & ZE(i + 1, 3) & ": NCM (" & aItem(i)(5) & ") diferente do cadastrado no produto (" & pgDadosEstoqueProduto(CInt(aItem(i)(0))).NCM & ")"
                        'Exit Function
            End If
            
        End If
        'Quantidade 8,13
        aItem(i)(8) = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvTotalProduto = 0, 0, CStr(aItem(i)(8))), 0, 4)
        aItem(i)(13) = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvTotalProduto = 0, 0, CStr(aItem(i)(13))), 0, 4)
        'Valor Unitario 9,14
        aItem(i)(9) = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvTotalProduto = 0, 0, CStr(aItem(i)(9))), 0, 10)
        aItem(i)(14) = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvTotalProduto = 0, 0, CStr(aItem(i)(14))), 0, 10)
        'Valor do Produto
        aItem(i)(10) = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvTotalProduto = 0, 0, CStr(aItem(i)(10))), 0, 2)
    'Next
    
    
    
    'aEstoque (cItens)
    '#### ICMS
    '0-Origem|1-CST|2-ModBC|3-pRedBC|4-vBC|5-pICMS|
    '6-vICMS|7-modBCST|8-pMVAST|9-pRedBCST|10-vBCST|11-pICMSST|12-vICMSST
    'For i = 0 To cItens
        aICMS(i)(3) = ChkVal(CStr(aICMS(i)(3)), 0, 2)
        aICMS(i)(4) = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpBCICMS = 0, 0, CStr(aICMS(i)(4))), 0, 2)
        aICMS(i)(5) = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpBCICMS = 0, 0, CStr(aICMS(i)(5))), 0, 2)
        aICMS(i)(6) = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvICMS = 0, 0, CStr(aICMS(i)(6))), 0, 2) 'ChkVal(CStr(aICMS(i)(6)), 0, 2)
        aICMS(i)(10) = ChkVal(CStr(aICMS(i)(10)), 0, 2)
        aICMS(i)(11) = ChkVal(CStr(aICMS(i)(11)), 0, 2)
        aICMS(i)(12) = ChkVal(CStr(aICMS(i)(12)), 0, 2)
   ' Next
    '#### IPI
    'cEnq|CST|vBC|pIPI|vIPI
    'For i = 0 To cItens
        'vBC
        aIPI(i)(2) = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvIPI = 0, 0, CStr(aIPI(i)(2))), 0, 2)
        'pIPI
        aIPI(i)(3) = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvIPI = 0, 0, CStr(aIPI(i)(3))), 0, 2)
        'vIPI
        aIPI(i)(4) = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvIPI = 0, 0, CStr(aIPI(i)(4))), 0, 2)
    Next
    'PIS
    'CST|vBC|pPIS|vPIS
    'aPIS (cItens)
    'CST|vBC|pCOFINS|vCOFINS
    'aCOFINS (cItens)
    'pComissao|vComissao
    'aComissao (cItens)
    
    'nfat|vOrig|vDesc|vLiq|nDup|dVenc|vDup|TpDoc|idCliente|Emissao|Mora|Protesto
    'Cobranca
    cobrAcertarParcelas
    
    '========================================================
    'Removido em 26/10/2012- Pois as faturas
    'devem ser geradas porem nao impressas na NFe.
    '
    'If PgDadosTpNotaFiscal(idTpNF).ImpCmpFatura = 0 Then
    '    For i = 0 To cCob
    '        Dim b As Integer
    '        For b = 0 To 12
    '            aCob(i)(b) = Empty
    '        Next
    '    Next
    'End If
    '========================================================
    
    'Transporte
    AlterarTransportadora
    'transp_modFrete = FreteConta
    'Entrega por transportadora
    'transp_CNPJ = RS(pgDadosTransportadora(IdTransp).CNPJ)
    'transp_xNome = RC(pgDadosTransportadora(IdTransp).Nome)
    'transp_IE = RS(pgDadosTransportadora(IdTransp).IE)
    
    transp_xEnder = IIf(Len(transp_xEnder) > 60, Mid(transp_xEnder, 1, 60), transp_xEnder)
    
    'transp_xMun = RC(pgDadosTransportadora(IdTransp).Mun)
    'transp_UF = pgDadosTransportadora(IdTransp).UF
    If PgDadosConfig.TranspVolumes = 1 Then
        If emit_CNPJ <> transp_CNPJ And dest_CNPJ <> transp_CNPJ Then
            If Trim(transp_qVol) = "" Then
                MsgBox "Campo TRANSPORTADOR/QUANTIDADE é obrigatorio!", vbInformation, "Aviso"
                ValidarVariaveis = False
                Exit Function
            End If
            If Trim(transp_esp) = "" Then
                MsgBox "Campo TRANSPORTADOR/ESPECIE é obrigatorio!", vbInformation, "Aviso"
                ValidarVariaveis = False
                Exit Function
            End If
    
            'transp_marca = IIf(Trim(Marca) = "", "S/M", RC(Marca))
            'transp_nVol = IIf(Trim(nVol) = "", "S/N", RC(nVol))
            If Trim(transp_pesoL) = 0 Then
                MsgBox "Campo TRANSPORTADOR/PESO LIQUIDO é obrigatorio!", vbInformation, "Aviso"
                ValidarVariaveis = False
                Exit Function
            End If
            If Trim(transp_pesoB) = 0 Then
                MsgBox "Campo TRANSPORTADOR/PESO BRUTO é obrigatorio!", vbInformation, "Aviso"
                ValidarVariaveis = False
                Exit Function
            End If
        End If
    End If
    'TOTAIS
    total_vBC = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpBCICMS = 0, 0, CalcICMS_Total(4)), 0, 2)
    total_vICMS = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvICMS = 0, 0, CalcICMS_Total(6)), 0, 2)
    total_vBCST = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpBCICMSST = 0, 0, CalcICMS_Total(10)), 0, 2)
    total_vICMSST = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvICMSST = 0, 0, CalcICMS_Total(12)), 0, 2)
    total_vProd = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvTotalProduto = 0, 0, total_vProd), 0, 2)
    total_vFrete = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvFrete = 0, 0, total_vFrete), 0, 2)
    total_vSeg = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvSeguro = 0, 0, total_vSeg), 0, 2)
    total_vDesc = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvDesconto = 0, 0, total_vDesc), 0, 2)
    total_vOutro = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvOutrasDesp = 0, 0, total_vOutro), 0, 2)
    total_vIPI = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvIPI = 0, 0, CalcIPI_Total(4)), 0, 2)
    'total_vPIS = ChkVal(CalcPIS_Total(3), 0, 2)
    'total_vCOFINS = ChkVal(CalcCOFINS_Total(3), 0, 2)
    total_vNF = ChkVal(IIf(PgDadosTpNotaFiscal(idTpNF).ImpvTotalNota = 0, 0, total_vNF), 0, 2)
    
    '############################################################################
    '### Checa o limite de credito do cliente
    '############################################################################
    If PgDadosConfig.ClienteAplLimiteCredito = 1 Then
        If PgDadosCliente(idCliente).LimiteCredito < ChkVal(Val(ChkVal(total_vNF, 0, cDecMoeda)) + Val(ClientePosicaoFinanceira(idCliente).Pagar), 0, cDecMoeda) Then
            MsgBox "Cliente ultrapassou o LIMITE DE CREDITO (" & ConvMoeda(PgDadosCliente(idCliente).LimiteCredito) & ")!", vbInformation, App.EXEName
            ValidarVariaveis = False
            Exit Function
        End If
    End If
    
    
    ValidarVariaveis = True
    AlterarVendedor
    
    
    
End Function
Private Function pgCSTIPI(tpNF As String, pIPI As String) As String
    Select Case tpNF
        Case 0 'Entrada
            If pIPI = 0 Then
                    pgCSTIPI = "49"
                Else
                    pgCSTIPI = "00"
            End If
        Case 1 'Saida
            If pIPI = 0 Then
                    pgCSTIPI = "99"
                Else
                    pgCSTIPI = "50"
            End If
        Case Else
            MsgBox "Não foi possivel classificar o CST do IPI devido o tipo de NFe não ter sido determinado.", vbInformation, "Aviso"
            pgCSTIPI = ""
    End Select
End Function

Private Sub MontarBaseDeDados()
    Dim vDados(1000)    As Variant
    Dim contReg         As Integer
    Dim i               As Integer
    
    contReg = 0
    
    'vDados(contReg) = Array("ChaveAcesso", "100", "S")
    'contReg = contReg + 1
    
    'Outros
    vDados(contReg) = Array("movfisco", "3", "N"): contReg = contReg + 1
    vDados(contReg) = Array("movfinanceiro", "3", "N"): contReg = contReg + 1
    vDados(contReg) = Array("enviorf", "3", "N"): contReg = contReg + 1
    vDados(contReg) = Array("impfatura", "3", "N"): contReg = contReg + 1
    '************************************************************************
    
    
    'NFe Autorizada
    vDados(contReg) = Array("nProt", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dhProt", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Lote", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("nRecibo", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cStat", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("xMotivo", "65000", "S"): contReg = contReg + 1
    vDados(contReg) = Array("StatusNFe", "65000", "S"): contReg = contReg + 1
    '*********************************************************************************
    'NFe Cancelada
    vDados(contReg) = Array("canc_nProt", "250", "S"): contReg = contReg + 1
    vDados(contReg) = Array("canc_dhRecbto", "250", "S"): contReg = contReg + 1
    vDados(contReg) = Array("canc_xJust", "250", "S"): contReg = contReg + 1
    vDados(contReg) = Array("canc_Status", "250", "S"): contReg = contReg + 1
    '*********************************************************************************
    'Numero de NFe Inutilizado
    vDados(contReg) = Array("inut_nProt", "250", "S"): contReg = contReg + 1
    vDados(contReg) = Array("inut_dhRecbto", "250", "S"): contReg = contReg + 1
    vDados(contReg) = Array("inut_xJust", "250", "S"): contReg = contReg + 1
    vDados(contReg) = Array("inut_Status", "250", "S"): contReg = contReg + 1
    '*********************************************************************************
    
    'cabecario do Pedido (ide)
    vDados(contReg) = Array("versao", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("IdNFe", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_cUF", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_cNF", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_natOp", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_indPag", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_mod", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_serie", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_nNF", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_dEmi", "15", "D"): contReg = contReg + 1
    vDados(contReg) = Array("ide_hEmi", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_dSaiEnt", "15", "D"): contReg = contReg + 1
    vDados(contReg) = Array("ide_hSaiEnt", "25", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_tpNF", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_idDest", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_cMunFG", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_refNFe", "50", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_tpImp", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_tpEmis", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_cDV", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_tpAmb", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_finNFe", "5", "N"): contReg = contReg + 1
    
    vDados(contReg) = Array("ide_indFinal", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_indPres", "5", "N"): contReg = contReg + 1
    
    vDados(contReg) = Array("ide_procEmi", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_verProc", "20", "S"): contReg = contReg + 1
    
    'Emitente
    vDados(contReg) = Array("emit_id", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("emit_CNPJ", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_xNome", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_xFant", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_xLgr", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_nro", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_xCpl", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_Bairro", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_cMun", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_xMun", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_UF", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_CEP", "8", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_cPais", "4", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_xPais", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_fone", "14", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_IE", "14", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("emit_IEST", "14", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_IM", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("emit_CNAE", "10", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("emit_CRT", "5", "N"): contReg = contReg + 1
    
    'Destinatario
    vDados(contReg) = Array("dest_idDest", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("dest_pessoa", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_CNPJ", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_xNome", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_xFant", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_xLgr", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_nro", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_xCpl", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_Bairro", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_cMun", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_xMun", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_UF", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_CEP", "8", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_cPais", "4", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_xPais", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_fone", "14", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_IE", "14", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_ISUF", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_email", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("infAdic_infCpl", "5000", "S"): contReg = contReg + 1
    vDados(contReg) = Array("dest_indIEDest", "2", "N"): contReg = contReg + 1
    
    
    'Local Entrega
    vDados(contReg) = Array("entr_CNPJ", "20", "S"): contReg = contReg + 1
    'vDados(contReg) = Array("entr_CPF", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("entr_xLgr", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("entr_nro", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("entr_xCpl", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("entr_xBairro", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("entr_cMun", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("entr_xMun", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("entr_UF", "2", "S"): contReg = contReg + 1
    
    'Transporte
    vDados(contReg) = Array("transp_modFrete", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("transp_Pessoa", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_CNPJ", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_xNome", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_IE", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_xEnder", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_xMun", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_UF", "2", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_qVol", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_esp", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_marca", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_nVol", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_pesoL", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_pesoB", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_VeicPlaca", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("transp_VeicUF", "15", "S"): contReg = contReg + 1
    
    
        
    'TOTAIS
    vDados(contReg) = Array("total_vBC", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vICMS", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vBCST", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vFCPUFDest", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vICMSUFDest", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vICMSUFRemet", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vFCP", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vICMSST", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vCredICMSSN", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vProd", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vFrete", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vSeg", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vDesc", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vIPI", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vPIS", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vCOFINS", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vOutro", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vNF", "15", "S"): contReg = contReg + 1
    
    
    vDados(contReg) = Array("ger_Vendedor", "15", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ger_idPV", "20", "N") ': contReg = contReg + 1
    
    
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg
    
    'Produto******************************************************************************
    contReg = 0
    vDados(contReg) = Array("IdNFe", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_IdProduto", "60", "N"): contReg = contReg + 1
    vDados(contReg) = Array("det_cProd", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_cEAN", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_xProd", "120", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_InfAdProd", "500", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_NCM", "8", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_CEST", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_EXTIPI", "8", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_CFOP", "4", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_uCom", "6", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_qCom", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_vUnCom", "50", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_vProd", "50", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("det_cEANTrib", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_uTrib", "6", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_qTrib", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_vUnTrib", "50", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_vFrete", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_vSeg", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_vDesc", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_vOutro", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_indTot", "1", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_xPed", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_nItemPed", "6", "S"): contReg = contReg + 1
    
    'vDados(contReg) = Array("det_indTot", "15", "S")
    'contReg = contReg + 1
    'IMPOSTOS
    'ICMS - 'Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST|vICMSST
    vDados(contReg) = Array("ICMS_origem", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_CST", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_modBC", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_pRedBC", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vBC", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_pICMS", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vICMS", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_modBCST", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_pMVAST", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_pRedBCST", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vBCST", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_pICMSST", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vICMSST", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_MotDesICMS", "2", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vBCSTRet", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vICMSSTRet", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_pBCOP", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_UFST", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_ICMSST", "30", "S"): contReg = contReg + 1 'N10b pag.137 - v.4.0.1 - NT2009.006
    vDados(contReg) = Array("ICMS_vBCSTDest", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vICMSSTDes", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_CSOSN", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_pCredSN", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vCredICMSSN", "30", "S"): contReg = contReg + 1
    
    '26/07/18 - FCP NT 2016.002 v.1.60
    vDados(contReg) = Array("ICMS_pFCP", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vFCP", "30", "S"): contReg = contReg + 1
    
    'ICMS Difal - 21/09/17
    vDados(contReg) = Array("ICMS_vBCUFDest", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_pFCPUFDest", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_pICMSUFDest", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_pICMSInter", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_pICMSInterPart", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vFCPUFDest", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vICMSUFDest", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS_vICMSUFRemet", "30", "S"): contReg = contReg + 1
    
    
    'IPI
    vDados(contReg) = Array("IPI_cEnq", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("IPI_CST", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("IPI_vBC", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("IPI_pIPI", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("IPI_vIPI", "15", "S"): contReg = contReg + 1
    'PIS
    vDados(contReg) = Array("PIS_CST", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("PIS_vBC", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("PIS_pPIS", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("PIS_vPIS", "15", "S"): contReg = contReg + 1
    'COFINS
    vDados(contReg) = Array("COFINS_CST", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("COFINS_vBC", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("COFINS_pCOFINS", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("COFINS_vCOFINS", "15", "S"): contReg = contReg + 1
    'Informacoes Gerenciais
    vDados(contReg) = Array("estoque_Unid", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("estoque_Qtd", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("estoque_vUnit", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("comissao_pComissao", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("comissao_vComissao", "20", "S") ': contReg = contReg + 1
   
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "Itens"
    
    
    contReg = 0
    'COBRANCA
    vDados(contReg) = Array("IdNFe", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_TpDoc", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_nFat", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_vOrig", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_vDesc", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_vLiq", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_nDup", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_dVenc", "10", "D"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_vDup", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_Emissao", "15", "D"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_Multa", "50", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_Mora", "50", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_Protesto", "50", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_idCliente", "15", "N") ': contReg = contReg + 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "Cobranca"
    
    contReg = 0
    'Email
    vDados(contReg) = Array("IdNFe", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("IdCliente", "60", "N"): contReg = contReg + 1
    vDados(contReg) = Array("Status", "200", "S") ': contReg = contReg + 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "SendMail"
End Sub

Private Function grvRegistro() As Boolean
    Dim vReg(1000)   As Variant
    Dim cReg         As Integer 'Contador de Registros
    Dim IdReg        As Integer  'Pega o Id do registro gravado
    Dim i            As Integer
    
    'BloquearTabela strTabela
'##########################################################################################################################
'# Usar o bloqueio de tabela para gerar a chv de acesso e o num da nfe automaticamente.
'# Usar a funcao PgPxNumNota com a tab bloq assim ninquem pega o prox num
'##########################################################################################################################
    If PgDadosConfig.BloqueionNFManual = 0 Then
        ide_nNF = txtNumNota.Text 'PgPxNumNota
        ide_nNF = Left(String(9, "0"), 9 - Len(Trim(txtNumNota.Text))) & Trim(txtNumNota.Text)
        If NumNotaFiscalExiste(ide_nNF) = True Then 'Verifica se o num da nota ja esta cadastrado
                'MontarVariaveis = False
                MsgBox "Numero de Nota Fiscal ja cadastrado.", vbInformation, "Aviso"
                grvRegistro = False
                Exit Function
            End If
        Else
            ide_nNF = PgPxNumNota
            ide_nNF = Left(String(9, "0"), 9 - Len(Trim(ide_nNF))) & Trim(ide_nNF)
    End If
    ide_cNF = Format(Now(), "DDHHMMSS"): ide_cNF = Mid(String(8, "0"), 1, 8 - Len(Trim(ide_cNF))) & Trim(ide_cNF)
    ide_hEmi = Format(Now(), "HH:MM:SS")
'##########################################################################################################################
    strChaveAcesso = ChaveAcesso(ide_nNF, _
                                dtpEmissao.Value, _
                                ide_cUF, _
                                PgDadosEmpresa(ID_Empresa).CNPJ, _
                                PgDadosTpNotaFiscal(idTpNF).Modelo, _
                                PgDadosTpNotaFiscal(idTpNF).Serie, _
                                CStr(PgDadosConfig().TpEmissao), _
                                ide_cNF)
    Id = strChaveAcesso

    ide_cDV = Right(strChaveAcesso, 1)

'###################################################################################################
    
    cReg = 0
    vReg(cReg) = Array("MovFisco", PgDadosTpNotaFiscal(idTpNF).MovFisco, "N"): cReg = cReg + 1
    vReg(cReg) = Array("movfinanceiro", PgDadosTpNotaFiscal(idTpNF).MovContasPR, "N"): cReg = cReg + 1
    vReg(cReg) = Array("enviorf", PgDadosTpNotaFiscal(idTpNF).EnvioRF, "N"): cReg = cReg + 1
    vReg(cReg) = Array("impfatura", PgDadosTpNotaFiscal(idTpNF).ImpCmpFatura, "N"): cReg = cReg + 1
    
    
    vReg(cReg) = Array("IdNFe", Id, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Versao", VersaoNFe, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_cUF", ide_cUF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_cNF", ide_cNF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_natOp", ide_natOp, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_indPag", ide_indPag, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_mod", ide_mod, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_Serie", ide_serie, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_nNF", ide_nNF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_dEmi", ide_dEmi, "D"): cReg = cReg + 1
    vReg(cReg) = Array("ide_hEmi", ide_hEmi, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_dSaiEnt", ide_dSaiEnt, "D"): cReg = cReg + 1
    vReg(cReg) = Array("ide_hSaiEnt", ide_hSaiEnt, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_tpNf", ide_tpNF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_idDest", ide_idDest, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ide_cMunFG", ide_cMunFG, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_refNFe", ide_refNFe, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_tpImp", ide_tpImp, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_tpEmis", ide_tpEmis, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_cDV", ide_cDV, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_tpAmb", ide_tpAmb, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_finNFe", ide_finNFe, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("ide_indFinal", ide_indFinal, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ide_indPres", ide_indPres, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ide_procEmi", ide_procEmi, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_verProc", ide_verProc, "S"): cReg = cReg + 1
    'Emitente
    vReg(cReg) = Array("emit_id", ID_Empresa, "N"): cReg = cReg + 1
    vReg(cReg) = Array("emit_CNPJ", emit_CNPJ, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_xNome", emit_xNome, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_xFant", emit_xFant, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_xlgr", emit_xLgr, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_nro", emit_nro, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_xCpl", emit_xCpl, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_Bairro", emit_Bairro, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_cMun", emit_cMun, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_xMun", emit_xMun, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_UF", emit_UF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_CEP", emit_CEP, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_cPais", emit_cPais, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_xPais", emit_xPais, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_fone", emit_fone, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_IE", emit_IE, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_IEST", emit_IEST, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_IM", emit_IM, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_CNAE", emit_CNAE, "S"): cReg = cReg + 1
    vReg(cReg) = Array("emit_CRT", emit_CRT, "S"): cReg = cReg + 1
    'Destinatario
    vReg(cReg) = Array("dest_idDest", dest_idDest, "N"): cReg = cReg + 1
    vReg(cReg) = Array("dest_pessoa", dest_pessoa, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_CNPJ", dest_CNPJ, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_xNome", dest_xNome, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_xFant", dest_xFant, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_xlgr", dest_xLgr, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_nro", dest_nro, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_xCpl", dest_xCpl, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_Bairro", dest_Bairro, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_cMun", dest_cMun, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_xMun", dest_xMun, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_UF", dest_UF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_CEP", dest_CEP, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_cPais", dest_cPais, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_xPais", dest_xPais, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_fone", dest_fone, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_IE", dest_IE, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dest_indIEDest", dest_indIEDest, "N"): cReg = cReg + 1
    
    
    'Local de Entrega
    vReg(cReg) = Array("entr_CNPJ", entr_CNPJ, "S"): cReg = cReg + 1
    vReg(cReg) = Array("entr_xLgr", entr_xLgr, "S"): cReg = cReg + 1
    vReg(cReg) = Array("entr_nro", entr_nro, "S"): cReg = cReg + 1
    vReg(cReg) = Array("entr_xCpl", entr_xCpl, "S"): cReg = cReg + 1
    vReg(cReg) = Array("entr_xBairro", entr_xBairro, "S"): cReg = cReg + 1
    vReg(cReg) = Array("entr_cMun", entr_cMun, "S"): cReg = cReg + 1
    vReg(cReg) = Array("entr_xMun", entr_xMun, "S"): cReg = cReg + 1
    vReg(cReg) = Array("entr_UF", entr_UF, "S"): cReg = cReg + 1
    
    'Infoemacoes complementares
    vReg(cReg) = Array("infAdic_infCpl", infAdic_infCpl, "S"): cReg = cReg + 1
    'Transportador
    vReg(cReg) = Array("transp_modFrete", transp_modFrete, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_Pessoa", transp_Pessoa, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_CNPJ", transp_CNPJ, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_xNome", transp_xNome, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_IE", transp_IE, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_xEnder", transp_xEnder, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_xMun", transp_xMun, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_UF", transp_UF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_qVol", transp_qVol, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_esp", transp_esp, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_marca", transp_marca, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_nVol", transp_nVol, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_pesoL", transp_pesoL, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_pesoB", transp_pesoB, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_VeicPlaca", transp_VeicPlaca, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_VeicUF", transp_VeicUF, "S"): cReg = cReg + 1
    
    'TOTAIS
    vReg(cReg) = Array("total_vBC", total_vBC, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vICMS", total_vICMS, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vFCP", total_vFCP, "S"): cReg = cReg + 1
     'total_vBCST = ChkVal(CalcICMS_Total(10), 0, 2)
     'total_vICMSST = ChkVal(CalcICMS_Total(12), 0, 2)
    vReg(cReg) = Array("total_vBCST", total_vBCST, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("total_vFCPUFDest", total_vFCPUFDest, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vICMSUFDest", total_vICMSUFDest, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vICMSUFRemet", total_vICMSUFRemet, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("total_vICMSST", total_vICMSST, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vCredICMSSN", total_vCredICMSSN, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("total_vProd", total_vProd, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vFrete", total_vFrete, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vSeg", total_vSeg, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vDesc", total_vDesc, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vOutro", total_vOutro, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vIPI", total_vIPI, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vPIS", total_vPIS, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vCOFINS", total_vCOFINS, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vNF", total_vNF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ger_Vendedor", ger_Vendedor, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ger_idPV", ger_idPV, "N") ': cReg = cReg + 1

    
    
'     If IdReg = 0 Then
            IdReg = RegistroIncluir(strTabela, vReg, cReg)
            If IdReg = 0 Then
                    MsgBox "FormFaturamentoNFe: Erro ao incluir cabeçalho da nota fiscal!", vbInformation, App.EXEName
                    grvRegistro = False
                Else
                    grvRegistro = True
            End If
 '       Else
 '           If RegistroAlterar(strTabela, vReg, cReg, "Id = " & IdReg) = False Then
 '                   MsgBox "Erro ao Alterar."
 '                   grvRegistro = False
 '               Else
 '                   grvRegistro = True
 '           End If
 '   End If
    'DesbloquearTabela strTabela
 
 '****************** Descricao dos itens da NFe *******************************************************
 
    cReg = 0
    For i = 0 To cItens
        '****** Movimenta o estoque **************************************
        If PgDadosTpNotaFiscal(idTpNF).MovEstoque <> 0 Then
            If MovimentarEstoque(IIf(ide_tpNF = 0, "e", "s"), _
                                CInt(aItem(i)(0)), _
                                CDate(ide_dEmi), _
                                ide_nNF, _
                                CStr(aEstoque(i)(1)), _
                                CStr(aEstoque(i)(2)), _
                                CStr(aItem(i)(10)), _
                                "Unid.: " & aItem(i)(7) & "  Qtd.: " & aItem(i)(8) & " Vl.Unit.: " & ConvMoeda(CStr(aItem(i)(9))), _
                                 dest_xNome, _
                                 Id, dest_idDest, dest_CNPJ) = False Then
                MsgBox "Erro ao Movimentar Estoque com o item n. " & i
            End If
        End If
        '*************************************************************************************************
        vReg(cReg) = Array("IdNFe", Id, "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_IdProduto", aItem(i)(0), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_cProd", aItem(i)(1), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_cEAN", aItem(i)(2), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_xProd", aItem(i)(3), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_EXTIPI", aItem(i)(4), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_NCM", aItem(i)(5), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_CFOP", aItem(i)(6), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_uCom", aItem(i)(7), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_qCom", aItem(i)(8), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_vUnCom", aItem(i)(9), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_vProd", aItem(i)(10), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_cEANTrib", aItem(i)(11), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_uTrib", aItem(i)(12), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_qTrib", aItem(i)(13), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_vUnTrib", aItem(i)(14), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_vFrete", aItem(i)(15), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_vSeg", aItem(i)(16), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_vDesc", aItem(i)(17), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_vOutro", aItem(i)(18), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_indTot", aItem(i)(19), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_xPed", aItem(i)(20), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_nItemPed", aItem(i)(21), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_InfAdProd", aItem(i)(22), "S"): cReg = cReg + 1
        vReg(cReg) = Array("det_CEST", aItem(i)(23), "S"): cReg = cReg + 1
'*****************************************************************************************************************
        'Dados em conjunto com estoque
        vReg(cReg) = Array("estoque_Unid", aEstoque(i)(0), "S"): cReg = cReg + 1
        vReg(cReg) = Array("estoque_Qtd", ChkVal(CStr(aEstoque(i)(1)), 0, cDecQtd), "S"): cReg = cReg + 1
        vReg(cReg) = Array("estoque_vUnit", ChkVal(CStr(aEstoque(i)(2)), 0, cDecMoeda), "S"): cReg = cReg + 1

'*****************************************************************************************************************
        'ICMS
        vReg(cReg) = Array("ICMS_origem", aICMS(i)(0), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_CST", aICMS(i)(1), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_modBC", aICMS(i)(2), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_pRedBC", aICMS(i)(3), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_vBC", aICMS(i)(4), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_pICMS", aICMS(i)(5), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_vICMS", aICMS(i)(6), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_modBCST", aICMS(i)(7), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_pMVAST", aICMS(i)(8), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_pRedBCST", aICMS(i)(9), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_vBCST", aICMS(i)(10), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_pICMSST", aICMS(i)(11), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_vICMSST", aICMS(i)(12), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_pCredSN", aICMS(i)(13), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_vCredICMSSN", aICMS(i)(14), "S"): cReg = cReg + 1
        
        'FCP - 27/07/18
        vReg(cReg) = Array("ICMS_pFCP", aICMS(i)(15), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_vFCP", aICMS(i)(16), "S"): cReg = cReg + 1
        
        
'       'DIFAL - 17.12.2017
        vReg(cReg) = Array("ICMS_vBCUFDest", aIcmsDifal(i)(0), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_pFCPUFDest", aIcmsDifal(i)(1), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_pICMSUFDest", aIcmsDifal(i)(2), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_pICMSInter", aIcmsDifal(i)(3), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_pICMSInterPart", aIcmsDifal(i)(4), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_vFCPUFDest", aIcmsDifal(i)(5), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_vICMSUFDest", aIcmsDifal(i)(6), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_vICMSUFRemet", aIcmsDifal(i)(7), "S"): cReg = cReg + 1

'*****************************************************************************************************************
        'IPI
        vReg(cReg) = Array("IPI_cEnq", aIPI(i)(0), "S"): cReg = cReg + 1
        vReg(cReg) = Array("IPI_CST", aIPI(i)(1), "S"): cReg = cReg + 1
        vReg(cReg) = Array("IPI_vBC", aIPI(i)(2), "S"): cReg = cReg + 1
        vReg(cReg) = Array("IPI_pIPI", aIPI(i)(3), "S"): cReg = cReg + 1
        vReg(cReg) = Array("IPI_vIPI", aIPI(i)(4), "S"): cReg = cReg + 1
        
'*****************************************************************************************************************
        'PIS
        vReg(cReg) = Array("PIS_CST", aPIS(i)(0), "S"): cReg = cReg + 1
        vReg(cReg) = Array("PIS_vBC", aPIS(i)(1), "S"): cReg = cReg + 1
        vReg(cReg) = Array("PIS_pPIS", aPIS(i)(2), "S"): cReg = cReg + 1
        vReg(cReg) = Array("PIS_vPIS", aPIS(i)(3), "S"): cReg = cReg + 1
        
        
'*****************************************************************************************************************
        'COFINS
        vReg(cReg) = Array("COFINS_CST", aCOFINS(i)(0), "S"): cReg = cReg + 1
        vReg(cReg) = Array("COFINS_vBC", aCOFINS(i)(1), "S"): cReg = cReg + 1
        vReg(cReg) = Array("COFINS_pCOFINS", aCOFINS(i)(2), "S"): cReg = cReg + 1
        vReg(cReg) = Array("COFINS_vCOFINS", aCOFINS(i)(3), "S"): cReg = cReg + 1
        
'*****************************************************************************************************************
        'COMISSAO
        vReg(cReg) = Array("comissao_pComissao", aComissao(i)(0), "S"): cReg = cReg + 1
        vReg(cReg) = Array("comissao_vComissao", ChkVal(CStr(aComissao(i)(1)), 0, cDecMoeda), "S") ': cReg = cReg + 1

        IdReg = RegistroIncluir(strTabela & "Itens", vReg, cReg)
        If IdReg = 0 Then
                MsgBox "Erro ao Incluir COMISSAO do item"
                cReg = 0
                grvRegistro = False
            Else
                cReg = 0
                grvRegistro = True
        End If

    Next
    'cReg = cReg - 1
    
     
'******************* COBRANCA da NFe *******************************************************
    cReg = 0
    'nfat|vOrig|vDesc|vLiq|nDup|dVenc|vDup|TpDoc|idCliente|Emissao|Mora|Protesto|Multa
    For i = 0 To cCob
    
     If IsEmpty(aCob(i)(3)) = True Then
                aCob(i)(0) = Empty
                aCob(i)(4) = Empty
            Else
                aCob(i)(0) = IIf(Trim(aCob(i)(0)) = "", ide_nNF, aCob(i)(0))
                aCob(i)(4) = ide_nNF & "-" & aCob(i)(4)
        End If
    
        vReg(cReg) = Array("IdNFe", Id, "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_nFat", aCob(i)(0), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_vOrig", aCob(i)(1), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_vDesc", aCob(i)(2), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_vLiq", aCob(i)(3), "S"): cReg = cReg + 1
        
       
        vReg(cReg) = Array("cobr_nDup", aCob(i)(4), "S"): cReg = cReg + 1
        'vReg(cReg) = Array("cobr_nDup", ide_nNF & "-" & aCob(i)(4), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_dVenc", aCob(i)(5), "D"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_vDup", aCob(i)(6), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_TpDoc", aCob(i)(7), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_idCliente", aCob(i)(8), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_Emissao", aCob(i)(9), "D"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_Mora", aCob(i)(10), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_Protesto", aCob(i)(11), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_Multa", aCob(i)(12), "S") ': cReg = cReg + 1
        
        
        IdReg = RegistroIncluir(strTabela & "Cobranca", vReg, cReg)
        If IdReg = 0 Then
                MsgBox "Erro ao Incluir COBRANÇA"
                cReg = 0
                grvRegistro = False
            Else
                cReg = 0
                grvRegistro = True
        End If
        'Checa se movimenta contas a receber e pagar
        
        If PgDadosTpNotaFiscal(idTpNF).MovContasPR = 1 Then
            Call MovimentarContasPagarReceber("R", _
                                          CDate(ide_dEmi), _
                                          ide_nNF, _
                                          CStr(aCob(i)(1)), _
                                          "Clientes", _
                                          dest_idDest, _
                                          dest_xNome, _
                                          dest_CNPJ, _
                                          PgDadosTpNotaFiscal(idTpNF).conta, _
                                          IIf(Trim(PgDadosCliente(dest_idDest).CentroCustos) = "0", PgDadosTpNotaFiscal(idTpNF).CentroCusto, PgDadosCliente(dest_idDest).CentroCustos), _
                                          CStr(aCob(i)(7)), _
                                          IIf(Trim(PgDadosCliente(dest_idDest).PlanoContas) = "0", PgDadosTpNotaFiscal(idTpNF).PlanoContas, PgDadosCliente(dest_idDest).PlanoContas), "", _
                                          "", _
                                          CStr(aCob(i)(5)), _
                                          CStr(aCob(i)(4)), _
                                          pgDadosConta(PgDadosTpNotaFiscal(idTpNF).conta).Multa, _
                                          pgDadosConta(PgDadosTpNotaFiscal(idTpNF).conta).Juros, _
                                          pgDadosConta(PgDadosTpNotaFiscal(idTpNF).conta).banco, _
                                          "0", "0", "0", "0", _
                                          CStr(aCob(i)(6)), _
                                          "", _
                                          Id)
    End If
        
    Next
End Function
Private Function NumNotaFiscalExiste(NumeroNota As String) As Boolean
    'Checa se o numero da NF ja foi cadastrada
    'True= Nota Ja cadastrada
    'False = Nota Não Cadastrada
    Dim Rst     As Recordset
    Dim sSQL    As String
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_serie = '" & ide_serie & "'" & _
           " AND ide_nNF = '" & NumeroNota & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            NumNotaFiscalExiste = False
        Else
            NumNotaFiscalExiste = True
    End If
End Function
Private Sub AlterarVendedor()
    Dim vReg(1) As Variant
    Dim cReg    As Integer
    If ger_Vendedor = 0 Then Exit Sub
    If Trim(PgDadosCliente(dest_idDest).Vendedor) = "0" And Trim(dest_CNPJ) <> String(14, "9") Then
        If MsgBox("Deseja vincular o cliente " & dest_xNome & " ao vendedor " & PgDadosRhFuncionario(ger_Vendedor).Nome & "?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            cReg = 0
            vReg(cReg) = Array("Vendedor", ger_Vendedor, "N")
            RegistroAlterar "Clientes", vReg, cReg, "Id = " & dest_idDest
        End If
    End If
End Sub
Private Sub AlterarTransportadora()
    Dim vReg(1) As Variant
    Dim cReg    As Integer
    'If IdTransp = 0 Then Exit Sub
    If pgDadosTransportadora(Trim(PgDadosCliente(dest_idDest).Transportadora)).CNPJ <> transp_CNPJ And transp_CNPJ <> emit_CNPJ And transp_CNPJ <> dest_CNPJ Then
        If MsgBox("A transportadora (" & transp_xNome & ") difere do cadastro (" & pgDadosTransportadora(CInt(PgDadosCliente(dest_idDest).Transportadora)).Nome & ")." & vbCrLf & _
                  "Deseja registrar a " & transp_xNome & " como transportadora padrão? ", vbInformation + vbYesNo, "Aviso") = vbYes Then
            cReg = 0
            vReg(cReg) = Array("Transportadora", IdTransp, "N")
            RegistroAlterar "Clientes", vReg, cReg, "Id = " & dest_idDest
        End If
    End If
End Sub
Private Sub Calculo_ICMS_CST_00(Item As Integer)
    'aICMS
    '  0Origem|1CST|2ModBC|3pRedBC|4vBC|5pICMS|6vICMS|7modBCST|8pMVAST|
    '  9pRedBCST|10vBCST|11pICMSST|12vICMSST|13pCredSN|14vCredICMSSN|15pFCP|16vFCP
    '
    
    'Checa se o mat. e para: Mercadoria/Industrializacao ou Total do item/Consumo
    'If optMaterialPara(0).Value = True Then 'Valor de Mercadoria/Industrializacao
    If bcICMS = 0 Then 'Valor de Mercadoria/Industrializacao
            aICMS(Item)(4) = (Val(aItem(Item)(10)) - Val(ChkVal(CStr(aItem(Item)(17)), 0, 2))) * Val(1)
            aICMS(Item)(4) = ChkVal(CStr(aICMS(Item)(4)), 0, 2)
        Else 'Valor Total do Item/Consumo
                                                                    ''cEnq|CST|vBC|pIPI|vIPI
            aICMS(Item)(4) = Val(ChkVal(CStr(aItem(Item)(10)), 0, 2)) + Val(ChkVal(CStr(aIPI(Item)(4)), 0, 2))
            'aICMS(item)(4) = Val(ChkVal(aICMS(item)(4), 0, 2))
    End If
    
    aICMS(Item)(5) = ManutencaoICMS(Item, CStr(aICMS(Item)(5)))
    aICMS(Item)(6) = Val(ChkVal(CStr(aICMS(Item)(5)), 0, 2)) * Val(ChkVal(CStr(aICMS(Item)(4)), 0, 2)) / 100
    aICMS(Item)(6) = ChkVal(CStr(aICMS(Item)(6)), 0, 2)
    
 'FCP
    Calculo_ICMS_FCP (Item)
    
    'Calcula o FCP
    
'    If Len(Trim(aICMS(Item)(15))) <> 0 Then
'        aICMS(Item)(15) = ChkVal(CStr(aICMS(Item)(15)), 0, 2)
'
'        '06.01.19 - Alterado o calculo pois a petrobras notificou que a BCICMS para o FCP nao
'        '           tem resducao na base de calculo
'        '03.05.21 - Alterado novamente o calculo pois desta vez a Petrobras notificou que a
'        '           BCICMS para FCP incide na reducao da BC
'        'Calcula com base no valor da BCICMS
'        aICMS(Item)(16) = Val(ChkVal(CStr(aICMS(Item)(15)), 0, 2)) * Val(ChkVal(CStr(aICMS(Item)(4)), 0, 2)) / 100
'        'Calcula com base no valor TOTAL DA NOTA
'        'aICMS(Item)(16) = Val(ChkVal(CStr(aICMS(Item)(15)), 0, 2)) * Val(ChkVal(CStr(aItem(Item)(10)), 0, 2)) / 100
'
'        aICMS(Item)(16) = ChkVal(CStr(aICMS(Item)(16)), 0, 2)
'
'        'Informa nas obs do item o valor do FCP
'        aItem(Item)(22) = Trim(aItem(Item)(22)) & " [pFCP: " & aICMS(Item)(15) & "% vFCP: " & aICMS(Item)(16) & "]"
'    End If
    
 
End Sub
Private Sub Calculo_ICMS_CST_10(Item As Integer)

    
    '         0     1    2     3    4    5      6     7       8       9     10      11     12
    '      Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST|vICMSST
    
    'Checa se o mat. e para: Mercadoria/Industrializacao ou Total do item/Consumo
    'If optMaterialPara(0).Value = True Then 'Valor de Mercadoria/Industrializacao
    If bcICMS = 0 Then 'Valor de Mercadoria/Industrializacao
            aICMS(Item)(4) = Val(aItem(Item)(10)) * Val(1)
            aICMS(Item)(4) = ChkVal(CStr(aICMS(Item)(4)), 0, 2)
        Else 'Valor Total do Item/Consumo
                                                                    ''cEnq|CST|vBC|pIPI|vIPI
            aICMS(Item)(4) = Val(ChkVal(CStr(aItem(Item)(10)), 0, 2)) + Val(ChkVal(CStr(aIPI(Item)(4)), 0, 2))
            'aICMS(item)(4) = Val(ChkVal(aICMS(item)(4), 0, 2))
    End If
    
    aICMS(Item)(5) = ManutencaoICMS(Item, CStr(aICMS(Item)(5)))
    aICMS(Item)(6) = Val(ChkVal(CStr(aICMS(Item)(5)), 0, 2)) * Val(ChkVal(CStr(aICMS(Item)(4)), 0, 2)) / 100
    aICMS(Item)(6) = ChkVal(CStr(aICMS(Item)(6)), 0, 2)
    
    '###################################################################################################
    '### CALCULO DO ICMS-ST
    '### 09/01/2012
    '###################################################################################################
    'Calcula ICMSST
    Dim pICMSemit   As String
    Dim pICMSdest   As String
    Dim vProd       As String
    
    pICMSemit = pgDadosICMS(emit_UF, 0).ICMS
    pICMSdest = pgDadosICMS(PgDadosCliente(idCliente).uf, 0).ICMS
    
    
    vProd = ChkVal(Val(ChkVal(CStr(aItem(Item)(10)), 0, cDecMoeda)) + Val(ChkVal(CStr(aItem(Item)(15)), 0, cDecMoeda)) + Val(ChkVal(CStr(aItem(Item)(16)), 0, cDecMoeda)) + Val(ChkVal(CStr(aIPI(Item)(4)), 0, cDecMoeda)), 0, cDecMoeda) 'Produto + Frete + Seguro + IPI
    
    aICMS(Item)(7) = 0 'PgDadosTpNotaFiscal(idTpNF).ModBCST
    aICMS(Item)(8) = pgDadosEstoqueProduto(CInt(aItem(Item)(0))).MVA
    aICMS(Item)(9) = 0
            
    If Trim(aICMS(Item)(8)) = "" Then
        MsgBox "Erro ao localizar o MVA do item " & ZE(Item, 3) & ".", vbInformation, "Aviso"
        msgValid "Item " & ZE(Item, 3) & " - Erro ao localizar MVA"
        Exit Sub
    End If
    'alterar
    aICMS(Item)(11) = pICMSdest
    aICMS(Item)(10) = Calculo_ICMSST(emit_UF, dest_UF, CStr(aICMS(Item)(8)), vProd, CStr(aICMS(Item)(6))).vBCICMSST
    aICMS(Item)(12) = Calculo_ICMSST(emit_UF, dest_UF, CStr(aICMS(Item)(8)), vProd, CStr(aICMS(Item)(6))).vICMSST
    
    
    
End Sub

Private Sub Calculo_ICMS_CST_20(Item As Integer)
   'aICMS=Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST|vICMSST|pCredSN|vCredICMSSN
   'Dim pICMS As String
   '
   ' If pgAliqDifICMS(CStr(aItem(Item)(5)), dest_UF) = "" Then
   '         pICMS = pgDadosICMS(PgDadosCliente(dest_idDest).UF, 0).ICMS 'pgDadosICMS(PgIdICMS(PgDadosCliente(dest_idDest).UF)).ICMS
   '     Else
   '         'Mudar a aliquota de icms conforme ncm
   '         pICMS = pgAliqDifICMS(CStr(aItem(Item)(5)), dest_UF)
   ' End If
    
    'Checa se o mat. e para: Mercadoria/Industrializacao ou Total do item/Consumo
    'If optMaterialPara(0).Value = True Then 'Valor de Mercadoria/Industrializacao
    If bcICMS = 0 Then 'Valor de Mercadoria/Industrializacao
            aICMS(Item)(4) = Val(aItem(Item)(10)) * Val(1)
            aICMS(Item)(4) = ChkVal(CStr(aICMS(Item)(4)), 0, 2)
        Else 'Valor Total do Item/Consumo
            'cEnq|CST|vBC|pIPI|vIPI
            aICMS(Item)(4) = Val(ChkVal(CStr(aItem(Item)(10)), 0, 2)) + Val(ChkVal(CStr(aIPI(Item)(4)), 0, 2))
            'aICMS(item)(4) = Val(ChkVal(aICMS(item)(4), 0, 2))
    End If
    
    aICMS(Item)(5) = ManutencaoICMS(Item, CStr(aICMS(Item)(5)))
    aICMS(Item)(6) = Val(ChkVal(CStr(aICMS(Item)(5)), 0, 2)) * Val(ChkVal(CStr(aICMS(Item)(4)), 0, 2)) / 100
    aICMS(Item)(6) = ChkVal(CStr(aICMS(Item)(6)), 0, 2)
    'FCP
    Calculo_ICMS_FCP (Item)
    
End Sub

Private Sub Calculo_ICMS_FCP(Item As Integer)
     'Calcula o FCP
    
    If Len(Trim(aICMS(Item)(15))) <> 0 Then
        aICMS(Item)(15) = ChkVal(CStr(aICMS(Item)(15)), 0, 2)
        
        '06.01.19 - Alterado o calculo pois a petrobras notificou que a BCICMS para o FCP nao
        '           tem resducao na base de calculo
        '03.05.21 - Alterado novamente o calculo pois desta vez a Petrobras notificou que a
        '           BCICMS para FCP incide na reducao da BC
        'Calcula com base no valor da BCICMS
        aICMS(Item)(16) = Val(ChkVal(CStr(aICMS(Item)(15)), 0, 2)) * Val(ChkVal(CStr(aICMS(Item)(4)), 0, 2)) / 100
        'Calcula com base no valor TOTAL DA NOTA
        'aICMS(Item)(16) = Val(ChkVal(CStr(aICMS(Item)(15)), 0, 2)) * Val(ChkVal(CStr(aItem(Item)(10)), 0, 2)) / 100
        
        aICMS(Item)(16) = ChkVal(CStr(aICMS(Item)(16)), 0, 2)
        
        'Informa nas obs do item o valor do FCP
        aItem(Item)(22) = Trim(aItem(Item)(22)) & " [pFCP: " & aICMS(Item)(15) & "% vFCP: " & aICMS(Item)(16) & "]"
    End If
End Sub

Private Function ManutencaoICMS(Item As Integer, icmsPV As String) As String
    '###################################################################
    '### RETORNA A ALIQUOTA DE ICMS QUE SERA USADA NO ITEM
    '### 31/10/2011
    '###################################################################
    Dim xItem       As String
    Dim pICMS       As String
    
    Dim indice      As String
    Dim bcICMSAnt   As String 'armazena a base de calculo sem alteração
    
    
    icmsPV = Replace(icmsPV, "%", "")
    
    '26/07/18 - Pega Aliquota de ICMS
    If Trim(dest_UF) <> Trim(PgDadosEmpresa(ID_Empresa).uf) Then
            'Pega o ICMS Interestadual
            pICMS = pgDadosICMS(dest_UF, 0).ICMS
        Else
            'Pega o ICMS Interno
            pICMS = pgDadosICMS(dest_UF, 0).ICMSInt
    End If
    
    
    'VERIFICA SE HA DIFERIMENTO NA ALIQUOTA DE ICMS ****************************************
    If pgAliqDifICMS(CStr(aItem(Item)(5)), dest_UF) = "" Then
            'Aliquota normal do Estado
            'pICMS = pgDadosICMS(PgDadosCliente(dest_idDest).uf, 0).ICMS
        Else
            'Mudar a aliquota de icms conforme NCM
            pICMS = pgAliqDifICMS(CStr(aItem(Item)(5)), dest_UF)
            '############################################################################################################
            '###      MUDA A BCICMS NO LUGAR DA ALIQUOTA DE ICMS
            '############################################################################################################
                
                 bcICMSAnt = ChkVal(CStr(aICMS(Item)(4)), 0, cDecMoeda)
                'Alterado em 04.05.2021
                indice = (Val(bcICMSAnt) * Val(pICMS)) / 100
                indice = Val(ChkVal(indice, 0, cDecMoeda)) / (Val(pgDadosICMS(dest_UF, 0).ICMS) + Val(pgDadosICMS(dest_UF, 0).ICMSFECP)) * 100
                aICMS(Item)(4) = ChkVal(CStr(indice), 0, cDecMoeda)
                pICMS = pgDadosICMS(PgDadosCliente(dest_idDest).uf, 0).ICMS
                
                'xItem = Item + 1
                xItem = ZE((Item + 1), 3)
                MsgBox "A BCICMS informado no ITEM " & xItem & " sofrerá redução conforme DECRETO N.º 28.494 DE 31 DE MAIO DE 2001.", vbInformation, "Aviso"
                msgValid "item " & xItem & ": Alteração da BCICMS de:" & ChkVal(bcICMSAnt, 0, cDecMoeda) & " para " & aICMS(Item)(4)
            '############################################################################################################
            'Incluindo msg nas informações complementares
            txtObs.Text = txtObs.Text & " // A Base de calculo do ICMS informado no ITEM " & xItem & _
                             " sofrerá redução conforme DECRETO N.º 28.494 DE 31 DE MAIO DE 2001."
            
            '############################################################################################################
    End If
    '**************************************************************************************

    
    
    If icmsPV <> pICMS Then
        xItem = Item + 1
        xItem = Left("000", 3 - Len(Trim(xItem))) & xItem
       If MsgBox("O ICMS informado no ITEM " & xItem & " da proposta diverge do ICMS calculado pelo sistema." & vbCrLf & vbCrLf & _
                "ICMS calculado: " & pICMS & vbCrLf & _
                "ICMS da Pre-Venda: " & icmsPV & vbCrLf & vbCrLf & _
                "Deseja ajustar a aliquota ICMS de acordo com o calculo do sistema para emissão do Documento Fiscal?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        
                    
                    icmsPV = IIf(Trim(icmsPV) = "", 0, icmsPV)
                    msgValid "Item " & xItem & ": Alteração do ICMS de " & icmsPV & " para " & pICMS
            Else
                pICMS = icmsPV
        End If
    End If
    '####################################################################################################
    '### 13/07/2012 - Recurso criado para facilitar o texto de
    '###              redução da BCICMS.
    '###
    If pICMS = "13" Then
        If MsgBox("Deseja incluir a mensagem abaixo na NFe?" & vbCrLf & _
                    "REDUCAO NA BC DO ICMS COM ALIQUOTA 13% CONFORME DECRETO 28.494 DE 31/05/2001." & vbCrLf _
                    , vbQuestion + vbYesNo, "Aviso") = vbYes Then
            msgRedICMS = "REDUCAO NA BC DO ICMS COM ALIQUOTA 13% CONFORME DECRETO 28.494 DE 31/05/2001."
        End If
    End If
        '####################################################################################################
    ManutencaoICMS = pICMS
End Function
Private Function Calculo_ICMS_CST_60(Item As Integer) As Boolean
    Dim vProd       As String 'Valor do Produto acrescido de impostos
    Dim pICMSemit   As String
    Dim pICMSdest   As String
    Dim BCST        As String
    
    pICMSemit = pgDadosICMS(emit_UF, 0).ICMS
    pICMSdest = pgDadosICMS(PgDadosCliente(idCliente).uf, 0).ICMS
    
    
    vProd = ChkVal(Val(ChkVal(CStr(aItem(Item)(10)), 0, cDecMoeda)) + Val(ChkVal(CStr(aItem(Item)(15)), 0, cDecMoeda)) + Val(ChkVal(CStr(aItem(Item)(16)), 0, cDecMoeda)) + Val(ChkVal(CStr(aIPI(Item)(4)), 0, cDecMoeda)), 0, cDecMoeda) 'Produto + Frete + Seguro + IPI
    
    If PgDadosCFOP(idTpNF, "60", dest_UF).ICMSST = 0 Then
            'Não calcula ICMSST
            
            aICMS(Item)(5) = 0 'ManutencaoICMS(Item, CStr(aICMS(Item)(5)))
            aICMS(Item)(7) = 0 'PgDadosTpNotaFiscal(idTpNF).ModBCST
            aICMS(Item)(8) = 0 'pgDadosEstoqueProduto(idProduto).MVA
            aICMS(Item)(9) = 0
            aICMS(Item)(10) = 0
            aICMS(Item)(11) = 0
            aICMS(Item)(12) = 0
        Else
            'Calcula ICMSST
            aICMS(Item)(7) = 0 'PgDadosTpNotaFiscal(idTpNF).ModBCST
            aICMS(Item)(8) = pgDadosEstoqueProduto(CInt(aItem(Item)(0))).MVA
            aICMS(Item)(9) = 0
            
            If Trim(aICMS(Item)(8)) = "" Then
                MsgBox "Erro ao localizar o MVA do item " & ZE(Item, 3) & ".", vbInformation, "Aviso"
                msgValid "Item " & ZE(Item, 3) & " - Erro ao localizar MVA"
                Calculo_ICMS_CST_60 = False
                Exit Function
            End If
            'vBCST
            'BCST = ChkVal((Val(vProd) * Val(ChkVal(CStr(aICMS(Item)(8)), 0, 3)) / 100), 0, cDecMoeda)
            
            'aICMS(Item)(10) = ChkVal(Val(BCST) + Val(vProd), 0, cDecMoeda)
            
            'pICMS
            aICMS(Item)(11) = pICMSdest
            'vICMSST
            'aICMS(Item)(12) = Val(ChkVal(Val(pICMSdest) * Val(aICMS(Item)(10)) / 100, 0, cDecMoeda)) - Val(ChkVal(Val(pICMSemit) * Val(aItem(Item)(10)) / 100, 0, cDecMoeda))
            'aICMS(Item)(12) = ChkVal(CStr(aICMS(Item)(12)), 0, cDecMoeda)
            
            aICMS(Item)(10) = Calculo_ICMSST(emit_UF, dest_UF, CStr(aICMS(Item)(8)), vProd, CStr(aICMS(Item)(6))).vBCICMSST
            aICMS(Item)(12) = Calculo_ICMSST(emit_UF, dest_UF, CStr(aICMS(Item)(8)), vProd, CStr(aICMS(Item)(6))).vICMSST
    End If
    
    '31.07.18 - NFe 4.0 - Inclui o codigo CEST no produto
    
    aItem(Item)(23) = PgDadosNCM("ncm", CStr(aItem(Item)(5)), "S").cest
    
    Calculo_ICMS_CST_60 = True
'aICMS=0-Origem|1-CST|2-ModBC|3-pRedBC|4-vBC|5-pICMS|6-vICMS|
'      7-modBCST|8-pMVAST|9-pRedBCST|10-vBCST|11-pICMSST|12-vICMSST

'----------------------------------------------------------------------------------
    'id_intProd|det_cProd|det_cEAN|det_xProd|EXTIPI|det_NCM|det_CFOP|det_uCom|
    'det_qCom|det_vUnCom|det_vProd|det_cEANTrib|det_uTrib|det_qTrib|
    'det_vUnTrib|
    'det_vFrete|det_vSeg|det_vDesc|det_vOutro|det_indTot|xPed|nItemPed
    'det_indTot = 0 = O valor do item compoe a NF / 1  = O valor do item nao compoe a NF
End Function
Private Sub Calculo_ICMS_CST_40(Item As Integer)
    '0-Origem|1-CST|2-ModBC|3-pRedBC|4-vBC|5-pICMS|6-vICMS|7-modBCST|8pMVAST|9-pRedBCST|10-vBCST|11-pICMSST|12-vICMSST
    aICMS(Item)(2) = 0
    aICMS(Item)(3) = 0
    aICMS(Item)(4) = 0
    aICMS(Item)(5) = 0
    aICMS(Item)(6) = 0
            
    aICMS(Item)(7) = 0
    aICMS(Item)(8) = 0
    aICMS(Item)(9) = 0
    aICMS(Item)(10) = 0
    aICMS(Item)(11) = 0
    aICMS(Item)(12) = 0
    
End Sub
Private Sub Calculo_ICMS_CST_41(Item As Integer)
    '0-Origem|1-CST|2-ModBC|3-pRedBC|4-vBC|5-pICMS|6-vICMS|7-modBCST|8pMVAST|9-pRedBCST|10-vBCST|11-pICMSST|12-vICMSST
    aICMS(Item)(2) = 0
    aICMS(Item)(3) = 0
    aICMS(Item)(4) = 0
    aICMS(Item)(5) = 0
    aICMS(Item)(6) = 0
            
    aICMS(Item)(7) = 0
    aICMS(Item)(8) = 0
    aICMS(Item)(9) = 0
    aICMS(Item)(10) = 0
    aICMS(Item)(11) = 0
    aICMS(Item)(12) = 0
    
End Sub

Private Sub Calculo_ICMS_CST_50(Item As Integer)
    '0-Origem|1-CST|2-ModBC|3-pRedBC|4-vBC|5-pICMS|6-vICMS|7-modBCST|8pMVAST|9-pRedBCST|10-vBCST|11-pICMSST|12-vICMSST
    aICMS(Item)(2) = 0
    aICMS(Item)(3) = 0
    aICMS(Item)(4) = 0
    aICMS(Item)(5) = 0
    aICMS(Item)(6) = 0
            
    aICMS(Item)(7) = 0
    aICMS(Item)(8) = 0
    aICMS(Item)(9) = 0
    aICMS(Item)(10) = 0
    aICMS(Item)(11) = 0
    aICMS(Item)(12) = 0
    
End Sub

Private Sub Calculo_ICMS_CST_51(Item As Integer)
    'Dim pICMS As String
    '0-Origem|1-CST|2-ModBC|3-pRedBC|4-vBC|5-pICMS|6-vICMS|7-modBCST|8pMVAST|9-pRedBCST|10-vBCST|11-pICMSST|12-vICMSST
            
    
    If PgDadosCFOP(idTpNF, CStr(aICMS(Item)(1)), dest_UF).ICMS = 0 Then
            aICMS(Item)(2) = 0
            aICMS(Item)(3) = 0
            aICMS(Item)(4) = 0
            aICMS(Item)(5) = 0
            aICMS(Item)(6) = 0
            
            aICMS(Item)(7) = 0
            aICMS(Item)(8) = 0
            aICMS(Item)(9) = 0
            aICMS(Item)(10) = 0
            aICMS(Item)(11) = 0
            aICMS(Item)(12) = 0
        Else
            'If pgAliqDifICMS(CStr(aItem(Item)(5)), dest_UF) = "" Then
            '        pICMS = pgDadosICMS(PgDadosCliente(dest_idDest).UF, 0).ICMS 'pgDadosICMS(PgIdICMS(PgDadosCliente(dest_idDest).UF)).ICMS
            '    Else
            '        'Mudar a aliquota de icms conforme ncm
            '        pICMS = pgAliqDifICMS(CStr(aItem(Item)(5)), dest_UF)
            'End If
    
            'If optMaterialPara(0).Value = True Then 'Valor de Mercadoria/Industrializacao
            If bcICMS = 0 Then 'Valor de Mercadoria/Industrializacao
                    aICMS(Item)(4) = Val(aItem(Item)(10)) * Val(1)
                    aICMS(Item)(4) = ChkVal(CStr(aICMS(Item)(4)), 0, 2)
                Else 'Valor Total do Item/Consumo
                                                                    ''cEnq|CST|vBC|pIPI|vIPI
                    aICMS(Item)(4) = Val(ChkVal(CStr(aItem(Item)(10)), 0, 2)) + Val(ChkVal(CStr(aIPI(Item)(4)), 0, 2))
            
            End If
            
            aICMS(Item)(5) = ManutencaoICMS(Item, CStr(aICMS(Item)(5)))
            aICMS(Item)(6) = Val(ChkVal(CStr(aICMS(Item)(5)), 0, 2)) * Val(ChkVal(CStr(aICMS(Item)(4)), 0, 2)) / 100
            aICMS(Item)(6) = ChkVal(CStr(aICMS(Item)(6)), 0, 2)

    End If
End Sub
Private Sub Calculo_ICMS_CSOSN_101(Item As Integer)
    'aICMS=Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST|vICMSST|pCredSN|vCredICMSSN
    
    'Checa se o mat. e para: Mercadoria/Industrializacao ou Total do item/Consumo
    
    'If optMaterialPara(0).Value = True Then 'Valor de Mercadoria/Industrializacao
    If bcICMS = 0 Then 'Valor de Mercadoria/Industrializacao
            aICMS(Item)(4) = (Val(aItem(Item)(10)) - Val(ChkVal(CStr(aItem(Item)(17)), 0, 2))) * Val(1)
            aICMS(Item)(4) = ChkVal(CStr(aICMS(Item)(4)), 0, 2)
        Else 'Valor Total do Item/Consumo
            aICMS(Item)(4) = Val(ChkVal(CStr(aItem(Item)(10)), 0, 2)) + Val(ChkVal(CStr(aIPI(Item)(4)), 0, 2))
            'aICMS(item)(4) = Val(ChkVal(aICMS(item)(4), 0, 2))
    End If
    
    aICMS(Item)(13) = aICMS(Item)(5)
    aICMS(Item)(14) = Val(ChkVal(CStr(aICMS(Item)(5)), 0, 2)) * Val(ChkVal(CStr(aICMS(Item)(4)), 0, 2)) / 100
    aICMS(Item)(14) = ChkVal(CStr(aICMS(Item)(14)), 0, 2)
    
    pCredICMS = aICMS(Item)(5)
    
    aICMS(Item)(4) = "0.00"
    aICMS(Item)(5) = "0.00" 'ManutencaoICMS(Item, CStr(aICMS(Item)(5)))
    aICMS(Item)(6) = "0.00" 'Val(ChkVal(CStr(aICMS(Item)(5)), 0, 2)) * Val(ChkVal(CStr(aICMS(Item)(4)), 0, 2)) / 100
    
End Sub
Public Sub Calculo_ICMS_DIFAL(Item As Integer)
    'Verifica se a empresa devera calcular o DIFAL
    
    'Empresa Nao optante do simples
'    If emit_CRT <> 3 Then
'       Exit Sub
'    End If
    'Consumidor final
    If bcICMS = 0 Then '0= Vl Mercadoria / 1=vl Total da NF
        Exit Sub
    End If
    'Outro Estado
    If emit_UF = dest_UF Then
        Exit Sub
    End If
    'Destinatario com IE nao calcula o DIFAL
    If Len(Trim(dest_IE)) <> 0 Then
        Exit Sub
    End If


'21/07/17 - Calcula a DIFAL (diferenca de aliquota de icms entre os estado)
    Dim bcDIFAL As String 'Valor da base de calculo
    
    Dim vBCUFDest As String
    Dim pFCPUFDest As String
    Dim pICMSUFDest As String
    Dim pICMSInter As String
    Dim pICMSInterPart As String
    Dim vFCPUFDest As String
    Dim vICMSUFDest As String
    Dim vICMSUFRemet As String

    'vBCUFDest|pFCPUFDest|pICMSUFDest|pICMSInter|pICMSInterPart|vFCPUFDest|vICMSUFDest|vICMSUFRemet|
    
    Dim vBCICMS As String
    Dim vDIFAL As String
    
    
    'vBCIcms - é a soma de todos os itens como IPI desc acresc frete etc
    vBCICMS = aICMS(Item)(4)
    
    'Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST|vICMSST|pCredSN|vCredICMSSN
                        
    
    Dim pICMSPartRemet As String
    pFCPUFDest = "2.00"
    pICMSUFDest = "18.00"
    pICMSInter = aICMS(Item)(5)
    
    pICMSInterPart = "100"
    pICMSPartRemet = "0"
    
    
    vFCPUFDest = Val(ChkVal(vBCICMS, 0, cDecMoeda)) * Val(ChkVal(pFCPUFDest, 0, cDecMoeda)) / 100
    
    'Valor do DIFAL
    vDIFAL = (Val(ChkVal(vBCICMS, 0, cDecMoeda)) * (Val(ChkVal(pICMSUFDest, 0, cDecMoeda)) - Val(ChkVal(pICMSInter, 0, cDecMoeda))) / 100)
    
    
    'vICMSUFDest = Val(ChkVal(vDIFAL, 0, cDecMoeda)) * (Val(ChkVal(pICMSUFDest, 0, cDecMoeda)) - Val(ChkVal(pICMSInter, 0, cDecMoeda)))
    vICMSUFDest = Val(ChkVal(vDIFAL, 0, cDecMoeda)) * (Val(ChkVal(pICMSInterPart, 0, cDecMoeda)) / 100)
    vICMSUFDest = ChkVal(vICMSUFDest, 0, cDecMoeda)
    
    
    
    vICMSUFRemet = Val(ChkVal(vDIFAL, 0, cDecMoeda)) * (Val(ChkVal(pICMSPartRemet, 0, cDecMoeda)) / 100)
    'vICMSUFRemet = Val(ChkVal(vICMSUFRemet, 0, cDecMoeda)) * (100 - Val(ChkVal(pICMSInterPart, 0, cDecMoeda))) / 100
    vICMSUFRemet = ChkVal(vICMSUFRemet, 0, 2)
    
    'Altera o a inficacao para consumidor final
    ide_indFinal = "1"
    
    dest_indIEDest = "9"
    
    vBCUFDest = vBCICMS
   
    aIcmsDifal(Item)(0) = ChkVal(vBCUFDest, 0, cDecMoeda)
    aIcmsDifal(Item)(1) = ChkVal(pFCPUFDest, 0, cDecMoeda)
    aIcmsDifal(Item)(2) = ChkVal(pICMSUFDest, 0, cDecMoeda)
    aIcmsDifal(Item)(3) = ChkVal(pICMSInter, 0, cDecMoeda)
    aIcmsDifal(Item)(4) = ChkVal(pICMSInterPart, 0, cDecMoeda)
    aIcmsDifal(Item)(5) = ChkVal(vFCPUFDest, 0, cDecMoeda)
    aIcmsDifal(Item)(6) = ChkVal(vICMSUFDest, 0, cDecMoeda)
    aIcmsDifal(Item)(7) = ChkVal(vICMSUFRemet, 0, cDecMoeda)
 
    'Soma os totalizadores
'    total_vFCPUFDest = Val(ChkVal(total_vFCPUFDest, 0, cDecMoeda)) + Val(ChkVal(vFCPUFDest, 0, cDecMoeda))
'    total_vFCPUFDest = ChkVal(total_vFCPUFDest, 0, cDecMoeda)
'    total_vICMSUFDest = Val(ChkVal(total_vICMSUFDest, 0, cDecMoeda)) + Val(ChkVal(vICMSUFDest, 0, cDecMoeda))
'    total_vICMSUFDest = ChkVal(total_vICMSUFDest, 0, cDecMoeda)
'    total_vICMSUFRemet = Val(ChkVal(total_vICMSUFRemet, 0, cDecMoeda)) + Val(ChkVal(vICMSUFRemet, 0, cDecMoeda))
'    total_vICMSUFRemet = ChkVal(total_vICMSUFRemet, 0, cDecMoeda)



End Sub
