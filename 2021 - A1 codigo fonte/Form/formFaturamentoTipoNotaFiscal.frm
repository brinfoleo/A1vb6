VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFaturamentoTipoNotaFiscal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Tipo de Nota Fiscal"
   ClientHeight    =   10755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10755
   ScaleWidth      =   11130
   Begin VB.Frame Frame1 
      Height          =   10035
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   10995
      Begin VB.Frame Frame10 
         Caption         =   "Movimenta Fisco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2280
         TabIndex        =   74
         Top             =   9300
         Width           =   2655
         Begin VB.CheckBox chkMovFisco 
            Caption         =   "Movimentar Fisco"
            Height          =   195
            Left            =   180
            TabIndex        =   75
            Top             =   300
            Width           =   2295
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Modalidade da Base de Calculo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   5040
         TabIndex        =   65
         Top             =   8580
         Width           =   5775
         Begin VB.ComboBox cboModBCST 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   780
            Width           =   4515
         End
         Begin VB.ComboBox cboModBC 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   360
            Width           =   4515
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "ICMS-ST:"
            Height          =   195
            Left            =   180
            TabIndex        =   68
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "ICMS:"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   420
            Width           =   675
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Imprimir na Nota Fiscal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   120
         TabIndex        =   50
         Top             =   5640
         Width           =   4815
         Begin VB.CheckBox chkImprCampoFatura 
            Caption         =   "Fatura/Duplicata da Nota Fiscal"
            Height          =   195
            Left            =   180
            TabIndex        =   63
            Top             =   300
            Width           =   4275
         End
         Begin VB.CheckBox chkImprDataSaida 
            Caption         =   "Data/hora de saída"
            Height          =   195
            Left            =   180
            TabIndex        =   62
            Top             =   555
            Width           =   4275
         End
         Begin VB.CheckBox chkImpvTotalNota 
            Caption         =   "Valor Total da Nota"
            Height          =   195
            Left            =   180
            TabIndex        =   61
            Top             =   3360
            Width           =   4275
         End
         Begin VB.CheckBox chkImpvIPI 
            Caption         =   "Valor do IPI"
            Height          =   195
            Left            =   180
            TabIndex        =   60
            Top             =   3105
            Width           =   4275
         End
         Begin VB.CheckBox chkImpvOutrasDesp 
            Caption         =   "Outras Despesas Acessorias"
            Height          =   195
            Left            =   180
            TabIndex        =   59
            Top             =   2850
            Width           =   4275
         End
         Begin VB.CheckBox chkImpvDesconto 
            Caption         =   "Valor do Desconto"
            Height          =   195
            Left            =   180
            TabIndex        =   58
            Top             =   2595
            Width           =   4275
         End
         Begin VB.CheckBox chkImpvSeguro 
            Caption         =   "Valor do Seguro"
            Height          =   195
            Left            =   180
            TabIndex        =   57
            Top             =   2340
            Width           =   4275
         End
         Begin VB.CheckBox chkImpvFrete 
            Caption         =   "Valor do Frete"
            Height          =   195
            Left            =   180
            TabIndex        =   56
            Top             =   2085
            Width           =   4275
         End
         Begin VB.CheckBox chkImpvTotalProduto 
            Caption         =   "Valor Total dos Produtos"
            Height          =   195
            Left            =   180
            TabIndex        =   55
            Top             =   1830
            Width           =   4275
         End
         Begin VB.CheckBox chkImpvICMSST 
            Caption         =   "Valor do ICMS ST"
            Height          =   195
            Left            =   180
            TabIndex        =   54
            Top             =   1575
            Width           =   4275
         End
         Begin VB.CheckBox chkImpBCICMSST 
            Caption         =   "Base de Calculo do ICMS ST"
            Height          =   195
            Left            =   180
            TabIndex        =   53
            Top             =   1320
            Width           =   4275
         End
         Begin VB.CheckBox chkImpvICMS 
            Caption         =   "Valor do ICMS"
            Height          =   195
            Left            =   180
            TabIndex        =   52
            Top             =   1065
            Width           =   4275
         End
         Begin VB.CheckBox chkImpBCICMS 
            Caption         =   "Base de Calculo do ICMS"
            Height          =   195
            Left            =   180
            TabIndex        =   51
            Top             =   810
            Width           =   4275
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "PIS / COFINS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5040
         TabIndex        =   45
         Top             =   7440
         Width           =   5775
         Begin VB.ComboBox cboCSTCOFINS 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   660
            Width           =   4815
         End
         Begin VB.ComboBox cboCSTPIS 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "COFINS:"
            Height          =   255
            Left            =   60
            TabIndex        =   47
            Top             =   660
            Width           =   675
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "PIS:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Financeiro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   5040
         TabIndex        =   37
         Top             =   120
         Width           =   5775
         Begin VB.ComboBox cboPlanoContas 
            Height          =   315
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   2040
            Width           =   3915
         End
         Begin VB.CheckBox chkMovComissao 
            Caption         =   "Movimenta comissão"
            Height          =   195
            Left            =   1740
            TabIndex        =   64
            Top             =   240
            Width           =   3615
         End
         Begin VB.CheckBox chkMovContasPR 
            Caption         =   "Movimentar Contas a Pagar/Receber"
            Height          =   255
            Left            =   1740
            TabIndex        =   44
            Top             =   540
            Width           =   3795
         End
         Begin VB.ComboBox cboCentroCusto 
            Height          =   315
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1680
            Width           =   3915
         End
         Begin VB.ComboBox cboTpDocumento 
            Height          =   315
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   1320
            Width           =   3915
         End
         Begin VB.ComboBox cboConta 
            Height          =   315
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   960
            Width           =   3915
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Plano de Contas:"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   2100
            Width           =   1515
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Centro de Custo:"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   1740
            Width           =   1455
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo de Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   1380
            Width           =   1575
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Conta:"
            Height          =   195
            Left            =   840
            TabIndex        =   38
            Top             =   1020
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Nota Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   4815
         Begin VB.CheckBox chkChaveAcessoRef 
            Caption         =   "Solicitar Chave Acesso Referencia"
            Height          =   195
            Left            =   1800
            TabIndex        =   73
            Top             =   3120
            Width           =   2835
         End
         Begin VB.CheckBox chkEnvioRF 
            Caption         =   "Enviar XML para Receira Federal"
            Height          =   255
            Left            =   1800
            TabIndex        =   70
            Top             =   3360
            Width           =   2715
         End
         Begin VB.ComboBox cboTipoNota 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   930
            Width           =   2835
         End
         Begin VB.TextBox txtSerie 
            Height          =   285
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   1290
            Width           =   2835
         End
         Begin VB.TextBox txtNumInicial 
            Height          =   285
            Left            =   1800
            MaxLength       =   9
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   1980
            Width           =   2835
         End
         Begin VB.ComboBox cboNaturezaOperacao 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   2325
            Width           =   2835
         End
         Begin VB.TextBox txtID 
            Height          =   285
            Left            =   1800
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   1800
            MaxLength       =   120
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   585
            Width           =   2475
         End
         Begin VB.ComboBox cboFinalidade 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   2700
            Width           =   2835
         End
         Begin VB.TextBox txtModelo 
            Height          =   285
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   1635
            Width           =   1515
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo de Nota:"
            Height          =   195
            Left            =   720
            TabIndex        =   36
            Top             =   1020
            Width           =   1035
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Serie:"
            Height          =   195
            Left            =   1080
            TabIndex        =   35
            Top             =   1320
            Width           =   675
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Número Inicial:"
            Height          =   195
            Left            =   660
            TabIndex        =   34
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Natureza da Operação:"
            Height          =   255
            Left            =   60
            TabIndex        =   33
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   720
            TabIndex        =   32
            Top             =   660
            Width           =   1035
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "ID:"
            Height          =   195
            Left            =   1140
            TabIndex        =   31
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Finalidade:"
            Height          =   195
            Left            =   900
            TabIndex        =   30
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Modelo:"
            Height          =   195
            Left            =   1020
            TabIndex        =   29
            Top             =   1680
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "CFOP/Tributação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   5040
         TabIndex        =   7
         Top             =   2700
         Width           =   5775
         Begin VB.CommandButton btoExcluir 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   4680
            TabIndex        =   19
            Top             =   4200
            Width           =   975
         End
         Begin VB.CommandButton btoIncluir 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   4680
            TabIndex        =   18
            Top             =   3840
            Width           =   975
         End
         Begin VB.CheckBox chkCalcICMS 
            Caption         =   "Calcular o Valor do ICMS"
            Height          =   195
            Left            =   1140
            TabIndex        =   16
            Top             =   4020
            Width           =   2355
         End
         Begin VB.CheckBox chkCalcICMSST 
            Caption         =   "Calcular o Valor do ICMS-ST"
            Height          =   255
            Left            =   1140
            TabIndex        =   15
            Top             =   4260
            Width           =   2655
         End
         Begin MSFlexGridLib.MSFlexGrid msfgCFOP 
            Height          =   2475
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   4366
            _Version        =   393216
            Cols            =   6
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   "ID |<Situação           |<CST         |^CFOP     |^Calc. ICMS |^Calc. ICMS-ST"
         End
         Begin VB.TextBox txtCFOP 
            Height          =   285
            Left            =   1140
            MaxLength       =   4
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   3600
            Width           =   675
         End
         Begin VB.ComboBox cboICMSCST 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   3240
            Width           =   4335
         End
         Begin VB.ComboBox cboSituacao 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2820
            Width           =   2535
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "ICMS:"
            Height          =   255
            Left            =   540
            TabIndex        =   17
            Top             =   4140
            Width           =   495
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "CFOP:"
            Height          =   195
            Left            =   300
            TabIndex        =   10
            Top             =   3660
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Cod. CST:"
            Height          =   195
            Left            =   300
            TabIndex        =   9
            Top             =   3300
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Situação:"
            Height          =   255
            Left            =   300
            TabIndex        =   8
            Top             =   2880
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Estoque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   5
         Top             =   9300
         Width           =   2115
         Begin VB.CheckBox chkMovEstoque 
            Caption         =   "Movimentar Estoque"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   300
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Observação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   120
         TabIndex        =   2
         Top             =   3840
         Width           =   4815
         Begin VB.TextBox txtObs 
            Height          =   1095
            Left            =   180
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Text            =   "formFaturamentoTipoNotaFiscal.frx":0000
            Top             =   540
            Width           =   4515
         End
         Begin VB.CheckBox chkImprCampoObs 
            Caption         =   "Imprimir no campo INFORMAÇÕES COMPLEMENTARES"
            Height          =   195
            Left            =   180
            TabIndex        =   3
            Top             =   240
            Width           =   4395
         End
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alterar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clonar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pesquisar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Manutenção da Tabela"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
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
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":0006
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":0458
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":0772
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":1004
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":2256
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":2B30
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":33C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":3C54
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":4EA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":51C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":54DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoTipoNotaFiscal.frx":58D1
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFaturamentoTipoNotaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg       As Integer
Dim strTabela   As String
Dim idCFOP      As Integer

Private Sub btoExcluir_Click()
    If MsgBox("Deseja realmente remover este item?", vbYesNo, "Removendo Item do Pedido") = vbYes Then
        If msfgCFOP.Rows = 2 Then
                msfgCFOP.Rows = 1
            Else
                msfgCFOP.RemoveItem msfgCFOP.Row
        End If
    End If
    LimpFormCFOP
End Sub

Private Sub btoIncluir_Click()
    With msfgCFOP
        If txtCFOP.Text = "" Or cboSituacao.Text = "" Or cboICMSCST.Text = "" Then Exit Sub
        If idCFOP = 0 Then
                .Rows = .Rows + 1
                idCFOP = .Rows - 1
            Else
                idCFOP = .Row
        End If
        .TextMatrix(idCFOP, 1) = cboSituacao.Text
        .TextMatrix(idCFOP, 2) = cboICMSCST.Text
        .TextMatrix(idCFOP, 3) = txtCFOP.Text
        .TextMatrix(idCFOP, 4) = IIf(chkCalcICMS.Value = 1, "SIM", "NÃO")
        .TextMatrix(idCFOP, 5) = IIf(chkCalcICMSST.Value = 1, "SIM", "NÃO")
    End With
    LimpFormCFOP
    idCFOP = 0
End Sub

Private Sub cboCentroCusto_DropDown()
    Dim Rst As Recordset
    Dim sSQL As String
    sSQL = "SELECT * FROM FinanceiroCentroCustos WHERE id_empresa=" & ID_Empresa

    Set Rst = RegistroBuscar(sSQL)
    cboCentroCusto.Clear
    If Rst.BOF And Rst.EOF Then
            'MsgBox "Nenhuma CONTA cadastrada"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCentroCusto.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("Id"))) & Rst.Fields("Id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Clone
End Sub

Private Sub cboConta_DropDown()
    Dim Rst As Recordset
    Dim sSQL As String
    sSQL = "SELECT * FROM FinanceiroConta WHERE id_empresa=" & ID_Empresa
    Set Rst = RegistroBuscar(sSQL)
    cboConta.Clear
    If Rst.BOF And Rst.EOF Then
            'MsgBox "Nenhuma CONTA cadastrada"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboConta.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("Id"))) & Rst.Fields("Id") & " - " & _
                                 Rst.Fields("Agencia") & "/" & Rst.Fields("Conta")
                Rst.MoveNext
            Loop
    End If
    Rst.Clone
End Sub

Private Sub cboCSTPIS_DropDown()
    Dim sSQL    As String
    Dim Rst     As Recordset
    cboCSTPIS.Clear
    sSQL = "SELECT * FROM TributacaoCST WHERE Tabela = 'P' ORDER BY CST"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCSTPIS.AddItem Rst.Fields("CST") & " - " & Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
End Sub
Private Sub cboCSTCOFINS_DropDown()
    Dim sSQL    As String
    Dim Rst     As Recordset
    cboCSTCOFINS.Clear
    sSQL = "SELECT * FROM TributacaoCST WHERE Tabela = 'P' ORDER BY CST"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCSTCOFINS.AddItem Rst.Fields("CST") & " - " & Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboFinalidade_DropDown()
    With cboFinalidade
        .Clear
        .AddItem "1 - NF-e Normal"
        .AddItem "2 - NF-e Complementar"
        .AddItem "3 - NF-e de Ajuste"
        .AddItem "4 - NF-e de Devolução"
    End With

End Sub

Private Sub cboModBC_DropDown()
    With cboModBC
        .Clear
        .AddItem "0" & " - " & pgDescrModBC(0)
        .AddItem "1" & " - " & pgDescrModBC(1)
        .AddItem "2" & " - " & pgDescrModBC(2)
        .AddItem "3" & " - " & pgDescrModBC(3)
    End With
End Sub


Private Sub cboModBCST_DropDown()
    With cboModBCST
        .Clear
        .AddItem "0" & " - " & pgDescrModBCST(0)
        .AddItem "1" & " - " & pgDescrModBCST(1)
        .AddItem "2" & " - " & pgDescrModBCST(2)
        .AddItem "3" & " - " & pgDescrModBCST(3)
        .AddItem "4" & " - " & pgDescrModBCST(4)
        .AddItem "5" & " - " & pgDescrModBCST(5)
    End With

End Sub

Private Sub cboNaturezaOperacao_DropDown()
    With cboNaturezaOperacao
        .Clear
        .AddItem "VENDA"
        .AddItem "COMPRA"
        .AddItem "TRANSFERENCIA"
        .AddItem "DEVOLUCAO"
        .AddItem "IMPORTAÇÃO"
        .AddItem "CONSIGUINAÇÃO"
        .AddItem "REMESSA"
    End With
End Sub




Private Sub cboPlanoContas_DropDown()
    Dim Rst As Recordset
    cboPlanoContas.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroPlanoContas WHERE id_empresa=" & ID_Empresa & " ORDER BY Codigo")
    If Rst.BOF And Rst.EOF Then
            'Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboPlanoContas.AddItem ZE(Rst.Fields("id"), 3) & " - (" & Rst.Fields("Codigo") & ") " & Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub cboTpDocumento_DropDown()
 Dim Rst As Recordset
    Dim sSQL As String
    sSQL = "SELECT * FROM FinanceiroTipoDocumento WHERE id_empresa=" & ID_Empresa

    Set Rst = RegistroBuscar(sSQL)
    cboTpDocumento.Clear
    If Rst.BOF And Rst.EOF Then
            'MsgBox "Nenhuma CONTA cadastrada"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboTpDocumento.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("Id"))) & Rst.Fields("Id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Clone
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
    HDForm Me, False
    HDMenu Me, True
    txtID.Enabled = True
    msfgCFOP.Enabled = True
    
End Sub


Private Sub LimpForm()
    LimpaFormulario Me
    msfgCFOP.Rows = 1
End Sub
Private Sub LimpFormCFOP()
    cboSituacao.Clear
    cboICMSCST.Clear
    txtCFOP.Text = ""
    chkCalcICMS.Value = 0
    chkCalcICMSST.Value = 0
End Sub
Private Sub cboICMSCST_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    cboICMSCST.Clear
    If PgDadosEmpresa(ID_Empresa).RegimeTrib = "3" Then
            sSQL = "SELECT * FROM TributacaoCST WHERE tabela = 'B' ORDER BY cst"
        Else
            sSQL = "SELECT * FROM TributacaoCST WHERE tabela = 'C' ORDER BY cst"
    End If
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao,localizar taberla CST"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboICMSCST.AddItem Rst.Fields("cst") & " - " & Rst.Fields("descricao")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub cboSituacao_DropDown()
    With cboSituacao
        .Clear
        .AddItem "0 - Dentro do Estado"
        .AddItem "1 - Fora do Estado"
    End With
End Sub

Private Sub cboTipoNota_DropDown()
    With cboTipoNota
        .Clear
        .AddItem "0 - Entrada"
        .AddItem "1 - Saída"
    End With
End Sub



Private Sub msfgCFOP_DblClick()
    With msfgCFOP
        If .TextMatrix(.Row, 1) = "" Or .Row = 0 Then Exit Sub
        idCFOP = .Row '.TextMatrix(.Row, 0)
        cboSituacao.Clear
        cboSituacao.AddItem .TextMatrix(.Row, 1)
        cboSituacao.Text = cboSituacao.List(0)
        cboICMSCST.Clear
        cboICMSCST.AddItem .TextMatrix(.Row, 2)
        cboICMSCST.Text = cboICMSCST.List(0)
        txtCFOP.Text = .TextMatrix(.Row, 3)
        chkCalcICMS.Value = IIf(.TextMatrix(.Row, 4) = "SIM", 1, 0)
        chkCalcICMSST.Value = IIf(.TextMatrix(.Row, 5) = "SIM", 1, 0)
    End With
End Sub

Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    IdReg = 0
    HDMenu Me, False
    HDForm Me, True
    LimpForm
    txtID.Enabled = False
    msfgCFOP.Enabled = True
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Selecione um Registro."
    Exit Sub
    End If
    HDForm Me, True
    HDMenu Me, False
    txtID.Enabled = False
    msfgCFOP.Enabled = True
End Sub
Private Sub Excluir()
    If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
            MsgBox "Selecione um Registro"
            Exit Sub
        Else
            If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                        vbCrLf & _
                        "ID: " & txtID.Text & vbCrLf & _
                        "Descrição: " & txtDescricao.Text, vbYesNo + vbQuestion) = vbYes Then

                If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                    LimpForm
                End If
            End If
    End If
End Sub
Private Sub Clonar()
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Selecione um Registro."
    Exit Sub
    End If
    HDForm Me, True
    HDMenu Me, False
    IdReg = 0
    txtID.Text = ""
    txtID.Enabled = False
    msfgCFOP.Enabled = True
End Sub
Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
            Incluir
            
        Case "Alterar"
            Alterar
        Case "Excluir"
            Excluir
        Case "Pesquisar"
            PesquisarRegistro
        Case "Clonar"
            Clonar
        Case "Salvar"
            If grvRegistro = True Then
                If grvRegistroGrade = True Then
                        HDMenu Me, True
                        HDForm Me, False
                        txtID.Enabled = True
                        msfgCFOP.Enabled = True
                    Else
                        MsgBox "Erro ao gravar a grade", vbInformation, "Aviso"
                End If
            End If

        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpForm
            txtID.Enabled = True
            msfgCFOP.Enabled = True
        
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select
End Sub



Private Sub txtCFOP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        PesquisarRegistro (txtID.Text)
    End If
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
    
End Sub
Private Function grvRegistro() As Boolean
    Dim vReg(199)   As Variant
    Dim cReg        As Integer 'Contador de Registros
    Dim l           As Integer
    Dim tmp         As Integer
    cReg = 0
    
    If ValidarDados = False Then
        Exit Function
    End If
    
    
    vReg(cReg) = Array("Descricao", txtDescricao.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("TipoNota", Left(cboTipoNota.Text, 1), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Serie", txtSerie.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("NumInicial", txtNumInicial.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("NaturezaOperacao", cboNaturezaOperacao.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Finalidade", Left(cboFinalidade.Text, 1), "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("MovFisco", chkMovFisco.Value, "S"): cReg = cReg + 1
    vReg(cReg) = Array("MovEstoque", chkMovEstoque.Value, "S"): cReg = cReg + 1
    vReg(cReg) = Array("MovContasPR", chkMovContasPR.Value, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("MovComissao", chkMovComissao.Value, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("EnvioRF", chkEnvioRF.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ChaveAcessoRef", chkChaveAcessoRef.Value, "N"): cReg = cReg + 1
    

    vReg(cReg) = Array("ImprCampoFatura", chkImprCampoFatura.Value, "S"): cReg = cReg + 1
    If chkMovContasPR.Value = 0 Then
            vReg(cReg) = Array("Conta", "0", "N"): cReg = cReg + 1
            vReg(cReg) = Array("TpDocumento", "0", "N"): cReg = cReg + 1
            vReg(cReg) = Array("CentroCusto", "0", "N"): cReg = cReg + 1
            vReg(cReg) = Array("PlanoContas", "0", "N"): cReg = cReg + 1
        Else
            vReg(cReg) = Array("Conta", Left(cboConta.Text, 3), "N"): cReg = cReg + 1
            vReg(cReg) = Array("TpDocumento", Left(cboTpDocumento.Text, 3), "N"): cReg = cReg + 1
            vReg(cReg) = Array("CentroCusto", Left(cboCentroCusto.Text, 3), "N"): cReg = cReg + 1
            vReg(cReg) = Array("Planocontas", Left(cboPlanoContas.Text, 3), "N"): cReg = cReg + 1
    End If
    vReg(cReg) = Array("ModBC", Trim(Left(cboModBC.Text, 1)), "S"): cReg = cReg + 1
    vReg(cReg) = Array("ModBCST", Trim(Left(cboModBCST.Text, 1)), "S"): cReg = cReg + 1
    vReg(cReg) = Array("ImprDataSaida", chkImprDataSaida.Value, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ImprCampoObs", chkImprCampoObs.Value, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Modelo", txtModelo.Text, "N"): cReg = cReg + 1
    vReg(cReg) = Array("CSTPIS", Left(cboCSTPIS.Text, 2), "S"): cReg = cReg + 1
    vReg(cReg) = Array("CSTCOFINS", Left(cboCSTCOFINS.Text, 2), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Obs", txtObs.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ImpBCICMS", chkImpBCICMS.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ImpvICMS", chkImpvICMS.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ImpBCICMSST", chkImpBCICMSST.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ImpvICMSST", chkImpvICMSST.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ImpvTotalProduto", chkImpvTotalProduto.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ImpvFrete", chkImpvFrete.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ImpvSeguro", chkImpvSeguro.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ImpvDesconto", chkImpvDesconto.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ImpvOutrasDesp", chkImpvOutrasDesp.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ImpvIPI", chkImpvIPI.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ImpvTotalNota", chkImpvTotalNota.Value, "N")   ': creg = creg + 1
    
    
    If IdReg = 0 Then
            IdReg = RegistroIncluir(strTabela, vReg, cReg)
            If IdReg = 0 Then
                    grvRegistro = False
                Else
                    grvRegistro = True
            End If
        Else
            If RegistroAlterar(strTabela, vReg, cReg, "Id = " & IdReg) = False Then
                    grvRegistro = False
                Else
                    grvRegistro = True
                
            End If
    End If
End Function
Private Function grvRegistroGrade() As Boolean
    Dim Tabela2     As String
    Dim vReg(199)   As Variant
    Dim i           As Integer
    Dim cReg        As Integer 'Contador de Registros
    
    cReg = 0
    Tabela2 = strTabela & "CFOP"
    If IdReg = 0 Then
        grvRegistroGrade = False
        Exit Function
    End If
    RegistroExcluir Tabela2, "idTipoNotaFiscal=" & IdReg
    With msfgCFOP
        For i = 1 To .Rows - 1
                vReg(cReg) = Array("idTipoNotaFiscal", IdReg, "S")
                cReg = cReg + 1
                vReg(cReg) = Array("Situacao", Left(.TextMatrix(i, 1), 1), "S")
                cReg = cReg + 1
                vReg(cReg) = Array("CST", Trim(Left(.TextMatrix(i, 2), 3)), "S")
                cReg = cReg + 1
                vReg(cReg) = Array("CFOP", .TextMatrix(i, 3), "S")
                cReg = cReg + 1
                vReg(cReg) = Array("ICMS", IIf(.TextMatrix(i, 4) = "SIM", 1, 0), "S")
                cReg = cReg + 1
                vReg(cReg) = Array("ICMSST", IIf(.TextMatrix(i, 5) = "SIM", 1, 0), "S")
                'cReg = cReg + 1
            If RegistroIncluir(Tabela2, vReg, cReg) = 0 Then
                    MsgBox "Erro ao Incluir."
                    grvRegistroGrade = False
                Else
                    grvRegistroGrade = True
                    cReg = 0
            End If
        Next
    End With
End Function
Private Sub PesquisarRegistro(Optional Id As Integer)
    ''Dim idreg  As String
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim Tabela2 As String

    Tabela2 = strTabela & "CFOP"
    
    
    If Trim(Id) = 0 Then
            IdReg = formBuscar.IniciarBusca(strTabela)
            ''IdReg = IIf(idreg = "", 0, idreg)
        Else
            IdReg = Id
    End If
    
    If IdReg = 0 Then
            LimpForm
        Else
            sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg
            
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    MsgBox "Erro ao localizar o Registro"
                    LimpForm
                Else
                    Rst.MoveFirst
                    txtID.Text = IdReg
                    txtDescricao.Text = PgDadosTpNotaFiscal(IdReg).Descricao  'Rst.Fields("Descricao")
                    
                    cboTipoNota.Clear
                    cboTipoNota.AddItem PgDadosTpNotaFiscal(IdReg).TipoNota & " - " & PgDadosTpNotaFiscal(IdReg).TipoNotaDescr  'IIf(IsNull(Rst.Fields("TipoNota")), " ", Rst.Fields("TipoNota"))
                    cboTipoNota.Text = cboTipoNota.List(0)
                    
                    txtSerie.Text = PgDadosTpNotaFiscal(IdReg).Serie  'IIf(IsNull(Rst.Fields("Serie")), "", Rst.Fields("Serie"))
                    txtModelo.Text = PgDadosTpNotaFiscal(IdReg).Modelo  'IIf(IsNull(Rst.Fields("Modelo")), "", Rst.Fields("Modelo"))
                    txtNumInicial.Text = PgDadosTpNotaFiscal(IdReg).NumInicial ' IIf(IsNull(Rst.Fields("NumInicial")), "", Rst.Fields("NumInicial"))
                    
                    cboNaturezaOperacao.Clear
                    cboNaturezaOperacao.AddItem PgDadosTpNotaFiscal(IdReg).Natureza  'Rst.Fields("NaturezaOperacao")
                    cboNaturezaOperacao.Text = cboNaturezaOperacao.List(0)
                    
                    cboFinalidade.Clear
                    cboFinalidade.AddItem PgDadosTpNotaFiscal(IdReg).Finalidade & " - " & PgDadosTpNotaFiscal(IdReg).FinalidadeDescr  'IIf(IsNull(Rst.Fields("Finalidade")), " ", Rst.Fields("Finalidade"))
                    cboFinalidade.Text = cboFinalidade.List(0)
                    
                    cboModBC.Clear
                    cboModBC.AddItem PgDadosTpNotaFiscal(IdReg).ModBC & " - " & pgDescrModBC(PgDadosTpNotaFiscal(IdReg).ModBC)
                    cboModBC.Text = cboModBC.List(0)
                    
                    cboModBCST.Clear
                    cboModBCST.AddItem PgDadosTpNotaFiscal(IdReg).ModBCST & " - " & pgDescrModBCST(PgDadosTpNotaFiscal(IdReg).ModBCST)
                    cboModBCST.Text = cboModBCST.List(0)
                    
                    chkEnvioRF.Value = PgDadosTpNotaFiscal(IdReg).EnvioRF
                    chkChaveAcessoRef.Value = PgDadosTpNotaFiscal(IdReg).ChaveAcessoRef
                    
                    chkMovFisco.Value = PgDadosTpNotaFiscal(IdReg).MovFisco
                    chkMovComissao.Value = PgDadosTpNotaFiscal(IdReg).MovComissao
                    chkMovEstoque.Value = PgDadosTpNotaFiscal(IdReg).MovEstoque  'Rst.Fields("MovEstoque")
                    chkMovContasPR.Value = PgDadosTpNotaFiscal(IdReg).MovContasPR  'Rst.Fields("MovContasPR")
                    chkImprCampoFatura.Value = PgDadosTpNotaFiscal(IdReg).ImpCmpFatura 'Rst.Fields("ImprCampoFatura")
                    'chkImprFatura.Value = PgDadosTpNotaFiscal(IdReg).ImpFatura  'IIf(IsNull(Rst.Fields("ImprFatura")), " ", Rst.Fields("ImprFatura"))
                    chkImprDataSaida.Value = PgDadosTpNotaFiscal(IdReg).ImpDtSaida  'Rst.Fields("ImprDataSaida")

                    chkImpBCICMS.Value = PgDadosTpNotaFiscal(IdReg).ImpBCICMS
                    chkImpvICMS.Value = PgDadosTpNotaFiscal(IdReg).ImpvICMS
                    chkImpBCICMSST.Value = PgDadosTpNotaFiscal(IdReg).ImpBCICMSST
                    chkImpvICMSST.Value = PgDadosTpNotaFiscal(IdReg).ImpvICMSST
                    chkImpvTotalProduto.Value = PgDadosTpNotaFiscal(IdReg).ImpvTotalProduto
                    chkImpvFrete.Value = PgDadosTpNotaFiscal(IdReg).ImpvFrete
                    chkImpvSeguro.Value = PgDadosTpNotaFiscal(IdReg).ImpvSeguro
                    chkImpvDesconto.Value = PgDadosTpNotaFiscal(IdReg).ImpvDesconto
                    chkImpvOutrasDesp.Value = PgDadosTpNotaFiscal(IdReg).ImpvOutrasDesp
                    chkImpvIPI.Value = PgDadosTpNotaFiscal(IdReg).ImpvIPI
                    chkImpvTotalNota.Value = PgDadosTpNotaFiscal(IdReg).ImpvTotalNota
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    chkImprCampoObs.Value = PgDadosTpNotaFiscal(IdReg).ImpInfCompl  'Rst.Fields("ImprCampoObs")
                    txtObs.Text = PgDadosTpNotaFiscal(IdReg).Obs  'IIf(IsNull(Rst.Fields("Obs")), " ", Rst.Fields("Obs"))
                    cboConta.Clear
                    cboConta.AddItem IIf(PgDadosTpNotaFiscal(IdReg).Conta = 0, _
                                         " ", _
                                         Left(String(3, "0"), 3 - Len(Trim(PgDadosTpNotaFiscal(IdReg).Conta))) & PgDadosTpNotaFiscal(IdReg).Conta & " - " & pgDadosConta(PgDadosTpNotaFiscal(IdReg).Conta).Agencia & "/" & pgDadosConta(PgDadosTpNotaFiscal(IdReg).Conta).Conta)
                    cboConta.Text = cboConta.List(0)
                    
                    cboCentroCusto.Clear
                    cboCentroCusto.AddItem IIf(PgDadosTpNotaFiscal(IdReg).CentroCusto = 0, _
                                               " ", _
                                               Left(String(3, "0"), 3 - Len(Trim(PgDadosTpNotaFiscal(IdReg).CentroCusto))) & PgDadosTpNotaFiscal(IdReg).CentroCusto & " - " & pgDadosCentroCustos(PgDadosTpNotaFiscal(IdReg).CentroCusto).Descricao)
                    cboCentroCusto.Text = cboCentroCusto.List(0)
                    
                    cboTpDocumento.Clear
                    cboTpDocumento.AddItem IIf(PgDadosTpNotaFiscal(IdReg).TipoDoc = 0, _
                                               " ", _
                                               Left(String(3, "0"), 3 - Len(Trim(PgDadosTpNotaFiscal(IdReg).TipoDoc))) & PgDadosTpNotaFiscal(IdReg).TipoDoc & " - " & pgDadosTipoDocumento(PgDadosTpNotaFiscal(IdReg).TipoDoc).Descricao)
                    cboTpDocumento.Text = cboTpDocumento.List(0)
                    
                    
                    cboPlanoContas.Clear
                    cboPlanoContas.AddItem IIf(PgDadosTpNotaFiscal(IdReg).PlanoContas = 0, _
                                               " ", _
                                               ZE(PgDadosTpNotaFiscal(IdReg).PlanoContas, 3) & " - (" & PgDadosPlanoContas("ID", CStr(PgDadosTpNotaFiscal(IdReg).PlanoContas)).Codigo & ") " & PgDadosPlanoContas("ID", CStr(PgDadosTpNotaFiscal(IdReg).PlanoContas)).Descricao)
                    cboPlanoContas.Text = cboPlanoContas.List(0)
                    
                    cboCSTPIS.Clear
                    If PgDadosTpNotaFiscal(IdReg).CSTPIS <> 0 Then
                        cboCSTPIS.AddItem PgDadosTpNotaFiscal(IdReg).CSTPIS & " - " & PgDadosCST(PgDadosTpNotaFiscal(IdReg).CSTPIS, "PIS").Descricao
                        cboCSTPIS.Text = cboCSTPIS.List(0)
                    End If
                    cboCSTCOFINS.Clear
                    If PgDadosTpNotaFiscal(IdReg).CSTCOFINS <> 0 Then
                        cboCSTCOFINS.AddItem PgDadosTpNotaFiscal(IdReg).CSTCOFINS & " - " & PgDadosCST(PgDadosTpNotaFiscal(IdReg).CSTCOFINS, "COFINS").Descricao
                        cboCSTCOFINS.Text = cboCSTCOFINS.List(0)
                    End If
                    
            End If
            Rst.Close
            
            'Carrega grade
            sSQL = "SELECT * FROM " & Tabela2 & " WHERE ID_Empresa = " & ID_Empresa & " AND IDTipoNotaFiscal = " & IdReg
    
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    msfgCFOP.Rows = 1
                Else
                    msfgCFOP.Rows = 1
                    Rst.MoveFirst
                    Do Until Rst.EOF
                        With msfgCFOP
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, 1) = IIf(Rst.Fields("Situacao") = 0, "0 - Dentro do Estado", "1 - Fora do Estado")
                            .TextMatrix(.Rows - 1, 2) = Rst.Fields("CST") & " - " & PgDadosCST(Rst.Fields("CST"), "ICMS").Descricao
                            .TextMatrix(.Rows - 1, 3) = Rst.Fields("CFOP")
                            .TextMatrix(.Rows - 1, 4) = IIf(Rst.Fields("ICMS") = 1, "SIM", "NÃO")
                            .TextMatrix(.Rows - 1, 5) = IIf(Rst.Fields("ICMSST") = 1, "SIM", "NÃO")
                            Rst.MoveNext
                        End With
                    Loop
            End If
    End If
End Sub

Private Sub MontarBaseDeDados()
    Dim vDados(199)  As Variant
    Dim contReg     As Integer
    Dim i           As Integer
    
    contReg = 0
    'formManutencaoTabelas.IniciarManutencao Me
    'cabecario do Pedido
    
    vDados(contReg) = Array("Descricao", "240", "S"): contReg = contReg + 1
    vDados(contReg) = Array("TipoNota", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Serie", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("NumInicial", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("NaturezaOperacao", "50", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Finalidade", "100", "S"): contReg = contReg + 1

    vDados(contReg) = Array("MovComissao", "5", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("EnvioRF", "5", "N"):  contReg = contReg + 1
    vDados(contReg) = Array("ChaveAcessoRef", "5", "N"):  contReg = contReg + 1
    
    
    vDados(contReg) = Array("MovEstoque", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("MovFisco", "5", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("MovContasPR", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ImprCampoFatura", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ModBC", "2", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ModBCST", "2", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ImprDataSaida", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ImprCampoObs", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Modelo", "2", "N"): contReg = contReg + 1
    vDados(contReg) = Array("tpEmis", "2", "N"): contReg = contReg + 1
    
    vDados(contReg) = Array("CSTPIS", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("CSTCOFINS", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Obs", "1000", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("Conta", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("TpDocumento", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("CentroCusto", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("PlanoContas", "10", "N"): contReg = contReg + 1
    
    vDados(contReg) = Array("ImpBCICMS", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ImpvICMS", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ImpBCICMSST", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ImpvICMSST", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ImpvTotalProduto", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ImpvFrete", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ImpvSeguro", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ImpvDesconto", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ImpvOutrasDesp", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ImpvIPI", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ImpvTotalNota", "1", "N") ': contReg = contReg + 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg
        
    'Outros Dados
    contReg = 0
    vDados(contReg) = Array("idTipoNotaFiscal", "50", "N"): contReg = contReg + 1
    vDados(contReg) = Array("Situacao", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("CST", "120", "S"): contReg = contReg + 1
    vDados(contReg) = Array("CFOP", "250", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMSST", "5", "S") ':contReg = contReg + 1

    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "CFOP"
End Sub

Private Sub txtmodelo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub
Private Function ValidarDados() As Boolean
    If Trim(txtDescricao.Text) = "" Then
        MsgBox "Os campos Descrição é obrigatorios!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    If Trim(cboCSTPIS.Text) = "" Or Trim(cboCSTCOFINS.Text) = "" Then
        MsgBox "Os campos PIS/COFINS são obrigatorios!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    If Trim(cboModBC.Text) = "" Then
        MsgBox "Os campos MODALIDADE DE BASE DE CALCULO é obrigatorios!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
        If Trim(cboModBCST.Text) = "" Then
        MsgBox "Os campos MODALIDADE DE BASE DE CALCULO é obrigatorios!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    ValidarDados = True
End Function
