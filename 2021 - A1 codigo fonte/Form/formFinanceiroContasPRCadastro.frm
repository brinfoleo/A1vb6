VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFinanceiroContasPRCadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  "
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   10215
   Begin VB.Frame Frame8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   55
      Top             =   7980
      Width           =   9975
      Begin VB.ComboBox cboContaQuitacao 
         Height          =   315
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   600
         Width           =   5115
      End
      Begin VB.CheckBox chkQuitado 
         Caption         =   "Quitado"
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
         Left            =   120
         TabIndex        =   59
         Top             =   0
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker dtpQuitacao 
         Height          =   315
         Left            =   1740
         TabIndex        =   57
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   101515265
         CurrentDate     =   40602
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Conta:"
         Height          =   195
         Left            =   3420
         TabIndex        =   63
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "O valor do Juros/Multa será alterado automaticamente de acordo com a data de quitação. "
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   300
         Width           =   7860
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Data do Pagamento:"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   630
         Width           =   1515
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Observações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   5580
      TabIndex        =   53
      Top             =   6300
      Width           =   4515
      Begin VB.TextBox txtObs 
         Height          =   1275
         Left            =   120
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Text            =   "formFinanceiroContasPRCadastro.frx":0000
         Top             =   240
         Width           =   4275
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Sacado/Cedente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   44
      Top             =   1380
      Width           =   9975
      Begin VB.ComboBox cboCadastro 
         Height          =   315
         ItemData        =   "formFinanceiroContasPRCadastro.frx":0006
         Left            =   1080
         List            =   "formFinanceiroContasPRCadastro.frx":0008
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   300
         Width           =   2415
      End
      Begin VB.ComboBox cboNome 
         Height          =   315
         Left            =   2760
         TabIndex        =   47
         Text            =   "Combo1"
         Top             =   720
         Width           =   7035
      End
      Begin VB.TextBox txtDoc 
         Height          =   315
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Cadastro:"
         Height          =   195
         Left            =   300
         TabIndex        =   51
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ/CFP:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   780
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Outros Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   5580
      TabIndex        =   33
      Top             =   4260
      Width           =   4515
      Begin VB.TextBox txtVlCobrado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2100
         MaxLength       =   30
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtMultaMora 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2100
         MaxLength       =   30
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtDeducoes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2100
         MaxLength       =   30
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   900
         Width           =   2295
      End
      Begin VB.TextBox txtAbatimento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2100
         MaxLength       =   30
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtAcrescimo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2100
         MaxLength       =   30
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Cobrado ( = ):"
         Height          =   195
         Left            =   420
         TabIndex        =   38
         Top             =   1605
         Width           =   1635
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Mora / Multa ( + ):"
         Height          =   195
         Left            =   480
         TabIndex        =   37
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Outras Deduções ( - ):"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Desconto Abatimento ( - ):"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Outros Acréscimos ( + ):"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Controle Interno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   120
      TabIndex        =   28
      Top             =   2640
      Width           =   5415
      Begin VB.ComboBox cboPlanoContas 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   1500
         Width           =   3915
      End
      Begin VB.ComboBox cboDocumento 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   1080
         Width           =   3915
      End
      Begin VB.ComboBox cboCentroCustos 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   660
         Width           =   3915
      End
      Begin VB.ComboBox cboConta 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   240
         Width           =   3915
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Plano de Contas:"
         Height          =   195
         Left            =   60
         TabIndex        =   65
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Documento:"
         Height          =   195
         Left            =   300
         TabIndex        =   49
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Centro de Custos:"
         Height          =   195
         Left            =   60
         TabIndex        =   31
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Conta:"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados da Duplicata"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   5415
      Begin VB.TextBox txtDiasProtesto 
         Height          =   285
         Left            =   4440
         MaxLength       =   3
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   2280
         Width           =   615
      End
      Begin VB.ComboBox cboBanco 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2760
         Width           =   3435
      End
      Begin VB.TextBox txtJuros 
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   2400
         Width           =   1275
      End
      Begin VB.TextBox txtMulta 
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2100
         Width           =   1275
      End
      Begin VB.TextBox txtVlDuplicata 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtNumDuplicata 
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1380
         Width           =   2895
      End
      Begin VB.TextBox txtNossoNumero 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   600
         Width           =   2835
      End
      Begin VB.TextBox txtLinhaDigitavel 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   240
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtpVencimento 
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   41222145
         CurrentDate     =   40557
      End
      Begin VB.Label Label24 
         Caption         =   "Protesto (dias):"
         Height          =   255
         Left            =   3240
         TabIndex        =   60
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Banco:"
         Height          =   195
         Left            =   720
         TabIndex        =   26
         Top             =   2820
         Width           =   915
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Juros (dia):"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Multa (mês):"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor da Duplicata:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Número da Duplicata:"
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Nosso Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Linha digitavel:"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Fatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5580
      TabIndex        =   4
      Top             =   2640
      Width           =   4515
      Begin VB.TextBox txtVlFatura 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1020
         Width           =   1635
      End
      Begin VB.TextBox txtNumFatura 
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   660
         Width           =   2715
      End
      Begin MSComCtl2.DTPicker dtpEmissao 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   101842945
         CurrentDate     =   40557
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor da Fatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Número da Fatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Data da Emissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Conta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9975
      Begin VB.OptionButton optContas 
         Caption         =   "A Receber"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton optContas 
         Caption         =   "A Pagar"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lblContas 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   2640
         TabIndex        =   48
         Top             =   180
         Width           =   7095
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
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
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pesquisar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Picture         =   "formFinanceiroContasPRCadastro.frx":000A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRCadastro.frx":045C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRCadastro.frx":0776
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRCadastro.frx":1008
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRCadastro.frx":225A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRCadastro.frx":2B34
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRCadastro.frx":33C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRCadastro.frx":3C58
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRCadastro.frx":4EAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRCadastro.frx":51C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRCadastro.frx":54DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRCadastro.frx":58D5
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFinanceiroContasPRCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg           As Integer ' ID do registro
Dim IDSacado        As Integer  'ID do Cliente ou Fornecedor
Dim strTabela       As String ' Nome da tabela do Contas a Pagar e Receber
Dim strTabelaSacado As String 'Nome da tabela do sacado transp /Cliente ou Fornecedor
Dim DocQuitado      As Boolean  'True = documento ja quitado / false - doc. em aberto
Dim gerarExtorno    As Boolean 'True=gerar extorno / false = nao gerar extorno
Private Sub HDQuitacao()
    dtpQuitacao.Enabled = IIf(chkQuitado.Value = 0, False, True)
    cboContaQuitacao.Enabled = IIf(chkQuitado.Value = 0, False, True)

End Sub

Public Sub LoadDocumento(iDoc As Integer)
    IdReg = iDoc
    dtpQuitacao.Value = Date
    PesquisarRegistro
    Me.Show
End Sub

Private Sub CalcQuitacao()
    Dim vlDupl      As String
    Dim vlCobrado   As String
    Dim vlDesc      As String
    Dim vlAcres     As String
    
    Dim dQuitacao   As Date
    Dim dVenc       As Date
    Dim DiasVenc    As Integer
    Dim pMulta      As String
    Dim pJuros      As String
    
    Dim vMulta      As String
    Dim vJuros      As String
    Dim vDupl       As String
    
    If chkQuitado.Value = 0 Then
            dQuitacao = dtpVencimento.Value
        Else
            dQuitacao = dtpQuitacao.Value
    End If
    
    dVenc = dtpVencimento.Value
    
    vlDupl = ChkVal(txtVlDuplicata.Text, 0, cDecMoeda)
    pMulta = ChkVal(txtMulta.Text, 0, 3)
    pJuros = ChkVal(txtJuros.Text, 0, 3)
    
    If DocQuitado = True Then Exit Sub
    'If chkQuitado.Value = 0 Then Exit Sub
    
    '07.06.2017
    'verifica se a data cai no sabado ou domening
    
    Select Case DatePart("w", dVenc)
        Case 1 'Domingo
            'DOMINGO"
            dVenc = dtpVencimento.Value + 1
        Case 7 'Sabado
            'SABADO"
            dVenc = dtpVencimento.Value + 2
        Case Else
            'Me.Caption = ""
    End Select
    
    If dQuitacao > dVenc Then
        dVenc = dtpVencimento.Value
    End If
    
    
    
    DiasVenc = dQuitacao - dVenc 'dtpVencimento.Value
    If DiasVenc <= 0 Then
            vMulta = "0"
            DiasVenc = 0
        Else
            vMulta = cobCalcMulta(vlDupl, pMulta) '(Val(ChkVal(txtMulta.Text, 0, 5)) * Val(ChkVal(txtVlDuplicata.Text, 0, cDecMoeda))) / 100
    End If
    vDupl = ChkVal(Val(ChkVal(vMulta, 0, cDecMoeda)) + Val(vlDupl), 0, cDecMoeda)
    
    vJuros = cobCalcMora(vDupl, DiasVenc, pJuros, "T") '(Val(ChkVal(txtJuros.Text, 0, 5)) * Val(DiasVenc)) * Val(vDupl) / 100
    
    txtMultaMora.Text = ConvMoeda(Val(ChkVal(vMulta, 0, cDecMoeda)) + Val(ChkVal(vJuros, 0, cDecMoeda)))
    
    vlDesc = Val(ChkVal(txtAbatimento.Text, 0, 2)) + Val(ChkVal(txtDeducoes.Text, 0, 2))
    vlAcres = Val(ChkVal(txtAcrescimo.Text, 0, 2)) + Val(ChkVal(txtMultaMora.Text, 0, 2))
    
    vlCobrado = Val(ChkVal(txtVlDuplicata.Text, 0, 2)) + Val(ChkVal(vlAcres, 0, 2)) - Val(ChkVal(vlDesc, 0, 2))
    
    txtVlCobrado.Text = ConvMoeda(vlCobrado)
End Sub

Private Function nomeTabela(sTabela As String) As String
        Select Case LCase(sTabela)
            Case "clientes"
                nomeTabela = "01 - Clientes"
            Case "fornecedores"
                nomeTabela = "02 - Fornecedores"
            Case "trasportadoras"
                nomeTabela = "03 - Transportadora"
            Case "rhfuncionariocadastro"
                nomeTabela = "04 - Funcionario"
            Case Else
                nomeTabela = "00 - Outros"
        End Select
End Function
Private Sub LimpForm()
    LimpaFormulario Me
    dtpEmissao.Value = Date
    dtpVencimento.Value = Date
End Sub

Private Sub PesquisarRegistro()
    On Error Resume Next
    Dim Rst     As Recordset
    Dim sSQL    As String
    LimpaFormulario Me
    

    If IdReg = 0 Then
        IdReg = formBuscar.IniciarBusca(strTabela)
    End If
    sSQL = "SELECT * FROM FinanceiroContasPRCadastro WHERE ID_Empresa = " & ID_Empresa & " AND ID=" & IdReg
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar Registro"
            Exit Sub
        Else
            Rst.MoveFirst
            DoEvents
            If Rst.Fields("ContaPR") = "P" Then
                    optContas.Item(0).Value = True
                    optContas.Item(1).Value = False
                Else
                    optContas.Item(0).Value = False
                    optContas.Item(1).Value = True
            End If
            cboCadastro.Clear
            cboCadastro.AddItem nomeTabela(cNull(Rst.Fields("Tabela")))
            cboCadastro.Text = cboCadastro.List(0)
            
            
            IDSacado = Rst.Fields("idSacado")
            cboNome.Text = Rst.Fields("Nome")
            txtDoc.Text = Rst.Fields("CNPJ")
            
                      
            dtpEmissao.Value = Rst.Fields("Emissao")
            txtNumFatura.Text = Rst.Fields("NumFatura")
            txtVlFatura.Text = ConvMoeda(IIf(IsNull(Rst.Fields("vlFatura")), "0", Rst.Fields("vlFatura")))
            
            If IsNull(Rst.Fields("DataQuitacao")) Then
                    DocQuitado = False
                    chkQuitado.Value = 0
                    dtpQuitacao.Value = Date
                    
                Else
                    DocQuitado = True
                    chkQuitado.Value = 1
                    dtpQuitacao.Value = Rst.Fields("DataQuitacao")
                    txtVlCobrado.Text = IIf(IsNull(Rst.Fields("VlCobrado")), "0.00", Rst.Fields("VlCobrado"))
                    cboContaQuitacao.Clear
                    If Rst.Fields("idContaQuitacao") <> 0 Or IsNull(Rst.Fields("idContaQuitacao")) = False Then
                        cboContaQuitacao.AddItem ZE(Rst.Fields("idContaQuitacao"), 3) & " - " & pgDadosConta(Rst.Fields("idContaQuitacao")).Agencia & "/" & pgDadosConta(Rst.Fields("idContaQuitacao")).Conta
                        cboContaQuitacao.Text = cboContaQuitacao.List(0)
                    End If
            End If
            'HDQuitacao
             If Not IsNull(Rst.Fields("IdBanco")) And Rst.Fields("IdBanco") <> 0 Then
                cboBanco.AddItem pgDadosBanco(Rst.Fields("IdBanco")).Id & " - " & pgDadosBanco(Rst.Fields("IdBanco")).Nome
                cboBanco.Text = cboBanco.List(0)
            End If
            If Not IsNull(Rst.Fields("Conta")) And Rst.Fields("Conta") <> 0 Then
                If Len(Trim(pgDadosConta(Rst.Fields("Conta")).Id)) <> 0 Then
                    cboConta.AddItem pgDadosConta(Rst.Fields("Conta")).Id & " - " & pgDadosConta(Rst.Fields("Conta")).Agencia & "/" & pgDadosConta(Rst.Fields("Conta")).Conta
                    cboConta.Text = cboConta.List(0)
                End If
                    
            End If
            
            txtLinhaDigitavel.Text = IIf(IsNull(Rst.Fields("LinhaDigitavel")), "", Rst.Fields("LinhaDigitavel"))
            txtNossoNumero.Text = IIf(IsNull(Rst.Fields("NossoNumero")), "", Rst.Fields("NossoNumero"))
            dtpVencimento.Value = Rst.Fields("Vencimento")
            txtNumDuplicata.Text = IIf(IsNull(Rst.Fields("NumDuplicata")), "", Rst.Fields("NumDuplicata"))
            txtVlDuplicata.Text = ConvMoeda(IIf(IsNull(Rst.Fields("vlDuplicata")), "0", Rst.Fields("vlDuplicata")))
            txtMulta.Text = IIf(IsNull(Rst.Fields("Multa")), "", Rst.Fields("Multa"))
            txtJuros.Text = IIf(IsNull(Rst.Fields("Juros")), "", Rst.Fields("Juros"))
            txtDiasProtesto.Text = IIf(IsNull(Rst.Fields("DiasProtesto")), "", Rst.Fields("DiasProtesto"))
           
            cboCentroCustos.Clear
            If pgDadosCentroCustos(Rst.Fields("CentroCusto")).Id <> 0 Then
                cboCentroCustos.AddItem pgDadosCentroCustos(Rst.Fields("CentroCusto")).Id & " - " & pgDadosCentroCustos(Rst.Fields("CentroCusto")).Descricao
                cboCentroCustos.Text = cboCentroCustos.List(0)
            End If
            
            cboPlanoContas.Clear
            cboPlanoContas.AddItem ZE(PgDadosPlanoContas("ID", Rst.Fields("PlanoContas")).Id, 3) & " - (" & PgDadosPlanoContas("ID", Rst.Fields("PlanoContas")).Codigo & ") " & PgDadosPlanoContas("ID", Rst.Fields("PlanoContas")).Descricao
            cboPlanoContas.Text = cboPlanoContas.List(0)
            
            If pgDadosTipoDocumento(Rst.Fields("tpDocumento")).Id <> 0 Then
                cboDocumento.AddItem pgDadosTipoDocumento(Rst.Fields("tpDocumento")).Id & " - " & pgDadosTipoDocumento(Rst.Fields("tpDocumento")).Descricao
                cboDocumento.Text = cboDocumento.List(0)
            End If
            txtAcrescimo.Text = ConvMoeda(IIf(IsNull(Rst.Fields("Acrescimo")), "0.00", Rst.Fields("Acrescimo")))
            txtAbatimento.Text = ConvMoeda(IIf(IsNull(Rst.Fields("Abatimento")), "0.00", Rst.Fields("Abatimento")))
            txtDeducoes.Text = ConvMoeda(IIf(IsNull(Rst.Fields("Deducoes")), "0.00", Rst.Fields("Deducoes")))
            txtMultaMora.Text = ConvMoeda(IIf(IsNull(Rst.Fields("MultaMora")), "0.00", Rst.Fields("MultaMora")))
            txtObs.Text = cNull(Rst.Fields("OBS"))

            
    End If
    Rst.Close
End Sub

Private Sub cboBanco_DropDown()
   Dim Rst As Recordset
    cboBanco.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroBancoCadastro")
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboBanco.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("nome")
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboCadastro_Click()
    If cboCadastro.Text = "" Then Exit Sub
    IDSacado = 0
    cboNome.Clear
    txtDoc.Text = ""
    Select Case Left(Trim(cboCadastro.Text), 2)
        Case 1 'Cliente
            strTabelaSacado = "Clientes"
        Case 2 'Fornecedor
            strTabelaSacado = "Fornecedores"
        Case 3 ' Transportadora
            strTabelaSacado = "Transportadoras"
        Case 4 ' Transportadora
            strTabelaSacado = "RHFuncionarioCadastro"
        Case Else
            strTabelaSacado = ""
    End Select
End Sub

Private Sub cboCadastro_DropDown()
    With cboCadastro
        .Clear
        .AddItem "00 - Outros"
        .AddItem "01 - Clientes"
        .AddItem "02 - Fornecedores"
        .AddItem "03 - Transportadora"
        .AddItem "04 - Funcionario"
    End With
End Sub


Private Sub cboCentroCustos_DropDown()
    Dim Rst As Recordset
    cboCentroCustos.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCentroCustos")
    If Rst.BOF And Rst.EOF Then
            'Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCentroCustos.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub cboConta_Click()
    On Error GoTo TrtErro
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim idConta As Integer
    
    If Trim(RS(cboConta.Text)) = "" Then Exit Sub
    idConta = Trim(Left(cboConta.Text, 3))
    sSQL = "SELECT * FROM FinanceiroConta WHERE ID_Empresa = " & ID_Empresa & " AND id = " & idConta
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar conta"
            Exit Sub
        Else
            Rst.MoveFirst
            txtMulta.Text = IIf(IsNull(Rst.Fields("Multa")), "0", Rst.Fields("Multa"))
            txtJuros.Text = IIf(IsNull(Rst.Fields("Juros")), "0", Rst.Fields("Juros"))
            '22.02.2017 - if removido pois nao atualizava quando mudava a conta
            'cboconta
            'If Trim(cboBanco.Text) = "" Then
                cboBanco.Clear
                cboBanco.AddItem IIf(IsNull(Rst.Fields("banco")), " ", Left("000", 3 - Len(Rst.Fields("banco"))) & Rst.Fields("banco")) & " - " & pgDadosBanco(Rst.Fields("banco")).Nome
                cboBanco.Text = cboBanco.List(0)
            'End If
            txtDiasProtesto.Text = IIf(IsNull(Rst.Fields("DiasProtesto")), "0", Rst.Fields("DiasProtesto"))
    End If
    Rst.Close
    Exit Sub
TrtErro:
    Exit Sub
End Sub

Private Sub cboConta_DropDown()
    Dim Rst As Recordset
    cboConta.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroConta")
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboConta.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Agencia") & "/" & Rst.Fields("Conta")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub cboDocumento_DropDown()
    Dim Rst As Recordset
    cboDocumento.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento WHERE ID_Empresa=" & ID_Empresa)
    If Rst.BOF And Rst.EOF Then
            'Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboDocumento.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub cboNome_Click()
    If Trim(cboNome.Text) = "" Then Exit Sub
    txtDoc.Text = ""
    IDSacado = Trim(Left(cboNome.Text, 5))
    PesquisarSacado
End Sub

Private Sub cboNome_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    If Trim(cboCadastro.Text) = "" Or strTabelaSacado = "" Then
        MsgBox "Selecione um tipo de cadastro."
        Exit Sub
    End If
    'cboNome.Clear
    If Left(cboCadastro.Text, 2) = "00" Then
        Exit Sub
    End If
    sSQL = "SELECT * FROM " & strTabelaSacado & _
           " WHERE ID_Empresa = " & ID_Empresa & _
           " AND xNome LIKE '" & Trim(cboNome.Text) & "%'" & _
           " ORDER BY xNome"
    
    If sSQL = "" Then Exit Sub
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboNome.AddItem Left(String(5, "0"), 5 - Len(Trim(Rst.Fields("id")))) & Trim(Rst.Fields("id")) & " - " & Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If
    Rst.Close

End Sub

Private Sub cboNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IDSacado = 0
        PesquisarSacado
    End If

End Sub



Private Sub cboPlanoContas_DropDown()
    Dim Rst As Recordset
    cboPlanoContas.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroPlanoContas ORDER BY Codigo")
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


Private Sub chkQuitado_Click()
    If txtVlDuplicata.Enabled = True Then
        HDQuitacao
    End If
    CalcQuitacao
End Sub

Private Sub dtpQuitacao_Change()
    'If chkQuitado.Value = 1 Then
        CalcQuitacao
    'End If
End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub
Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    IdReg = 0
    HDMenu Me, False
    HDForm Me, True
    LimpForm
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Selecione um Registro."
        Exit Sub
    End If
    If chkQuitado.Value = 1 Then
        If MsgBox("Documento já esta QUITADO!" & vbCrLf & _
                  "A reabertura gerará um extorno no movimento da conta a qual foi feito o lancamento de quitação." & vbCrLf & vbCrLf & _
                  "Deseja reabrir e alterar assim mesmo?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                DocQuitado = False
                gerarExtorno = True
            Else
                gerarExtorno = False
                DocQuitado = True
                Exit Sub
        End If
    End If
    HDForm Me, True
    HDMenu Me, False
    
    If cboContaQuitacao.Text = "" Then
        If cboConta.Text <> "" Then
            cboContaQuitacao.Clear
            cboContaQuitacao.AddItem cboConta.Text
            cboContaQuitacao.Text = cboContaQuitacao.List(0)
        End If
    End If
    HDQuitacao
End Sub
Private Sub Clonar()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Selecione um Registro.", vbInformation, App.EXEName
        Exit Sub
    End If
    If MsgBox("Deseja realmente CLONAR esta fatura?", vbInformation + vbYesNo, App.EXEName) = vbNo Then
        Exit Sub
    End If
    IdReg = 0
    
    If chkQuitado.Value = 1 Then
        If MsgBox("Documento já esta QUITADO!" & vbCrLf & _
                  "A reabertura gerará um extorno no movimento da conta a qual foi feito o lancamento de quitação." & vbCrLf & vbCrLf & _
                  "Deseja reabrir e Clonar assim mesmo?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                DocQuitado = False
                ExtornarLancamentoEmConta
            Else
                DocQuitado = True
                Exit Sub
        End If
    End If
    HDForm Me, True
    HDMenu Me, False
    
    If cboContaQuitacao.Text = "" Then
        If cboConta.Text <> "" Then
            cboContaQuitacao.Clear
            cboContaQuitacao.AddItem cboConta.Text
            cboContaQuitacao.Text = cboContaQuitacao.List(0)
        End If
    End If
End Sub
Private Sub ExtornarLancamentoEmConta()
    Dim cd As String
     If optContas.Item(0).Value = True Then
            'Debito reverter para crediro
            cd = "C"
        Else
            'Credito reverter para debito
            cd = "D"
    End If
    If Trim(cboContaQuitacao.Text) = "" Then Exit Sub
    '########################################################################################
    '### Modificado em 27/04/2012
    '### Os valores Extornados Serao apagados da tabela
    '########################################################################################
    'MovimentarConta Left(cboContaQuitacao.Text, 3), _
                    cd, _
                    IdReg, _
                    dtpQuitacao.Value, _
                    txtNumDuplicata.Text, _
                    Left(cboDocumento.Text, 3), _
                    "EXTORNO: " & cboNome.Text, _
                    txtVlCobrado.Text
                    
   'Public Sub MovimentarConta(IdConta As Integer, _
                           cd As String, _
                           IdRegDoc As Integer, _
                           Data As String, _
                           nDoc As String, _
                           tDoc As Integer, _
                           Descricao As String, _
                           Valor As String)
     
    'nDoc - Numero do Documeto
    'tDoc - Codigo interno do Tipo de Documento
    
    Dim Saldo   As String
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim vReg(20)    As Variant
    Dim cReg    As Integer
    Dim Valor   As String
    Dim idConta As Integer
    
    Valor = ChkVal(txtVlCobrado.Text, 0, cDecMoeda)
    idConta = Left(cboContaQuitacao.Text, 3)
    sSQL = "SELECT * FROM FinanceiroConta WHERE id = " & idConta
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar conta.", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            Saldo = ChkVal(IIf(IsNull(Rst.Fields("Saldo")), 0, Rst.Fields("Saldo")), 0, cDecMoeda)
            If cd = "C" Then
                    Saldo = Val(Saldo) + Val(ChkVal(Valor, 0, cDecMoeda))
                Else
                    Saldo = Val(Saldo) - Val(ChkVal(Valor, 0, cDecMoeda))
            End If
            Saldo = ChkVal(Saldo, 0, cDecMoeda)
    End If
    Rst.Close
    'cReg = 0
    'vReg(cReg) = Array("IdConta", IdConta, "N"): cReg = cReg + 1
    'vReg(cReg) = Array("IdRegDoc", IdRegDoc, "N"): cReg = cReg + 1
    'vReg(cReg) = Array("Data", Data, "D"): cReg = cReg + 1
    'vReg(cReg) = Array("Documento", nDoc, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("TpDoc", tDoc, "N"): cReg = cReg + 1
    'vReg(cReg) = Array("Descricao", Descricao, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("Valor", ChkVal(Valor, 0, cDecMoeda), "S"): cReg = cReg + 1
    'vReg(cReg) = Array("CD", cd, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("Saldo", ChkVal(Saldo, 0, cDecMoeda), "S") ': cReg = cReg + 1
    
    'RegistroIncluir "FinanceiroContaHistorico", vReg, cReg
    
    cReg = 0
    vReg(cReg) = Array("Saldo", Saldo, "S")
    RegistroAlterar "FinanceiroConta", vReg, cReg, "id = " & idConta
    
    RegistroExcluir "FinanceiroContaHistorico", "idRegDoc=" & IdReg
 
 
    
    
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
                        "CNPJ: " & txtDoc.Text & vbCrLf & _
                        "Nome: " & cboNome.Text, vbYesNo + vbQuestion) = vbYes Then
                              
                If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                    LimpaFormulario Me
                End If
            End If
    End If
End Sub
Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
            Incluir
        Case "Alterar"
            Alterar
        Case "Excluir"
            Excluir
        Case "Imprimir"
            ImprimirDocumento
        Case "Pesquisar"
            IdReg = 0
            PesquisarRegistro
        Case "Clonar"
            Clonar
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDForm Me, False
            End If
            
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            txtDoc.Enabled = True
        
        Case "Manutenção da Tabela"
            'formManutencaoTabelas.IniciarManutencao Me, "SELECT * FROM Clientes"
            MontarBaseDeDados
    End Select
End Sub
Private Sub ImprimirDocumento()
    Dim docPrint    As String
    
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If
    
    If IdReg = 0 Then Exit Sub
    docPrint = pgDadosTipoDocumento(PgDadosFinanceiroFatura(IdReg).idTpDoc).Impressao
    Select Case Left(docPrint, 2)
        Case "01" 'Boleto Bancario
            If PgDadosConfig.ImpBoleto = 1 Then
                    ImprBB_Pre (IdReg)
                Else
                    BoletoBancario IdReg, True
            End If
        Case "02" 'Duplicata
            impDuplicata (IdReg)
        Case "03"
        Case Else
            MsgBox "Erro ao localizar documento de impressão.", vbInformation, "Aviso"
    End Select
    
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    HDForm Me, False
    HDMenu Me, True
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    'optContas.Item(0).Enabled = True
    'optContas.Item(1).Enabled = True
    dtpQuitacao.Value = Date
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub optContas_Click(Index As Integer)
    'LimpaFormulario Me
    Select Case Index
        Case 0 ' A Pagar
            lblContas.Caption = "A PAGAR"
            lblContas.BackColor = vbRed
            'strTabelaSacado = "Fornecedores"
        Case 1 ' A Receber
            lblContas.Caption = "A RECEBER"
            lblContas.BackColor = vbBlue
            'strTabelaSacado = "Clientes"
    End Select
    
End Sub
Private Function grvRegistro() As Boolean
    Dim vReg(199)   As Variant
    Dim cReg        As Integer 'Contador de Registros
    Dim l           As Integer
    Dim tmp         As Integer
    Dim ContaTp     As String
    cReg = 0
    
    If ValidarDados = False Then
        grvRegistro = False
        Exit Function
    End If
    If optContas.Item(0).Value = True Then
            ContaTp = "P"
        Else
            ContaTp = "R"
    End If
    If chkQuitado.Value = 1 Then
        If Trim(cboContaQuitacao.Text) = "" Then
            MsgBox "Selecione uma conta para lançamento da quitação!", vbInformation, "Aviso"
            grvRegistro = False
            Exit Function
        End If
    End If
    
   
    vReg(cReg) = Array("contaPR", ContaTp, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Emissao", dtpEmissao.Value, "D"): cReg = cReg + 1
    vReg(cReg) = Array("NumFatura", txtNumFatura.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("VlFatura", ChkVal(Trim(txtVlFatura.Text), 0, cDecMoeda), "S"): cReg = cReg + 1
       
    vReg(cReg) = Array("Tabela", strTabelaSacado, "S"): cReg = cReg + 1
    vReg(cReg) = Array("IdSacado", IDSacado, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Nome", cboNome.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("CNPJ", txtDoc.Text, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("Conta", Left(cboConta.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("CentroCusto", Left(cboCentroCustos.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("TpDocumento", Left(cboDocumento.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("PlanoContas", Left(cboPlanoContas.Text, 3), "S"): cReg = cReg + 1
    
    
    Dim tmppc As String
    If Len(Trim(cboPlanoContas.Text)) > 0 Then
            tmppc = Mid(Trim(cboPlanoContas.Text), InStr(Trim(cboPlanoContas.Text), "(") + 1, Len(Trim(cboPlanoContas.Text)))
    
            tmppc = Mid(Trim(tmppc), 1, InStr(Trim(tmppc), ")") - 1)
        Else
            tmppc = ""
    End If
    vReg(cReg) = Array("PlanoContasCodigo", Trim(tmppc), "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("LinhaDigitavel", txtLinhaDigitavel.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("NossoNumero", txtNossoNumero.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Vencimento", dtpVencimento.Value, "D"): cReg = cReg + 1
    vReg(cReg) = Array("NumDuplicata", txtNumDuplicata.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("VlDuplicata", ChkVal(Trim(txtVlDuplicata.Text), 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Multa", txtMulta.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Juros", txtJuros.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("DiasProtesto", txtDiasProtesto.Text, "N"): cReg = cReg + 1
    
    
    vReg(cReg) = Array("idBanco", Left(cboBanco.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Acrescimo", ChkVal(txtAcrescimo.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Abatimento", ChkVal(txtAbatimento.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Deducoes", ChkVal(txtDeducoes.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("MultaMora", ChkVal(txtMultaMora.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("vlCobrado", ChkVal(Trim(txtVlCobrado.Text), 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Obs", txtObs.Text, "S"): cReg = cReg + 1
    If IDSacado <> 0 Then
        vReg(cReg) = Array("ObsBol2", PgDadosCliente(IDSacado).ObsCobBoleto, "S"): cReg = cReg + 1
    End If
    'Quita o Documento
    If chkQuitado.Value = 1 Then
            vReg(cReg) = Array("DataQuitacao", dtpQuitacao.Value, "D"): cReg = cReg + 1
            vReg(cReg) = Array("IdContaQuitacao", Trim(Left(cboContaQuitacao.Text, 3)), "N"): cReg = cReg + 1
        Else
            vReg(cReg) = Array("DataQuitacao", "", "D"): cReg = cReg + 1
            vReg(cReg) = Array("IdContaQuitacao", "", "N"): cReg = cReg + 1
    'IdContaQuitacao
    End If
    
    cReg = cReg - 1
    If IdReg = 0 Then
            IdReg = RegistroIncluir(strTabela, vReg, cReg)
            If IdReg = 0 Then
                    MsgBox "Erro ao Incluir"
                    grvRegistro = False
                Else
                    grvRegistro = True
            End If
        Else
            If RegistroAlterar(strTabela, vReg, cReg, "Id = " & IdReg) = False Then
                    MsgBox "Erro ao Alterar."
                    grvRegistro = False
                Else
                    grvRegistro = True
                
            End If
    End If
    If gerarExtorno = True Then
        ExtornarLancamentoEmConta
    End If
    'MovimentarConta
    If chkQuitado.Value = 1 Then
        MovimentarConta Trim(Left(cboContaQuitacao.Text, 3)), _
                        IIf(ContaTp = "R", "C", "D"), _
                        IdReg, _
                        dtpQuitacao.Value, Trim(txtNumDuplicata.Text), _
                        Left(cboDocumento.Text, 3), _
                        Trim(cboNome.Text), _
                        ChkVal(Trim(txtVlCobrado.Text), 0, cDecMoeda)
    End If
End Function
Private Sub MontarBaseDeDados()
    Dim vDados(199)  As Variant
    Dim contReg     As Integer
    Dim i           As Integer
    
    contReg = 0
    vDados(contReg) = Array("ContaPR", "10", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("Emissao", "10", "D"): contReg = contReg + 1
    vDados(contReg) = Array("NumFatura", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("VlFatura", "100", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("Conta", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("CentroCusto", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("TpDocumento", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("PlanoContas", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("PlanoContasCodigo", "30", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("Tabela", "50", "S"): contReg = contReg + 1
    vDados(contReg) = Array("idSacado", "50", "N"): contReg = contReg + 1
    vDados(contReg) = Array("CNPJ", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Nome", "120", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("CodigoBarras", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("LinhaDigitavel", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("NossoNumero", "30", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("Vencimento", "10", "D"): contReg = contReg + 1
    vDados(contReg) = Array("NumDuplicata", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("VlDuplicata", "30", "DC"): contReg = contReg + 1
    
    vDados(contReg) = Array("Multa", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Juros", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("DiasProtesto", "10", "N"): contReg = contReg + 1
    
    
    vDados(contReg) = Array("IdBanco", "10", "N"): contReg = contReg + 1
    
    vDados(contReg) = Array("Acrescimo", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Abatimento", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Deducoes", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("MultaMora", "30", "S"): contReg = contReg + 1
    vDados(contReg) = Array("VlCobrado", "30", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("DataQuitacao", "30", "D"): contReg = contReg + 1
    vDados(contReg) = Array("IdContaQuitacao", "10", "N"): contReg = contReg + 1
    
    vDados(contReg) = Array("Obs", "2000", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("ObsBol1", "2000", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ObsBol2", "2000", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ObsBol3", "2000", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("ide_NFe", "60", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("IdFixa", "50", "N"): contReg = contReg + 1
    
    contReg = contReg - 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg
    'formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "Fixo"
End Sub
Private Sub txtAbatimento_Change()
    CalcQuitacao
End Sub

Private Sub txtAbatimento_GotFocus()
    txtAbatimento.Text = ChkVal(txtAbatimento, 0, 2)
    
End Sub

Private Sub txtAbatimento_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtAbatimento.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtAbatimento_LostFocus()
    txtAbatimento.Text = ConvMoeda(txtAbatimento)
End Sub

Private Sub txtAcrescimo_Change()
    CalcQuitacao
End Sub

Private Sub txtAcrescimo_GotFocus()
    txtAcrescimo.Text = ChkVal(txtAcrescimo.Text, 0, 2)
End Sub

Private Sub txtAcrescimo_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtAcrescimo.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtAcrescimo_LostFocus()
    txtAcrescimo.Text = ConvMoeda(txtAcrescimo.Text)
End Sub

Private Sub txtDeducoes_Change()
    CalcQuitacao
End Sub

Private Sub txtDeducoes_GotFocus()
    txtDeducoes.Text = ChkVal(txtDeducoes.Text, 0, 2)
End Sub

Private Sub txtDeducoes_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtDeducoes.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtDeducoes_LostFocus()
    txtDeducoes.Text = ConvMoeda(txtDeducoes.Text)
End Sub

Private Sub txtDiasProtesto_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(IsNumeric(Chr(KeyAscii)), KeyAscii, 0)
End Sub

Private Sub txtDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IDSacado = 0
        PesquisarSacado
    End If
End Sub
Private Sub PesquisarSacado()
    
     If Left(cboCadastro.Text, 2) = "00" Then
        Exit Sub
    End If
    
    If IDSacado = 0 Then
        IDSacado = formBuscar.IniciarBusca(strTabelaSacado)
    End If
   
    Select Case strTabelaSacado
        Case "Fornecedores"
            txtDoc.Text = PgDadosFornecedor(IDSacado).Doc
            'cboNome.Clear
            cboNome.Text = PgDadosFornecedor(IDSacado).Nome
                    
        Case "Clientes"
            txtDoc.Text = PgDadosCliente(IDSacado).Doc
            'cboNome.Clear
            cboNome.Text = PgDadosCliente(IDSacado).Nome
                    
        Case "Transportadoras"
            txtDoc.Text = pgDadosTransportadora(IDSacado).CNPJ
            'cboNome.Clear
            cboNome.Text = pgDadosTransportadora(IDSacado).Nome
        Case "RHFuncionarioCadastro"
            txtDoc.Text = PgDadosRhFuncionario(IDSacado).CPF
            'cboNome.Clear
            cboNome.Text = PgDadosRhFuncionario(IDSacado).Nome
        Case Else
            MsgBox "Selecione um tipo de Cadastro"
    End Select
End Sub



Private Sub txtDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PesquisarSacadoCNPJ txtDoc.Text
    End If
End Sub

Private Sub txtJuros_Change()
    CalcQuitacao
End Sub

Private Sub txtJuros_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtJuros.Text, KeyAscii, 3)
End Sub


Private Sub txtMulta_Change()
    CalcQuitacao
End Sub

Private Sub txtMulta_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtMulta.Text, KeyAscii, 3)
End Sub


Private Sub txtMultaMora_Change()
    CalcQuitacao
End Sub

Private Sub txtMultaMora_GotFocus()
    txtMultaMora.Text = ChkVal(txtMultaMora.Text, 0, 2)
End Sub

Private Sub txtMultaMora_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtMultaMora.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtMultaMora_LostFocus()
    txtMultaMora.Text = ConvMoeda(txtMultaMora.Text)
End Sub

Private Sub txtNumFatura_Change()
    If Trim(txtNumFatura.Text) <> "" And optContas(1).Value = True Then
            txtNumDuplicata.Text = txtNumFatura.Text & "-1/1"
        Else
            txtNumDuplicata.Text = txtNumFatura.Text
    End If
End Sub

Private Sub txtVlCobrado_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtVlDuplicata_Change()
    CalcQuitacao
End Sub

Private Sub txtVlDuplicata_GotFocus()
    txtVlDuplicata.Text = ChkVal(txtVlDuplicata.Text, 0, 2)
    txtVlDuplicata.SelStart = 0
    txtVlDuplicata.SelLength = Len(txtVlDuplicata.Text)
End Sub

Private Sub txtVlDuplicata_KeyPress(KeyAscii As Integer)
    If txtVlDuplicata.SelLength = Len(txtVlDuplicata.Text) Then
        txtVlDuplicata.Text = ""
    End If

    KeyAscii = ChkVal(txtVlDuplicata.Text, KeyAscii, cDecMoeda)
    
End Sub


Private Sub txtVlDuplicata_LostFocus()
    txtVlDuplicata.Text = ConvMoeda(txtVlDuplicata.Text)
End Sub

Private Sub txtVlFatura_Change()
    If Trim(txtVlFatura.Text) <> "" Then
        txtVlDuplicata.Text = txtVlFatura.Text
    End If
End Sub

Private Sub txtVlFatura_GotFocus()
    txtVlFatura.Text = ChkVal(txtVlFatura.Text, 0, 2)
End Sub

Private Sub txtVlFatura_KeyPress(KeyAscii As Integer)
    If txtVlFatura.SelLength = Len(txtVlFatura.Text) Then
        txtVlFatura.Text = ""
    End If
    KeyAscii = ChkVal(txtVlFatura.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtVlFatura_LostFocus()
    txtVlFatura.Text = ConvMoeda(txtVlFatura.Text)
End Sub


Private Sub PesquisarSacadoCNPJ(sCNPJ As String)
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim Campo   As String
    If Trim(sCNPJ) = "" Then Exit Sub
    If Trim(strTabelaSacado) = "" Then Exit Sub
    
    
    Select Case strTabelaSacado
        Case "Clientes"
            Campo = "Doc"
        Case "Fornecedores"
            Campo = "Doc"
        Case "Transportador"
            Campo = "CNPJ"
        Case "RHFuncionarioCadastro"
            Campo = "CPF"
        Case Else
            Exit Sub
    End Select

    
    sSQL = "SELECT * FROM " & strTabelaSacado & " WHERE " & Campo & "= '" & sCNPJ & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'Exit Sub
        Else
            Rst.MoveFirst
            IDSacado = Rst.Fields("ID")
            PesquisarSacado
    End If
    Rst.Close
End Sub
Private Sub cboContaQuitacao_DropDown()
    Dim sSQL    As String
    Dim Rst     As Recordset
    cboContaQuitacao.Clear
    sSQL = "SELECT * FROM FinanceiroConta"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma conta cadastrada!", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboContaQuitacao.AddItem ZE(Rst.Fields("ID"), 3) & " - " & Rst.Fields("Agencia") & "/" & Rst.Fields("Conta")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub
Private Function ValidarDados() As Boolean
    If cboDocumento.Text = "" Then
        MsgBox "Selecione o tipo de documento.", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    If dtpEmissao.Value > dtpVencimento.Value Then
        MsgBox "A data de emissão nao pode ser posterior a data de vencimento.", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    If IDSacado = 0 And Trim(strTabelaSacado) <> "" Then
        MsgBox "O Sacado/Cedente não esta vinculado ao banco de dados." & vbCrLf & "Favor verificar!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
        ValidarDados = True
End Function

