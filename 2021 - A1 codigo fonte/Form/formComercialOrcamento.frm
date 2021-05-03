VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formFaturamentoPV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pré Venda"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   13335
   Begin VB.Frame Frame2 
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   120
      TabIndex        =   16
      Top             =   2940
      Width           =   13095
      Begin VB.TextBox txtDescItem 
         Height          =   285
         Left            =   8760
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   1080
         Width           =   1395
      End
      Begin VB.TextBox txtItemID 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   420
         Width           =   495
      End
      Begin VB.Frame Frame3 
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
         TabIndex        =   40
         Top             =   2880
         Width           =   12975
         Begin VB.Frame Frame11 
            Caption         =   "Outros"
            Height          =   675
            Left            =   9582
            TabIndex        =   66
            Top             =   180
            Width           =   1635
            Begin VB.TextBox txtOutros 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   15
               TabIndex        =   67
               Text            =   "Text1"
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Seguro"
            Height          =   675
            Left            =   7885
            TabIndex        =   64
            Top             =   180
            Width           =   1635
            Begin VB.TextBox txtSeguro 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   15
               TabIndex        =   65
               Text            =   "Text1"
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Frete"
            Height          =   675
            Left            =   6188
            TabIndex        =   62
            Top             =   180
            Width           =   1635
            Begin VB.TextBox txtFrete 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   15
               TabIndex        =   63
               Text            =   "Text1"
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Total"
            Height          =   675
            Left            =   11280
            TabIndex        =   60
            Top             =   180
            Width           =   1575
            Begin VB.Label lblTotalPV 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "R$ 0,00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   61
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Valor do IPI"
            Height          =   675
            Left            =   2794
            TabIndex        =   58
            Top             =   180
            Width           =   1635
            Begin VB.Label lblIPI 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "R$ 0,00"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   60
               TabIndex        =   59
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Valor da Mercadoria"
            Height          =   675
            Left            =   1097
            TabIndex        =   56
            Top             =   180
            Width           =   1635
            Begin VB.Label lblMercadoria 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "R$ 0,00"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Itens"
            Height          =   675
            Left            =   60
            TabIndex        =   54
            Top             =   180
            Width           =   975
            Begin VB.Label lblItens 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0000"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   240
               Width           =   795
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Descontos"
            Height          =   675
            Left            =   4491
            TabIndex        =   52
            Top             =   180
            Width           =   1635
            Begin VB.Label lblDesconto 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "R$ 0,00"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   60
               TabIndex        =   53
               Top             =   240
               Width           =   1455
            End
         End
      End
      Begin VB.CommandButton btoRemoverItem 
         Caption         =   "Remover Item"
         Height          =   315
         Left            =   11580
         TabIndex        =   37
         Top             =   540
         Width           =   1395
      End
      Begin VB.CommandButton btoAdicionarItem 
         Caption         =   "Adicionar Item"
         Height          =   375
         Left            =   11640
         TabIndex        =   36
         Top             =   120
         Width           =   1395
      End
      Begin MSFlexGridLib.MSFlexGrid msfgItens 
         Height          =   1395
         Left            =   120
         TabIndex        =   35
         Top             =   1500
         Width           =   12915
         _ExtentX        =   22781
         _ExtentY        =   2461
         _Version        =   393216
         Cols            =   11
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"formComercialOrcamento.frx":0000
      End
      Begin VB.TextBox txtSubTotalProduto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   1080
         Width           =   1515
      End
      Begin VB.TextBox txtValorIPI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7560
         MaxLength       =   15
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1080
         Width           =   1035
      End
      Begin VB.TextBox txtAliquotaIPI 
         Height          =   285
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtTotalProduto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtValorUnitario 
         Height          =   285
         Left            =   2760
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtQuantidade 
         Height          =   315
         Left            =   1260
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cboUnidade 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   2400
         MaxLength       =   250
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   420
         Width           =   8535
      End
      Begin VB.TextBox txtProdutoID 
         Height          =   285
         Left            =   780
         MaxLength       =   10
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Desconto:"
         Height          =   195
         Left            =   8820
         TabIndex        =   50
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "ID:"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label17 
         Caption         =   "Sub Total Produto"
         Height          =   195
         Left            =   4800
         TabIndex        =   34
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label Label16 
         Caption         =   "IPI (Valor):"
         Height          =   255
         Left            =   7620
         TabIndex        =   30
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label15 
         Caption         =   "IPI (%):"
         Height          =   195
         Left            =   6600
         TabIndex        =   29
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Total do Produto"
         Height          =   195
         Left            =   10320
         TabIndex        =   27
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label13 
         Caption         =   "Unidade:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label12 
         Caption         =   "Preço Unitário:"
         Height          =   195
         Left            =   2760
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   1260
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   2400
         TabIndex        =   18
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Referencia:"
         Height          =   195
         Left            =   780
         TabIndex        =   17
         Top             =   180
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   13155
      Begin VB.CheckBox chkFreteConta 
         Caption         =   "Frete: Cliente"
         Height          =   195
         Left            =   6060
         TabIndex        =   49
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtValidade 
         Height          =   285
         Left            =   10260
         MaxLength       =   10
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   1980
         Width           =   1515
      End
      Begin VB.TextBox txtRefCliente 
         Height          =   285
         Left            =   10260
         MaxLength       =   50
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   1620
         Width           =   2775
      End
      Begin VB.TextBox txtPrazoEntrega 
         Height          =   285
         Left            =   10260
         MaxLength       =   60
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   1260
         Width           =   2775
      End
      Begin VB.ComboBox cboTransportadora 
         Height          =   315
         Left            =   1260
         TabIndex        =   39
         Text            =   "Combo1"
         Top             =   1080
         Width           =   4695
      End
      Begin VB.ComboBox cboFormaPagamento 
         Height          =   315
         Left            =   10260
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   900
         Width           =   2775
      End
      Begin VB.TextBox txtObs 
         Height          =   675
         Left            =   1260
         MaxLength       =   65000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "formComercialOrcamento.frx":00C1
         Top             =   1560
         Width           =   6435
      End
      Begin VB.ComboBox cboCondicoesPagamento 
         Height          =   315
         Left            =   10260
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   540
         Width           =   2775
      End
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   10260
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   180
         Width           =   2775
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   1260
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   660
         Width           =   6435
      End
      Begin MSComCtl2.DTPicker dtpEmissao 
         Height          =   315
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55508993
         CurrentDate     =   40517
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1260
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Validade (em dias):"
         Height          =   195
         Left            =   8760
         TabIndex        =   47
         Top             =   2040
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Ref.Cliente:"
         Height          =   195
         Left            =   8520
         TabIndex        =   45
         Top             =   1680
         Width           =   1635
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Prazo de Entrega:"
         Height          =   195
         Left            =   8220
         TabIndex        =   43
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Transportadora:"
         Height          =   195
         Left            =   60
         TabIndex        =   38
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "Forma de Pagamento:"
         Height          =   195
         Left            =   8580
         TabIndex        =   14
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Observações:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Condições de Pagamento:"
         Height          =   195
         Left            =   8220
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Vendedor:"
         Height          =   255
         Left            =   9420
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   540
         TabIndex        =   4
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Data Emissão:"
         Height          =   195
         Left            =   2520
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Orçamento:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   300
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
            Object.ToolTipText     =   "Pesquisar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialOrcamento.frx":00C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialOrcamento.frx":0519
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialOrcamento.frx":0833
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialOrcamento.frx":10C5
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialOrcamento.frx":2317
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialOrcamento.frx":2BF1
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialOrcamento.frx":3483
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialOrcamento.frx":3D15
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialOrcamento.frx":4F67
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialOrcamento.frx":5281
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialOrcamento.frx":559B
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFaturamentoPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim IdReg       As Integer 'ID do Pedido
Dim IdItem      As Integer 'Id dos itens do pedido
Dim strTabela   As String
Dim strTabela2   As String


Private Sub CalcVlPV()
    Dim VlMercadoria    As String
    Dim VlPV     As String
    Dim VlDesconto      As String
    Dim VlIPI           As String
    Dim itens           As Integer
    
    
    For itens = 1 To msfgItens.Rows - 1
        VlMercadoria = Val(ChkVal(msfgItens.TextMatrix(itens, 6), 0, 0)) + Val(ChkVal(VlMercadoria, 0, 0))
        VlIPI = Val(ChkVal(msfgItens.TextMatrix(itens, 8), 0, 0)) + Val(ChkVal(VlIPI, 0, 0))
        VlDesconto = Val(ChkVal(msfgItens.TextMatrix(itens, 9), 0, 0)) + Val(ChkVal(VlDesconto, 0, 0))
        VlPV = Val(ChkVal(msfgItens.TextMatrix(itens, 10), 0, 0)) + Val(ChkVal(VlPV, 0, 0))
    
    Next
    itens = itens - 1
    
    VlPV = Val(ChkVal(VlPV, 0, 0)) + Val(ChkVal(txtFrete.Text, 0, 0)) + Val(ChkVal(txtSeguro.Text, 0, 0)) + Val(ChkVal(txtOutros.Text, 0, 0))
    VlPV = Val(ChkVal(VlPV, 0, 0)) - Val(ChkVal(VlDesconto, 0, 0))
    
    lblItens.Caption = Left(String(5, "0"), 5 - Len(itens)) & itens
    lblMercadoria = ConvMoeda(VlMercadoria)
    lblIPI.Caption = ConvMoeda(VlIPI)
    lblDesconto.Caption = ConvMoeda(VlDesconto)
    lblTotalPV.Caption = ConvMoeda(VlPV)
    'txtFrete.Text = ConvMoeda(ChkVal(IIf(Trim(txtFrete.Text) = "", 0, txtFrete.Text), 0, 0))
    'txtSeguro.Text = ConvMoeda(ChkVal(IIf(Trim(txtSeguro.Text) = "", 0, txtSeguro.Text), 0, 0))
    'txtOutros.Text = ConvMoeda(ChkVal(IIf(Trim(txtOutros.Text) = "", 0, txtOutros.Text), 0, 0))
End Sub

Private Sub ImprimirTeste()
    ImpPV (IdReg)


End Sub

Private Sub LimparGrid()
    msfgItens.Rows = 1
    lblItens.Caption = "0000"
    lblMercadoria.Caption = "R$ 0,00"
    lblIPI.Caption = "R$ 0,00"
    lblTotalPV.Caption = "R$ 0,00"
End Sub

Private Sub LimpProduto()
        txtItemID.Text = ""
        txtProdutoID.Text = ""
        txtDescricao.Text = ""
        cboUnidade.Clear
        txtQuantidade.Text = ""
        txtValorUnitario.Text = ""
        txtSubTotalProduto.Text = ""
        txtAliquotaIPI.Text = ""
        txtValorIPI.Text = ""
        txtDescItem.Text = ""
        txtTotalProduto.Text = ""

End Sub

Private Sub MontarBaseDeDados()
    Dim vDados(20)  As Variant
    Dim contReg     As Integer
    Dim I           As Integer
    
    contReg = 0
    
    'cabecario do Pedido
    
    vDados(contReg) = Array("Emissao", "10", "D")
    contReg = contReg + 1
    vDados(contReg) = Array("IdCliente", "10", "N")
    contReg = contReg + 1
    vDados(contReg) = Array("Cliente", "100", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("Transportadora", "100", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("Vendedor", "100", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("CondicoesPagamento", "100", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("FormaPagamento", "100", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("PrazoEntrega", "100", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("RefCliente", "50", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("Obs", "65000", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("Itens", "10", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("VlMercadoria", "50", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("VlIPI", "50", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("Frete", "10", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("Seguro", "10", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("Outros", "10", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("Desconto", "10", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("VlTotalPV", "50", "S")
    contReg = contReg + 1
    vDados(contReg) = Array("FreteConta", "1", "N")
    contReg = contReg + 1
    vDados(contReg) = Array("Validade", "10", "S")
    'contReg = contReg + 1
    
    
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg
        
    'Outros Dados
    contReg = 0
    'Set vDados = Empty
    'For I = 1 To msfgItens.Rows - 1
        vDados(contReg) = Array("idPV", "50", "N")
        contReg = contReg + 1
        vDados(contReg) = Array("idProduto", "100", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("referencia", "100", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("descricao", "250", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("unidade", "100", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("quantidade", "100", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("ValorUnitario", "100", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("SubTotal", "100", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("ipi", "10", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("VlIPI", "30", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("DescItem", "30", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("TotalProduto", "30", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("EstoqueVlCusto", "30", "S")
        contReg = contReg + 1
        vDados(contReg) = Array("EstoqueUnidade", "30", "S")
        'contReg = contReg + 1
    'Next
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "Itens"
End Sub
Private Function grvRegistro() As Boolean
    Dim vReg(199)   As Variant
    Dim cReg        As Integer 'Contador de Registros
    Dim l           As Integer
    Dim tmp         As Integer
    cReg = 0
    If ValidarPV = False Then
        grvRegistro = False
        Exit Function
    End If
    
    vReg(cReg) = Array("Emissao", dtpEmissao.Value, "D")
    cReg = cReg + 1
    vReg(cReg) = Array("IdCliente", Trim(Left(cboCliente.Text, 6)), "N")
    cReg = cReg + 1
    vReg(cReg) = Array("Cliente", Trim(Mid(cboCliente.Text, 10, Len(cboCliente.Text))), "S")
    cReg = cReg + 1
    vReg(cReg) = Array("Transportadora", Trim(Left(cboTransportadora.Text, 6)), "S")
    cReg = cReg + 1
    vReg(cReg) = Array("Vendedor", Left(Trim(cboVendedor.Text), 4), "S")
    cReg = cReg + 1
    vReg(cReg) = Array("CondicoesPagamento", Trim(Left(cboCondicoesPagamento.Text, 3)), "S")
    cReg = cReg + 1
    vReg(cReg) = Array("FormaPagamento", Trim(Left(cboFormaPagamento.Text, 3)), "S")
    cReg = cReg + 1
    vReg(cReg) = Array("PrazoEntrega", txtPrazoEntrega.Text, "S")
    cReg = cReg + 1
    vReg(cReg) = Array("RefCliente", txtRefCliente.Text, "S")
    cReg = cReg + 1
    vReg(cReg) = Array("Obs", txtObs.Text, "S")
    cReg = cReg + 1
    vReg(cReg) = Array("Itens", lblItens.Caption, "S")
    cReg = cReg + 1
    vReg(cReg) = Array("VlMercadoria", lblMercadoria.Caption, "S")
    cReg = cReg + 1
    vReg(cReg) = Array("vlIPI", lblIPI.Caption, "S")
    cReg = cReg + 1
    vReg(cReg) = Array("Desconto", lblDesconto.Caption, "S")
    cReg = cReg + 1
    vReg(cReg) = Array("Frete", txtFrete.Text, "S")
    cReg = cReg + 1
    vReg(cReg) = Array("Seguro", txtSeguro.Text, "S")
    cReg = cReg + 1
    vReg(cReg) = Array("Outros", txtOutros.Text, "S")
    cReg = cReg + 1
    vReg(cReg) = Array("vlTotalPV", lblTotalPV.Caption, "S")
    cReg = cReg + 1
    vReg(cReg) = Array("FreteConta", chkFreteConta.Value, "N")
    cReg = cReg + 1
    vReg(cReg) = Array("Validade", txtValidade.Text, "S")
    'cReg = cReg + 1
    
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
    'Gravar os dados do grig
    
    If RegistroExcluir(strTabela2, "idPV = " & IdReg) = False Then
        MsgBox "Erro interno - Ao apagar os dados para novo registro Produtos"
        Exit Function
    End If

    cReg = 0
    For l = 1 To msfgItens.Rows - 1
        vReg(cReg) = Array("idPV", IdReg, "S")
        cReg = cReg + 1
        vReg(cReg) = Array("idProduto", msfgItens.TextMatrix(l, 0), "S")
        cReg = cReg + 1
        vReg(cReg) = Array("referencia", msfgItens.TextMatrix(l, 1), "S")
        cReg = cReg + 1
        vReg(cReg) = Array("Descricao", msfgItens.TextMatrix(l, 2), "S")
        cReg = cReg + 1
        vReg(cReg) = Array("unidade", msfgItens.TextMatrix(l, 3), "S")
        cReg = cReg + 1
        vReg(cReg) = Array("quantidade", msfgItens.TextMatrix(l, 4), "S")
        cReg = cReg + 1
        vReg(cReg) = Array("ValorUnitario", msfgItens.TextMatrix(l, 5), "S")
        cReg = cReg + 1
        vReg(cReg) = Array("SubTotal", msfgItens.TextMatrix(l, 6), "S")
        cReg = cReg + 1
        vReg(cReg) = Array("ipi", msfgItens.TextMatrix(l, 7), "S")
        cReg = cReg + 1
        vReg(cReg) = Array("vlipi", msfgItens.TextMatrix(l, 8), "S")
        cReg = cReg + 1
        vReg(cReg) = Array("DescItem", msfgItens.TextMatrix(l, 9), "S")
        cReg = cReg + 1
        vReg(cReg) = Array("TotalProduto", msfgItens.TextMatrix(l, 10), "S")
        cReg = cReg + 1
        vReg(cReg) = Array("EstoqueVlCusto", pgDadosEstoqueProduto(msfgItens.TextMatrix(l, 0)).VlCusto, "S")
        cReg = cReg + 1
        vReg(cReg) = Array("EstoqueUnidade", pgDadosEstoqueProduto(msfgItens.TextMatrix(l, 0)).Unidade, "S")
        'cReg = cReg + 1
        tmp = RegistroIncluir(strTabela2, vReg, cReg)
        If tmp = 0 Then
                MsgBox "Erro ao Incluir o Produto"
                grvRegistro = False
            Else
                grvRegistro = True
                cReg = 0
        End If
    Next
 
End Function
Private Sub PesquisarRegistro(Optional Id As Integer)
    Dim psqTMP  As String
    If Trim(Id) = 0 Then
            psqTMP = FormBusca.IniciarBusca(strTabela)
            IdReg = IIf(psqTMP = "", 0, psqTMP)
        Else
            IdReg = Id
    End If
    
    If IdReg = 0 Then
            LimpForm 'me
            Exit Sub
        Else
            Dim sSQL    As String
            Dim Rst     As Recordset
    
            sSQL = "SELECT * FROM " & strTabela & " WHERE Id = " & IdReg
    
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    MsgBox "Registro nao encontrado."
                    LimpForm
                Else
                    txtID.Text = Rst.Fields("id")
                    dtpEmissao.Value = Rst.Fields("Emissao")
                    
                    cboCliente.Clear
                    cboCliente.AddItem IIf(IsNull(Rst.Fields("Cliente")), " ", Left(String(6, "0"), 6 - Len(Rst.Fields("IdCliente"))) & Rst.Fields("IdCliente") & " - " & Rst.Fields("Cliente"))
                    cboCliente.Text = cboCliente.List(0)
                    
                    cboTransportadora.Clear
                    If Not IsNull(Rst.Fields("Transportadora")) Then
                        cboTransportadora.AddItem IIf(IsNull(Rst.Fields("Transportadora")), " ", Left(String(6, "0"), 6 - Len(Rst.Fields("Transportadora"))) & Rst.Fields("Transportadora") & " - " & pgDadosTransportadora(Rst.Fields("Transportadora")).Nome)
                        cboTransportadora.Text = cboTransportadora.List(0)
                    End If
                    
                    cboVendedor.Clear
                    If Not IsNull(Rst.Fields("Vendedor")) Then
                        cboVendedor.AddItem Rst.Fields("Vendedor") & " - " & PgDadosRhFuncionario(Rst.Fields("Vendedor")).Nome
                        cboVendedor.Text = cboVendedor.List(0)
                    End If
                    
                    cboCondicoesPagamento.Clear
                    If Not IsNull(Rst.Fields("CondicoesPagamento")) Then
                        cboCondicoesPagamento.AddItem Rst.Fields("CondicoesPagamento") & " - " & pgDescrCondPag(Rst.Fields("CondicoesPagamento"))
                        cboCondicoesPagamento.Text = cboCondicoesPagamento.List(0)
                    End If
                    
                    cboFormaPagamento.Clear
                    If Not IsNull(Rst.Fields("FormaPagamento")) Then
                        cboFormaPagamento.AddItem Rst.Fields("FormaPagamento") & " - " & pgDescrTipoDoc((Rst.Fields("FormaPagamento")))
                        cboFormaPagamento.Text = cboFormaPagamento.List(0)
                    End If
                    
                    txtPrazoEntrega.Text = IIf(IsNull(Rst.Fields("PrazoEntrega")), " ", Rst.Fields("PrazoEntrega"))
                    
                    txtRefCliente.Text = IIf(IsNull(Rst.Fields("RefCliente")), " ", Rst.Fields("RefCliente"))
                    txtObs.Text = IIf(IsNull(Rst.Fields("Obs")), "", Rst.Fields("Obs"))
                    txtValidade.Text = IIf(IsNull(Rst.Fields("Validade")), "", Rst.Fields("Validade"))
                    chkFreteConta.Value = IIf(IsNull(Rst.Fields("FreteConta")), "0", Rst.Fields("Freteconta"))
                    
                    lblDesconto.Caption = IIf(IsNull(Rst.Fields("Desconto")), ConvMoeda("0,00"), Rst.Fields("Desconto"))
                    txtFrete.Text = IIf(IsNull(Rst.Fields("Frete")), ConvMoeda("0,00"), Rst.Fields("Frete"))
                    txtSeguro.Text = IIf(IsNull(Rst.Fields("Seguro")), ConvMoeda("0,00"), Rst.Fields("Seguro"))
                    txtOutros.Text = IIf(IsNull(Rst.Fields("outros")), ConvMoeda("0,00"), Rst.Fields("Outros"))
            End If
            Rst.Close
            
            'Carregar Corpo do PV
            sSQL = "SELECT * FROM " & strTabela2 & " WHERE IdPV = " & IdReg
    
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    'MsgBox "Registro nao encontrado."
                    LimparGrid
                Else
                    Rst.MoveFirst
                    With msfgItens
                        .Rows = 1
                        Do Until Rst.EOF
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, 0) = Rst.Fields("IDproduto")
                            .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rst.Fields("Referencia")), " ", Rst.Fields("Referencia"))
                            .TextMatrix(.Rows - 1, 2) = Rst.Fields("Descricao")
                            .TextMatrix(.Rows - 1, 3) = Rst.Fields("Unidade")
                            .TextMatrix(.Rows - 1, 4) = Rst.Fields("Quantidade")
                            .TextMatrix(.Rows - 1, 5) = Rst.Fields("ValorUnitario")
                            .TextMatrix(.Rows - 1, 6) = Rst.Fields("SubTotal")
                            .TextMatrix(.Rows - 1, 7) = Rst.Fields("IPI")
                            .TextMatrix(.Rows - 1, 8) = Rst.Fields("vlIPI")
                            .TextMatrix(.Rows - 1, 9) = IIf(IsNull(Rst.Fields("DescItem")), "0,00", Rst.Fields("DescItem"))
                            .TextMatrix(.Rows - 1, 10) = IIf(IsNull(Rst.Fields("TotalProduto")), "0,00", Rst.Fields("TotalProduto"))
                            Rst.MoveNext
                        Loop
                    End With
            End If
            Rst.Close
            CalcVlPV
    End If
End Sub

Private Sub cboCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 114 Then
        PesquisarCliente
    End If
End Sub

Private Sub cboTransportadora_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarTransp
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub PesquisarProduto(Optional Id As String)
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    If Trim(Id) = "" Then
            Id = FormBusca.IniciarBusca("EstoqueProduto")
            If Trim(Id) = "" Then Exit Sub
    End If
    
    sSQL = "SELECT * FROM EstoqueProduto WHERE Id = " & Id
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado"
            LimpProduto
        Else
            Rst.MoveFirst
            txtItemID.Text = Id
            txtProdutoID.Text = IIf(IsNull(Rst.Fields("Referencia")), "", Rst.Fields("Referencia"))
            txtDescricao.Text = IIf(IsNull(Rst.Fields("Descricao")), "", Rst.Fields("Descricao"))
            cboUnidade.AddItem IIf(IsNull(Rst.Fields("Unidade")), " ", Rst.Fields("Unidade"))
            cboUnidade.Text = cboUnidade.List(0)
            txtQuantidade.Text = "0"
            txtValorUnitario.Text = IIf(IsNull(Rst.Fields("preco")), "0.00", ConvMoeda(Rst.Fields("preco")))
            txtAliquotaIPI.Text = IIf(IsNull(Rst.Fields("ipialiquota")), "0.00", Rst.Fields("ipialiquota"))
            txtDescItem.Text = ConvMoeda("0")
    End If
    Rst.Close
End Sub
Private Sub PesquisarCliente(Optional Id As String)
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim IdCliente   As String
    
    'fgrProdutos.Rows = 1
    If Trim(Id) = "" Then
            IdCliente = FormBusca.IniciarBusca("Clientes")
            If Trim(IdCliente) = "" Then Exit Sub
        Else
            IdCliente = Id
    End If
    
    sSQL = "SELECT * FROM Clientes WHERE Id = " & IdCliente
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado"
            LimpProduto
        Else
            Rst.MoveFirst
            'txtItemID.Text = idCliente
            'txtProdutoID.Text = IIf(IsNull(Rst.Fields("Referencia")), "", Rst.Fields("Referencia"))
            'txtDescricao.Text = IIf(IsNull(Rst.Fields("Descricao")), "", Rst.Fields("Descricao"))
            cboCliente.AddItem Left(String(6, "0"), 6 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                               Rst.Fields("Nome")
            cboCliente.Text = cboCliente.List(0)
            'txtQuantidade.Text = "0"
            'txtValorUnitario.Text = IIf(IsNull(Rst.Fields("preco")), "0.00", ConvMoeda(Rst.Fields("preco")))
            'txtAliquotaIPI.Text = IIf(IsNull(Rst.Fields("ipialiquota")), "0.00", Rst.Fields("ipialiquota"))
    End If
    Rst.Close
End Sub


Private Sub btoAdicionarItem_Click()
    Dim l As Integer
    If Trim(txtItemID.Text) = "" Or txtItemID.Text = "0" Then
        MsgBox "Selecione um Produto do Estoque!", vbInformation, "Aviso"
        Exit Sub
    End If
    With msfgItens
        If IdItem = 0 Then
                .Rows = .Rows + 1
                l = .Rows - 1
            Else
                l = .Row
        End If
        .TextMatrix(l, 0) = txtItemID.Text
        .TextMatrix(l, 1) = txtProdutoID.Text
        .TextMatrix(l, 2) = txtDescricao.Text
        .TextMatrix(l, 3) = cboUnidade.Text
        .TextMatrix(l, 4) = txtQuantidade.Text
        .TextMatrix(l, 5) = ConvMoeda(txtValorUnitario.Text)
        .TextMatrix(l, 6) = txtSubTotalProduto.Text
        .TextMatrix(l, 7) = IIf(Trim(txtAliquotaIPI.Text) = "", "0", txtAliquotaIPI.Text)
        .TextMatrix(l, 8) = txtValorIPI.Text
        .TextMatrix(l, 9) = txtDescItem.Text
        .TextMatrix(l, 10) = txtTotalProduto.Text
    End With
    IdItem = 0
    LimpProduto
    CalcVlPV
End Sub

Private Sub btoRemoverItem_Click()
    If MsgBox("Deseja realmente remover este item?", vbYesNo, "Removendo Item do Pedido") = vbYes Then
        If msfgItens.Rows = 2 Then
                msfgItens.Rows = 1
            Else
                msfgItens.RemoveItem msfgItens.Row
        End If
    End If
    CalcVlPV
End Sub


Private Sub cboCliente_DropDown()
    Dim Rst As Recordset
    
    Set Rst = RegistroBuscar("SELECT * FROM Clientes WHERE Nome LIKE '" & cboCliente.Text & "%'")
    If Rst.BOF And Rst.EOF Then
            cboCliente.Clear
            Exit Sub
        Else
            cboCliente.Clear
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCliente.AddItem Left(String(6, "0"), 6 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & _
                " - " & _
                Rst.Fields("Nome")
                Rst.MoveNext
            Loop
    End If

End Sub
Private Sub cbotransportadora_DropDown()
    Dim Rst As Recordset
    
    Set Rst = RegistroBuscar("SELECT * FROM Transportadoras WHERE Nome LIKE '" & cboTransportadora.Text & "%'")
    If Rst.BOF And Rst.EOF Then
            cboTransportadora.Clear
            Exit Sub
        Else
            cboTransportadora.Clear
            Rst.MoveFirst
            Do Until Rst.EOF
                cboTransportadora.AddItem Left(String(6, "0"), 6 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & _
                " - " & _
                Rst.Fields("Nome")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub PesquisarTransp(Optional Id As String)
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim IdTransp    As String
    
    'fgrProdutos.Rows = 1
    If Trim(Id) = "" Then
            IdTransp = FormBusca.IniciarBusca("Transportadoras")
            If Trim(IdTransp) = "" Then Exit Sub
        Else
            IdTransp = Id
    End If
    
    sSQL = "SELECT * FROM Transportadoras WHERE Id = " & IdTransp
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado"
            LimpProduto
        Else
            Rst.MoveFirst
            'txtItemID.Text = idTransp
            'txtProdutoID.Text = IIf(IsNull(Rst.Fields("Referencia")), "", Rst.Fields("Referencia"))
            'txtDescricao.Text = IIf(IsNull(Rst.Fields("Descricao")), "", Rst.Fields("Descricao"))
            cboTransportadora.AddItem Left(String(6, "0"), 6 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                               Rst.Fields("Nome")
            cboTransportadora.Text = cboTransportadora.List(0)
            'txtQuantidade.Text = "0"
            'txtValorUnitario.Text = IIf(IsNull(Rst.Fields("preco")), "0.00", ConvMoeda(Rst.Fields("preco")))
            'txtAliquotaIPI.Text = IIf(IsNull(Rst.Fields("ipialiquota")), "0.00", Rst.Fields("ipialiquota"))
    End If
    Rst.Close
End Sub
'Private Sub cboTransp_KeyPress(KeyAscii As Integer)
'    If KeyCode = 114 Then
'        PesquisarRegistro
'    End If
'End Sub

Private Sub cboCondicoesPagamento_DropDown()
    Dim Rst As Recordset
    cboCondicoesPagamento.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCondicoespagamento")
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCondicoesPagamento.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                              Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
            
End Sub


Private Sub cboFormaPagamento_DropDown()
    Dim Rst As Recordset
    cboFormaPagamento.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento")
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboFormaPagamento.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                          Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If

End Sub



Private Sub cboUnidade_DropDown()
  Dim Rst As Recordset
    cboUnidade.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueUnidadeMedida")
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboUnidade.AddItem Rst.Fields("Sigla")
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboVendedor_DropDown()
    Dim Rst As Recordset
    cboVendedor.Clear
    Set Rst = RegistroBuscar("SELECT * FROM RHFuncionarioCadastro")
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboVendedor.AddItem Left(String(4, "0"), 4 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & " - " & Rst.Fields("Nome")
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub Form_Load()
    LimpForm
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    strTabela2 = strTabela & "itens"
    HDForm Me, False
    HDMenu Me, True
    txtID.Enabled = True
    IdReg = 0
    
End Sub

Private Sub msfgItens_DblClick()
    If msfgItens.TextMatrix(msfgItens.Row, 0) = "" Then Exit Sub
    IdItem = msfgItens.TextMatrix(msfgItens.Row, 0)
    
    txtItemID.Text = msfgItens.TextMatrix(msfgItens.Row, 0)
    txtProdutoID.Text = msfgItens.TextMatrix(msfgItens.Row, 1)
    txtDescricao.Text = msfgItens.TextMatrix(msfgItens.Row, 2)
    cboUnidade.AddItem IIf(Trim(msfgItens.TextMatrix(msfgItens.Row, 3)) = "", ".", msfgItens.TextMatrix(msfgItens.Row, 3))
    cboUnidade.Text = cboUnidade.List(0)
    txtQuantidade.Text = msfgItens.TextMatrix(msfgItens.Row, 4)
    txtValorUnitario.Text = ChkVal(msfgItens.TextMatrix(msfgItens.Row, 5), 0, 0)
    txtSubTotalProduto.Text = msfgItens.TextMatrix(msfgItens.Row, 6)
    txtAliquotaIPI.Text = msfgItens.TextMatrix(msfgItens.Row, 7)
    txtValorIPI.Text = msfgItens.TextMatrix(msfgItens.Row, 8)
    txtDescItem.Text = msfgItens.TextMatrix(msfgItens.Row, 9)
    txtTotalProduto.Text = msfgItens.TextMatrix(msfgItens.Row, 10)

End Sub






Private Sub msfgItens_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        btoRemoverItem_Click
    End If
End Sub


Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
            IdReg = 0
            HDMenu Me, False
            HDForm Me, True
            LimpForm
            msfgItens.Rows = 1
            txtID.Enabled = False
            
        Case "Alterar"
            If IdReg = 0 Then
                MsgBox "Selecione uma Registro"
                Exit Sub
            End If
            HDForm Me, True
            HDMenu Me, False
            txtID.Enabled = False
        Case "Excluir"
            If IdReg = 0 Then
                    MsgBox "Selecione um Registro"
                    Exit Sub
                Else
                    If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                               vbCrLf & _
                               "Descrição.: " & txtDescricao.Text, vbYesNo + vbCritical) = vbYes Then
                        If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                            LimpForm
                        End If
                    End If
            End If
            
        Case "Imprimir"
            ImprimirTeste
        Case "Pesquisar"
            PesquisarRegistro
            
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDForm Me, False
                'LimpForm
                txtID.Enabled = True
            End If
            
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpForm
            txtID.Enabled = True
        Case "Manutenção da Tabela"
            'formManutencaoTabelas.IniciarManutencao Me
            MontarBaseDeDados
    End Select
End Sub
Private Sub LimpForm()
    LimpaFormulario Me
    dtpEmissao.Value = Date
    'txtID.Enabled = False
    chkFreteConta.Value = 0
End Sub










Private Sub txtAliquotaIPI_Change()
    CalcVlItem
End Sub

Private Sub txtAliquotaIPI_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtAliquotaIPI.Text, KeyAscii, 2)
End Sub

Private Sub txtDescItem_Change()
    CalcVlItem
End Sub

Private Sub txtDescItem_GotFocus()
        txtDescItem.Text = ChkVal(txtDescItem.Text, 0, 0)
End Sub

Private Sub txtDescItem_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtDescItem.Text, KeyAscii, CDecMoeda)
End Sub

Private Sub txtDescItem_LostFocus()
    txtDescItem.Text = ConvMoeda(txtDescItem.Text)
End Sub

Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
End Sub






Private Sub txtFrete_Change()
    CalcVlPV
End Sub

Private Sub txtFrete_GotFocus()
        txtFrete.Text = ChkVal(txtFrete.Text, 0, 0)

End Sub


Private Sub txtFrete_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtFrete.Text, KeyAscii, CDecMoeda)
End Sub


Private Sub txtFrete_LostFocus()
    txtFrete.Text = ConvMoeda(IIf(txtFrete.Text = "", 0, txtFrete.Text))
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PesquisarRegistro (txtID.Text)
    End If
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub

Private Sub txtItemID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarProduto
    End If

End Sub


Private Sub txtItemID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PesquisarProduto (Trim(txtItemID.Text))
    End If
End Sub


Private Sub txtOutros_Change()
    CalcVlPV
End Sub

Private Sub txtOutros_GotFocus()
    txtOutros.Text = ChkVal(txtOutros.Text, 0, 0)

End Sub


Private Sub txtOutros_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtOutros.Text, KeyAscii, CDecMoeda)
End Sub


Private Sub txtOutros_LostFocus()
    txtOutros.Text = ConvMoeda(IIf(txtOutros.Text = "", 0, txtOutros.Text))
End Sub

Private Sub txtProdutoID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarProduto
    End If
End Sub

Private Sub txtQuantidade_Change()
    CalcVlItem
End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtQuantidade.Text, KeyAscii, CDecQtd)
End Sub



Private Sub txtSeguro_Change()
    CalcVlPV
End Sub

Private Sub txtSeguro_GotFocus()
    txtSeguro.Text = ChkVal(txtSeguro.Text, 0, 0)

End Sub


Private Sub txtSeguro_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtSeguro.Text, KeyAscii, CDecMoeda)
End Sub


Private Sub txtSeguro_LostFocus()
    txtSeguro.Text = ConvMoeda(IIf(txtSeguro.Text = "", 0, txtSeguro.Text))
End Sub

Private Sub txtSubTotalProduto_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtTotalProduto_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtValidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub


Private Sub txtValorIPI_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub txtValorUnitario_Change()
    CalcVlItem
End Sub

Private Sub txtValorUnitario_GotFocus()
    txtValorUnitario.Text = ChkVal(txtValorUnitario.Text, 0, 0)
End Sub

Private Sub txtValorUnitario_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtValorUnitario.Text, KeyAscii, CDecMoeda)
End Sub
Private Sub CalcVlItem()
    Dim SubTotalItem    As String
    Dim IPIItem       As String
    Dim TotalItem       As String
    
    SubTotalItem = Val(ChkVal(txtQuantidade.Text, 0, 0)) * Val(ChkVal(txtValorUnitario.Text, 0, 0))
    IPIItem = (Val(ChkVal(SubTotalItem, 0, 0)) * Val(ChkVal(txtAliquotaIPI.Text, 0, 0))) / 100
    TotalItem = (Val(ChkVal(SubTotalItem, 0, 0)) + Val(ChkVal(IPIItem, 0, 0))) - Val(ChkVal(txtDescItem.Text, 0, 0))
    

    txtSubTotalProduto.Text = ConvMoeda(SubTotalItem)
    txtValorIPI.Text = ConvMoeda(IPIItem)
    txtTotalProduto.Text = ConvMoeda(TotalItem)
End Sub

Private Sub txtValorUnitario_LostFocus()
    txtValorUnitario.Text = ConvMoeda(txtValorUnitario.Text)
End Sub
Private Function ValidarPV() As Boolean
    ValidarPV = False
    If Trim(cboCliente.Text) = "" Then
        MsgBox "Favor selecionar um Cliente!", vbExclamation, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If Trim(cboTransportadora.Text) = "" Then
        MsgBox "Favor selecionar uma Transportadora!", vbExclamation, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If Trim(cboVendedor.Text) = "" Then
        MsgBox "Favor selecionar um Vendedor!", vbExclamation, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If Trim(cboCondicoesPagamento.Text) = "" Then
        MsgBox "Favor selecionar uma condicao de pagamento!", vbExclamation, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If Trim(cboFormaPagamento.Text) = "" Then
        MsgBox "Favor selecionar um forma de pagamento!", vbExclamation, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If Trim(txtValidade.Text) = "" Then
        MsgBox "Favor informar o Prazo de Validade.", vbExclamation, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If msfgItens.Rows = 1 Then
        MsgBox "Favor informar pelo menos um item.", vbExclamation, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    ValidarPV = True
End Function

