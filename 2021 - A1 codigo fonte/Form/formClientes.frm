VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9810
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   7140
      TabIndex        =   6
      Top             =   480
      Width           =   2535
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1635
      End
      Begin VB.ComboBox cboPessoa 
         Height          =   315
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label42 
         Caption         =   "Status:"
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label33 
         Caption         =   "Pessoa:"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   6975
      Begin VB.TextBox txtFant 
         Height          =   285
         Left            =   1260
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   840
         Width           =   5535
      End
      Begin VB.TextBox txtDoc 
         Height          =   285
         Left            =   1260
         MaxLength       =   14
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   180
         Width           =   2295
      End
      Begin VB.TextBox txtxNome 
         Height          =   285
         Left            =   1260
         MaxLength       =   60
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   540
         Width           =   5535
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ/CPF:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome Fantasia:"
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   195
         Left            =   540
         TabIndex        =   1
         Top             =   540
         Width           =   675
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9810
      _ExtentX        =   17304
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
               Picture         =   "formClientes.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formClientes.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formClientes.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formClientes.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formClientes.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formClientes.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formClientes.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formClientes.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formClientes.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formClientes.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formClientes.frx":54D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4095
      Left            =   60
      TabIndex        =   13
      Top             =   1800
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7223
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dados do Cliente"
      TabPicture(0)   =   "formClientes.frx":58CB
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Histórico de Notas Fiscais"
      TabPicture(1)   =   "formClientes.frx":58E7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   3555
         Left            =   -74880
         TabIndex        =   14
         Top             =   60
         Width           =   9375
         Begin VB.CommandButton btFinanceiro 
            Caption         =   "&Financeiro"
            DragIcon        =   "formClientes.frx":5903
            Height          =   375
            Left            =   8040
            TabIndex        =   100
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton btoAtualizarHist 
            Height          =   375
            Left            =   1440
            Picture         =   "formClientes.frx":75CD
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtAno 
            Height          =   285
            Left            =   600
            MaxLength       =   4
            TabIndex        =   98
            Text            =   "Text1"
            Top             =   180
            Width           =   795
         End
         Begin MSFlexGridLib.MSFlexGrid msfgHist 
            Height          =   2895
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Clique duas vezes para visualizar a Nota Fiscal..."
            Top             =   540
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   5106
            _Version        =   393216
            Cols            =   7
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            FormatString    =   $"formClientes.frx":7CB7
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Ano:"
            Height          =   195
            Left            =   180
            TabIndex        =   97
            Top             =   240
            Width           =   375
         End
      End
      Begin TabDlg.SSTab SSt 
         Height          =   3555
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   6271
         _Version        =   393216
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Basico"
         TabPicture(0)   =   "formClientes.frx":7D5F
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Entrega"
         TabPicture(1)   =   "formClientes.frx":7D7B
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame9"
         Tab(1).Control(1)=   "Frame4"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Cobrança"
         TabPicture(2)   =   "formClientes.frx":7D97
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame11"
         Tab(2).Control(1)=   "Frame13"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Contatos"
         TabPicture(3)   =   "formClientes.frx":7DB3
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame12"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Outros"
         TabPicture(4)   =   "formClientes.frx":7DCF
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "Frame10"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "Frame8"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "Frame6"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "Frame7"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).ControlCount=   4
         Begin VB.Frame Frame3 
            Height          =   3015
            Left            =   -74880
            TabIndex        =   78
            Top             =   360
            Width           =   9015
            Begin VB.TextBox txtFone 
               Height          =   315
               Left            =   1140
               MaxLength       =   10
               TabIndex        =   87
               Text            =   "Text1"
               Top             =   2520
               Width           =   2955
            End
            Begin VB.TextBox txteMail 
               Height          =   285
               Left            =   1140
               TabIndex        =   86
               Text            =   "Text1"
               Top             =   2160
               Width           =   3915
            End
            Begin VB.TextBox txtCEP 
               Height          =   285
               Left            =   1140
               MaxLength       =   8
               TabIndex        =   85
               Text            =   "Text1"
               Top             =   1800
               Width           =   2175
            End
            Begin VB.ComboBox cboxMun 
               Height          =   315
               Left            =   3180
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   1380
               Width           =   2655
            End
            Begin VB.ComboBox cboUF 
               Height          =   315
               Left            =   1140
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   1380
               Width           =   915
            End
            Begin VB.TextBox txtxBairro 
               Height          =   285
               Left            =   1140
               MaxLength       =   60
               TabIndex        =   82
               Text            =   "Text1"
               Top             =   1020
               Width           =   2955
            End
            Begin VB.TextBox txtxCpl 
               Height          =   285
               Left            =   4620
               MaxLength       =   60
               TabIndex        =   81
               Text            =   "Text1"
               Top             =   600
               Width           =   3255
            End
            Begin VB.TextBox txtNro 
               Height          =   285
               Left            =   1140
               MaxLength       =   60
               TabIndex        =   80
               Text            =   "Text1"
               Top             =   600
               Width           =   2175
            End
            Begin VB.TextBox txtxLgr 
               Height          =   285
               Left            =   1140
               MaxLength       =   60
               TabIndex        =   79
               Text            =   "Text1"
               Top             =   240
               Width           =   6735
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Caption         =   "Telefone:"
               Height          =   195
               Left            =   420
               TabIndex        =   96
               Top             =   2580
               Width           =   675
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Caption         =   "e-mail:"
               Height          =   195
               Left            =   660
               TabIndex        =   95
               Top             =   2220
               Width           =   435
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   "CEP:"
               Height          =   195
               Left            =   600
               TabIndex        =   94
               Top             =   1860
               Width           =   495
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               Caption         =   "UF:"
               Height          =   195
               Left            =   600
               TabIndex        =   93
               Top             =   1440
               Width           =   495
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Municipio:"
               Height          =   255
               Left            =   2340
               TabIndex        =   92
               Top             =   1440
               Width           =   795
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Bairro:"
               Height          =   255
               Left            =   360
               TabIndex        =   91
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "Complemento:"
               Height          =   195
               Left            =   3480
               TabIndex        =   90
               Top             =   660
               Width           =   1035
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Número:"
               Height          =   255
               Left            =   240
               TabIndex        =   89
               Top             =   660
               Width           =   855
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Endereço:"
               Height          =   255
               Left            =   240
               TabIndex        =   88
               Top             =   300
               Width           =   855
            End
         End
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   -74880
            TabIndex        =   60
            Top             =   420
            Width           =   9075
            Begin VB.TextBox txtEntregaCEP 
               Height          =   285
               Left            =   7380
               MaxLength       =   8
               TabIndex        =   69
               Text            =   "Text1"
               Top             =   1380
               Width           =   1515
            End
            Begin VB.ComboBox cboEntregaMun 
               Height          =   315
               Left            =   2940
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   1380
               Width           =   3135
            End
            Begin VB.ComboBox cboEntregaUF 
               Height          =   315
               Left            =   1020
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   1380
               Width           =   915
            End
            Begin VB.TextBox txtEntregaBairro 
               Height          =   285
               Left            =   4620
               MaxLength       =   60
               TabIndex        =   66
               Text            =   "Text1"
               Top             =   1020
               Width           =   2955
            End
            Begin VB.TextBox txtEntregaCpl 
               Height          =   285
               Left            =   1020
               MaxLength       =   60
               TabIndex        =   65
               Text            =   "Text1"
               Top             =   1020
               Width           =   2715
            End
            Begin VB.TextBox txtEntregaNro 
               Height          =   285
               Left            =   7920
               MaxLength       =   60
               TabIndex        =   64
               Text            =   "Text1"
               Top             =   660
               Width           =   975
            End
            Begin VB.TextBox txtEntregaLgr 
               Height          =   285
               Left            =   1020
               MaxLength       =   60
               TabIndex        =   63
               Text            =   "Text1"
               Top             =   660
               Width           =   5955
            End
            Begin VB.CheckBox chkEntrega 
               Caption         =   "Local de ENTREGA diferente do emitente."
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
               TabIndex        =   62
               Top             =   0
               Width           =   3975
            End
            Begin VB.TextBox txtEntregaDoc 
               Height          =   285
               Left            =   1020
               MaxLength       =   14
               TabIndex        =   61
               Text            =   "Text1"
               Top             =   300
               Width           =   3435
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "CEP:"
               Height          =   195
               Left            =   6840
               TabIndex        =   77
               Top             =   1440
               Width           =   495
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "UF:"
               Height          =   195
               Left            =   480
               TabIndex        =   76
               Top             =   1440
               Width           =   495
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "Bairro:"
               Height          =   195
               Left            =   4080
               TabIndex        =   75
               Top             =   1080
               Width           =   495
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Compl.:"
               Height          =   195
               Left            =   420
               TabIndex        =   74
               Top             =   1080
               Width           =   555
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "Número:"
               Height          =   255
               Left            =   7020
               TabIndex        =   73
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "Endereço:"
               Height          =   255
               Left            =   240
               TabIndex        =   72
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Municipio:"
               Height          =   255
               Left            =   2040
               TabIndex        =   71
               Top             =   1440
               Width           =   795
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               Caption         =   "CNPJ/CPF:"
               Height          =   195
               Left            =   120
               TabIndex        =   70
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Transporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -74880
            TabIndex        =   57
            Top             =   2400
            Width           =   9075
            Begin VB.ComboBox cboTransportadora 
               Height          =   315
               Left            =   1440
               TabIndex        =   58
               Text            =   "cboTransportadora"
               Top             =   240
               Width           =   7335
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               Caption         =   "Transportadora:"
               Height          =   195
               Left            =   120
               TabIndex        =   59
               Top             =   300
               Width           =   1155
            End
         End
         Begin VB.Frame Frame11 
            Height          =   1995
            Left            =   -74820
            TabIndex        =   46
            Top             =   360
            Width           =   9075
            Begin VB.Frame Frame14 
               Caption         =   "Obs para o boleto de cobrança "
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
               Left            =   5640
               TabIndex        =   101
               Top             =   240
               Width           =   3315
               Begin VB.TextBox txtObsBoleto 
                  Height          =   1095
                  Left            =   180
                  MultiLine       =   -1  'True
                  TabIndex        =   102
                  Text            =   "formClientes.frx":7DEB
                  Top             =   300
                  Width           =   2955
               End
            End
            Begin VB.ComboBox cboCentroCustos 
               Height          =   315
               Left            =   2220
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   180
               Width           =   3315
            End
            Begin VB.ComboBox cboCondicoesPagamento 
               Height          =   315
               Left            =   2220
               Style           =   2  'Dropdown List
               TabIndex        =   50
               Top             =   840
               Width           =   3315
            End
            Begin VB.ComboBox cboTipoDocumento 
               Height          =   315
               Left            =   2220
               Style           =   2  'Dropdown List
               TabIndex        =   49
               Top             =   1200
               Width           =   3315
            End
            Begin VB.TextBox txtLimiteCredito 
               Height          =   285
               Left            =   2220
               TabIndex        =   48
               Text            =   "Text12"
               Top             =   1560
               Width           =   3315
            End
            Begin VB.ComboBox cboPlanoContas 
               Height          =   315
               Left            =   2220
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   540
               Width           =   3315
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               Caption         =   "Centro de Custos:"
               Height          =   255
               Left            =   60
               TabIndex        =   56
               Top             =   300
               Width           =   2115
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               Caption         =   "Condições de Pagamento:"
               Height          =   255
               Left            =   300
               TabIndex        =   55
               Top             =   960
               Width           =   1875
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "Forma de Pagamento:"
               Height          =   195
               Left            =   600
               TabIndex        =   54
               Top             =   1260
               Width           =   1575
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               Caption         =   "Limite Credito:"
               Height          =   195
               Left            =   1020
               TabIndex        =   53
               Top             =   1620
               Width           =   1155
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               Caption         =   "Plano de Contas:"
               Height          =   195
               Left            =   300
               TabIndex        =   52
               Top             =   600
               Width           =   1815
            End
         End
         Begin VB.Frame Frame12 
            Height          =   1935
            Left            =   -74700
            TabIndex        =   37
            Top             =   540
            Width           =   7335
            Begin VB.TextBox txtWebSite 
               Height          =   285
               Left            =   1200
               TabIndex        =   41
               Text            =   "Text1"
               Top             =   300
               Width           =   4575
            End
            Begin VB.TextBox txtEmailCom 
               Height          =   285
               Left            =   1200
               TabIndex        =   40
               Text            =   "Text1"
               Top             =   660
               Width           =   4575
            End
            Begin VB.TextBox txtEmailFin 
               Height          =   285
               Left            =   1200
               TabIndex        =   39
               Text            =   "Text1"
               Top             =   1020
               Width           =   4575
            End
            Begin VB.TextBox txtEmailNFe 
               Height          =   285
               Left            =   1200
               TabIndex        =   38
               Text            =   "Text1"
               Top             =   1380
               Width           =   4575
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               Caption         =   "Web site:"
               Height          =   195
               Left            =   300
               TabIndex        =   45
               Top             =   300
               Width           =   855
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               Caption         =   "e-mail Com.:"
               Height          =   255
               Left            =   300
               TabIndex        =   44
               Top             =   660
               Width           =   855
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               Caption         =   "e-mail Financ.:"
               Height          =   195
               Left            =   120
               TabIndex        =   43
               Top             =   1080
               Width           =   1035
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               Caption         =   "e-mail NF-e:"
               Height          =   195
               Left            =   240
               TabIndex        =   42
               Top             =   1440
               Width           =   915
            End
         End
         Begin VB.Frame Frame7 
            Height          =   1995
            Left            =   3840
            TabIndex        =   26
            Top             =   1320
            Width           =   4095
            Begin VB.TextBox txtCNAE 
               Height          =   285
               Left            =   1560
               TabIndex        =   31
               Text            =   "Text1"
               Top             =   1260
               Width           =   2235
            End
            Begin VB.TextBox txtIM 
               Height          =   285
               Left            =   1560
               TabIndex        =   30
               Text            =   "Text1"
               Top             =   900
               Width           =   2235
            End
            Begin VB.TextBox txtIEST 
               Height          =   285
               Left            =   1560
               TabIndex        =   29
               Text            =   "Text1"
               Top             =   540
               Width           =   2235
            End
            Begin VB.TextBox txtIE 
               Height          =   285
               Left            =   1560
               TabIndex        =   28
               Text            =   "Text1"
               Top             =   180
               Width           =   2235
            End
            Begin VB.TextBox txtSUFRAMA 
               Height          =   285
               Left            =   1560
               MaxLength       =   15
               TabIndex        =   27
               Text            =   "Text1"
               Top             =   1620
               Width           =   2235
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               Caption         =   "CNAE:"
               Height          =   195
               Left            =   540
               TabIndex        =   36
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               Caption         =   "Insc. Municipal:"
               Height          =   195
               Left            =   300
               TabIndex        =   35
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               Caption         =   "Insc. Estadual ST:"
               Height          =   195
               Left            =   180
               TabIndex        =   34
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               Caption         =   "Insc. Estadual:"
               Height          =   195
               Left            =   60
               TabIndex        =   33
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               Caption         =   "Insc. SUFRAMA:"
               Height          =   195
               Left            =   180
               TabIndex        =   32
               Top             =   1680
               Width           =   1275
            End
         End
         Begin VB.Frame Frame6 
            Height          =   795
            Left            =   60
            TabIndex        =   23
            Top             =   1380
            Width           =   2835
            Begin VB.CheckBox chkSimplesNacional 
               Caption         =   "Optante do Simples Nacional"
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   240
               Width           =   2535
            End
            Begin VB.CheckBox chkME 
               Caption         =   "Micro Empresa (M.E.)"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   480
               Width           =   2595
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Observações:"
            Height          =   915
            Left            =   60
            TabIndex        =   21
            Top             =   360
            Width           =   9015
            Begin VB.TextBox txtObs 
               Height          =   495
               Left            =   180
               TabIndex        =   22
               Text            =   "Text11"
               Top             =   240
               Width           =   8715
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Vendedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   120
            TabIndex        =   19
            Top             =   2400
            Width           =   3555
            Begin VB.ComboBox cboVendedor 
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   300
               Width           =   3135
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Obs. para a NFe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   -74820
            TabIndex        =   17
            Top             =   2400
            Width           =   9075
            Begin VB.TextBox txtObsNFe 
               Height          =   735
               Left            =   120
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               Text            =   "formClientes.frx":7DF1
               Top             =   240
               Width           =   8835
            End
         End
      End
   End
End
Attribute VB_Name = "formClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg As Integer
Dim strTabela As String



Private Sub PesquisarRegistro()
    IdReg = formBuscar.IniciarBusca(strTabela) ', "xNome,xlgr,nro,xcpl,xbairro,xmun,uf,fone")
    
    If IdReg = 0 Then
            LimpaFormulario Me
        Else
            MostrarDados
    End If
End Sub


Private Sub btFinanceiro_Click()
    Dim lSQL    As String
    Dim lData   As String
    Dim lNome   As String
    
    If IdReg = 0 Then
        MsgBox "Selecione um cliente.", vbInformation, App.EXEName
        Exit Sub
    End If
    
    lData = txtAno.Text & "-01-01"
    
    If Trim(txtxNome.Text) <> "" Then
        lNome = Trim(txtxNome.Text)
    End If

    lSQL = "SELECT * FROM FinanceiroContasPRCadastro " & _
           "WHERE ID_Empresa = " & ID_Empresa & " " & _
           "AND Emissao >= '" & lData & "' AND Emissao <= '" & Format(Date, "YYYY-MM-DD") & "' " & _
           "AND  Nome IN ('" & lNome & "')"
    
    formFinanceiroContasPRGerenciador.filtroExterno lSQL
    
    
End Sub

Private Sub btoAtualizarHist_Click()
    ListHistoricoNotasFiscais
    
        
End Sub

Private Sub cboCentroCustos_DropDown()
    Dim Rst As Recordset
    Dim sSQL As String
    sSQL = "SELECT * FROM FinanceiroCentroCustos WHERE ID_Empresa = " & ID_Empresa
    Set Rst = RegistroBuscar(sSQL)
    cboCentroCustos.Clear
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma CONTA cadastrada"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCentroCustos.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("Id"))) & Rst.Fields("Id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Clone
End Sub


Private Sub cboCondicoesPagamento_DropDown()
    Dim Rst As Recordset
    cboCondicoesPagamento.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCondicoespagamento WHERE ID_Empresa = " & ID_Empresa)
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

Private Sub cboEntregaMun_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    If Trim(cboEntregaUF.Text) = "" Then
        MsgBox "Selecione uma Unidade Federal (UF)."
        Exit Sub
    End If
    'sSQL = "SELECT * FROM TributacaoMunicipio WHERE ID_Empresa = " & ID_Empresa & " AND codUF = " & pgDadosICMS(cboEntregaUF.Text, 0).Id & " ORDER BY Descricao"
    sSQL = "SELECT * FROM TributacaoMunicipio WHERE UF = '" & Trim(UCase(cboEntregaUF.Text)) & "' ORDER BY Descricao"
    cboEntregaMun.Clear
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboEntregaMun.AddItem Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If

End Sub


Private Sub cboEntregaUF_DropDown()
    Dim Rst As Recordset
    cboEntregaUF.Clear
    Set Rst = RegistroBuscar("SELECT * FROM TributacaoUF ORDER BY Sigla")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboEntregaUF.AddItem Rst.Fields("Sigla")
                Rst.MoveNext
            Loop
    End If

End Sub



'Private Sub cboMun_DropDown()
'    Dim Rst     As Recordset
'    Dim sSQL    As String
 '   If Trim(cboUF.Text) = "" Then
'        MsgBox "Selecione uma Unidade Federal (UF)."
'        Exit Sub
'    End If
'    sSQL = "SELECT * FROM TributacaoMunicipio WHERE codUF = " & PgDadosUF(cboUF.Text).Id & " ORDER BY Descricao"
'    cboMun.Clear
'    Set Rst = RegistroBuscar(sSQL)
'    If Rst.BOF And Rst.EOF Then
'        Else
'            Rst.MoveFirst
'            Do Until Rst.EOF
'                cboMun.AddItem Rst.Fields("Descricao")
'                Rst.MoveNext
'            Loop
'    End If
'End Sub

'Private Sub cboPais_DropDown()
'    Dim Rst As Recordset
'    cboPais.Clear
'    Set Rst = RegistroBuscar("SELECT * FROM TributacaoPais ORDER BY Pais")
'    If Rst.BOF And Rst.EOF Then
'        Else
'            Rst.MoveFirst
'            Do Until Rst.EOF
'                cboPais.AddItem Rst.Fields("Pais")
'                Rst.MoveNext
'            Loop
'    End If
'End Sub

Private Sub cboPessoa_DropDown()
    cboPessoa.Clear
    cboPessoa.AddItem "Fisica"
    cboPessoa.AddItem "Juridica"
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

Private Sub cboStatus_DropDown()
    cboStatus.Clear
    cboStatus.AddItem "Ativo"
    cboStatus.AddItem "Inativo"
End Sub

Private Sub cboTipoDocumento_DropDown()
    Dim Rst As Recordset
    cboTipoDocumento.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento WHERE ID_Empresa = " & ID_Empresa)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboTipoDocumento.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                          Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub cbotransportadora_DropDown()
    Dim Rst As Recordset
    
    Set Rst = RegistroBuscar("SELECT * FROM Transportadoras WHERE ID_Empresa = " & ID_Empresa & " AND xNome LIKE '" & cboTransportadora.Text & "%'")
    cboTransportadora.Clear
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboTransportadora.AddItem Left("00000", 5 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub cboUF_DropDown()
    Dim Rst As Recordset
    cboUF.Clear
    Set Rst = RegistroBuscar("SELECT * FROM TributacaoUF ORDER BY sigla")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboUF.AddItem Rst.Fields("sigla")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub cboVendedor_DropDown()
    Dim Rst As Recordset
    cboVendedor.Clear
    Set Rst = RegistroBuscar("SELECT * FROM RHFuncionarioCadastro WHERE ID_Empresa = " & ID_Empresa & " ORDER BY xNome")
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboVendedor.AddItem Left(String(4, "0"), 4 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & " - " & Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If
End Sub


Private Sub cboxMun_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    If Trim(cboUF.Text) = "" Then
        MsgBox "Selecione uma Unidade Federal (UF)."
        Exit Sub
    End If
    sSQL = "SELECT * FROM TributacaoMunicipio WHERE UF = '" & Trim(UCase(cboUF.Text)) & "' ORDER BY Descricao"
    cboxMun.Clear
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboxMun.AddItem UCase(Rst.Fields("Descricao"))
                Rst.MoveNext
            Loop
    End If
End Sub


Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    
    Me.Top = 0
    Me.Left = 0
    LimpaFormulario Me
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    SSt.Tab = 0
    HDFormulario False
    HDMenu Me, True
    SSTab.Tab = 0
    IdReg = 0
    
    
    
End Sub



Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    
    IdReg = 0
    LimpaFormulario Me
    HDMenu Me, False
    HDFormulario True
    txtDoc.Enabled = True
    
    If Trim(PgDadosEmpresa(ID_Empresa).uf) <> "" Then
        cboUF.Clear
        cboUF.AddItem PgDadosEmpresa(ID_Empresa).uf
        cboUF.Text = cboUF.List(0)
    End If
    
    If Trim(PgDadosEmpresa(ID_Empresa).Mun) <> "" Then
        cboxMun.Clear
        cboxMun.AddItem PgDadosEmpresa(ID_Empresa).Mun
        cboxMun.Text = cboxMun.List(0)
    End If
    
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Selecione um Registro."
        Exit Sub
    End If
    HDFormulario True
    HDMenu Me, False
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
                        "Nome: " & txtxNome.Text, vbYesNo + vbQuestion) = vbYes Then
                               
                If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                    LimpaFormulario Me
                    End If
                End If
    End If
End Sub
Private Sub Salvar()
    If grvRegistro = True Then
        HDMenu Me, True
        HDFormulario False
        
        
    End If
End Sub
Private Sub Cancelar()
    HDMenu Me, True
    HDFormulario False
    LimpaFormulario Me
    
    
End Sub

Private Sub HDFormulario(op As Boolean)
    HDForm Me, op
    msfgHist.Enabled = IIf(op = True, False, True)
    txtDoc.Enabled = IIf(op = True, False, True)
    txtAno.Enabled = IIf(op = True, False, True)
    btoAtualizarHist.Enabled = IIf(op = True, False, True)
    btFinanceiro.Enabled = IIf(op = True, False, True)
    
End Sub


Private Sub msfgHist_DblClick()
    If msfgHist.TextMatrix(msfgHist.Row, 0) = "ID" Or msfgHist.Rows = 1 Then Exit Sub
    ImprimirDANFE2 (msfgHist.TextMatrix(msfgHist.Row, 4))
    'lnPv = msfgItens.Row
    'IdItem = IIf(Trim(msfgItens.TextMatrix(msfgItens.Row, 0)) = "", 0, msfgItens.TextMatrix(msfgItens.Row, 0))
    
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
            ImprimirListaClientes
        Case "Pesquisar"
            PesquisarRegistro
        Case "Salvar"
            Salvar
        Case "Cancelar"
            Cancelar
        Case "Manutenção da Tabela"
            formManutencaoTabelas.IniciarManutencao Me, "SELECT * FROM Clientes"
    End Select
End Sub
Private Function ValidarDados() As Boolean
    If Trim(txtDoc.Text) <> "99999999999999" Then
            If Validar_CNPJ_CPF(Trim(txtDoc.Text)) = False Then
                MsgBox "O campo CNPJ/CPF esta com valor invalido. Favor verificar!"
                ValidarDados = False
                Exit Function
            End If
            'If Validar_IE(Trim(txtIE.Text), cboUF.Text) = False Then
            '    MsgBox "O campo Inscrição Estadual esta com valor invalido. Favor verificar!"
            '    ValidarDados = False
            '    Exit Function
            'End If
        Else
            '26/10/2012 - Checa se CONSUMIDOR ja não esta cadastrado
            If BuscarDados(Trim(txtDoc.Text)) = True Then
                MsgBox "Cliente ja cadastrado!", vbInformation, App.EXEName
                ValidarDados = False
                Exit Function
            End If
    End If
    If cboUF.Text = "" Then
        MsgBox "O campo UF invalido. Favor verificar!"
        ValidarDados = False
        Exit Function
    End If
    If cboPessoa.Text = "" Then
        MsgBox "O campo PESSOA esta com valor invalido. Favor verificar!"
        ValidarDados = False
        Exit Function
    End If
    If Trim(txtxLgr.Text) <> "" And Trim(txtNro.Text) = "" Then
        MsgBox "O campo NUMERO esta com valor invalido. Favor verificar!"
        ValidarDados = False
        Exit Function
    End If
    If Trim(txtxNome.Text) = "" Then
        MsgBox "O campo NOME/RAZÃO SOCIAL está com valor invalido. Favor verificar!"
        ValidarDados = False
        Exit Function
    End If
    
    
    ValidarDados = True

End Function
Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim i           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer
    
    If ValidarDados = False Then
        grvRegistro = False
        Exit Function
    End If
    
    cReg = 0
    
    vReg(cReg) = Array("Doc", txtDoc.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("xNome", txtxNome.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Fant", txtFant.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Pessoa", cboPessoa.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Status", cboStatus.Text, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("xLgr", txtxLgr.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Nro", txtNro.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("xCpl", txtxCpl.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("xBairro", txtxBairro.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("UF", cboUF.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("xMun", cboxMun.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("CEP", txtCEP.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("eMail", txteMail.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Fone", txtFone.Text, "S"): cReg = cReg + 1
    'Entrega
    vReg(cReg) = Array("Entrega", chkEntrega.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("EntregaDoc", txtEntregaDoc.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("EntregaLgr", txtEntregaLgr.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("EntregaNro", txtEntregaNro.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("EntregaCpl", txtEntregaCpl.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("EntregaBairro", txtEntregaBairro.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("EntregaUF", cboEntregaUF.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("EntregaMun", cboEntregaMun.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("EntregaCEP", txtEntregaCEP.Text, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("Transportadora", Left(cboTransportadora.Text, 5), "N"): cReg = cReg + 1
    'Cobranca
    vReg(cReg) = Array("CentroCustos", IIf(Trim(cboCentroCustos.Text) = "", 0, Left(cboCentroCustos.Text, 3)), "N"): cReg = cReg + 1
    vReg(cReg) = Array("CondicoesPagamento", IIf(Trim(cboCondicoesPagamento.Text) = "", 0, Left(cboCondicoesPagamento.Text, 3)), "N"): cReg = cReg + 1
    vReg(cReg) = Array("TipoDocumento", IIf(Trim(cboTipoDocumento.Text) = "", 0, Left(cboTipoDocumento.Text, 3)), "N"): cReg = cReg + 1
    vReg(cReg) = Array("PlanoContas", IIf(Trim(cboPlanoContas.Text) = "", 0, Left(cboPlanoContas.Text, 3)), "N"): cReg = cReg + 1
    vReg(cReg) = Array("LimiteCredito", txtLimiteCredito.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ObsNFe", txtObsNFe.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ObsBoleto", txtObsBoleto.Text, "S"): cReg = cReg + 1
    
    'Contatos
    vReg(cReg) = Array("WebSite", txtWebSite.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("EmailCom", txtEmailCom.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("EmailFin", txtEmailFin.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("EmailNFe", txtEmailNFe.Text, "S"): cReg = cReg + 1
    'Outros
    vReg(cReg) = Array("Obs", txtObs.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("SimplesNacional", chkSimplesNacional.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("ME", chkME.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("Vendedor", IIf(Trim(cboVendedor.Text) = "", 0, Left(cboVendedor.Text, 4)), "N"): cReg = cReg + 1
    vReg(cReg) = Array("IE", txtIE.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("IEST", txtIEST.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("IM", txtIM.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("CNAE", txtCNAE.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("SUFRAMA", txtSuframa.Text, "S") ': cReg = cReg + 1
    
    
    
    
    
    
    If IdReg = 0 Then
            If RegistroIncluir(strTabela, vReg, cReg) = 0 Then
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


End Function





Private Sub txtAno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        ListHistoricoNotasFiscais
    End If
    KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub txtDoc_Change()
    If Trim(txtDoc.Text) = String(14, "9") And txtxLgr.Enabled = True Then
        txtxNome.Text = "*** CONSUMIDOR ***"
        HDForm Me, False
    End If
End Sub

Private Sub txtDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
    
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        If BuscarDados(txtDoc.Text) = True Then
            MostrarDados
        End If
    End If
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub
Private Function BuscarDados(strCNPJ As String) As Boolean
    Dim Rst     As ADODB.Recordset
    Dim strSQL  As String
    
    If Trim(strCNPJ) = "" Then Exit Function
    
    SSt.Tab = 0
    
    strSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Doc = '" & strCNPJ & "'"

    Set Rst = RegistroBuscar(strSQL)
    If Rst.BOF And Rst.EOF Then
            BuscarDados = False
            MsgBox "Nenhum Registro encontrado!", vbExclamation, App.EXEName
            Rst.Close
            Exit Function
        Else
            BuscarDados = True
            Rst.MoveFirst
            IdReg = Rst.Fields("Id")
            Rst.Close
    End If
    
End Function
Private Sub MostrarDados()
    Dim sSQL    As String
    Dim tmp     As String
    
    ListHistoricoNotasFiscais
   
    SSTab.Tab = 0
            
    sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & IdReg
    ExibirDados Me, sSQL
    
     With cboTransportadora
        If Trim(.Text) <> "0" And Trim(.Text) <> "" Then
                tmp = Trim(.Text)
                .Clear
                .AddItem Left("00000", 5 - Len(tmp)) & tmp & " - " & pgDadosTransportadora(CInt(tmp)).Nome
                .Text = .List(0)
            Else
                .Clear
        End If
    End With
 
    
    If Trim(cboVendedor.Text) <> "" And Trim(cboVendedor.Text) <> "0" Then
        tmp = cboVendedor.Text
        cboVendedor.Clear
        cboVendedor.AddItem Left("0000", 4 - Len(tmp)) & tmp & " - " & PgDadosRhFuncionario(CInt(tmp)).Nome
        cboVendedor.Text = cboVendedor.List(0)
    End If
    
    With cboCentroCustos
        If Trim(.Text) <> "0" And Trim(.Text) <> "" Then
                tmp = Trim(.Text)
                .Clear
                .AddItem Left("000", 3 - Len(tmp)) & tmp & " - " & pgDadosCentroCustos(CInt(tmp)).Descricao
                .Text = .List(0)
            Else
                .Clear
        End If
    End With
    With cboCondicoesPagamento
        If Trim(.Text) <> "0" And Trim(.Text) <> "" Then
                tmp = Trim(.Text)
                .Clear
                .AddItem Left("000", 3 - Len(tmp)) & tmp & " - " & pgDescrCondPag(tmp)
                .Text = .List(0)
            Else
                .Clear
        End If
    End With
    With cboTipoDocumento
        If Trim(.Text) <> "0" And Trim(.Text) <> "" Then
                tmp = Trim(.Text)
                .Clear
                .AddItem Left("000", 3 - Len(tmp)) & tmp & " - " & pgDadosTipoDocumento(CInt(tmp)).Descricao
                .Text = .List(0)
            Else
                .Clear
        End If
    End With
    With cboPlanoContas
        If Trim(.Text) <> "0" And Trim(.Text) <> "" Then
                tmp = Trim(.Text)
                .Clear
                .AddItem ZE(CInt(tmp), 3) & " - (" & PgDadosPlanoContas("ID", CInt(tmp)).Codigo & ") " & PgDadosPlanoContas("ID", CInt(tmp)).Descricao
                .Text = .List(0)
            Else
                .Clear
        End If
    End With
End Sub
Private Sub txtFone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtIE_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtLimiteCredito_GotFocus()
    txtLimiteCredito.Text = ChkVal(txtLimiteCredito.Text, 0, 2)
End Sub

Private Sub txtLimiteCredito_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtLimiteCredito.Text, KeyAscii, cDecMoeda)
End Sub
Private Sub txtLimiteCredito_LostFocus()
    txtLimiteCredito.Text = ConvMoeda(txtLimiteCredito.Text)
End Sub
Private Sub ListHistoricoNotasFiscais()
    Dim sSQL    As String
    Dim Rst     As Recordset
    msfgHist.Rows = 1
    If IdReg = 0 Then
        Exit Sub
    End If
    'If Trim(txtAno.Text) = "" Then Exit Sub
    

    If Trim(txtAno.Text) = "" Then
        txtAno.Text = Format(Date, "YYYY")
    End If
    If Val(txtAno.Text) < 2000 Then
        MsgBox "Ano (" & txtAno.Text & ") invalido!", vbInformation, App.EXEName
        txtAno.Text = Format(Date, "YYYY")
        Exit Sub
    End If

    sSQL = "SELECT * FROM FaturamentoNFe " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND dest_IdDest = " & IdReg & _
           " AND movfisco=1 " & _
           " AND YEAR(ide_dEmi)>='" & txtAno.Text & "'"
           '" ORDER BY ide_nNF LIMIT 200"
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
        txtAno.Enabled = True
        btoAtualizarHist.Enabled = True
            Rst.MoveFirst
            With msfgHist
                Do Until Rst.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("ID")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("ide_dEmi")
                    .TextMatrix(.Rows - 1, 2) = Rst.Fields("ide_Serie")
                    .TextMatrix(.Rows - 1, 3) = Rst.Fields("ide_nnf")
                    .TextMatrix(.Rows - 1, 4) = Rst.Fields("idNFe")
                    .TextMatrix(.Rows - 1, 5) = ConvMoeda(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                    .TextMatrix(.Rows - 1, 6) = IIf(IsNull(Rst.Fields("canc_nProt")), IIf(IsNull(Rst.Fields("StatusNFe")), "", Rst.Fields("StatusNFe")), Rst.Fields("canc_Status"))
                    
                    Rst.MoveNext
                Loop
            End With
    End If
End Sub

