VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formEstoqueProduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estoque - Cadastro de Produto"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   10920
   Begin VB.Frame Frame9 
      Caption         =   "Gerenciador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   47
      Top             =   3120
      Width           =   8055
      Begin VB.CheckBox chkIncluirBalanco 
         Caption         =   "Incluir no balanço"
         Height          =   255
         Left            =   180
         TabIndex        =   48
         Top             =   240
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab sstProdutos 
      Height          =   4335
      Left            =   60
      TabIndex        =   26
      Top             =   3780
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Especificações"
      TabPicture(0)   =   "formEstoqueProduto.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Fiscal / Tributario"
      TabPicture(1)   =   "formEstoqueProduto.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Preço / Movimentações"
      TabPicture(2)   =   "formEstoqueProduto.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Documentos Anexados"
      TabPicture(3)   =   "formEstoqueProduto.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame13"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Kit"
      TabPicture(4)   =   "formEstoqueProduto.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame15"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame15 
         Height          =   3675
         Left            =   -74880
         TabIndex        =   81
         Top             =   480
         Width           =   10395
         Begin VB.Frame Frame16 
            Height          =   1275
            Left            =   120
            TabIndex        =   82
            Top             =   2280
            Width           =   10155
            Begin VB.CommandButton btPesqKitItem 
               Height          =   315
               Left            =   2340
               Picture         =   "formEstoqueProduto.frx":008C
               Style           =   1  'Graphical
               TabIndex        =   91
               ToolTipText     =   "Buscar local do Arquivo"
               Top             =   240
               Width           =   435
            End
            Begin VB.TextBox txtKitId 
               Height          =   315
               Left            =   1080
               TabIndex        =   90
               Text            =   "Text2"
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtKitQtd 
               Height          =   285
               Left            =   1080
               MaxLength       =   250
               TabIndex        =   85
               Text            =   "Text1"
               Top             =   600
               Width           =   1215
            End
            Begin VB.CommandButton btKitAdd 
               Height          =   375
               Left            =   9360
               Picture         =   "formEstoqueProduto.frx":0416
               Style           =   1  'Graphical
               TabIndex        =   84
               Top             =   300
               Width           =   675
            End
            Begin VB.CommandButton btKitDel 
               Height          =   375
               Left            =   9360
               Picture         =   "formEstoqueProduto.frx":07A0
               Style           =   1  'Graphical
               TabIndex        =   83
               Top             =   720
               Width           =   675
            End
            Begin VB.Label lblKitDescr 
               Caption         =   "Label27"
               Height          =   255
               Left            =   2880
               TabIndex        =   92
               Top             =   300
               Width           =   6375
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               Caption         =   "Descrição:"
               Height          =   195
               Left            =   180
               TabIndex        =   87
               Top             =   660
               Width           =   795
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               Caption         =   "Produto ID:"
               Height          =   195
               Left            =   60
               TabIndex        =   86
               Top             =   300
               Width           =   915
            End
         End
         Begin MSFlexGridLib.MSFlexGrid msfgKit 
            Height          =   1935
            Left            =   120
            TabIndex        =   88
            Top             =   180
            Width           =   10155
            _ExtentX        =   17912
            _ExtentY        =   3413
            _Version        =   393216
            Cols            =   4
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"formEstoqueProduto.frx":0B2A
         End
         Begin VB.Label Label26 
            Caption         =   "Duplo click para editar o item..."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   180
            TabIndex        =   89
            Top             =   2100
            Width           =   10035
         End
      End
      Begin VB.Frame Frame13 
         Height          =   3675
         Left            =   -74880
         TabIndex        =   67
         Top             =   480
         Width           =   10395
         Begin VB.Frame Frame14 
            Height          =   1275
            Left            =   120
            TabIndex        =   69
            Top             =   2280
            Width           =   10155
            Begin VB.CommandButton btoFileExcluir 
               Height          =   375
               Left            =   9360
               Picture         =   "formEstoqueProduto.frx":0BCF
               Style           =   1  'Graphical
               TabIndex        =   76
               Top             =   720
               Width           =   675
            End
            Begin VB.CommandButton btoFileIncluir 
               Height          =   375
               Left            =   9360
               Picture         =   "formEstoqueProduto.frx":0F59
               Style           =   1  'Graphical
               TabIndex        =   75
               Top             =   300
               Width           =   675
            End
            Begin MSComDlg.CommonDialog cdFile 
               Left            =   8700
               Top             =   180
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton btoFileBuscar 
               Height          =   315
               Left            =   8100
               Picture         =   "formEstoqueProduto.frx":12E3
               Style           =   1  'Graphical
               TabIndex        =   74
               ToolTipText     =   "Buscar local do Arquivo"
               Top             =   240
               Width           =   435
            End
            Begin VB.TextBox txtFileDescricao 
               Height          =   285
               Left            =   1020
               MaxLength       =   250
               TabIndex        =   73
               Text            =   "Text1"
               Top             =   600
               Width           =   6975
            End
            Begin VB.TextBox txtFile 
               Height          =   285
               Left            =   1020
               TabIndex        =   72
               Text            =   "Text1"
               Top             =   240
               Width           =   6975
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               Caption         =   "Arquivo:"
               Height          =   195
               Left            =   360
               TabIndex        =   71
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Descrição:"
               Height          =   195
               Left            =   180
               TabIndex        =   70
               Top             =   660
               Width           =   795
            End
         End
         Begin MSFlexGridLib.MSFlexGrid msfgDocAnexo 
            Height          =   1935
            Left            =   120
            TabIndex        =   68
            Top             =   180
            Width           =   10155
            _ExtentX        =   17912
            _ExtentY        =   3413
            _Version        =   393216
            Cols            =   3
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"formEstoqueProduto.frx":166D
         End
         Begin VB.Label Label21 
            Caption         =   "Duplo click para abrir o documento..."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   180
            TabIndex        =   77
            Top             =   2100
            Width           =   10035
         End
      End
      Begin VB.Frame Frame10 
         Height          =   3675
         Left            =   -74880
         TabIndex        =   49
         Top             =   480
         Width           =   10515
         Begin VB.Frame Frame12 
            Caption         =   "Dez ultimas SAIDAS"
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
            Left            =   3000
            TabIndex        =   65
            Top             =   1920
            Width           =   7395
            Begin MSFlexGridLib.MSFlexGrid msfgUltSaidas 
               Height          =   1215
               Left            =   180
               TabIndex        =   66
               Top             =   300
               Width           =   7035
               _ExtentX        =   12409
               _ExtentY        =   2143
               _Version        =   393216
               Cols            =   5
               SelectionMode   =   1
               AllowUserResizing=   1
               FormatString    =   "^NFe      |^Emissão      |<Fornecedor                                             |>Valor Unitario  |>Quantidade"
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Dez ultimas ENTRADAS"
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
            Left            =   3000
            TabIndex        =   63
            Top             =   240
            Width           =   7395
            Begin MSFlexGridLib.MSFlexGrid msfgUltEntradas 
               Height          =   1155
               Left            =   180
               TabIndex        =   64
               Top             =   300
               Width           =   7035
               _ExtentX        =   12409
               _ExtentY        =   2037
               _Version        =   393216
               Cols            =   5
               SelectionMode   =   1
               AllowUserResizing=   1
               FormatString    =   "^NFe      |^Emissão      |<Fornecedor                                             |>Valor Unitario  |>Quantidade"
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Composição de Preço"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   180
            TabIndex        =   50
            Top             =   240
            Width           =   2715
            Begin VB.TextBox txtCusto 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1020
               TabIndex        =   60
               Text            =   "Text1"
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox txtVlIPI 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1020
               TabIndex        =   59
               Text            =   "Text1"
               Top             =   780
               Width           =   1575
            End
            Begin VB.TextBox txtMarkup 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1020
               TabIndex        =   58
               Text            =   "Text1"
               Top             =   1500
               Width           =   1575
            End
            Begin VB.TextBox txtPreco 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1020
               TabIndex        =   57
               Text            =   "Text1"
               Top             =   2040
               Width           =   1575
            End
            Begin VB.TextBox txtOutros 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1020
               TabIndex        =   56
               Text            =   "Text1"
               Top             =   1140
               Width           =   1575
            End
            Begin VB.CheckBox chkCalcAutomCusto 
               Alignment       =   1  'Right Justify
               Caption         =   "Custo:"
               Height          =   195
               Left            =   60
               TabIndex        =   55
               Top             =   480
               Width           =   915
            End
            Begin VB.CheckBox chkCalcAutomIPI 
               Alignment       =   1  'Right Justify
               Caption         =   "IPI:"
               Height          =   195
               Left            =   60
               TabIndex        =   54
               Top             =   840
               Width           =   915
            End
            Begin VB.CheckBox chkCalcAutomOutros 
               Alignment       =   1  'Right Justify
               Caption         =   "Outros:"
               Height          =   255
               Left            =   60
               TabIndex        =   53
               Top             =   1140
               Width           =   915
            End
            Begin VB.CheckBox chkCalcAutomMarkup 
               Alignment       =   1  'Right Justify
               Caption         =   "Markup:"
               Height          =   195
               Left            =   60
               TabIndex        =   52
               Top             =   1560
               Width           =   915
            End
            Begin VB.CheckBox chkCalcAutomPreco 
               Alignment       =   1  'Right Justify
               Caption         =   "Tabela:"
               Height          =   315
               Left            =   60
               TabIndex        =   51
               Top             =   2040
               Width           =   915
            End
            Begin VB.Line Line1 
               BorderWidth     =   2
               X1              =   600
               X2              =   2595
               Y1              =   1920
               Y2              =   1920
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Fisco/Tributos:"
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
         TabIndex        =   29
         Top             =   480
         Width           =   10515
         Begin VB.Frame Frame7 
            Caption         =   "ICMS"
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
            TabIndex        =   39
            Top             =   960
            Width           =   7815
            Begin VB.ComboBox cboICMSCST 
               Height          =   315
               Left            =   720
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   720
               Width           =   6975
            End
            Begin VB.ComboBox cboICMSOrigem 
               Height          =   315
               Left            =   720
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   240
               Width           =   6975
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Caption         =   "CST:"
               Height          =   195
               Left            =   240
               TabIndex        =   43
               Top             =   780
               Width           =   375
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "Origem:"
               Height          =   195
               Left            =   120
               TabIndex        =   42
               Top             =   300
               Width           =   555
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "IPI"
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
            Left            =   120
            TabIndex        =   32
            Top             =   2280
            Width           =   7815
            Begin VB.TextBox txtIPICodEnquadramento 
               Height          =   285
               Left            =   6060
               MaxLength       =   3
               TabIndex        =   35
               Text            =   "999"
               Top             =   180
               Width           =   615
            End
            Begin VB.ComboBox cboIPICST 
               Height          =   315
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   540
               Width           =   6855
            End
            Begin VB.TextBox txtIPIAliquota 
               Height          =   285
               Left            =   840
               MaxLength       =   10
               TabIndex        =   33
               Text            =   "Text1"
               Top             =   180
               Width           =   675
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "Código de Enquadramento:"
               Height          =   195
               Left            =   4020
               TabIndex        =   38
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "CST:"
               Height          =   195
               Left            =   120
               TabIndex        =   37
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Aliq. IPI:"
               Height          =   195
               Left            =   120
               TabIndex        =   36
               Top             =   180
               Width           =   615
            End
         End
         Begin VB.TextBox txtMVA 
            Height          =   285
            Left            =   600
            MaxLength       =   5
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtNCM 
            Height          =   285
            Left            =   600
            MaxLength       =   8
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblDescrNCM 
            Height          =   675
            Left            =   2400
            TabIndex        =   46
            Top             =   180
            Width           =   7935
         End
         Begin VB.Label Label9 
            Caption         =   "MVA:"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   660
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "NCM:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Informações Complementares:"
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
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   10515
         Begin VB.TextBox txtInformacoesComplementares 
            Height          =   3195
            Left            =   120
            MaxLength       =   65000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Text            =   "formEstoqueProduto.frx":1724
            Top             =   300
            Width           =   10215
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quantidades:"
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
      Left            =   8160
      TabIndex        =   6
      Top             =   1980
      Width           =   2655
      Begin VB.TextBox txtSaldo 
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1380
         Width           =   1635
      End
      Begin VB.ComboBox cboUnidade 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1020
         Width           =   1635
      End
      Begin VB.TextBox txtQtdMinima 
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   660
         Width           =   1635
      End
      Begin VB.TextBox txtQtdMedia 
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo Atual:"
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Unidade:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Minima:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Média:"
         Height          =   255
         Left            =   420
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Classificação:"
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
      Left            =   60
      TabIndex        =   4
      Top             =   1980
      Width           =   8055
      Begin VB.ComboBox cboFabricante 
         Height          =   315
         Left            =   5340
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   240
         Width           =   2595
      End
      Begin VB.ComboBox cboSubGrupo 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   3255
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Fabricante:"
         Height          =   255
         Left            =   4500
         TabIndex        =   61
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Subgrupo:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   540
         TabIndex        =   5
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   10755
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   180
         Width           =   4335
      End
      Begin VB.CommandButton btoGerarReferencia 
         Caption         =   "..."
         Height          =   315
         Left            =   2820
         TabIndex        =   78
         ToolTipText     =   "Gerar referencia..."
         Top             =   600
         Width           =   315
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   960
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   195
         Width           =   1095
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   180
         Width           =   2115
      End
      Begin VB.TextBox txtCodigoBarras 
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   4020
         MaxLength       =   120
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   615
         Width           =   6615
      End
      Begin VB.TextBox txtReferencia 
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   615
         Width           =   1815
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Depósito:"
         Height          =   195
         Left            =   2340
         TabIndex        =   79
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "ID:"
         Height          =   195
         Left            =   720
         TabIndex        =   24
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Status:"
         Height          =   195
         Left            =   7620
         TabIndex        =   16
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Código de Barras:"
         Height          =   465
         Left            =   60
         TabIndex        =   9
         Top             =   930
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   3180
         TabIndex        =   3
         Top             =   690
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Referencia:"
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   660
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
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
            Object.ToolTipText     =   "Clonar Produto"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Kardex"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":172A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":1B7C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":1E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":2728
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":397A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":4254
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":4AE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":5378
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":65CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":68E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":6BFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":6FF5
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueProduto.frx":76EF
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formEstoqueProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IdReg       As Long
Dim strTabela   As String
Dim lKit        As Integer 'Linha do grid da tabela kit


Private Sub Cancelar()
    HDMenu Me, True
    HDFormLocal False
    IdReg = 0
    LpForm
    'txtID.Enabled = True
    'msfgDocAnexo.Enabled = True
End Sub

Private Sub HDFormLocal(op As Boolean)
    HDForm Me, op
    txtID.Enabled = IIf(op = True, False, True)
    msfgDocAnexo.Enabled = True 'IIf(op = True, False, True)
    msfgUltEntradas.Enabled = True
    msfgUltSaidas.Enabled = True
End Sub

Private Sub ListarItensDaGrade()
'    '**************************************************************************************************************************
'    '**** Alterar os dados do Kit
'    '**** Função criada em 11/12/2017
'    '**** Alterar os dados do Kit do produto
'    '**************************************************************************************************************************
'    If msfgKit.Rows = 1 Then Exit Function
'    With msfgKit
'        'Exclui registros anteriores
'        RegistroExcluir strTabela & "Kit", "idProduto = " & IdReg
'
'        For i = 1 To .Rows - 1
'           cReg = 0
'           vReg(cReg) = Array("IdProduto", IdReg, "S"): cReg = cReg + 1
'           vReg(cReg) = Array("IdItemKit", .TextMatrix(i, 1), "N"): cReg = cReg + 1
'           vReg(cReg) = Array("qtd", .TextMatrix(i, 3), "N"): cReg = cReg + 1
'
'           cReg = cReg - 1
'
'           RegistroIncluir strTabela & "Kit", vReg, cReg
'        Next
'    End With

    Dim Rst     As Recordset
    Dim sSQL    As String
    
    msfgKit.Rows = 1
    sSQL = "SELECT * " & _
           "FROM estoqueprodutokit " & _
           "WHERE IdProduto=" & IdReg
           
    Set Rst = RegistroBuscar(sSQL)
    'Incluso em 17.01.18
    If Rst Is Nothing Then Exit Sub
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                With msfgKit
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("Id")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("IdItemKit")
                    .TextMatrix(.Rows - 1, 2) = pgDadosEstoqueProduto(Rst.Fields("IdItemKit")).Descricao
                    .TextMatrix(.Rows - 1, 3) = Rst.Fields("qtd")
                    '.TextMatrix(.Rows - 1, 3) = ChkVal(Rst.Fields("det_vUnCom"), 0, cDecMoeda)
                    '.TextMatrix(.Rows - 1, 4) = ChkVal(Rst.Fields("det_qCom"), 0, cDecQtd) & "/" & Rst.Fields("det_uCom")
                    
                End With
                Rst.MoveNext
                
            Loop
    End If
    Rst.Close
End Sub

Private Sub LpForm()
    LimpaFormulario Me
    sstProdutos.Tab = 0
    txtIPICodEnquadramento.Text = "999"
    msfgDocAnexo.Rows = 1
    msfgKit.Rows = 1
    lblKitDescr.Caption = ""
End Sub

Private Sub PesquisarRegistro()
    LpForm 'me
    If IdReg = 0 Then
        IdReg = formBuscar.IniciarBusca(strTabela) ', "Descricao, Referencia, NCM,IPIAliquota,ICMSCST, Deposito, Status")
    End If
    
    If IdReg <> 0 Then
        MostrarDados
    End If
    ''txtID.Enabled = True
End Sub
Private Function ChkNCM(sNCM As String) As String
    'Checa se a NCM e valida e retorna sua descricao
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    If Trim(sNCM) = "" Then
        ChkNCM = ""
        Exit Function
    End If
    sSQL = "SELECT * FROM TributacaoNCM WHERE NCM ='" & sNCM & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            ChkNCM = "<<< Codigo NCM invalido. >>>"
        Else
            Rst.MoveFirst
            ChkNCM = Rst.Fields("Descricao") & _
                    " - IPI: " & IIf(IsNull(Rst.Fields("IPI")), "0%", Rst.Fields("IPI") & "%") & _
                    IIf(IsNull(Rst.Fields("CEST")), "", " - CEST: " & Rst.Fields("CEST"))
                    
    End If
    Rst.Close
    
End Function






Private Sub btKitAdd_Click()
    Dim sSQL As String
    If Trim(txtKitId.Text) = "" Or Trim(txtKitQtd.Text) = "" Then
        MsgBox "Não pode haver campos em branco. Por favor verifique!", vbInformation, App.EXEName
        Exit Sub
    End If
    
    '17.12.17 - Verifica se o item ja nao e um kit
    Dim Rst1 As Recordset
    sSQL = "SELECT * FROM EstoqueProdutoKit WHERE idProduto=" & Trim(txtKitId.Text)
    Set Rst1 = RegistroBuscar(sSQL)
    If Rst1.BOF And Rst1.EOF Then
            'MsgBox "Erro ao localizar Produto no Estoque.", vbInformation, "Aviso"
            'MovimentarEstoque = False
            'Exit Function
            Rst1.Close
        Else
            MsgBox "Este item não pode ser incluso pois já se trata de um KIT!", vbInformation, App.EXEName
            Rst1.Close
            Exit Sub
    End If
    If Trim(txtKitId.Text) = IdReg Then
        MsgBox "O item não pode ser incluso nele mesmo!", vbInformation, App.EXEName
        Exit Sub
    End If
    
    
    
    
    
    
    With msfgKit
        If lKit = 0 Then
            .Rows = .Rows + 1
            lKit = .Rows - 1
        End If
        '.TextMatrix(.Rows - 1, 0) = idDoc
        .TextMatrix(lKit, 1) = txtKitId.Text
        .TextMatrix(lKit, 2) = lblKitDescr.Caption
        .TextMatrix(lKit, 3) = txtKitQtd.Text
    End With
    txtKitId.Text = ""
    txtKitQtd.Text = ""
    lblKitDescr.Caption = ""
    lKit = 0
End Sub

Private Sub btKitDel_Click()
    If lKit = 0 Then
        MsgBox "Selecione um item da grade!", vbInformation, App.EXEName
        Exit Sub
    End If
    With msfgKit
        If .Rows <= 2 Then
                .Rows = 1
            Else
                .RemoveItem lKit
        End If
       
    End With
    txtKitId.Text = ""
    txtKitQtd.Text = ""
    lblKitDescr.Caption = ""
    lKit = 0

End Sub

Private Sub btoFileBuscar_Click()
    Dim sFile As String
    cdFile.ShowOpen
    sFile = cdFile.filename
    txtFile.Text = sFile
End Sub



Private Sub btoFileExcluir_Click()
    With msfgDocAnexo
        If .Rows <= 2 Then
                .Rows = 1
            Else
                .RemoveItem .Row
        End If
        
        
    End With
'    Dim idFile      As Integer
'    Dim sDescrFile  As String
'    With msfgDocAnexo
'        idFile = .TextMatrix(.Row, 0)
'        sDescrFile = .TextMatrix(.Row, 1)
'        If MsgBox("Deseja realmente EXCLUIR o arquivo abaixo?" & vbCrLf & vbCrLf & _
'                  "Arquivo: " & sDescrFile, vbInformation + vbYesNo, App.EXEName) = vbYes Then
'            If ExcluirFile(idFile) = True Then
'                RegistroExcluir "EstoqueProdutoArquivos", "id=" & idFile
'                ListarArquivos
'            End If
'        End If
'    End With
End Sub

Private Sub btoFileIncluir_Click()
    If Trim(txtFile.Text) = "" Or Trim(txtFileDescricao.Text) = "" Then
        MsgBox "Não pode haver campos em branco. Por favor verifique!", vbInformation, App.EXEName
        Exit Sub
    End If
    With msfgDocAnexo
        .Rows = .Rows + 1
        '.TextMatrix(.Rows - 1, 0) = idDoc
        .TextMatrix(.Rows - 1, 1) = txtFileDescricao.Text
        .TextMatrix(.Rows - 1, 2) = txtFile.Text
    End With
    txtFile.Text = ""
    txtFileDescricao.Text = ""
End Sub

Private Sub btoGerarReferencia_Click()
    txtReferencia.Text = MontarReferenciaProduto(IdReg)
End Sub

Private Sub btPesqKitItem_Click()

    pesqItemKit
End Sub
Private Sub pesqItemKit(Optional Id As Integer)
    Dim prodIdKit As Long
    If Id <> 0 Then
        prodIdKit = Id
    End If
    
    If prodIdKit = 0 Then
        prodIdKit = formBuscar.IniciarBusca(strTabela) ', "Descricao, Referencia, NCM,IPIAliquota,ICMSCST, Deposito, Status")
    End If
    
    If prodIdKit = 0 Then
        Exit Sub
    End If
    txtKitId.Text = pgDadosEstoqueProduto(prodIdKit).Id
    lblKitDescr.Caption = pgDadosEstoqueProduto(prodIdKit).Descricao
End Sub
Private Sub cboDeposito_DropDown()
    Dim Rst As Recordset
    cboDeposito.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueDeposito WHERE ID_Empresa=" & ID_Empresa & " ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
            
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboDeposito.AddItem Left(String(5, "0"), 5 - Len(Trim(Rst.Fields("id")))) & Trim(Rst.Fields("id")) & " - " & Rst.Fields("descricao")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub cboFabricante_dropdown()
    Dim Rst As Recordset
    cboFabricante.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueFabricante WHERE id_empresa=" & ID_Empresa & " ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboFabricante.AddItem ZE(Rst.Fields("Id"), 5) & _
                                 " - " & Rst.Fields("descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub cboGrupo_DropDown()
    Dim Rst As Recordset
    cboGrupo.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueGrupos WHERE id_empresa=" & ID_Empresa & " ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboGrupo.AddItem Left(String(5, "0"), 5 - Len(Rst.Fields("Id"))) & Rst.Fields("Id") & _
                                 " - " & Rst.Fields("descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
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
            MsgBox "Erro ao localizar CST na tabela de ICMS!", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboICMSCST.AddItem Rst.Fields("cst") & " - " & Rst.Fields("descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub cboICMSOrigem_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    cboICMSOrigem.Clear
    sSQL = "SELECT * FROM TributacaoCST WHERE tabela = 'A' ORDER BY cst"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao,localizar taberla CST"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboICMSOrigem.AddItem Rst.Fields("cst") & " - " & Rst.Fields("descricao")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub cboIPICST_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    cboIPICST.Clear
    sSQL = "SELECT * FROM TributacaoCST WHERE tabela = 'I' ORDER BY cst"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao,localizar taberla CST"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboIPICST.AddItem Left(String(2, "0"), 2 - Len(Trim(Rst.Fields("cst")))) & Rst.Fields("cst") & " - " & _
                                  Rst.Fields("descricao")
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub cbosubGrupo_DropDown()
    Dim Rst As Recordset
    cboSubgrupo.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueSubGrupo WHERE id_empresa=" & ID_Empresa & " ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
            
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboSubgrupo.AddItem Left(String(5, "0"), 5 - Len(Rst.Fields("Id"))) & Rst.Fields("Id") & _
                                 " - " & Rst.Fields("descricao")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub cboStatus_DropDown()
    cboStatus.Clear
    cboStatus.AddItem "Ativo"
    cboStatus.AddItem "Inativo"
End Sub

Private Sub cboUnidade_DropDown()
    Dim Rst As Recordset
    cboUnidade.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueUnidadeMedida WHERE id_empresa=" & ID_Empresa & " ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboUnidade.AddItem IIf(IsNull(Rst.Fields("sigla")), "", Rst.Fields("sigla"))
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub chkCalcAutomCusto_Click()
    If chkCalcAutomCusto.Value = 1 Then
            txtCusto.Enabled = True
        Else
            txtCusto.Enabled = False
    End If
    CalcPrecoTabela
End Sub

Private Sub chkCalcAutomIPI_Click()
    If chkCalcAutomIPI.Value = 1 Then
            txtVlIPI.Enabled = True
        Else
            txtVlIPI.Enabled = False
    End If
    CalcPrecoTabela
End Sub



Private Sub chkCalcAutomMarkup_Click()
    If chkCalcAutomMarkup.Value = 1 Then
            txtMarkup.Enabled = True
        Else
            txtMarkup.Enabled = False
    End If
    CalcPrecoTabela
End Sub

Private Sub chkCalcAutomOutros_Click()
    If chkCalcAutomOutros.Value = 1 Then
            txtOutros.Enabled = True
        Else
            txtOutros.Enabled = False
    End If
    CalcPrecoTabela
End Sub

Private Sub chkCalcAutomPreco_Click()
    If chkCalcAutomPreco.Value = 1 Then
            txtPreco.Enabled = True
            txtPreco.Text = ChkVal(txtPreco.Text, 0, 2)
        Else
            txtPreco.Enabled = False
    End If
    CalcPrecoTabela
End Sub







Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'Me.Left = 0
    'Me.Height = 0
    LpForm
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    HDFormLocal False
    HDMenu Me, True
    ''txtID.Enabled = True
End Sub
Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    IdReg = 0
    HDMenu Me, False
    HDFormLocal True
    LpForm
    
    '** 09.11.2012 **
    'Incluir Deposito automaticamente e nao deixar o usuario mexer
    cboDeposito.Clear
    If ID_Deposito <> 0 Then
        cboDeposito.AddItem Left(String(5, "0"), 5 - Len(ID_Deposito)) & ID_Deposito & " - " & pgDescrDeposito(ID_Deposito)
        cboDeposito.Text = cboDeposito.List(0)
        'cboDeposito.Enabled = False
    End If
    '************************
    ValidarTabelaPreco
    IncluirProduto
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Selecione uma Produto.", vbInformation, "Aviso"
        Exit Sub
    End If
    HDFormLocal True
    cboDeposito.Enabled = False
    HDMenu Me, False
    ValidarTabelaPreco
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
                        "Descrição: " & txtDescricao.Text & vbCrLf, _
                        vbYesNo + vbQuestion, "Aviso") = vbYes Then
                               
                'If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                 If RegistroExcluir(strTabela, "Deposito = " & ID_Deposito & " AND Id = " & IdReg) = True Then
                    '###########################
                    '# Leonardo Aquino
                    '# 27/09/2012 - Exclui o kardex tbm
                    RegistroExcluir "EstoqueKardex", "Deposito = " & ID_Deposito & " AND IdProduto = " & IdReg
                    '###########################
                    LpForm
                End If
            End If
    End If
End Sub



Private Sub msfgDocAnexo_DblClick()
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim nmFile  As String
    
    With msfgDocAnexo
        If .TextMatrix(.Row, 0) = "" Or IdReg = 0 Then Exit Sub
        sSQL = "SELECT * FROM " & strTabela & "Arquivos " & _
             "WHERE Id_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND " & _
             "idProduto = " & IdReg & " AND " & _
             "id=" & .TextMatrix(.Row, 0)
        Set Rst = RegistroBuscar(sSQL)
        If Rst.BOF And Rst.EOF Then
                nmFile = ""
            Else
                Rst.MoveFirst
                nmFile = PgDadosConfig.pFileArmazenamento & "\EstoqueProdutos\" & LCase(cNull(Rst.Fields("NomeArquivo")))
        End If
        Rst.Close
        
        If Dir(nmFile) = "" Then Exit Sub
        ShellExecute Hwnd, "open", (nmFile), "", "", 1
    End With
End Sub





Private Sub msfgKit_DblClick()
     With msfgKit
        lKit = .Row
        txtKitId.Text = .TextMatrix(.Row, 1)
        txtKitQtd.Text = .TextMatrix(.Row, 3)
        lblKitDescr.Caption = .TextMatrix(.Row, 2)
    End With
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
            IdReg = 0
            PesquisarRegistro
            HDFormLocal False
            ''txtID.Enabled = True
        Case "Clonar Produto"
            ClonarProduto
        Case "Kardex"
            Kardex
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDFormLocal False
                MsgBox "Registro gravado com sucesso!", vbInformation, App.EXEName
                'txtID.Enabled = True
                txtID.Text = IdReg
            End If
            
        
        Case "Cancelar"
            Cancelar
            
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select
End Sub
Private Sub MontarBaseDeDados()
    Dim cReg        As Integer
    Dim vReg(100)   As Variant
    'MsgBox "Ira gerar um erro no Id do produto"
    'formManutencaoTabelas.IniciarManutencao Me, "ALTER TABLE EstoqueProduto ADD COLUMN  Deposito VARCHAR(10) default Null"
    cReg = 0
    vReg(cReg) = Array("Deposito", "10", "N"): cReg = cReg + 1
    vReg(cReg) = Array("Referencia", "50", "S"): cReg = cReg + 1
    vReg(cReg) = Array("CodigoBarras", "50", "S"): cReg = cReg + 1
    vReg(cReg) = Array("Descricao", txtDescricao.MaxLength, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Status", "20", "S"): cReg = cReg + 1
    vReg(cReg) = Array("CalcAutomPreco", "1", "N"): cReg = cReg + 1
    vReg(cReg) = Array("CalcAutomMarkUp", "1", "N"): cReg = cReg + 1
    vReg(cReg) = Array("CalcAutomOutros", "1", "N"): cReg = cReg + 1
    vReg(cReg) = Array("CalcAutomIPI", "1", "N"): cReg = cReg + 1
    vReg(cReg) = Array("CalcAutomCusto", "1", "N"): cReg = cReg + 1
    
    vReg(cReg) = Array("Outros", "15", "S"): cReg = cReg + 1
    vReg(cReg) = Array("Preco", "15", "S"): cReg = cReg + 1
    vReg(cReg) = Array("MarkUP", "15", "S"): cReg = cReg + 1
    vReg(cReg) = Array("VlIPI", "15", "S"): cReg = cReg + 1
    vReg(cReg) = Array("Custo", "15", "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("InformacoesComplementares", txtInformacoesComplementares.MaxLength, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ICMSCST", "250", "S"): cReg = cReg + 1
    vReg(cReg) = Array("ICMSOrigem", "250", "S"): cReg = cReg + 1
    vReg(cReg) = Array("IPICodEnquadramento", txtIPICodEnquadramento.MaxLength, "S"): cReg = cReg + 1
    vReg(cReg) = Array("IPICST", "250", "S"): cReg = cReg + 1
    vReg(cReg) = Array("IPIAliquota", txtIPIAliquota.MaxLength, "S"): cReg = cReg + 1
    vReg(cReg) = Array("MVA", txtMVA.MaxLength, "S"): cReg = cReg + 1
    vReg(cReg) = Array("NCM", txtNCM.MaxLength, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Saldo", txtSaldo.MaxLength, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Unidade", "250", "S"): cReg = cReg + 1
    vReg(cReg) = Array("QtdMinima", txtQtdMinima.MaxLength, "S"): cReg = cReg + 1
    vReg(cReg) = Array("QtdMedia", txtQtdMedia.MaxLength, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Grupo", "11", "N"): cReg = cReg + 1
    vReg(cReg) = Array("SubGrupo", "11", "N"): cReg = cReg + 1
    vReg(cReg) = Array("Fabricante", "11", "N"): cReg = cReg + 1
    vReg(cReg) = Array("IncluirBalanco", "1", "N"): cReg = cReg + 1
    cReg = cReg - 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vReg, cReg
    'Armazena arquivos
    cReg = 0
    vReg(cReg) = Array("idProduto", "11", "N"): cReg = cReg + 1
    vReg(cReg) = Array("Deposito", "11", "N"): cReg = cReg + 1
    vReg(cReg) = Array("Descricao", txtFileDescricao.MaxLength, "S"): cReg = cReg + 1
    vReg(cReg) = Array("NomeArquivo", "250", "S"): cReg = cReg + 1
    
    cReg = cReg - 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vReg, cReg, "Arquivos"
    
    'Armazena caso tenha kit
   
    cReg = 0
    vReg(cReg) = Array("idProduto", "11", "N"): cReg = cReg + 1
    vReg(cReg) = Array("idItemKit", "11", "N"): cReg = cReg + 1
    vReg(cReg) = Array("qtd", "11", "N"): cReg = cReg + 1
    'vReg(cReg) = Array("Descricao", txtFileDescricao.MaxLength, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("NomeArquivo", "250", "S"): cReg = cReg + 1
    
    cReg = cReg - 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vReg, cReg, "Kit"
End Sub
Private Sub Kardex()
    If IdReg = 0 Then
        MsgBox "Selecione um Produto no Estoque.", vbInformation, "Aviso"
        Exit Sub
    End If
    formEstoqueKardex.ReceberConsultaExterna (IdReg)
End Sub
Private Sub ClonarProduto()
    If IdReg = 0 Then
        MsgBox "Favor selecionar um produto.", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Deseja realmente Clonar a produto: " & Trim(txtDescricao.Text) & "?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        IdReg = 0
        'txtID.Text = ""
        'IdReg = grvRegistro
        If grvRegistro = True Then
                MsgBox "Novo produto " & Left(String(6, "0"), 6 - Len(Trim(IdReg))) & IdReg & " - " & pgDadosEstoqueProduto(IdReg).Descricao & vbCrLf & "Produto clonado com sucesso.", vbInformation, "Aviso"
                PesquisarRegistro
            Else
                MsgBox "Erro ao criar clone do produto.", vbInformation, "Aviso"
        End If
    End If
        
    
End Sub
Private Function ValidaDados() As Boolean
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    If Trim(cboDeposito.Text) = "" Then
        MsgBox "Depósito invalido. Favor verificar.", vbInformation, "Aviso"
        ValidaDados = False
        Exit Function
    End If
    
    If Trim(cboStatus.Text) = "" Then
        MsgBox "Status invalido. Favor verificar.", vbInformation, "Aviso"
        ValidaDados = False
        Exit Function
    End If
    
    If lblDescrNCM.Caption = "<<< Codigo NCM invalido. >>>" Or Trim(lblDescrNCM.Caption) = "" Then
        MsgBox "NCM invalido. Favor verificar.", vbInformation, "Aviso"
        ValidaDados = False
        Exit Function
    End If
    
    If Trim(cboGrupo.Text) = "" Then
        MsgBox "Grupo invalido. Favor verificar.", vbInformation, "Aviso"
        ValidaDados = False
        Exit Function
    End If
    If Trim(cboSubgrupo.Text) = "" Then
        MsgBox "Subgrupo invalido. Favor verificar.", vbInformation, "Aviso"
        ValidaDados = False
        Exit Function
    End If
    If cboUnidade.Text = "" Then
            MsgBox "Unidade de armazenamento invalida. Favor verificar.", vbInformation, "Aviso"
            ValidaDados = False
            Exit Function
        Else
            sSQL = "SELECT * FROM EstoqueUnidadeMedida WHERE Sigla ='" & Trim(cboUnidade.Text) & "'"
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                MsgBox "Unidade não cadastrada no sistema!", vbInformation, "Aviso"
                ValidaDados = False
                Rst.Close
                Exit Function
            End If
            Rst.Close
    End If
    
    If cboIPICST.Text = "" Then
        MsgBox "Codigo de Situação Tributaria IPI (CST) invalido. Favor verificar.", vbInformation, "Aviso"
        ValidaDados = False
        Exit Function
    End If

    If ChkVal(Trim(PgDadosNCM("NCM", Trim(txtNCM.Text), "S").pIPI), 0, 2) <> ChkVal(Trim(txtIPIAliquota.Text), 0, 2) Then
        If MsgBox("Aliquota de IPI difere da cadastrada no sistema! Deseja continuare?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            ValidaDados = False
            Exit Function
        End If
    End If
    'CST ICMS
    If Trim(cboICMSCST.Text) = "" Then
        MsgBox "Codigo de Situação Tributaria ICMS (CST) invalido. Favor verificar.", vbInformation, "Aviso"
        ValidaDados = False
        Exit Function
    End If
    If PgDadosEmpresa(ID_Empresa).RegimeTrib = "3" Then
            sSQL = "SELECT * FROM TributacaoCST WHERE tabela = 'B' AND CST = '" & Trim(Left(cboICMSCST.Text, 3)) & "'"
        Else
            sSQL = "SELECT * FROM TributacaoCST WHERE tabela = 'C' AND CST = '" & Trim(Left(cboICMSCST.Text, 3)) & "'"
    End If
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        MsgBox "Erro ao localizar CST na tabela de ICMS!", vbInformation, "Aviso"
        ValidaDados = False
        Exit Function
    End If
    Rst.Close




    ValidaDados = True
End Function
Private Function grvRegistro() As Boolean
    Dim vReg(199)   As Variant
    Dim i           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer 'Contador de Registros
    
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim SaldoKardex As String
    Dim Movimento   As String
    Dim difSaldo    As String
    
    'Validar os dados para gravacao
    If ValidaDados = False Then
        'MsgBox "Dados do formulario nao gravados!", vbInformation, "Aviso"
        grvRegistro = False
        Exit Function
    End If
  
    
    '**************************************************************************************************************************
    '**** Alterar os dados do Produto
    '**** Função criada em 04/07/2011
    '**** Modificada em 09/11/2012 - Inclusao do registro do deposito
    '**************************************************************************************************************************
    cReg = 0
    'vReg(cReg) = Array("Deposito", ID_Deposito, "N"): cReg = cReg + 1
    For i = 0 To Me.Controls.Count - 1
        Set Controle = Me.Controls(i)
        
        If TypeOf Controle Is TextBox Then
            '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            Select Case Controle.Name
                Case "txtID"
                Case "txtKitId"
                Case "txtKitQtd"
                Case "txtFile"
                Case "txtFileDescricao"
                Case "txtSaldo"
                    If IdReg = 0 Then
                            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), "0", "S"): cReg = cReg + 1
                        'Else
                            'vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Trim(Controle.Text), "S"): cReg = cReg + 1
                    End If
                Case Else
                    vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Trim(Controle.Text), "S"): cReg = cReg + 1
            End Select
            '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        End If
        
        
        
        If TypeOf Controle Is ComboBox Then
            Select Case Controle.Name
                
                Case "cboDeposito"
                    'vReg(cReg) = Array("Deposito", ID_Deposito, "N"): cReg = cReg + 1
                    vReg(cReg) = Array("Deposito", Left(Trim(Controle.Text), 5), "N")
                    cReg = cReg + 1
                Case "cboFabricante", "cboFabricante"
                    vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Left(Trim(Controle.Text), 5), "S")
                    cReg = cReg + 1
                Case "cboGrupo", "cboSubGrupo"
                    vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Left(Trim(Controle.Text), 5), "S")
                    cReg = cReg + 1
                Case "cboICMSOrigem"
                    vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Left(Trim(Controle.Text), 1), "S")
                    cReg = cReg + 1
                Case "cboICMSCST"
                    vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Trim(Left(Controle.Text, 3)), "S")
                    cReg = cReg + 1
                Case "cboIPICST"
                    vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Left(Trim(Controle.Text), 2), "S")
                    cReg = cReg + 1
                Case Else
                    vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), UCase(Controle.Text), "S")
                    cReg = cReg + 1
            End Select
        End If
        If TypeOf Controle Is CheckBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Value, "S")
            cReg = cReg + 1
        End If
        
        
    Next
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
    '**************************************************************************************************************************
    'Armazena os Arquivos em anexo
    '
    Dim nmFileDestino   As String
    Dim idFile          As Integer
    Dim sFile           As String
    With msfgDocAnexo
        For i = 1 To .Rows - 1
            nmFileDestino = Dir(Trim(.TextMatrix(i, 2)), vbDirectory)
            cReg = 0
            vReg(cReg) = Array("IdProduto", IdReg, "S"): cReg = cReg + 1
            vReg(cReg) = Array("Deposito", ID_Deposito, "S"): cReg = cReg + 1
            vReg(cReg) = Array("Descricao", .TextMatrix(i, 1), "S"): cReg = cReg + 1
            If Trim(nmFileDestino) <> "" Then
                vReg(cReg) = Array("NomeArquivo", nmFileDestino, "S"): cReg = cReg + 1
            End If
            cReg = cReg - 1
            
            If Trim(.TextMatrix(i, 0)) = "" Then
            'Incluir novo Documento
                    If MoverPastaArquivos(.TextMatrix(i, 2), nmFileDestino) = True Then
                        idFile = RegistroIncluir(strTabela & "Arquivos", vReg, cReg)
                        .TextMatrix(i, 0) = idFile
                        .TextMatrix(i, 2) = "< Armazenado >"
                    End If
                Else
                    'Altera o Documento Existente
                    If .TextMatrix(i, 2) <> "< Armazenado >" Then
                        If ExcluirFile(.TextMatrix(i, 0)) = True Then
                            If MoverPastaArquivos(.TextMatrix(i, 2), nmFileDestino) = True Then
                                .TextMatrix(i, 2) = "< Armazenado >"
                                RegistroAlterar strTabela & "Arquivos", vReg, cReg, "id = " & .TextMatrix(i, 0)
                            End If
                        End If
                    End If
            End If
        Next
        'Exclui os registros dos arquivos removidos
        'Cria uma string com os IDs
        sFile = ""
        For i = 1 To .Rows - 1
            sFile = sFile & IIf(sFile = "", "", ",") & .TextMatrix(i, 0)
        Next
    
    End With
    If Trim(sFile) <> "" Then
            sSQL = "SELECT * FROM " & strTabela & "Arquivos" & " WHERE Id_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND IdProduto = " & IdReg & " AND Id NOT IN (" & sFile & ")"
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                Else
                    Rst.MoveFirst
                    Do Until Rst.EOF
                        ExcluirFile Rst.Fields("ID")
                        Rst.MoveNext
                    Loop
            End If
            Rst.Close
    
   
            RegistroExcluir strTabela & "Arquivos", "Deposito = " & ID_Deposito & " AND idProduto = " & IdReg & " AND id NOT IN (" & sFile & ")"
        Else
            sSQL = "SELECT * FROM " & strTabela & "Arquivos" & " WHERE Id_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND IdProduto = " & IdReg
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                Else
                    Rst.MoveFirst
                    Do Until Rst.EOF
                        ExcluirFile Rst.Fields("ID")
                        Rst.MoveNext
                    Loop
            End If
            Rst.Close
            RegistroExcluir strTabela & "Arquivos", "Deposito = " & ID_Deposito & " AND idProduto = " & IdReg
    End If
    
    '**************************************************************************************************************************
    '****  Alterar os dados do KARDEX
    '**** Função criada em 04/07/2011
    '**** Alterar o KARDEX primeiro pois a funcao MovimentarEstoque pega o saldo da tabela base EstoqueProduto
    '**************************************************************************************************************************
    sSQL = "SELECT * FROM EstoqueKardex WHERE ID_Empresa =" & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND IdProduto = " & IdReg & " ORDER BY ID"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            SaldoKardex = pgDadosEstoqueProduto(IdReg).Saldo '"0" 'ChkVal(txtSaldo.Text, 0, cDecQtd)
        Else
            Rst.MoveLast
            SaldoKardex = ChkVal(IIf(IsNull(Rst.Fields("Saldo")), "0", Rst.Fields("Saldo")), 0, cDecQtd)
    End If
    Dim Mov As String
    If Val(SaldoKardex) > Val(txtSaldo.Text) Then
            Movimento = "s"
            difSaldo = Val(SaldoKardex) - Val(txtSaldo.Text)
        Else
            Movimento = "e"
            difSaldo = Val(txtSaldo.Text) - Val(SaldoKardex)
    End If
    '14.11.2012 - So deixa entrar se o saldo do sistema for diferente do saldo lancado
    If Val(SaldoKardex) <> Val(txtSaldo.Text) Then
        MovimentarEstoque Movimento, IdReg, Date, Format(Time, "HHMMSS"), ChkVal(difSaldo, 0, cDecQtd), ChkVal(txtCusto.Text, 0, cDecMoeda), _
                        ChkVal(Val(ChkVal(difSaldo, 0, cDecQtd)) * Val(ChkVal(txtCusto.Text, 0, cDecMoeda)), 0, cDecMoeda), "AJUSTE AUTOMATICO DE SALDO NO KARDEX - [Cadastro Produto]"
    End If
    Rst.Close
    '****************************************************************************************************************************
    
    '**************************************************************************************************************************
    '**** Alterar os dados do Kit
    '**** Função criada em 11/12/2017
    '**** Alterar os dados do Kit do produto
    '**************************************************************************************************************************
    'Exclui registros anteriores
    RegistroExcluir strTabela & "Kit", "idProduto = " & IdReg
    
    If msfgKit.Rows = 1 Then Exit Function
    With msfgKit
        For i = 1 To .Rows - 1
           cReg = 0
           vReg(cReg) = Array("IdProduto", IdReg, "S"): cReg = cReg + 1
           vReg(cReg) = Array("IdItemKit", .TextMatrix(i, 1), "N"): cReg = cReg + 1
           vReg(cReg) = Array("qtd", .TextMatrix(i, 3), "N"): cReg = cReg + 1
           
           cReg = cReg - 1
        
           RegistroIncluir strTabela & "Kit", vReg, cReg
        Next
    End With
End Function
Private Function MoverPastaArquivos(fileOrigem As String, nmFileDestino As String) As Boolean
    On Error GoTo TrtErroFile
    Dim fileXMLDestino As String
    
    If Trim(fileOrigem) = "" Then
        MoverPastaArquivos = False
        Exit Function
    End If
    fileXMLDestino = PgDadosConfig.pFileArmazenamento & "\EstoqueProdutos"
    If Dir(fileXMLDestino, vbDirectory) = "" Then
        MkDir fileXMLDestino
    End If
    fileXMLDestino = fileXMLDestino & "\" & rc(nmFileDestino)
    FileCopy fileOrigem, fileXMLDestino
    MoverPastaArquivos = True
    Exit Function
TrtErroFile:
    MsgBox "Erro ao armazenar o arquivo." & vbCrLf & _
    Err.Description, vbInformation, "Aviso - Erro n.: " & Err.Number
    'Resume Next
    MoverPastaArquivos = False
End Function
Private Function ExcluirFile(Id As Integer) As Boolean
    'On Error GoTo TrtErroFile
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim nmFile  As String
    sSQL = "SELECT * FROM   s WHERE Id_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND id = " & Id
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            ExcluirFile = True
            Rst.Close
            Exit Function
        Else
            Rst.MoveFirst
            nmFile = cNull(Rst.Fields("NomeArquivo"))
            Rst.Close
    End If
    nmFile = PgDadosConfig.pFileArmazenamento & "\EstoqueProdutos\" & nmFile
    If Dir(nmFile) = "" Then
        ExcluirFile = True
        Exit Function
    End If
    Kill nmFile
    ExcluirFile = True
    Exit Function
TrtErroFile:
    MsgBox "Erro ao Excluir Arquivo." & vbCrLf & _
    Err.Description, vbInformation, "Aviso - Erro n.: " & Err.Number
    'Resume Next
    ExcluirFile = False
        
        
        
End Function
Private Sub MostrarDados()
    Dim sSQL    As String
    Dim tmp     As String
    sSQL = "SELECT * FROM " & strTabela & " WHERE Id = " & IdReg
    
    ExibirDados Me, sSQL
    
    
    'Carregar os dados do Desposito
    tmp = cboDeposito.Text
    If Trim(tmp) <> "" Then
            cboDeposito.Clear
            cboDeposito.AddItem Left(String(5, "0"), 5 - Len(tmp)) & tmp & " - " & pgDescrDeposito(CInt(tmp))
            cboDeposito.Text = cboDeposito.List(0)
        Else
            cboDeposito.Clear
    End If
    
    
    
    'Carregar os dados do GRUPO
    tmp = cboGrupo.Text
    If Trim(tmp) <> "" Then
            cboGrupo.Clear
            cboGrupo.AddItem Left(String(5, "0"), 5 - Len(tmp)) & tmp & " - " & pgDescrGrupo(tmp)
            cboGrupo.Text = cboGrupo.List(0)
        Else
            cboGrupo.Clear
    End If
    
     'Carregar os dados do SUBGRUPO
    tmp = cboSubgrupo.Text
    If Trim(tmp) <> "" Then
            cboSubgrupo.Clear
            cboSubgrupo.AddItem Left(String(5, "0"), 5 - Len(tmp)) & tmp & " - " & pgDescrSubGrupo(tmp)
            cboSubgrupo.Text = cboSubgrupo.List(0)
        Else
            cboSubgrupo.Clear
    End If
    
    
     'Carregar os dados do Fabricante
    tmp = cboFabricante.Text
    If Trim(tmp) <> "" And Trim(tmp) <> "0" Then
            cboFabricante.Clear
            cboFabricante.AddItem Left(String(5, "0"), 5 - Len(tmp)) & tmp & " - " & pgDescrFabricante(CInt(tmp))
            cboFabricante.Text = cboFabricante.List(0)
        Else
            cboFabricante.Clear
    End If
    
    
    'Carregar os dados do ICMS ORIGEM
    With cboICMSOrigem
        tmp = .Text
        .Clear
        If Trim(tmp) <> "" Then
            .AddItem tmp & " - " & PgDadosCST(tmp, "ORIGEM").Descricao
            .Text = .List(0)
        End If
    End With
    
    '*********************************************
    With cboICMSCST
        tmp = .Text
        .Clear
        If Trim(tmp) <> "" Then
            .AddItem tmp & " - " & PgDadosCST(tmp, "ICMS").Descricao
            .Text = .List(0)
        End If
    End With
    '*********************************************
    With cboIPICST
        tmp = Trim(.Text)
        .Clear
        If Trim(tmp) <> "" Then
            .AddItem tmp & " - " & PgDadosCST(tmp, "IPI").Descricao
            .Text = .List(0)
        End If
    End With
    '*********************************************
    
    txtCusto.Text = ConvMoeda(txtCusto.Text)
    
    '*************************************************

    ListarDezUltEntradas
    ListarDezUltSaidas
    ListarArquivos
    ListarItensDaGrade
    

End Sub
Private Sub ListarDezUltEntradas()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    msfgUltEntradas.Rows = 1
    sSQL = "SELECT FNFeEntI.det_uCom, FNFeEntI.det_qCom,FNFeEntI.det_vUnCom, FNFeEntI.idNFe, FNFeEntI.det_idProduto, " & _
                   "FNFeEnt.Id, FNFeEnt.IdNFe, FNFeEnt.emit_xNome,FNFeEnt.ide_nNF, FNFeEnt.ide_dEmi " & _
           "FROM FaturamentoNFeEntrada AS FNFeEnt ,FaturamentoNFeEntradaItens AS FNFeEntI " & _
           "WHERE FNFeEntI.IdNFe=FNFeEnt.idNFe AND FNFeEntI.det_IdProduto=" & IdReg & " " & _
           "ORDER BY FNFeEnt.Id DESC LIMIT 10"
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveLast
            
            Do Until Rst.BOF
                With msfgUltEntradas
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("ide_nNF")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("ide_dEmi")
                    .TextMatrix(.Rows - 1, 2) = Rst.Fields("emit_xNome")
                    .TextMatrix(.Rows - 1, 3) = ChkVal(Rst.Fields("det_vUnCom"), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 4) = ChkVal(Rst.Fields("det_qCom"), 0, cDecQtd) & "/" & Rst.Fields("det_uCom")
                    
                End With
                Rst.MovePrevious
                
            Loop
    End If
    Rst.Close
End Sub
Private Sub ListarDezUltSaidas()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    msfgUltSaidas.Rows = 1
    sSQL = "SELECT FNFeI.det_uCom, FNFeI.det_qCom,FNFeI.det_vUnCom, FNFeI.idNFe, FNFeI.det_idProduto, " & _
                   "FNFe.id, FNFe.IdNFe, FNFe.dest_xNome,FNFe.ide_nNF, FNFe.ide_dEmi " & _
           "FROM FaturamentoNFe AS FNFe ,FaturamentoNFeItens AS FNFeI " & _
           "WHERE FNFeI.IdNFe=FNFe.idNFe AND FNFeI.det_IdProduto=" & IdReg & " " & _
           "ORDER BY FNFe.id DESC LIMIT 10"
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveLast
            Do Until Rst.BOF
                With msfgUltSaidas
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("ide_nNF")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("ide_dEmi")
                    .TextMatrix(.Rows - 1, 2) = Rst.Fields("dest_xNome")
                    .TextMatrix(.Rows - 1, 3) = ChkVal(Rst.Fields("det_vUnCom"), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 4) = ChkVal(Rst.Fields("det_qCom"), 0, cDecQtd) & "/" & Rst.Fields("det_uCom")
                    
                End With
                Rst.MovePrevious
                
            Loop
    End If
    Rst.Close
End Sub

Private Sub ListarArquivos()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    msfgDocAnexo.Rows = 1
    sSQL = "SELECT * " & _
           "FROM " & strTabela & "Arquivos " & _
           "WHERE IdProduto=" & IdReg
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                With msfgDocAnexo
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("id")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("Descricao")
                    .TextMatrix(.Rows - 1, 2) = "< Armazenado >"
                    '.TextMatrix(.Rows - 1, 3) = ChkVal(Rst.Fields("det_vUnCom"), 0, cDecMoeda)
                    '.TextMatrix(.Rows - 1, 4) = ChkVal(Rst.Fields("det_qCom"), 0, cDecQtd) & "/" & Rst.Fields("det_uCom")
                End With
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub


Private Sub IncluirProduto()
    'Seleciona o DEPOSITO
    'Dim Rst As Recordset
    'cboDeposito.Clear
    'Set Rst = RegistroBuscar("SELECT * FROM EstoqueDeposito ORDER BY Descricao")
    'If Rst.BOF And Rst.EOF Then
    '
    '        Rst.Close
    '        Exit Sub
    '    Else
    '        Rst.MoveFirst
    '        cboDeposito.AddItem Left(String(5, "0"), 5 - Len(Rst.Fields("Id"))) & Rst.Fields("Id") & _
    '                            " - " & Rst.Fields("Referencia") & " - " & Rst.Fields("descricao")
    '        cboDeposito.Text = cboDeposito.List(0)
    'End If
    'Rst.Close
    'Seleciona o status
   
    cboStatus.AddItem "Ativo"
    cboStatus.Text = cboStatus.List(0)
    chkIncluirBalanco.Value = 1
End Sub


Private Sub txtCusto_Change()
    CalcPrecoTabela
End Sub

Private Sub txtCusto_GotFocus()
    With txtCusto
        .Text = ChkVal(.Text, 0, cDecMoeda)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtCusto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Exit Sub
    If txtCusto.SelLength = Len(txtCusto.Text) Then
        txtCusto.Text = ""
    End If
    KeyAscii = ChkVal(txtCusto.Text, KeyAscii, cDecMoeda)
    
End Sub

Private Sub txtCusto_LostFocus()
    txtCusto.Text = ConvMoeda(txtCusto.Text)
End Sub


Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IdReg = 0
        PesquisarRegistro
    End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        IdReg = Trim(txtID.Text)
        PesquisarRegistro
        Exit Sub
    End If
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtIPIAliquota_Change()
    CalcPrecoTabela
End Sub

Private Sub txtIPIAliquota_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtIPIAliquota.Text, KeyAscii, 2)
End Sub

Private Sub txtIPICodEnquadramento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If

End Sub



Private Sub txtKitId_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 114 Then
        IdReg = 0
        pesqItemKit
        
    End If
End Sub

Private Sub txtKitId_KeyPress(KeyAscii As Integer)
  If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        
        pesqItemKit (Trim(txtKitId.Text))
        
        Exit Sub
    End If
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtKitQtd_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub txtMarkup_Change()
    CalcPrecoTabela
End Sub

Private Sub txtMarkup_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtMarkup.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtMVA_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtMVA.Text, KeyAscii, 3)
End Sub


Private Sub txtNCM_Change()
    lblDescrNCM.Caption = ChkNCM(txtNCM.Text)
    If lblDescrNCM.Caption = "<<< Codigo NCM invalido. >>>" Then
            lblDescrNCM.ForeColor = vbRed
        Else
            lblDescrNCM.ForeColor = vbBlue
    End If
    
End Sub

Private Sub txtNCM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarNCM
    End If
End Sub

Private Sub txtNCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If

End Sub


Private Sub txtOutros_Change()
    CalcPrecoTabela
End Sub

Private Sub txtOutros_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtOutros.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtPreco_Change()
    CalcPrecoTabela
End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtPreco.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtQtdMedia_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtQtdMedia.Text, KeyAscii, cDecQtd)
End Sub


Private Sub txtQtdMinima_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtQtdMinima.Text, KeyAscii, cDecQtd)
End Sub

Private Sub CalcPrecoTabela()
    Dim VlCusto     As String
    Dim VlIPI       As String
    Dim VlOutros    As String
    Dim vlMarkup    As String
    Dim vlReposicao As String
    
    VlCusto = ChkVal(txtCusto.Text, 0, 2)
    VlOutros = ChkVal(txtOutros.Text, 0, 2)
    
    If chkCalcAutomIPI.Value = 0 Then
            VlIPI = Val(ChkVal(txtIPIAliquota.Text, 0, 2)) * Val(ChkVal(txtCusto.Text, 0, 2)) / 100
            VlIPI = ChkVal(VlIPI, 0, 2)
            txtVlIPI.Text = ConvMoeda(VlIPI)
        Else
            VlIPI = ChkVal(txtVlIPI.Text, 0, 2)
    End If
    
    vlReposicao = Val(VlCusto) + Val(VlIPI) + Val(VlOutros)
    
    vlMarkup = Val(ChkVal(txtMarkup.Text, 0, 2)) * Val(ChkVal(vlReposicao, 0, 2)) / 100
    
    If chkCalcAutomPreco.Value = 0 Then
            txtPreco.Text = ConvMoeda(Val(ChkVal(vlReposicao, 0, 2)) + Val(ChkVal(vlMarkup, 0, 2)))
        Else
            txtPreco.Text = txtPreco.Text
    End If
End Sub
Private Sub ValidarTabelaPreco()
    chkCalcAutomCusto_Click
    chkCalcAutomIPI_Click
    chkCalcAutomOutros_Click
    chkCalcAutomMarkup_Click
    chkCalcAutomPreco_Click
    
    
End Sub

Private Sub txtSaldo_GotFocus()
    With txtSaldo
        .Text = ChkVal(.Text, 0, cDecQtd)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
    If txtSaldo.SelLength = Len(txtSaldo.Text) Then
        txtSaldo.Text = ""
    End If
      KeyAscii = ChkVal(txtSaldo.Text, KeyAscii, cDecQtd)
End Sub

Private Sub txtVlIPI_Change()
    CalcPrecoTabela
End Sub

Private Sub txtVlIPI_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtVlIPI.Text, KeyAscii, cDecMoeda)
End Sub
Private Sub PesquisarNCM()
    Dim idNCM   As Integer
    Dim Rst     As Recordset
    Dim sSQL    As String
    idNCM = formBuscar.IniciarBusca("TributacaoNCM")
    If idNCM = 0 Then
        Exit Sub
    End If
    sSQL = "SELECT * FROM TributacaoNCM WHERE id = " & idNCM
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar NCM", vbInformation, "Aviso"
            txtNCM.Text = ""
        Else
            Rst.MoveFirst
            txtNCM.Text = Trim(Rst.Fields("NCM"))
    End If
    Rst.Close
End Sub
Public Sub RecebendoDadosProduto(Optional SRef As String, _
                                 Optional sDescr As String, _
                                 Optional sCodBar As String, _
                                 Optional sNCM As String, _
                                 Optional sMVA As String, _
                                 Optional sOrig As String, _
                                 Optional sCST_ICMS As String, _
                                 Optional spIPI As String, _
                                 Optional sCST_IPI As String, _
                                 Optional sPeso As String, _
                                 Optional sUnid As String, _
                                 Optional sCusto As String)
    IdReg = 0
    HDMenu Me, False
    HDFormLocal True
    LpForm
    ValidarTabelaPreco
    IncluirProduto
    txtReferencia.Text = SRef
    txtDescricao.Text = sDescr
    txtCodigoBarras.Text = sCodBar
    txtNCM.Text = sNCM
    txtMVA.Text = sMVA
    If Trim(sOrig) <> "" Then
        With cboICMSOrigem
            .Clear
            .AddItem sOrig & " - " & PgDadosCST(sOrig, "Origem").Descricao
            .Text = .List(0)
        End With
    End If
    If Trim(sCST_ICMS) <> "" Then
        With cboICMSCST
            .Clear
            .AddItem sCST_ICMS & " - " & PgDadosCST(sCST_ICMS, "ICMS").Descricao
            .Text = .List(0)
        End With
    End If
    txtIPIAliquota.Text = spIPI
    If Trim(sCST_IPI) <> "" Then
        With cboIPICST
            .Clear
            .AddItem sCST_IPI & " - " & PgDadosCST(sCST_IPI, "IPI").Descricao
            .Text = .List(0)
        End With
    End If
    If Trim(sUnid) <> "" Then
        With cboUnidade
            .Clear
            .AddItem sUnid
            .Text = .List(0)
        End With
    End If
    txtSaldo.Text = sPeso
    txtCusto.Text = ConvMoeda(ChkVal(sCusto, 0, cDecMoeda))
    Me.Show
End Sub
Public Sub pesqLoadForm(idProduto As Integer, Optional xp As String)
    On Error Resume Next
    If idProduto = 0 Then
            IdReg = 0
            Me.Show
        Else
            IdReg = idProduto
            PesquisarRegistro
            'Alterar
            Me.Show
    End If
End Sub
Public Sub LoadFormExcluir(idProduto As Integer)
    On Error Resume Next
    If idProduto = 0 Then
            IdReg = 0
            Me.Show
        Else
            
            IdReg = idProduto
            PesquisarRegistro
            
            Excluir
            
    End If
End Sub

