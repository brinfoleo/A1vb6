VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formFaturamentoNFeEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Nota Fiscal de Entrada"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   12120
   Begin MSComDlg.CommonDialog cd 
      Left            =   6420
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab sstNF 
      Height          =   5595
      Left            =   60
      TabIndex        =   6
      Top             =   3840
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   9869
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dados dos Produtos"
      TabPicture(0)   =   "formFaturamentoNFeEntrada.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Faturamento/Transportador"
      TabPicture(1)   =   "formFaturamentoNFeEntrada.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Total da Nota Fiscal / Obs."
      TabPicture(2)   =   "formFaturamentoNFeEntrada.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame9"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame9 
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
         Height          =   1575
         Left            =   180
         TabIndex        =   113
         Top             =   3840
         Width           =   11535
         Begin VB.TextBox txtObs 
            Height          =   1215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   114
            Text            =   "formFaturamentoNFeEntrada.frx":0054
            Top             =   240
            Width           =   11235
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Total da Nota Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         TabIndex        =   85
         Top             =   420
         Width           =   11595
         Begin VB.TextBox txtvCredICMSSN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1980
            TabIndex        =   118
            Text            =   "Text1"
            Top             =   1920
            Width           =   2115
         End
         Begin VB.TextBox txtBCICMS 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1980
            TabIndex        =   99
            Text            =   "Text1"
            Top             =   360
            Width           =   2115
         End
         Begin VB.TextBox txtvICMS 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1980
            TabIndex        =   98
            Text            =   "Text1"
            Top             =   750
            Width           =   2115
         End
         Begin VB.TextBox txtBCICMSST 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1980
            TabIndex        =   97
            Text            =   "Text1"
            Top             =   1140
            Width           =   2115
         End
         Begin VB.TextBox txtvICMSST 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1980
            TabIndex        =   96
            Text            =   "Text1"
            Top             =   1530
            Width           =   2115
         End
         Begin VB.TextBox txtvProduto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7140
            TabIndex        =   95
            Text            =   "Text1"
            Top             =   300
            Width           =   2115
         End
         Begin VB.TextBox txtvFrete 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7140
            TabIndex        =   94
            Text            =   "Text1"
            Top             =   960
            Width           =   2115
         End
         Begin VB.TextBox txtvSeguro 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7140
            TabIndex        =   93
            Text            =   "Text1"
            Top             =   1290
            Width           =   2115
         End
         Begin VB.TextBox txtvDesconto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7140
            TabIndex        =   92
            Text            =   "Text1"
            Top             =   1620
            Width           =   2115
         End
         Begin VB.TextBox txtvOutras 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7140
            TabIndex        =   91
            Text            =   "Text1"
            Top             =   1950
            Width           =   2115
         End
         Begin VB.TextBox txtvIPI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7140
            TabIndex        =   90
            Text            =   "Text1"
            Top             =   630
            Width           =   2115
         End
         Begin VB.TextBox txtvTotalNF 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7140
            TabIndex        =   89
            Text            =   "Text1"
            Top             =   2340
            Width           =   2115
         End
         Begin VB.CheckBox chkTotaisAutomatico 
            Caption         =   "Calcular os totais automaticamente com base na grade de produtos."
            Height          =   195
            Left            =   6300
            TabIndex        =   88
            Top             =   3060
            Width           =   5115
         End
         Begin VB.TextBox txtvPIS 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1980
            TabIndex        =   87
            Text            =   "Text1"
            Top             =   2340
            Width           =   2115
         End
         Begin VB.TextBox txtvCOFINS 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1980
            TabIndex        =   86
            Text            =   "Text1"
            Top             =   2700
            Width           =   2115
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            Caption         =   "Credito de ICMS:"
            Height          =   255
            Left            =   180
            TabIndex        =   117
            Top             =   1980
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Base de Calculo ICMS:"
            Height          =   195
            Left            =   180
            TabIndex        =   112
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor do ICMS:"
            Height          =   195
            Left            =   720
            TabIndex        =   111
            Top             =   750
            Width           =   1155
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Base Calc. ICMS-ST:"
            Height          =   195
            Left            =   300
            TabIndex        =   110
            Top             =   1155
            Width           =   1575
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor do ICMS-ST:"
            Height          =   195
            Left            =   420
            TabIndex        =   109
            Top             =   1545
            Width           =   1455
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor Total do Produto:"
            Height          =   195
            Left            =   5340
            TabIndex        =   108
            Top             =   345
            Width           =   1695
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor do Frete:"
            Height          =   195
            Left            =   5820
            TabIndex        =   107
            Top             =   1050
            Width           =   1215
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor do Seguro:"
            Height          =   195
            Left            =   5700
            TabIndex        =   106
            Top             =   1380
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor do Desconto:"
            Height          =   195
            Left            =   5460
            TabIndex        =   105
            Top             =   1710
            Width           =   1575
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Outras Desp. Acess.:"
            Height          =   195
            Left            =   5400
            TabIndex        =   104
            Top             =   1995
            Width           =   1635
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor do IPI:"
            Height          =   195
            Left            =   6060
            TabIndex        =   103
            Top             =   690
            Width           =   975
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor Total da Nota:"
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
            Left            =   5220
            TabIndex        =   102
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor do PIS:"
            Height          =   195
            Left            =   840
            TabIndex        =   101
            Top             =   2400
            Width           =   1035
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor COFINS:"
            Height          =   195
            Left            =   840
            TabIndex        =   100
            Top             =   2760
            Width           =   1035
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Transportador"
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
         Left            =   -74760
         TabIndex        =   78
         Top             =   3900
         Width           =   11415
         Begin VB.ComboBox cboTranspNome 
            Height          =   315
            Left            =   1140
            TabIndex        =   81
            Text            =   "Combo1"
            Top             =   720
            Width           =   7035
         End
         Begin VB.TextBox txtTranspCNPJ 
            Height          =   315
            Left            =   1140
            MaxLength       =   14
            TabIndex        =   80
            Text            =   "Text1"
            Top             =   300
            Width           =   2895
         End
         Begin VB.TextBox txtFreteConta 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6300
            MaxLength       =   1
            TabIndex        =   79
            Text            =   "Text1"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "CNPJ:"
            Height          =   195
            Left            =   240
            TabIndex        =   84
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Nome:"
            Height          =   255
            Left            =   180
            TabIndex        =   83
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Frete por Conta:"
            Height          =   255
            Left            =   4920
            TabIndex        =   82
            Top             =   300
            Width           =   1275
         End
      End
      Begin VB.Frame Frame6 
         Height          =   5055
         Left            =   -74820
         TabIndex        =   39
         Top             =   420
         Width           =   11655
         Begin VB.Frame Frame8 
            Height          =   2775
            Left            =   120
            TabIndex        =   40
            Top             =   2160
            Width           =   11355
            Begin VB.OptionButton optBCICMS 
               Caption         =   "Valor do Produto + IPI"
               Height          =   195
               Index           =   1
               Left            =   6600
               TabIndex        =   33
               Top             =   2400
               Width           =   2115
            End
            Begin VB.OptionButton optBCICMS 
               Caption         =   "Valor do Produto"
               Height          =   195
               Index           =   0
               Left            =   6600
               TabIndex        =   32
               Top             =   2160
               Value           =   -1  'True
               Width           =   1515
            End
            Begin VB.TextBox txtvICMSp 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   6600
               TabIndex        =   31
               Text            =   "Text1"
               Top             =   1800
               Width           =   1815
            End
            Begin VB.TextBox txtpICMSp 
               Height          =   285
               Left            =   6600
               TabIndex        =   30
               Text            =   "Text1"
               Top             =   1440
               Width           =   1815
            End
            Begin VB.TextBox txtBCICMSp 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   6600
               TabIndex        =   29
               Text            =   "Text1"
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox txtvIPIp 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3780
               TabIndex        =   28
               Text            =   "Text1"
               Top             =   1440
               Width           =   1575
            End
            Begin VB.TextBox txtpIPIp 
               Height          =   285
               Left            =   3780
               TabIndex        =   27
               Text            =   "Text1"
               Top             =   1080
               Width           =   1155
            End
            Begin VB.TextBox txtvProd 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1260
               TabIndex        =   26
               Text            =   "Text1"
               Top             =   2160
               Width           =   1635
            End
            Begin VB.TextBox txtCFOP 
               Height          =   285
               Left            =   6600
               MaxLength       =   4
               TabIndex        =   22
               Text            =   "Text1"
               Top             =   690
               Width           =   1155
            End
            Begin VB.TextBox txtCST 
               Height          =   285
               Left            =   3780
               MaxLength       =   5
               TabIndex        =   21
               Text            =   "Text1"
               Top             =   690
               Width           =   1155
            End
            Begin VB.TextBox txtNCM 
               Height          =   315
               Left            =   1260
               MaxLength       =   8
               TabIndex        =   20
               Text            =   "Text1"
               Top             =   660
               Width           =   1635
            End
            Begin VB.CommandButton btoCadastrarItem 
               Caption         =   "Incluir Item"
               Height          =   495
               Left            =   9480
               TabIndex        =   36
               ToolTipText     =   "Incluir item no estoque..."
               Top             =   1260
               Width           =   1575
            End
            Begin VB.CommandButton btoRemover 
               Caption         =   "&Remover"
               Height          =   495
               Left            =   9480
               TabIndex        =   35
               Top             =   720
               Width           =   1575
            End
            Begin VB.CommandButton btoAdicionar 
               Caption         =   "&Adicionar"
               Height          =   495
               Left            =   9480
               TabIndex        =   34
               Top             =   180
               Width           =   1575
            End
            Begin VB.ComboBox cboUnidade 
               Height          =   315
               Left            =   1260
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   1065
               Width           =   1635
            End
            Begin VB.TextBox txtQtd 
               Height          =   285
               Left            =   1260
               TabIndex        =   24
               Text            =   "Text2"
               Top             =   1440
               Width           =   1635
            End
            Begin VB.TextBox txtvUnit 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1260
               TabIndex        =   25
               Text            =   "Text2"
               Top             =   1800
               Width           =   1635
            End
            Begin VB.TextBox txtIdProd 
               Height          =   285
               Left            =   1260
               TabIndex        =   17
               Text            =   "Text1"
               Top             =   240
               Width           =   915
            End
            Begin VB.CommandButton btoPesqProduto 
               Height          =   315
               Left            =   2160
               Picture         =   "formFaturamentoNFeEntrada.frx":005A
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   240
               Width           =   315
            End
            Begin VB.TextBox txtDescricao 
               Height          =   285
               Left            =   2640
               MaxLength       =   120
               TabIndex        =   19
               Text            =   "Text4"
               Top             =   240
               Width           =   5775
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               Caption         =   "Calculo do ICMS:"
               Height          =   195
               Left            =   4860
               TabIndex        =   73
               Top             =   2280
               Width           =   1635
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               Caption         =   "% ICMS:"
               Height          =   195
               Left            =   5760
               TabIndex        =   67
               Top             =   1485
               Width           =   735
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor Produto:"
               Height          =   195
               Left            =   60
               TabIndex        =   66
               Top             =   2220
               Width           =   1095
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor ICMS:"
               Height          =   195
               Left            =   5640
               TabIndex        =   65
               Top             =   1845
               Width           =   855
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               Caption         =   "BC ICMS:"
               Height          =   195
               Left            =   5760
               TabIndex        =   64
               Top             =   1125
               Width           =   735
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor IPI:"
               Height          =   195
               Left            =   2940
               TabIndex        =   63
               Top             =   1485
               Width           =   795
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               Caption         =   "% IPI:"
               Height          =   195
               Left            =   3120
               TabIndex        =   62
               Top             =   1125
               Width           =   615
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               Caption         =   "CFOP:"
               Height          =   195
               Left            =   5880
               TabIndex        =   61
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               Caption         =   "CST:"
               Height          =   195
               Left            =   3360
               TabIndex        =   60
               Top             =   720
               Width           =   375
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               Caption         =   "NCM:"
               Height          =   195
               Left            =   420
               TabIndex        =   59
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Unidade:"
               Height          =   195
               Left            =   480
               TabIndex        =   45
               Top             =   1125
               Width           =   675
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               Caption         =   "Quant.:"
               Height          =   195
               Left            =   600
               TabIndex        =   44
               Top             =   1485
               Width           =   555
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor Unit.:"
               Height          =   195
               Left            =   360
               TabIndex        =   43
               Top             =   1845
               Width           =   795
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Caption         =   "ID:"
               Height          =   195
               Left            =   540
               TabIndex        =   42
               Top             =   300
               Width           =   615
            End
         End
         Begin MSFlexGridLib.MSFlexGrid msfgProdutos 
            Height          =   1755
            Left            =   120
            TabIndex        =   41
            Top             =   180
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   3096
            _Version        =   393216
            Cols            =   15
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"formFaturamentoNFeEntrada.frx":03E4
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Deixe o mouse parado encima do item para saber qual a quantidade sera armazenada no estoque..."
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   1980
            Width           =   8715
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Duplicatas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3315
         Left            =   -74760
         TabIndex        =   37
         Top             =   480
         Width           =   11415
         Begin VB.ComboBox cboPlanoContas 
            Height          =   315
            Left            =   4740
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   2820
            Width           =   3795
         End
         Begin VB.ComboBox cboCentroCustos 
            Height          =   315
            Left            =   4740
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   1980
            Width           =   3795
         End
         Begin VB.ComboBox cboDocumento 
            Height          =   315
            Left            =   4740
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   2400
            Width           =   3795
         End
         Begin VB.CommandButton btoDuplRemover 
            Caption         =   "&Remover"
            Height          =   555
            Left            =   9420
            TabIndex        =   55
            Top             =   2520
            Width           =   1695
         End
         Begin VB.CommandButton btoDuplAdicionar 
            Caption         =   "&Adicionar"
            Height          =   555
            Left            =   9420
            TabIndex        =   54
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox txtvDupl 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   2820
            Width           =   1755
         End
         Begin MSComCtl2.DTPicker dtpdDupl 
            Height          =   315
            Left            =   1260
            TabIndex        =   50
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   50593793
            CurrentDate     =   40591
         End
         Begin VB.TextBox txtnDupl 
            Height          =   285
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   1980
            Width           =   1755
         End
         Begin MSFlexGridLib.MSFlexGrid msfgDupl 
            Height          =   1515
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   2672
            _Version        =   393216
            Cols            =   7
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"formFaturamentoNFeEntrada.frx":04AC
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            Caption         =   "Plano de Contas:"
            Height          =   195
            Left            =   3180
            TabIndex        =   115
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Centro de Custos:"
            Height          =   195
            Left            =   3420
            TabIndex        =   58
            Top             =   2040
            Width           =   1275
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Documento:"
            Height          =   195
            Left            =   3660
            TabIndex        =   57
            Top             =   2460
            Width           =   975
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor:"
            Height          =   195
            Left            =   660
            TabIndex        =   48
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Vencimento:"
            Height          =   195
            Left            =   180
            TabIndex        =   47
            Top             =   2460
            Width           =   975
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "Duplicata:"
            Height          =   195
            Left            =   360
            TabIndex        =   46
            Top             =   2040
            Width           =   795
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados da Nota Fiscal"
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
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   11955
      Begin VB.CheckBox chkMovFisco 
         Caption         =   "Movimentar FISCO"
         Height          =   195
         Left            =   9360
         TabIndex        =   77
         Top             =   180
         Width           =   2295
      End
      Begin VB.CheckBox chkMovFinanceiro 
         Caption         =   "Movimentar FINANCEIRO"
         Height          =   195
         Left            =   6780
         TabIndex        =   76
         Top             =   420
         Width           =   2235
      End
      Begin VB.CheckBox chkMovEstoque 
         Caption         =   "Movimentar ESTOQUE"
         Height          =   195
         Left            =   6780
         TabIndex        =   75
         Top             =   180
         Width           =   2175
      End
      Begin VB.CheckBox chkNFDevolucao 
         Caption         =   "Nota Fiscal de DEVOLUÇÃO"
         Height          =   195
         Left            =   9360
         TabIndex        =   74
         Top             =   420
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker dtpEmissao 
         Height          =   315
         Left            =   3420
         TabIndex        =   8
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   50593793
         CurrentDate     =   40591
      End
      Begin VB.TextBox txtnNF 
         Height          =   285
         Left            =   780
         MaxLength       =   9
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   315
         Width           =   1395
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   2700
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fornecedor"
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
      Left            =   60
      TabIndex        =   2
      Top             =   2880
      Width           =   11955
      Begin VB.ComboBox cboFornecedor 
         Height          =   315
         Left            =   3000
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   300
         Width           =   6375
      End
      Begin VB.TextBox txtDoc 
         Height          =   285
         Left            =   1080
         MaxLength       =   14
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ/CPF:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chave de Acesso da NF-e"
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
      Left            =   60
      TabIndex        =   1
      Top             =   1380
      Width           =   11955
      Begin VB.CheckBox chkretICMSST 
         Caption         =   "ICMS Retido por ST"
         Height          =   195
         Left            =   10080
         TabIndex        =   119
         Top             =   1020
         Width           =   1755
      End
      Begin VB.TextBox txtModelo 
         Height          =   285
         Left            =   3780
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Left            =   1980
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtnProt 
         Height          =   285
         Left            =   1980
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   975
         Width           =   3015
      End
      Begin VB.TextBox txtVersaoXML 
         Height          =   285
         Left            =   7140
         MaxLength       =   4
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   975
         Width           =   735
      End
      Begin VB.TextBox txtChaveAcesso 
         Height          =   285
         Left            =   1980
         MaxLength       =   44
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Modelo:"
         Height          =   195
         Left            =   2880
         TabIndex        =   72
         Top             =   285
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Serie:"
         Height          =   195
         Left            =   1500
         TabIndex        =   71
         Top             =   285
         Width           =   435
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "Versão da NFe:"
         Height          =   195
         Left            =   5700
         TabIndex        =   70
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "Protocolo de Autorização:"
         Height          =   195
         Left            =   60
         TabIndex        =   69
         Top             =   1020
         Width           =   1875
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Chave Acesso:"
         Height          =   195
         Left            =   840
         TabIndex        =   68
         Top             =   660
         Width           =   1095
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12120
      _ExtentX        =   21378
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Importar NF-e"
            ImageIndex      =   12
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
               Picture         =   "formFaturamentoNFeEntrada.frx":0534
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeEntrada.frx":0986
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeEntrada.frx":0CA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeEntrada.frx":1532
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeEntrada.frx":2784
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeEntrada.frx":305E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeEntrada.frx":38F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeEntrada.frx":4182
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeEntrada.frx":53D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeEntrada.frx":56EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeEntrada.frx":5A08
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeEntrada.frx":5DFF
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFaturamentoNFeEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 
 '##################################################################
 '### 19/03/2012
 '### Rever quando a empressa for do Simples Nacional e der credito
 '### de ICMS pois o campo no BD total_vCredICMSSN nao existe no
 '### manual, ate pq o cred. de ICMS é unico ou seja total_vICMS.
 '##################################################################
 
Dim IdReg           As Integer
Dim strTabela       As String
Dim lnProd          As Integer
Dim idProd          As Long
Dim idFornecedor    As Integer
Dim lnDupl          As Integer
Dim IdTransp        As Integer

Dim fileXMLOrigem   As String 'Armazena o local de origem do XML do Fornecedor

'
'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************
'Gerenciamento (ger)
Dim ger_Vendedor    As Integer
'cabecario do Pedido (ide)
Dim Versao          As String
Dim Id              As String
Dim ide_cUF         As String
Dim ide_cNF         As String
Dim ide_natOp       As String
Dim ide_indPag      As String
Dim ide_mod         As String
Dim ide_serie       As String
Dim ide_nNF         As String
Dim ide_dEmi        As String
Dim ide_dSaiEnt     As String
Dim ide_hSaiEnt     As String
Dim ide_tpNF        As String
Dim ide_cMunFG      As String
Dim ide_refNFe      As String
Dim ide_tpImp       As String
Dim ide_tpEmis      As String
Dim ide_cDV         As String
Dim ide_tpAmb       As String
Dim ide_finNFe      As String
Dim ide_procEmi     As String
Dim ide_verProc     As String
Dim nProt           As String
    'Emitente
Dim emit_id         As Integer
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
Dim dest            As String
Dim dest_pessoa     As String 'Variavel particular para saber que tipo de Pesoa F/J
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
Dim infAdic_infCpl  As String
    'Produtos
Dim aItem(1000)     As Variant
Dim aICMS(1000)     As Variant
Dim aIPI(1000)      As Variant
Dim aPIS(1000)      As Variant
Dim aCOFINS(1000)   As Variant
Dim aEstoque(1000)  As Variant 'Variavel gerenciado ra para estoque
Dim cItens          As Integer 'Contador dos itens
    'Fatura
Dim fat_nFat        As String
Dim fat_vOrig       As String
Dim fat_vDesc       As String
Dim fat_vLiq        As String
    
    'Cobranca
'Dim aFat(100)       As Variant
Dim aCob(100)       As Variant
Dim cCob            As Integer 'Contador das cobrancas

    
    
    'Transporte
Dim transp_modFrete     As String
Dim transp_CNPJ         As String
Dim transp_xNome        As String
Dim transp_IE           As String
Dim transp_xEnder       As String
Dim transp_xMun         As String
Dim transp_UF           As String
Dim transp_qVol         As String
Dim transp_esp          As String
Dim transp_marca        As String
Dim transp_nVol         As String
Dim transp_pesoL        As String
Dim transp_pesoB        As String
    'TOTAIS
Dim total_vBC           As String
Dim total_vICMS         As String
Dim total_vBCST         As String
Dim total_vICMSST       As String
Dim total_vCredICMSSN   As String
Dim total_vProd         As String
Dim total_vFrete        As String
Dim total_vSeg          As String
Dim total_vDesc         As String
Dim total_vIPI          As String
Dim total_vPIS          As String
Dim total_vCOFINS       As String
Dim total_vOutro        As String
Dim total_vNF           As String

Dim infCpl              As String
'***********************************************************************************************
'***********************************************************************************************
'***********************************************************************************************



Private Sub AtualizarCustos(iProduto As Integer, sCusto As String)
    '******************************************************************************************
    '*** Data: 08/07/2011
    '*** Obj.: Atualiza o preco de custo do material
    '******************************************************************************************
    Dim vReg(10)    As Variant
    Dim cReg        As Integer
    
    '##############################################################
    '### Caso a NFe nao mov. estoque, nao atualizar o preco de custo
    '##############################################################
    If chkMovEstoque.Value = False Then
        'ValidarProdutos = True
        Exit Sub
    End If
    '##############################################################
    cReg = 0
    vReg(cReg) = Array("Custo", sCusto, "S")
    If RegistroAlterar("EstoqueProduto", vReg, cReg, "id=" & iProduto) = False Then
        MsgBox "Erro ao localizar o produto no estoque para atualização do custo." & vbCrLf & "Custo não atualizado.", vbInformation, "Aviso"
    End If
    'Nao atuializa o preço de venda para uma atualização futura
End Sub

Private Sub CalcItem()
    txtvProd.Text = ConvMoeda(Val(ChkVal(txtQtd.Text, 0, cDecQtd)) * Val(ChkVal(txtvUnit.Text, 0, cDecMoeda)))
    
    txtvIPIp.Text = ConvMoeda(Val(ChkVal(txtpIPIp.Text, 0, cDecMoeda)) * Val(ChkVal(txtvProd.Text, 0, cDecMoeda)) / 100)
    
    If optBCICMS(0).Value = True Then 'Valor do produto
            txtBCICMSp.Text = ConvMoeda(ChkVal(txtvProd.Text, 0, cDecMoeda))
        ElseIf optBCICMS(1).Value = True Then 'Valor do Produto + IPI
            txtBCICMSp.Text = ConvMoeda(Val(ChkVal(txtvProd.Text, 0, cDecMoeda)) + Val(ChkVal(txtvIPIp.Text, 0, cDecMoeda)))
        Else
            txtBCICMSp.Text = "0.00"
    End If
    
    txtvICMSp.Text = ConvMoeda(Val(ChkVal(txtpICMSp.Text, 0, cDecMoeda)) * Val(ChkVal(txtBCICMSp.Text, 0, cDecMoeda)) / 100)
    
End Sub

Private Sub ExcluirNFe()
    Dim i As Integer
    
    If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If

    
    If IdReg = 0 Then
            MsgBox "Selecione um registro!", vbInformation, "Aviso"
            Exit Sub
        Else
            If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                                vbCrLf & _
                                "ID: " & IdReg & vbCrLf & _
                                "Nota Fiscal: " & txtnNF.Text & vbCrLf & _
                                "Chave NF-e: " & txtChaveAcesso.Text & vbCrLf & _
                                "Fornecedor: " & cboFornecedor.Text, vbYesNo + vbQuestion) = vbYes Then
                        RegistroExcluir strTabela, "Id = " & IdReg
                        RegistroExcluir strTabela & "Itens", "IdReg = " & IdReg
                        RegistroExcluir strTabela & "Cobranca", "IdReg = " & IdReg
                        If Trim(Id) = "" Then
                                RegistroExcluir "FinanceiroContasPRCadastro", "Ide_NFe = '" & IdReg & "'"
                            Else
                                RegistroExcluir "FinanceiroContasPRCadastro", "Ide_NFe = '" & Id & "'"
                        End If
                        'Baixar Estoque
                        '##############################################################
                        '### Caso a NFe nao mov. estoque, nao atualizar o preco de custo
                        '##############################################################
                        If chkMovEstoque.Value = 1 Then
                            For i = 1 To cItens
                                '****** Movimenta o estoque **************************************
                                ide_tpNF = 0
                                If MovimentarEstoque("s", _
                                                    CLng(aItem(i)(0)), _
                                                    CDate(ide_dEmi), _
                                                    ide_nNF, _
                                                    CStr(aEstoque(i)(1)), _
                                                    CStr(aEstoque(i)(2)), _
                                                    CStr(aItem(i)(10)), _
                                                    "Exclusao da NFe em ( " & Now() & ") por: " & PgDadosUsuario(ID_Usuario).Nome, _
                                                    emit_xNome, Id, emit_id, emit_CNPJ) = False Then
                                                    MsgBox "Erro ao Movimentar Estoque com o item n. " & i
                                End If
                            Next
                        
                        End If
                        
                        RemoverNFe
                        LimpaFormulario Me
                        IdReg = 0
                        sstNF.Tab = 0
                        MsgBox "Nota Fiscal Excluida com sucesso!", vbInformation, "Aviso"
                        LimpaFormulario Me
            End If
    End If
End Sub

Private Sub LimpFormProd()
    txtIdProd.Text = ""
    txtDescricao.Text = ""
    
    txtNCM.Text = ""
    txtCST.Text = ""
    txtCFOP.Text = ""
    
    cboUnidade.Clear
    txtQtd.Text = ""
    txtvUnit.Text = ""
    txtvProd.Text = ""
    
    txtpIPIp.Text = ""
    txtvIPIp.Text = ""
    
    txtBCICMSp.Text = ""
    txtpICMSp.Text = ""
    txtvICMSp.Text = ""
    
    
    
    
End Sub

Private Sub MontarArrayFornecedor(idFor As Integer)
    emit_CNPJ = PgDadosFornecedor(idFor).Doc
    emit_xNome = PgDadosFornecedor(idFor).Nome
    emit_xFant = PgDadosFornecedor(idFor).Fant
    emit_xLgr = PgDadosFornecedor(idFor).Lgr
    emit_nro = PgDadosFornecedor(idFor).Nro
    emit_xCpl = PgDadosFornecedor(idFor).Cpl
    emit_Bairro = PgDadosFornecedor(idFor).Bairro
    
    'ide_cUF
    'ide_cMunFG

    'emit_cMun= PgDadosFornecedor(idFor).Mun
    emit_xMun = PgDadosFornecedor(idFor).Mun
    emit_UF = PgDadosFornecedor(idFor).uf
    emit_CEP = PgDadosFornecedor(idFor).CEP
    'emit_cPais= PgDadosFornecedor(idFor).
    'emit_xPais= PgDadosFornecedor(idFor)
    emit_fone = PgDadosFornecedor(idFor).Fone
    emit_IE = PgDadosFornecedor(idFor).IE
    emit_IEST = PgDadosFornecedor(idFor).iest
    emit_IM = PgDadosFornecedor(idFor).im
    emit_CNAE = PgDadosFornecedor(idFor).cnae
    'emit_CRT= PgDadosFornecedor(idFor)
End Sub

Private Sub MontarArrayIde()
    Id = txtChaveAcesso.Text
    Versao = txtVersaoXML.Text
    ide_cNF = ""
    ide_natOp = ""
    ide_indPag = ""
    ide_mod = txtModelo.Text
    ide_serie = txtSerie.Text
    ide_nNF = txtnNF.Text
    ide_dEmi = dtpEmissao.Value
    ide_dSaiEnt = ""
    ide_hSaiEnt = ""
    ide_tpNF = "1"
    ide_refNFe = ""
    ide_tpImp = ""
    ide_tpEmis = ""
    ide_cDV = ""
    ide_tpAmb = ""
    ide_finNFe = ""
    ide_procEmi = ""
    ide_verProc = ""
End Sub

Private Sub MostrarDadosForm()
    Dim i As Integer
    
    'IDENTIFICACAO DA NFe
    txtChaveAcesso.Text = Id
    txtVersaoXML.Text = Versao
    txtModelo.Text = ide_mod
    txtSerie.Text = ide_serie
    txtnNF.Text = ide_nNF
    dtpEmissao.Value = Format(ide_dEmi, "DD/MM/YYYY")
    txtnProt.Text = nProt
    'chkretICMSST.Value = IIf(Trim(emit_IEST) = "", 0, 1)
    
    'txtChaveAcesso.Enabled = False
    'txtVersaoXML.Enabled = False
    'txtModelo.Enabled = False
    'txtSerie.Enabled = False
    'txtnNF.Enabled = False
    'dtpEmissao.Enabled = False
    'txtnProt.Enabled = False
    
    'DADOS DO EMISSOR
    txtDoc.Text = emit_CNPJ
    cboFornecedor.Text = emit_xNome
    
    'txtDoc.Enabled = False
    'cboFornecedor.Enabled = False
    
    'PRODUTOS
    'id_intProd|det_cProd|det_cEAN|det_xProd|EXTIPI|det_NCM|det_CFOP|det_uCom|
    'det_qCom|det_vUnCom|det_vProd|det_cEANTrib|det_uTrib|det_qTrib|
    'det_vUnTrib|
    'det_vFrete|det_vSeg|det_vDesc|det_vOutro|det_indTot|xPed|nItemPed
    'det_indTot = 0 = O valor do item compoe a NF / 1  = O valor do item nao compoe a NF
    total_vCredICMSSN = 0
    msfgProdutos.Rows = 1
    For i = 1 To cItens
        With msfgProdutos
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = aItem(i)(0)
            .TextMatrix(i, 1) = aItem(i)(1)
            .TextMatrix(i, 2) = aItem(i)(3)
            .TextMatrix(i, 3) = aItem(i)(5)
            .TextMatrix(i, 4) = aICMS(i)(1)
            .TextMatrix(i, 5) = aItem(i)(6)
            .TextMatrix(i, 6) = aItem(i)(7)
            .TextMatrix(i, 7) = aItem(i)(8)
            .TextMatrix(i, 8) = ChkVal(CStr(aItem(i)(9)), 0, cDecMoeda)
            .TextMatrix(i, 9) = aItem(i)(10)
            .TextMatrix(i, 10) = aICMS(i)(4)
            .TextMatrix(i, 11) = aICMS(i)(6)
            .TextMatrix(i, 12) = aIPI(i)(4)
            .TextMatrix(i, 13) = aIPI(i)(3)
            .TextMatrix(i, 14) = aICMS(i)(5)
            total_vCredICMSSN = Val(ChkVal(total_vCredICMSSN, 0, cDecMoeda)) + Val(ChkVal(CStr(aICMS(i)(15)), 0, cDecMoeda))
            total_vCredICMSSN = ChkVal(total_vCredICMSSN, 0, cDecMoeda)
        End With
    Next
    
    
    'COBRANCA
    'IdNFe|nfat|vOrig|vDesc|vLiq|nDup|dVenc|vDup|CC|tpDoc|PlanoContas
     msfgDupl.Rows = 1
    For i = 1 To cCob
        With msfgDupl
            .Rows = .Rows + 1
            .TextMatrix(i, 1) = aCob(i)(5)
            .TextMatrix(i, 2) = aCob(i)(6)
            .TextMatrix(i, 3) = aCob(i)(7)
            If Trim(aCob(i)(8)) <> "0" Then
                        .TextMatrix(i, 4) = Left("000", 3 - Len(Trim(aCob(i)(8)))) & aCob(i)(8) & " - " & pgDadosCentroCustos(CInt(aCob(i)(8))).Descricao
                    Else
                        .TextMatrix(i, 4) = ""
            End If
            If Trim(aCob(i)(9)) <> "0" Then
                    .TextMatrix(i, 5) = Left("000", 3 - Len(Trim(aCob(i)(9)))) & aCob(i)(9) & " - " & pgDadosTipoDocumento(CInt(aCob(i)(9))).Descricao
                Else
                    .TextMatrix(i, 5) = ""
            End If
            'Plano de Contas
            If Trim(aCob(i)(10)) <> "0" Then
                        .TextMatrix(i, 6) = ZE(CInt(aCob(i)(10)), 3) & " - " & PgDadosPlanoContas("ID", CInt(aCob(i)(10))).Descricao
                    Else
                        .TextMatrix(i, 6) = ""
            End If
            
        End With
    Next
    'TRANSPORTADORA
    txtTranspCNPJ.Text = transp_CNPJ
    cboTranspNome.Text = transp_xNome
    txtFreteConta.Text = transp_modFrete
    
    'TOTAL DA NOTA FISCAL
    txtBCICMS.Text = total_vBC
    txtvICMS.Text = total_vICMS
    txtBCICMSST.Text = total_vBCST
    txtvICMSST.Text = total_vICMSST
    txtvCredICMSSN.Text = total_vCredICMSSN
    txtvFrete.Text = total_vFrete
    txtvSeguro.Text = total_vSeg
    txtvDesconto.Text = total_vDesc
    txtvOutras.Text = total_vOutro
    txtvProduto.Text = total_vProd
    txtvIPI.Text = total_vIPI
    txtvTotalNF.Text = total_vNF
    txtvCOFINS.Text = total_vCOFINS
    txtvPIS.Text = total_vPIS
    txtObs.Text = infCpl
    
    BloqFormNFe
    'btoRemover.Enabled = False
    'If Trim(txtModelo.Text) = "55" Then
        If PgDadosConfig.EntradaNFSemAutorSEFAZ = 0 And Trim(txtnProt.Text) = "" Then
            MsgBox "Nota Fiscal sem autorização da SEFAZ!", vbInformation, "Aviso"
            HDMenu Me, True
        End If
    'End If
End Sub

Private Function nfeCadastrada() As Boolean
    ' 12.12.2012
    'Checa se a nfe foi ja inserida
    nfeCadastrada = False
    Dim sSQL As String
    Dim Rst As Recordset
    sSQL = "SELECT * FROM FaturamentoNFeEntrada WHERE id_empresa=" & ID_Empresa & " AND ide_nNF = '" & Trim(txtnNF.Text) & "' AND emit_CNPJ = '" & txtDoc.Text & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            nfeCadastrada = False
        Else
            nfeCadastrada = True
    End If
    Rst.Close
    If dest_CNPJ <> PgDadosEmpresa(ID_Empresa).CNPJ Then
        MsgBox "O CNPJ (" & dest_CNPJ & ") do destinatário da NFe difere do CNPJ (" & PgDadosEmpresa(ID_Empresa).CNPJ & ")da empresa.", vbInformation, "Aviso!"
    End If
  
End Function

Private Function pgTagXML(tagI As String, tagF As String, sDoc As String) As String
    Dim str As String
    If InStr(sDoc, tagI) = 0 Then
        pgTagXML = ""
        Exit Function
    End If
    str = Mid(sDoc, InStr(sDoc, tagI) + Len(tagI), Len(sDoc))
    'str = Mid(str, 1, InStrRev(str, tagF))
    str = Left(str, InStr(str, tagF) - 1)
    pgTagXML = str
End Function

Private Sub btoAdicionar_Click()
    Dim p As Integer
    With msfgProdutos
        If idProd = 0 Then
            MsgBox "Selecione um produto que esteja cadastrado na base de dados!", vbInformation, "Aviso"
            Exit Sub
        End If
        'If Trim(txtChaveAcesso.Text) <> "" And lnProd = 0 Then
        If Trim(fileXMLOrigem) <> "" And lnProd = 0 Then
        'If Trim(fileXMLOrigem) <> "" Then
            MsgBox "Não é permitido inclusao de item com a importação da NFe." & vbCrLf & _
                   "Favor selecionar um item!", vbInformation, "Aviso"
            Exit Sub
        End If
        If lnProd = 0 Then
            .Rows = .Rows + 1
            lnProd = .Rows - 1
        End If
        
        'Registra as variaveis caso nao seja importacao do XML
        If txtChaveAcesso.Enabled = True Then
            MontarArrayBaseEstoque (idProd)
        End If
        aItem(lnProd)(0) = idProd
        pgDadosEstoque
        If pgDadosEstoqueProduto(idProd).NCM <> aItem(lnProd)(5) Then
            If MsgBox("O NCM constante na Nota Fiscal diverge do NCM constante no item do estoque. Deseja continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                Exit Sub
            End If
        End If
        
        
        .TextMatrix(lnProd, 0) = idProd
        .TextMatrix(lnProd, 2) = txtDescricao.Text
        .TextMatrix(lnProd, 3) = txtNCM.Text
        .TextMatrix(lnProd, 4) = txtCST.Text
        .TextMatrix(lnProd, 5) = txtCFOP.Text
        .TextMatrix(lnProd, 6) = cboUnidade.Text
        
        .TextMatrix(lnProd, 7) = txtQtd.Text
        .TextMatrix(lnProd, 8) = txtvUnit.Text
        .TextMatrix(lnProd, 9) = txtvProd.Text
        
        .TextMatrix(lnProd, 13) = txtpIPIp.Text
        .TextMatrix(lnProd, 12) = txtvIPIp.Text
        
        .TextMatrix(lnProd, 10) = txtBCICMSp.Text
        .TextMatrix(lnProd, 14) = txtpICMSp.Text
        .TextMatrix(lnProd, 11) = txtvICMSp.Text
        '.TextMatrix(lnProd, 0) = txtIdProd.Text
        '.TextMatrix(lnProd, 1) = pgDadosEstoqueProduto(idProd).Referencia
        '.TextMatrix(lnProd, 2) = pgDadosEstoqueProduto(idProd).Descricao
        '.TextMatrix(lnProd, 3) = pgDadosEstoqueProduto(idProd).NCM
        '.TextMatrix(lnProd, 4) = pgDadosEstoqueProduto(idProd).ICMSCST
        '.TextMatrix(lnProd, 5) = cboUnidade.Text
        '.TextMatrix(lnProd, 6) = ChkVal(txtQtd.Text, 0, cDecQtd)
        '.TextMatrix(lnProd, 7) = ConvMoeda(txtvUnit.Text)
        '.TextMatrix(lnProd, 8) = ConvMoeda(Val(ChkVal(txtQtd.Text, 0, cDecQtd)) * Val(ChkVal(txtvUnit.Text, 0, cDecMoeda)))
        
        'aItem(lnProd)(8) = cboUnidade.Text
        'aItem(lnProd)(13) = cboUnidade.Text
        
        'aItem(lnProd)(9) = ConvMoeda(txtvUnit.Text)
        'aItem(lnProd)(14) = ConvMoeda(txtvUnit.Text)
        
        'aItem(lnProd)(10) = .TextMatrix(lnProd, 8)
        lnProd = 0
        idProd = 0
        
    End With
    
   LimpFormProd
    
    calcTotais
End Sub
Private Sub btoDuplAdicionar_Click()
    With msfgDupl
        If Trim(txtnDupl.Text) = "" Then Exit Sub
        If lnDupl = 0 Then
            .Rows = .Rows + 1
            lnDupl = .Rows - 1
        End If
        
        
        '
        .TextMatrix(lnDupl, 1) = txtnDupl.Text
        .TextMatrix(lnDupl, 2) = dtpdDupl.Value
        .TextMatrix(lnDupl, 3) = ConvMoeda(txtvDupl.Text)
        .TextMatrix(lnDupl, 4) = cboCentroCustos.Text
        .TextMatrix(lnDupl, 5) = cboDocumento.Text
        .TextMatrix(lnDupl, 6) = cboPlanoContas.Text
        If cCob = 0 Then
            cCob = lnDupl
            '************************************************************
        
        'vReg(cReg) = Array("cobr_nFat", aCob(i)(0), "S"): cReg = cReg + 1
        'vReg(cReg) = Array("cobr_vOrig", aCob(i)(1), "S"): cReg = cReg + 1
        'vReg(cReg) = Array("cobr_vDesc", aCob(i)(2), "S"): cReg = cReg + 1
        'vReg(cReg) = Array("cobr_vLiq", aCob(i)(3), "S"): cReg = cReg + 1
        
        '
            '************************************************************
            
                               'IdNFe|nfat|vOrig|vDesc|vLiq|nDup|dVenc|vDup|CC|tpDoc|PlanoContas
            aCob(cCob) = Array(ide_nNF, total_vNF, "0.00", total_vNF, total_vNF) ', ide_nNF, dtpdDupl.Value, txtvDupl.Text)
        End If
    End With
    lnDupl = 0
    txtnDupl.Text = ""
    txtvDupl.Text = ""
    cboCentroCustos.Clear
    cboDocumento.Clear
    cboPlanoContas.Clear
End Sub

Private Sub btoDuplRemover_Click()
    If lnDupl = 0 Then Exit Sub
    If lnDupl = 1 And msfgDupl.Rows = 2 Then
            msfgDupl.Rows = 1
        Else
        msfgDupl.RemoveItem lnDupl
    End If
    lnDupl = 0
End Sub

Private Sub btoPesqProduto_Click()
    idProd = 0
    PesquisarProduto
End Sub

Private Sub PesquisarProduto()
    'Dim Rst         As Recordset
    'Dim sSQL        As String
    If idProd = 0 Then
        If Trim(txtDescricao.Text) = "" Then
                idProd = formBuscar.IniciarBusca("EstoqueProduto")
            Else
                idProd = formBuscar.IniciarBusca("EstoqueProduto", , "Descricao", txtDescricao.Text)
        End If
    End If
    
    If Trim(idProd) = 0 Then Exit Sub
    
   
    txtIdProd.Text = pgDadosEstoqueProduto(idProd).Id
    txtDescricao.Text = pgDadosEstoqueProduto(idProd).Descricao
    
    'Nao termina de preecher a tela pois tratase de NFe
    If txtChaveAcesso.Enabled = False Then Exit Sub

    txtNCM.Text = pgDadosEstoqueProduto(idProd).NCM
    txtCST.Text = pgDadosEstoqueProduto(idProd).ICMSCST
    cboUnidade.Clear
    cboUnidade.AddItem pgDadosEstoqueProduto(idProd).Unidade
    cboUnidade.Text = cboUnidade.List(0)
    txtpIPIp.Text = pgDadosEstoqueProduto(idProd).IPIAliquota
End Sub

Private Sub btoRemover_Click()
    If lnProd = 0 Then Exit Sub
    If MsgBox("Deseja realmente remover este item?", vbYesNo, "Removendo Item do Pedido") = vbYes Then
        If msfgProdutos.Rows = 2 Then
                msfgProdutos.Rows = 1
            Else
                msfgProdutos.RemoveItem lnProd
                lnProd = 0
                txtIdProd.Text = ""
                txtDescricao.Text = ""
                cboUnidade.Clear
                txtQtd.Text = ""
                txtvUnit.Text = ""
        End If
    End If
End Sub

Private Sub MontarBaseDeDados()
    Dim vDados(1000)    As Variant
    Dim contReg         As Integer
    Dim i               As Integer
    
    contReg = 0
    
    vDados(contReg) = Array("MovEstoque", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("MovFinanceiro", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("NFDevolucao", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("MovFisco", "1", "N"): contReg = contReg + 1
    
    vDados(contReg) = Array("retICMSST", "1", "N"): contReg = contReg + 1
    
    vDados(contReg) = Array("nProt", "50", "S"): contReg = contReg + 1
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
    vDados(contReg) = Array("ide_dSaiEnt", "15", "D"): contReg = contReg + 1
    vDados(contReg) = Array("ide_hSaiEnt", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_tpNF", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_cMunFG", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_refNFe", "50", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ide_tpImp", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_tpEmis", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_cDV", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_tpAmb", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_finNFe", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_procEmi", "5", "N"): contReg = contReg + 1
    vDados(contReg) = Array("ide_verProc", "20", "S"): contReg = contReg + 1
    
    'Emitente
    vDados(contReg) = Array("emit_id", "10", "S"): contReg = contReg + 1
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
    'Transporte
    vDados(contReg) = Array("transp_modFrete", "5", "N"): contReg = contReg + 1
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
        
    'TOTAIS
    vDados(contReg) = Array("total_vBC", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vICMS", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("total_vBCST", "15", "S"): contReg = contReg + 1
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
    vDados(contReg) = Array("ger_Vendedor", "15", "N") ': contReg = contReg + 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg
    
    'Produto******************************************************************************
    contReg = 0
    vDados(contReg) = Array("IdReg", "60", "N"): contReg = contReg + 1
    vDados(contReg) = Array("IdNFe", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_IdProduto", "60", "N"): contReg = contReg + 1
    vDados(contReg) = Array("det_cProd", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_cEAN", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_xProd", "120", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_NCM", "8", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_EXTIPI", "8", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_CFOP", "4", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_uCom", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_qCom", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_vUnCom", "21", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_vProd", "15", "S"): contReg = contReg + 1
    
    vDados(contReg) = Array("det_cEANTrib", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_uTrib", "10", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_qTrib", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("det_vUnTrib", "21", "S"): contReg = contReg + 1
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
    vDados(contReg) = Array("Estoque_Unid", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Estoque_Qtd", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Estoque_vUnit", "15", "S"): contReg = contReg + 1
   
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "Itens"
    
    
    contReg = 0
    'COBRANCA
    vDados(contReg) = Array("IdReg", "60", "N"): contReg = contReg + 1
    vDados(contReg) = Array("IdNFe", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_nFat", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_vOrig", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_vDesc", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_vLiq", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_nDup", "60", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_dVenc", "10", "D"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_vDup", "15", "S"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_tpDoc", "15", "N"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_PlanoContas", "15", "N"): contReg = contReg + 1
    vDados(contReg) = Array("cobr_CC", "15", "N") ': contReg = contReg + 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "Cobranca"
End Sub


Private Sub btoCadastrarItem_Click()
    If Trim(txtIdProd.Text) = "" Or lnProd = 0 Then
        MsgBox "Selecione um produto da grade!", vbInformation, "Aviso"
        Exit Sub
    End If
    With msfgProdutos
    formEstoqueProduto.RecebendoDadosProduto .TextMatrix(lnProd, 1), _
                                             .TextMatrix(lnProd, 2), _
                                             "", _
                                             .TextMatrix(lnProd, 3), _
                                             CStr(aICMS(lnProd)(8)), _
                                             CStr(aICMS(lnProd)(0)), _
                                             CStr(aICMS(lnProd)(1)), _
                                             CStr(aIPI(lnProd)(3)), _
                                             CStr(aIPI(lnProd)(1)), _
                                             "0", _
                                             .TextMatrix(lnProd, 6), _
                                             .TextMatrix(lnProd, 8)
                                             
                                             
                                                
    End With
End Sub

Private Sub cboCentroCustos_DropDown()
    Dim Rst As Recordset
    cboCentroCustos.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCentroCustos WHERE ID_Empresa = " & ID_Empresa)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCentroCustos.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If

End Sub
Private Sub cboDocumento_DropDown()
    Dim Rst As Recordset
    cboDocumento.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento WHERE ID_Empresa = " & ID_Empresa)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboDocumento.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub cboFornecedor_Click()
    If Trim(cboFornecedor.Text) = "" Then Exit Sub
    idFornecedor = Left(Trim(cboFornecedor.Text), 6)
    PesquisarForn
    'cboFornecedor.Text = PgDadosFornecedor(idFornecedor).Nome
    'txtDoc.Text = PgDadosFornecedor(idFornecedor).Doc
End Sub

Private Sub cboFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        idFornecedor = 0
        PesquisarForn
    End If
End Sub



Private Sub BloqFormNFe()
    HDForm Me, False
    
    chkMovEstoque.Enabled = True
    chkNFDevolucao.Enabled = True
    chkMovFinanceiro.Enabled = True
    
    
    msfgProdutos.Enabled = True
    txtIdProd.Enabled = True
    btoPesqProduto.Enabled = True
    
    btoAdicionar.Enabled = True
    btoCadastrarItem.Enabled = True
    
    msfgDupl.Enabled = True
    txtvCredICMSSN.Enabled = True
    
    chkretICMSST.Enabled = True
End Sub



Private Sub cboPlanoContas_DropDown()
    Dim Rst As Recordset
    cboPlanoContas.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroPlanoContas WHERE ID_Empresa = " & ID_Empresa & " ORDER BY Codigo")
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

Private Sub cboTranspNome_Click()
    If Trim(cboTranspNome.Text) = "" Then Exit Sub
    IdTransp = Left(cboTranspNome.Text, 6)
    cboTranspNome.Text = pgDadosTransportadora(IdTransp).Nome
    txtTranspCNPJ.Text = pgDadosTransportadora(IdTransp).CNPJ
End Sub

Private Sub cboTranspNome_DropDown()

    Dim Rst As Recordset
    
    Set Rst = RegistroBuscar("SELECT * FROM Transportadoras WHERE xNome LIKE '" & cboTranspNome.Text & "%'")
    If Rst.BOF And Rst.EOF Then
            cboTranspNome.Clear
            Exit Sub
        Else
            cboTranspNome.Clear
            Rst.MoveFirst
            Do Until Rst.EOF
                cboTranspNome.AddItem Left(String(6, "0"), 6 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & _
                " - " & _
                Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If



End Sub

Private Sub cboTranspNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarTransp
    End If
End Sub




Private Sub chkretICMSST_Click()
    
    txtBCICMSST.Enabled = IIf(chkretICMSST.Value = 1, True, False)
    txtvICMSST.Enabled = IIf(chkretICMSST.Value = 1, True, False)
    
End Sub

Private Sub chkTotaisAutomatico_Click()
    calcTotais
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
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    LimpaFormulario Me
    HDMenu Me, True
    HDForm Me, False
    msfgDupl.Rows = 1
    sstNF.Tab = 0
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub cboFornecedor_DropDown()
    Dim Rst As Recordset
    
    Set Rst = RegistroBuscar("SELECT * FROM Fornecedores WHERE ID_Empresa = " & ID_Empresa & " AND xNome LIKE '" & cboFornecedor.Text & "%'")
    If Rst.BOF And Rst.EOF Then
            cboFornecedor.Clear
            Exit Sub
        Else
            cboFornecedor.Clear
            Rst.MoveFirst
            Do Until Rst.EOF
                cboFornecedor.AddItem Left(String(6, "0"), 6 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & _
                " - " & _
                Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If

End Sub
Private Sub PesquisarForn(Optional CNPJ As String)
    Dim sSQL    As String
    Dim Rst     As Recordset
    'emit_id = 0
        
    If idFornecedor <> 0 Then
        CNPJ = PgDadosFornecedor(idFornecedor).Doc
    End If
        
    If Trim(CNPJ) <> "" Then
        sSQL = "SELECT * FROM Fornecedores WHERE ID_Empresa = " & ID_Empresa & " AND Doc = '" & CNPJ & "'"
       
        Set Rst = RegistroBuscar(sSQL)
        If Rst.BOF And Rst.EOF Then
                idFornecedor = 0
            Else
                idFornecedor = Rst.Fields("id")
        End If
        Rst.Close
    End If
    If Trim(idFornecedor) = 0 Then
            idFornecedor = formBuscar.IniciarBusca("Fornecedores") ', "xNome,xlgr,nro,xcpl,xbairro,xmun,uf,fone")
    End If
    If Trim(idFornecedor) = 0 Then Exit Sub
    
    cboFornecedor.Clear
    cboFornecedor.Text = PgDadosFornecedor(idFornecedor).Nome
    txtDoc.Text = PgDadosFornecedor(idFornecedor).Doc
    emit_id = idFornecedor
    'Registra as variaveis caso nao seja importacao do XML
    If txtChaveAcesso.Enabled = True Then
        MontarArrayFornecedor (idFornecedor)
    End If
End Sub

Private Sub msfgDupl_Click()
    With msfgDupl
        If Trim(.TextMatrix(.Row, 1)) = "" Then Exit Sub
        lnDupl = .Row
        txtnDupl.Text = .TextMatrix(.Row, 1)
        dtpdDupl.Value = .TextMatrix(.Row, 2)
        txtvDupl.Text = ChkVal(.TextMatrix(.Row, 3), 0, cDecMoeda)
        cboCentroCustos.Clear
        cboCentroCustos.AddItem IIf(Trim(.TextMatrix(.Row, 4)) = "", " ", .TextMatrix(.Row, 4))
        cboCentroCustos.Text = cboCentroCustos.List(0)
        
        cboDocumento.Clear
        cboDocumento.AddItem IIf(Trim(.TextMatrix(.Row, 5)) = "", " ", .TextMatrix(.Row, 5))
        cboDocumento.Text = cboDocumento.List(0)
    End With
End Sub

Private Sub msfgProdutos_DblClick()
   With msfgProdutos
        If .RowSel = 0 Then Exit Sub
        lnProd = .Row
        idProd = IIf(Trim(.TextMatrix(lnProd, 0)) = "", 0, .TextMatrix(lnProd, 0))
        
        
        txtIdProd.Text = idProd
        txtDescricao.Text = .TextMatrix(lnProd, 2)
        txtNCM.Text = .TextMatrix(lnProd, 3)
        txtCST.Text = .TextMatrix(lnProd, 4)
        txtCFOP.Text = .TextMatrix(lnProd, 5)
        cboUnidade.Clear
        cboUnidade.AddItem IIf(Trim(.TextMatrix(lnProd, 6)) = "", " ", .TextMatrix(lnProd, 6))
        cboUnidade.Text = cboUnidade.List(0)
        txtQtd.Text = .TextMatrix(lnProd, 7)
        txtvUnit.Text = .TextMatrix(lnProd, 8)
        txtvProd.Text = .TextMatrix(lnProd, 9)
        
        txtpIPIp.Text = .TextMatrix(lnProd, 13)
        txtvIPIp.Text = .TextMatrix(lnProd, 12)
        
        txtBCICMSp.Text = .TextMatrix(lnProd, 10)
        txtpICMSp.Text = .TextMatrix(lnProd, 14)
        txtvICMSp.Text = .TextMatrix(lnProd, 11)
        
        
    End With
End Sub


Private Sub msfgProdutos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    On Error Resume Next
'    Dim i As Integer
'    With msfgProdutos
'        If .Rows = 1 Then Exit Sub
'        i = IIf(.MouseRow = 0, 1, .MouseRow)
'        .ToolTipText = "item: " & Left("000", 3 - Len(i)) & i & " - [Unid.: " & aEstoque(i)(0) & "] [Qtd.: " & ChkVal(CStr(aEstoque(i)(1)), 0, cDecQtd) & "] [Valor Unitario: " & aEstoque(i)(2) & "]"
'    End With
'--------------------------------------------------
    On Error Resume Next
    Dim i As Integer
    Dim ii As Integer
    With msfgProdutos
    
        If .Rows = 1 Then Exit Sub
        If Trim(.TextMatrix(1, 0)) = "" Then Exit Sub

        i = IIf(.MouseRow = 0, 1, .MouseRow)
        If .MouseCol <= 0 Then
                ii = IIf(Trim(.TextMatrix(i, 0)) = "", 0, Trim(.TextMatrix(i, 0)))
                .ToolTipText = pgDescricaoMaterial(ii) '.TextMatrix(.MouseRow, .MouseCol)
            Else
                .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        End If
    End With

End Sub
Private Function pgDescricaoMaterial(i As Integer) As String
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim sTexto  As String
    sTexto = ""
    sSQL = "SELECT * FROM estoqueproduto WHERE id=" & i
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            sTexto = ""
        Else
            Rst.MoveFirst
            sTexto = "Estoque: " & UCase(Rst.Fields("Descricao")) & "  " & _
                                "Unid: " & UCase(Rst.Fields("unidade"))
    End If
    Rst.Close
    pgDescricaoMaterial = sTexto
End Function
Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    IdReg = 0
    dtpEmissao.Value = Date
    dtpdDupl.Value = Date
    HDMenu Me, False
    HDForm Me, True
    LimpaFormulario Me
    chkMovEstoque.Value = 1
    chkMovFinanceiro.Value = 1
    chkMovFisco.Value = 1
    msfgProdutos.Rows = 1
    msfgDupl.Rows = 1
    btoRemover.Enabled = True
End Sub

Private Sub optBCICMS_Click(Index As Integer)
    CalcItem
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
            Incluir
        Case "Excluir"
            ExcluirNFe
            
        Case "Imprimir"
            ImprimirDANFeEntrada
        Case "Pesquisar"
            IdReg = 0
            PesquisarRegistro
            
        Case "Salvar"
            If ValidarProdutos = False Then Exit Sub
            If grvRegistro = True Then
                HDMenu Me, True
                HDForm Me, False
                ArmazenarNFe
            End If
            
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            'LimpForm
            'txtID.Enabled = True
        Case "Importar NF-e"
            ImportarNFe
        Case "Manutenção da Tabela"
            'formManutencaoTabelas.IniciarManutencao Me
            MontarBaseDeDados
    End Select
End Sub
Private Sub ImprimirDANFeEntrada()
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If

    If Trim(Id) = "" Then Exit Sub
    ImprimirDANFEFornecedor (Id)
End Sub
Private Sub PesquisarRegistro()
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim NFe         As String
    
    If IdReg = 0 Then
        IdReg = formBuscar.IniciarBusca("FaturamentoNFeEntrada")
    End If
    If IdReg = 0 Then Exit Sub
    
    sSQL = "SELECT * FROM FaturamentoNFeEntrada WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum Registro encontrado."
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            IdReg = Rst.Fields("ID")
    End If
    
    '***************************************************************
    '*** Data: 25/07/2011
    '*** Obj.: carrecgar os chk para inf do sistema
    chkMovEstoque.Value = IIf(IsNull(Rst.Fields("MovEstoque")), 0, Rst.Fields("MovEstoque"))
    chkMovFinanceiro.Value = IIf(IsNull(Rst.Fields("MovFinanceiro")), 0, Rst.Fields("MovFinanceiro"))
    chkNFDevolucao.Value = IIf(IsNull(Rst.Fields("NFDevolucao")), 0, Rst.Fields("NFDevolucao"))
    chkMovFisco.Value = IIf(IsNull(Rst.Fields("MovFisco")), 0, Rst.Fields("MovFisco"))
    chkretICMSST.Value = IIf(IsNull(Rst.Fields("retICMSST")), 0, Rst.Fields("retICMSST"))
    
    '***************************************************************
    
      
  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>IDENTIFICADOR DA NOTA
    NFe = IIf(IsNull(Rst.Fields("idNFe")), "", Rst.Fields("idNFe"))
    Id = NFe 'Replace(Mid(NFe, InStr(NFe, "NFe"), 47), "NFe", "")
   
    Versao = IIf(IsNull(Rst.Fields("versao")), "", Rst.Fields("versao"))
    nProt = IIf(IsNull(Rst.Fields("nProt")), "", Rst.Fields("nProt"))
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> IDE
    'ide = pgTagXML("<ide>", "</ide>", docNFe.xml)
    ide_cUF = IIf(IsNull(Rst.Fields("ide_cUF")), "", Rst.Fields("ide_cUF"))
    ide_cNF = IIf(IsNull(Rst.Fields("ide_cNF")), "", Rst.Fields("ide_cNF"))
    ide_natOp = IIf(IsNull(Rst.Fields("ide_natOp")), "", Rst.Fields("ide_natOp"))
    ide_indPag = IIf(IsNull(Rst.Fields("ide_indPag")), "", Rst.Fields("ide_indPag"))
    ide_mod = IIf(IsNull(Rst.Fields("ide_mod")), "", Rst.Fields("ide_mod"))
    ide_serie = IIf(IsNull(Rst.Fields("ide_serie")), "", Rst.Fields("ide_serie"))
    ide_nNF = IIf(IsNull(Rst.Fields("ide_nNF")), "", Rst.Fields("ide_nNF"))
    ide_dEmi = IIf(IsNull(Rst.Fields("ide_dEmi")), "", Rst.Fields("ide_dEmi"))
    ide_dSaiEnt = IIf(IsNull(Rst.Fields("ide_dSaiEnt")), "", Rst.Fields("ide_dSaiEnt"))
    ide_hSaiEnt = IIf(IsNull(Rst.Fields("ide_hSaiEnt")), "", Rst.Fields("ide_hSaiEnt"))
    ide_tpNF = IIf(IsNull(Rst.Fields("ide_tpNF")), "", Rst.Fields("ide_tpNF"))
    ide_cMunFG = IIf(IsNull(Rst.Fields("ide_cMunFG")), "", Rst.Fields("ide_cMunFG"))
    ide_refNFe = IIf(IsNull(Rst.Fields("ide_refNFe")), "", Rst.Fields("ide_refNFe"))
    ide_tpImp = IIf(IsNull(Rst.Fields("ide_tpImp")), "", Rst.Fields("ide_tpImp"))
    ide_tpEmis = IIf(IsNull(Rst.Fields("ide_tpEmis")), "", Rst.Fields("ide_tpEmis"))
    ide_cDV = IIf(IsNull(Rst.Fields("ide_cDV")), "", Rst.Fields("ide_cDV"))
    ide_tpAmb = IIf(IsNull(Rst.Fields("ide_tpAmb")), "", Rst.Fields("ide_tpAmb"))
    ide_finNFe = IIf(IsNull(Rst.Fields("ide_finNFe")), "", Rst.Fields("ide_finNFe"))
    ide_procEmi = IIf(IsNull(Rst.Fields("ide_procEmi")), "", Rst.Fields("ide_procEmi"))
    ide_verProc = IIf(IsNull(Rst.Fields("ide_verProc")), "", Rst.Fields("ide_verProc"))
    'Emitente
'    emit = IIf(IsNull(Rst.Fields("emit")), "", Rst.Fields("emit"))
    emit_CNPJ = IIf(IsNull(Rst.Fields("emit_CNPJ")), "", Rst.Fields("emit_CNPJ"))
    emit_xNome = IIf(IsNull(Rst.Fields("emit_xNome")), "", Rst.Fields("emit_xNome"))
    emit_xFant = IIf(IsNull(Rst.Fields("emit_xFant")), "", Rst.Fields("emit_xFant"))
    emit_xLgr = IIf(IsNull(Rst.Fields("emit_xLgr")), "", Rst.Fields("emit_xLgr"))
    emit_nro = IIf(IsNull(Rst.Fields("emit_nro")), "", Rst.Fields("emit_nro"))
    emit_xCpl = IIf(IsNull(Rst.Fields("emit_xCpl")), "", Rst.Fields("emit_xCpl"))
    emit_Bairro = IIf(IsNull(Rst.Fields("emit_Bairro")), "", Rst.Fields("emit_Bairro"))
    emit_cMun = IIf(IsNull(Rst.Fields("emit_cMun")), "", Rst.Fields("emit_cMun"))
    emit_xMun = IIf(IsNull(Rst.Fields("emit_xMun")), "", Rst.Fields("emit_xMun"))
    emit_UF = IIf(IsNull(Rst.Fields("emit_UF")), "", Rst.Fields("emit_UF"))
    emit_CEP = IIf(IsNull(Rst.Fields("emit_CEP")), "", Rst.Fields("emit_CEP"))
    emit_cPais = IIf(IsNull(Rst.Fields("emit_cPais")), "", Rst.Fields("emit_cPais"))
    emit_xPais = IIf(IsNull(Rst.Fields("emit_xPais")), "", Rst.Fields("emit_xPais"))
    emit_fone = IIf(IsNull(Rst.Fields("emit_fone")), "", Rst.Fields("emit_fone"))
    emit_IE = IIf(IsNull(Rst.Fields("emit_IE")), "", Rst.Fields("emit_IE"))
    emit_IEST = IIf(IsNull(Rst.Fields("emit_IEST")), "", Rst.Fields("emit_IEST"))
    emit_IM = IIf(IsNull(Rst.Fields("emit_IM")), "", Rst.Fields("emit_IM"))
    emit_CNAE = IIf(IsNull(Rst.Fields("emit_CNAE")), "", Rst.Fields("emit_CNAE"))
    emit_CRT = IIf(IsNull(Rst.Fields("emit_CRT")), "", Rst.Fields("emit_CRT"))
    'Destinatario
    'dest_pessoa = PgDadosClienteFornecedor(IdFornecedor).pessoa 'NAO PODE SER DEIXADO EM  BRANCO
    'dest_CNPJ = RS(PgDadosClienteFornecedor(idFornecedor).Doc)
    'dest_xNome = RC(PgDadosClienteFornecedor(IdFornecedor).Nome)
    'dest_xFant = RC(PgDadosClienteFornecedor(IdFornecedor).Fant)
    'dest_xLgr = RC(PgDadosClienteFornecedor(IdFornecedor).Lgr)
    'dest_nro = RC(PgDadosClienteFornecedor(IdFornecedor).Nro)
    'dest_xCpl = RC(PgDadosClienteFornecedor(IdFornecedor).Cpl)
    'dest_Bairro = RC(PgDadosClienteFornecedor(IdFornecedor).Bairro)
    'dest_cMun = PgDadosMunicipio(PgDadosClienteFornecedor(IdFornecedor).UF, PgDadosClienteFornecedor(IdFornecedor).Mun).codMun
    'dest_xMun = RC(PgDadosClienteFornecedor(IdFornecedor).Mun)
    'dest_UF = PgDadosClienteFornecedor(IdFornecedor).UF
    'dest_CEP = RS(PgDadosClienteFornecedor(IdFornecedor).CEP)
    'dest_cPais = emit_cPais
    'dest_xPais = RC(emit_xPais)
    'dest_fone = RS(PgDadosClienteFornecedor(IdFornecedor).Fone)
    'dest_IE = RS(PgDadosClienteFornecedor(IdFornecedor).IE)
    'dest_ISUF = RS(PgDadosClienteFornecedor(IdFornecedor).SUFRAMA)
    'dest_email = RC(PgDadosClienteFornecedor(IdFornecedor).Mail)
    'infAdic_infCpl = IIf(IsNull(Rst.Fields("infAdic_infCpl")), "", Rst.Fields("infAdic_infCpl"))
    
    '****************************************************************************************
'    transp = IIf(IsNull(Rst.Fields("transp")), "", Rst.Fields("transp"))
    transp_modFrete = IIf(IsNull(Rst.Fields("transp_modFrete")), "", Rst.Fields("transp_modFrete"))
    transp_CNPJ = IIf(IsNull(Rst.Fields("transp_CNPJ")), "", Rst.Fields("transp_CNPJ"))
    transp_xNome = IIf(IsNull(Rst.Fields("transp_xNome")), "", Rst.Fields("transp_xNome"))
    transp_IE = IIf(IsNull(Rst.Fields("transp_IE")), "", Rst.Fields("transp_IE"))
    transp_xEnder = IIf(IsNull(Rst.Fields("transp_xEnder")), "", Rst.Fields("transp_xEnder"))
    transp_xMun = IIf(IsNull(Rst.Fields("transp_xMun")), "", Rst.Fields("transp_xMun"))
    transp_UF = IIf(IsNull(Rst.Fields("transp_UF")), "", Rst.Fields("transp_UF"))
    transp_qVol = IIf(IsNull(Rst.Fields("transp_qVol")), "", Rst.Fields("transp_qVol"))
    transp_esp = IIf(IsNull(Rst.Fields("transp_esp")), "", Rst.Fields("transp_esp"))
    transp_marca = IIf(IsNull(Rst.Fields("transp_marca")), "", Rst.Fields("transp_marca"))
    transp_nVol = IIf(IsNull(Rst.Fields("transp_nVol")), "", Rst.Fields("transp_nVol"))
    transp_pesoL = IIf(IsNull(Rst.Fields("transp_pesoL")), "", Rst.Fields("transp_pesoL"))
    transp_pesoB = IIf(IsNull(Rst.Fields("transp_PesoB")), "", Rst.Fields("transp_pesoB"))
    'TOTAIS
'    total = IIf(IsNull(Rst.Fields("total")), "0", Rst.Fields("total"))
    total_vBC = IIf(IsNull(Rst.Fields("total_vBC")), "0", Rst.Fields("total_vBC"))
    total_vICMS = IIf(IsNull(Rst.Fields("total_vICMS")), "0", Rst.Fields("total_vICMS"))
    total_vBCST = IIf(IsNull(Rst.Fields("total_vBCST")), "0", Rst.Fields("total_vBCST"))
    total_vICMSST = IIf(IsNull(Rst.Fields("total_vICMSST")), "0", Rst.Fields("total_vICMSST"))
    total_vCredICMSSN = IIf(IsNull(Rst.Fields("total_vcredICMSSN")), "0", Rst.Fields("total_vcredICMSSN"))
    
    total_vProd = IIf(IsNull(Rst.Fields("total_vProd")), "0", Rst.Fields("total_vProd"))
    total_vFrete = IIf(IsNull(Rst.Fields("total_vFrete")), "0", Rst.Fields("total_vFrete"))
    total_vSeg = IIf(IsNull(Rst.Fields("total_vSeg")), "0", Rst.Fields("total_vSeg"))
    total_vDesc = IIf(IsNull(Rst.Fields("total_vDesc")), "0", Rst.Fields("total_vDesc"))
    total_vOutro = IIf(IsNull(Rst.Fields("total_vOutro")), "0", Rst.Fields("total_vOutro"))
    total_vIPI = IIf(IsNull(Rst.Fields("total_vIPI")), "0", Rst.Fields("total_vIPI"))
    total_vPIS = IIf(IsNull(Rst.Fields("total_vPIS")), "0", Rst.Fields("total_vPIS"))
    total_vCOFINS = IIf(IsNull(Rst.Fields("total_vCOFINS")), "0", Rst.Fields("total_vCOFINS"))
    total_vNF = IIf(IsNull(Rst.Fields("total_vNF")), "0", Rst.Fields("total_vNF"))
    infCpl = IIf(IsNull(Rst.Fields("infAdic_infCpl")), "", Rst.Fields("infAdic_infCpl"))
    
    Rst.Close
    cItens = 1
    'Set Rst = RegistroBuscar("SELECT * FROM FaturamentoNFeENtradaItens WHERE IdNFe='" & NFe & "' ORDER BY Id")
    Set Rst = RegistroBuscar("SELECT * FROM FaturamentoNFeENtradaItens WHERE ID_Empresa = " & ID_Empresa & " AND IdReg=" & IdReg & " ORDER BY Id")
    If Rst.BOF And Rst.EOF Then
        Else
        Rst.MoveFirst
        Do Until Rst.EOF
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> PRODUTO
                                        'id_intProd|det_cProd|det_cEAN|det_xProd|EXTIPI|det_NCM|det_CFOP|det_uCom|
                                        'det_qCom|det_vUnCom|det_vProd|det_cEANTrib|det_uTrib|det_qTrib|
                                        'det_vUnTrib|
                                        'det_vFrete|det_vSeg|det_vDesc|det_vOutro|det_indTot|xPed|nItemPed
                                         'det_indTot = 0 = O valor do item compoe a NF / 1  = O valor do item nao compoe a NF
                        aItem(cItens) = Array(CStr(CStr(IIf(IsNull(Rst.Fields("det_idProduto")), "0", Rst.Fields("det_idProduto")))), _
                                            CStr(CStr(IIf(IsNull(Rst.Fields("det_cProd")), "0", Rst.Fields("det_cProd")))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_cEAN")), "0", Rst.Fields("det_cEAN"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_xProd")), "0", Rst.Fields("det_xProd"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_EXTIPI")), "0", Rst.Fields("det_EXTIPI"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_NCM")), "0", Rst.Fields("det_NCM"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_CFOP")), "0", Rst.Fields("det_CFOP"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_uCom")), "0", Rst.Fields("det_uCom"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_qCom")), "0", Rst.Fields("det_qCom"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_vUnCom")), "0", Rst.Fields("det_vUnCom"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_vProd")), "0", Rst.Fields("det_vProd"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_cEANTrib")), "0", Rst.Fields("det_cEANTrib"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_uTrib")), "0", Rst.Fields("det_uTrib"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_qTrib")), "0", Rst.Fields("det_qTrib"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_vUnTrib")), "0", Rst.Fields("det_vUnTrib"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_vFrete")), "0", Rst.Fields("det_vFrete"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_vSeg")), "0", Rst.Fields("det_vSeg"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_vDesc")), "0", Rst.Fields("det_vDesc"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_vOutro")), "0", Rst.Fields("det_vOutro"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_indTot")), "0", Rst.Fields("det_indTot"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_xPed")), "0", Rst.Fields("det_xPed"))), _
                                            CStr(IIf(IsNull(Rst.Fields("det_nItemPed")), "0", Rst.Fields("det_nItemPed"))))
                            
                      aEstoque(cItens) = Array("0", aItem(cItens)(8), "0")
                        
                                            
                                            'Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST|vICMSST
                        aICMS(cItens) = Array(CStr(IIf(IsNull(Rst.Fields("icms_origem")), "0", Rst.Fields("icms_origem"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_CST")), "0", Rst.Fields("icms_CST"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_ModBC")), "0", Rst.Fields("icms_ModBC"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_pRedBC")), "0", Rst.Fields("icms_pRedBC"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_vBC")), "0", Rst.Fields("icms_vBC"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_pICMS")), "0", Rst.Fields("icms_pICMS"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_vICMS")), "0", Rst.Fields("icms_vICMS"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_ModBCST")), "0", Rst.Fields("icms_ModBCST"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_pMVAST")), "0", Rst.Fields("icms_pMVAST"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_pRedBCST")), "0", Rst.Fields("icms_pRedBCST"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_vBCST")), "0", Rst.Fields("icms_vBCST"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_pICMSST")), "0", Rst.Fields("icms_pICMSST"))), _
                                            CStr(IIf(IsNull(Rst.Fields("icms_vICMSST")), "0", Rst.Fields("icms_vICMSST"))), _
                                            cNull(Rst.Fields("icms_CSOSN")), _
                                            cNull(Rst.Fields("icms_pCredSN")), _
                                            cNull(Rst.Fields("icms_vCredICMSSN")))
                                            
                                            'cEnq|CST|vBC|pIPI|vIPI
                        aIPI(cItens) = Array(CStr(IIf(IsNull(Rst.Fields("ipi_cEnq")), "0", Rst.Fields("ipi_cEnq"))), _
                                            CStr(IIf(IsNull(Rst.Fields("ipi_CST")), "0", Rst.Fields("ipi_CST"))), _
                                            CStr(IIf(IsNull(Rst.Fields("ipi_vBC")), "0", Rst.Fields("ipi_vBC"))), _
                                            CStr(IIf(IsNull(Rst.Fields("ipi_pIPI")), "0", Rst.Fields("ipi_pIPI"))), _
                                            CStr(IIf(IsNull(Rst.Fields("ipi_vIPI")), "0", Rst.Fields("ipi_vIPI"))))
                                            
                                            
                                            'CST|vBC|pPIS|vPIS
                        'tagPIS = pgTagXML("<PIS>", "</PIS>", prod)
                        aPIS(cItens) = Array(CStr(IIf(IsNull(Rst.Fields("PIS_CST")), "", Rst.Fields("PIS_CST"))), _
                                            CStr(IIf(IsNull(Rst.Fields("PIS_vBC")), "0", Rst.Fields("PIS_vBC"))), _
                                            CStr(IIf(IsNull(Rst.Fields("PIS_pPIS")), "0", Rst.Fields("PIS_pPIS"))), _
                                            CStr(IIf(IsNull(Rst.Fields("PIS_vPIS")), "0", Rst.Fields("PIS_vPIS"))))
                                            
                                            'aPIS(cItens)(3) = ChkVal(CStr(aPIS(cItens)(3)), 0, 2)
                                            
                                            'CST|vBC|pCOFINS|vCOFINS
                        'tagCOFINS = pgTagXML("<COFINS>", "</COFINS>", prod)
                        aCOFINS(cItens) = Array(CStr(IIf(IsNull(Rst.Fields("COFINS_CST")), "0", Rst.Fields("COFINS_CST"))), _
                                            CStr(IIf(IsNull(Rst.Fields("COFINS_vBC")), "0", Rst.Fields("COFINS_vBC"))), _
                                            CStr(IIf(IsNull(Rst.Fields("COFINS_pCOFINS")), "0", Rst.Fields("COFINS_pCOFINS"))), _
                                            CStr(IIf(IsNull(Rst.Fields("COFINS_vCOFINS")), "0", Rst.Fields("COFINS_vCOFINS"))))
                                            
                                            'aCOFINS(cItens)(3) = ChkVal(CStr(aCOFINS(cItens)(3)), 0, 2)

                        
                        cItens = cItens + 1
                        Rst.MoveNext
        Loop
                        'Tag = "<det nItem=""" & cItens & """>"
     End If
                    cItens = cItens - 1
            
    'Transporte
  
    '>>>>>>>>>>>>>>>>>>>>>>>>>> COBRANCA
 '   cob = pgTagXML("<cobr>", "</cobr>", docNFe.xml)
    
  '  fat = pgTagXML("<fat>", "</fat>", docNFe.xml)
    
   ' fat_nFat = pgTagXML("<nFat>", "</nFat>", fat)
   ' fat_vOrig = pgTagXML("<vOrig>", "</vOrig>", fat)
   ' fat_vDesc = pgTagXML("<vDesc>", "</vDesc>", fat)
   ' fat_vLiq = pgTagXML("<vLiq>", "</vLiq>", fat)
    cCob = 1
    'Set Rst = RegistroBuscar("SELECT * FROM FaturamentoNFeENtradaCobranca WHERE IdNFe='" & NFe & "' ORDER BY Id")
    Set Rst = RegistroBuscar("SELECT * FROM FaturamentoNFeENtradaCobranca WHERE ID_Empresa = " & ID_Empresa & " AND IdReg=" & IdReg & " ORDER BY Id")
    If Rst.BOF And Rst.EOF Then
        Else
        Rst.MoveFirst
        Do Until Rst.EOF
    
    
        'IdNFe|nfat|vOrig|vDesc|vLiq|nDup|dVenc|vDup|CC|tpDoc
        aCob(cCob) = Array(ide_nNF, _
                            CStr(IIf(IsNull(Rst.Fields("cobr_nFat")), "0", Rst.Fields("cobr_nFat"))), _
                            CStr(IIf(IsNull(Rst.Fields("cobr_vOrig")), "0", Rst.Fields("cobr_vOrig"))), _
                            CStr(IIf(IsNull(Rst.Fields("cobr_vDesc")), "0", Rst.Fields("cobr_vDesc"))), _
                            CStr(IIf(IsNull(Rst.Fields("cobr_vLiq")), "0", Rst.Fields("cobr_vLiq"))), _
                            CStr(IIf(IsNull(Rst.Fields("cobr_nDup")), "0", Rst.Fields("cobr_nDup"))), _
                            CStr(IIf(IsNull(Rst.Fields("cobr_dVenc")), "0", Rst.Fields("cobr_dVenc"))), _
                            CStr(IIf(IsNull(Rst.Fields("cobr_vDup")), "0", Rst.Fields("cobr_vDup"))), _
                            CStr(IIf(IsNull(Rst.Fields("cobr_CC")), "0", Rst.Fields("cobr_CC"))), _
                            CStr(IIf(IsNull(Rst.Fields("cobr_tpDoc")), "0", Rst.Fields("cobr_tpDoc"))), _
                            CStr(IIf(IsNull(Rst.Fields("cobr_PlanoContas")), "0", Rst.Fields("cobr_PlanoContas"))))
    
        cCob = cCob + 1
        Rst.MoveNext 'cob = Mid(cob, InStr(cob, "</dup>") + Len("</dup>"), Len(cob))
    Loop
    End If
    cCob = cCob - 1
    MostrarDadosForm
    
    
    sstNF.Tab = 0
    HDForm Me, False
    msfgProdutos.Enabled = True
    msfgDupl.Enabled = True
    
End Sub

Private Sub ImportarNFe()
    With cd
        .Filter = "XML Files (*.xml) |*.xml"
        .ShowOpen
        If Len(.filename) = 0 Then Exit Sub
        fileXMLOrigem = .filename
        If LoadXML(fileXMLOrigem) = True Then
            MostrarDadosForm
            'Checa se a nota ja foi cadastrada
            If nfeCadastrada = True Then
                MsgBox "Nota fiscal ja registrada no sistema.", vbInformation, App.EXEName
            End If
        End If
    End With
End Sub

Private Function LoadXML(fArquivo As String) As Boolean
    On Error GoTo TrtErroXML
    Dim docNFe      As DOMDocument60
    Dim NFe         As String
    Dim fat         As String
    Dim cob         As String
    Dim ide         As String
    Dim emit        As String
    Dim prod        As String
    Dim transp      As String
    Dim total       As String
    Dim tagICMS     As String
    Dim tagIPI      As String
    Dim tagPIS      As String
    Dim tagCOFINS   As String
    Dim i           As Integer
    Set docNFe = New DOMDocument60
    
    docNFe.resolveExternals = True
    docNFe.validateOnParse = True
    docNFe.async = False
    
    Call docNFe.Load(fArquivo)
    
    'Checa se houve algum erro ao carregar
    If docNFe.parseError.reason <> "" Then
        MsgBox "Erro ao ler XML : " & docNFe.parseError.reason & vbCrLf & _
            "Linha: " & docNFe.parseError.line & vbCrLf & _
            "Texto: " & docNFe.parseError.srcText & vbCrLf & _
            "Codigo do Erro: " & docNFe.parseError.errorCode, vbInformation, "Aviso"
        Exit Function
    End If
   '***************************************************
   '* Modificacoes para atender a NFE 3.1
   '* Conforme NT 2013.005_V.1.10
   '***************************************************
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>IDENTIFICADOR DA NOTA
    NFe = pgTagXML("infNFe", ">", docNFe.xml)
    Id = Replace(Mid(NFe, InStr(NFe, "NFe"), 47), "NFe", "")
    Versao = pgTagXML("versao=""", """", NFe)
    
    Versao = Replace(Versao, """", "") 'Replace(Replace(Mid(NFe, InStr(NFe, "versao"), Len(NFe) - 2), "versao=", ""), """", "")
    nProt = pgTagXML("<nProt>", "</nProt>", docNFe.xml)
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> IDE
    ide = pgTagXML("<ide>", "</ide>", docNFe.xml)
    ide_cUF = pgTagXML("<cUF>", "</cUF>", ide)
    ide_cNF = pgTagXML("<cNF>", "</cNF>", ide)
    ide_natOp = pgTagXML("<natOp>", "</natOp>", ide)
    ide_indPag = pgTagXML("<indPag>", "</indPag>", ide)
    ide_mod = pgTagXML("<mod>", "</mod>", ide)
    ide_serie = pgTagXML("<serie>", "</serie>", ide)
    ide_nNF = pgTagXML("<nNF>", "</nNF>", ide)
    '***************************************************
   '* Modificacoes para atender a NFE 3.1 ou superiores
   '* Conforme NT 2013.005_V.1.10
    Dim lTag As String
    If Val(RS(Versao)) >= 310 Then
            ide_dEmi = pgTagXML("<dhEmi>", "</dhEmi>", ide)
            ide_dEmi = Mid(ide_dEmi, 1, InStr(ide_dEmi, "T") - 1)
            ide_dSaiEnt = Left(pgTagXML("<dhSaiEnt>", "</dhSaiEnt>", ide), 10)
        Else
         ide_dEmi = pgTagXML("<dEmi>", "</dEmi>", ide)
         ide_dSaiEnt = pgTagXML("<dSaiEnt>", "</dSaiEnt>", ide)
    End If
   '***************************************************
    
    ide_hSaiEnt = pgTagXML("<hSaiEnt>", "</hSaiEnt>", ide)
    ide_tpNF = pgTagXML("<tpNF>", "</tpNF>", ide)
    ide_cMunFG = pgTagXML("<cMunFG>", "</cMunFG>", ide)
    ide_refNFe = pgTagXML("<refNFe>", "</refNFe>", ide)
    ide_tpImp = pgTagXML("<tpImp>", "</tpImp>", ide)
    ide_tpEmis = pgTagXML("<tpEmis>", "</tpEmis>", ide)
    ide_cDV = pgTagXML("<cDV>", "</cDV>", ide)
    ide_tpAmb = pgTagXML("<tpAmb>", "</tpAmb>", ide)
    ide_finNFe = pgTagXML("<finNFe>", "</finNFe>", ide)
    ide_procEmi = pgTagXML("<procEmi>", "</procEmi>", ide)
    ide_verProc = pgTagXML("<verProc>", "</verProc>", ide)
    'Emitente
    emit = pgTagXML("<emit>", "</emit>", docNFe.xml)
    emit_CNPJ = pgTagXML("<CNPJ>", "</CNPJ>", emit)
    emit_xNome = pgTagXML("<xNome>", "</xNome>", emit)
    emit_xFant = pgTagXML("<xFant>", "</xFant>", emit)
    emit_xLgr = pgTagXML("<xLgr>", "</xLgr>", emit)
    emit_nro = pgTagXML("<nro>", "</nro>", emit)
    emit_xCpl = pgTagXML("<xCpl>", "</xCpl>", emit)
    emit_Bairro = pgTagXML("<xBairro>", "</xBairro>", emit)
    emit_cMun = pgTagXML("<cMun>", "</cMun>", emit)
    emit_xMun = pgTagXML("<xMun>", "</xMun>", emit)
    emit_UF = pgTagXML("<UF>", "</UF>", emit)
    emit_CEP = pgTagXML("<CEP>", "</CEP>", emit)
    emit_cPais = pgTagXML("<cPais>", "</cPais>", emit)
    emit_xPais = pgTagXML("<xPais>", "</xPais>", emit)
    emit_fone = pgTagXML("<fone>", "</fone>", emit)
    emit_IE = pgTagXML("<IE>", "</IE>", emit)
    emit_IEST = pgTagXML("<IEST>", "</IEST>", emit)
    emit_IM = pgTagXML("<IM>", "</IM>", emit)
    emit_CNAE = pgTagXML("<CNAE>", "</CNAE>", emit)
    emit_CRT = pgTagXML("<CRT>", "</CRT>", emit)
    'Destinatario
    dest = pgTagXML("<dest>", "</dest>", docNFe.xml)
    dest_CNPJ = pgTagXML("<CNPJ>", "</CNPJ>", dest)
    'dest_pessoa = PgDadosClienteFornecedor(IdFornecedor).pessoa 'NAO PODE SER DEIXADO EM  BRANCO
    'dest_CNPJ = RS(PgDadosClienteFornecedor(idFornecedor).Doc)
    'dest_xNome = RC(PgDadosClienteFornecedor(IdFornecedor).Nome)
    'dest_xFant = RC(PgDadosClienteFornecedor(IdFornecedor).Fant)
    'dest_xLgr = RC(PgDadosClienteFornecedor(IdFornecedor).Lgr)
    'dest_nro = RC(PgDadosClienteFornecedor(IdFornecedor).Nro)
    'dest_xCpl = RC(PgDadosClienteFornecedor(IdFornecedor).Cpl)
    'dest_Bairro = RC(PgDadosClienteFornecedor(IdFornecedor).Bairro)
    'dest_cMun = PgDadosMunicipio(PgDadosClienteFornecedor(IdFornecedor).UF, PgDadosClienteFornecedor(IdFornecedor).Mun).codMun
    'dest_xMun = RC(PgDadosClienteFornecedor(IdFornecedor).Mun)
    'dest_UF = PgDadosClienteFornecedor(IdFornecedor).UF
    'dest_CEP = RS(PgDadosClienteFornecedor(IdFornecedor).CEP)
    'dest_cPais = emit_cPais
    'dest_xPais = RC(emit_xPais)
    'dest_fone = RS(PgDadosClienteFornecedor(IdFornecedor).Fone)
    'dest_IE = RS(PgDadosClienteFornecedor(IdFornecedor).IE)
    'dest_ISUF = RS(PgDadosClienteFornecedor(IdFornecedor).SUFRAMA)
    'dest_email = RC(PgDadosClienteFornecedor(IdFornecedor).Mail)
    'infAdic_infCpl = pgTagXML("<infCpl>", "</infCpl>", ide)
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> PRODUTO
     cItens = 0
     'Tag = "<det nItem=""" & cItens & """>"
     'Do Until InStr(docNFe.xml, Tag) = 0
     For i = 1 To 999
        Tag = "<det nItem=""" & i & """>"
        If InStr(docNFe.xml, Tag) <> 0 Then
            cItens = cItens + 1
            prod = revCarac(pgTagXML(Tag, "</det>", docNFe.xml))
                                        'id_intProd|det_cProd|det_cEAN|det_xProd|EXTIPI|det_NCM|det_CFOP|det_uCom|
                                        'det_qCom|det_vUnCom|det_vProd|det_cEANTrib|det_uTrib|det_qTrib|
                                        'det_vUnTrib|
                                        'det_vFrete|det_vSeg|det_vDesc|det_vOutro|det_indTot|xPed|nItemPed
                                         'det_indTot = 0 = O valor do item compoe a NF / 1  = O valor do item nao compoe a NF
                        aItem(cItens) = Array("0", _
                                            CStr(pgTagXML("<cProd>", "</cProd>", prod)), _
                                            pgTagXML("<cEAN>", "</cEAN>", prod), _
                                            CStr(pgTagXML("<xProd>", "</xProd>", prod)), _
                                            pgTagXML("<EXTIPI>", "</EXTIPI>", prod), _
                                            pgTagXML("<NCM>", "</NCM>", prod), _
                                            pgTagXML("<CFOP>", "</CFOP>", prod), _
                                            CStr(pgTagXML("<uCom>", "</uCom>", prod)), _
                                            CStr(pgTagXML("<qCom>", "</qCom>", prod)), _
                                            CStr(pgTagXML("<vUnCom>", "</vUnCom>", prod)), _
                                            pgTagXML("<vProd>", "</vProd>", prod), _
                                            pgTagXML("<cEANTrib>", "</cEANTrib>", prod), _
                                            CStr(pgTagXML("<uTrib>", "</uTrib>", prod)), _
                                            CStr(pgTagXML("<qTrib>", "</qTrib>", prod)), _
                                            CStr(pgTagXML("<vUnTrib>", "</vUnTrib>", prod)), _
                                            pgTagXML("<vFrete>", "</vFrete", prod), _
                                            pgTagXML("<vSeg>", "</vSeg>", prod), _
                                            pgTagXML("<vDesc>", "</vDesc>", prod), _
                                            pgTagXML("<vOutro>", "</vOutro>", prod), _
                                            pgTagXML("<indTot>", "</indTot>", prod), _
                                            CStr(pgTagXML("<xPed>", "</xPed>", prod)), _
                                            pgTagXML("<nItemPed>", "</nItemPed>", prod))
                      aItem(cItens)(0) = PgIDMaterial(CStr(aItem(cItens)(1)), emit_CNPJ)
                      aEstoque(cItens) = Array("0", aItem(cItens)(8), "0")
                        
                                            
                                            'Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST|vICMSST|CSOSN
                        tagICMS = pgTagXML("<ICMS>", "</ICMS>", prod)
                        aICMS(cItens) = Array(pgTagXML("<orig>", "</orig>", tagICMS), _
                                            pgTagXML("<CST>", "</CST>", tagICMS), _
                                            pgTagXML("<ModBC>", "</ModBC>", tagICMS), _
                                            pgTagXML("<pRedBC>", "</pRedBC>", tagICMS), _
                                            pgTagXML("<vBC>", "</vBC>", tagICMS), _
                                            ChkVal(pgTagXML("<pICMS>", "</pICMS>", tagICMS), 0, 2), _
                                            pgTagXML("<vICMS>", "</vICMS>", tagICMS), _
                                            pgTagXML("<modBCST>", "</modBCST>", tagICMS), _
                                            pgTagXML("<pMVAST>", "</pMVAST>", tagICMS), _
                                            pgTagXML("<pRedBCST>", "</pRedBCST>", tagICMS), _
                                            pgTagXML("<vBCST>", "</vBCST>", tagICMS), _
                                            pgTagXML("<pICMSST>", "</pICMSST>", tagICMS), _
                                            pgTagXML("<vICMSST>", "</vICMSST>", tagICMS), _
                                            pgTagXML("<CSOSN>", "</CSOSN>", tagICMS), _
                                            pgTagXML("<pCredSN>", "</pCredSN>", tagICMS), _
                                            pgTagXML("<vCredICMSSN>", "</vCredICMSSN>", tagICMS))
                                            
                                            'a partir do CSOSN para contemplar Empresas no Simples Nacional
                                            
                                            
                                            'cEnq|CST|vBC|pIPI|vIPI
                        tagIPI = pgTagXML("<IPI>", "</IPI>", prod)
                        aIPI(cItens) = Array(pgTagXML("<cEnq>", "</cEnq>", tagIPI), _
                                            pgTagXML("<CST>", "</CST>", tagIPI), _
                                            pgTagXML("<vBC>", "</vBC>", tagIPI), _
                                            pgTagXML("<pIPI>", "</pIPI>", tagIPI), _
                                            pgTagXML("<vIPI>", "</vIPI>", tagIPI))
                                            
                                            
                                            'CST|vBC|pPIS|vPIS
                        tagPIS = pgTagXML("<PIS>", "</PIS>", prod)
                        aPIS(cItens) = Array(pgTagXML("<CST>", "</CST>", tagPIS), _
                                            CStr(pgTagXML("<vBC>", "</vBC>", tagPIS)), _
                                            pgTagXML("<pPIS>", "</pPIS>", tagPIS), _
                                            pgTagXML("<vPIS>", "</vPIS>", tagPIS))
                                            
                                            'aPIS(cItens)(3) = ChkVal(CStr(aPIS(cItens)(3)), 0, 2)
                                            
                                            'CST|vBC|pCOFINS|vCOFINS
                        tagCOFINS = pgTagXML("<COFINS>", "</COFINS>", prod)
                        aCOFINS(cItens) = Array(pgTagXML("<CST>", "</CST>", tagCOFINS), _
                                            CStr(pgTagXML("<vBC>", "</vBC>", tagCOFINS)), _
                                            pgTagXML("<pCOFINS>", "</pCOFINS>", tagCOFINS), _
                                            pgTagXML("<vCOFINS>", "</vCOFINS>", tagCOFINS))
                                            
                                            'aCOFINS(cItens)(3) = ChkVal(CStr(aCOFINS(cItens)(3)), 0, 2)

                        
        'cItens = cItens + 1
        'Tag = "<det nItem=""" & cItens & """>"
    'Loop
        End If
    Next
    'cItens = cItens - 1
            'End If
    'Transporte
    transp = pgTagXML("<transp>", "</transp>", docNFe.xml)
    transp_modFrete = pgTagXML("<modFrete>", "</modFrete>", transp)
    transp_CNPJ = pgTagXML("<CNPJ>", "</CNPJ>", transp)
    transp_xNome = pgTagXML("<xNome>", "</xNome>", transp)
    transp_IE = pgTagXML("<IE>", "</IE>", transp)
    transp_xEnder = pgTagXML("<xEnder>", "</xEnder>", transp)
    transp_xMun = pgTagXML("<xMun>", "</xMun>", transp)
    transp_UF = pgTagXML("<UF>", "</UF>", transp)
    transp_qVol = pgTagXML("<qVol>", "</qVol>", transp)
    transp_esp = pgTagXML("<esp>", "</esp>", transp)
    transp_marca = pgTagXML("<marca>", "</marca>", transp)
    transp_nVol = pgTagXML("<nVol>", "</nVol>", transp)
    transp_pesoL = pgTagXML("<pesoL>", "</pesoL>", transp)
    transp_pesoB = pgTagXML("<pesoB>", "</pesoB>", transp)
    'TOTAIS
    total = pgTagXML("<total>", "</total>", docNFe.xml)
    total_vBC = pgTagXML("<vBC>", "</vBC>", total)
    
    
    '#############################################################################################
    '### 15/03/2012
    '### Funcao removida do bloco visto q empresas no simples tbm podem dar cred de ICMS
    '### Falta armazenar mais informacoes quanto a empresas no simples
    '#############################################################################################
    'If emit_CRT = 3 Then
            total_vICMS = pgTagXML("<vICMS>", "</vICMS>", total)
   '     Else
   '         total_vICMS = "0"
            'For i = 1 To cItens
            '    total_vICMS = Val(ChkVal(CStr(aICMS(i)(6)), 0, cDecMoeda)) + Val(ChkVal(total_vICMS, 0, cDecMoeda))
            'Next
            total_vICMS = ChkVal(total_vICMS, 0, cDecMoeda)
    'End If
    '#############################################################################################
    
    total_vBCST = pgTagXML("<vBCST>", "</vBCST>", total)
    total_vICMSST = pgTagXML("<vST>", "</vST>", total)
    total_vProd = pgTagXML("<vProd>", "</vProd>", total)
    total_vFrete = pgTagXML("<vFrete>", "</vFrete>", total)
    total_vSeg = pgTagXML("<vSeg>", "</vSeg>", total)
    total_vDesc = pgTagXML("<vDesc>", "</vDesc>", total)
    total_vOutro = pgTagXML("<vOutro>", "</vOutro>", total)
    total_vIPI = pgTagXML("<vIPI>", "</vIPI>", total)
    total_vPIS = pgTagXML("<vPIS>", "</vPIS>", total)
    total_vCOFINS = pgTagXML("<vCOFINS>", "</vCOFINS>", total)
    total_vNF = pgTagXML("<vNF>", "</vNF>", total)
    '>>>>>>>>>>>>>>>>>>>>>>>>>> COBRANCA
    cob = pgTagXML("<cobr>", "</cobr>", docNFe.xml)
    
    fat = pgTagXML("<fat>", "</fat>", docNFe.xml)
    
    fat_nFat = pgTagXML("<nFat>", "</nFat>", fat)
    fat_vOrig = pgTagXML("<vOrig>", "</vOrig>", fat)
    fat_vDesc = pgTagXML("<vDesc>", "</vDesc>", fat)
    fat_vLiq = pgTagXML("<vLiq>", "</vLiq>", fat)
    
    
    cCob = 1
    Do Until InStr(cob, "</dup>") = 0
        'IdNFe|nfat|vOrig|vDesc|vLiq|nDup|dVenc|vDup|CC|tpDoc|PlanoContas
        aCob(cCob) = Array(ide_nNF, _
                            fat_nFat, _
                            fat_vOrig, _
                            fat_vDesc, _
                            fat_vLiq, _
                            pgTagXML("<nDup>", "</nDup>", cob), _
                            Format(pgTagXML("<dVenc>", "</dVenc>", cob), "DD/MM/YYYY"), _
                            ConvMoeda(pgTagXML("<vDup>", "</vDup>", cob)), _
                            PgDadosConfig.FornecedorCC, PgDadosConfig.FornecedorTpDoc, _
                            PgDadosConfig.FornecedorPlanoContas)
    
        cCob = cCob + 1
        cob = Mid(cob, InStr(cob, "</dup>") + Len("</dup>"), Len(cob))
    Loop
    cCob = cCob - 1
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>> INFORMACOES COMPLEMENTARES
    infCpl = pgTagXML("<infCpl>", "</infCpl>", docNFe.xml)
    
    
    
    LoadXML = True
    Exit Function
TrtErroXML:
    RegLog "0", "formFaturamentoNFeEntrada.LoadXML", Err.Number & " - " & Err.Description
    LoadXML = False
End Function
Private Sub pgDadosEstoque()
    'Pega os dados para baixa no estoque
    Dim i As Integer
    i = lnProd 'msfgProdutos.Row
                        If UCase(aItem(i)(7)) <> UCase(pgDadosEstoqueProduto(idProd).Unidade) Then
                                'Estoque_Unid|Estoque_Qtd|Estoque_vUnit
                        
                                aEstoque(i)(0) = pgDadosEstoqueProduto(idProd).Unidade
                                aEstoque(i)(1) = InputBox("A UNIDADE vendida (" & UCase(aItem(i)(7)) & ") diverge da unidade de armazenamento." & vbCrLf & vbCrLf & _
                                                            "Informe a quantidade TOTAL em " & aEstoque(i)(0) & " do:" & vbCrLf & vbCrLf & _
                                                            "Item " & i & " - " & aItem(i)(3), _
                                                            "Dados para Baixa de Estoque", aItem(i)(8))
                                aEstoque(i)(1) = ChkVal(IIf(Trim(aEstoque(i)(1)) = "", aItem(i)(8), aEstoque(i)(1)), 0, cDecQtd)
                        
                        
                                aEstoque(i)(2) = ChkVal(Val(aItem(i)(10)) / Val(aEstoque(i)(1)), 0, cDecMoeda)
                        
                                'Me.Height = 8355
                                'List1.AddItem "item: " & Left("000", 3 - Len(i)) & i + 1 & " - Unidade: " & aEstoque(i)(0) & " - Quantidade: " & aEstoque(i)(1) & " - Valor Unitario: " & aEstoque(i)(2)
                            Else
                                aEstoque(i)(0) = pgDadosEstoqueProduto(idProd).Unidade
                                aEstoque(i)(1) = aItem(i)(8)
                                aEstoque(i)(2) = aItem(i)(9)
                        End If
End Sub

Private Function grvRegistro() As Boolean
    Dim vReg(1000)   As Variant
    Dim cReg         As Integer 'Contador de Registros
    Dim idCob        As Integer  'Pega o Id de inclusao da Cobranca
    Dim i            As Integer
    
    'Registra as variaveis caso nao seja importacao do XML
    If txtChaveAcesso.Enabled = True Then
        MontarArrayIde
    End If
    MontarArrayTotais
    
    If ValidarDados = False Then
        grvRegistro = False
        Exit Function
    End If
    
    
    cReg = 0
     
    vReg(cReg) = Array("MovEstoque", chkMovEstoque.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("MovFinanceiro", chkMovFinanceiro.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("NFDevolucao", chkNFDevolucao, "N"): cReg = cReg + 1
    vReg(cReg) = Array("MovFisco", chkMovFisco.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("retICMSST", chkretICMSST.Value, "N"): cReg = cReg + 1
    
    vReg(cReg) = Array("IdNFe", Id, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Versao", Versao, "S"): cReg = cReg + 1
    vReg(cReg) = Array("nProt", nProt, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_cUF", ide_cUF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_cNF", ide_cNF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_natOp", ide_natOp, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_indPag", ide_indPag, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_mod", ide_mod, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_Serie", ide_serie, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_nNF", ide_nNF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_dEmi", ide_dEmi, "D"): cReg = cReg + 1
    vReg(cReg) = Array("ide_dSaiEnt", ide_dSaiEnt, "D"): cReg = cReg + 1
    vReg(cReg) = Array("ide_hSaiEnt", ide_hSaiEnt, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_tpNf", ide_tpNF, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_cMunFG", ide_cMunFG, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_refNFe", ide_refNFe, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_tpImp", ide_tpImp, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_tpEmis", ide_tpEmis, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_cDV", ide_cDV, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_tpAmb", ide_tpAmb, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_finNFe", ide_finNFe, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_procEmi", ide_procEmi, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_verProc", ide_verProc, "S"): cReg = cReg + 1
    
    'Emitente
    vReg(cReg) = Array("emit_id", emit_id, "N"): cReg = cReg + 1
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
    vReg(cReg) = Array("dest_idDest", "10", "N"): cReg = cReg + 1
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
    'Infoemacoes complementares
    'vReg(cReg) = Array("infAdic_infCpl", infAdic_infCpl, "S"): cReg = cReg + 1
    'Transportador
    vReg(cReg) = Array("transp_modFrete", transp_modFrete, "S"): cReg = cReg + 1
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
    'TOTAIS
    vReg(cReg) = Array("total_vBC", IIf(Trim(total_vBC) = "", "0.00", total_vBC), "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vICMS", total_vICMS, "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vBCST", IIf(Trim(total_vBCST) = "", "0.00", total_vBCST), "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vICMSST", IIf(Trim(total_vICMSST) = "", "0.00", total_vICMSST), "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vCredICMSSN", IIf(Trim(total_vCredICMSSN) = "", "0.00", total_vCredICMSSN), "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("total_vProd", IIf(Trim(total_vProd) = "", "0.00", total_vProd), "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vFrete", IIf(Trim(total_vFrete) = "", "0.00", total_vFrete), "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vSeg", IIf(Trim(total_vSeg) = "", "0.00", total_vSeg), "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vDesc", IIf(Trim(total_vDesc) = "", "0.00", total_vDesc), "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vOutro", IIf(Trim(total_vOutro) = "", "0.00", total_vOutro), "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vIPI", IIf(Trim(total_vIPI) = "", "0.00", total_vIPI), "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vPIS", IIf(Trim(total_vPIS) = "", "0.00", total_vPIS), "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vCOFINS", IIf(Trim(total_vCOFINS) = "", "0.00", total_vCOFINS), "S"): cReg = cReg + 1
    vReg(cReg) = Array("total_vNF", IIf(Trim(total_vNF) = "", "0.00", total_vNF), "S"): cReg = cReg + 1
    vReg(cReg) = Array("infAdic_infCpl", IIf(Trim(infCpl) = "", "", infCpl), "S"): cReg = cReg + 1
    vReg(cReg) = Array("ger_Vendedor", ger_Vendedor, "N") ': cReg = cReg + 1
    
'     If IdReg = 0 Then
            IdReg = RegistroIncluir(strTabela, vReg, cReg)
            If IdReg = 0 Then
                    MsgBox "Erro ao Incluir CAB"
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
 '****************** Descricao dos itens da NFe *******************************************************
 
    cReg = 0
    For i = 1 To cItens
        '****** Movimenta o estoque **************************************
        If chkMovEstoque.Value = 1 Then
            If MovimentarEstoque("e", _
                                CLng(aItem(i)(0)), _
                                CDate(ide_dEmi), _
                                ide_nNF, _
                                CStr(aEstoque(i)(1)), _
                                CStr(aEstoque(i)(2)), _
                                CStr(aItem(i)(10)), _
                                "Unid.: " & aItem(i)(7) & "  Qtd.: " & aItem(i)(8) & " Vl.Unit.: " & ConvMoeda(CStr(aItem(i)(9))), _
                                emit_xNome, Id, emit_id, emit_CNPJ) = False Then
                MsgBox "Erro ao Movimentar Estoque com o item n. " & i
            End If
        End If
        '*******************************************************************
        
        '*******************************************************************
        '*** Objetivo: Atualizar Custo do Produto
        '*** Data: 08/07/2011
        If PgDadosConfig.EstoqueAtualizarCusto = 1 Then
            AtualizarCustos CInt(aItem(i)(0)), CStr(aEstoque(i)(2))
        End If
        '*******************************************************************
        cReg = 0
        vReg(cReg) = Array("IdReg", IdReg, "S"): cReg = cReg + 1
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
        
        vReg(cReg) = Array("ICMS_CSOSN", aICMS(i)(13), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_pCredSN", aICMS(i)(14), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ICMS_vCredICMSSN", aICMS(i)(15), "S"): cReg = cReg + 1
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
        'ESTOQUE
        If chkMovEstoque.Value = 1 Then
            vReg(cReg) = Array("Estoque_Unid", pgDadosEstoqueProduto(CLng(aItem(i)(0))).Unidade, "S"): cReg = cReg + 1
            vReg(cReg) = Array("Estoque_Qtd", aEstoque(i)(1), "S"): cReg = cReg + 1
            vReg(cReg) = Array("Estoque_vUnit", aEstoque(i)(2), "S"): cReg = cReg + 1
        End If
'*****************************************************************************************************************
         cReg = cReg - 1
        RegistroIncluir strTabela & "Itens", vReg, cReg
        'If IdReg = 0 Then
        '        MsgBox "Erro ao Incluir ITEM"
        '        cReg = 0
        '        grvRegistro = False
        '    Else
        '        cReg = 0
        '        grvRegistro = True
        'End If

    Next
    'cReg = cReg - 1
    
     
'******************* COBRANCA da NFe *******************************************************

'Modificado para atender o que manda o grid
'IdNFe|nfat|vOrig|vDesc|vLiq|nDup|dVenc|vDup|CC|tpDoc
'    cReg = 0
'    For i = 1 To cCob
'        vReg(cReg) = Array("IdNFe", Id, "S"): cReg = cReg + 1
'        vReg(cReg) = Array("cobr_nFat", aCob(i)(0), "S"): cReg = cReg + 1
'        vReg(cReg) = Array("cobr_vOrig", aCob(i)(1), "S"): cReg = cReg + 1
'        vReg(cReg) = Array("cobr_vDesc", aCob(i)(2), "S"): cReg = cReg + 1
'        vReg(cReg) = Array("cobr_vLiq", aCob(i)(3), "S"): cReg = cReg + 1
'        vReg(cReg) = Array("cobr_nDup", aCob(i)(4), "S"): cReg = cReg + 1
'        vReg(cReg) = Array("cobr_dVenc", aCob(i)(5), "D"): cReg = cReg + 1
'        vReg(cReg) = Array("cobr_vDup", aCob(i)(6), "S") ': cReg = cReg + 1
'        IdReg = RegistroIncluir(strTabela & "Cobranca", vReg, cReg)
    cReg = 0
    For i = 1 To msfgDupl.Rows - 1
     'IdNFe|nfat|vOrig|vDesc|vLiq|nDup|dVenc|vDup|CC|tpDoc|PlanoContas
        vReg(cReg) = Array("IdReg", IdReg, "S"): cReg = cReg + 1
        vReg(cReg) = Array("idNFe", Id, "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_nFat", aCob(i)(1), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_vOrig", aCob(i)(2), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_vDesc", aCob(i)(3), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_vLiq", aCob(i)(4), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_nDup", msfgDupl.TextMatrix(i, 1), "S"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_dVenc", msfgDupl.TextMatrix(i, 2), "D"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_vDup", ChkVal(msfgDupl.TextMatrix(i, 3), 0, 2), "S"): cReg = cReg + 1
        
        vReg(cReg) = Array("cobr_PlanoContas", Left(msfgDupl.TextMatrix(i, 6), 3), "N"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_CC", Left(msfgDupl.TextMatrix(i, 4), 3), "N"): cReg = cReg + 1
        vReg(cReg) = Array("cobr_tpDoc", Left(msfgDupl.TextMatrix(i, 5), 3), "N") ': cReg = cReg +
        
        idCob = RegistroIncluir(strTabela & "Cobranca", vReg, cReg)
        
        
        
        If idCob = 0 Then
                MsgBox "Erro ao Incluir COBRANÇA", vbCritical, "Aviso"
                cReg = 0
                grvRegistro = False
            Else
                cReg = 0
                grvRegistro = True
        End If
        'MovimentarContasPagarReceber "P", CDate(ide_dEmi), ide_nNF, CStr(aCob(i)(4)), "Fornecedores", emit_id, emit_xNome, emit_CNPJ, "0", Left(msfgDupl.TextMatrix(i, 4), 3), Left(msfgDupl.TextMatrix(i, 5), 3), "", "", CDate(aCob(i)(6)), _
                                    CStr(aCob(i)(5)), "0", "0", "0", "0", "0", "0", "0", CStr(aCob(i)(7)), "", Id
        If chkMovFinanceiro.Value = 1 Then
            MovimentarContasPagarReceber "P", CDate(ide_dEmi), ide_nNF, CStr(aCob(i)(4)), "Fornecedores", emit_id, emit_xNome, emit_CNPJ, "0", Left(msfgDupl.TextMatrix(i, 4), 3), Left(msfgDupl.TextMatrix(i, 5), 3), Left(msfgDupl.TextMatrix(i, 6), 3), "", "", msfgDupl.TextMatrix(i, 2), _
                                        msfgDupl.TextMatrix(i, 1), "0", "0", "0", "0", "0", "0", "0", ChkVal(msfgDupl.TextMatrix(i, 3), 0, 2), "", IIf(Trim(Id) = "", IdReg, Id)
        End If
    Next
        
'    Next
End Function
Private Sub cboUnidade_DropDown()
    Dim Rst As Recordset
    cboUnidade.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueUnidadeMedida ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboUnidade.AddItem Rst.Fields("sigla")
                Rst.MoveNext
            Loop
    End If

End Sub


Private Sub txtBCICMS_Change()
    calcTotais
End Sub

Private Sub txtBCICMS_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtBCICMS.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtBCICMSp_Change()
    CalcItem
End Sub

Private Sub txtBCICMSST_Change()
    calcTotais
End Sub

Private Sub txtBCICMSST_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtBCICMSST.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarProduto
    End If

End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    
End Sub


Private Sub txtDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        idFornecedor = 0
        PesquisarForn
    End If

End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDoc.Text) <> "" Then
            PesquisarForn (Trim(txtDoc.Text))
        End If
    End If
End Sub


Private Sub txtFreteConta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    KeyAscii = IIf(IsNumeric(Chr(KeyAscii)), KeyAscii, 0)
End Sub

Private Sub txtIdProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarProduto
    End If
End Sub


Private Sub txtIdProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        idProd = txtIdProd.Text
        PesquisarProduto
    End If
End Sub






Private Sub txtpICMSp_Change()
    CalcItem
End Sub

Private Sub txtpIPIp_Change()
    CalcItem
End Sub

Private Sub txtQtd_Change()
    CalcItem
End Sub

Private Sub txtQtd_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtQtd.Text, KeyAscii, cDecQtd)
    
End Sub



Private Sub txtTranspCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarTransp
    End If
End Sub

Private Sub txtvCOFINS_Change()
    calcTotais
End Sub

Private Sub txtvCOFINS_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvCOFINS.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtvCredICMSSN_Change()
    calcTotais
End Sub

Private Sub txtvCredICMSSN_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvCredICMSSN.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtvDesconto_Change()
    calcTotais
End Sub

Private Sub txtvDesconto_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvDesconto.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvDupl_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvDupl.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvFrete_Change()
    calcTotais
End Sub

Private Sub txtvFrete_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvFrete.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvICMS_Change()
    calcTotais
End Sub

Private Sub txtvICMS_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvICMS.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvICMSST_Change()
    calcTotais
End Sub

Private Sub txtvICMSST_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvICMSST.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvIPI_Change()
    calcTotais
End Sub

Private Sub txtvIPI_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvIPI.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvOutras_Change()
    calcTotais
End Sub

Private Sub txtvOutras_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvOutras.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvPIS_Change()
    calcTotais
End Sub

Private Sub txtvPIS_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvPIS.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtvProd_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub txtvProduto_Change()
    calcTotais
End Sub

Private Sub txtvProduto_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvProduto.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvSeguro_Change()
    calcTotais
End Sub

Private Sub txtvSeguro_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvSeguro.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvTotalNF_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvTotalNF.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvUnit_Change()
    CalcItem
End Sub

Private Sub txtvUnit_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvUnit.Text, KeyAscii, cDecMoeda)
End Sub

Private Function revCarac(sCampo As String) As String 'Reverter Caracteres
    sCampo = Replace(sCampo, "&lt;", "<")
    sCampo = Replace(sCampo, "&gt;", ">")
    sCampo = Replace(sCampo, "&amp;", "&")
    sCampo = Replace(sCampo, "&quot;", """")
    sCampo = Replace(sCampo, "&#39;", "´")
    revCarac = sCampo
End Function
Private Sub PesquisarTransp(Optional Id As String)
    If Trim(IdTransp) = 0 Then
        IdTransp = formBuscar.IniciarBusca("Transportadoras")
    End If
    If Trim(IdTransp) = 0 Then Exit Sub
    cboTranspNome.Text = pgDadosTransportadora(IdTransp).Nome
    txtTranspCNPJ.Text = pgDadosTransportadora(IdTransp).CNPJ
    
End Sub


Private Function ValidarProdutos() As Boolean
    Dim i As Integer
    If chkMovEstoque.Value = False Then
        ValidarProdutos = True
        Exit Function
    End If
    With msfgProdutos
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = 0 Then
                MsgBox "O item n. " & ZE(i, 3) & " não foi identificado junto ao estoque da empresa. Favor verificar!", vbInformation, "Aviso"
                ValidarProdutos = False
                Exit Function
            End If
        Next
        ValidarProdutos = True
    End With
End Function
Private Sub ArmazenarNFe()
    'ide_dEmi
    Dim fileXMLDestino As String
    On Error GoTo TrtErroFile
    If Trim(fileXMLOrigem) = "" Then Exit Sub
    fileXMLDestino = PgDadosConfig.pXMLFornecedor & "\" & Format(ide_dEmi, "YYYYMM")
    If Dir(fileXMLDestino, vbDirectory) = "" Then
        MkDir fileXMLDestino
    End If
    fileXMLDestino = fileXMLDestino & "\NFe" & Id & ".xml"
    FileCopy fileXMLOrigem, fileXMLDestino
    Exit Sub
TrtErroFile:
    MsgBox "Erro ao armazenar o XML." & vbCrLf & _
    Err.Description, vbInformation, "Aviso - Erro n.: " & Err.Number
    Resume Next
End Sub
Private Sub RemoverNFe()
    Dim fileXMLDestino As String
    On Error GoTo TrtErroFile
    fileXMLDestino = PgDadosConfig.pXMLFornecedor & "\" & Format(ide_dEmi, "YYYYMM") & "\NFe" & Id & ".xml"
    If Dir(fileXMLDestino) = "" Then Exit Sub
    Kill fileXMLDestino 'PgDadosConfig.pXMLFornecedor & "\NFe" & Id & ".xml"
    Exit Sub
TrtErroFile:
    MsgBox "Erro ao armazenar o XML." & vbCrLf & _
    Err.Description, vbInformation, "Aviso - Erro n.: " & Err.Number
    Resume Next
End Sub
Private Function ValidarDados() As Boolean
    Dim i       As Integer
    Dim Rst     As Recordset
    Dim sSQL    As String
    ValidarDados = False
    
    
    'Checa se a nNF foi digitada
    If Trim(ide_nNF) = "" Or Trim(txtnNF.Text) = "" Then
        MsgBox "Favor informar o numero da nota fiscal!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
     'Fornecedor
    If Trim(emit_CNPJ) <> "" Then
            sSQL = "SELECT * FROM Fornecedores WHERE Doc = '" & emit_CNPJ & "'"
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    If MsgBox("Fornecedor não registrada no sistema. Deseja cadastrar agora?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                        formFornecedores.ReceberDadosFornecedores emit_CNPJ, emit_xNome, emit_IE, emit_IEST, emit_IM, emit_CNAE, , , , , emit_xLgr, _
                                                          emit_nro, emit_xCpl, emit_Bairro, emit_UF, emit_xMun, _
                                                          emit_CEP, , emit_fone
                        formFornecedores.Show
                    End If
                    ValidarDados = False
                    Exit Function
                Else
                    Rst.MoveFirst
                    emit_id = Rst.Fields("Id")
            End If
            Rst.Close
        Else
            MsgBox "Selecione um Fornecedor.", vbInformation, "Aviso"
            ValidarDados = False
            Exit Function
    End If
    
    'Checa se a nota ja foi cadastrada
    If nfeCadastrada = True Then
        MsgBox "Nota fiscal ja registrada no sistema.", vbInformation, App.EXEName
        ValidarDados = False
        Exit Function
    End If

   
     'Transportadora
    If Trim(transp_CNPJ) <> "" Then
        sSQL = "SELECT * FROM Transportadoras WHERE CNPJ = '" & transp_CNPJ & "'"
        Set Rst = RegistroBuscar(sSQL)
        If Rst.BOF And Rst.EOF Then
                If MsgBox("Transportadora não registrada no sistema. Deseja cadastrar agora?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                        formTransportadoras.ReceberDadosTransportadora transp_CNPJ, transp_xNome, transp_IE, transp_xEnder, transp_UF, transp_xMun
                        formTransportadoras.Show
                        ValidarDados = False
                        Exit Function
                    Else
                        ValidarDados = True
                End If
            Else
                ValidarDados = True
        End If
        Rst.Close
    End If
    
    'Produtos
    'For i = 1 To msfgProdutos.Rows - 1
    '    If msfgProdutos.TextMatrix(i, 0) = "0" Then
    '        MsgBox "Favor vincular os produtos da NFe ao Estoque.", vbInformation, "Aviso"
    '        ValidarDados = False
    '        Exit Function
    '    End If
    'Next
    'Duplicatas
    For i = 1 To msfgDupl.Rows - 1
        If Trim(msfgDupl.TextMatrix(i, 4)) = "" Or Trim(msfgDupl.TextMatrix(i, 5)) = "" Then
            MsgBox "Favor vincular as Duplicatas a um centro de custos.", vbInformation, "Aviso"
            ValidarDados = False
            Exit Function
        End If
    Next
    ValidarDados = True
End Function
Private Sub MontarArrayBaseEstoque(nProd As Integer)
'id_intProd|det_cProd|det_cEAN|det_xProd|EXTIPI|det_NCM|det_CFOP|det_uCom|
'det_qCom|det_vUnCom|det_vProd|det_cEANTrib|det_uTrib|det_qTrib|
'det_vUnTrib|det_vFrete|det_vSeg|det_vDesc|det_vOutro|det_indTot|xPed|nItemPed
'det_indTot = 0 = O valor do item compoe a NF / 1  = O valor do item nao compoe a NF
    aItem(lnProd) = Array(nProd, _
                        "", _
                        "", _
                        Trim(txtDescricao.Text), _
                        "", _
                        Trim(txtNCM.Text), _
                        Trim(txtCFOP.Text), _
                        Trim(cboUnidade.Text), _
                        Trim(txtQtd.Text), _
                        Trim(txtvUnit.Text), _
                        Trim(txtvProd.Text), _
                        "", _
                        Trim(cboUnidade.Text), _
                        Trim(txtQtd.Text), _
                        Trim(txtvUnit.Text), _
                        "0.00", _
                        "0.00", _
                        "0.00", _
                        "0.00", _
                        "1", _
                        "0", _
                        "0")
    aEstoque(lnProd) = Array("0", aItem(lnProd)(8), "0")
    '******* ICMS ****************
    'Origem|CST|ModBC|pRedBC|vBC|pICMS|vICMS|modBCST|pMVAST|pRedBCST|vBCST|pICMSST|vICMSST
    aICMS(lnProd) = Array(0, Trim(txtCST.Text), 0, 0, Trim(txtBCICMSp.Text), Trim(txtpICMSp.Text), Trim(txtvICMSp.Text), 0, 0, 0, 0, 0, 0, 0, 0, 0)
        'vReg(cReg) = Array("ICMS_CSOSN", aICMS(i)(13), "S"): cReg = cReg + 1
        'vReg(cReg) = Array("ICMS_pCredSN", aICMS(i)(14), "S"): cReg = cReg + 1
        'vReg(cReg) = Array("ICMS_vCredICMSSN", aICMS(i)(15), "S"): cReg = cReg + 1
    '********* IPI ***************
    'cEnq|CST|vBC|pIPI|vIPI
    aIPI(lnProd) = Array(0, 0, Trim(txtvProd.Text), Trim(txtpIPIp.Text), Trim(txtvIPIp.Text))
    '******** PIS / COFINS ****************
    'CST|vBC|pCOFINS|vCOFINS
    aPIS(lnProd) = Array(0, 0, 0, 0)
    aCOFINS(lnProd) = Array(0, 0, 0, 0)
    cItens = msfgProdutos.Rows - 1
End Sub
Private Sub MontarArrayTotais()
    total_vBC = ChkVal(txtBCICMS.Text, 0, cDecMoeda)
    total_vICMS = ChkVal(txtvICMS.Text, 0, cDecMoeda)
    total_vICMSST = ChkVal(txtvICMSST.Text, 0, cDecMoeda)
    total_vBCST = ChkVal(txtBCICMSST.Text, 0, cDecMoeda)
    total_vCredICMSSN = ChkVal(txtvCredICMSSN.Text, 0, cDecMoeda)
    
    total_vProd = ChkVal(txtvProduto.Text, 0, cDecMoeda)
    total_vFrete = ChkVal(txtvFrete.Text, 0, cDecMoeda)
    total_vSeg = ChkVal(txtvSeguro.Text, 0, cDecMoeda)
    total_vDesc = ChkVal(txtvDesconto.Text, 0, cDecMoeda)
    total_vOutro = ChkVal(txtvOutras.Text, 0, cDecMoeda)
    total_vIPI = ChkVal(txtvIPI.Text, 0, cDecMoeda)
    total_vPIS = ChkVal(txtvPIS.Text, 0, cDecMoeda)
    total_vCOFINS = ChkVal(txtvCOFINS.Text, 0, cDecMoeda)
    total_vNF = ChkVal(txtvTotalNF.Text, 0, cDecMoeda)
      
End Sub
Private Sub calcTotais()
    Dim vProd   As String
    Dim vBCICMS As String
    Dim vICMS   As String
    Dim vIPI    As String
    Dim vNF     As String
    Dim vPIS    As String
    Dim vCOFINS As String
    Dim vCalc   As String 'Armazena os calculos variaveir para auxiliar o resultado (TMP)
    vProd = "0"
    vBCICMS = "0"
    vICMS = "0"
    vIPI = "0"
    vNF = "0"
    vPIS = "0"
    vCOFINS = "0"
    
    If chkTotaisAutomatico.Value = 0 Then
        txtBCICMS.Enabled = True
        txtvICMS.Enabled = True
        txtvProduto.Enabled = True
        txtvIPI.Enabled = True
        txtvPIS.Enabled = True
        txtvCOFINS.Enabled = True
    
        txtvTotalNF.Enabled = True
        Exit Sub
    End If
    Dim i As Integer
    With msfgProdutos
        For i = 1 To .Rows - 1
            vProd = Val(ChkVal(vProd, 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 9), 0, cDecMoeda))
            vBCICMS = Val(ChkVal(vBCICMS, 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 10), 0, cDecMoeda))
            vICMS = Val(ChkVal(vICMS, 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 11), 0, cDecMoeda))
            vIPI = Val(ChkVal(vIPI, 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 12), 0, cDecMoeda))
            
            vCalc = Val(ChkVal(.TextMatrix(i, 9), 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 12), 0, cDecMoeda))
            vCalc = (Val(ChkVal(vCalc, 0, cDecMoeda)) * Val(ChkVal(PgDadosEmpresa(ID_Empresa).PISAliquota, 0, cDecMoeda))) / 100
            
            vPIS = Val(ChkVal(vPIS, 0, cDecMoeda)) + Val(ChkVal(vCalc, 0, cDecMoeda))
            
            vCalc = Val(ChkVal(.TextMatrix(i, 9), 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 12), 0, cDecMoeda))
            vCalc = (Val(ChkVal(vCalc, 0, cDecMoeda)) * Val(ChkVal(PgDadosEmpresa(ID_Empresa).COFINSAliquota, 0, cDecMoeda))) / 100
            
            vCOFINS = Val(ChkVal(vCOFINS, 0, cDecMoeda)) + Val(ChkVal(vCalc, 0, cDecMoeda))
            
        Next
    End With
    'vNF = Val(ChkVal(vNF, 0, cDecMoeda)) + Val(ChkVal(vProd, 0, cDecMoeda)) + Val(ChkVal(vIPI, 0, cDecMoeda))
    vNF = Val(ChkVal(vProd, 0, cDecMoeda)) + Val(ChkVal(vIPI, 0, cDecMoeda))
    'ST
    vNF = Val(ChkVal(vNF, 0, cDecMoeda)) + Val(ChkVal(Trim(txtvICMSST.Text), 0, cDecMoeda))
    'Frete
    vNF = Val(ChkVal(vNF, 0, cDecMoeda)) + Val(ChkVal(Trim(txtvFrete.Text), 0, cDecMoeda))
    'Seguro
    vNF = Val(ChkVal(vNF, 0, cDecMoeda)) + Val(ChkVal(Trim(txtvSeguro.Text), 0, cDecMoeda))
    'Desconto
    vNF = Val(ChkVal(vNF, 0, cDecMoeda)) - Val(ChkVal(Trim(txtvDesconto.Text), 0, cDecMoeda))
    'Outras Desp
    vNF = Val(ChkVal(vNF, 0, cDecMoeda)) - Val(ChkVal(Trim(txtvOutras.Text), 0, cDecMoeda))
 
    txtBCICMS.Enabled = False
    txtvICMS.Enabled = False
    txtvProduto.Enabled = False
    txtvIPI.Enabled = False
    txtvPIS.Enabled = False
    txtvCOFINS.Enabled = False
    
    txtvTotalNF.Enabled = False
    
    
    txtBCICMS.Text = ConvMoeda(vBCICMS)
    txtvICMS.Text = ConvMoeda(vICMS)
    txtvProduto.Text = ConvMoeda(vProd)
    txtvIPI.Text = ConvMoeda(vIPI)
    txtvPIS.Text = ConvMoeda(vPIS)
    txtvCOFINS.Text = ConvMoeda(vCOFINS)
    txtvTotalNF.Text = ConvMoeda(vNF)
    MontarArrayTotais
  
End Sub
Private Function PgIDMaterial(cFab As String, sCNPJ As String) As Long
    Dim sSQL    As String
    Dim Rst     As Recordset
    sSQL = "SELECT FNFe.emit_CNPJ, FNFe.idNFe, FNFeI.idNFe, FNFeI.det_cProd, FNFeI.det_idproduto " & _
           "FROM FaturamentoNFeEntrada AS FNFe, FaturamentoNFeEntradaItens AS FNFeI " & _
           "WHERE FNFe.emit_CNPJ = '" & sCNPJ & "' AND FNFe.idNFe=FNFeI.idNFe AND det_cProd = '" & Trim(cFab) & "' " & _
           "ORDER BY FNFeI.ID"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            PgIDMaterial = 0
        Else
            Rst.MoveLast
            PgIDMaterial = Rst.Fields("det_idProduto")
            '*** 12.11.2012
            '*** Valida o id do produto quanto esta ativo e no mesmo deposito
            If LCase(pgDadosEstoqueProduto(Rst.Fields("det_idProduto")).status) <> "ativo" Then
                PgIDMaterial = 0
            End If
            If pgDadosEstoqueProduto(Rst.Fields("det_idProduto")).IdDeposito <> ID_Deposito Then
                PgIDMaterial = 0
            End If
    End If
    Rst.Close
End Function
