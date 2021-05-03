VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFornecedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fornecedores"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9885
   Begin TabDlg.SSTab SSTab 
      Height          =   4215
      Left            =   60
      TabIndex        =   11
      Top             =   1800
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   7435
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dados Cadastrais"
      TabPicture(0)   =   "formFornecedores.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "formFornecedores.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame10 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   101
         Top             =   60
         Width           =   9435
         Begin VB.TextBox txtAno 
            Height          =   285
            Left            =   480
            MaxLength       =   4
            TabIndex        =   104
            Text            =   "Text1"
            Top             =   180
            Width           =   795
         End
         Begin VB.CommandButton btoAtualizarHist 
            Height          =   375
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   120
            Width           =   375
         End
         Begin MSFlexGridLib.MSFlexGrid msfgHist 
            Height          =   3075
            Left            =   120
            TabIndex        =   102
            Top             =   540
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   5424
            _Version        =   393216
            Cols            =   6
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            FormatString    =   "^id|^Chave de Acesso         |^Emissão            |^Serie     |^Num.Nota               |>Valor                              "
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "Ano:"
            Height          =   195
            Left            =   60
            TabIndex        =   105
            Top             =   240
            Width           =   375
         End
      End
      Begin TabDlg.SSTab SSt 
         Height          =   3555
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   6271
         _Version        =   393216
         Tabs            =   5
         Tab             =   1
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Basico"
         TabPicture(0)   =   "formFornecedores.frx":0038
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame3"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Entrega"
         TabPicture(1)   =   "formFornecedores.frx":0054
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame4"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame9"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Cobrança"
         TabPicture(2)   =   "formFornecedores.frx":0070
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame5"
         Tab(2).Control(1)=   "Frame11"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Contatos"
         TabPicture(3)   =   "formFornecedores.frx":008C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame12"
         Tab(3).Control(1)=   "Frame13"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Outros"
         TabPicture(4)   =   "formFornecedores.frx":00A8
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame7"
         Tab(4).Control(1)=   "Frame6"
         Tab(4).Control(2)=   "Frame8"
         Tab(4).ControlCount=   3
         Begin VB.Frame Frame13 
            Height          =   2595
            Left            =   -69840
            TabIndex        =   99
            Top             =   480
            Width           =   3975
            Begin VB.Label Label41 
               Caption         =   "colocar nome / tel  / email / obs"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   240
               TabIndex        =   100
               Top             =   240
               Width           =   3315
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Observações:"
            Height          =   915
            Left            =   -74940
            TabIndex        =   97
            Top             =   360
            Width           =   9015
            Begin VB.TextBox txtObs 
               Height          =   495
               Left            =   180
               TabIndex        =   98
               Text            =   "Text11"
               Top             =   240
               Width           =   8715
            End
         End
         Begin VB.Frame Frame6 
            Height          =   795
            Left            =   -74940
            TabIndex        =   94
            Top             =   1380
            Width           =   2835
            Begin VB.CheckBox chkME 
               Caption         =   "Micro Empresa (M.E.)"
               Height          =   255
               Left            =   120
               TabIndex        =   96
               Top             =   480
               Width           =   2595
            End
            Begin VB.CheckBox chkSimplesNacional 
               Caption         =   "Optante do Simples Nacional"
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame Frame7 
            Height          =   1995
            Left            =   -71160
            TabIndex        =   85
            Top             =   1320
            Width           =   4095
            Begin VB.TextBox txtIE 
               Height          =   285
               Left            =   1560
               TabIndex        =   89
               Text            =   "Text1"
               Top             =   300
               Width           =   2235
            End
            Begin VB.TextBox txtIEST 
               Height          =   285
               Left            =   1560
               TabIndex        =   88
               Text            =   "Text1"
               Top             =   720
               Width           =   2235
            End
            Begin VB.TextBox txtIM 
               Height          =   285
               Left            =   1560
               TabIndex        =   87
               Text            =   "Text1"
               Top             =   1140
               Width           =   2235
            End
            Begin VB.TextBox txtCNAE 
               Height          =   285
               Left            =   1560
               TabIndex        =   86
               Text            =   "Text1"
               Top             =   1560
               Width           =   2235
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               Caption         =   "Insc. Estadual/RG:"
               Height          =   195
               Left            =   60
               TabIndex        =   93
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               Caption         =   "Insc. Estadual ST:"
               Height          =   195
               Left            =   180
               TabIndex        =   92
               Top             =   780
               Width           =   1335
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               Caption         =   "Insc. Municipal:"
               Height          =   195
               Left            =   300
               TabIndex        =   91
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               Caption         =   "CNAE:"
               Height          =   195
               Left            =   540
               TabIndex        =   90
               Top             =   1620
               Width           =   975
            End
         End
         Begin VB.Frame Frame12 
            Height          =   2115
            Left            =   -74700
            TabIndex        =   76
            Top             =   540
            Width           =   3675
            Begin VB.TextBox txtEmailNFe 
               Height          =   285
               Left            =   1200
               TabIndex        =   80
               Text            =   "Text1"
               Top             =   1380
               Width           =   2055
            End
            Begin VB.TextBox txtEmailFin 
               Height          =   285
               Left            =   1200
               TabIndex        =   79
               Text            =   "Text1"
               Top             =   1020
               Width           =   2055
            End
            Begin VB.TextBox txtEmailCom 
               Height          =   285
               Left            =   1200
               TabIndex        =   78
               Text            =   "Text1"
               Top             =   660
               Width           =   2055
            End
            Begin VB.TextBox txtWebSite 
               Height          =   285
               Left            =   1200
               TabIndex        =   77
               Text            =   "Text1"
               Top             =   300
               Width           =   2055
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               Caption         =   "e-mail NF-e:"
               Height          =   195
               Left            =   240
               TabIndex        =   84
               Top             =   1500
               Width           =   915
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               Caption         =   "e-mail Financ.:"
               Height          =   195
               Left            =   120
               TabIndex        =   83
               Top             =   1080
               Width           =   1035
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               Caption         =   "e-mail Com.:"
               Height          =   255
               Left            =   300
               TabIndex        =   82
               Top             =   660
               Width           =   855
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               Caption         =   "Web site:"
               Height          =   195
               Left            =   300
               TabIndex        =   81
               Top             =   300
               Width           =   855
            End
         End
         Begin VB.Frame Frame11 
            Height          =   1455
            Left            =   -74760
            TabIndex        =   67
            Top             =   1980
            Width           =   8955
            Begin VB.TextBox txtLimiteCredito 
               Height          =   285
               Left            =   6300
               TabIndex        =   71
               Text            =   "Text12"
               Top             =   840
               Width           =   1815
            End
            Begin VB.ComboBox cboTipoDocumento 
               Height          =   315
               Left            =   5940
               Style           =   2  'Dropdown List
               TabIndex        =   70
               Top             =   360
               Width           =   2715
            End
            Begin VB.ComboBox cboCondicoesPagamento 
               Height          =   315
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   960
               Width           =   2595
            End
            Begin VB.ComboBox cboPlanoContas 
               Height          =   315
               Left            =   2040
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   180
               Width           =   3315
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               Caption         =   "Limite Credito:"
               Height          =   195
               Left            =   5100
               TabIndex        =   75
               Top             =   900
               Width           =   1155
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "Tipo de Documento:"
               Height          =   195
               Left            =   5640
               TabIndex        =   74
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               Caption         =   "Condições de Pagamento:"
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   1020
               Width           =   1875
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               Caption         =   "Plano de Contas:"
               Height          =   255
               Left            =   780
               TabIndex        =   72
               Top             =   240
               Width           =   1215
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
            Height          =   1455
            Left            =   120
            TabIndex        =   64
            Top             =   1920
            Width           =   9075
            Begin VB.ComboBox cboTransportadora 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   240
               Width           =   7335
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               Caption         =   "Transportadora:"
               Height          =   195
               Left            =   120
               TabIndex        =   66
               Top             =   300
               Width           =   1155
            End
         End
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   -74880
            TabIndex        =   48
            Top             =   420
            Width           =   9075
            Begin VB.TextBox txtCobrancaLgr 
               Height          =   285
               Left            =   900
               MaxLength       =   60
               TabIndex        =   56
               Text            =   "Text1"
               Top             =   240
               Width           =   6075
            End
            Begin VB.TextBox txtCobrancaNro 
               Height          =   285
               Left            =   7920
               MaxLength       =   60
               TabIndex        =   55
               Text            =   "Text1"
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox txtCobrancaCpl 
               Height          =   285
               Left            =   900
               MaxLength       =   60
               TabIndex        =   54
               Text            =   "Text1"
               Top             =   600
               Width           =   2715
            End
            Begin VB.TextBox txtCobrancaBairro 
               Height          =   285
               Left            =   4620
               MaxLength       =   60
               TabIndex        =   53
               Text            =   "Text1"
               Top             =   600
               Width           =   2955
            End
            Begin VB.ComboBox cboCobrancaUF 
               Height          =   315
               Left            =   900
               Style           =   2  'Dropdown List
               TabIndex        =   52
               Top             =   960
               Width           =   915
            End
            Begin VB.ComboBox cboCobrancaMun 
               Height          =   315
               Left            =   2940
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   960
               Width           =   3135
            End
            Begin VB.TextBox txtCobrancaCEP 
               Height          =   285
               Left            =   7380
               MaxLength       =   8
               TabIndex        =   50
               Text            =   "Text1"
               Top             =   960
               Width           =   1515
            End
            Begin VB.CheckBox chkCobranca 
               Caption         =   "Local de COBRANÇA diferente do emitente."
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
               TabIndex        =   49
               Top             =   0
               Width           =   4155
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               Caption         =   "Municipio:"
               Height          =   255
               Left            =   2040
               TabIndex        =   63
               Top             =   1020
               Width           =   795
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               Caption         =   "Endereço:"
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "Número:"
               Height          =   255
               Left            =   7020
               TabIndex        =   61
               Top             =   300
               Width           =   855
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               Caption         =   "Compl.:"
               Height          =   195
               Left            =   300
               TabIndex        =   60
               Top             =   660
               Width           =   555
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               Caption         =   "Bairro:"
               Height          =   195
               Left            =   4080
               TabIndex        =   59
               Top             =   660
               Width           =   495
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               Caption         =   "UF:"
               Height          =   195
               Left            =   360
               TabIndex        =   58
               Top             =   1020
               Width           =   495
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               Caption         =   "CEP:"
               Height          =   195
               Left            =   6840
               TabIndex        =   57
               Top             =   1020
               Width           =   495
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
            Height          =   1455
            Left            =   120
            TabIndex        =   32
            Top             =   420
            Width           =   9075
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
               TabIndex        =   40
               Top             =   0
               Width           =   3975
            End
            Begin VB.TextBox txtEntregaLgr 
               Height          =   285
               Left            =   900
               MaxLength       =   60
               TabIndex        =   39
               Text            =   "Text1"
               Top             =   240
               Width           =   6075
            End
            Begin VB.TextBox txtEntregaNro 
               Height          =   285
               Left            =   7920
               MaxLength       =   60
               TabIndex        =   38
               Text            =   "Text1"
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox txtEntregaCpl 
               Height          =   285
               Left            =   900
               MaxLength       =   60
               TabIndex        =   37
               Text            =   "Text1"
               Top             =   600
               Width           =   2715
            End
            Begin VB.TextBox txtEntregaBairro 
               Height          =   285
               Left            =   4620
               MaxLength       =   60
               TabIndex        =   36
               Text            =   "Text1"
               Top             =   600
               Width           =   2955
            End
            Begin VB.ComboBox cboEntregaUF 
               Height          =   315
               Left            =   900
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   960
               Width           =   915
            End
            Begin VB.ComboBox cboEntregaMun 
               Height          =   315
               Left            =   2940
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   960
               Width           =   3135
            End
            Begin VB.TextBox txtEntregaCEP 
               Height          =   285
               Left            =   7380
               MaxLength       =   8
               TabIndex        =   33
               Text            =   "Text1"
               Top             =   960
               Width           =   1515
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Municipio:"
               Height          =   255
               Left            =   2040
               TabIndex        =   47
               Top             =   1020
               Width           =   795
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "Endereço:"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "Número:"
               Height          =   255
               Left            =   7020
               TabIndex        =   45
               Top             =   300
               Width           =   855
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Compl.:"
               Height          =   195
               Left            =   300
               TabIndex        =   44
               Top             =   660
               Width           =   555
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "Bairro:"
               Height          =   195
               Left            =   4080
               TabIndex        =   43
               Top             =   660
               Width           =   495
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "UF:"
               Height          =   195
               Left            =   360
               TabIndex        =   42
               Top             =   1020
               Width           =   495
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "CEP:"
               Height          =   195
               Left            =   6840
               TabIndex        =   41
               Top             =   1020
               Width           =   495
            End
         End
         Begin VB.Frame Frame3 
            Height          =   3015
            Left            =   -74880
            TabIndex        =   13
            Top             =   360
            Width           =   9015
            Begin VB.TextBox txtLgr 
               Height          =   285
               Left            =   1140
               MaxLength       =   60
               TabIndex        =   22
               Text            =   "Text1"
               Top             =   240
               Width           =   6735
            End
            Begin VB.TextBox txtNro 
               Height          =   285
               Left            =   1140
               MaxLength       =   60
               TabIndex        =   21
               Text            =   "Text1"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox txtCpl 
               Height          =   285
               Left            =   3180
               MaxLength       =   60
               TabIndex        =   20
               Text            =   "Text1"
               Top             =   600
               Width           =   4695
            End
            Begin VB.TextBox txtBairro 
               Height          =   285
               Left            =   1140
               MaxLength       =   60
               TabIndex        =   19
               Text            =   "Text1"
               Top             =   1020
               Width           =   2955
            End
            Begin VB.ComboBox cboUF 
               Height          =   315
               Left            =   1140
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   1380
               Width           =   915
            End
            Begin VB.ComboBox cboMun 
               Height          =   315
               Left            =   3180
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   1380
               Width           =   2655
            End
            Begin VB.TextBox txtCEP 
               Height          =   285
               Left            =   1080
               MaxLength       =   8
               TabIndex        =   16
               Text            =   "Text1"
               Top             =   1800
               Width           =   2175
            End
            Begin VB.TextBox txtMail 
               Height          =   285
               Left            =   1080
               TabIndex        =   15
               Text            =   "Text1"
               Top             =   2160
               Width           =   3915
            End
            Begin VB.TextBox txtFone 
               Height          =   315
               Left            =   1080
               MaxLength       =   10
               TabIndex        =   14
               Text            =   "Text1"
               Top             =   2520
               Width           =   2955
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Endereço:"
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   300
               Width           =   855
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Número:"
               Height          =   255
               Left            =   240
               TabIndex        =   30
               Top             =   660
               Width           =   855
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "Complemento:"
               Height          =   195
               Left            =   2100
               TabIndex        =   29
               Top             =   660
               Width           =   1035
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Bairro:"
               Height          =   255
               Left            =   360
               TabIndex        =   28
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Municipio:"
               Height          =   255
               Left            =   2340
               TabIndex        =   27
               Top             =   1440
               Width           =   795
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               Caption         =   "UF:"
               Height          =   195
               Left            =   600
               TabIndex        =   26
               Top             =   1440
               Width           =   495
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   "CEP:"
               Height          =   195
               Left            =   540
               TabIndex        =   25
               Top             =   1860
               Width           =   495
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Caption         =   "e-mail:"
               Height          =   195
               Left            =   600
               TabIndex        =   24
               Top             =   2220
               Width           =   435
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Caption         =   "Telefone:"
               Height          =   195
               Left            =   360
               TabIndex        =   23
               Top             =   2520
               Width           =   675
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   60
      TabIndex        =   4
      Top             =   480
      Width           =   7275
      Begin VB.TextBox txtxNome 
         Height          =   285
         Left            =   1260
         MaxLength       =   60
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox txtDoc 
         Height          =   285
         Left            =   1260
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   180
         Width           =   2295
      End
      Begin VB.TextBox txtFant 
         Height          =   285
         Left            =   1260
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   5895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   195
         Left            =   540
         TabIndex        =   10
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome Fantasia:"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ/CPF:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   7440
      TabIndex        =   1
      Top             =   480
      Width           =   2355
      Begin VB.ComboBox cboPessoa 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1995
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   1995
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
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
               Picture         =   "formFornecedores.frx":00C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFornecedores.frx":0516
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFornecedores.frx":0830
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFornecedores.frx":10C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFornecedores.frx":2314
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFornecedores.frx":2BEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFornecedores.frx":3480
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFornecedores.frx":3D12
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFornecedores.frx":4F64
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFornecedores.frx":527E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFornecedores.frx":5598
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg As Integer
Dim strTabela As String

Private Sub PesquisarRegistro()
    ''Dim idreg  As String
    IdReg = formBuscar.IniciarBusca(strTabela)
    ''IdReg = IIf(idreg = "", 0, idreg)
    
    If IdReg = 0 Then
            LimpaFormulario Me
        Else
            MostrarDados
    End If
End Sub
Private Sub txtAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ListHistoricoNotasFiscais
    End If
    KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub btoAtualizarHist_Click()
    ListHistoricoNotasFiscais
    
        
End Sub





Private Sub cboCobrancaMun_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    If Trim(cboCobrancaUF.Text) = "" Then
        MsgBox "Selecione uma Unidade Federal (UF)."
        Exit Sub
    End If
    sSQL = "SELECT * FROM TributacaoMunicipio WHERE codUF = " & pgDadosICMS(cboCobrancaUF.Text, 0).Id & " ORDER BY Descricao"
    cboCobrancaMun.Clear
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCobrancaMun.AddItem Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboCobrancaUF_DropDown()
    Dim Rst As Recordset
    cboCobrancaUF.Clear
    Set Rst = RegistroBuscar("SELECT * FROM TributacaoUF ORDER BY UF")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCobrancaUF.AddItem Rst.Fields("UF")
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
    sSQL = "SELECT * FROM TributacaoMunicipio WHERE codUF = " & pgDadosICMS(cboEntregaUF.Text, 0).Id & " ORDER BY Descricao"
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
    Set Rst = RegistroBuscar("SELECT * FROM TributacaoUF ORDER BY UF")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboEntregaUF.AddItem Rst.Fields("UF")
                Rst.MoveNext
            Loop
    End If

End Sub



Private Sub cboMun_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    If Trim(cboUF.Text) = "" Then
        MsgBox "Selecione uma Unidade Federal (UF)."
        Exit Sub
    End If
    sSQL = "SELECT * FROM TributacaoMunicipio WHERE UF = '" & Trim(UCase(cboUF.Text)) & "' ORDER BY Descricao"
    cboMun.Clear
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboMun.AddItem UCase(Rst.Fields("Descricao"))
                Rst.MoveNext
            Loop
    End If
End Sub

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



Private Sub cboStatus_DropDown()
    cboStatus.Clear
    cboStatus.AddItem "Ativo"
    cboStatus.AddItem "Inativo"
End Sub

Private Sub cbotransportadora_DropDown()
    Dim Rst As Recordset
    Set Rst = RegistroBuscar("SELECT * FROM Transportadoras WHERE ID_Empresa = " & ID_Empresa)
    If Rst.BOF And Rst.EOF Then
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboTransportadora.AddItem ZE(Rst.Fields("id"), 6) & " - " & Rst.Fields("xNome")
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

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    IdReg = 0
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    SSt.Tab = 0
    HDFormulario False
    HDMenu Me, True
    SSTab.Tab = 0
    
    
End Sub

Private Sub ListHistoricoNotasFiscais()
    Dim sSQL    As String
    Dim Rst     As Recordset
    msfgHist.Rows = 1
    If IdReg = 0 Then
        Exit Sub
    End If
    
    If Trim(txtAno.Text) = "" Then
        txtAno.Text = Format(Date, "YYYY")
    End If

    If Val(txtAno.Text) < 2000 Then
        MsgBox "Ano (" & txtAno.Text & ") invalido!", vbInformation, App.EXEName
        txtAno.Text = Format(Date, "YYYY")
        Exit Sub
    End If

    sSQL = "SELECT * FROM FaturamentoNFeEntrada" & _
           " WHERE ID_Empresa = " & ID_Empresa & _
           " AND emit_CNPJ = '" & PgDadosFornecedor(IdReg).Doc & "'" & _
           " AND YEAR(ide_dEmi)>='" & txtAno.Text & "'"
           '" ORDER BY ide_nNF LIMIT 200"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            With msfgHist
                Do Until Rst.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("ID")
                    .TextMatrix(.Rows - 1, 1) = cNull(Rst.Fields("idNFe"))
                    .TextMatrix(.Rows - 1, 2) = Rst.Fields("ide_dEmi")
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("ide_Serie")), "", Rst.Fields("ide_Serie"))
                    .TextMatrix(.Rows - 1, 4) = IIf(IsNull(Rst.Fields("ide_nnf")), "", Rst.Fields("ide_nnf"))
                    .TextMatrix(.Rows - 1, 5) = ConvMoeda(ChkVal(IIf(IsNull(Rst.Fields("total_vNF")), "0", Rst.Fields("total_vNF")), 0, cDecMoeda))
                    Rst.MoveNext
                Loop
            End With
    End If
End Sub
Private Sub HDFormulario(op As Boolean)
    HDForm Me, op
    msfgHist.Enabled = IIf(op = True, False, True)
    txtDoc.Enabled = IIf(op = True, False, True)
    txtAno.Enabled = IIf(op = True, False, True)
    btoAtualizarHist.Enabled = IIf(op = True, False, True)
    
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




Private Sub msfgHist_DblClick()
    If msfgHist.TextMatrix(msfgHist.Row, 0) = "ID" Or msfgHist.Rows = 1 Then Exit Sub
    ImprimirDANFEFornecedor msfgHist.TextMatrix(msfgHist.Row, 1)

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
            
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDFormulario False
                'LimpaFormulario me
                'txtDoc.Enabled = True
            End If
            
        
        Case "Cancelar"
            HDMenu Me, True
            HDFormulario False
            LimpaFormulario Me
            txtDoc.Enabled = True
        
        Case "Manutenção da Tabela"
            formManutencaoTabelas.IniciarManutencao Me, "SELECT * FROM Clientes"
    End Select
End Sub

Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim i           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer
    cReg = 0
    If Len(Trim(cboUF.Text)) = 0 Then
        MsgBox "Selecione uma UF!", vbInformation, App.EXEName
        grvRegistro = False
        Exit Function
    End If
     If Len(Trim(cboMun.Text)) = 0 Then
        MsgBox "Selecione um MUNICIPIO!", vbInformation, App.EXEName
        grvRegistro = False
        Exit Function
    End If
    
    For i = 0 To Me.Controls.Count - 1
        Set Controle = Me.Controls(i)
        
        If TypeOf Controle Is TextBox Then
            If Controle.Name <> "txtAno" Then
                vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Text, "S")
                cReg = cReg + 1
            End If
        End If
        If TypeOf Controle Is ComboBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Text, "S")
            cReg = cReg + 1
        End If
        If TypeOf Controle Is CheckBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Value, "S")
            cReg = cReg + 1
        End If
        
    Next
   
    cReg = cReg - 1
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


'*************************
'  Dim Controle    As Control
'    Dim I           As Integer
'
'    For I = 0 To Formulario.Controls.Count - 1
'        Set Controle = Formulario.Controls(I)
'        If TypeOf Controle Is TextBox Then
'            Controle.Enabled = sModo
'        End If
'        If TypeOf Controle Is ComboBox Then
'            Controle.Enabled = sModo
'        End If
End Function
Private Sub txtDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
    
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        BuscarDados (txtDoc.Text)
    End If
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub
Private Sub BuscarDados(strCNPJ As String)
    Dim Rst     As ADODB.Recordset
    Dim strSQL  As String
    
    If Trim(strCNPJ) = "" Then Exit Sub
    
    SSt.Tab = 0
    
    strSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Doc = '" & strCNPJ & "'"

    Set Rst = RegistroBuscar(strSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum Registro encontrado"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            IdReg = Rst.Fields("Id")
            Rst.Close
            MostrarDados
    End If
    
    
    
    
End Sub
Private Sub MostrarDados()
    Dim sSQL As String
    sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & IdReg
    ExibirDados Me, sSQL
    SSTab.Tab = 0
    ListHistoricoNotasFiscais
    
End Sub


Public Sub ReceberDadosFornecedores(Doc As String, _
                                    Optional Nome As String, _
                                    Optional IE As String, _
                                    Optional iest As String, _
                                    Optional im As String, _
                                    Optional cnae As String, _
                                    Optional emailnfe As String, _
                                    Optional emailfin As String, _
                                    Optional emailcom As String, _
                                    Optional website As String, _
                                    Optional Lgr As String, _
                                    Optional Nro As String, _
                                    Optional Cpl As String, _
                                    Optional Bairro As String, _
                                    Optional uf As String, _
                                    Optional Mun As String, _
                                    Optional CEP As String, _
                                    Optional Mail As String, _
                                    Optional Fone As String)
    cboStatus.Clear
    cboStatus.AddItem "Ativo"
    cboStatus.Text = cboStatus.List(0)
    
    cboPessoa.Clear
    cboPessoa.AddItem IIf(Len(Doc) >= 14, "Juridica", "Fisica")
    cboPessoa.Text = cboPessoa.List(0)
    
    
    txtDoc.Text = Doc
    txtxNome.Text = Nome
    txtIE.Text = IE
    txtIEST.Text = iest
    txtIM.Text = im
    txtCNAE.Text = cnae
    txtEmailNFe.Text = emailnfe
    txtEmailFin.Text = emailfin
    txtEmailCom.Text = emailcom
    txtWebSite.Text = website
    txtLgr.Text = Lgr
    txtNro.Text = Nro
    txtCpl.Text = Cpl
    txtBairro.Text = Bairro
    cboUF.Clear
    If Trim(uf) <> "" Then
        cboUF.AddItem uf
        cboUF.Text = cboUF.List(0)
    End If
    If Trim(Mun) <> "" Then
        cboMun.AddItem Mun
        cboMun.Text = cboMun.List(0)
    End If
    txtCEP.Text = CEP
    txtMail.Text = Mail
    txtFone.Text = Fone
    
    IdReg = 0
    HDMenu Me, False
    HDFormulario True
End Sub
