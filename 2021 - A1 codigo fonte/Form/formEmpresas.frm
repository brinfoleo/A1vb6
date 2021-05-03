VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form formEmpresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresas"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   9135
   Begin TabDlg.SSTab sstEmpresa 
      Height          =   5355
      Left            =   60
      TabIndex        =   4
      Top             =   2040
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9446
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Basico"
      TabPicture(0)   =   "formEmpresas.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Documentos"
      TabPicture(1)   =   "formEmpresas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Logotipo"
      TabPicture(2)   =   "formEmpresas.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Contador"
      TabPicture(3)   =   "formEmpresas.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         Caption         =   "Dados do Contador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   8715
         Begin VB.TextBox txtContCodigoID 
            Height          =   285
            Left            =   60
            MaxLength       =   50
            TabIndex        =   94
            Text            =   "Text1"
            Top             =   4320
            Width           =   2295
         End
         Begin VB.TextBox txtContIM 
            Height          =   315
            Left            =   1500
            TabIndex        =   93
            Text            =   "Text1"
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox txtContIE 
            Height          =   285
            Left            =   120
            TabIndex        =   92
            Text            =   "Text1"
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox txtContCNPJ 
            Height          =   285
            Left            =   120
            TabIndex        =   91
            Text            =   "Text1"
            Top             =   2880
            Width           =   2235
         End
         Begin VB.ComboBox cboContMunicipio 
            Height          =   315
            Left            =   3360
            TabIndex        =   83
            Text            =   "Combo1"
            Top             =   1740
            Width           =   3555
         End
         Begin VB.TextBox txtContMail 
            Height          =   285
            Left            =   4980
            MaxLength       =   120
            TabIndex        =   76
            Text            =   "Text1"
            Top             =   2340
            Width           =   3555
         End
         Begin VB.TextBox txtContFone2 
            Height          =   285
            Left            =   2520
            MaxLength       =   15
            TabIndex        =   75
            Text            =   "Text1"
            Top             =   2340
            Width           =   2355
         End
         Begin VB.TextBox txtContFone1 
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   74
            Text            =   "Text1"
            Top             =   2340
            Width           =   2355
         End
         Begin VB.TextBox txtContCEP 
            Height          =   285
            Left            =   7020
            MaxLength       =   10
            TabIndex        =   73
            Text            =   "Text1"
            Top             =   1740
            Width           =   1515
         End
         Begin VB.ComboBox cboContUF 
            Height          =   315
            Left            =   2640
            TabIndex        =   72
            Text            =   "Combo1"
            Top             =   1740
            Width           =   615
         End
         Begin VB.TextBox txtContBairro 
            Height          =   285
            Left            =   120
            MaxLength       =   60
            TabIndex        =   71
            Text            =   "Text1"
            Top             =   1740
            Width           =   2415
         End
         Begin VB.TextBox txtContComplemento 
            Height          =   285
            Left            =   6720
            TabIndex        =   70
            Text            =   "Text1"
            Top             =   1140
            Width           =   1815
         End
         Begin VB.TextBox txtContNumero 
            Height          =   285
            Left            =   5280
            MaxLength       =   30
            TabIndex        =   69
            Text            =   "Text1"
            Top             =   1140
            Width           =   1395
         End
         Begin VB.TextBox txtContEndereco 
            Height          =   285
            Left            =   120
            MaxLength       =   120
            TabIndex        =   68
            Text            =   "Text1"
            Top             =   1140
            Width           =   5055
         End
         Begin VB.TextBox txtContNome 
            Height          =   285
            Left            =   120
            MaxLength       =   120
            TabIndex        =   67
            Text            =   "Text1"
            Top             =   540
            Width           =   8415
         End
         Begin VB.Frame Frame6 
            Caption         =   "Contador Responsável"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   2880
            TabIndex        =   61
            Top             =   2700
            Width           =   5715
            Begin VB.TextBox txtContRMail 
               Height          =   285
               Left            =   2160
               MaxLength       =   120
               TabIndex        =   81
               Text            =   "Text1"
               Top             =   1680
               Width           =   3435
            End
            Begin VB.TextBox txtContRFone 
               Height          =   285
               Left            =   120
               MaxLength       =   15
               TabIndex        =   80
               Text            =   "Text1"
               Top             =   1680
               Width           =   1935
            End
            Begin VB.TextBox txtContRCRC 
               Height          =   285
               Left            =   2880
               MaxLength       =   50
               TabIndex        =   79
               Text            =   "Text1"
               Top             =   1080
               Width           =   2715
            End
            Begin VB.TextBox txtContRCPF 
               Height          =   285
               Left            =   120
               MaxLength       =   50
               TabIndex        =   78
               Text            =   "Text1"
               Top             =   1080
               Width           =   2655
            End
            Begin VB.TextBox txtContRNome 
               Height          =   285
               Left            =   120
               MaxLength       =   120
               TabIndex        =   77
               Text            =   "Text1"
               Top             =   480
               Width           =   5475
            End
            Begin VB.Label Label37 
               Caption         =   "Email"
               Height          =   195
               Left            =   2220
               TabIndex        =   66
               Top             =   1440
               Width           =   555
            End
            Begin VB.Label Label36 
               Caption         =   "Telefone"
               Height          =   195
               Left            =   120
               TabIndex        =   65
               Top             =   1440
               Width           =   735
            End
            Begin VB.Label Label35 
               Caption         =   "CRC"
               Height          =   195
               Left            =   2880
               TabIndex        =   64
               Top             =   840
               Width           =   375
            End
            Begin VB.Label Label34 
               Caption         =   "CPF"
               Height          =   195
               Left            =   120
               TabIndex        =   63
               Top             =   840
               Width           =   435
            End
            Begin VB.Label Label33 
               Caption         =   "Nome"
               Height          =   195
               Left            =   120
               TabIndex        =   62
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Label Label38 
            Caption         =   "Codigo Identificador"
            Height          =   195
            Left            =   60
            TabIndex        =   95
            Top             =   4080
            Width           =   1515
         End
         Begin VB.Label Label43 
            Caption         =   "Insc. Municipal"
            Height          =   195
            Left            =   1500
            TabIndex        =   90
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label Label42 
            Caption         =   "Insc. Estadual"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   3240
            Width           =   1155
         End
         Begin VB.Label Label41 
            Caption         =   "CNPJ"
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   2640
            Width           =   435
         End
         Begin VB.Label Label27 
            Caption         =   "Municipio"
            Height          =   195
            Left            =   3360
            TabIndex        =   82
            Top             =   1500
            Width           =   1155
         End
         Begin VB.Label Label32 
            Caption         =   "Email"
            Height          =   195
            Left            =   4980
            TabIndex        =   60
            Top             =   2100
            Width           =   555
         End
         Begin VB.Label Label31 
            Caption         =   "Telefone 2 "
            Height          =   195
            Left            =   2520
            TabIndex        =   59
            Top             =   2100
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "Telefone 1"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   2100
            Width           =   1095
         End
         Begin VB.Label Label29 
            Caption         =   "CEP"
            Height          =   195
            Left            =   7020
            TabIndex        =   57
            Top             =   1500
            Width           =   795
         End
         Begin VB.Label Label28 
            Caption         =   "UF"
            Height          =   195
            Left            =   2640
            TabIndex        =   56
            Top             =   1500
            Width           =   375
         End
         Begin VB.Label Label26 
            Caption         =   "Bairro"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   1500
            Width           =   855
         End
         Begin VB.Label Label25 
            Caption         =   "Complemento"
            Height          =   195
            Left            =   6780
            TabIndex        =   54
            Top             =   900
            Width           =   975
         End
         Begin VB.Label Label24 
            Caption         =   "Numero"
            Height          =   195
            Left            =   5280
            TabIndex        =   53
            Top             =   900
            Width           =   795
         End
         Begin VB.Label Label23 
            Caption         =   "Endereço"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   900
            Width           =   1155
         End
         Begin VB.Label Label22 
            Caption         =   "Nome/Razão Social"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   300
            Width           =   1515
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   46
         Top             =   360
         Width           =   8715
         Begin MSComDlg.CommonDialog cmdLogo 
            Left            =   4680
            Top             =   3780
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtfLogotipo 
            Height          =   285
            Left            =   120
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   4200
            Width           =   7815
         End
         Begin VB.CommandButton btoBusca 
            Height          =   255
            Index           =   7
            Left            =   8040
            Picture         =   "formEmpresas.frx":0070
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   4230
            Width           =   315
         End
         Begin VB.PictureBox pctLogo 
            Height          =   3435
            Left            =   180
            ScaleHeight     =   3375
            ScaleWidth      =   8235
            TabIndex        =   47
            Top             =   240
            Width           =   8295
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   26
         Top             =   360
         Width           =   8715
         Begin VB.TextBox txtSuframa 
            Height          =   285
            Left            =   1560
            MaxLength       =   9
            TabIndex        =   87
            Text            =   "Text1"
            Top             =   3900
            Width           =   1815
         End
         Begin VB.ComboBox cboTipoAtividade 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   3480
            Width           =   4515
         End
         Begin VB.TextBox txtCOFINSAliquota 
            Height          =   285
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   45
            Text            =   "Text1"
            Top             =   3060
            Width           =   795
         End
         Begin VB.TextBox txtPISAliquota 
            Height          =   285
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   44
            Text            =   "Text1"
            Top             =   2580
            Width           =   795
         End
         Begin VB.ComboBox cboRegimeTrib 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1980
            Width           =   4575
         End
         Begin VB.TextBox txtCNAE 
            Height          =   285
            Left            =   1560
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   1560
            Width           =   2235
         End
         Begin VB.TextBox txtIM 
            Height          =   285
            Left            =   1560
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   1140
            Width           =   2235
         End
         Begin VB.TextBox txtIEST 
            Height          =   285
            Left            =   1560
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   720
            Width           =   2235
         End
         Begin VB.TextBox txtIE 
            Height          =   285
            Left            =   1560
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   300
            Width           =   2235
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            Caption         =   "SUFRAMA:"
            Height          =   195
            Left            =   300
            TabIndex        =   86
            Top             =   3900
            Width           =   1215
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo de Atividade:"
            Height          =   195
            Left            =   180
            TabIndex        =   85
            Top             =   3540
            Width           =   1335
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "COFINS(%):"
            Height          =   255
            Left            =   600
            TabIndex        =   43
            Top             =   3120
            Width           =   915
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "PIS(%):"
            Height          =   195
            Left            =   720
            TabIndex        =   42
            Top             =   2640
            Width           =   795
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Regime Tributário:"
            Height          =   255
            Left            =   180
            TabIndex        =   40
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "CNAE:"
            Height          =   195
            Left            =   540
            TabIndex        =   30
            Top             =   1620
            Width           =   975
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Insc. Municipal:"
            Height          =   195
            Left            =   300
            TabIndex        =   29
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Insc. Estadual ST:"
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Insc. Estadual:"
            Height          =   195
            Left            =   300
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   8715
         Begin VB.TextBox txtFone 
            Height          =   315
            Left            =   1140
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   3900
            Width           =   2955
         End
         Begin VB.TextBox txtMail 
            Height          =   285
            Left            =   1140
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   3540
            Width           =   3915
         End
         Begin VB.TextBox txtCEP 
            Height          =   285
            Left            =   1140
            MaxLength       =   8
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   2820
            Width           =   2175
         End
         Begin VB.ComboBox cboMun 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   2400
            Width           =   3015
         End
         Begin VB.ComboBox cboUF 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1980
            Width           =   915
         End
         Begin VB.TextBox txtBairro 
            Height          =   285
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   1560
            Width           =   2955
         End
         Begin VB.ComboBox cboPais 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1080
            Width           =   2955
         End
         Begin VB.TextBox txtCpl 
            Height          =   285
            Left            =   3180
            MaxLength       =   60
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   720
            Width           =   4695
         End
         Begin VB.TextBox txtNro 
            Height          =   285
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtLgr 
            Height          =   285
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   360
            Width           =   6735
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Telefone:"
            Height          =   195
            Left            =   420
            TabIndex        =   36
            Top             =   3900
            Width           =   675
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "e-mail:"
            Height          =   195
            Left            =   660
            TabIndex        =   35
            Top             =   3600
            Width           =   435
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "CEP:"
            Height          =   195
            Left            =   600
            TabIndex        =   13
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "UF:"
            Height          =   195
            Left            =   600
            TabIndex        =   12
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "País:"
            Height          =   195
            Left            =   480
            TabIndex        =   11
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Municipio:"
            Height          =   255
            Left            =   300
            TabIndex        =   10
            Top             =   2460
            Width           =   795
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Bairro:"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   1620
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Complemento:"
            Height          =   195
            Left            =   2100
            TabIndex        =   8
            Top             =   780
            Width           =   1035
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Número:"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   780
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Endereço:"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   420
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   8955
      Begin VB.TextBox txtFant 
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1020
         Width           =   5295
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   1320
         MaxLength       =   160
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   660
         Width           =   5295
      End
      Begin VB.TextBox txtCNPJ 
         Height          =   285
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   255
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome Fantasia:"
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Razão Social:"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   660
         Width           =   1095
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
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
               Picture         =   "formEmpresas.frx":03FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresas.frx":084C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresas.frx":0B66
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresas.frx":13F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresas.frx":264A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresas.frx":2F24
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresas.frx":37B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresas.frx":4048
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresas.frx":529A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresas.frx":55B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresas.frx":58CE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Pressione <F3> para consulta..."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   7440
      Width           =   8955
   End
End
Attribute VB_Name = "formEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg   As Integer
Dim strTabela   As String

Private Sub PesquisarRegistro()
    IdReg = formBuscar.IniciarBusca(strTabela)
    
    If IdReg = 0 Then
            LimpaFormulario Me
        Else
            MostrarDados
            cboRegimeTrib.Clear
            If Trim(PgDadosEmpresa(IdReg).RegimeTrib) <> "" Then
                cboRegimeTrib.AddItem PgDadosEmpresa(IdReg).RegimeTrib & " - " & PgDescrRegTrib(CInt(PgDadosEmpresa(IdReg).RegimeTrib))
                cboRegimeTrib.Text = cboRegimeTrib.List(0)
            End If
    End If
End Sub

 



Private Sub btoBusca_Click(Index As Integer)
    On Error GoTo TrtErro
    With cmdLogo
        .Filter = "Imagem JPEG|*.jpg"
        .ShowOpen
        pctLogo.Picture = LoadPicture(.filename)
        txtfLogotipo.Text = .filename
    End With
    Exit Sub
TrtErro:
    MsgBox Err.Description, vbInformation, Err.Number
End Sub







Private Sub cboContMunicipio_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    If Trim(cboContUF.Text) = "" Then
        MsgBox "Selecione uma Unidade Federal (UF).", vbInformation, App.EXEName
        Exit Sub
    End If
    sSQL = "SELECT * FROM TributacaoMunicipio WHERE UF = '" & Trim(UCase(cboContUF.Text)) & "' ORDER BY Descricao"
    cboContMunicipio.Clear
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboContMunicipio.AddItem UCase(Rst.Fields("Descricao"))
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboContUF_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    cboContUF.Clear
    sSQL = "SELECT * FROM TributacaoUF ORDER BY sigla"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboContUF.AddItem Rst.Fields("sigla")
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

Private Sub cboPais_DropDown()
    Dim Rst As Recordset
    cboPais.Clear
    Set Rst = RegistroBuscar("SELECT * FROM TributacaoPais ORDER BY Pais")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboPais.AddItem Rst.Fields("Pais")
                Rst.MoveNext
            Loop
    End If
End Sub





Private Sub cboRegimeTrib_DropDown()
    With cboRegimeTrib
        .Clear
        .AddItem "1 - " & PgDescrRegTrib(1)
        .AddItem "2 - " & PgDescrRegTrib(2)
        .AddItem "3 - " & PgDescrRegTrib(3)
    End With
End Sub

Private Sub cboTipoAtividade_DropDown()
    With cboTipoAtividade
        .AddItem "0 - Industria ou equiparado a industria"
        .AddItem "1 - Outros"
    End With
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
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    sstEmpresa.Tab = 0
    HDForm Me, False
    HDMenu Me, True
    
    txtCNPJ.Enabled = True
    
End Sub
Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    IdReg = 0
    HDMenu Me, False
    HDForm Me, True
    LimpaFormulario Me
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Selecione uma empresa"
        Exit Sub
    End If
    HDForm Me, True
    HDMenu Me, False

End Sub
Private Sub Excluir()
    If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
            MsgBox "Selecione uma Empresa"
            Exit Sub
        Else
            If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                        vbCrLf & _
                        "CNPJ: " & txtCNPJ.Text & vbCrLf & _
                        "Nome: " & txtNome.Text, vbYesNo + vbQuestion) = vbYes Then
                               
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
        Case "Pesquisar"
            PesquisarRegistro
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDForm Me, False
                'LimpaFormulario me
                'txtCNPJ.Enabled = True
            End If
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            txtCNPJ.Enabled = True
        
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select
End Sub
Public Sub MontarBaseDeDados()
    formManutencaoTabelas.IniciarManutencao Me
End Sub
Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim i           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer 'Contador de Registros
    cReg = 0
    For i = 0 To Me.Controls.Count - 1
        Set Controle = Me.Controls(i)
        
        If TypeOf Controle Is TextBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Text, "S")
            cReg = cReg + 1
        End If
        If TypeOf Controle Is ComboBox Then
            If Controle.Name = "cboRegimeTrib" Then
                    vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Left(Controle.Text, 1), "S")
                Else
                    vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Text, "S")
            End If
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



End Function
Private Sub txtCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
    
End Sub

Private Sub txtCNPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        BuscarDados (txtCNPJ.Text)
    End If
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub
Private Sub BuscarDados(strCNPJ As String)
    Dim Rst     As ADODB.Recordset
    Dim strSQL  As String
    
    sstEmpresa.Tab = 0
    
    strSQL = "SELECT * FROM " & strTabela & " WHERE CNPJ = '" & strCNPJ & "'"

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
Public Sub MostrarDados()
    On Error Resume Next
    Dim sSQL As String
    sSQL = "SELECT * FROM " & strTabela & " WHERE Id = " & IdReg

    ExibirDados Me, sSQL
    DoEvents
    pctLogo.Picture = LoadPicture(PgDadosEmpresa(IdReg).Logotipo)

End Sub

Private Sub txtIE_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub txtPISAliquota_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtPISAliquota.Text, KeyAscii, 2)
End Sub
Private Sub txtCOFINSAliquota_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtCOFINSAliquota.Text, KeyAscii, 2)
End Sub
