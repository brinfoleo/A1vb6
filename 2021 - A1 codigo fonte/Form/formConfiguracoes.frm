VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formConfiguracoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurações"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   11640
   Begin TabDlg.SSTab sstConfig 
      Height          =   6435
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11351
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "1 - Geral"
      TabPicture(0)   =   "formConfiguracoes.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "sstGeralConfig"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "2 - Emissao de NF-e"
      TabPicture(1)   =   "formConfiguracoes.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "sstNFe"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin TabDlg.SSTab sstNFe 
         Height          =   5895
         Left            =   180
         TabIndex        =   25
         Top             =   420
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   10398
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "2.1 - Geral"
         TabPicture(0)   =   "formConfiguracoes.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame12"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "2.2 - UniNFe"
         TabPicture(1)   =   "formConfiguracoes.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sstNFeUnimaker"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame12 
            Height          =   5175
            Left            =   180
            TabIndex        =   70
            Top             =   420
            Width           =   10635
            Begin VB.Frame Frame17 
               Caption         =   "Fuso Horário"
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
               Left            =   7680
               TabIndex        =   129
               Top             =   300
               Width           =   2715
               Begin VB.TextBox txtfusoHorario 
                  Height          =   315
                  Left            =   180
                  MaxLength       =   6
                  TabIndex        =   130
                  Text            =   "Text1"
                  Top             =   540
                  Width           =   1335
               End
               Begin VB.Label Label42 
                  Caption         =   "Fuso horário local:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   131
                  Top             =   300
                  Width           =   1455
               End
            End
            Begin VB.Frame Frame13 
               Caption         =   "Certificado Digital"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1875
               Left            =   240
               TabIndex        =   79
               Top             =   3060
               Width           =   5535
               Begin VB.CommandButton btoConsultarValCertDigital 
                  Caption         =   "Consultar Certificado Digital"
                  Height          =   615
                  Left            =   3420
                  TabIndex        =   84
                  Top             =   1140
                  Width           =   1875
               End
               Begin VB.TextBox txtFinalValidadeCertDigital 
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   83
                  Text            =   "Text1"
                  Top             =   660
                  Width           =   3555
               End
               Begin VB.TextBox txtInicioValidadeCertDigital 
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   82
                  Text            =   "Text1"
                  Top             =   300
                  Width           =   3555
               End
               Begin VB.Label Label34 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Final da Validade:"
                  Height          =   255
                  Left            =   180
                  TabIndex        =   81
                  Top             =   720
                  Width           =   1395
               End
               Begin VB.Label Label33 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Inicio da Validade:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   80
                  Top             =   360
                  Width           =   1335
               End
            End
            Begin VB.CheckBox chkBloqueionNFManual 
               Caption         =   "Bloquear inclusão de numero de Nota Fiscal Manual"
               Height          =   255
               Left            =   300
               TabIndex        =   78
               Top             =   2220
               Width           =   9375
            End
            Begin VB.TextBox txtNFePrazoCancelamento 
               Height          =   285
               Left            =   3360
               MaxLength       =   2
               TabIndex        =   77
               Text            =   "Text1"
               Top             =   1800
               Width           =   615
            End
            Begin VB.CheckBox chkEmissaoNFesPV 
               Caption         =   "Permitir emissão de varias NFe.s de uma unica Pré-Venda."
               Height          =   195
               Left            =   300
               TabIndex        =   74
               Top             =   480
               Width           =   9675
            End
            Begin VB.CheckBox chkInserirNomeVendXML 
               Caption         =   "Inserir o 1° nome do vendedor na Nota Fiscal"
               Height          =   255
               Left            =   300
               TabIndex        =   73
               Top             =   1140
               Width           =   3495
            End
            Begin VB.CheckBox chkTranspVolumes 
               Caption         =   "Exigir VOLUME, ESPECIE e PESO B/L caso o material seja entregue por transportadora."
               Height          =   195
               Left            =   300
               TabIndex        =   72
               Top             =   840
               Width           =   7815
            End
            Begin VB.ComboBox cboCodProdImpresso 
               Height          =   315
               Left            =   3660
               Style           =   2  'Dropdown List
               TabIndex        =   71
               Top             =   1440
               Width           =   3375
            End
            Begin VB.Label Label31 
               Caption         =   "Código do Produto a ser usado na Nota Fiscal:"
               Height          =   255
               Left            =   300
               TabIndex        =   76
               Top             =   1500
               Width           =   3675
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               Caption         =   "Prazo para cancelamento e NF-e (Horas):"
               Height          =   255
               Left            =   300
               TabIndex        =   75
               Top             =   1860
               Width           =   2955
            End
         End
         Begin TabDlg.SSTab sstNFeUnimaker 
            Height          =   5295
            Left            =   -74880
            TabIndex        =   26
            Top             =   420
            Width           =   10755
            _ExtentX        =   18971
            _ExtentY        =   9340
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "2.2.1 - Geral"
            TabPicture(0)   =   "formConfiguracoes.frx":0070
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame2"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "2.2.2 - Pastas"
            TabPicture(1)   =   "formConfiguracoes.frx":008C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame1"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "2.2.3 - DANFe"
            TabPicture(2)   =   "formConfiguracoes.frx":00A8
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame6"
            Tab(2).ControlCount=   1
            Begin VB.Frame Frame1 
               Height          =   4755
               Left            =   -74820
               TabIndex        =   48
               Top             =   360
               Width           =   10395
               Begin VB.CommandButton btoBusca 
                  Height          =   255
                  Index           =   6
                  Left            =   9960
                  Picture         =   "formConfiguracoes.frx":00C4
                  Style           =   1  'Graphical
                  TabIndex        =   62
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.CommandButton btoBusca 
                  Height          =   255
                  Index           =   5
                  Left            =   9960
                  Picture         =   "formConfiguracoes.frx":044E
                  Style           =   1  'Graphical
                  TabIndex        =   61
                  Top             =   3720
                  Width           =   315
               End
               Begin VB.CommandButton btoBusca 
                  Height          =   255
                  Index           =   4
                  Left            =   9960
                  Picture         =   "formConfiguracoes.frx":07D8
                  Style           =   1  'Graphical
                  TabIndex        =   60
                  Top             =   3060
                  Width           =   315
               End
               Begin VB.CommandButton btoBusca 
                  Height          =   255
                  Index           =   3
                  Left            =   9960
                  Picture         =   "formConfiguracoes.frx":0B62
                  Style           =   1  'Graphical
                  TabIndex        =   59
                  Top             =   2460
                  Width           =   315
               End
               Begin VB.CommandButton btoBusca 
                  Height          =   255
                  Index           =   2
                  Left            =   9960
                  Picture         =   "formConfiguracoes.frx":0EEC
                  Style           =   1  'Graphical
                  TabIndex        =   58
                  Top             =   1800
                  Width           =   315
               End
               Begin VB.CommandButton btoBusca 
                  Height          =   255
                  Index           =   1
                  Left            =   9960
                  Picture         =   "formConfiguracoes.frx":1276
                  Style           =   1  'Graphical
                  TabIndex        =   57
                  Top             =   1140
                  Width           =   315
               End
               Begin VB.CommandButton btoBusca 
                  Height          =   255
                  Index           =   0
                  Left            =   9960
                  Picture         =   "formConfiguracoes.frx":1600
                  Style           =   1  'Graphical
                  TabIndex        =   56
                  Top             =   540
                  Width           =   315
               End
               Begin VB.TextBox txtpValidar 
                  Height          =   285
                  Left            =   180
                  TabIndex        =   55
                  Text            =   "Text1"
                  Top             =   4320
                  Width           =   9735
               End
               Begin VB.TextBox txtpBackup 
                  Height          =   315
                  Left            =   180
                  TabIndex        =   54
                  Text            =   "Text1"
                  Top             =   3660
                  Width           =   9735
               End
               Begin VB.TextBox txtpErro 
                  Height          =   315
                  Left            =   180
                  TabIndex        =   53
                  Text            =   "Text1"
                  Top             =   3015
                  Width           =   9735
               End
               Begin VB.TextBox txtpEnviados 
                  Height          =   285
                  Left            =   180
                  TabIndex        =   52
                  Text            =   "Text1"
                  Top             =   2400
                  Width           =   9735
               End
               Begin VB.TextBox txtpRetorno 
                  Height          =   285
                  Left            =   180
                  TabIndex        =   51
                  Text            =   "Text1"
                  Top             =   1770
                  Width           =   9735
               End
               Begin VB.TextBox txtpEnviadosLote 
                  Height          =   315
                  Left            =   180
                  TabIndex        =   50
                  Text            =   "Text1"
                  Top             =   1125
                  Width           =   9735
               End
               Begin VB.TextBox txtpEnvio 
                  Height          =   285
                  Left            =   180
                  TabIndex        =   49
                  Text            =   "Text1"
                  Top             =   510
                  Width           =   9735
               End
               Begin MSComDlg.CommonDialog cmd 
                  Left            =   7980
                  Top             =   120
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin VB.Label Label7 
                  Caption         =   "Pasta onde será gravado os arquivos XML's a serem somente validados:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   69
                  Top             =   4050
                  Width           =   7875
               End
               Begin VB.Label Label6 
                  Caption         =   "Pasta para Backup dos XML's enviados:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   68
                  Top             =   3405
                  Width           =   7875
               End
               Begin VB.Label Label5 
                  Caption         =   "Pasta para arquivamento temporário dos XML's que apresentam erro na tentativa de envio:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   67
                  Top             =   2745
                  Width           =   7875
               End
               Begin VB.Label Label4 
                  Caption         =   "Pasta onde será gravado os arquivos XML's enviados:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   66
                  Top             =   2130
                  Width           =   7875
               End
               Begin VB.Label Label3 
                  Caption         =   "Pasta onde será gravado os arquivos XML's de retornodos WebServices:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   65
                  Top             =   1515
                  Width           =   7875
               End
               Begin VB.Label Label2 
                  Caption         =   "Pasta onde será gravado os arquivos XML's de NF-e a serem enviadas em lote para os WebServices:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   64
                  Top             =   855
                  Width           =   7875
               End
               Begin VB.Label Label1 
                  Caption         =   "Pasta onde será gravado os arquivos XML's a serem enviados individualmente para o WebServices:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   63
                  Top             =   240
                  Width           =   7875
               End
            End
            Begin VB.Frame Frame2 
               Height          =   4755
               Left            =   180
               TabIndex        =   36
               Top             =   360
               Width           =   10395
               Begin VB.Frame frmDtHrContigencia 
                  Caption         =   "Data/Hora entrada em contingência"
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
                  Left            =   6720
                  TabIndex        =   119
                  Top             =   180
                  Width           =   3555
                  Begin VB.TextBox txtMotivoContigencia 
                     Height          =   285
                     Left            =   120
                     MaxLength       =   500
                     TabIndex        =   125
                     Text            =   "Text1"
                     Top             =   1200
                     Width           =   3315
                  End
                  Begin MSComCtl2.DTPicker dtpDataContigencia 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   123
                     Top             =   540
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   105709569
                     CurrentDate     =   41057
                  End
                  Begin MSComCtl2.DTPicker dtpHoraContigencia 
                     Height          =   315
                     Left            =   1680
                     TabIndex        =   122
                     Top             =   540
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   105709570
                     CurrentDate     =   41057
                  End
                  Begin VB.Label Label41 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Motivo:"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   124
                     Top             =   960
                     Width           =   555
                  End
                  Begin VB.Label Label40 
                     Caption         =   "Hora:"
                     Height          =   255
                     Left            =   1680
                     TabIndex        =   121
                     Top             =   300
                     Width           =   495
                  End
                  Begin VB.Label Label39 
                     Caption         =   "Data:"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   120
                     Top             =   300
                     Width           =   435
                  End
               End
               Begin VB.CheckBox chkRetornoTXT 
                  Caption         =   "Gravar os retornos do webservices também no formato texto (TXT)"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   42
                  Top             =   4020
                  Width           =   5595
               End
               Begin VB.ComboBox cboEstadoUF 
                  Height          =   315
                  Left            =   360
                  Style           =   2  'Dropdown List
                  TabIndex        =   41
                  Top             =   646
                  Width           =   4575
               End
               Begin VB.ComboBox cboAmbiente 
                  Height          =   315
                  Left            =   360
                  Style           =   2  'Dropdown List
                  TabIndex        =   40
                  Top             =   1338
                  Width           =   4575
               End
               Begin VB.ComboBox cboTipoEmissao 
                  Height          =   315
                  Left            =   360
                  Style           =   2  'Dropdown List
                  TabIndex        =   39
                  Top             =   2030
                  Width           =   4575
               End
               Begin VB.ComboBox cboFormatoPasta 
                  Height          =   315
                  Left            =   360
                  Style           =   2  'Dropdown List
                  TabIndex        =   38
                  Top             =   2722
                  Width           =   4575
               End
               Begin VB.TextBox txtDiasXMLTemp 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   360
                  MaxLength       =   10
                  TabIndex        =   37
                  Text            =   "Text1"
                  Top             =   3420
                  Width           =   2175
               End
               Begin VB.Label Label8 
                  Caption         =   "Unidade Federativa (UF):"
                  Height          =   195
                  Left            =   360
                  TabIndex        =   47
                  Top             =   360
                  Width           =   1995
               End
               Begin VB.Label Label9 
                  Caption         =   "Ambiente:"
                  Height          =   195
                  Left            =   360
                  TabIndex        =   46
                  Top             =   1052
                  Width           =   1695
               End
               Begin VB.Label Label10 
                  Caption         =   "Tipo de Emissão:"
                  Height          =   195
                  Left            =   360
                  TabIndex        =   45
                  Top             =   1744
                  Width           =   1815
               End
               Begin VB.Label Label11 
                  Caption         =   "Como devem ser criados os diretórios baseados na data de emissão?"
                  Height          =   195
                  Left            =   360
                  TabIndex        =   44
                  Top             =   2436
                  Width           =   4995
               End
               Begin VB.Label Label12 
                  Caption         =   "Quantos dias devem ser mantidos os arquivos na pasta temporária e retorno? (0 para infinito): "
                  Height          =   195
                  Left            =   360
                  TabIndex        =   43
                  Top             =   3128
                  Width           =   6855
               End
            End
            Begin VB.Frame Frame6 
               Height          =   4635
               Left            =   -74880
               TabIndex        =   27
               Top             =   360
               Width           =   10395
               Begin VB.Frame Frame7 
                  Caption         =   "UniDANFe"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   915
                  Left            =   180
                  TabIndex        =   31
                  Top             =   180
                  Width           =   9795
                  Begin VB.CommandButton btoBusca 
                     Height          =   255
                     Index           =   8
                     Left            =   9360
                     Picture         =   "formConfiguracoes.frx":198A
                     Style           =   1  'Graphical
                     TabIndex        =   33
                     Top             =   540
                     Width           =   315
                  End
                  Begin VB.TextBox txtpUniDANFe 
                     Height          =   285
                     Left            =   120
                     TabIndex        =   32
                     Text            =   "Text1"
                     Top             =   510
                     Width           =   9135
                  End
                  Begin VB.Label Label16 
                     Caption         =   "Pasta onde esta o executavel do UniDANFe:"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   34
                     Top             =   240
                     Width           =   7875
                  End
               End
               Begin VB.TextBox txtDANFEnCopias 
                  Height          =   285
                  Left            =   1560
                  MaxLength       =   2
                  TabIndex        =   30
                  Text            =   "Text1"
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   1035
               End
               Begin VB.CheckBox chkDANFEVisualizar 
                  Caption         =   "Visualizar DANFe (validado pela SEFAZ) antes de imprimir."
                  Height          =   255
                  Left            =   120
                  TabIndex        =   29
                  Top             =   2220
                  Width           =   4575
               End
               Begin VB.CheckBox chkPreviewDanfe 
                  Caption         =   "Vizualizar DANFe ANTES do envio para Receita Ferderal"
                  Height          =   435
                  Left            =   120
                  TabIndex        =   28
                  Top             =   1740
                  Width           =   5055
               End
               Begin VB.Label Label17 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Numero de cópias:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   35
                  Top             =   1500
                  Visible         =   0   'False
                  Width           =   1395
               End
            End
         End
      End
      Begin TabDlg.SSTab sstGeralConfig 
         Height          =   5715
         Left            =   -74820
         TabIndex        =   2
         Top             =   480
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   10081
         _Version        =   393216
         Style           =   1
         Tabs            =   7
         Tab             =   2
         TabsPerRow      =   7
         TabHeight       =   520
         TabCaption(0)   =   "1.1 - Sistema"
         TabPicture(0)   =   "formConfiguracoes.frx":1D14
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "dtpDtUltMov"
         Tab(0).Control(1)=   "Frame11"
         Tab(0).Control(2)=   "chkGerClientesVisualizarOutrosFunc"
         Tab(0).Control(3)=   "chkMenuManutencaoTabelas"
         Tab(0).Control(4)=   "Frame4"
         Tab(0).Control(5)=   "Frame9"
         Tab(0).Control(6)=   "Label38"
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "1.2 - RH"
         TabPicture(1)   =   "formConfiguracoes.frx":1D30
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame8"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "1.3 - Cliente"
         TabPicture(2)   =   "formConfiguracoes.frx":1D4C
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame14"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "1.4 - Fornecedor"
         TabPicture(3)   =   "formConfiguracoes.frx":1D68
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame5"
         Tab(3).Control(1)=   "Frame3"
         Tab(3).Control(2)=   "chkNFDevolucaoCompra"
         Tab(3).Control(3)=   "chkAceitarEntradaNFSemAutorizacaoSEFAZ"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "1.5 - Financeiro"
         TabPicture(4)   =   "formConfiguracoes.frx":1D84
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame15"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "1.6 - Email"
         TabPicture(5)   =   "formConfiguracoes.frx":1DA0
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "btnTestarEmail"
         Tab(5).Control(1)=   "Frame10"
         Tab(5).ControlCount=   2
         TabCaption(6)   =   "1.7 - Estoque"
         TabPicture(6)   =   "formConfiguracoes.frx":1DBC
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Frame16"
         Tab(6).Control(1)=   "chkEstoqueAtualizarCusto"
         Tab(6).Control(2)=   "Frame18"
         Tab(6).ControlCount=   3
         Begin VB.Frame Frame18 
            Caption         =   "Estoque padrão:"
            Height          =   1095
            Left            =   -74880
            TabIndex        =   133
            Top             =   2400
            Width           =   4155
            Begin VB.ComboBox cboDeposito 
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   134
               Top             =   360
               Width           =   3675
            End
         End
         Begin VB.CommandButton btnTestarEmail 
            Caption         =   "Testar"
            Height          =   375
            Left            =   -69480
            TabIndex        =   132
            Top             =   840
            Width           =   1635
         End
         Begin VB.CheckBox chkEstoqueAtualizarCusto 
            Caption         =   "Atualizar Custo do Produto na entrada da Nota Fiscal"
            Height          =   255
            Left            =   -74700
            TabIndex        =   128
            Top             =   1680
            Width           =   4395
         End
         Begin VB.Frame Frame16 
            Caption         =   "Depósito"
            Height          =   1035
            Left            =   -74880
            TabIndex        =   126
            Top             =   480
            Width           =   4035
            Begin VB.CheckBox chkEstoqueSUverDepositos 
               Caption         =   "Super Usuario pode ver mais de um depósito"
               Height          =   615
               Left            =   180
               TabIndex        =   127
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.Frame Frame15 
            Height          =   3915
            Left            =   -74820
            TabIndex        =   118
            Top             =   600
            Width           =   9435
         End
         Begin VB.Frame Frame10 
            Caption         =   "Servidor de E-mail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2835
            Left            =   -74760
            TabIndex        =   105
            Top             =   540
            Width           =   4875
            Begin VB.TextBox txtMailSMTPPorta 
               Height          =   285
               Left            =   3960
               MaxLength       =   5
               TabIndex        =   112
               Text            =   "Text1"
               Top             =   360
               Width           =   555
            End
            Begin VB.CheckBox chkMailAutenticacao 
               Caption         =   "Meu servidor requer autenticação"
               Height          =   255
               Left            =   720
               TabIndex        =   111
               Top             =   1800
               Width           =   3255
            End
            Begin VB.TextBox txtMailSMTP 
               Height          =   285
               Left            =   720
               TabIndex        =   110
               Text            =   "Text1"
               Top             =   360
               Width           =   2655
            End
            Begin VB.TextBox txtMailEndereco 
               Height          =   285
               Left            =   720
               TabIndex        =   109
               Text            =   "Text1"
               Top             =   720
               Width           =   3795
            End
            Begin VB.TextBox txtMailLogin 
               Height          =   285
               Left            =   720
               TabIndex        =   108
               Text            =   "Text1"
               Top             =   1080
               Width           =   3795
            End
            Begin VB.TextBox txtMailSenha 
               Height          =   285
               Left            =   720
               TabIndex        =   107
               Text            =   "Text1"
               Top             =   1440
               Width           =   3795
            End
            Begin VB.CheckBox chkMailRecCopia 
               Caption         =   "Receber cópia do e-mail enviado"
               Height          =   195
               Left            =   720
               TabIndex        =   106
               Top             =   2100
               Width           =   3855
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "Porta:"
               Height          =   195
               Left            =   3480
               TabIndex        =   117
               Top             =   420
               Width           =   435
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               Caption         =   "SMTP:"
               Height          =   195
               Left            =   60
               TabIndex        =   116
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               Caption         =   "E-mail:"
               Height          =   195
               Left            =   180
               TabIndex        =   115
               Top             =   780
               Width           =   495
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               Caption         =   "Login:"
               Height          =   195
               Left            =   240
               TabIndex        =   114
               Top             =   1140
               Width           =   435
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               Caption         =   "Senha:"
               Height          =   195
               Left            =   180
               TabIndex        =   113
               Top             =   1500
               Width           =   495
            End
         End
         Begin MSComCtl2.DTPicker dtpDtUltMov 
            Height          =   315
            Left            =   -72180
            TabIndex        =   104
            Top             =   4260
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   165806081
            CurrentDate     =   41001
         End
         Begin VB.Frame Frame14 
            Height          =   4935
            Left            =   120
            TabIndex        =   101
            Top             =   480
            Width           =   10695
            Begin VB.CheckBox chkClienteLimiteCredito 
               Caption         =   "Aplicar o LIMITE DE CRÉDITO do cliente"
               Height          =   195
               Left            =   180
               TabIndex        =   102
               Top             =   300
               Width           =   3555
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Armazenamento de NF-e de Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   -74760
            TabIndex        =   97
            Top             =   480
            Width           =   9915
            Begin VB.CommandButton btoBusca 
               Height          =   255
               Index           =   7
               Left            =   9480
               Picture         =   "formConfiguracoes.frx":1DD8
               Style           =   1  'Graphical
               TabIndex        =   99
               Top             =   540
               Width           =   315
            End
            Begin VB.TextBox txtpXMLFornecedor 
               Height          =   285
               Left            =   120
               TabIndex        =   98
               Text            =   "Text1"
               Top             =   510
               Width           =   9195
            End
            Begin VB.Label Label15 
               Caption         =   "Pasta onde será gravado os arquivos XML's recebido dos fornecedores:"
               Height          =   195
               Left            =   120
               TabIndex        =   100
               Top             =   240
               Width           =   7875
            End
         End
         Begin VB.Frame Frame3 
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
            Height          =   2055
            Left            =   -74760
            TabIndex        =   89
            Top             =   1500
            Width           =   5775
            Begin VB.ComboBox cboFornecedorTpDoc 
               Height          =   315
               Left            =   1680
               Style           =   2  'Dropdown List
               TabIndex        =   92
               Top             =   1200
               Width           =   3555
            End
            Begin VB.ComboBox cboFornecedorCC 
               Height          =   315
               Left            =   1680
               Style           =   2  'Dropdown List
               TabIndex        =   91
               Top             =   780
               Width           =   3555
            End
            Begin VB.ComboBox cboFornecedorPlanoContas 
               Height          =   315
               Left            =   1680
               Style           =   2  'Dropdown List
               TabIndex        =   90
               Top             =   1560
               Width           =   3555
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               Caption         =   "Centro de Custos:"
               Height          =   195
               Left            =   300
               TabIndex        =   96
               Top             =   840
               Width           =   1275
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               Caption         =   "Tipo de Documento:"
               Height          =   195
               Left            =   120
               TabIndex        =   95
               Top             =   1260
               Width           =   1455
            End
            Begin VB.Label Label30 
               Caption         =   "Os dados abaixo serão usados no preenchimento da(s) fatura(s):"
               Height          =   315
               Left            =   120
               TabIndex        =   94
               Top             =   360
               Width           =   4995
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               Caption         =   "Plano de Contas:"
               Height          =   195
               Left            =   180
               TabIndex        =   93
               Top             =   1620
               Width           =   1335
            End
         End
         Begin VB.CheckBox chkNFDevolucaoCompra 
            Caption         =   "Considerar NF de devolução como compra."
            Height          =   195
            Left            =   -74820
            TabIndex        =   88
            Top             =   4200
            Width           =   3915
         End
         Begin VB.CheckBox chkAceitarEntradaNFSemAutorizacaoSEFAZ 
            Caption         =   "Aceitar NF-e de Entrada sem autorização da SEFAZ"
            Height          =   195
            Left            =   -74820
            TabIndex        =   87
            Top             =   4500
            Width           =   4215
         End
         Begin VB.Frame Frame11 
            Caption         =   "Boleto Bancario"
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
            Left            =   -72420
            TabIndex        =   22
            Top             =   2880
            Width           =   5175
            Begin VB.ComboBox cboBoleto 
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   300
               Width           =   3855
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               Caption         =   "Impresso em:"
               Height          =   195
               Left            =   120
               TabIndex        =   24
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.CheckBox chkGerClientesVisualizarOutrosFunc 
            Caption         =   "Gerenciador de Clientes: Permitir que veja o movimento de outros Funcionarios"
            Height          =   195
            Left            =   -74760
            TabIndex        =   21
            Top             =   2100
            Width           =   6135
         End
         Begin VB.CheckBox chkMenuManutencaoTabelas 
            Caption         =   "Mostrar icone para Manutenção de Tabela"
            Height          =   315
            Left            =   -74760
            TabIndex        =   20
            Top             =   1680
            Width           =   4995
         End
         Begin VB.Frame Frame4 
            Caption         =   "Casas decimais"
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
            Left            =   -74820
            TabIndex        =   15
            Top             =   2880
            Width           =   2175
            Begin VB.TextBox txtcDecMoeda 
               Height          =   285
               Left            =   1020
               MaxLength       =   1
               TabIndex        =   17
               Text            =   "Text1"
               Top             =   300
               Width           =   975
            End
            Begin VB.TextBox txtcDecQtd 
               Height          =   285
               Left            =   1020
               MaxLength       =   5
               TabIndex        =   16
               Text            =   "Text1"
               Top             =   780
               Width           =   975
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "Moeda:"
               Height          =   195
               Left            =   420
               TabIndex        =   19
               Top             =   360
               Width           =   555
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Quantidade:"
               Height          =   195
               Left            =   120
               TabIndex        =   18
               Top             =   840
               Width           =   855
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Armazenamento de Arquivos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   -74760
            TabIndex        =   11
            Top             =   660
            Width           =   10035
            Begin VB.CommandButton btoBusca 
               Height          =   255
               Index           =   9
               Left            =   9540
               Picture         =   "formConfiguracoes.frx":2162
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   540
               Width           =   315
            End
            Begin VB.TextBox txtpFileArmazenamento 
               Height          =   285
               Left            =   120
               TabIndex        =   12
               Text            =   "Text1"
               Top             =   510
               Width           =   9255
            End
            Begin VB.Label Label22 
               Caption         =   "Pasta onde será gravado todos os arquivos de uso do Sistema:"
               Height          =   195
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   7875
            End
         End
         Begin VB.Frame Frame8 
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
            Height          =   2775
            Left            =   -74760
            TabIndex        =   3
            Top             =   660
            Width           =   5775
            Begin VB.ComboBox cboRHPlanoContas 
               Height          =   315
               Left            =   1620
               Style           =   2  'Dropdown List
               TabIndex        =   86
               Top             =   1200
               Width           =   3555
            End
            Begin VB.ComboBox cboRHDocumento 
               Height          =   315
               Left            =   1620
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   1620
               Width           =   3555
            End
            Begin VB.ComboBox cboRHCentroCustos 
               Height          =   315
               Left            =   1620
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   780
               Width           =   3555
            End
            Begin VB.ComboBox cboRHConta 
               Height          =   315
               Left            =   1620
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   2040
               Width           =   3555
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               Caption         =   "Plano de Contas:"
               Height          =   195
               Left            =   180
               TabIndex        =   85
               Top             =   1260
               Width           =   1335
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Caption         =   "Centro de Custos:"
               Height          =   195
               Left            =   240
               TabIndex        =   10
               Top             =   840
               Width           =   1275
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Tipo de Documento:"
               Height          =   195
               Left            =   60
               TabIndex        =   9
               Top             =   1680
               Width           =   1455
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               Caption         =   "Conta:"
               Height          =   195
               Left            =   1020
               TabIndex        =   8
               Top             =   2100
               Width           =   495
            End
            Begin VB.Label Label21 
               Caption         =   "Os dados abaixo serão usados no preenchimento da comissão:"
               Height          =   315
               Left            =   120
               TabIndex        =   7
               Top             =   360
               Width           =   4995
            End
         End
         Begin VB.Label Label38 
            Caption         =   "Data do ultimo acesso ao sistema:"
            Height          =   255
            Left            =   -74640
            TabIndex        =   103
            Top             =   4320
            Width           =   2475
         End
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11640
      _ExtentX        =   20532
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Incluir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alterar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
               Picture         =   "formConfiguracoes.frx":24EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConfiguracoes.frx":293E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConfiguracoes.frx":2C58
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConfiguracoes.frx":34EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConfiguracoes.frx":473C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConfiguracoes.frx":5016
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConfiguracoes.frx":58A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConfiguracoes.frx":613A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConfiguracoes.frx":738C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConfiguracoes.frx":76A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConfiguracoes.frx":79C0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formConfiguracoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IdReg       As Integer
Dim strTabela   As String
Dim caminho     As String

'Coloque estas declarações em um módulo
'Se colocar em um formulário lembre-se de não usar como 'private'

'Existem outras flags para parametrizar a pesquisa
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                  ByVal lpString2 As String) As Long
'Tipo para def
Private Type BrowseInfo
    hWndOwner       As Long
    pIDLRoot        As Long
    pszDisplayName  As Long
    lpszTitle       As Long
    ulFlags         As Long
    lpfnCallback    As Long
    lParam          As Long
    iImage          As Long
End Type


Private Sub grvReg(nmArquivo As String, Dados As String)
    'Grava em txt para atualizacao do uninfe
    On Error GoTo TrtErro
    
    If Trim(caminho) = "" Then
        MsgBox "Erro ao definir o caminho."
        Exit Sub
    End If
    'define o ObjPreview filesystem e demais variaveis
    Dim fso As New FileSystemObject
    Dim Arquivo As File
    Dim arquivoLog As TextStream
    Dim Msg As String
    'Dim caminho As String


    'se o arquivo não existir então cria
    If fso.FileExists(caminho) Then
            Set Arquivo = fso.GetFile(caminho)
        Else
            Set arquivoLog = fso.CreateTextFile(caminho)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(caminho)
    End If
    'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
    
    'monta informações para gerar a linha da mensagem
    Msg = Dados

    'inclui linhas no arquivo texto
    arquivoLog.WriteLine Msg
    
    'escreve uma linha em branco no arquivo - se voce quiser
    'arquivoLog.WriteBlankLines (1)
    'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing
    Exit Sub
TrtErro:
        MsgBox "Erro ao gerar registro da NFe em Texto .                                   " & _
           vbCrLf & vbCrLf & _
           "Erro n.: " & Err.Number & _
           vbCrLf & vbCrLf & _
           "Descrição: " & Err.Description & _
           vbCrLf

End Sub


Private Sub MontarArquivoConfiguracao_UniNFe()
    
    If MsgBox("Reconfigurar o sistema de envio de NFe (UniNFe)?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    
    

    If Dir(SistemPath & "\cfg", vbDirectory) = "" Then
        MkDir SistemPath & "\cfg"
    End If

    caminho = SistemPath & "\cfg\uninfe-alt-con.txt" ' & LocalArq
    If Dir(Trim(caminho)) <> "" Then
            Kill caminho
    End If
    '#################################################################
    '### UniNFe v.4.0.4352.20650                                   ###
    '#################################################################
    grvReg caminho, "PastaXmlEnvio|" & PgDadosConfig.pEnvio
    grvReg caminho, "PastaXmlRetorno|" & PgDadosConfig.pRetorno
    grvReg caminho, "PastaXmlEnviado|" & PgDadosConfig.pEnviados
    grvReg caminho, "PastaXmlErro|" & PgDadosConfig.pErro
    grvReg caminho, "PastaBackup|" & PgDadosConfig.pBackup
    grvReg caminho, "PastaXmlEmLote|" & PgDadosConfig.pEnviadosLote
    grvReg caminho, "PastaValidar|" & PgDadosConfig.pValidar
    grvReg caminho, "UnidadeFederativaCodigo|" & PgDadosConfig.uf
    grvReg caminho, "AmbienteCodigo|" & PgDadosConfig.Ambiente
    grvReg caminho, "tpEmis|" & PgDadosConfig.TpEmissao
    grvReg caminho, "GravarRetornoTXTNFe|" & IIf(PgDadosConfig.GravarRetornoTXT = 0, "False", "True")
    grvReg caminho, "DiretorioSalvarComo|" & PgDadosConfig.FormatoPasta
    grvReg caminho, "DiasLimpeza|" & PgDadosConfig.DiasXMLTemp
    'grvReg caminho, "PastaExeUniDanfe|" & ""
    'grvReg caminho, "PastaConfigUniDanfe|" & ""
    'grvReg caminho, "XMLDanfeMonNFe|" & ""
    'grvReg caminho, "XMLDanfeMonProcNFe|" & ""
    grvReg caminho, "TempoConsulta|" & "3"
    'grvReg caminho, "Proxy|" & ""
    'grvReg caminho, "ProxyServidor|" & ""
    'grvReg caminho, "ProxyUsuario|" & ""
    'grvReg caminho, "ProxySenha|" & ""
    'grvReg caminho, "ProxyPorta|" & ""
    'grvReg caminho, "SenhaConfig|" & ""
    'grvReg caminho, "FTPAtivo|" & ""
    'grvReg caminho, "FTPGravaXMLPAstaUnica|" & ""
    'grvReg caminho, "FTPNomeDoUsuario|" & ""
    'grvReg caminho, "FTPNomeDoServidor|" & ""
    'grvReg caminho, "FTPPastaAutorizados|" & ""
    'grvReg caminho, "FTPPastaRetornos|" & ""
    'grvReg caminho, "FTPPorta|" & ""
    'grvReg caminho, "FTPSenha|" & ""
    
    
    MoverPastaConfig_UniNFe
    
End Sub
Private Sub MoverPastaConfig_UniNFe()
    On Error Resume Next
    FileCopy caminho, PgDadosConfig.pEnvio & "\uninfe-alt-con.txt"


End Sub

Private Sub btnTestarEmail_Click()
    On Error GoTo trtErroSendMail
    Dim poSendMail As New vbSendMail.clsSendMail
    Dim StatusEnvio As Integer '0 - Cancelado,1 - enviado , -1 - erro no envio
    Dim bHtml       As Boolean 'Informa se o email sera enviado em HTML ou TXT

    'HDForm Me, False
    'stbConexao.Panels(1).Text = "Enviando email para " & LCase(Trim(txtTO.Text))
    With poSendMail
        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        '.SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        '.EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = txtMailSMTP.Text   '"smtp.metalcentermetais.com.br" 'txtServer.Text                  ' Required the fist time, optional thereafter
        .from = txtMailEndereco.Text '"nfe@metalcentermetais.com.br"      'txtFrom.Text                        ' Required the fist time, optional thereafter
        .FromDisplayName = txtMailEndereco.Text          'txtFromName.Text         ' Optional, saved after first use
        
        .Recipient = Trim(txtMailEndereco.Text)                 'txtTO.Text                     ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = "Destinatario"       'txtToName.Text      ' Optional, separate multiple entries with delimiter character
        '.CcRecipient = txtCc                        ' Optional, separate multiple entries with delimiter character
        '.CcDisplayName = txtCcName                  ' Optional, separate multiple entries with delimiter character
        '.BccRecipient = txtBcc                      ' Optional, separate multiple entries with delimiter character
        '.ReplyToAddress = txtFrom.Text              ' Optional, used when different than 'From' address
        .Subject = "Teste de email"                ' Optional
        .Message = "Email enviado para teste"                      ' Optional
        '.Attachment = Trim(App.Path & "\a1.ini")           'Trim(txtAttach.Text)          ' Optional, separate multiple entries with delimiter character

        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MIME_ENCODE                   'MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = NORMAL_PRIORITY                 ' etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = False                            ' bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = True                  ' bAuthLogin             ' Optional, default = FALSE
        '.UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .Username = txtMailLogin.Text   '"nfe@metalcentermetais.com.br"  'txtUserName                     ' Optional, default = Null String
        .Password = txtMailSenha.Text '"qwe123"                        'txtPassword                     ' Optional, default = Null String, value is NOT saved
        '.POP3Host = txtPopServer
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
        .SMTPPort = txtMailSMTPPorta.Text       ' Optional, default = 25

        ' **************************************************************************
        ' OK, all of the properties are set, send the email...
        ' **************************************************************************
        ' .Connect                                  ' Optional, use when sending bulk mail
        .Send                                       ' Required
        ' .Disconnect                               ' Optional, use when sending bulk mail
        'txtServer.Text = .SMTPHost                  ' Optional, re-populate the Host in case
                                                    ' MX look up was used to find a host    End With
    End With
    'Screen.MousePointer = vbDefault
    'cmdSend.Enabled = True
    MsgBox "Email enviado com sucesso...", vbInformation, App.EXEName
    'HDForm Me, True
    'stbConexao.Panels(1).Text = ""
    Exit Sub
trtErroSendMail:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub btoBusca_Click(Index As Integer)
   Dim lpIDList        As Long
    Dim sBuffer         As String
    Dim szTitle         As String
    Dim tBrowseInfo     As BrowseInfo

    'Personaliza a procura
    szTitle = "Selecione a Pasta"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN  ' + BIF_EDITBOX
    End With

    'Abre a janela de procura
    'E retorna o caminho da pasta selecionada
    lpIDList = SHBrowseForFolder(tBrowseInfo)

    'Se existir alguma pasta selecionada extrair
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        Select Case Index
            Case 0
                txtpEnvio.Text = sBuffer
            Case 1
                txtpEnviadosLote.Text = sBuffer
            Case 2
                txtpRetorno.Text = sBuffer
            Case 3
                txtpEnviados.Text = sBuffer
            Case 4
                txtpErro.Text = sBuffer
            Case 5
                txtpBackup.Text = sBuffer
            Case 6
                txtpValidar.Text = sBuffer
            Case 7
                txtpXMLFornecedor.Text = sBuffer
            Case 8
                txtpUniDANFe.Text = sBuffer
            Case 9
                txtpFileArmazenamento.Text = sBuffer
        End Select
    End If
End Sub


Private Sub btoConsultarValCertDigital_Click()
    
    On Error GoTo TrtErroCD
    
    If MsgBox("Esse Processo levara até 5 segundos. Deseja continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    txtInicioValidadeCertDigital.Text = ""
    txtFinalValidadeCertDigital.Text = ""
    
    If Dir(SistemPath & "\cfg", vbDirectory) = "" Then
        MkDir SistemPath & "\cfg"
    End If

    caminho = SistemPath & "\cfg\uninfe-cons-inf.txt" ' & LocalArq
    If Dir(Trim(caminho)) <> "" Then
            Kill caminho
    End If
    grvReg caminho, "xServ|CONS-INF"
    FileCopy caminho, PgDadosConfig.pEnvio & "\uninfe-cons-inf.txt"
    
    
    '########################################################################################
    '### Gerar loop de 10 segundos para saber se ja houve retorno dos dados
    '########################################################################################
    Dim lFile       As Long
    Dim linha       As String
    Dim DtIni As String
    Dim DtFin As String
    Dim pRetorno As String
    
    pRetorno = PgDadosConfig.pRetorno & "\uninfe-ret-cons-inf.txt"
    
    Pausa (5)
    If Dir(pRetorno) = "" Then
        MsgBox "Não foi encontrado o arquivo de retorno! Solicite novamente mais tarde!", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lFile = FreeFile
    Open pRetorno For Input As lFile   'abre o arquivo texto
    Do While Not EOF(lFile)
        Line Input #lFile, linha 'lê uma linha do arquivo texto
        If InStr(LCase(linha), "dvalini") <> 0 Then
            DtIni = Trim(Mid(linha, InStr(linha, "|") + 1, Len(linha)))
        End If
        If InStr(LCase(linha), "dvalfin") <> 0 Then
            DtFin = Trim(Mid(linha, InStr(linha, "|") + 1, Len(linha)))
        End If
    Loop
    Close #lFile
   
    'RegistroAlterar "Configuracoes", vReg, 1 ', "ID=1"
    
    txtInicioValidadeCertDigital.Text = DtIni 'PgDadosConfig.IniValCertDigital
    txtFinalValidadeCertDigital.Text = DtFin 'PgDadosConfig.FinValCertDigital
    Exit Sub
TrtErroCD:
    txtInicioValidadeCertDigital.Text = ""
    txtFinalValidadeCertDigital.Text = ""
End Sub

Private Sub Pausa(Seconds As Single)
'Pausa o sistema em X segundos
     Dim EndTime As Date
     EndTime = DateAdd("s", Seconds, Now)

     Do
       DoEvents
     Loop Until Now >= EndTime

 End Sub
Private Sub cboAmbiente_DropDown()
    With cboAmbiente
        .Clear
        .AddItem "1 - Produção"
        .AddItem "2 - Homologação"
    End With
    
End Sub


Private Sub cboBoleto_DropDown()
    With cboBoleto
        .Clear
        .AddItem "1 - Formulario"
        .AddItem "2 - Papel A4"
    End With
    
End Sub



Private Sub cboCodProdImpresso_DropDown()
    With cboCodProdImpresso
        .Clear
        .AddItem "01 - Código interno"
        .AddItem "02 - Referencia"
    End With
End Sub

Private Sub cboDeposito_DropDown()
    Dim Rst As Recordset
    cboDeposito.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueDeposito WHERE ID_Empresa = " & ID_Empresa & " ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
            
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboDeposito.AddItem Left(String(3, "0"), 3 - Len(Trim(Rst.Fields("id")))) & Trim(Rst.Fields("id")) & " - " & Rst.Fields("descricao")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub cboEstadoUF_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    sSQL = "SELECT * FROM TributacaoUF ORDER BY Sigla"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum estado Cadastrado"
        Else
            Rst.MoveFirst
            cboEstadoUF.Clear
            Do Until Rst.EOF
                cboEstadoUF.AddItem Rst.Fields("CodUF") & " - " & Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub


Private Sub cboFormatoPasta_DropDown()
    Dim sDados As String
    sDados = "AMD|AM|AD|MDA|MD|MA|DMA|DM|DA|A\M\D|A\M|A\D|M\D\A|M\D|M\A|D\M\A|D\M|D\A|"
    With cboFormatoPasta
        Do Until InStr(sDados, "|") = 0
            .AddItem Mid(sDados, 1, InStr(sDados, "|") - 1)
            sDados = Mid(sDados, InStr(sDados, "|") + 1, Len(sDados))
        Loop

    End With
End Sub


Private Sub cboFornecedorCC_DropDown()
    Dim Rst As Recordset
    cboFornecedorCC.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCentroCustos WHERE id_empresa=" & ID_Empresa)

    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboFornecedorCC.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
End Sub


Private Sub cboFornecedorPlanoContas_DropDown()
    Dim Rst As Recordset
    cboFornecedorPlanoContas.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroPlanoContas WHERE id_empresa=" & ID_Empresa & " ORDER BY Codigo")
    If Rst.BOF And Rst.EOF Then
            'Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboFornecedorPlanoContas.AddItem ZE(Rst.Fields("id"), 3) & " - (" & Rst.Fields("Codigo") & ") " & Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub cboFornecedorTpDoc_DropDown()
    Dim Rst As Recordset
    cboFornecedorTpDoc.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento WHERE id_empresa=" & ID_Empresa)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboFornecedorTpDoc.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboRHPlanoContas_DropDown()
    Dim Rst As Recordset
    cboRHPlanoContas.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroPlanoContas  WHERE id_empresa=" & ID_Empresa & " ORDER BY Codigo")
    If Rst.BOF And Rst.EOF Then
            'Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboRHPlanoContas.AddItem ZE(Rst.Fields("id"), 3) & " - (" & Rst.Fields("Codigo") & ") " & Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close

End Sub



Private Sub cboTipoEmissao_DropDown()
    Dim i As Integer
    With cboTipoEmissao
        .Clear
        For i = 1 To 5
            .AddItem i & " - " & PgDescrTipoEmissao(i)
        
        Next
    End With
End Sub



Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    
    sstConfig.Tab = 0
    sstNFeUnimaker.Tab = 0
    sstGeralConfig.Tab = 0
    sstNFe.Tab = 0
    
    LimpaFormulario Me
    HDForm Me, False
    HDMenu Me, True
    'IdReg = 0
    MostrarDados
    'txteMailCC.Enabled = False
End Sub


Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Alterar"
            Alterar
        Case "Salvar"
            If ValidarDados = False Then Exit Sub
            
            If grvRegistro = True Then
                HDMenu Me, True
                HDForm Me, False
            End If
            
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
        
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select
End Sub
Public Sub MontarBaseDeDados()
    formManutencaoTabelas.IniciarManutencao Me ', "SELECT * FROM Clientes"
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    HDForm Me, True
    HDMenu Me, False
End Sub


Private Function ValidarDados() As Boolean
    ValidarDados = False
    If Trim(cboBoleto.Text) = "" Then
        MsgBox "Favor selecionar o tipo de impressão do boleto bancario.", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    If Trim(cboCodProdImpresso.Text) = "" Then
        MsgBox "Favor selecionar o tipo de Código do Produto a ser Impresso na NF-e.", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
      If Len(Trim(txtfusoHorario.Text)) < 6 Then
        MsgBox "Favor informar fuso horário local válido.", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    ValidarDados = True
End Function
Private Function grvRegistro() As Boolean
    'Grava na base de dados
    Dim vReg(1000)      As Variant
    Dim cReg            As Integer
    
        
    cReg = 0
    vReg(cReg) = Array("DtUltMov", dtpDtUltMov.Value, "D"): cReg = cReg + 1
    
    vReg(cReg) = Array("cDecQtd", txtcDecQtd.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("cDecMoeda", txtcDecMoeda.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("EmissaoNFesPV", chkEmissaoNFesPV.Value, "S"): cReg = cReg + 1
    vReg(cReg) = Array("GerClientesVisualizarOutrosFunc", chkGerClientesVisualizarOutrosFunc.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("Boleto", Left(cboBoleto.Text, 1), "S"): cReg = cReg + 1
    vReg(cReg) = Array("TranspVolumes", chkTranspVolumes.Value, "N"): cReg = cReg + 1
    
    
    'E-MAIL
    vReg(cReg) = Array("MailSMTPPorta", Trim(txtMailSMTPPorta.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("MailSMTP", Trim(txtMailSMTP.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("MailEndereco", Trim(txtMailEndereco.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("MailLogin", Trim(txtMailLogin.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("MailSenha", Trim(txtMailSenha.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("MailAutenticacao", chkMailAutenticacao.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("MailRecCopia", chkMailRecCopia.Value, "N"): cReg = cReg + 1
    
    
    'RH
    vReg(cReg) = Array("RHCentroCustos", Left(cboRHCentroCustos.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("RHDocumento", Left(cboRHDocumento.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("RHConta", Left(cboRHConta.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("RHPlanoContas", Left(cboRHPlanoContas.Text, 3), "S"): cReg = cReg + 1
    
    'Fornecedor
    vReg(cReg) = Array("NFDevolucaoCompra", chkNFDevolucaoCompra.Value, "S"): cReg = cReg + 1
    vReg(cReg) = Array("AceitarEntradaNFSemAutorizacaoSEFAZ", chkAceitarEntradaNFSemAutorizacaoSEFAZ.Value, "S"): cReg = cReg + 1
    
    'Estoque
    vReg(cReg) = Array("EstoqueAtualizarCusto", chkEstoqueAtualizarCusto.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("EstoqueSUverDepositos", chkEstoqueSUverDepositos.Value, "N"): cReg = cReg + 1
    If Len(Trim(cboDeposito.Text)) <> 0 Then
        vReg(cReg) = Array("EstoqueDepositoPadrao", Mid(cboDeposito.Text, 1, 3), "N"): cReg = cReg + 1
    End If
    
    'Cliente
    vReg(cReg) = Array("ClienteLimiteCredito", chkClienteLimiteCredito.Value, "S"): cReg = cReg + 1
    
    'NFe
    vReg(cReg) = Array("EstadoUF", Left(cboEstadoUF.Text, 2), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Ambiente", Left(cboAmbiente.Text, 1), "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("TipoEmissao", Left(cboTipoEmissao.Text, 1), "S"): cReg = cReg + 1
    If Left(cboTipoEmissao.Text, 1) <> "1" Then
    
            vReg(cReg) = Array("DataContigencia", dtpDataContigencia.Value, "S"): cReg = cReg + 1
            vReg(cReg) = Array("HoraContigencia", Format(dtpHoraContigencia.Value, "HH:MM:SS"), "S"): cReg = cReg + 1
            vReg(cReg) = Array("MotivoContigencia", Trim(txtMotivoContigencia.Text), "S"): cReg = cReg + 1
        Else
            vReg(cReg) = Array("DataContigencia", "", "S"): cReg = cReg + 1
            vReg(cReg) = Array("HoraContigencia", "", "S"): cReg = cReg + 1
            vReg(cReg) = Array("MotivoContigencia", "", "S"): cReg = cReg + 1
    End If
    vReg(cReg) = Array("FormatoPasta", cboFormatoPasta.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("DiasXMLTemp", txtDiasXMLTemp.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("MenuManutencaoTabelas", chkMenuManutencaoTabelas.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("RetornoTXT", chkRetornoTXT.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("InserirNomeVendXML", chkInserirNomeVendXML.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("CodProdImpresso", Left(cboCodProdImpresso.Text, 2), "S"): cReg = cReg + 1
    vReg(cReg) = Array("NFePrazoCancelamento", Left(txtNFePrazoCancelamento.Text, 2), "N"): cReg = cReg + 1
    vReg(cReg) = Array("BloqueionNFManual", chkBloqueionNFManual.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("fusoHorario", Trim(txtfusoHorario.Text), "S"): cReg = cReg + 1
    '
    
    vReg(cReg) = Array("pEnviados", txtpEnviados.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("pEnviadosLote", txtpEnviadosLote.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("pRetorno", txtpRetorno.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("pEnvio", txtpEnvio.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("pErro", txtpErro.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("pBackup", txtpBackup.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("pValidar", txtpValidar.Text, "S"): cReg = cReg + 1
    
    'Validade do Certificado Digital
    vReg(cReg) = Array("InicioValidadeCertDigital", txtInicioValidadeCertDigital.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("FinalValidadeCertDigital", txtFinalValidadeCertDigital.Text, "S"): cReg = cReg + 1
    
    
    'Geral
    'Fornecedor
    vReg(cReg) = Array("pXMLFornecedor", txtpXMLFornecedor.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("FornecedorCC", Left(cboFornecedorCC.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("FornecedorTpDoc", Left(cboFornecedorTpDoc.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("FornecedorPlanoContas", Left(cboFornecedorPlanoContas.Text, 3), "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("pFileArmazenamento", txtpFileArmazenamento.Text, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("pUNIDANFe", txtpUniDANFe.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("DANFenCopias", txtDANFEnCopias.Text, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("DANFEEnviarMail", chkDANFEEnviarMail.Value, "S"): cReg = cReg + 1
    
    'vReg(cReg) = Array("DANFEEnviarMailCC", chkDANFEEnviarMailCC.Value, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("eMailCC", txteMailCC.Text, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("DANFEVisualizar", chkDANFEVisualizar.Value, "S"): cReg = cReg + 1
    vReg(cReg) = Array("PreviewDanfe", chkPreviewDanfe.Value, "S") ': cReg = cReg + 1
    
    If IdReg = 0 Then
            If RegistroIncluir(strTabela, vReg, cReg) = 0 Then
                    MsgBox "Erro ao Incluir"
                    grvRegistro = False
                Else
                    grvRegistro = True
                    IdReg = 1
            End If
        Else
            If RegistroAlterar(strTabela, vReg, cReg, "Id = " & IdReg) = False Then
                    MsgBox "Erro ao Alterar."
                    grvRegistro = False
                Else
                    grvRegistro = True
                    IdReg = 1
                
            End If
    End If
    
    MontarArquivoConfiguracao_UniNFe
End Function
Private Sub MostrarDados()
    On Error Resume Next
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim sTMP    As String
    sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            IdReg = 0
            Exit Sub
        Else
            Rst.MoveFirst
            IdReg = Rst.Fields("Id")
    End If
    sSQL = "SELECT * FROM " & strTabela & " WHERE  ID_Empresa = " & ID_Empresa & " AND ID = " & IdReg
    ExibirDados Me, sSQL
    
    With cboEstadoUF
        sTMP = .Text
        .Clear
        If Trim(sTMP) <> "" Then
            .AddItem IIf(pgDadosICMS(sTMP, 1).Descricao = "", " ", sTMP & " - " & pgDadosICMS(sTMP, 1).Descricao)
            .Text = .List(0)
        End If
    End With
    
     
    With cboAmbiente
        sTMP = .Text
        .Clear
        .AddItem IIf(sTMP = "1", "1 - Produção", "2 - Homologação")
        .Text = .List(0)
    End With
    
    With cboTipoEmissao
        sTMP = IIf(Trim(.Text) = "", 0, .Text)
        .Clear
        .AddItem IIf(Trim(PgDescrTipoEmissao(CInt(sTMP))) = "", " ", sTMP & " - " & PgDescrTipoEmissao(CInt(sTMP)))
        .Text = .List(0)
        'If Left(.Text, 1) = "1" Then
        '    dtpDataContigencia.Value = "Date"
        '    dtpHoraContigencia.Value = "00:00:00"
        '    dtpDataContigencia.Value = ""
        'End If
            
    End With
    
    
    With cboRHConta
        sTMP = IIf(Trim(.Text) = "", 0, .Text)
        .Clear
        .AddItem IIf(Trim(pgDadosConta(CInt(sTMP)).agencia) = "", " ", sTMP & " - " & pgDadosConta(CInt(sTMP)).agencia & "/" & pgDadosConta(CInt(sTMP)).conta)
        .Text = .List(0)
    End With
     With cboRHDocumento
        sTMP = IIf(Trim(.Text) = "", 0, .Text)
        .Clear
        .AddItem IIf(Trim(pgDadosTipoDocumento(CInt(sTMP)).Descricao) = "", " ", sTMP & " - " & pgDadosTipoDocumento(CInt(sTMP)).Descricao)
        .Text = .List(0)
    End With
    With cboRHCentroCustos
        sTMP = IIf(Trim(.Text) = "", 0, .Text)
        .Clear
        .AddItem IIf(Trim(pgDadosCentroCustos(CInt(sTMP)).Descricao) = "", " ", sTMP & " - " & pgDadosCentroCustos(CInt(sTMP)).Descricao)
        .Text = .List(0)
    End With
    
    With cboRHPlanoContas
        sTMP = IIf(Trim(.Text) = "", 0, .Text)
        .Clear
        .AddItem IIf(Trim(PgDadosPlanoContas("ID", CInt(sTMP)).Descricao) = "", " ", sTMP & " - (" & PgDadosPlanoContas("ID", CInt(sTMP)).Codigo & ") " & PgDadosPlanoContas("ID", CInt(sTMP)).Descricao)
        .Text = .List(0)
    End With
    
    With cboDeposito
        sTMP = IIf(Trim(.Text) = "", 0, .Text)
        .Clear
        .AddItem IIf(cNull(Rst.Fields("EstoqueDepositoPadrao")) = "", " ", Mid(String(3, "0"), 1, 3 - Len(Trim(Rst.Fields("EstoqueDepositoPadrao")))) & Rst.Fields("EstoqueDepositoPadrao") & " - " & pgDescrDeposito(Rst.Fields("EstoqueDepositoPadrao")))
        .Text = .List(0)
    End With
    
    
    With cboFornecedorPlanoContas
        sTMP = IIf(Trim(.Text) = "", 0, .Text)
        .Clear
        .AddItem IIf(Trim(PgDadosPlanoContas("ID", CInt(sTMP)).Descricao) = "", " ", sTMP & " - (" & PgDadosPlanoContas("ID", CInt(sTMP)).Codigo & ") " & PgDadosPlanoContas("ID", CInt(sTMP)).Descricao)
        .Text = .List(0)
    End With
    With cboFornecedorCC
        sTMP = IIf(Trim(.Text) = "", 0, .Text)
        .Clear
        .AddItem IIf(Trim(pgDadosCentroCustos(CInt(sTMP)).Descricao) = "", " ", sTMP & " - " & pgDadosCentroCustos(CInt(sTMP)).Descricao)
        .Text = .List(0)
    End With
    With cboFornecedorTpDoc
        sTMP = IIf(Trim(.Text) = "", 0, .Text)
        .Clear
        .AddItem IIf(Trim(pgDadosTipoDocumento(CInt(sTMP)).Descricao) = "", " ", sTMP & " - " & pgDadosTipoDocumento(CInt(sTMP)).Descricao)
        .Text = .List(0)
    End With
    With cboBoleto
        sTMP = IIf(Trim(.Text) = "", 0, .Text)
        .Clear
        Select Case sTMP
            Case 1
                .AddItem "1 - Formulario"
            Case 2
                .AddItem "2 - Papel A4"
            Case Else
                .AddItem " "
        End Select
        .Text = .List(0)
    End With
    
    With cboCodProdImpresso
        sTMP = IIf(Trim(.Text) = "", 0, .Text)
        .Clear
        Select Case sTMP
            Case "01"
                .AddItem "01 - Código Interno"
            Case "02"
                .AddItem "02 - Referencia"
            Case Else
                .AddItem " "
        End Select
        .Text = .List(0)
    End With
End Sub




Private Sub txtcDecMoeda_Change()
    
    If IIf(Trim(txtcDecMoeda.Text) = "", 0, txtcDecMoeda.Text) > 5 Then
        MsgBox "Valor entre 0 e 5", vbInformation, "Aviso"
        txtcDecMoeda.Text = ""
    End If
End Sub

Private Sub txtDANFEnCopias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDiasXMLTemp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If

End Sub
Private Sub cboRHConta_DropDown()
    Dim Rst As Recordset
    cboRHConta.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroConta WHERE id_empresa=" & ID_Empresa)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboRHConta.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Agencia") & "/" & Rst.Fields("Conta")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub cboRHDocumento_DropDown()
    Dim Rst As Recordset
    cboRHDocumento.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento WHERE id_empresa=" & ID_Empresa)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboRHDocumento.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If

End Sub
Private Sub cboRHCentroCustos_DropDown()
    Dim Rst As Recordset
    cboRHCentroCustos.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCentroCustos WHERE id_empresa=" & ID_Empresa)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboRHCentroCustos.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If

End Sub


Private Sub txtFinalValidadeCertDigital_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtInicioValidadeCertDigital_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtMailSMTPPorta_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub txtNFePrazoCancelamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If

End Sub
