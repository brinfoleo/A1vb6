VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFaturamentoPV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Pré Venda"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   13470
   Begin VB.Frame Frame17 
      Height          =   2415
      Left            =   7920
      TabIndex        =   50
      Top             =   480
      Width           =   5415
      Begin VB.CommandButton btocobrVisualizarParcelas 
         Caption         =   "..."
         Height          =   255
         Left            =   5100
         TabIndex        =   82
         Top             =   540
         Width           =   255
      End
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   180
         Width           =   2775
      End
      Begin VB.ComboBox cboCondicoesPagamento 
         Height          =   315
         Left            =   2340
         TabIndex        =   9
         Text            =   "cboCondicoesPagamento"
         Top             =   540
         Width           =   2775
      End
      Begin VB.ComboBox cboFormaPagamento 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   900
         Width           =   2775
      End
      Begin VB.TextBox txtPrazoEntrega 
         Height          =   285
         Left            =   2340
         MaxLength       =   60
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1260
         Width           =   2775
      End
      Begin VB.TextBox txtRefCliente 
         Height          =   285
         Left            =   2340
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1620
         Width           =   2775
      End
      Begin VB.TextBox txtValidade 
         Height          =   285
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1980
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Vendedor:"
         Height          =   255
         Left            =   1500
         TabIndex        =   56
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Condições de Pagamento:"
         Height          =   195
         Left            =   300
         TabIndex        =   55
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Forma de Pagamento:"
         Height          =   195
         Left            =   660
         TabIndex        =   54
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Prazo de Entrega:"
         Height          =   195
         Left            =   300
         TabIndex        =   53
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Ref.Cliente:"
         Height          =   195
         Left            =   600
         TabIndex        =   52
         Top             =   1680
         Width           =   1635
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Validade (em dias):"
         Height          =   195
         Left            =   840
         TabIndex        =   51
         Top             =   2040
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   60
      TabIndex        =   27
      Top             =   7200
      Width           =   13335
      Begin VB.TextBox txtTotalPV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11040
         TabIndex        =   89
         Text            =   "R$ 0,00"
         Top             =   840
         Width           =   2115
      End
      Begin VB.TextBox txtvICMSST 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   11280
         TabIndex        =   88
         Text            =   "R$ 0,00"
         Top             =   540
         Width           =   1875
      End
      Begin VB.TextBox txtMercadoria 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   11280
         TabIndex        =   87
         Text            =   "R$ 0,00"
         Top             =   180
         Width           =   1875
      End
      Begin VB.TextBox txtIPI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7080
         TabIndex        =   86
         Text            =   "R$ 0,00"
         Top             =   720
         Width           =   1875
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7080
         TabIndex        =   85
         Text            =   "R$ 0,00"
         Top             =   360
         Width           =   1875
      End
      Begin VB.TextBox txtItens 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   780
         TabIndex        =   84
         Text            =   "000"
         Top             =   360
         Width           =   1875
      End
      Begin VB.TextBox txtOutros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   69
         Text            =   "0,00"
         Top             =   705
         Width           =   1875
      End
      Begin VB.TextBox txtSeguro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   63
         Text            =   "0,00"
         Top             =   345
         Width           =   1875
      End
      Begin VB.TextBox txtFrete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   780
         MaxLength       =   15
         TabIndex        =   61
         Text            =   "0,00"
         Top             =   705
         Width           =   1875
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outros:"
         Height          =   195
         Left            =   3120
         TabIndex        =   90
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor ICMS-ST:"
         Height          =   195
         Left            =   9900
         TabIndex        =   74
         Top             =   570
         Width           =   1275
      End
      Begin VB.Label lblvICMSST 
         Alignment       =   1  'Right Justify
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
         Left            =   11280
         TabIndex        =   73
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total:"
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
         Left            =   10200
         TabIndex        =   72
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblTotalPV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   11040
         TabIndex        =   71
         Top             =   840
         Width           =   2115
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor da Mercadoria:"
         Height          =   195
         Left            =   9600
         TabIndex        =   70
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label lblIPI 
         Alignment       =   1  'Right Justify
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
         Left            =   7080
         TabIndex        =   68
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor IPI:"
         Height          =   195
         Left            =   6180
         TabIndex        =   67
         Top             =   750
         Width           =   795
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descontos:"
         Height          =   195
         Left            =   6120
         TabIndex        =   66
         Top             =   390
         Width           =   855
      End
      Begin VB.Label lblDesconto 
         Alignment       =   1  'Right Justify
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
         Left            =   7080
         TabIndex        =   65
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seguro:"
         Height          =   195
         Left            =   3120
         TabIndex        =   64
         Top             =   390
         Width           =   675
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   750
         Width           =   555
      End
      Begin VB.Label lblItens 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000"
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
         Left            =   780
         TabIndex        =   60
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Itens:"
         Height          =   195
         Left            =   240
         TabIndex        =   59
         Top             =   390
         Width           =   435
      End
   End
   Begin TabDlg.SSTab sstDados 
      Height          =   4155
      Left            =   60
      TabIndex        =   20
      Top             =   3000
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   7329
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&1 - Produtos"
      TabPicture(0)   =   "formFaturamentoPV.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2 - Transporte"
      TabPicture(1)   =   "formFaturamentoPV.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame12"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame12 
         Height          =   3675
         Left            =   -74880
         TabIndex        =   28
         Top             =   360
         Width           =   13095
         Begin VB.Frame Frame16 
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
            Height          =   1875
            Left            =   180
            TabIndex        =   80
            Top             =   1620
            Width           =   8715
            Begin VB.TextBox txtObs 
               Height          =   1455
               Left            =   120
               MaxLength       =   65000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   81
               Text            =   "formFaturamentoPV.frx":0038
               Top             =   300
               Width           =   8475
            End
         End
         Begin VB.Frame Frame15 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3135
            Left            =   9120
            TabIndex        =   36
            Top             =   240
            Width           =   3675
            Begin VB.TextBox txtVol 
               Height          =   315
               Left            =   180
               MaxLength       =   15
               TabIndex        =   37
               Text            =   "Text1"
               Top             =   579
               Width           =   1515
            End
            Begin VB.TextBox txtEsp 
               Height          =   315
               Left            =   1920
               MaxLength       =   60
               TabIndex        =   39
               Text            =   "Text1"
               Top             =   579
               Width           =   1515
            End
            Begin VB.TextBox txtMarca 
               Height          =   315
               Left            =   180
               MaxLength       =   60
               TabIndex        =   41
               Text            =   "Text1"
               Top             =   1305
               Width           =   1515
            End
            Begin VB.TextBox txtNumVol 
               Height          =   315
               Left            =   1920
               MaxLength       =   60
               TabIndex        =   43
               Text            =   "Text1"
               Top             =   1305
               Width           =   1515
            End
            Begin VB.TextBox txtPesoL 
               Height          =   315
               Left            =   180
               MaxLength       =   15
               TabIndex        =   45
               Text            =   "Text1"
               Top             =   2070
               Width           =   1515
            End
            Begin VB.TextBox txtPesoB 
               Height          =   315
               Left            =   1920
               MaxLength       =   15
               TabIndex        =   47
               Text            =   "Text1"
               Top             =   2070
               Width           =   1515
            End
            Begin VB.CheckBox chkAutoSoma 
               Caption         =   "Somar automaticamente as quantidades"
               Height          =   195
               Left            =   180
               TabIndex        =   49
               Top             =   2760
               Width           =   3195
            End
            Begin VB.Label Label28 
               Caption         =   "Quantidade(s)"
               Height          =   195
               Left            =   180
               TabIndex        =   48
               Top             =   270
               Width           =   1155
            End
            Begin VB.Label Label27 
               Caption         =   "Especie"
               Height          =   195
               Left            =   1920
               TabIndex        =   46
               Top             =   270
               Width           =   915
            End
            Begin VB.Label Label26 
               Caption         =   "Marca"
               Height          =   195
               Left            =   180
               TabIndex        =   44
               Top             =   1065
               Width           =   1035
            End
            Begin VB.Label Label25 
               Caption         =   "Numero"
               Height          =   195
               Left            =   1920
               TabIndex        =   42
               Top             =   1035
               Width           =   1215
            End
            Begin VB.Label Label24 
               Caption         =   "Peso Liquido"
               Height          =   195
               Left            =   180
               TabIndex        =   40
               Top             =   1830
               Width           =   1155
            End
            Begin VB.Label Label23 
               Caption         =   "Peso Bruto"
               Height          =   195
               Left            =   1920
               TabIndex        =   38
               Top             =   1830
               Width           =   1155
            End
         End
         Begin VB.Frame Frame13 
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
            Height          =   1275
            Left            =   180
            TabIndex        =   29
            Top             =   240
            Width           =   8715
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Height          =   435
               Left            =   180
               TabIndex        =   75
               Top             =   240
               Width           =   3675
               Begin VB.OptionButton optEntrega 
                  Caption         =   "A Retirar"
                  Height          =   195
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   77
                  Top             =   120
                  Width           =   915
               End
               Begin VB.OptionButton optEntrega 
                  Caption         =   "Entregar"
                  Height          =   195
                  Index           =   1
                  Left            =   2340
                  TabIndex        =   76
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   915
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Modalidade:"
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
                  Left            =   60
                  TabIndex        =   78
                  Top             =   120
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame14 
               BorderStyle     =   0  'None
               Height          =   435
               Left            =   4380
               TabIndex        =   31
               Top             =   180
               Width           =   4155
               Begin VB.OptionButton optFrete 
                  Caption         =   "Emitente"
                  Height          =   195
                  Index           =   0
                  Left            =   1680
                  TabIndex        =   33
                  Top             =   120
                  Width           =   975
               End
               Begin VB.OptionButton optFrete 
                  Caption         =   "Destinatário"
                  Height          =   195
                  Index           =   1
                  Left            =   2700
                  TabIndex        =   32
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   1335
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Frete por conta:"
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
                  Left            =   0
                  TabIndex        =   34
                  Top             =   120
                  Width           =   1635
               End
            End
            Begin VB.ComboBox cboTransportadora 
               Height          =   315
               Left            =   1440
               TabIndex        =   30
               Text            =   "Combo1"
               Top             =   720
               Width           =   5835
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               Caption         =   "Transportador:"
               Height          =   195
               Left            =   180
               TabIndex        =   35
               Top             =   780
               Width           =   1155
            End
         End
      End
      Begin VB.Frame Frame2 
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
         TabIndex        =   21
         Top             =   360
         Width           =   13095
         Begin VB.CommandButton btoMovAcima 
            Height          =   555
            Left            =   12660
            Picture         =   "formFaturamentoPV.frx":003E
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Mover item para cima..."
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton btoMovAbaixo 
            Height          =   555
            Left            =   12660
            Picture         =   "formFaturamentoPV.frx":0728
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Mover item para baixo..."
            Top             =   900
            Width           =   375
         End
         Begin VB.CommandButton btoNovoItem 
            Height          =   555
            Left            =   12660
            Picture         =   "formFaturamentoPV.frx":0E12
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Novo item..."
            Top             =   1620
            Width           =   375
         End
         Begin VB.CommandButton btpExcluirItem 
            Height          =   555
            Left            =   12660
            Picture         =   "formFaturamentoPV.frx":14FC
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   2220
            Width           =   375
         End
         Begin MSFlexGridLib.MSFlexGrid msfgItens 
            Height          =   3255
            Left            =   60
            TabIndex        =   26
            Top             =   180
            Width           =   12555
            _ExtentX        =   22146
            _ExtentY        =   5741
            _Version        =   393216
            Cols            =   23
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"formFaturamentoPV.frx":1886
         End
         Begin VB.Label Label32 
            Caption         =   "Duplo click para editar / <Insert> - Inserir Linha / <Delete> - Remove item / <F7> - Duplica item..."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   3420
            Width           =   9195
         End
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
      Height          =   2415
      Left            =   60
      TabIndex        =   15
      Top             =   480
      Width           =   7755
      Begin VB.ComboBox cboStatusPV 
         Height          =   315
         ItemData        =   "formFaturamentoPV.frx":1A83
         Left            =   5760
         List            =   "formFaturamentoPV.frx":1A85
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   180
         Width           =   1875
      End
      Begin VB.Frame Frame8 
         Caption         =   "Material para:"
         Height          =   855
         Left            =   4920
         TabIndex        =   79
         Top             =   1500
         Width           =   2715
         Begin VB.OptionButton optMaterialPara 
            Caption         =   "Consumo"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   540
            Width           =   1335
         End
         Begin VB.OptionButton optMaterialPara 
            Caption         =   "Industrialização"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   1995
         End
      End
      Begin VB.TextBox txtDoc 
         Height          =   285
         Left            =   1260
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cboUF 
         Height          =   315
         ItemData        =   "formFaturamentoPV.frx":1A87
         Left            =   1260
         List            =   "formFaturamentoPV.frx":1A89
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   915
      End
      Begin VB.TextBox txtTel 
         Height          =   285
         Left            =   1260
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1080
         Width           =   6435
      End
      Begin MSComCtl2.DTPicker dtpEmissao 
         Height          =   315
         Left            =   3600
         TabIndex        =   14
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   122159105
         CurrentDate     =   40517
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1260
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Status:"
         Height          =   255
         Left            =   5220
         TabIndex        =   91
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ/CPF:"
         Height          =   195
         Left            =   180
         TabIndex        =   58
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "UF:"
         Height          =   195
         Left            =   780
         TabIndex        =   57
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   540
         TabIndex        =   18
         Top             =   1140
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Data Emissão:"
         Height          =   195
         Left            =   2520
         TabIndex        =   17
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero:"
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   300
         Width           =   1155
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Pedido"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Romaneio"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clonar Pré-Venda"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Enviar por Email"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pesquisar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":1A8B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":1EDD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":21F7
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":2A89
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":3CDB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":45B5
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":4E47
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":56D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":692B
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":6C45
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":6F5F
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":7356
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":7A50
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPV.frx":814A
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
Dim IdReg           As Integer 'ID do Pedido
Dim iditem          As Integer 'Id dos itens do pedido
Dim IdTransp        As Integer 'Id da transportadora
Dim idCliente       As Integer 'Id do cliente
Dim idCobr          As Integer 'Id Condicoes Pagamento
Dim strTabela       As String
Dim strTabela2      As String
Dim strTabela3      As String
Dim lnPv            As Integer ' Linha da Prevenda
Dim modCalcICMS     As Integer '0 - industrializacao / 1 - consumo
Dim aCob(100)       As Variant 'Registra as parcelas a serem cobradas
Dim cCob            As Integer 'Registra o num de Parcelas

Public Function ClonarPV(pvOriginal As Integer, idPV As Integer, cItens As Integer, qtdItens As Variant) As Integer
    On Error GoTo Terrclone
    'Classe Publica
    'Clona uma PV e retorna o numero da mesma
    'cItens - Quantidade de itens na PV
    'qtdItens () -  array com qtd do item
    'IdPV = 0(zero) Nova PV
    
    
    '1 - Carregar PV
    'Me.Show
    PesquisarRegistro (pvOriginal)
    If IdReg = 0 Then
        ClonarPV = 0
        Exit Function
    End If
    If idPV = 0 Then
        IdReg = 0
    End If
    'Seleciona a linha
    tbMenu.Buttons(1).Enabled = False
    For lnPv = 1 To cItens
    
        iditem = IIf(Trim(msfgItens.TextMatrix(msfgItens.Row, 0)) = "", 0, msfgItens.TextMatrix(msfgItens.Row, 0))
        MovItem (qtdItens(lnPv))
    Next
    
    'formFaturamentoPVItem.txtQuantidade.Text = "50"
    'formFaturamentoPVItem.btoAdicionarItem_Click
    
    
    
    
    '2 - Clona a PV e devolve o numero
    'IdReg = 0
    If grvRegistro = True Then
            ClonarPV = Left(String(6, "0"), 6 - Len(Trim(IdReg))) & IdReg
        Else
            ClonarPV = 0
    End If
    'MsgBox "novaID " & IdReg
    Exit Function
Terrclone:
    MsgBox Err.Description, vbInformation, Err.Number
    
End Function

Private Sub ZerarParcelas()
    idCobr = 0
    cCob = 0
    aCob(cCob) = Array("00/00/0000", ChkVal(txtTotalPV.Text, 0, cDecMoeda))
End Sub

Private Sub cboCliente_KeyPress(KeyAscii As Integer)
    idCliente = 0
End Sub

Private Sub cboCondicoesPagamento_Click()
    If Trim(cboCondicoesPagamento.Text) = "" Then
        ZerarParcelas
        Exit Sub
    End If
    cobrMontarParcelas
End Sub



Private Sub cboCondicoesPagamento_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

'Private Sub Label31_Click()
'    Dim l As Integer
'        For l = 0 To cCob
'        MsgBox "idPV: " & IdReg & vbCrLf & _
'               "Parcela " & l & vbCrLf & _
'               "Vencimento " & aCob(l)(0) & vbCrLf & _
'               "Valor " & aCob(l)(1)
'    Next
'End Sub

Private Sub btocobrVisualizarParcelas_Click()
    Dim NovaCobr As Variant
    If Trim(cboCondicoesPagamento.Text) = "" Then Exit Sub
    cobrMontarParcelas
    formFaturamentoPVPagamento.CarregarFormulario aCob, cCob
End Sub

Private Sub cboStatusPV_DropDown()
    cboStatusPV.Clear
    cboStatusPV.AddItem "01 - Orçamento"
    cboStatusPV.AddItem "02 - Pedido"
End Sub



Private Sub dtpEmissao_Click()
    cobrMontarParcelas
End Sub

Private Sub msfgItens_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim i As Integer
    Dim ii As Integer
    With msfgItens
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
            sTexto = "Descrição: " & UCase(Rst.Fields("Descricao")) & "  " & _
                                "Unidade: " & UCase(Rst.Fields("unidade"))
    End If
    Rst.Close
    pgDescricaoMaterial = sTexto
End Function
Private Sub optMaterialPara_Click(Index As Integer)
    modCalcICMS = Index
End Sub

Private Sub tbMenu_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case LCase(ButtonMenu.Text)
        Case "pedido"
            ImpPV IdReg
        Case "romaneio"
            ImpRomaneio IdReg
    End Select
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
    idCliente = 0
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        PesquisarCliente "Doc", Trim(txtDoc.Text), "S"
    End If
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub
Private Sub QtdAutoSoma()
    Dim i       As Integer
    Dim pesoBL  As String
    For i = 1 To msfgItens.Rows - 1
        pesoBL = Val(ChkVal(pesoBL, 0, cDecQtd)) + Val(ChkVal(msfgItens.TextMatrix(i, 6), 0, cDecQtd))
    Next
    
    pesoBL = ChkVal(pesoBL, 0, 3) 'Convertendo com 03 casas decimais conf. Manual Integracao 4.0.1
    txtPesoB.Text = pesoBL
    txtPesoL.Text = pesoBL
End Sub
Private Sub CalcVlPV()
    Dim VlMercadoria    As String
    Dim vlPV            As String
    Dim VlDesconto      As String
    Dim VlIPI           As String
    Dim itens           As Integer
    Dim vBCICMSST       As String
    Dim vICMSST         As String
    
    VlMercadoria = "0"
    vlPV = "0"
    VlDesconto = "0"
    VlIPI = "0"
    vBCICMSST = "0"
    vICMSST = "0"
    
    
    For itens = 1 To msfgItens.Rows - 1
        If msfgItens.TextMatrix(itens, 21) = "S" Then
            VlMercadoria = Val(ChkVal(msfgItens.TextMatrix(itens, 8), 0, cDecMoeda)) + Val(ChkVal(VlMercadoria, 0, cDecMoeda))
            VlIPI = Val(ChkVal(msfgItens.TextMatrix(itens, 9), 0, cDecMoeda)) + Val(ChkVal(VlIPI, 0, cDecMoeda))
            VlDesconto = Val(ChkVal(msfgItens.TextMatrix(itens, 10), 0, cDecMoeda)) + Val(ChkVal(VlDesconto, 0, cDecMoeda))
        End If
        If ChkVal(msfgItens.TextMatrix(itens, 15), 0, cDecMoeda) <> 0 Then
            vICMSST = Val(ChkVal(msfgItens.TextMatrix(itens, 15), 0, cDecMoeda)) + Val(ChkVal(vICMSST, 0, cDecMoeda))
        End If
        'VlPV = Val(ChkVal(msfgItens.TextMatrix(itens, 10), 0, cDecMoeda)) + Val(ChkVal(VlPV, 0, 2))
    Next
    itens = itens - 1
    
    vlPV = Val(ChkVal(VlMercadoria, 0, cDecMoeda)) + Val(ChkVal(VlIPI, 0, cDecMoeda))
    vlPV = Val(ChkVal(vlPV, 0, cDecMoeda)) + Val(ChkVal(txtFrete.Text, 0, cDecMoeda)) + Val(ChkVal(txtSeguro.Text, 0, cDecMoeda)) + Val(ChkVal(txtOutros.Text, 0, cDecMoeda))
    vlPV = Val(ChkVal(vlPV, 0, 2)) + Val(ChkVal(vICMSST, 0, 2))
    vlPV = Val(ChkVal(vlPV, 0, 2)) - Val(ChkVal(VlDesconto, 0, 2))
    
    txtItens.Text = Left(String(5, "0"), 5 - Len(itens)) & itens
    txtMercadoria.Text = ConvMoeda(VlMercadoria)
    txtIPI.Text = ConvMoeda(VlIPI)
    txtDesconto.Text = ConvMoeda(VlDesconto)
    txtTotalPV.Text = ConvMoeda(vlPV)
    txtvICMSST.Text = ConvMoeda(vICMSST)
    'txtFrete.Text = ConvMoeda(ChkVal(IIf(Trim(txtFrete.Text) = "", 0, txtFrete.Text), 0, 2))
    'txtSeguro.Text = ConvMoeda(ChkVal(IIf(Trim(txtSeguro.Text) = "", 0, txtSeguro.Text), 0, 2))
    'txtOutros.Text = ConvMoeda(ChkVal(IIf(Trim(txtOutros.Text) = "", 0, txtOutros.Text), 0, 2))
    cobrMontarParcelas
End Sub


Private Sub MovItem(Optional nQtd As String)
    '08/12/2014 - Atualizacao para atender a WTL Distribuidora
    ' - Incluir nQtd para que o sistema abra a tela do item da pv inclua
    '   a nova quantidade, faca os calculos e feche o sistema
    
    '09.02.2017 - Atualizacao MRubber
    'Informar se o item sera somado no VTotNFe
    
    If tbMenu.Buttons(1).Enabled = True Then Exit Sub
    Dim Envio(1000)         As Variant 'Armazena os registros
    Dim cReg            As Integer 'Conta os registros
    Dim Retorno         As Variant
    Dim cont            As Integer 'Conta as colunas
    If Trim(cboUF.Text) = "" Then
        MsgBox "Antes de incluir um item, selecione uma UF!", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If lnPv <> 0 Then
            With msfgItens
                iditem = IIf(Trim(msfgItens.TextMatrix(lnPv, 0)) = "", 0, msfgItens.TextMatrix(lnPv, 0))
                cReg = 0
                Envio(cReg) = idCliente & "|" & lnPv & "|" & Trim(cboUF.Text) & "|" & modCalcICMS
                For cont = 0 To .Cols - 1
                    cReg = cReg + 1
                    Envio(cReg) = .TextMatrix(lnPv, cont)
                Next
            End With
        Else
            Envio(0) = idCliente & "|0|" & Trim(cboUF.Text) & "|" & modCalcICMS
    End If
    '08.12.2014- nQtd inclusa para que o sistema faca novos calculos e retorne
    'sem a necessidade de pressionar nenhum bt
    Retorno = formFaturamentoPVItem.CarregarFormulario(Envio, cReg, nQtd)
    
    
    
    'ID|Referencia|Descricao|NCM|Unid|Qtd|vUnit|SubTotal|vIPI|vDesc|vTotal|pIPI|pICMS|pFCP|N.Ped|item ped|Obs
    With msfgItens
        If IsArray(Retorno) Then
            If iditem = 0 And lnPv = 0 Then
                .Rows = .Rows + 1
                lnPv = .Rows - 1
            End If
            .TextMatrix(lnPv, 0) = Retorno(0)
            .TextMatrix(lnPv, 1) = Retorno(1)
            .TextMatrix(lnPv, 2) = Retorno(2)
            .TextMatrix(lnPv, 3) = Retorno(3)
            .TextMatrix(lnPv, 4) = Retorno(4)
            .TextMatrix(lnPv, 5) = Retorno(5)
            .TextMatrix(lnPv, 6) = Retorno(6)
            .TextMatrix(lnPv, 7) = ConvMoeda(ChkVal(IIf(Trim(Retorno(7)) = "", "0", Retorno(7)), 0, cDecMoeda))
            .TextMatrix(lnPv, 8) = Retorno(8)
            .TextMatrix(lnPv, 9) = ConvMoeda(ChkVal(IIf(Trim(Retorno(9)) = "", "0", Retorno(9)), 0, cDecMoeda))
            .TextMatrix(lnPv, 10) = Retorno(10)
            .TextMatrix(lnPv, 11) = Retorno(11)
            .TextMatrix(lnPv, 12) = Retorno(12)
            .TextMatrix(lnPv, 13) = Retorno(13)
            .TextMatrix(lnPv, 14) = ConvMoeda(ChkVal(IIf(Trim(Retorno(19)) = "", "0", Retorno(19)), 0, cDecMoeda))
            .TextMatrix(lnPv, 15) = Retorno(14)
            .TextMatrix(lnPv, 16) = Retorno(21)
            .TextMatrix(lnPv, 17) = Retorno(15)
            .TextMatrix(lnPv, 18) = Retorno(16)
            
            .TextMatrix(lnPv, 19) = Retorno(17)
            .TextMatrix(lnPv, 20) = Retorno(18)
            .TextMatrix(lnPv, 21) = Retorno(20)
        End If
    End With
    CalcVlPV
    'txtItemID.SetFocus
    If chkAutoSoma.Value = 1 Then
        QtdAutoSoma
    End If
End Sub

Private Sub HDFormulario(op As Boolean)
    HDForm Me, op
    'Call optEntrega_Click
    msfgItens.Enabled = True
    If optEntrega(0).Value = True Then
        cboTransportadora.Enabled = False
    End If
End Sub

Private Sub ImprimirPV()
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If
    ImpPV IdReg


End Sub
Private Sub ImprimirRo()
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If
    ImpRomaneio IdReg


End Sub
Private Sub LimparGrid()
    msfgItens.Rows = 1
    txtItens.Text = "0000"
    txtMercadoria.Text = "R$ 0,00"
    txtIPI.Text = "R$ 0,00"
    txtTotalPV.Text = "R$ 0,00"
End Sub

'Private Sub LimpProduto()
'        txtItemID.Text = ""
'        txtProdutoID.Text = ""
'        txtDescricao.Text = ""
'        cboUnidade.Clear
'        txtQuantidade.Text = ""
'        txtValorUnitario.Text = ""
'        txtSubTotalProduto.Text = ""
'        txtAliquotaIPI.Text = ""
'        txtValorIPI.Text = ""
'        txtDescItem.Text = ""
'        txtTotalProduto.Text = ""
'End Sub

Private Sub MontarBaseDeDados()
    Dim vDados(1000)   As Variant
    Dim cReg           As Integer
    Dim i              As Integer
    '##############################################################################################
    '### DADOS DA PRE VENDA
    '##############################################################################################
   cReg = 0
    vDados(cReg) = Array("status", "2", "N"): cReg = cReg + 1
    vDados(cReg) = Array("Emissao", "10", "D"): cReg = cReg + 1
    vDados(cReg) = Array("IdCliente", "10", "N"): cReg = cReg + 1
    vDados(cReg) = Array("Cliente", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("Tel", "30", "S"): cReg = cReg + 1
    vDados(cReg) = Array("CNPJ", "30", "S"): cReg = cReg + 1
    vDados(cReg) = Array("UF", "3", "S"): cReg = cReg + 1
    vDados(cReg) = Array("Transportadora", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("Vendedor", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("CondicoesPagamento", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("FormaPagamento", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("PrazoEntrega", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("RefCliente", "50", "S"): cReg = cReg + 1
    vDados(cReg) = Array("Obs", "65000", "S"): cReg = cReg + 1
    vDados(cReg) = Array("Itens", "10", "S"): cReg = cReg + 1
    vDados(cReg) = Array("VlMercadoria", "50", "L"): cReg = cReg + 1
    vDados(cReg) = Array("VlIPI", "50", "L"): cReg = cReg + 1
    
    vDados(cReg) = Array("bcICMS", "1", "N"): cReg = cReg + 1
    
    vDados(cReg) = Array("vICMSST", "50", "S"): cReg = cReg + 1
    
    vDados(cReg) = Array("Frete", "10", "S"): cReg = cReg + 1
    vDados(cReg) = Array("transp_RetEnt", "10", "S"): cReg = cReg + 1
    
    vDados(cReg) = Array("Seguro", "10", "S"): cReg = cReg + 1
    vDados(cReg) = Array("Outros", "10", "S"): cReg = cReg + 1
    vDados(cReg) = Array("Desconto", "10", "S"): cReg = cReg + 1
    vDados(cReg) = Array("VlTotalPV", "50", "S"): cReg = cReg + 1
    vDados(cReg) = Array("FreteConta", "1", "N"): cReg = cReg + 1
    vDados(cReg) = Array("Validade", "10", "S"): cReg = cReg + 1
    
    vDados(cReg) = Array("transpID", "10", "S"): cReg = cReg + 1
    vDados(cReg) = Array("transp_qVol", "15", "S"): cReg = cReg + 1
    vDados(cReg) = Array("transp_esp", "60", "S"): cReg = cReg + 1
    vDados(cReg) = Array("transp_marca", "60", "S"): cReg = cReg + 1
    vDados(cReg) = Array("transp_nVol", "60", "S"): cReg = cReg + 1
    vDados(cReg) = Array("transp_pesoL", "15", "S"): cReg = cReg + 1
    vDados(cReg) = Array("transp_pesoB", "15", "S"): cReg = cReg + 1
    
    'vDados(cReg) = Array("idLote", "150", "S"): cReg = cReg + 1
    
    cReg = cReg - 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, cReg
    
    '##############################################################################################
    '### ITENS DA PRE VENDA
    '##############################################################################################
    cReg = 0
    vDados(cReg) = Array("idPV", "50", "N"): cReg = cReg + 1
    vDados(cReg) = Array("idProduto", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("referencia", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("descricao", "500", "S"): cReg = cReg + 1
    vDados(cReg) = Array("NCM", "10", "S"): cReg = cReg + 1
    vDados(cReg) = Array("CST", "10", "S"): cReg = cReg + 1
    vDados(cReg) = Array("vICMSST", "15", "S"): cReg = cReg + 1
    vDados(cReg) = Array("vBCICMSST", "15", "S"): cReg = cReg + 1
    '"vBCICMSST"
    vDados(cReg) = Array("pICMS", "10", "S"): cReg = cReg + 1
    
    vDados(cReg) = Array("pICMSFCP", "10", "S"): cReg = cReg + 1
    vDados(cReg) = Array("unidade", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("quantidade", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("ValorUnitario", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("VlItem", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("SubTotal", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("ipi", "10", "S"): cReg = cReg + 1
    vDados(cReg) = Array("VlIPI", "30", "S"): cReg = cReg + 1
    vDados(cReg) = Array("DescItem", "30", "S"): cReg = cReg + 1
    vDados(cReg) = Array("TotalProduto", "30", "S"): cReg = cReg + 1
    vDados(cReg) = Array("Obs", "65000", "S"): cReg = cReg + 1
    vDados(cReg) = Array("ComplDescricaoNFe", "500", "S"): cReg = cReg + 1
        
    vDados(cReg) = Array("nPedido", "30", "S"): cReg = cReg + 1
    vDados(cReg) = Array("iPedido", "30", "S"): cReg = cReg + 1
        
    vDados(cReg) = Array("EstoqueVlCusto", "30", "S"): cReg = cReg + 1
    vDados(cReg) = Array("EstoqueUnidade", "30", "S"): cReg = cReg + 1
    
    vDados(cReg) = Array("destino", "5", "S"): cReg = cReg + 1
    
    cReg = cReg - 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, cReg, "Itens"
    
    '##############################################################################################
    '### PARCELA DAS COBRANCAS
    '##############################################################################################
    cReg = 0
    vDados(cReg) = Array("idPV", "50", "N"): cReg = cReg + 1
    vDados(cReg) = Array("Parcela", "2", "N"): cReg = cReg + 1
    vDados(cReg) = Array("Vencimento", "10", "D"): cReg = cReg + 1
    vDados(cReg) = Array("Valor", "30", "S"): cReg = cReg + 1
    cReg = cReg - 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, cReg, "Cobranca"
End Sub
Private Function grvRegistro() As Boolean
    On Error GoTo TrtErro
    Dim vReg(199)   As Variant
    Dim cReg        As Integer 'Contador de Registros
    Dim l           As Integer
    Dim tmp         As Long
    cReg = 0
    If ValidarPV = False Then
        grvRegistro = False
        Exit Function
    End If
    vReg(cReg) = Array("status", Left(Trim(cboStatusPV.Text), 2), "N"): cReg = cReg + 1
    vReg(cReg) = Array("Emissao", dtpEmissao.Value, "D"): cReg = cReg + 1
    vReg(cReg) = Array("IdCliente", idCliente, "N"): cReg = cReg + 1 'Trim(Left(cboCliente.Text, 6)), "N"): cReg = cReg + 1
    vReg(cReg) = Array("Cliente", Trim(rc(cboCliente.Text)), "S"): cReg = cReg + 1 'Trim(Mid(cboCliente.Text, 10, Len(cboCliente.Text))), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Tel", Trim(txtTel.Text), "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("CNPJ", Trim(txtDoc.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("UF", Trim(cboUF.Text), "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("Transportadora", Trim(Left(cboTransportadora.Text, 6)), "N"): cReg = cReg + 1
    vReg(cReg) = Array("Vendedor", Left(Trim(cboVendedor.Text), 4), "S"): cReg = cReg + 1
    vReg(cReg) = Array("CondicoesPagamento", Trim(Left(cboCondicoesPagamento.Text, 3)), "S"): cReg = cReg + 1
    vReg(cReg) = Array("FormaPagamento", Trim(Left(cboFormaPagamento.Text, 3)), "S"): cReg = cReg + 1
    vReg(cReg) = Array("PrazoEntrega", txtPrazoEntrega.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("RefCliente", txtRefCliente.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Obs", rc(txtObs.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Itens", lblItens.Caption, "S"): cReg = cReg + 1
    vReg(cReg) = Array("VlMercadoria", ChkVal(txtMercadoria.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("vlIPI", ChkVal(txtIPI.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Desconto", ChkVal(txtDesconto.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Frete", ChkVal(txtFrete.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Seguro", ChkVal(txtSeguro.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Outros", ChkVal(txtOutros.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("vICMSST", ChkVal(txtvICMSST.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("bcICMS", modCalcICMS, "N"): cReg = cReg + 1
    
    vReg(cReg) = Array("vlTotalPV", ChkVal(txtTotalPV.Text, 0, cDecMoeda), "S"): cReg = cReg + 1
    
    
    If optFrete(0).Value = True Then
            vReg(cReg) = Array("FreteConta", 0, "N"): cReg = cReg + 1
        Else
            vReg(cReg) = Array("FreteConta", 1, "N"): cReg = cReg + 1
    End If
    If optEntrega(0).Value = True Then
            vReg(cReg) = Array("transp_RetEnt", 0, "N"): cReg = cReg + 1
        Else
            vReg(cReg) = Array("transp_RetEnt", 1, "N"): cReg = cReg + 1
    End If
    vReg(cReg) = Array("Validade", txtValidade.Text, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("transpID", IdTransp, "N"): cReg = cReg + 1
    
    vReg(cReg) = Array("transp_qVol", txtVol.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_esp", txtEsp.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_marca", txtMarca.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_nVol", txtNumVol.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_pesoL", txtPesoL.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("transp_pesoB", txtPesoB.Text, "S") ': cReg = cReg + 1
    
    
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
    
    '*************************************************************************************************************
    '*********************************** GRAVAR DADOS DA COBRANCA ************************************************
    '*************************************************************************************************************
    
    If RegistroExcluir(strTabela3, "idPV = " & IdReg) = False Then
        MsgBox "Erro interno - Ao apagar os dados para novo registro Cobranca!", vbInformation, App.EXEName
        Exit Function
    End If





    cReg = 0
    For l = 0 To cCob
        vReg(cReg) = Array("idPV", IdReg, "S"): cReg = cReg + 1
        vReg(cReg) = Array("Parcela", l, "S"): cReg = cReg + 1
        vReg(cReg) = Array("Vencimento", aCob(l)(0), "D"): cReg = cReg + 1
        vReg(cReg) = Array("Valor", aCob(l)(1), "S"): cReg = cReg + 1
        cReg = cReg - 1
        tmp = RegistroIncluir(strTabela3, vReg, cReg)
        If tmp = 0 Then
                MsgBox "Erro ao Incluir o Cobranca"
                grvRegistro = False
                cReg = 0
            Else
                grvRegistro = True
                cReg = 0
        End If
    Next
    
    
    
    
    '*************************************************************************************************************
    '*********************************** GRAVAR DADOS DA GRADE ***************************************************
    '*************************************************************************************************************
    '12/04/2017
    'Autor: Leonardo Aquino
    'Motivo: Acao modificada pois ao causar erro todos os dados eram perdidos
    '
    cReg = 0
    vReg(cReg) = Array("idPV", "0", "S") ': cReg = cReg + 1
    If RegistroAlterar(strTabela2, vReg, cReg, "idPV=" & IdReg) = False Then
    
    'If RegistroExcluir(strTabela2, "idPV = " & IdReg) = False Then
        MsgBox "Erro interno - Ao apagar os dados para novo registro Produtos"
        grvRegistro = False
        Exit Function
    End If

    

    cReg = 0
    For l = 1 To msfgItens.Rows - 1
        vReg(cReg) = Array("idPV", IdReg, "S"): cReg = cReg + 1
        vReg(cReg) = Array("idProduto", msfgItens.TextMatrix(l, 0), "S"): cReg = cReg + 1
        vReg(cReg) = Array("referencia", msfgItens.TextMatrix(l, 1), "S"): cReg = cReg + 1
        vReg(cReg) = Array("Descricao", rc(msfgItens.TextMatrix(l, 2)), "S"): cReg = cReg + 1

        vReg(cReg) = Array("NCM", msfgItens.TextMatrix(l, 3), "S"): cReg = cReg + 1
        vReg(cReg) = Array("CST", msfgItens.TextMatrix(l, 4), "S"): cReg = cReg + 1
        'vReg(cReg) = Array("pICMS", msfgItens.TextMatrix(l, 13), "S"): cReg = cReg + 1
       
        vReg(cReg) = Array("unidade", msfgItens.TextMatrix(l, 5), "S"): cReg = cReg + 1
        vReg(cReg) = Array("quantidade", msfgItens.TextMatrix(l, 6), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ValorUnitario", ChkVal(msfgItens.TextMatrix(l, 7), 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("vlItem", ChkVal(msfgItens.TextMatrix(l, 8), 0, cDecMoeda), "S"): cReg = cReg + 1
        'vReg(cReg) = Array("SubTotal", ChkVal(Val(ChkVal(msfgItens.TextMatrix(l, 7), 0, 2)) + Val(ChkVal(msfgItens.TextMatrix(l, 8), 0, 2)), 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("SubTotal", ChkVal(msfgItens.TextMatrix(l, 8), 0, 2), "S"): cReg = cReg + 1
        'vReg(cReg) = Array("ipi", msfgItens.TextMatrix(l, 12), "S"): cReg = cReg + 1
        vReg(cReg) = Array("vlipi", ChkVal(msfgItens.TextMatrix(l, 9), 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("DescItem", ChkVal(msfgItens.TextMatrix(l, 10), 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("TotalProduto", ChkVal(msfgItens.TextMatrix(l, 11), 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ipi", msfgItens.TextMatrix(l, 12), "S"): cReg = cReg + 1
        vReg(cReg) = Array("pICMS", msfgItens.TextMatrix(l, 13), "S"): cReg = cReg + 1
        
        vReg(cReg) = Array("vBCICMSST", msfgItens.TextMatrix(l, 14), "S"): cReg = cReg + 1
        vReg(cReg) = Array("vICMSST", msfgItens.TextMatrix(l, 15), "S"): cReg = cReg + 1
        '.TextMatrix(.Rows - 1, 14) = ConvMoeda(IIf(IsNull(Rst.Fields("vICMSST")), "0", Rst.Fields("vICMSST")))
        
        vReg(cReg) = Array("pICMSFCP", Trim(msfgItens.TextMatrix(l, 16)), "S"): cReg = cReg + 1
        
        vReg(cReg) = Array("nPedido", Trim(msfgItens.TextMatrix(l, 17)), "S"): cReg = cReg + 1
        vReg(cReg) = Array("iPedido", Trim(msfgItens.TextMatrix(l, 18)), "S"): cReg = cReg + 1
        
        vReg(cReg) = Array("Obs", msfgItens.TextMatrix(l, 19), "S"): cReg = cReg + 1
 
        'vReg(cReg) = Array("Destino", msfgItens.TextMatrix(l, 20), "S"): cReg = cReg + 1
    
        vReg(cReg) = Array("ComplDescricaoNFe", msfgItens.TextMatrix(l, 20), "S"): cReg = cReg + 1
 
        
        
        If Trim(msfgItens.TextMatrix(l, 0)) <> "" Then
                vReg(cReg) = Array("EstoqueVlCusto", Val(ChkVal(pgDadosEstoqueProduto(msfgItens.TextMatrix(l, 0)).VlCusto, 0, cDecMoeda)), "S"): cReg = cReg + 1
                vReg(cReg) = Array("EstoqueUnidade", pgDadosEstoqueProduto(msfgItens.TextMatrix(l, 0)).Unidade, "S"): cReg = cReg + 1
            Else
                vReg(cReg) = Array("EstoqueVlCusto", ChkVal(0, 0, cDecMoeda), "S"): cReg = cReg + 1
                'vReg(cReg) = Array("EstoqueUnidade", ChkVal(0, 0, 2), "S")
        End If
        
        vReg(cReg) = Array("indTot", IIf(msfgItens.TextMatrix(l, 21) = "S", 1, 0), "N"): cReg = cReg + 1

        cReg = cReg - 1
               
        
               
        tmp = RegistroIncluir(strTabela2, vReg, cReg)
        If tmp = 0 Then
                MsgBox "Erro ao Incluir o Produto"
                grvRegistro = False
                cReg = 0
                Exit Function
            Else
                grvRegistro = True
                cReg = 0
        End If
    Next
    '*** Fim  da gravacao da grade
    
    '* 12/04/2017
    '* Autor: Leonardo Aquino
    '* Motivo: Apos tudo ser gravado corretamente o sistema limpa os itens velhos
    If RegistroExcluir(strTabela2, "idPV = 0 AND UsuId=" & ID_Usuario) = False Then
        MsgBox "Erro interno - Erro ao apagar itens obsoletos da tab itens"
        grvRegistro = False
        Exit Function
    End If

    
    
    
    
 Exit Function
TrtErro:
    MsgBox Err.Number, vbInformation, Err.Description
    Resume Next
End Function
Public Sub PesquisarRegistro(Optional Id As Integer)
On Error GoTo TrtErro
    If Trim(Id) = 0 Then
            IdReg = formBuscar.IniciarBusca(strTabela, "cliente,emissao")
        Else
            IdReg = Id
    End If
    LimpForm
    If IdReg = 0 Then
            Exit Sub
        Else
            Dim sSQL    As String
            Dim Rst     As Recordset
    
            sSQL = "SELECT * FROM FaturamentoPV WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg
    
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    MsgBox "Registro nao encontrado.", vbInformation, "Aviso"
                    LimpForm
                    Rst.Close
                    Exit Sub
                Else
                    'Filtra se outros usuarios podem ver as vendas de outro vendedor
                    If PgDadosUsuario(ID_Usuario).SuperUsuario = 0 Then 'Checa se e super usuario
                        If PgDadosConfig.VisualizarOutrosFunc = 0 Then
                            If CInt(Rst.Fields("Vendedor")) <> CInt(Left(PgDadosUsuario(ID_Usuario).Nome, 3)) Then
                                MsgBox "Somente sera permitido visualizar os dados da pre-venda pelo seu emissor (" & PgDadosRhFuncionario(Rst.Fields("Vendedor")).Nome & ")!", vbInformation, "Aviso"
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    txtID.Text = Rst.Fields("id")
                    
                    cboStatusPV_DropDown
                    
                    Select Case Rst.Fields("status")
                        Case "1"
                            cboStatusPV.Text = cboStatusPV.List(0)
                        Case "2"
                            cboStatusPV.Text = cboStatusPV.List(1)
                        Case Else
                            cboStatusPV.Clear
    
                    End Select
                    
                    
                    dtpEmissao.Value = Rst.Fields("Emissao")
                    
                    idCliente = Rst.Fields("IdCliente")
                    
                    cboCliente.Clear
                    cboCliente.AddItem IIf(IsNull(Rst.Fields("Cliente")), " ", Rst.Fields("Cliente"))
                    cboCliente.Text = cboCliente.List(0)
                    
                    txtTel.Text = cNull(Rst.Fields("Tel"))
                    txtDoc.Text = cNull(Rst.Fields("CNPJ"))
                    With cboUF
                        .Clear
                        .AddItem IIf(Trim(cNull(Rst.Fields("UF"))) = "", " ", cNull(Rst.Fields("UF")))
                        .Text = .List(0)
                    End With
                    cboTransportadora.Clear
                    If Not IsNull(Rst.Fields("Transportadora")) Then
                        cboTransportadora.AddItem IIf(IsNull(Rst.Fields("Transportadora")), " ", _
                                                   Left(String(6, "0"), 6 - Len(Rst.Fields("Transportadora"))) & Rst.Fields("Transportadora") & " - " & pgDadosTransportadora(Rst.Fields("Transportadora")).Nome)
                        cboTransportadora.Text = cboTransportadora.List(0)
                    End If
                    
                    cboVendedor.Clear
                    If Not IsNull(Rst.Fields("Vendedor")) Then
                        cboVendedor.AddItem Rst.Fields("Vendedor") & " - " & PgDadosRhFuncionario(Rst.Fields("Vendedor")).Nome
                        cboVendedor.Text = cboVendedor.List(0)
                    End If
                    
                    cboCondicoesPagamento.Clear
                    If Not IsNull(Rst.Fields("CondicoesPagamento")) Then
                        cboCondicoesPagamento.Text = Rst.Fields("CondicoesPagamento") & " - " & pgDescrCondPag(Rst.Fields("CondicoesPagamento"))
                        'cboCondicoesPagamento.AddItem Rst.Fields("CondicoesPagamento") & " - " & pgDescrCondPag(Rst.Fields("CondicoesPagamento"))
                        'cboCondicoesPagamento.Text = cboCondicoesPagamento.List(0)
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
                    If Rst.Fields("FreteConta") = 1 Then
                            '1 - Destinatario
                            optFrete(1).Value = True
                        Else
                            '0 - Emitente
                            optFrete(0).Value = True
                    End If
                    If Rst.Fields("transp_RetEnt") = 0 Then
                            '0 - Retira
                            optEntrega(0).Value = True
                        Else
                            '1 - Entrega
                            optEntrega(1).Value = True
                    End If
                    cboTransportadora.Clear
                    If Rst.Fields("Transportadora") <> 0 Then
                        cboTransportadora.AddItem Left(String(6, "0"), 6 - Len(Rst.Fields("Transportadora"))) & Rst.Fields("Transportadora") & " - " & pgDadosTransportadora(Rst.Fields("Transportadora")).Nome
                        cboTransportadora.Text = cboTransportadora.List(0)
                    End If
                    txtDesconto.Text = ConvMoeda(IIf(IsNull(Rst.Fields("Desconto")), "0,00", Rst.Fields("Desconto")))
                    txtFrete.Text = ConvMoeda(IIf(IsNull(Rst.Fields("Frete")), "0,00", Rst.Fields("Frete")))
                    txtSeguro.Text = ConvMoeda(IIf(IsNull(Rst.Fields("Seguro")), "0,00", Rst.Fields("Seguro")))
                    txtOutros.Text = ConvMoeda(IIf(IsNull(Rst.Fields("outros")), "0,00", Rst.Fields("Outros")))
                    
                    txtIPI.Text = ConvMoeda(IIf(IsNull(Rst.Fields("vlIPI")), "0,00", Rst.Fields("vlIPI")))
                    
                    txtMercadoria.Text = ConvMoeda(IIf(IsNull(Rst.Fields("vlMercadoria")), "0,00", Rst.Fields("vlMercadoria")))
                    txtvICMSST.Text = ConvMoeda(IIf(IsNull(Rst.Fields("vICMSST")), "0,00", Rst.Fields("vICMSST")))
                    txtTotalPV.Text = ConvMoeda(IIf(IsNull(Rst.Fields("vlTotalPV")), "0,00", Rst.Fields("vlTotalPV")))
                    
                    modCalcICMS = IIf(IsNull(Rst.Fields("bcICMS")), 0, Rst.Fields("bcICMS"))
                    optMaterialPara(modCalcICMS).Value = True
                    
                    txtVol.Text = IIf(IsNull(Rst.Fields("transp_qVol")), "", Rst.Fields("transp_qVol"))
                    txtEsp.Text = IIf(IsNull(Rst.Fields("transp_Esp")), "", Rst.Fields("transp_Esp"))
                    txtMarca.Text = IIf(IsNull(Rst.Fields("transp_Marca")), "", Rst.Fields("transp_Marca"))
                    txtNumVol.Text = IIf(IsNull(Rst.Fields("transp_nVol")), "", Rst.Fields("transp_nVol"))
                    txtPesoL.Text = IIf(IsNull(Rst.Fields("transp_PesoL")), "", Rst.Fields("transp_PesoL"))
                    txtPesoB.Text = IIf(IsNull(Rst.Fields("transp_PesoB")), "", Rst.Fields("transp_PesoB"))
                    
            End If
            Rst.Close
            
            '**********************************************************************************************************
            '************************************* Carregar Cobranca **************************************************
            '**********************************************************************************************************
            sSQL = "SELECT * FROM " & strTabela3 & " WHERE ID_Empresa = " & ID_Empresa & " AND IdPV = " & IdReg & " ORDER BY Parcela"
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    'LimparGrid
                Else
                    Rst.MoveFirst
                    'With msfgItens
                        cCob = 0
                        Do Until Rst.EOF
                            aCob(cCob) = Array(CStr(cNull(Rst.Fields("Vencimento"))), CStr(Rst.Fields("Valor"))): cCob = cCob + 1
                            '.TextMatrix(.Rows - 1, 0) = IIf(IsNull(Rst.Fields("IDproduto")), "", Rst.Fields("idProduto"))
                            '.TextMatrix(.Rows - 1, 19) = IIf(IsNull(Rst.Fields("ComplDescricaoNFe")), "", Rst.Fields("ComplDescricaoNFe"))
                            Rst.MoveNext
                        Loop
                        cCob = cCob - 1
                        'txtItens.text= ZE(.Rows - 1, 5)
                    'End With
                    
            End If
            Rst.Close
            
            
            '**********************************************************************************************************
            '************************************* Carregar Corpo do PV ***********************************************
            '**********************************************************************************************************
            sSQL = "SELECT * FROM " & strTabela2 & " WHERE ID_Empresa = " & ID_Empresa & " AND IdPV = " & IdReg
    
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    LimparGrid
                Else
                    Rst.MoveFirst
                    With msfgItens
                        .Rows = 1
                        Do Until Rst.EOF
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, 0) = IIf(IsNull(Rst.Fields("IDproduto")), "", Rst.Fields("idProduto"))
                            .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rst.Fields("Referencia")), " ", Rst.Fields("Referencia"))
                            .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rst.Fields("Descricao")), "", Rst.Fields("Descricao"))
                            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("NCM")), "", Rst.Fields("NCM"))
                            .TextMatrix(.Rows - 1, 4) = cNull(Rst.Fields("CST"))
                            .TextMatrix(.Rows - 1, 5) = IIf(IsNull(Rst.Fields("Unidade")), "", Rst.Fields("Unidade"))
                            .TextMatrix(.Rows - 1, 6) = IIf(IsNull(Rst.Fields("Quantidade")), "0", Rst.Fields("Quantidade"))
                            .TextMatrix(.Rows - 1, 7) = ConvMoeda(Rst.Fields("ValorUnitario"))
                            .TextMatrix(.Rows - 1, 8) = ConvMoeda(IIf(IsNull(Rst.Fields("vlItem")), "0", Rst.Fields("vlItem")))
                            .TextMatrix(.Rows - 1, 9) = ConvMoeda(Rst.Fields("vlIPI"))
                            .TextMatrix(.Rows - 1, 10) = ConvMoeda(IIf(IsNull(Rst.Fields("DescItem")), "0,00", Rst.Fields("DescItem")))
                            .TextMatrix(.Rows - 1, 11) = ConvMoeda(IIf(IsNull(Rst.Fields("TotalProduto")), "0,00", Rst.Fields("TotalProduto")))
                            .TextMatrix(.Rows - 1, 12) = IIf(IsNull(Rst.Fields("IPI")), "0", Rst.Fields("IPI"))
                            .TextMatrix(.Rows - 1, 13) = IIf(IsNull(Rst.Fields("pICMS")), "0", Rst.Fields("pICMS"))
                            
                            .TextMatrix(.Rows - 1, 14) = ConvMoeda(IIf(IsNull(Rst.Fields("vBCICMSST")), "0", Rst.Fields("vBCICMSST")))
                            .TextMatrix(.Rows - 1, 15) = ConvMoeda(IIf(IsNull(Rst.Fields("vICMSST")), "0", Rst.Fields("vICMSST")))
                            .TextMatrix(.Rows - 1, 16) = IIf(IsNull(Rst.Fields("pICMSFCP")), "0", Rst.Fields("pICMSFCP"))
                            .TextMatrix(.Rows - 1, 17) = IIf(IsNull(Rst.Fields("nPedido")), "", Rst.Fields("nPedido"))
                            .TextMatrix(.Rows - 1, 18) = IIf(IsNull(Rst.Fields("iPedido")), "", Rst.Fields("iPedido"))
                            .TextMatrix(.Rows - 1, 19) = IIf(IsNull(Rst.Fields("Obs")), "", Rst.Fields("Obs"))
                            .TextMatrix(.Rows - 1, 20) = IIf(IsNull(Rst.Fields("ComplDescricaoNFe")), "", Rst.Fields("ComplDescricaoNFe"))
                            
                            '.TextMatrix(.Rows - 1, 20) = IIf(IsNull(Rst.fields("destino")), "", Rst.fields("destino"))
                            
                            .TextMatrix(.Rows - 1, 21) = IIf(cNull(Rst.Fields("indTot")) = 0, "N", "S")
                            Rst.MoveNext
                        Loop
                        txtItens.Text = ZE(.Rows - 1, 5)
                    End With
                    
            End If
            Rst.Close
    End If
    Exit Sub
TrtErro:
    MsgBox "Erro - verifique arquivo de log.", vbCritical, App.EXEName
    RegLog "", "", "[formFaturamentoPv.PesquisarRegistro] " & Err.Number & " - " & Err.Description
End Sub
Private Sub btoMovAbaixo_Click()
    If lnPv = 0 Then Exit Sub
    MoveRow lnPv, "Down"
    lnPv = 0
End Sub
Private Sub btoMovAcima_Click()
    If lnPv = 0 Then Exit Sub
    MoveRow lnPv, "UP"
    lnPv = 0
End Sub
Private Sub btoNovoItem_Click()
    iditem = 0
    lnPv = 0
    MovItem
End Sub
Private Sub btpExcluirItem_Click()
    RemoverItem
End Sub
Private Sub cboCliente_Click()
    If Trim(cboCliente.Text) = "" Then
        idCliente = 0
        Exit Sub
    End If
    PesquisarCliente "ID", Trim(Left(Trim(cboCliente.Text), 6)), "N"
End Sub

Private Sub cboCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarCliente
    End If
End Sub

Private Sub cboTransportadora_Click()
    If Trim(cboTransportadora.Text) = "" Then Exit Sub
    IdTransp = Left(Trim(cboTransportadora.Text), 6)
End Sub

Private Sub cboTransportadora_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarTransp
    End If
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

Private Sub chkAutoSoma_Click()
    If chkAutoSoma.Value = 1 Then
        QtdAutoSoma
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

Private Sub PesquisarCliente(Optional sCampo As String, Optional sBusca As String, Optional SN As String)
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    If Trim(sCampo) = "" Then
        sBusca = formBuscar.IniciarBusca("Clientes", , , , , "Status='Ativo'") ', "xNome,xlgr,nro,xcpl,xbairro,xmun,uf,fone")
        sCampo = "Id"
        SN = "N"
        If Trim(sBusca) = 0 Then Exit Sub
    End If
    If SN = "N" Then
            sSQL = "SELECT * FROM Clientes WHERE ID_Empresa = " & ID_Empresa & " AND Status='Ativo' AND " & sCampo & " = '" & sBusca & "'"
        Else
            sSQL = "SELECT * FROM Clientes WHERE ID_Empresa = " & ID_Empresa & " AND Status='Ativo' AND " & sCampo & " = " & sBusca
    End If
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado"
        Else
            Rst.MoveFirst
            idCliente = Rst.Fields("Id")
            cboCliente.Text = Trim(Rst.Fields("xNome"))
            txtDoc.Text = cNull(Rst.Fields("Doc"))
            txtTel.Text = cNull(Rst.Fields("Fone"))
            
            With cboUF
                .Clear
                .AddItem cNull(Rst.Fields("UF"))
                .Text = .List(0)
            End With
            
            If Trim(cboVendedor.Text) = "" Then
                If Not IsNull(Rst.Fields("vendedor")) Then
                    cboVendedor.Clear
                    cboVendedor.AddItem ZE(Rst.Fields("vendedor"), 4) & " - " & PgDadosRhFuncionario(Rst.Fields("vendedor")).Nome
                    cboVendedor.Text = cboVendedor.List(0)
                End If
            End If
            If Trim(cboCondicoesPagamento.Text) = "" Then
                If Not IsNull(Rst.Fields("CondicoesPagamento")) And Trim(Rst.Fields("CondicoesPagamento")) <> 0 Then
                    cboCondicoesPagamento.Clear
                    cboCondicoesPagamento.AddItem ZE(Rst.Fields("CondicoesPagamento"), 3) _
                                                    & " - " & pgDescrCondPag(Rst.Fields("CondicoesPagamento"))
                    cboCondicoesPagamento.Text = cboCondicoesPagamento.List(0)
                End If
            End If
            If Trim(cboFormaPagamento.Text) = "" Then
                If Not IsNull(Rst.Fields("TipoDocumento")) And Trim(Rst.Fields("TipoDocumento")) <> 0 Then
                    cboFormaPagamento.Clear
                    cboFormaPagamento.AddItem ZE(Rst.Fields("tipodocumento"), 3) _
                                            & " - " & pgDadosTipoDocumento(Rst.Fields("tipodocumento")).Descricao
                    cboFormaPagamento.Text = cboFormaPagamento.List(0)
                End If
            End If
            If Trim(cboTransportadora.Text) = "" Then
                If Not IsNull(Rst.Fields("Transportadora")) And Trim(Rst.Fields("Transportadora")) <> 0 Then
                    cboTransportadora.Clear
                    IdTransp = Rst.Fields("Transportadora")
                    cboTransportadora.AddItem ZE(Trim(IdTransp), 6) _
                                            & " - " & pgDadosTransportadora(Rst.Fields("Transportadora")).Nome
                    cboTransportadora.Text = cboTransportadora.List(0)
                    
                End If
            End If
            
    End If
    Rst.Close
End Sub

Private Sub RemoverItem()
    If lnPv = 0 Then Exit Sub
    If MsgBox("Deseja realmente remover este item?", vbYesNo, "Removendo Item do Pedido") = vbYes Then
        If msfgItens.Rows = 2 Then
                msfgItens.Rows = 1
                iditem = 0
                lnPv = 0
            Else
                msfgItens.RemoveItem msfgItens.Row
                iditem = 0
                lnPv = 0
        End If
        iditem = 0
        lnPv = 0
    End If
    CalcVlPV
End Sub
Private Sub cboCliente_DropDown()
    Dim Rst As Recordset
    idCliente = 0
    Set Rst = RegistroBuscar("SELECT * FROM Clientes WHERE ID_Empresa = " & ID_Empresa & " AND xNome LIKE '" & cboCliente.Text & "%'")
    If Rst.BOF And Rst.EOF Then
            cboCliente.Clear
            Exit Sub
        Else
            cboCliente.Clear
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCliente.AddItem Left(String(6, "0"), 6 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & _
                                   " - " & _
                                   Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If
End Sub
Private Sub cbotransportadora_DropDown()
    Dim Rst As Recordset
    
    Set Rst = RegistroBuscar("SELECT * FROM Transportadoras WHERE ID_Empresa = " & ID_Empresa & " AND xNome LIKE '" & cboTransportadora.Text & "%'")
    If Rst.BOF And Rst.EOF Then
            cboTransportadora.Clear
            Exit Sub
        Else
            cboTransportadora.Clear
            Rst.MoveFirst
            Do Until Rst.EOF
                cboTransportadora.AddItem Left(String(6, "0"), 6 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & _
                " - " & _
                Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub PesquisarTransp(Optional Id As String)
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim IdTransp    As String
    
    If Trim(Id) = "" Then
            IdTransp = formBuscar.IniciarBusca("Transportadoras")
            If Trim(IdTransp) = 0 Then Exit Sub
        Else
            IdTransp = Id
    End If
    
    sSQL = "SELECT * FROM Transportadoras WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdTransp
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado"
        Else
            Rst.MoveFirst
            cboTransportadora.AddItem Left(String(6, "0"), 6 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                               Rst.Fields("xNome")
            cboTransportadora.Text = cboTransportadora.List(0)
    End If
    Rst.Close
End Sub
Private Sub cboCondicoesPagamento_DropDown()
    Dim Rst As Recordset
    cboCondicoesPagamento.Clear
    ZerarParcelas
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCondicoespagamento WHERE id_empresa=" & ID_Empresa & " ORDER BY Descricao")
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
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento WHERE id_empresa=" & ID_Empresa)
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
Private Sub cboVendedor_DropDown()
    Dim Rst As Recordset
    cboVendedor.Clear
    Set Rst = RegistroBuscar("SELECT * FROM RHFuncionarioCadastro WHERE id_empresa=" & ID_Empresa & " ORDER BY xNome")
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

Private Sub Form_Load()

    Me.Top = 0
    Me.Left = 0
    LimpForm
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    strTabela2 = strTabela & "itens"
    strTabela3 = strTabela & "Cobranca"
    HDFormulario (False)
    HDMenu Me, True
    txtID.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    IdReg = 0
End Sub
Private Sub msfgItens_Click()
    If msfgItens.TextMatrix(msfgItens.Row, 0) = "ID" Or msfgItens.Rows = 1 Then Exit Sub
    lnPv = msfgItens.Row
End Sub
Private Sub msfgItens_DblClick()
    If msfgItens.TextMatrix(msfgItens.Row, 0) = "ID" Or msfgItens.Rows = 1 Then Exit Sub
    lnPv = msfgItens.Row
    iditem = IIf(Trim(msfgItens.TextMatrix(msfgItens.Row, 0)) = "", 0, msfgItens.TextMatrix(msfgItens.Row, 0))
    MovItem
End Sub

Private Sub msfgItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If tbMenu.Buttons(1).Enabled = True Then Exit Sub
    If msfgItens.TextMatrix(msfgItens.Row, 0) = "ID" Or msfgItens.Rows = 1 Then Exit Sub 'Or Trim(msfgItens.TextMatrix(msfgItens.Row, 2)) = "" Then Exit Sub
    If KeyCode = 45 Then 'Adicionar linha em branco
        msfgItens.AddItem "", msfgItens.Row
    End If
    If KeyCode = 46 Then 'Tecla Delete
        iditem = IIf(msfgItens.TextMatrix(msfgItens.Row, 0) = "", 0, msfgItens.TextMatrix(msfgItens.Row, 0))
        RemoverItem
    End If
    If KeyCode = 118 Then 'F7 Clonar o item
        With msfgItens
            .Rows = .Rows + 1
            For i = 0 To .Cols - 1
                .TextMatrix(.Rows - 1, i) = .TextMatrix(lnPv, i)
            Next
        End With
    End If
    'CalcVlPV
End Sub
Private Sub optEntrega_Click(Index As Integer)
    If optEntrega(0).Enabled = False Then Exit Sub
    If Index = 0 Then
            cboTransportadora.Enabled = False
            cboTransportadora.Clear
            IdTransp = 0
            optFrete(1).Value = True
        Else
            cboTransportadora.Enabled = True
            optFrete(0).Value = True
    End If
End Sub
Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    IdReg = 0
    ZerarParcelas
    HDMenu Me, False
    HDFormulario (True)
    LimpForm
    msfgItens.Rows = 1
    txtID.Enabled = False
    txtDesconto.Enabled = False
    
'    cboStatusPV.AddItem "02 - Pedido"
'    cboStatusPV.Text = cboStatusPV.List(0)
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
                        "Pre Venda.: " & txtID.Text, vbYesNo + vbQuestion) = vbYes Then
                If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                    RegistroExcluir strTabela2, "Id = " & IdReg
                    LimpForm
                End If
            End If
    End If
End Sub
Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
           Incluir
        Case "Alterar"
            AlterarPV
        Case "Excluir"
            Excluir
        Case "Imprimir"
            ImprimirPV
        Case "Pesquisar"
            PesquisarRegistro
        Case "Clonar Pré-Venda"
            ClonarPedido
        Case "Enviar por Email"
            EnviarViaEmail
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDFormulario (False)
                txtID.Enabled = True
                txtID.Text = IdReg
            End If
            
        
        Case "Cancelar"
            Cancelar
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select
End Sub
Private Sub EnviarViaEmail()
    
    If ID_Usuario <> "2" Then
        MsgBox "Função incompleta! Por favor aguarde atualizações futuras." & vbCrLf _
        & vbCrLf & vbCrLf & _
        "DTI - Departamento de TI", vbInformation, App.EXEName
        Exit Sub
    End If
    
    Dim Rst                 As Recordset
    Dim RstI                As Recordset
    Dim sSQL                As String
    
    
    'Dim linha               As String * 80
    '##################################################################################################
    '### 27/03/2012 - Dados da Pre Venda
    
    'Cabecalho
    Dim cNome               As String
    Dim cCNPJ               As String
    Dim cEnd                As String
    Dim cTransp             As String
    Dim cTotalPV            As String
    
    'Itens
    Dim iQtd                 As String
    Dim iUnid                As String
    Dim iDescricao           As String
    Dim iNCM                 As String
    Dim iICMS                As String
    Dim iIPI                 As String
    Dim ivUnitario           As String
    Dim iTotalProd           As String
    Dim msg                  As String
    '##################################################################################################
    
    If IdReg = 0 Then Exit Sub
    'linha = String(80, "=")
    
    sSQL = "SELECT * FROM FaturamentoPV WHERE ID=" & IdReg
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar Pre-Venda número " & IdReg & "!", vbInformation, App.EXEName
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
    End If
    cNome = cNull(Rst.Fields("Cliente"))
    cCNPJ = cNull(Rst.Fields("CNPJ"))
    cEnd = ""
    cTransp = pgDadosTransportadora(Trim(IIf(Trim(cNull(Rst.Fields("TranspID"))) = "", "0", cNull(Rst.Fields("TranspID"))))).Nome
    cTotalPV = cNull(Rst.Fields("vlTotalPV"))
    
    'Dados da Empresa
    
    msg = "<HTML><HEAD><TITLE>PROPOSTA N." & ZE(IdReg, 6) & "</TITLE><HEAD><BODY>"
    msg = msg & "<H2>" & PgDadosEmpresa(ID_Empresa).Nome & "</H2>"
    msg = msg & "<FONT SIZE=1>"
    msg = msg & PgDadosEmpresa(ID_Empresa).Lgr & ", " & PgDadosEmpresa(ID_Empresa).Nro
    msg = msg & "- " & PgDadosEmpresa(ID_Empresa).Bairro & "<BR>" & PgDadosEmpresa(ID_Empresa).Mun & "/" & PgDadosEmpresa(ID_Empresa).uf & ". CEP:" & PgDadosEmpresa(ID_Empresa).CEP & "<BR>Tel.:" & PgDadosEmpresa(ID_Empresa).Fone & "<BR>email: " & PgDadosEmpresa(ID_Empresa).Mail & "</FONT>"
    
    msg = msg & "<BR><H1><CENTER>PRE-VENDA N." & ZE(IdReg, 6) & "</CENTER></H1>"
    msg = msg & "<BR>Cliente: <B>" & cNome & "</B><BR>Endereço: <B>" & cEnd & "</B><BR>CNPJ/CPF: <B>" & cCNPJ & "</B><BR>"
    msg = msg & "<BR><BR>"
    
    'MsgCorpo = PgDadosEmpresa(ID_Empresa).Nome & vbCrLf & _
             "CNPJ: " & PgDadosEmpresa(ID_Empresa).CNPJ & "   " & _
             "I.E." & PgDadosEmpresa(ID_Empresa).IE & vbCrLf & _
             PgDadosEmpresa(ID_Empresa).Lgr & " " & _
             PgDadosEmpresa(ID_Empresa).Nro & " - " & _
             PgDadosEmpresa(ID_Empresa).Bairro & " " & _
             PgDadosEmpresa(ID_Empresa).Mun & "/" & _
             PgDadosEmpresa(ID_Empresa).UF & vbCrLf & _
             "Fone: " & PgDadosEmpresa(ID_Empresa).Fone & vbCrLf & _
             vbCrLf & _
             vbCrLf

    'MsgCorpo = MsgCorpo & _
               "PRE-VENDA" & vbCrLf & _
               "Data: " & Rst.fields("Emissao") & "            Número: " & Left(String(5, "0"), 5 - Len(IdReg)) & IdReg & vbCrLf & _
               vbCrLf & _
               vbCrLf
               
               
    'Dados do Cliente
    'MsgCorpo = MsgCorpo & _
               "Cliente: " & cNome & vbCrLf & _
               "CNPJ: " & cCNPJ & vbCrLf & _
               "REF: " & cNull(Rst.fields("RefCliente")) & vbCrLf & _
               vbCrLf & _
               "Transportadora: " & cTransp & vbCrLf & _
               "Frete por conta: " & IIf(cNull(Rst.fields("FreteConta")) = 0, "EMITENTE", "DESTINATARIO") & vbCrLf & _
               vbCrLf & _
               "Condições de Pagamento: " & IIf(IsNull(Rst.fields("CondicoesPagamento")), "", pgDescrCondPag(Rst.fields("CondicoesPagamento"))) & IIf(IsNull(Rst.fields("FormaPagamento")), "", " (" & pgDescrTipoDoc(Rst.fields("FormaPagamento")) & ")") & vbCrLf & _
               vbCrLf & _
               "Observações: " & IIf(IsNull(Rst.fields("Obs")), "", Rst.fields("Obs")) & vbCrLf
                

    
    
    
    sSQL = "SELECT * FROM FaturamentoPVItens WHERE IdPV=" & IdReg
    Set RstI = RegistroBuscar(sSQL)
    If RstI.BOF And RstI.EOF Then
            MsgBox "Erro ao localizar os itens da Pre-Venda!", vbInformation, App.EXEName
            RstI.Close
        Else
            'MsgCorpo = MsgCorpo & _
                       vbCrLf & _
                       vbCrLf & _
                       String(80, "=") & vbCrLf & _
                       "DESCRIÇÃO" & vbCrLf & _
                       String(80, "=") & vbCrLf


        
            RstI.MoveFirst
            'Campos da Listagem
            msg = msg & "<TABLE>"
            'msg = msg & "<TH>"
            msg = msg & "<TH WIDTH=""5%"" ALIGN=""center"">Quantidade</TH>"
            msg = msg & "<TH WIDTH=""5%"" ALIGN=""center"">Unid.</TH>"
            msg = msg & "<TH WIDTH=""50%"">Descrição</TH>"
            msg = msg & "<TH WIDTH=""15%"" ALIGN=""right"">Vl.Unit.</TH>"
            msg = msg & "<TH WIDTH=""10%"" ALIGN=""right"">IPÌ(%)</TH>"
            msg = msg & "<TH WIDTH=""15%"" ALIGN=""right"">Subtotal</TH>"
            'msg = msg & "</TR>"
    
            Do Until RstI.EOF
                iQtd = cNull(RstI.Fields("Quantidade"))
                iUnid = cNull(RstI.Fields("Unidade"))
                iDescricao = cNull(RstI.Fields("Descricao"))
                iNCM = cNull(RstI.Fields("NCM"))
                iICMS = ConvMoeda(cNull(RstI.Fields("pICMS")))
                iIPI = ConvMoeda(cNull(RstI.Fields("IPI")))
                ivUnitario = cNull(RstI.Fields("ValorUnitario"))
                'iTotalProd = Space(200)
                'iTotalProd = Right(iTotalProd, Len(Trim(ConvMoeda(cNull(RstI.Fields("TotalProduto")))))) & ConvMoeda(cNull(RstI.Fields("TotalProduto")))
                iTotalProd = ConvMoeda(cNull(RstI.Fields("TotalProduto")))
                iTotalProd = Mid(String(120, " "), 1, Len(Trim(iTotalProd))) & Trim(iTotalProd)
                
                'MsgCorpo = MsgCorpo & _
                           vbCrLf & _
                           iDescricao & _
                           vbCrLf & _
                           "NCM: " & iNCM & "      " & _
                           "ICMS(%): " & iICMS & "     " & _
                           "IPI(%): " & iIPI & _
                           vbCrLf & _
                           iUnid & " - " & ChkVal(iQtd, 0, cDecQtd) & " x " & ConvMoeda(ivUnitario) & " " & vbTab & vbTab & vbTab & vbTab & iTotalProd & _
                           vbCrLf
                           
    
    
                'Lista os materiais da PV
               
                msg = msg & "<TH></TH><TR>"
                msg = msg & "<TD WIDTH=""5%"" ALIGN=""center""><I>" & iQtd & "</I></TD>"
                msg = msg & "<TD WIDTH=""5%"" ALIGN=""center""><I>" & iUnid & "</I></TD>"
                msg = msg & "<TD WIDTH=""50%""><I>" & iDescricao & "</I></TD>"
                msg = msg & "<TD WIDTH=""15%"" ALIGN=""right""><I>" & ivUnitario & "</I></TD>"
                msg = msg & "<TD WIDTH=""10%"" ALIGN=""right""><I>" & iIPI & "</I></TD>"
                msg = msg & "<TD WIDTH=""15%"" ALIGN=""right""><I>" & iTotalProd & "</I></TD>"
                msg = msg & "</TR>"
                                                    
                RstI.MoveNext
            Loop
            RstI.Close
            msg = msg & "</TABLE>"
    End If
    'MsgCorpo = MsgCorpo & vbCrLf & _
               String(80, "=") & vbCrLf

    'Resumo dos Materiais
    msg = msg & "<BR><BR><TABLE>"
    msg = msg & "<TH></TH>"
    msg = msg & "<TR><TD WIDTH=""70%"" ALIGN=""right"">(+)Frete:</TD>"
    msg = msg & "<TD WIDTH=""30%"" ALIGN=""right"">" & ConvMoeda(IIf(IsNull(Rst.Fields("Frete")), "0,00", Rst.Fields("Frete"))) & "</TD></TR>"
    
    msg = msg & "<TR><TD WIDTH=""70%"" ALIGN=""right"">(+)Seguro:</TD>"
    msg = msg & "<TD WIDTH=""30%"" ALIGN=""right"">" & ConvMoeda(IIf(IsNull(Rst.Fields("Seguro")), "0,00", Rst.Fields("Seguro"))) & "</TD></TR>"
    
    msg = msg & "<TR><TD WIDTH=""70%"" ALIGN=""right"">(+)Outros:</TD>"
    msg = msg & "<TD WIDTH=""30%"" ALIGN=""right"">" & ConvMoeda(IIf(IsNull(Rst.Fields("Outros")), "0,00", Rst.Fields("Outros"))) & "</TD></TR>"
    
    msg = msg & "<TR><TD WIDTH=""70%"" ALIGN=""right"">(-)Desconto:</TD>"
    msg = msg & "<TD WIDTH=""30%"" ALIGN=""right"">" & ConvMoeda(IIf(IsNull(Rst.Fields("Desconto")), "0,00", Rst.Fields("Desconto"))) & "</TD></TR>"
    
    msg = msg & "<TR><TD WIDTH=""70%"" ALIGN=""right"">(+)ICMS ST:</TD>"
    msg = msg & "<TD WIDTH=""30%"" ALIGN=""right"">" & ConvMoeda(IIf(IsNull(Rst.Fields("vICMSST")), "0,00", Rst.Fields("vICMSST"))) & "</TD></TR>"
    
    msg = msg & "<TR><TD WIDTH=""70%"" ALIGN=""right"">VALOR TOTAL:</TD>"
    msg = msg & "<TD WIDTH=""30%"" ALIGN=""right""><B>" & ConvMoeda(cTotalPV) & "</B></TD></TR>"
    
    
    msg = msg & "</TABLE>"

    
    'MsgCorpo = MsgCorpo & _
               vbCrLf & _
               vbCrLf & _
               vbCrLf & _
               "Frete(+): " & ConvMoeda(IIf(IsNull(Rst.fields("Frete")), "0,00", Rst.fields("Frete"))) & vbCrLf & _
               "Seguro(+): " & ConvMoeda(IIf(IsNull(Rst.fields("Seguro")), "0,00", Rst.fields("Seguro"))) & vbCrLf & _
               "Outros(+): " & ConvMoeda(IIf(IsNull(Rst.fields("Outros")), "0,00", Rst.fields("Outros"))) & vbCrLf & _
               "Desconto(-): " & ConvMoeda(IIf(IsNull(Rst.fields("Desconto")), "0,00", Rst.fields("Desconto"))) & vbCrLf & _
               "ICMSST(+): " & ConvMoeda(IIf(IsNull(Rst.fields("vICMSST")), "0,00", Rst.fields("vICMSST"))) & vbCrLf & _
               "==================================="
    
    'MsgCorpo = MsgCorpo & vbCrLf & _
             vbCrLf & vbCrLf & _
             vbCrLf & _
             "                   VALOR TOTAL: " & ConvMoeda(cTotalPV)
    
    '           "REF: " & cNull(Rst.fields("RefCliente")) & vbCrLf & _
               vbCrLf & _
               "Transportadora: " & cTransp & vbCrLf & _
               "Frete por conta: " & IIf(cNull(Rst.fields("FreteConta")) = 0, "EMITENTE", "DESTINATARIO") & vbCrLf & _
               vbCrLf & _
               "Condições de Pagamento: " & IIf(IsNull(Rst.fields("CondicoesPagamento")), "", pgDescrCondPag(Rst.fields("CondicoesPagamento"))) & IIf(IsNull(Rst.fields("FormaPagamento")), "", " (" & pgDescrTipoDoc(Rst.fields("FormaPagamento")) & ")") & vbCrLf & _
               vbCrLf & _
               "Observações: " & IIf(IsNull(Rst.fields("Obs")), "", Rst.fields("Obs")) & vbCrLf
    
    Dim cVend As String
    
    cVend = IIf(IsNull(Rst.Fields("Vendedor")), "", PgDadosRhFuncionario(Rst.Fields("Vendedor")).Nome) & _
                IIf(IsNull(Rst.Fields("Vendedor")), "", " - " & Trim(Mid(PgDadosRhFuncionario(Rst.Fields("Vendedor")).Cargo, 5, Len(PgDadosRhFuncionario(Rst.Fields("Vendedor")).Cargo)))) & vbCrLf
    
    msg = msg & "<BR><BR><BR>Transporte: <B>" & cTransp & "</B>"
    msg = msg & "<BR>Frete por conta: <B>" & IIf(cNull(Rst.Fields("FreteConta")) = 0, "EMITENTE", "DESTINATARIO") & "</B>"
    msg = msg & "<BR><BR><BR>Observações:<BR><B>" & IIf(IsNull(Rst.Fields("Obs")), "", Rst.Fields("Obs")) & "</B>"
    msg = msg & "<BR><BR><BR>Vendedor: <B>" & cVend & "</B>"
    msg = msg & Time
    msg = msg & "</BODY></HTML>"

    Rst.Close
    'grvFile "c:\prop" & ZE(IdReg, 6) & ".html", msg
    formSendMail.CarregarForm "brinfo.leo@gmail.com", "Teste de cotação online", msg, , True
    
End Sub
Private Sub Cancelar()
    If MsgBox("Deseja realmente abandorar, Pré-Venda?", vbYesNo + vbExclamation, "Cancelar") = vbNo Then
        Exit Sub
    End If
    HDMenu Me, True
    HDFormulario (False)
    LimpForm
    txtID.Enabled = True

    IdReg = 0
    iditem = 0
    IdTransp = 0
    idCliente = 0
    'idCobr = 0
    lnPv = 0
    'cCob = 0
    ZerarParcelas
End Sub
Private Sub AlterarPV()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Selecione uma Registro", vbInformation, "Aviso"
        Exit Sub
    End If
    If ChkPvTemNFe(IdReg) = True Then
        MsgBox "Pré-venda não poderá ser alterada, pois possui Nota Fiscal vinculada.", vbInformation, "Aviso"
        If MsgBox("Deseja gerar novo pedido tendo como base este?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            IdReg = 0
            HDMenu Me, False
            HDFormulario (True)
            txtID.Text = ""
            txtID.Enabled = False
            txtDesconto.Enabled = False
        End If
        Exit Sub
    End If
    HDFormulario (True)
    HDMenu Me, False
    txtID.Enabled = False
    txtDesconto.Enabled = False
End Sub

Private Sub ClonarPedido()
     If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Favor selecionar uma pré-venda.", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Deseja realmente Clonar a Pré-Venda n. " & Left(String(6, "0"), 6 - Len(Trim(IdReg))) & IdReg & "?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        IdReg = 0
        If dtpEmissao.Value <> Date Then
            If MsgBox("Deseja alterar a data de emissão da prevenda para hoje(" & Date & ")?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                dtpEmissao.Value = Date
            End If
        End If
'        If grvRegistro = True Then
'                MsgBox "Nova pré-venda n. " & Left(String(6, "0"), 6 - Len(Trim(IdReg))) & IdReg & vbCrLf & "Pré-Venda clonada com sucesso.", vbInformation, "Aviso"
'                PesquisarRegistro (IdReg)
'            Else
'                MsgBox "Erro ao criar clone da Pré-Venda.", vbInformation, "Aviso"
'        End If
        IdReg = 0
        HDMenu Me, False
        HDFormulario (True)
        txtID.Text = ""
        txtID.Enabled = False
        txtDesconto.Enabled = False
        
    End If
End Sub
Private Sub LimpForm()
    On Error GoTo TrtErro
    LimpaFormulario Me

    cboStatusPV_DropDown
    cboStatusPV.Text = cboStatusPV.List(0)

    cboVendedor.AddItem PgDadosUsuario(ID_Usuario).Nome
    cboVendedor.Text = cboVendedor.List(0)
    
    optMaterialPara(0).Value = False
    optMaterialPara(1).Value = False
    
    sstDados.Tab = 0
    dtpEmissao.Value = Date
    txtItens.Text = "0000"
    txtTotalPV.Text = ConvMoeda("0")
    txtMercadoria.Text = ConvMoeda("0")
    txtDesconto.Text = ConvMoeda("0")
    Exit Sub
TrtErro:
    MsgBox Err.Description, vbCritical, Err.Number
    Resume Next
End Sub



Private Sub cobrAcertarParcelas(sValor As String)
    Dim i           As Integer
    Dim smCobTot    As String
    smCobTot = 0
    For i = 0 To cCob
        smCobTot = Val(ChkVal(CStr(aCob(i)(1)), 0, cDecMoeda)) + Val(ChkVal(smCobTot, 0, cDecMoeda))
    Next
    If ChkVal(smCobTot, 0, cDecMoeda) > ChkVal(sValor, 0, cDecMoeda) Then
            aCob(cCob)(1) = Val(ChkVal(CStr(aCob(cCob)(1)), 0, cDecMoeda)) - (Val(ChkVal(smCobTot, 0, cDecMoeda)) - Val(ChkVal(sValor, 0, cDecMoeda)))
            aCob(cCob)(1) = ChkVal(CStr(aCob(cCob)(1)), 0, 2)
        ElseIf ChkVal(smCobTot, 0, cDecMoeda) < ChkVal(sValor, 0, cDecMoeda) Then
            aCob(0)(1) = Val(ChkVal(CStr(aCob(0)(1)), 0, cDecMoeda)) + (Val(ChkVal(sValor, 0, cDecMoeda)) - Val(ChkVal(smCobTot, 0, cDecMoeda)))
            aCob(0)(1) = ChkVal(CStr(aCob(0)(1)), 0, 2)
    End If
    
        
End Sub

Private Function cobrMontarParcelas() As Boolean
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim Valor       As String
    Dim Data        As String
    Dim sValor      As String
    Dim DtEmissao   As Date
    
    sValor = ChkVal(txtTotalPV.Text, 0, cDecMoeda)
    DtEmissao = dtpEmissao.Value
   
    If Trim(cboCondicoesPagamento.Text) = "" Then
            'MsgBox "Selecione as Condições de Pagamento!", vbInformation, App.EXEName
            cobrMontarParcelas = False
            Exit Function
        Else
            ZerarParcelas
            idCobr = Left(cboCondicoesPagamento.Text, 3)
    End If
   sSQL = "SELECT * FROM FinanceiroCondicoesPagamentoParcelas WHERE IdCondicoes = " & idCobr & " ORDER BY Parcela"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao encontrar indice financeiro da(s) parcela(s)." & vbCrLf & vbCrLf & _
                   "idCobr: " & idCobr & vbCrLf & _
                   "Tabela: FinanceiroCondicoesPagamentoParcelas", vbInformation, App.EXEName
                   cobrMontarParcelas = False
        Else
            Rst.MoveFirst
            cCob = 0
            Do Until Rst.EOF
                Data = DtEmissao + IIf(IsNull(Rst.Fields("DiasCorridos")), 0, Rst.Fields("DiasCorridos"))
                Valor = ChkVal(Val(ChkVal(IIf(IsNull(Rst.Fields("Percentual")), 0, Rst.Fields("Percentual")), 0, 3)) * Val(ChkVal(sValor, 0, 2)) / 100, 0, cDecMoeda)
                aCob(cCob) = Array(Data, Valor): cCob = cCob + 1
                Rst.MoveNext
            Loop
            cCob = cCob - 1
            cobrAcertarParcelas sValor
            cobrMontarParcelas = True
    End If
    Rst.Close
    
End Function
Private Sub txtDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarCliente
    End If
End Sub

Private Sub txtFrete_Change()
    CalcVlPV
End Sub

Private Sub txtFrete_GotFocus()
        txtFrete.Text = ChkVal(txtFrete.Text, 0, 2)

End Sub


Private Sub txtFrete_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtFrete.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtFrete_LostFocus()
    txtFrete.Text = ConvMoeda(IIf(txtFrete.Text = "", 0, txtFrete.Text))
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 114 Then
        PesquisarRegistro
    End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtID.Text) = "" Then
                MsgBox "Digite o número da Pre-Venda"
                Exit Sub
            Else
                PesquisarRegistro (txtID.Text)
        End If
    End If
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub
Private Sub txtObs_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = 13, 0, KeyAscii)
End Sub

Private Sub txtOutros_Change()
    CalcVlPV
End Sub

Private Sub txtOutros_GotFocus()
    txtOutros.Text = ChkVal(txtOutros.Text, 0, 2)
End Sub

Private Sub txtOutros_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtOutros.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtOutros_LostFocus()
    txtOutros.Text = ConvMoeda(IIf(txtOutros.Text = "", 0, txtOutros.Text))
End Sub

Private Sub txtPesoB_GotFocus()
    With txtPesoB
        .Text = ChkVal(.Text, 0, 3)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPesoB_LostFocus()
    If Trim(txtPesoB.Text) <> "" Then
        txtPesoB.Text = ChkVal(txtPesoB.Text, 0, 3)
    End If
End Sub

Private Sub txtPesoL_GotFocus()
        With txtPesoL
        .Text = ChkVal(.Text, 0, 3)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtPesoL_LostFocus()
    If Trim(txtPesoL.Text) <> "" Then
        txtPesoL.Text = ChkVal(txtPesoL.Text, 0, 3)
    End If
End Sub

Private Sub txtValidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub
Private Function ValidarPV() As Boolean
    
    On Error GoTo TrtErroValidacao
    
    Dim l As Integer
    
    ValidarPV = False
    
    If Trim(cboStatusPV.Text) = "" Then
        cboStatusPV_DropDown
        cboStatusPV.Text = cboStatusPV.List(0)
'        MsgBox "Favor informar o STATUS do documento!", vbCritical, "Aviso"
'        ValidarPV = False
'        Exit Function
    End If
    
    If Trim(cboCliente.Text) = "" Then
        MsgBox "Favor informar o NOME/RAZÃO SOCIAL do cliente!", vbCritical, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    
    If idCliente <> "0" And PgDadosCliente(idCliente).Doc <> txtDoc.Text Then
        MsgBox "O CNPJ do cliente não condiz com o CNPJ Informado!", vbCritical, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    
    If Trim(cboUF.Text) = "" Then
        MsgBox "Favor selecionar uma UF!", vbCritical, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If Trim(cboVendedor.Text) = "" Then
        MsgBox "Favor selecionar um vendedor!", vbCritical, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If Trim(cboCondicoesPagamento.Text) = "" Then
        MsgBox "Favor selecionar uma condicao de pagamento!", vbCritical, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If Trim(cboFormaPagamento.Text) = "" Then
        MsgBox "Favor selecionar um forma de pagamento!", vbCritical, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If Trim(txtValidade.Text) = "" Then
        MsgBox "Favor informar o prazo de Validade.", vbCritical, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If msfgItens.Rows = 1 Then
        MsgBox "Favor informar pelo menos um item.", vbCritical, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    If optMaterialPara(0).Value = False And optMaterialPara(1).Value = False Then
        MsgBox "Favor selecionar o uso final para o material!", vbCritical, "Aviso"
        ValidarPV = False
        Exit Function
    End If
    
    If Trim(cboTransportadora.Text) <> "" Then
        If Not IsNumeric(Mid(cboTransportadora.Text, 1, 6)) Then
            MsgBox "Favor informar transportadora cadastrada no sistema.", vbCritical, "Aviso"
            ValidarPV = False
            Exit Function
        End If
    End If
    
    For l = 0 To cCob
        If ChkVal(CStr(aCob(l)(1)), 0, cDecMoeda) = "0.00" Then
            'MsgBox "Existe Duplicata sem valor definido.", vbCritical, App.EXEName
            'ValidarPV = False
            'Exit Function
            If cobrMontarParcelas = False Then
                ValidarPV = False
                Exit Function
            End If
        End If
    Next
    
    
    ValidarPV = True
    Exit Function
TrtErroValidacao:
    MsgBox "Falha na validação dos dados.", vbCritical, "Aviso"
    ValidarPV = False
End Function
Private Sub txtPesoB_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Exit Sub
    If txtPesoB.SelLength = Len(txtPesoB.Text) Then
        txtPesoB.Text = ""
    End If
    KeyAscii = ChkVal(txtPesoB.Text, KeyAscii, 3)
End Sub

Private Sub txtPesoL_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then Exit Sub
    If txtPesoL.SelLength = Len(txtPesoL.Text) Then
        txtPesoL.Text = ""
    End If

    KeyAscii = ChkVal(txtPesoL.Text, KeyAscii, 3)
End Sub

Private Sub txtvICMSST_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtVol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Public Sub MoveRow(oldRow As Integer, UpDown As String)
    Dim i       As Integer
    Dim newRow  As Integer
    With msfgItens
    Select Case UCase(UpDown)
        Case "UP"
            If oldRow = 1 Then Exit Sub
            newRow = oldRow - 1
            oldRow = oldRow + 1
            .AddItem "", newRow
        Case "DOWN"
            If .Rows - 1 = oldRow Then Exit Sub
            newRow = oldRow + 2
            .AddItem "", newRow
    End Select
    For i = 0 To .Cols - 1
        .TextMatrix(newRow, i) = .TextMatrix(oldRow, i)
    Next
        .RemoveItem oldRow
        .Row = newRow - 1
        .RowSel = newRow - 1
    End With
End Sub
Private Sub txtSeguro_Change()
    CalcVlPV
End Sub

Private Sub txtSeguro_GotFocus()
    txtSeguro.Text = ChkVal(txtSeguro.Text, 0, 2)
End Sub

Private Sub txtSeguro_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtSeguro.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtSeguro_LostFocus()
    txtSeguro.Text = ConvMoeda(IIf(txtSeguro.Text = "", 0, txtSeguro.Text))
End Sub
