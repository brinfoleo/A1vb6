VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formRHFuncionarioCadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RH - Funcionario"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8835
   Begin VB.Frame Frame2 
      Height          =   1395
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   8655
      Begin VB.ComboBox cboCargo 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox txtCPF 
         Height          =   285
         Left            =   840
         MaxLength       =   14
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   255
         Width           =   1935
      End
      Begin VB.TextBox txtxNome 
         Height          =   285
         Left            =   840
         MaxLength       =   60
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker dtpAdmissao 
         Height          =   315
         Left            =   7140
         TabIndex        =   19
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   122814465
         CurrentDate     =   41052
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Admissão:"
         Height          =   255
         Left            =   5940
         TabIndex        =   20
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Cargo:"
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   660
         Width           =   555
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "CPF:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   330
         Width           =   555
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8835
      _ExtentX        =   15584
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
               Picture         =   "formRHFuncionarioCadastro.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioCadastro.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioCadastro.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioCadastro.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioCadastro.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioCadastro.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioCadastro.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioCadastro.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioCadastro.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioCadastro.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioCadastro.frx":54D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab sstFunc 
      Height          =   3795
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6694
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Dados Pessoais"
      TabPicture(0)   =   "formRHFuncionarioCadastro.frx":58CB
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Documentos"
      TabPicture(1)   =   "formRHFuncionarioCadastro.frx":58E7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Financeiro"
      TabPicture(2)   =   "formRHFuncionarioCadastro.frx":5903
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Outros"
      TabPicture(3)   =   "formRHFuncionarioCadastro.frx":591F
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   3195
         Left            =   -74820
         TabIndex        =   34
         Top             =   480
         Width           =   8295
         Begin VB.TextBox txtLgr 
            Height          =   285
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   540
            Width           =   6735
         End
         Begin VB.TextBox txtNro 
            Height          =   285
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox txtCpl 
            Height          =   285
            Left            =   3180
            MaxLength       =   60
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   900
            Width           =   4695
         End
         Begin VB.TextBox txtBairro 
            Height          =   285
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   1260
            Width           =   2955
         End
         Begin VB.ComboBox cboUF 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1680
            Width           =   915
         End
         Begin VB.ComboBox cboMun 
            Height          =   315
            Left            =   3060
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1620
            Width           =   2655
         End
         Begin VB.TextBox txtCEP 
            Height          =   285
            Left            =   6360
            MaxLength       =   8
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   1560
            Width           =   1515
         End
         Begin VB.TextBox txtMail 
            Height          =   285
            Left            =   1140
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   2100
            Width           =   3915
         End
         Begin VB.TextBox txtFone 
            Height          =   315
            Left            =   1140
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   2460
            Width           =   2055
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Endereço:"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Número:"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Complemento:"
            Height          =   195
            Left            =   2100
            TabIndex        =   50
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Bairro:"
            Height          =   255
            Left            =   360
            TabIndex        =   49
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Municipio:"
            Height          =   255
            Left            =   2220
            TabIndex        =   48
            Top             =   1680
            Width           =   795
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "UF:"
            Height          =   195
            Left            =   600
            TabIndex        =   47
            Top             =   1740
            Width           =   495
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "CEP:"
            Height          =   195
            Left            =   5820
            TabIndex        =   46
            Top             =   1620
            Width           =   495
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "e-mail:"
            Height          =   195
            Left            =   660
            TabIndex        =   45
            Top             =   2160
            Width           =   435
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Telefone:"
            Height          =   195
            Left            =   420
            TabIndex        =   44
            Top             =   2460
            Width           =   675
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3075
         Left            =   180
         TabIndex        =   28
         Top             =   480
         Width           =   8235
         Begin VB.TextBox txtNumFilhos 
            Height          =   285
            Left            =   1980
            TabIndex        =   76
            Text            =   "Text1"
            Top             =   1380
            Width           =   795
         End
         Begin VB.TextBox txtCargaHoraria 
            Height          =   285
            Left            =   1980
            MaxLength       =   100
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   900
            Width           =   4095
         End
         Begin VB.CheckBox chkFolhaPonto 
            Alignment       =   1  'Right Justify
            Caption         =   "Imprimir Folha de Ponto:"
            Height          =   315
            Left            =   180
            TabIndex        =   31
            Top             =   420
            Width           =   1995
         End
         Begin VB.Frame Frame5 
            Caption         =   "Nome Abreviado:"
            Height          =   675
            Left            =   240
            TabIndex        =   29
            Top             =   2220
            Width           =   3615
            Begin VB.TextBox txtAssinatura 
               Height          =   285
               Left            =   180
               MaxLength       =   60
               TabIndex        =   30
               Text            =   "Text1"
               Top             =   240
               Width           =   3255
            End
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "Filho menor 14 anos:"
            Height          =   195
            Left            =   360
            TabIndex        =   75
            Top             =   1440
            Width           =   1515
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Carga Horaria:"
            Height          =   195
            Left            =   840
            TabIndex        =   33
            Top             =   960
            Width           =   1035
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3075
         Left            =   -74880
         TabIndex        =   12
         Top             =   540
         Width           =   8295
         Begin VB.TextBox txtNumRioCard 
            Height          =   285
            Left            =   2280
            TabIndex        =   74
            Text            =   "Text1"
            Top             =   2100
            Width           =   1635
         End
         Begin VB.TextBox txtValeRefeicao 
            Height          =   315
            Left            =   2280
            TabIndex        =   73
            Text            =   "Text1"
            Top             =   1380
            Width           =   1635
         End
         Begin VB.TextBox txtValeTransp 
            Height          =   285
            Left            =   2280
            TabIndex        =   72
            Text            =   "Text1"
            Top             =   720
            Width           =   1635
         End
         Begin VB.TextBox txtAdicionais 
            Height          =   315
            Left            =   180
            TabIndex        =   64
            Text            =   "Text1"
            Top             =   1980
            Width           =   1635
         End
         Begin VB.Frame Frame7 
            Caption         =   "Dados Bancários:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   4680
            TabIndex        =   54
            Top             =   420
            Width           =   3495
            Begin VB.TextBox txtBcoCP 
               Height          =   285
               Left            =   1500
               TabIndex        =   62
               Text            =   "Text1"
               Top             =   1800
               Width           =   1815
            End
            Begin VB.TextBox txtBcoCC 
               Height          =   285
               Left            =   1500
               TabIndex        =   61
               Text            =   "Text1"
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox txtBcoAgencia 
               Height          =   285
               Left            =   1500
               TabIndex        =   60
               Text            =   "Text1"
               Top             =   840
               Width           =   1815
            End
            Begin VB.TextBox txtBcoNome 
               Height          =   285
               Left            =   1500
               TabIndex        =   59
               Text            =   "Text1"
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               Caption         =   "Conta Poupança:"
               Height          =   195
               Left            =   60
               TabIndex        =   58
               Top             =   1860
               Width           =   1275
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               Caption         =   "Conta Corrente:"
               Height          =   195
               Left            =   180
               TabIndex        =   57
               Top             =   1380
               Width           =   1155
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               Caption         =   "Agência:"
               Height          =   195
               Left            =   660
               TabIndex        =   56
               Top             =   900
               Width           =   675
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               Caption         =   "Banco:"
               Height          =   195
               Left            =   780
               TabIndex        =   55
               Top             =   420
               Width           =   555
            End
         End
         Begin VB.TextBox txtComissao 
            Height          =   285
            Left            =   180
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   1260
            Width           =   1635
         End
         Begin VB.TextBox txtSalario 
            Height          =   285
            Left            =   180
            MaxLength       =   15
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   660
            Width           =   1635
         End
         Begin VB.Label Label32 
            Caption         =   "Número Rio Card:"
            Height          =   195
            Left            =   2280
            TabIndex        =   71
            Top             =   1800
            Width           =   1515
         End
         Begin VB.Label Label31 
            Caption         =   "Vale Refeição:"
            Height          =   255
            Left            =   2280
            TabIndex        =   70
            Top             =   1140
            Width           =   1515
         End
         Begin VB.Label Label30 
            Caption         =   "Vale Transporte:"
            Height          =   195
            Left            =   2280
            TabIndex        =   69
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label27 
            Caption         =   "Adicionais:"
            Height          =   255
            Left            =   180
            TabIndex        =   63
            Top             =   1680
            Width           =   795
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Comissão (%):"
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   1020
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Salário Base:"
            Height          =   255
            Left            =   180
            TabIndex        =   15
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3195
         Left            =   -74880
         TabIndex        =   9
         Top             =   420
         Width           =   8295
         Begin VB.TextBox txtPai 
            Height          =   285
            Left            =   660
            TabIndex        =   68
            Text            =   "Text1"
            Top             =   2760
            Width           =   6255
         End
         Begin VB.TextBox txtMae 
            Height          =   285
            Left            =   660
            TabIndex        =   67
            Text            =   "Text1"
            Top             =   2340
            Width           =   6255
         End
         Begin VB.TextBox txtRGEmissao 
            Height          =   285
            Left            =   3660
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   870
            Width           =   1275
         End
         Begin VB.TextBox txtTitEleitor 
            Height          =   285
            Left            =   900
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   1380
            Width           =   1275
         End
         Begin VB.TextBox txtPIS 
            Height          =   285
            Left            =   4560
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   360
            Width           =   1515
         End
         Begin VB.TextBox txtRGExpedidor 
            Height          =   285
            Left            =   6060
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   885
            Width           =   1935
         End
         Begin VB.TextBox txtRG 
            Height          =   285
            Left            =   900
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   870
            Width           =   1935
         End
         Begin VB.TextBox txtCTPS 
            Height          =   285
            Left            =   900
            MaxLength       =   50
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Pai:"
            Height          =   255
            Left            =   180
            TabIndex        =   66
            Top             =   2760
            Width           =   435
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Mãe:"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Tit. Eleitor:"
            Height          =   255
            Left            =   60
            TabIndex        =   26
            Top             =   1440
            Width           =   795
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "PIS nº:"
            Height          =   255
            Left            =   3720
            TabIndex        =   24
            Top             =   360
            Width           =   675
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Expedidor:"
            Height          =   255
            Left            =   5040
            TabIndex        =   22
            Top             =   900
            Width           =   915
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Emissão:"
            Height          =   195
            Left            =   2940
            TabIndex        =   21
            Top             =   930
            Width           =   615
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "RG:"
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "CTPS:"
            Height          =   255
            Left            =   300
            TabIndex        =   11
            Top             =   420
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "formRHFuncionarioCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IdReg    As Integer
Dim strTabela           As String


Private Sub PesquisarRegistro()
    ''Dim idreg  As String
    IdReg = formBuscar.IniciarBusca(strTabela, "xNome, CPF, RG,Lgr,Nro,Cpl, Bairro, Mun,UF")
    ''IdReg = IIf(idreg = "", 0, idreg)
    
    If IdReg = 0 Then
            LimpaFormulario Me
        Else
            MostrarDados
    End If
End Sub



Private Sub cboCargo_DropDown()
    Dim rst     As Recordset
    Dim sSQL    As String
    
    cboCargo.Clear
    sSQL = "SELECT * FROM RHFuncionarioCargo"
    Set rst = RegistroBuscar(sSQL)
    If rst.BOF And rst.EOF Then
            rst.Close
        Else
            rst.MoveFirst
            Do Until rst.EOF
                cboCargo.AddItem Left("00", 2 - Len(Trim(rst.Fields("id")))) & rst.Fields("id") & " - " & _
                                 rst.Fields("Descricao")
                rst.MoveNext
            Loop
    End If
    'With cboCargo
    '    .AddItem "01 - Diretor Geral"
    '    .AddItem "02 - Administrador"
    '    .AddItem "03 - Auxiliar Administrativo"
    '    .AddItem "04 - Gerente"
    '    .AddItem "05 - Vendedor"
    '    .AddItem "06 - Estoquista"
    '    .AddItem "07 - Auxilioar de almoxarifado"
    'End With
End Sub

Private Sub cboMun_DropDown()
    Dim rst     As Recordset
    Dim sSQL    As String
    If Trim(cboUF.Text) = "" Then
        MsgBox "Selecione uma Unidade Federal (UF)."
        Exit Sub
    End If
    sSQL = "SELECT * FROM TributacaoMunicipio WHERE UF = '" & Trim(UCase(cboUF.Text)) & "' ORDER BY Descricao"
    cboMun.Clear
    Set rst = RegistroBuscar(sSQL)
    If rst.BOF And rst.EOF Then
        Else
            rst.MoveFirst
            Do Until rst.EOF
                cboMun.AddItem UCase(rst.Fields("Descricao"))
                rst.MoveNext
            Loop
    End If
End Sub




Private Sub cboUF_DropDown()
    Dim rst As Recordset
    cboUF.Clear
    Set rst = RegistroBuscar("SELECT * FROM TributacaoUF ORDER BY SIGLA")
    If rst.BOF And rst.EOF Then
        Else
            rst.MoveFirst
            Do Until rst.EOF
                cboUF.AddItem rst.Fields("Sigla")
                rst.MoveNext
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
    HDForm Me, False
    HDMenu Me, True
    txtCPF.Enabled = True
    sstFunc.Tab = 0
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
        MsgBox "Selecione um registro!", vbInformation, "Aviso"
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
            MsgBox "Selecione uma Transportadora"
            Exit Sub
        Else
            If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                        vbCrLf & _
                        "CPF: " & txtCPF.Text & vbCrLf & _
                        "Nome: " & txtxNome.Text, vbYesNo + vbQuestion) = vbYes Then
                               
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
                'txtCPF.Enabled = True
            End If
            
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            txtCPF.Enabled = True
        
        Case "Manutenção da Tabela"
            formManutencaoTabelas.IniciarManutencao Me
    End Select
End Sub

Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim i           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer 'Contador de Registros
    If ValidarRegistros = False Then
        grvRegistro = False
        Exit Function
    End If
    cReg = 0
    For i = 0 To Me.Controls.Count - 1
        Set Controle = Me.Controls(i)
        
        If TypeOf Controle Is TextBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Text, "S")
            cReg = cReg + 1
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



End Function

Private Sub txtComissao_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtComissao.Text, KeyAscii, 3)
End Sub

Private Sub txtCPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
    
End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        BuscarDados (txtCPF.Text)
    End If
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub
Private Sub BuscarDados(strCPF As String)
    Dim rst     As ADODB.Recordset
    Dim strSQL  As String
    
    'sstTransportadora.Tab = 0
    
    strSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND CPF = '" & strCPF & "'"

    Set rst = RegistroBuscar(strSQL)
    If rst.BOF And rst.EOF Then
            MsgBox "Nenhum Registro encontrado"
            rst.Close
            Exit Sub
        Else
            rst.MoveFirst
            IdReg = rst.Fields("Id")
            rst.Close
            MostrarDados
    End If
End Sub
Private Sub MostrarDados()
    Dim sSQL As String
    sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg

    ExibirDados Me, sSQL

End Sub


Private Sub txtNumFilhos_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
End Sub


Private Sub txtSalario_GotFocus()
    txtSalario.Text = ChkVal(txtSalario.Text, 0, cDecMoeda)
End Sub

Private Sub txtSalario_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtSalario.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtSalario_LostFocus()
    txtSalario.Text = ConvMoeda(txtSalario.Text)
End Sub
Private Function ValidarRegistros() As Boolean
    If Trim(txtxNome.Text) = "" Then
        ValidarRegistros = False
        MsgBox "Campo NOME invalido!", vbCritical, App.EXEName
        Exit Function
    End If
    ValidarRegistros = True
End Function
