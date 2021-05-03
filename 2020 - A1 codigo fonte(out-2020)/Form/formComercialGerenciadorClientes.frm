VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formComercialGerenciadorClientes 
   Caption         =   "Gerenciador de Clientes"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10740
   ScaleWidth      =   14895
   Begin VB.Frame frmConsCliente 
      Height          =   8415
      Left            =   420
      TabIndex        =   3
      Top             =   1380
      Width           =   13455
      Begin VB.Frame frmTitulosVencidos 
         Caption         =   "Titulos Vencidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         Left            =   5640
         TabIndex        =   17
         Top             =   3600
         Width           =   7215
         Begin MSFlexGridLib.MSFlexGrid msfgTitulos 
            Height          =   2355
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   6915
            _ExtentX        =   12197
            _ExtentY        =   4154
            _Version        =   393216
            Cols            =   6
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   "^ID |>Titulo                       |>Valor                   |^Vencimento    |^Atraso  |>Valor Atualizado      "
         End
      End
      Begin VB.Frame frmClientes 
         Caption         =   "Clientes"
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
         TabIndex        =   13
         Top             =   180
         Width           =   13155
         Begin VB.TextBox txtNome 
            Height          =   285
            Left            =   660
            TabIndex        =   14
            Text            =   "Text1"
            ToolTipText     =   "Digite o NOME do cliente e pressione <Entre>..."
            Top             =   240
            Width           =   6675
         End
         Begin MSFlexGridLib.MSFlexGrid msfgClientes 
            Height          =   2655
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   12915
            _ExtentX        =   22781
            _ExtentY        =   4683
            _Version        =   393216
            Cols            =   8
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"formComercialGerenciadorClientes.frx":0000
         End
         Begin VB.Label Label2 
            Caption         =   "Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.Frame frmPV 
         Caption         =   "Pré-Vendas"
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
         Left            =   120
         TabIndex        =   11
         Top             =   3600
         Width           =   4995
         Begin MSFlexGridLib.MSFlexGrid msfgPV 
            Height          =   2355
            Left            =   120
            TabIndex        =   12
            Top             =   300
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   4154
            _Version        =   393216
            Cols            =   5
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   "^PV           |^Emissão PV     |>Valor                         |^Validade  |<NFe                         "
         End
      End
      Begin VB.Frame frmHC 
         Caption         =   "Ultimos Contatos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   180
         TabIndex        =   4
         Top             =   6360
         Width           =   13035
         Begin VB.Frame frmDescrCont 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   180
            TabIndex        =   5
            Top             =   1200
            Width           =   10755
            Begin VB.TextBox txtDescricao 
               Height          =   285
               Left            =   900
               MaxLength       =   5000
               TabIndex        =   8
               Top             =   180
               Width           =   7575
            End
            Begin VB.CommandButton btoIncluir 
               Caption         =   "&Incluir"
               Height          =   315
               Left            =   8580
               TabIndex        =   7
               Top             =   180
               Width           =   1035
            End
            Begin VB.CommandButton btoExcluir 
               Caption         =   "&Excluir"
               Height          =   315
               Left            =   9660
               TabIndex        =   6
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label3 
               Caption         =   "Descricão:"
               Height          =   195
               Left            =   60
               TabIndex        =   9
               Top             =   240
               Width           =   855
            End
         End
         Begin MSFlexGridLib.MSFlexGrid msfgHCont 
            Height          =   1035
            Left            =   120
            TabIndex        =   10
            Top             =   180
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   1826
            _Version        =   393216
            Cols            =   3
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"formComercialGerenciadorClientes.frx":00C4
         End
      End
   End
   Begin VB.Frame frmConsProduto 
      Caption         =   "Produto Faturado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   4140
      TabIndex        =   29
      Top             =   2940
      Width           =   14115
      Begin VB.Frame frmPesqProd 
         Height          =   975
         Left            =   240
         TabIndex        =   32
         Top             =   6780
         Width           =   12135
         Begin VB.TextBox txtPesqProd 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   480
            Width           =   9255
         End
         Begin VB.Label Label5 
            Caption         =   "Digite o material que deseja consultar:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   3195
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid 
         Height          =   6075
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   12675
         _ExtentX        =   22357
         _ExtentY        =   10716
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   "Chave Acesso"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   "Emissão"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   ""
            Caption         =   "Num. Nota Fiscal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   ""
            Caption         =   "Cliente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   ""
            Caption         =   "Produto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frmConsNF 
      Height          =   7515
      Left            =   0
      TabIndex        =   19
      Top             =   3180
      Width           =   14535
      Begin VB.Frame frmDescricao 
         Caption         =   "Descrição dos Produtos"
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
         Left            =   180
         TabIndex        =   24
         Top             =   3720
         Width           =   13515
         Begin MSFlexGridLib.MSFlexGrid msfgDescricao 
            Height          =   2355
            Left            =   180
            TabIndex        =   25
            Top             =   240
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   4154
            _Version        =   393216
            Cols            =   15
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"formComercialGerenciadorClientes.frx":018C
         End
      End
      Begin VB.Frame frmNF 
         Caption         =   "Notas Fiscais"
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
         TabIndex        =   20
         Top             =   240
         Width           =   13155
         Begin VB.TextBox txtNFNome 
            Height          =   285
            Left            =   660
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   240
            Width           =   6675
         End
         Begin MSFlexGridLib.MSFlexGrid msfgNF 
            Height          =   2655
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   12915
            _ExtentX        =   22781
            _ExtentY        =   4683
            _Version        =   393216
            Cols            =   6
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"formComercialGerenciadorClientes.frx":027B
         End
         Begin VB.Label Label4 
            Caption         =   "Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   300
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtrar por:"
      Height          =   555
      Left            =   8640
      TabIndex        =   26
      Top             =   420
      Width           =   5715
      Begin VB.OptionButton optConsulta 
         Caption         =   "Produto Faturado"
         Height          =   195
         Index           =   2
         Left            =   3540
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optConsulta 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optConsulta 
         Caption         =   "Nota Fiscal"
         Height          =   195
         Index           =   1
         Left            =   1740
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.ComboBox cboFuncionario 
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   540
      Width           =   7275
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir PV"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Visualizar PV"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir DANFe"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Titulos Vencidos"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Picture         =   "formComercialGerenciadorClientes.frx":0386
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":07D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":0AF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":1384
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":25D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":2EB0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":3742
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":3FD4
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":5226
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":5540
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":585A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":5C51
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":692B
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formComercialGerenciadorClientes.frx":6EC5
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Funcionario:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1035
   End
End
Attribute VB_Name = "formComercialGerenciadorClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTabela   As String
Dim idFunc      As Integer
'Dim lHC         As Integer
Dim idCliente   As Integer
Dim idContato   As Integer
Dim idPV        As Integer
Dim chvNFe      As String
Private Sub LstContatos()
    Dim Rst     As Recordset
    Dim sSQL    As String
    msfgHCont.Rows = 1
    sSQL = "SELECT * FROM ComercialGerenciadorClientes WHERE  ID_Empresa = " & ID_Empresa & " AND IdCliente = " & idCliente
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            With msfgHCont
                Rst.MoveFirst
                Do Until Rst.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("id")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("dthr")
                    .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rst.Fields("Descricao")), "", Rst.Fields("Descricao"))
                    Rst.MoveNext
                Loop
            End With
    End If
    Rst.Close
End Sub

Private Function MontStringConsultaProduto(CampoBusca As String, sTexto As String) As String
    Dim sBtmp       As String
    Dim sBusca      As String
    Dim sParte      As String
    'Dim CampoBusca  As String
    'CampoBusca = "Descricao"
    sBtmp = ""
    sBusca = Replace(Trim(sTexto), " ", "|") & "|"
    Do Until InStr(sBusca, "|") = 0
        sParte = Trim(Mid(sBusca, 1, InStr(sBusca, "|") - 1))
        sParte = Replace(sParte, "'", "''")
                            
        sBtmp = IIf(Trim(sBtmp) = "", "", sBtmp & " AND ") & CampoBusca & " LIKE '%" & Trim(sParte) & "%'"
        sBusca = Mid(sBusca, InStr(sBusca, "|") + 1, Len(sBusca))
    Loop
    MontStringConsultaProduto = sBtmp '& " ORDER BY " & CampoBusca
End Function

Private Function PgDtUltimaCompra(idCliente As Integer) As String
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND Dest_idDest = " & idCliente & " ORDER BY ide_dEmi"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            PgDtUltimaCompra = "00/00/0000"
        Else
            Rst.MoveLast
            PgDtUltimaCompra = Rst.Fields("ide_dEmi")
    End If
    Rst.Close
End Function

Private Sub ResizeClientes()
    On Error Resume Next
    DoEvents
    'Frame CONSULTA CLIENTES
    frmConsCliente.Top = 1020
    frmConsCliente.Left = 120
    frmConsCliente.Height = Me.Height - (frmConsCliente.Top + 600)
    frmConsCliente.Width = Me.Width - 350
    
    'Frame Clientes
    frmClientes.Height = frmConsCliente.Height / 3.2 '(frmConsCliente.Height - frmClientes.Top) / 3.28
    frmClientes.Width = frmConsCliente.Width - 250
    msfgClientes.Width = frmClientes.Width - 250
    msfgClientes.Height = frmClientes.Height - (msfgClientes.Top + 150)
    
   
    'Frame PV
    frmPV.Top = frmClientes.Height + frmClientes.Top + 100
    frmPV.Left = frmClientes.Left
    frmPV.Width = frmClientes.Width / 2
    frmPV.Height = frmClientes.Height
    msfgPV.Top = 300
    msfgPV.Width = frmPV.Width - 250
    msfgPV.Height = frmPV.Height - (msfgPV.Top + 150)
    
    'Frame Titulos
    frmTitulosVencidos.Top = frmClientes.Height + frmClientes.Top + 100
    frmTitulosVencidos.Left = frmPV.Width + frmPV.Left + 100
    frmTitulosVencidos.Width = (frmClientes.Width / 2) - 100
    frmTitulosVencidos.Height = frmClientes.Height
    msfgTitulos.Top = 300
    msfgTitulos.Height = frmTitulosVencidos.Height - (msfgTitulos.Top + 150)
    msfgTitulos.Width = frmTitulosVencidos.Width - 250
    
    'Frame Historico de Contato
    frmHC.Top = frmPV.Height + frmPV.Top + 100
    frmHC.Height = frmClientes.Height
    frmHC.Left = frmClientes.Left
    frmHC.Width = frmClientes.Width
    msfgHCont.Top = 300
    msfgHCont.Width = frmHC.Width - 250
    msfgHCont.Height = frmHC.Height - (msfgHCont.Top + frmDescrCont.Height + 150)
    
    frmDescrCont.Top = msfgHCont.Top + msfgHCont.Height
End Sub
Private Sub ResizeProdutos()
    On Error Resume Next
    DoEvents
    'Frame CONSULTA POR PRODUTO
    frmConsProduto.Top = 1020
    frmConsProduto.Left = 120
    frmConsProduto.Height = Me.Height - (frmConsProduto.Top + 600)
    frmConsProduto.Width = Me.Width - 350
    
    'DataGrid
    DataGrid.Left = 100
    DataGrid.Width = frmConsProduto.Width - 300
    DataGrid.Height = frmConsProduto.Height - (frmPesqProd.Height + 800)
    'frmClientes.Width = frmConsProduto.Width - 250
    'msfgClientes.Width = frmClientes.Width - 250
    'msfgClientes.Height = frmClientes.Height - (msfgClientes.Top + 150)
    
   
    'frmPesqProd
    frmPesqProd.Left = DataGrid.Left
    frmPesqProd.Top = DataGrid.Height + (frmPesqProd.Height / 2)
    frmPesqProd.Width = DataGrid.Width
    'frmPV.Left = frmClientes.Left
    'frmPV.Width = frmClientes.Width / 2
    'frmPV.Height = frmClientes.Height
    'msfgPV.Top = 300
    'msfgPV.Width = frmPV.Width - 250
    'msfgPV.Height = frmPV.Height - (msfgPV.Top + 150)
    
    'Frame Titulos
    'frmTitulosVencidos.Top = frmClientes.Height + frmClientes.Top + 100
    'frmTitulosVencidos.Left = frmPV.Width + frmPV.Left + 100
    'frmTitulosVencidos.Width = (frmClientes.Width / 2) - 100
    'frmTitulosVencidos.Height = frmClientes.Height
    'msfgTitulos.Top = 300
    'msfgTitulos.Height = frmTitulosVencidos.Height - (msfgTitulos.Top + 150)
    'msfgTitulos.Width = frmTitulosVencidos.Width - 250
    
    'Frame Historico de Contato
    'frmHC.Top = frmPV.Height + frmPV.Top + 100
    'frmHC.Height = frmClientes.Height
    'frmHC.Left = frmClientes.Left
    'frmHC.Width = frmClientes.Width
    'msfgHCont.Top = 300
    'msfgHCont.Width = frmHC.Width - 250
    'msfgHCont.Height = frmHC.Height - (msfgHCont.Top + frmDescrCont.Height + 150)
    
    'frmDescrCont.Top = msfgHCont.Top + msfgHCont.Height
End Sub

Private Sub ResizeNF()
    On Error Resume Next
    DoEvents
    'Frame CONSULTA CLIENTES
    frmConsNF.Top = 1020
    frmConsNF.Left = 120
    frmConsNF.Height = Me.Height - (frmConsNF.Top + 600)
    frmConsNF.Width = Me.Width - 350
    
    'Frame Notas Fiscais
    frmNF.Height = frmConsNF.Height / 3.2 '(frmConsCliente.Height - frmClientes.Top) / 3.28
    frmNF.Width = frmConsNF.Width - 250
    msfgNF.Width = frmNF.Width - 250
    msfgNF.Height = frmNF.Height - (msfgNF.Top + 150)
    
   
    'Frame Descricao
    frmDescricao.Top = frmNF.Height + frmNF.Top + 100
    frmDescricao.Left = frmNF.Left
    frmDescricao.Width = frmNF.Width
    frmDescricao.Height = frmNF.Height
    msfgDescricao.Top = 300
    msfgDescricao.Width = frmDescricao.Width - 250
    msfgDescricao.Height = frmDescricao.Height - (msfgDescricao.Top + 150)
    
    'Frame Titulos
    'frmTitulosVencidos.Top = frmClientes.Height + frmClientes.Top + 100
    'frmTitulosVencidos.Left = frmPV.Width + frmPV.Left + 100
    'frmTitulosVencidos.Width = (frmClientes.Width / 2) - 100
    'frmTitulosVencidos.Height = frmClientes.Height
    'msfgTitulos.Top = 300
    'msfgTitulos.Height = frmTitulosVencidos.Height - (msfgTitulos.Top + 150)
    'msfgTitulos.Width = frmTitulosVencidos.Width - 250
    
    'Frame Historico de Contato
    'frmHC.Top = frmPV.Height + frmPV.Top + 100
    'frmHC.Height = frmClientes.Height
    'frmHC.Left = frmClientes.Left
    'frmHC.Width = frmClientes.Width
    'msfgHCont.Top = 300
    'msfgHCont.Width = frmHC.Width - 250
    'msfgHCont.Height = frmHC.Height - (msfgHCont.Top + frmDescrCont.Height + 150)
    
    'frmDescrCont.Top = msfgHCont.Top + msfgHCont.Height
End Sub



Private Sub btoExcluir_Click()
    If idContato = 0 Then Exit Sub
    If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
              vbCrLf & _
              "Descrição.: " & txtDescricao.Text, vbYesNo + vbQuestion) = vbYes Then
        If RegistroExcluir(strTabela, "Id = " & idContato) = True Then
            txtDescricao.Text = ""
            idContato = 0
            LstContatos
        End If
    End If
    
End Sub

Private Sub btoIncluir_Click()
     If grvRegistro = True Then
        txtDescricao.Text = ""
        idContato = 0
        LstContatos
    End If
End Sub

Private Function grvRegistro() As Boolean
    Dim vReg(1000)   As Variant
    Dim cReg         As Integer 'Contador de Registros
    'Dim idContato        As Integer  'Pega o Id do registro gravado
    Dim i            As Integer
    cReg = 0
    vReg(cReg) = Array("idCliente", idCliente, "N"): cReg = cReg + 1
    vReg(cReg) = Array("DtHrReg", CStr(Now), "S"): cReg = cReg + 1
    vReg(cReg) = Array("descricao", Trim(txtDescricao.Text), "S") ': cReg = cReg + 1
    
    If idContato = 0 Then
            If RegistroIncluir(strTabela, vReg, cReg) = 0 Then
                    MsgBox "Erro ao Incluir"
                    grvRegistro = False
                Else
                    grvRegistro = True
            End If
        Else
            If RegistroAlterar(strTabela, vReg, cReg, "Id = " & idContato) = False Then
                    MsgBox "Erro ao Alterar."
                    grvRegistro = False
                Else
                    grvRegistro = True
                
            End If
    End If
 End Function
Private Sub cboFuncionario_Click()
    If cboFuncionario.Text = "" Then
        LimpForm
        Exit Sub
    End If
    idFunc = Left(cboFuncionario.Text, 3)
    If optConsulta.Item(0).Value = True Then
            LstClientes
        Else
            LstNotasFiscais
    End If
End Sub

Private Sub cboFuncionario_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    cboFuncionario.Clear
    sSQL = "SELECT * FROM RHFuncionarioCadastro WHERE ID_Empresa = " & ID_Empresa & " ORDER BY xNome"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboFuncionario.AddItem Left("000", 3 - Len(Trim(Rst.Fields("id")))) & Trim(Rst.Fields("id")) & " - " & Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub LstClientes()
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim sSQLa   As String
    Dim sBusca  As String
    Dim sBtmp   As String
    If idFunc = 0 Then Exit Sub
    msfgClientes.Rows = 1
    LimpForm
    sBusca = Replace(Trim(txtNome.Text), " ", "|") & "|"
    If Trim(txtNome.Text) = "" Then
            sSQL = "SELECT * FROM Clientes WHERE ID_Empresa = " & ID_Empresa & " AND Vendedor = " & idFunc & " LIMIT 50"
        Else
            Do Until InStr(sBusca, "|") = 0
                sSQLa = IIf(Trim(sBtmp) = "", "", sBtmp & " AND ") & "xNome" & " LIKE '%" & Trim(Mid(sBusca, 1, InStr(sBusca, "|") - 1)) & "%'"
                sBusca = Mid(sBusca, InStr(sBusca, "|") + 1, Len(sBusca))
            Loop
            sSQL = "SELECT * FROM Clientes WHERE ID_Empresa = " & ID_Empresa & " AND Vendedor = " & idFunc & IIf(Trim(sSQLa) = "", "", " AND " & sSQLa) & " LIMIT 50"
    End If
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            With msfgClientes
                .Rows = 1
                Do Until Rst.EOF
                DoEvents
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("id")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("xNome")
                    .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rst.Fields("Doc")), "", Rst.Fields("Doc"))
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("fone")), "", Rst.Fields("fone"))
                    .TextMatrix(.Rows - 1, 5) = PgDtUltimaCompra(Rst.Fields("id"))
                    
                    .TextMatrix(.Rows - 1, 6) = PgNumTitulosAtraso(Rst.Fields("id"))
                    
                    .TextMatrix(.Rows - 1, 7) = PgNumTitulosAvencer(Rst.Fields("id"))
                    
                    Rst.MoveNext
                Loop
            End With
    End If
    Rst.Close
    

End Sub

Private Sub LstNotasFiscais()
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim sSQLa   As String
    Dim sBusca  As String
    Dim sBtmp   As String
    If idFunc = 0 Then Exit Sub
    msfgNF.Rows = 1
    msfgDescricao.Rows = 1
    LimpForm
    sBusca = Replace(Trim(txtNFNome.Text), " ", "|") & "|"
    If Trim(txtNFNome.Text) = "" Then
            sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND ger_Vendedor = " & idFunc
        Else
            Do Until InStr(sBusca, "|") = 0
                sSQLa = IIf(Trim(sBtmp) = "", "", sBtmp & " AND ") & "dest_xNome" & " LIKE '%" & Trim(Mid(sBusca, 1, InStr(sBusca, "|") - 1)) & "%'"
                sBusca = Mid(sBusca, InStr(sBusca, "|") + 1, Len(sBusca))
            Loop
            sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND ger_Vendedor = " & idFunc & IIf(Trim(sSQLa) = "", "", " AND " & sSQLa)
    End If
    sSQL = sSQL & " ORDER BY ide_nNF DESC LIMIT 200"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            With msfgNF
                .Rows = 1
                Do Until Rst.EOF
                    DoEvents
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("id")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("dest_xNome")
                    .TextMatrix(.Rows - 1, 2) = Rst.Fields("ide_dEmi")
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("ide_nNF")), "", Rst.Fields("ide_nNF"))
                    .TextMatrix(.Rows - 1, 4) = IIf(IsNull(Rst.Fields("IdNFe")), "", Rst.Fields("IdNFe"))
                    .TextMatrix(.Rows - 1, 5) = IIf(IsNull(Rst.Fields("ger_IdPV")), "0", Rst.Fields("ger_IdPV"))
                    '.TextMatrix(.Rows - 1, 7) = PgNumTitulosAvencer(Rst.Fields("id"))
                    
                    Rst.MoveNext
                Loop
            End With
    End If
    Rst.Close
    

End Sub



Private Sub LstPV(idCliente As Integer)
    Dim Rst     As Recordset
    Dim sSQL    As String
    'If idFunc = 0 Then Exit Sub
    msfgPV.Rows = 1
    sSQL = "SELECT * FROM FaturamentoPV WHERE  ID_Empresa = " & ID_Empresa & " AND idCliente = " & idCliente & " LIMIT 300" ' ORDER Emissao DESC"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            With msfgPV
                .Rows = 1
                Do Until Rst.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("id")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("Emissao")
                    .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rst.Fields("vlTotalPV")), "", ConvMoeda(Rst.Fields("vlTotalPV")))
                    .TextMatrix(.Rows - 1, 3) = Rst.Fields("Validade")
                    .TextMatrix(.Rows - 1, 4) = pgNumNFe(Rst.Fields("id"))
                    Rst.MoveNext
                Loop
            End With
    End If
    Rst.Close
End Sub
Private Function pgNumNFe(intPV As Integer) As String
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE  ID_Empresa = " & ID_Empresa & " AND ger_idPV = " & intPV & " ORDER BY Ide_nNF"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            pgNumNFe = ""
        Else
            Rst.MoveLast
            pgNumNFe = Rst.Fields("ide_nNF") & " [" & Rst.Fields("idNFe") & "]"
            
    End If
    Rst.Close
End Function

Private Sub DataGrid_Click()
    On Error GoTo TrtErroFatagrid
    Dim sSQL As String
    Dim Rst As Recordset
    chvNFe = DataGrid.Columns(0).Text
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa =" & ID_Empresa & " AND " & "IdNFe='" & chvNFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            chvNFe = ""
            idPV = 0
        Else
            idPV = Rst.Fields("ger_idPV")
    End If
    Rst.Close
    Exit Sub
TrtErroFatagrid:
End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9870
    Me.Width = 15015
    optConsulta_Click (0) 'LstClientes
    HDMenu Me, True
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    LimpaFormulario Me
    LimpForm
    cboFuncionario.Clear
    cboFuncionario.AddItem Left("000", 3 - Len(Trim(PgDadosUsuario(ID_Usuario).idFunc))) & PgDadosUsuario(ID_Usuario).idFunc & " - " & PgDadosRhFuncionario(PgDadosUsuario(ID_Usuario).idFunc).Nome
    '
    'cboFuncionario.AddItem PgDadosUsuario(ID_Usuario).Nome
    '
    cboFuncionario.Text = cboFuncionario.List(0)
    idFunc = Left(cboFuncionario.Text, 3)
    'Caso seja super usuario desabilita o cbo de funcionario impedindo de ver os dados dos outros
    If PgDadosUsuario(ID_Usuario).SuperUsuario <> 1 Then
        cboFuncionario.Enabled = IIf(PgDadosConfig.VisualizarOutrosFunc = 0, False, True)
    End If
    
End Sub

Private Sub Form_Resize()
    ResizeClientes
    ResizeNF
    ResizeProdutos
End Sub

Private Sub msfgClientes_Click()
    
    idPV = 0
    chvNFe = ""
    If msfgClientes.TextMatrix(msfgClientes.Row, 0) = "ID" Or msfgClientes.TextMatrix(msfgClientes.Row, 0) = "" Then
        LimpForm
        Exit Sub
    End If
    idCliente = msfgClientes.TextMatrix(msfgClientes.Row, 0)
    LstPV (idCliente)
    LstTitulosVencidos (idCliente)
    LstContatos
End Sub
Private Sub LimpForm()
    msfgPV.Rows = 1
    msfgTitulos.Rows = 1
    msfgHCont.Rows = 1
    
End Sub

Private Sub LstTodosTitulosVencidos()
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim dtVenc      As Date
    Dim vDupl       As String
    Dim vTotal      As String
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    
    dtVenc = Date - 1
    sSQL = "SELECT * " & _
         "FROM FinanceiroContasPRCadastro " & _
         "WHERE ID_Empresa = " & ID_Empresa & " AND ContaPR = 'R' AND Vencimento <= '" & Format(dtVenc, "YYYY-MM-DD") & "' AND DataQuitacao IS NULL " & _
         "ORDER BY Nome, Vencimento"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Rst.Close
            MsgBox "Nenhuma Duplicata em atraso!", vbInformation, App.EXEName
            Exit Sub
        Else
            Rst.MoveFirst
    End If


    
    MontarTabelaTemporaria

    vTotal = "0"
    Do Until Rst.EOF
        cReg = 0
        
        vDupl = AtualizaCobranca(Rst.Fields("id"), CStr(Date)).vTotal
        
        
        vReg(cReg) = Array("Nome", Rst.Fields("Nome"), "S"): cReg = cReg + 1
        vReg(cReg) = Array("Duplicata", Rst.Fields("NumDuplicata"), "S"): cReg = cReg + 1
        vReg(cReg) = Array("Vencimento", Rst.Fields("Vencimento"), "S"): cReg = cReg + 1
        vReg(cReg) = Array("CalcPara", Date, "S"): cReg = cReg + 1
        vReg(cReg) = Array("DiasVencidos", AtualizaCobranca(Rst.Fields("id"), CStr(Date)).DiasVencidos, "S"): cReg = cReg + 1
        vReg(cReg) = Array("ValorAtualizado", ConvMoeda(vDupl), "S"): cReg = cReg + 1
        cReg = cReg - 1
        Debug.Print PgDadosCliente(Rst.Fields("IdSacado")).Vendedor
        If PgDadosCliente(Rst.Fields("IdSacado")).Vendedor = idFunc Then
            RegistroIncluir "tmp_Titulos", vReg, cReg
            vTotal = Val(ChkVal(vTotal, 0, cDecMoeda)) + Val(ChkVal(vDupl, 0, cDecMoeda))
        End If
        Rst.MoveNext
    Loop



    sSQL = "SELECT * FROM tmp_titulos"

    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
        Set rptListaTitulos.DataSource = Rst.DataSource
        rptListaTitulos.Sections("Section2").Controls.Item("lblFunc").Caption = "Listagem de Titulos Vencidos"
        rptListaTitulos.Sections("Section5").Controls.Item("lblValorDuplicataC").Caption = ConvMoeda("0")
        rptListaTitulos.Sections("Section5").Controls.Item("lblValorDuplicataD").Caption = ConvMoeda(vTotal)
        
        rptListaTitulos.Show 1
    End If
    Rst.Close


End Sub
Private Sub MontarTabelaTemporaria()

    Dim sCampos     As String
    Dim i           As Integer
    
    BD.Execute "DROP TABLE IF EXISTS tmp_titulos"
    'sCampos = ""
    'For i = 1 To msfgContas.Cols - 1
    '    sCampos = sCampos & RS(msfgContas.TextMatrix(0, i)) & " VARCHAR(100) default Null,"
    'Next
    sCampos = "CREATE TABLE IF NOT EXISTS tmp_titulos " & _
              "(Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "UsuID VARCHAR(10) default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "Nome VARCHAR(120) default Null," & _
               "Vencimento VARCHAR(20) default Null," & _
               "CalcPara VARCHAR(20) default Null," & _
               "DiasVencidos VARCHAR(20) default Null," & _
               "Duplicata VARCHAR(20) default Null," & _
               "ValorAtualizado VARCHAR(20) default Null," & _
               "DataQuitação VARCHAR(20) default Null," & _
               " PRIMARY KEY (Id))"
    BD.Execute sCampos
End Sub

Private Sub LstTitulosVencidos(idCliente As Integer)
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim vDupl       As String
    Dim DiasVenc    As Integer
    Dim pMulta      As String
    Dim pJuros      As String
    Dim vMulta      As String
    Dim vJuros      As String
    'If idFunc = 0 Then Exit Sub
    msfgTitulos.Rows = 1
    sSQL = "SELECT * FROM FinanceiroContasPRCadastro WHERE ID_Empresa = " & ID_Empresa & " AND idSacado = " & idCliente & " AND Tabela = 'Clientes' AND Vencimento <= '" & Format(Date, "YYYY-MM-DD") & "' AND DataQuitacao IS NULL"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            With msfgTitulos
                Do Until Rst.EOF
                    .Rows = .Rows + 1
                    
                    DiasVenc = Date - Rst.Fields("Vencimento")
                    
                    vDupl = ChkVal(Rst.Fields("vlDuplicata"), 0, cDecMoeda)
                    
                    pMulta = IIf(IsNull(Rst.Fields("Multa")), "0", Rst.Fields("Multa"))
                    
                    pJuros = IIf(IsNull(Rst.Fields("Juros")), "0", Rst.Fields("Juros"))
                    
                    vMulta = ChkVal(Val(pMulta) * Val(vDupl) / 100, 0, cDecMoeda)
                    
                    vJuros = ChkVal((Val(pJuros) * Val(DiasVenc)) * (Val(vDupl) + Val(vMulta)) / 100, 0, cDecMoeda)
                    
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("id")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("NumDuplicata")
                    .TextMatrix(.Rows - 1, 2) = ConvMoeda(vDupl)
                    .TextMatrix(.Rows - 1, 3) = Rst.Fields("Vencimento")
                    
                    .TextMatrix(.Rows - 1, 4) = DiasVenc
                    .TextMatrix(.Rows - 1, 5) = ConvMoeda(Val(vMulta) + Val(vJuros) + Val(vDupl))
                    Rst.MoveNext
                Loop
            End With
    End If
    Rst.Close
End Sub
Private Function PgNumTitulosAtraso(idCliente As Integer) As Integer
    Dim Rst     As Recordset
    Dim sSQL    As String
  
    sSQL = "SELECT * FROM FinanceiroContasPRCadastro" & _
           " WHERE  ID_Empresa = " & ID_Empresa & _
           " AND idSacado = " & idCliente & _
           " AND Tabela = 'Clientes'" & _
           " AND Vencimento <= '" & Format(Date, "YYYY-MM-DD") & "'" & _
           " AND DataQuitacao IS NULL"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            PgNumTitulosAtraso = 0
        Else
            Rst.MoveLast
            PgNumTitulosAtraso = Rst.RecordCount
    End If
    Rst.Close
    
End Function
Private Function PgNumTitulosAvencer(idCliente As Integer) As Integer
    Dim Rst     As Recordset
    Dim sSQL    As String
    sSQL = "SELECT * FROM FinanceiroContasPRCadastro WHERE ID_Empresa = " & ID_Empresa & " AND idSacado = " & idCliente & " AND Tabela = 'Clientes' AND Vencimento > '" & Format(Date, "YYYY-MM-DD") & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            PgNumTitulosAvencer = 0
        Else
            Rst.MoveLast
            PgNumTitulosAvencer = Rst.RecordCount
    End If
    Rst.Close
    
End Function

Private Sub msfgHCont_Click()

    If msfgHCont.TextMatrix(msfgHCont.Row, 0) = "id" Or msfgHCont.TextMatrix(msfgHCont.Row, 0) = "" Then Exit Sub
    idContato = msfgHCont.TextMatrix(msfgHCont.Row, 0)
    txtDescricao.Text = msfgHCont.TextMatrix(msfgHCont.Row, 2)
End Sub

Private Sub msfgNF_Click()
    chvNFe = ""
    chvNFe = msfgNF.TextMatrix(msfgNF.Row, 4)
    idPV = msfgNF.TextMatrix(msfgNF.Row, 5)
    LstNFDescricao
End Sub

Private Sub msfgPV_Click()
    If Trim(msfgPV.TextMatrix(msfgPV.Row, 0)) = "PV" Or Trim(msfgPV.TextMatrix(msfgPV.Row, 0)) = "" Then
        idPV = 0
        Exit Sub
    End If
    idPV = Trim(msfgPV.TextMatrix(msfgPV.Row, 0))
    If Trim(msfgPV.TextMatrix(msfgPV.Row, 4)) = "" Then
        chvNFe = ""
        Exit Sub
    End If
    chvNFe = Mid(Trim(msfgPV.TextMatrix(msfgPV.Row, 4)), 1 + InStr(Trim(msfgPV.TextMatrix(msfgPV.Row, 4)), "["), Len(Trim(msfgPV.TextMatrix(msfgPV.Row, 4))))
    chvNFe = Left(chvNFe, Len(chvNFe) - 1)
End Sub



Private Sub optConsulta_Click(Index As Integer)
    frmConsCliente.Visible = False
    frmConsNF.Visible = False
    frmConsProduto.Visible = False
    Select Case Index
        Case 0
            frmConsCliente.Visible = True
            'frmConsNF.Visible = False
            LstClientes
            ResizeClientes
        Case 1
            'frmConsCliente.Visible = False
            frmConsNF.Visible = True
            LstNotasFiscais
            ResizeNF
        Case 2
            frmConsProduto.Visible = True
    End Select
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir PV"
            formFaturamentoPV.Show
        Case "Visualizar PV"
            ImpPV (idPV)
        Case "Imprimir DANFe"
            ImprimirDANFE2 (chvNFe)
        Case "Titulos Vencidos"
            LstTodosTitulosVencidos
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select
End Sub


Private Sub txtNFNome_Change()
    LstNotasFiscais
End Sub

Private Sub MontarBaseDeDados()
    Dim vDados(1000)    As Variant
    Dim contReg         As Integer
    Dim i               As Integer
    
    contReg = 0
    vDados(contReg) = Array("DtHrReg", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("idCliente", "50", "N"): contReg = contReg + 1
    vDados(contReg) = Array("Descricao", "5000", "S") ': contReg = contReg + 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg
    
End Sub
Private Sub LstNFDescricao()
    If Trim(chvNFe) = "" Then Exit Sub
    Dim sSQL As String
    Dim Rst As Recordset
    
    msfgDescricao.Rows = 1
    sSQL = "SELECT * FROM FaturamentoNFeItens WHERE  ID_Empresa = " & ID_Empresa & " AND idNFe = '" & chvNFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Rst.Close
        Else
            Rst.MoveFirst
    End If
    With msfgDescricao
        Do Until Rst.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Rst.Fields("ID")
            .TextMatrix(.Rows - 1, 1) = Rst.Fields("det_cProd")
            .TextMatrix(.Rows - 1, 2) = Rst.Fields("det_xProd")
            .TextMatrix(.Rows - 1, 3) = Rst.Fields("det_NCM")
            .TextMatrix(.Rows - 1, 4) = Rst.Fields("ICMS_Origem") & Rst.Fields("ICMS_CST")
            .TextMatrix(.Rows - 1, 5) = Rst.Fields("det_CFOP")
            .TextMatrix(.Rows - 1, 6) = Rst.Fields("det_uCom")
            .TextMatrix(.Rows - 1, 7) = ChkVal(Rst.Fields("det_qCom"), 0, cDecQtd)
            .TextMatrix(.Rows - 1, 8) = ChkVal(Rst.Fields("det_vUnCom"), 0, cDecMoeda)
            .TextMatrix(.Rows - 1, 9) = ChkVal(Rst.Fields("det_vProd"), 0, cDecMoeda)
            .TextMatrix(.Rows - 1, 10) = Rst.Fields("ICMS_vBC")
            .TextMatrix(.Rows - 1, 11) = Rst.Fields("ICMS_vICMS")
            .TextMatrix(.Rows - 1, 12) = Rst.Fields("IPI_vIPI")
            .TextMatrix(.Rows - 1, 13) = Rst.Fields("ICMS_pICMS")
            .TextMatrix(.Rows - 1, 14) = Rst.Fields("IPI_pIPI")
            Rst.MoveNext
        Loop
    End With
End Sub
Private Sub LstNFporProduto()
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim sTexto  As String
    
    'sTexto = Replace(Trim(txtPesqProd.Text), " ", "%")
    sTexto = MontStringConsultaProduto("FNFeI.det_xProd", Trim(txtPesqProd.Text))
   ' sSQL = "SELECT FNFe.idNFe, FNFe.ide_nNF, FNFe.ide_dEmi, FNFe.dest_xNome, " & _
                  "FNFeI.idNFe, FNFeI.det_xProd " & _
           "FROM FaturamentoNFe AS FNFe, FaturamentoNFeItens AS FNFeI " & _
           "WHERE FNFe.idNFe = FNFeI.idNFe AND FNFeI.det_xProd LIKE '%" & sTexto & "%'"
           
    sSQL = "SELECT Clientes.id, Clientes.Vendedor, FNFe.dest_idDest, FNFe.idNFe, FNFe.ide_nNF, FNFe.ide_dEmi, FNFe.dest_xNome, " & _
                  "FNFeI.idNFe, FNFeI.det_xProd " & _
           "FROM Clientes, FaturamentoNFe AS FNFe, FaturamentoNFeItens AS FNFeI " & _
           "WHERE Clientes.vendedor=" & idFunc & " AND FNFe.dest_idDest=Clientes.id AND " & _
                 "FNFe.idNFe = FNFeI.idNFe AND " & sTexto & _
           " ORDER BY FNFe.ide_nNF"
           
   
    Set Rst = RegistroBuscar(sSQL)
    If Rst Is Nothing Then
        DataGrid.Enabled = False
        'Text1.Enabled = False
        'Me.Caption = "Busca - [ 00000 Registros]"
        Rst.Close
        Exit Sub
        
    End If
    frmConsProduto.Caption = "Produto Faturado - [ " & ZE(Rst.RecordCount, 6) & " Registro(s) encontrado(s)]"
    'Rst.MoveFirst
            'dgProdFat.Enabled = True
            'DoEvents
            'DataGrid.ClearSelCols
            'DataGrid.Columns.Add 4
            DataGrid.Columns.Item(0).DataField = "IdNFe"
            DataGrid.Columns.Item(1).DataField = "ide_dEmi"
            DataGrid.Columns.Item(2).DataField = "ide_nNF"
            
            DataGrid.Columns.Item(3).DataField = "dest_xNome"
            DataGrid.Columns.Item(4).DataField = "det_xProd"
            
            'Set DataGrid.DataSource = Rst.DataSource
            'DataGrid.Columns(1).DataField = "FNFeI.det_xProd"
            'DataGrid.Columns(2).DataField = "FNFeI.det_xProd"
            Set DataGrid.DataSource = Rst.DataSource
            'DataGrid.Refresh
 
    'End If
    'Rst.Close
End Sub

Private Sub txtNome_Change()
    LstClientes
End Sub


Private Sub txtPesqProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtPesqProd.Text) <> "" Then
        LstNFporProduto
    End If

End Sub
