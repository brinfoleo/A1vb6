VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formContratosCRUD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contratos"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9390
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   60
      TabIndex        =   10
      Top             =   1800
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   9128
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Descrição "
      TabPicture(0)   =   "formContratosCRUD.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtDescricaoContrato"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtVlTotalContrato"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Materiais"
      TabPicture(1)   =   "formContratosCRUD.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "frmMaterialManutencao"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Funcionários"
      TabPicture(2)   =   "formContratosCRUD.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmFuncionarioManutencao"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      Begin VB.TextBox txtVlTotalContrato 
         Height          =   345
         Left            =   300
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   4140
         Width           =   1635
      End
      Begin VB.Frame frmFuncionarioManutencao 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   -72120
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   3615
         Begin VB.ComboBox cboFuncionario 
            Height          =   315
            Left            =   180
            TabIndex        =   31
            Text            =   "Combo1"
            Top             =   660
            Width           =   3195
         End
         Begin VB.CommandButton btnFunCancelar 
            Caption         =   "&Cancelar"
            Height          =   675
            Left            =   1860
            TabIndex        =   29
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnFunOk 
            Caption         =   "&Adicionar"
            Height          =   675
            Left            =   120
            TabIndex        =   28
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Funcionário:"
            Height          =   195
            Left            =   180
            TabIndex        =   30
            Top             =   420
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   25
         Top             =   540
         Width           =   8835
         Begin VB.TextBox txtFunTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   660
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   3900
            Width           =   1635
         End
         Begin VB.CommandButton btnFunRem 
            Caption         =   "-"
            Height          =   435
            Left            =   7620
            TabIndex        =   36
            Top             =   3960
            Width           =   495
         End
         Begin VB.CommandButton btnFunAdd 
            Caption         =   "+"
            Height          =   435
            Left            =   7020
            TabIndex        =   35
            Top             =   3960
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid msfgFuncionarios 
            Height          =   3615
            Left            =   180
            TabIndex        =   26
            Top             =   300
            Width           =   8475
            _ExtentX        =   14949
            _ExtentY        =   6376
            _Version        =   393216
            Cols            =   3
            SelectionMode   =   1
            FormatString    =   "^id  |<Funcionário                                    |>Salario                 "
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Total:"
            Height          =   195
            Left            =   180
            TabIndex        =   41
            Top             =   4080
            Width           =   435
         End
      End
      Begin VB.Frame frmMaterialManutencao 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   3075
         Left            =   -72420
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CommandButton btnMatCancelar 
            Caption         =   "&Cancelar"
            Height          =   615
            Left            =   2040
            TabIndex        =   24
            Top             =   2220
            Width           =   1875
         End
         Begin VB.CommandButton btnMatOk 
            Caption         =   "&Ok"
            Height          =   615
            Left            =   120
            TabIndex        =   23
            Top             =   2220
            Width           =   1875
         End
         Begin VB.TextBox txtVlTotal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   1740
            Width           =   1695
         End
         Begin VB.TextBox txtVlUnit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtQtd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   900
            Width           =   1695
         End
         Begin VB.ComboBox cboMaterial 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Text            =   "Combo1"
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Caption         =   "Vl. Total:"
            Height          =   255
            Left            =   1380
            TabIndex        =   19
            Top             =   1800
            Width           =   675
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Caption         =   "Vl. Unitário:"
            Height          =   195
            Left            =   1200
            TabIndex        =   18
            Top             =   1380
            Width           =   855
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Caption         =   "Quantidade:"
            Height          =   255
            Left            =   1200
            TabIndex        =   17
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Material:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4575
         Left            =   -74820
         TabIndex        =   12
         Top             =   480
         Width           =   8895
         Begin VB.TextBox txtMatTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   600
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   4020
            Width           =   1635
         End
         Begin VB.CommandButton btnMatRem 
            Caption         =   "-"
            Height          =   435
            Left            =   8100
            TabIndex        =   34
            Top             =   4020
            Width           =   555
         End
         Begin VB.CommandButton btnMatAdd 
            Caption         =   "+"
            Height          =   435
            Left            =   7560
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   4020
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid msfgMateriais 
            Height          =   3735
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   6588
            _Version        =   393216
            Cols            =   5
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"formContratosCRUD.frx":0054
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Total:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   4140
            Width           =   435
         End
      End
      Begin VB.TextBox txtDescricaoContrato 
         Height          =   3135
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "formContratosCRUD.frx":00DD
         Top             =   540
         Width           =   8835
      End
      Begin VB.Label Label11 
         Caption         =   "Valor do Contrato:"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   3900
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   9195
      Begin VB.CheckBox chkFinal 
         Alignment       =   1  'Right Justify
         Caption         =   "Termino:"
         Height          =   195
         Left            =   6720
         TabIndex        =   32
         Top             =   720
         Width           =   915
      End
      Begin VB.ComboBox cboTipoContrato 
         Height          =   315
         ItemData        =   "formContratosCRUD.frx":00E3
         Left            =   4380
         List            =   "formContratosCRUD.frx":00E5
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2115
      End
      Begin VB.TextBox txtNumContrato 
         Height          =   285
         Left            =   1020
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   780
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   315
         Left            =   7740
         TabIndex        =   3
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   90243073
         CurrentDate     =   42455
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   315
         Left            =   7740
         TabIndex        =   4
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   90243073
         CurrentDate     =   42455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Contrato:"
         Height          =   195
         Left            =   3060
         TabIndex        =   8
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Contrato Nº:"
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Início:"
         Height          =   195
         Left            =   6720
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   555
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Picture         =   "formContratosCRUD.frx":00E7
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratosCRUD.frx":0539
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratosCRUD.frx":0853
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratosCRUD.frx":10E5
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratosCRUD.frx":2337
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratosCRUD.frx":2C11
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratosCRUD.frx":34A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratosCRUD.frx":3D35
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratosCRUD.frx":4F87
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratosCRUD.frx":52A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratosCRUD.frx":55BB
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formContratosCRUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IdReg       As Integer 'Id do contrato na base de dados
Private idCliente   As Integer
Private idFun       As Integer
Private idMat       As Integer

Private lnMat       As Integer 'Linha do material selecionado
Private lnFun       As Integer 'Linha do Funcionario (mao de obra) selecionado
Private Sub carregaMatForm()
    If lnMat = 0 Then
        Exit Sub
    End If
    With msfgMateriais
        idMat = .TextMatrix(lnMat, 0)
        cboMaterial.Text = .TextMatrix(lnMat, 1)
        txtQtd.Text = .TextMatrix(lnMat, 2)
        txtVlUnit.Text = .TextMatrix(lnMat, 3)
        txtVlTotal.Text = .TextMatrix(lnMat, 4)
    End With
    frmMaterialManutencao.Visible = True
End Sub
Private Sub carregaFunForm()
    If lnFun = 0 Then
        Exit Sub
    End If
    With msfgFuncionarios
        idFun = .TextMatrix(lnFun, 0)
        cboFuncionario.Text = .TextMatrix(lnFun, 1)
    End With
    frmFuncionarioManutencao.Visible = True
End Sub
Private Sub salvar()
'Salva as alteracoes do form
    grvRegistro
End Sub
Private Function grvDB(complNomeTabela As String, vReg As Variant, cReg As Integer) As Boolean
    'Funcao responsavel na persistencia na base de dados
    Dim tmp As Boolean
    Dim strTabela As String
    strTabela = Mid(Me.Name, 5, Len(Me.Name)) & complNomeTabela
    
    If IdReg = 0 Then
            IdReg = RegistroIncluir(strTabela, vReg, cReg)
        Else
            tmp = RegistroAlterar(strTabela, vReg, cReg, "id=" & IdReg)
    End If
      
End Function

Private Function grvRegistro() As Boolean
  'On Error GoTo TrtErro
    Dim vDados(200) As Variant
    Dim cReg        As Integer 'Contador de Registros
    Dim l           As Integer
    Dim tmp         As Long
    cReg = 0
'    If ValidarPV = False Then
'        grvRegistro = False
'        Exit Function
'    End If
'
    
    
 '****************************************************************************
     
    'Cab
    vDados(cReg) = Array("numContrato", txtNumContrato.Text, "S"): cReg = cReg + 1
    vDados(cReg) = Array("tpContrato", cboTipoContrato.Text, "S"): cReg = cReg + 1
    vDados(cReg) = Array("idCliente", idCliente, "N"): cReg = cReg + 1
    vDados(cReg) = Array("dtIni", dtpInicial.Value, "D"): cReg = cReg + 1
    If chkFinal.Value = 1 Then
        vDados(cReg) = Array("dtFin", dtpFinal.Value, "D"): cReg = cReg + 1
    End If

    
    'Descricao
    vDados(cReg) = Array("descricao", txtDescricaoContrato.Text, "S"): cReg = cReg + 1
    vDados(cReg) = Array("vTotContrato", txtVlTotalContrato.Text, "S")  ': cReg = cReg + 1
    
    grvDB "", vDados, cReg
    
    
    'Material
    cReg = 0
    vDados(cReg) = Array("IdMaterial", "60", "S"): cReg = cReg + 1
    vDados(cReg) = Array("descricao", "150", "N"): cReg = cReg + 1
    vDados(cReg) = Array("qtd", "20", "N"): cReg = cReg + 1
    vDados(cReg) = Array("vUnit", "15", "S"): cReg = cReg + 1
    vDados(cReg) = Array("vTotalItem", "15", "S") ': cReg = cReg + 1
    
    grvDB "Materiais", vDados, cReg
    
    'Funcionarios
    cReg = 0
    vDados(cReg) = Array("IdFunc", "60", "S"): cReg = cReg + 1
    vDados(cReg) = Array("nome", "150", "N"): cReg = cReg + 1
    vDados(cReg) = Array("salario", "20", "N") ': cReg = cReg + 1
    
    grvDB "Funcionarios", vDados, cReg
 '****************************************************************************
'
'
'
'        tmp = RegistroIncluir(strTabela2, vReg, cReg)
'        If tmp = 0 Then
'                MsgBox "Erro ao Incluir o Produto"
'                grvRegistro = False
'                cReg = 0
'            Else
'                grvRegistro = True
'                cReg = 0
'        End If
End Function

Private Sub limpFormMaterial()
    idMat = 0
    cboMaterial.Text = ""
    txtQtd.Text = ""
    txtVlUnit.Text = ""
    txtVlTotal.Text = ""
    frmMaterialManutencao.Visible = False
    msfgMateriais.Enabled = True
End Sub

Private Sub limpFormFun()
    idFun = 0
    cboFuncionario.Text = ""
    frmFuncionarioManutencao.Visible = False
    msfgFuncionarios.Enabled = True
End Sub

Private Sub btnFunAdd_Click()
    adicionarFuncionario
    
End Sub
Private Sub adicionarFuncionario()
    lnFun = 0
    idFun = 0
    frmFuncionarioManutencao.Visible = True
    msfgFuncionarios.Enabled = False
End Sub
Private Sub btnFunCancelar_Click()
    limpFormFun
End Sub

Private Sub btnFunOk_Click()
    adicionarFunGrid
End Sub

Private Sub btnFunRem_Click()
    RemoverFunGrid
End Sub

Private Sub btnMatAdd_Click()
    addMaterial
End Sub

Private Sub addMaterial()
    idMat = 0
    lnMat = 0
    frmMaterialManutencao.Visible = True
    msfgMateriais.Enabled = False
End Sub
Private Sub btnMatCancelar_Click()
    cancelarAddMaterial
    
End Sub
Private Sub cancelarAddMaterial()
    limpFormMaterial
    msfgMateriais.Enabled = True
End Sub
Private Sub btnMatOk_Click()
    adicionarMaterialGrid
End Sub

Private Sub btnMatRem_Click()
    RemoverMatGrid
End Sub



Private Sub cboCliente_Click()
 If Trim(cboCliente.Text) = "" Then
        idCliente = 0
        Exit Sub
    End If
    PesquisarCliente "ID", Trim(Left(Trim(cboCliente.Text), 6)), "N"
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
            Rst.Close
    End If
End Sub

Private Sub cboCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarCliente
    End If
    
End Sub
Private Sub cboCliente_KeyPress(KeyAscii As Integer)
    idCliente = 0
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
    End If
    Rst.Close
End Sub
Private Sub PesquisarFuncionario(Optional sCampo As String, Optional sBusca As String, Optional SN As String)
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    If Trim(sCampo) = "" Then
        sBusca = formBuscar.IniciarBusca("rhfuncionariocadastro", "xNome,cpf,Cargo")
        sCampo = "Id"
        SN = "N"
        If Trim(sBusca) = 0 Then Exit Sub
    End If
    If SN = "N" Then
            sSQL = "SELECT * FROM rhfuncionariocadastro WHERE ID_Empresa = " & ID_Empresa & " AND " & sCampo & " = '" & sBusca & "'"
        Else
            sSQL = "SELECT * FROM rhfuncionariocadastro WHERE ID_Empresa = " & ID_Empresa & " AND " & sCampo & " = " & sBusca
    End If
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado"
        Else
            Rst.MoveFirst
            idFun = Rst.Fields("Id")
            cboFuncionario.Text = Trim(Rst.Fields("xNome"))
    End If
    Rst.Close
End Sub

Private Sub cboFuncionario_Click()
 If Trim(cboFuncionario.Text) = "" Then
        idFun = 0
        Exit Sub
    End If
    PesquisarFuncionario "ID", Trim(Left(Trim(cboFuncionario.Text), 6)), "N"
End Sub

Private Sub cboFuncionario_DropDown()
    Dim Rst As Recordset
    idFun = 0
    Set Rst = RegistroBuscar("SELECT * FROM rhfuncionariocadastro WHERE ID_Empresa = " & ID_Empresa & " AND xNome LIKE '" & cboFuncionario.Text & "%'")
    If Rst.BOF And Rst.EOF Then
            cboFuncionario.Clear
            Exit Sub
        Else
            cboFuncionario.Clear
            Rst.MoveFirst
            Do Until Rst.EOF
                cboFuncionario.AddItem Left(String(6, "0"), 6 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & _
                                   " - " & _
                                   Rst.Fields("xNome")
                Rst.MoveNext
            Loop
            Rst.Close
    End If
End Sub

Private Sub cboFuncionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarFuncionario
    End If
End Sub

Private Sub cboFuncionario_KeyPress(KeyAscii As Integer)
    idFun = 0
    If KeyAscii = 13 And IsNumeric(Trim(cboFuncionario.Text)) Then
        PesquisarFuncionario "ID", Trim(Left(Trim(cboFuncionario.Text), 6)), "N"
    End If

End Sub
'******************************************************
Private Sub PesquisarMaterial(Optional sCampo As String, Optional sBusca As String, Optional SN As String)
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    If Trim(sCampo) = "" Then
        sBusca = formBuscar.IniciarBusca("estoqueproduto", , , , , "status='ativo'")
        sCampo = "Id"
        SN = "N"
        If Trim(sBusca) = 0 Then Exit Sub
    End If
    If SN = "N" Then
            sSQL = "SELECT * FROM estoqueproduto WHERE ID_Empresa = " & ID_Empresa & " AND Deposito=" & ID_Deposito & " AND status = 'ativo' AND " & sCampo & " = '" & sBusca & "'"
        Else
            sSQL = "SELECT * FROM estoqueproduto WHERE ID_Empresa = " & ID_Empresa & " AND Deposito=" & ID_Deposito & " AND status = 'ativo' AND " & sCampo & " = " & sBusca
    End If
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado"
        Else
            Rst.MoveFirst
            idMat = Rst.Fields("Id")
            cboMaterial.Text = Trim(Rst.Fields("descricao"))
            txtVlUnit.Text = ChkVal(Trim(Rst.Fields("preco")), 0, cDecMoeda)
    End If
    Rst.Close
End Sub

Private Sub cboMaterial_Click()
 If Trim(cboMaterial.Text) = "" Then
        idFun = 0
        Exit Sub
    End If
    PesquisarMaterial "ID", Trim(Left(Trim(cboMaterial.Text), 6)), "N"
End Sub

Private Sub cboMaterial_DropDown()
    Dim Rst As Recordset
    idMat = 0
    Set Rst = RegistroBuscar("SELECT * FROM estoqueproduto WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND status='ativo' AND descricao LIKE '" & cboMaterial.Text & "%'")
    If Rst.BOF And Rst.EOF Then
            cboMaterial.Clear
            Exit Sub
        Else
            cboMaterial.Clear
            Rst.MoveFirst
            Do Until Rst.EOF
                cboMaterial.AddItem Left(String(6, "0"), 6 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & _
                                   " - " & _
                                   Rst.Fields("descricao")
                Rst.MoveNext
            Loop
            Rst.Close
    End If
End Sub

Private Sub cboMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarMaterial
    End If
End Sub

Private Sub cboMaterial_KeyPress(KeyAscii As Integer)
    idMat = 0
    If KeyAscii = 13 And IsNumeric(Trim(cboMaterial.Text)) Then
        PesquisarMaterial "ID", Trim(Left(Trim(cboMaterial.Text), 6)), "N"
    End If
End Sub
'******************************************************
Private Sub cboTipoContrato_DropDown()
With cboTipoContrato
    .Clear
    .AddItem ("01 - Show")
    .AddItem ("02 - Festa")
    .AddItem ("03 - Feira")
    .AddItem ("04 - Recepção")
    .AddItem ("05 - Fixo")
End With
End Sub

Private Sub chkFinal_Click()
    If chkFinal.Value = 0 Then
            dtpFinal.Enabled = False
        Else
            dtpFinal.Enabled = True
    End If
End Sub

Private Sub Form_Load()
   LimpForm
End Sub
Private Sub LimpForm()
    chkFinal.Value = 0
    dtpFinal.Enabled = False
    LimpaFormulario Me
    SSTab1.Tab = 0
    msfgMateriais.Rows = 1
    msfgFuncionarios.Rows = 1
    lnMat = 0
End Sub
Private Sub calcMatTotalItem()
    
    Dim a, b, t As String
    a = txtQtd.Text
    b = txtVlUnit.Text
    t = Val(a) * Val(b)
    txtVlTotal.Text = ChkVal(t, 0, cDecMoeda)
End Sub

Private Sub msfgFuncionarios_Click()
    lnFun = msfgFuncionarios.RowSel
    End Sub


Private Sub msfgFuncionarios_DblClick()
    If lnFun <> 0 Then
        carregaFunForm
    End If

End Sub

Private Sub msfgMateriais_Click()
    lnMat = msfgMateriais.RowSel
End Sub

Private Sub msfgMateriais_DblClick()
    
    If lnMat <> 0 Then
        carregaMatForm
    End If
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Salvar"
            salvar
        Case "Cancelar"
            'Cancelar
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select
End Sub

Private Sub txtFunTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub txtMatTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub txtQtd_Change()
    calcMatTotalItem
End Sub

Private Sub txtQtd_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtQtd.Text, KeyAscii, cDecQtd)
End Sub

Private Sub txtVlTotalContrato_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtVlTotalContrato.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtVlTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtVlUnit_Change()
    calcMatTotalItem
End Sub

Private Sub txtVlUnit_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtVlUnit.Text, KeyAscii, cDecMoeda)
End Sub
Private Sub adicionarMaterialGrid()
    On Error GoTo TrtErroGridMaterial
    
    'Valida Adicao
    If idMat = 0 Or Val(ChkVal(Trim(txtQtd.Text), 0, cDecQtd)) <= 0 Then
        MsgBox "Selecione um material e insira a quantidade!", vbInformation, App.EXEName
        Exit Sub
    End If
    
    
    With msfgMateriais
        'Add +1 linha ao grid
        If lnMat = 0 Then
            .Rows = .Rows + 1
            lnMat = .Rows - 1
        End If
    
        .TextMatrix(lnMat, 0) = idMat
        .TextMatrix(lnMat, 1) = cboMaterial.Text
        .TextMatrix(lnMat, 2) = ChkVal(txtQtd.Text, 0, cDecQtd)
        .TextMatrix(lnMat, 3) = ChkVal(txtVlUnit.Text, 0, cDecMoeda)
        .TextMatrix(lnMat, 4) = ChkVal(txtVlTotal.Text, 0, cDecMoeda)
    End With
    lnMat = 0
    idMat = 0
    limpFormMaterial
    calcTotMat
    
    Exit Sub
TrtErroGridMaterial:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub adicionarFunGrid()
    On Error GoTo TrtErroGridFun
    
    'Valida Adicao
    If idFun = 0 Or Trim(cboFuncionario.Text) = "" Then
        MsgBox "Selecione um funcionário!", vbInformation, App.EXEName
        Exit Sub
    End If
    
    
    With msfgFuncionarios
        'Add +1 linha ao grid
        If lnFun = 0 Then
            .Rows = .Rows + 1
            lnFun = .Rows - 1
        End If
    
        .TextMatrix(lnFun, 0) = idFun
        .TextMatrix(lnFun, 1) = cboFuncionario.Text
        .TextMatrix(lnFun, 2) = ChkVal(PgDadosRhFuncionario(idFun).Salario, 0, cDecMoeda)
        
    End With
    lnFun = 0
    idFun = 0
    limpFormFun
    calcTotFun
    
    Exit Sub
TrtErroGridFun:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub
Private Sub RemoverMatGrid()
    If lnMat = 0 Then Exit Sub
    If MsgBox("Deseja realmente remover este item?", vbYesNo, App.EXEName) = vbYes Then
        If msfgMateriais.Rows = 2 Then
                msfgMateriais.Rows = 1
                'lnMat = 0
                'idMat = 0
            Else
                msfgMateriais.RemoveItem msfgMateriais.Row
                'lnMat = 0
                'idMat = 0
        End If
        lnMat = 0
        idMat = 0
        limpFormMaterial
    End If
    calcTotMat
End Sub
Private Sub RemoverFunGrid()
    If lnFun = 0 Then Exit Sub
    If MsgBox("Deseja realmente remover este funcionario?", vbYesNo, App.EXEName) = vbYes Then
        If msfgFuncionarios.Rows = 2 Then
                msfgFuncionarios.Rows = 1
                
                
            Else
                msfgFuncionarios.RemoveItem msfgFuncionarios.Row
                
        End If
        limpFormFun
        lnFun = 0
        idFun = 0
    End If
    calcTotFun
End Sub

Private Function calcTotMat() As String
'Calcula o total do material lancado
    Dim l As Integer
    Dim vlLin As String
    Dim vlTotal As String
    
    With msfgMateriais
        For i = 1 To .Rows - 1
            vlLin = ChkVal(.TextMatrix(i, 4), 0, cDecMoeda)
            vlTotal = Val(ChkVal(vlTotal, 0, cDecMoeda)) + Val(ChkVal(vlLin, 0, cDecMoeda))
        Next
    End With
    txtMatTotal.Text = ChkVal(vlTotal, 0, cDecMoeda)
    calcTotMat = ChkVal(vlTotal, 0, cDecMoeda)
    
End Function
Private Function calcTotFun() As String
'Calcula o total do salario dos funcionarios
    Dim l As Integer
    Dim vlLin As String
    Dim vlTotal As String
    
    With msfgFuncionarios
        For i = 1 To .Rows - 1
            vlLin = ChkVal(.TextMatrix(i, 2), 0, cDecMoeda)
            vlTotal = Val(ChkVal(vlTotal, 0, cDecMoeda)) + Val(ChkVal(vlLin, 0, cDecMoeda))
        Next
    End With
    txtFunTotal.Text = ChkVal(vlTotal, 0, cDecMoeda)
    calcTotFun = ChkVal(vlTotal, 0, cDecMoeda)
    
End Function

Private Sub MontarBaseDeDados()
    Dim vDados(1000)    As Variant
    Dim cReg         As Integer
    Dim i               As Integer
    
    cReg = 0
    
   
    
    'Cab
    vDados(cReg) = Array("numContrato", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("tpContrato", "100", "S"): cReg = cReg + 1
    vDados(cReg) = Array("idCliente", "6", "N"): cReg = cReg + 1
    vDados(cReg) = Array("dtIni", "15", "D"): cReg = cReg + 1
    vDados(cReg) = Array("dtFin", "15", "D"): cReg = cReg + 1
    '************************************************************************
    
    
    'Descricao
    vDados(cReg) = Array("descricao", "65000", "S"): cReg = cReg + 1
    vDados(cReg) = Array("vTotContrato", "15", "S") ': cReg = cReg + 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, cReg, ""
    
    
    'Material
    cReg = 0
    vDados(cReg) = Array("IdMaterial", "60", "S"): cReg = cReg + 1
    vDados(cReg) = Array("descricao", "150", "N"): cReg = cReg + 1
    vDados(cReg) = Array("qtd", "20", "N"): cReg = cReg + 1
    vDados(cReg) = Array("vUnit", "15", "S"): cReg = cReg + 1
    vDados(cReg) = Array("vTotalItem", "15", "S") ': cReg = cReg + 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, cReg, "Materiais"
    
    'funcionarios
    cReg = 0
    vDados(cReg) = Array("IdFunc", "60", "S"): cReg = cReg + 1
    vDados(cReg) = Array("nome", "150", "N"): cReg = cReg + 1
    vDados(cReg) = Array("salario", "20", "N") ': cReg = cReg + 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, cReg, "Funcionarios"
End Sub

