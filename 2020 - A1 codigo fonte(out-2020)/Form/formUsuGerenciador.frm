VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formUsuGerenciador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciamento de Usuarios & Grupos"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10665
   Begin TabDlg.SSTab sstGerenciador 
      Height          =   6375
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Usuarios"
      TabPicture(0)   =   "formUsuGerenciador.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Grupos"
      TabPicture(1)   =   "formUsuGerenciador.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   5775
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   10275
         Begin VB.Frame Frame5 
            Height          =   3555
            Left            =   5160
            TabIndex        =   17
            Top             =   1980
            Width           =   4995
            Begin MSComctlLib.TreeView trvGrupos 
               Height          =   3195
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   5636
               _Version        =   393217
               Style           =   7
               Checkboxes      =   -1  'True
               Appearance      =   1
            End
         End
         Begin VB.Frame Frame4 
            Height          =   1815
            Left            =   5160
            TabIndex        =   12
            Top             =   180
            Width           =   4995
            Begin VB.TextBox txtGrupoDescricao 
               Height          =   675
               Left            =   180
               TabIndex        =   16
               Text            =   "Text1"
               Top             =   960
               Width           =   4695
            End
            Begin VB.TextBox txtGrupoNome 
               Height          =   285
               Left            =   1380
               TabIndex        =   15
               Text            =   "Text1"
               Top             =   240
               Width           =   3495
            End
            Begin VB.Label Label4 
               Caption         =   "Descrição:"
               Height          =   195
               Left            =   120
               TabIndex        =   14
               Top             =   660
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Nome do grupo:"
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   300
               Width           =   1155
            End
         End
         Begin MSFlexGridLib.MSFlexGrid msfgGrupos 
            Height          =   5415
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   9551
            _Version        =   393216
            Cols            =   3
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   "^id   |<Nome do grupo            |<Descrição                         "
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   10275
         Begin VB.Frame Frame6 
            Caption         =   "Menus"
            Height          =   2115
            Left            =   4980
            TabIndex        =   24
            Top             =   3480
            Width           =   5175
            Begin VB.ListBox lstUsuMenu 
               Height          =   1635
               Left            =   180
               Style           =   1  'Checkbox
               TabIndex        =   25
               Top             =   300
               Width           =   4695
            End
         End
         Begin VB.Frame Frame2 
            Height          =   3195
            Left            =   4980
            TabIndex        =   3
            Top             =   180
            Width           =   5175
            Begin VB.CheckBox chkSuperUsuario 
               Caption         =   "Super-usuário"
               Height          =   195
               Left            =   240
               TabIndex        =   23
               Top             =   2880
               Width           =   2775
            End
            Begin VB.ComboBox cboUsuNome 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   540
               Width           =   4755
            End
            Begin VB.ComboBox cboGrupo 
               Height          =   315
               Left            =   720
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   1440
               Width           =   4095
            End
            Begin VB.TextBox txtUsuLogin 
               Height          =   285
               Left            =   720
               TabIndex        =   7
               Text            =   "Text1"
               Top             =   960
               Width           =   3015
            End
            Begin VB.CheckBox chkSenhaNuncaExpira 
               Caption         =   "Senha nunca expira"
               Height          =   255
               Left            =   240
               TabIndex        =   6
               Top             =   2580
               Width           =   2895
            End
            Begin VB.CheckBox chkTrocarSenha 
               Caption         =   "Trocar senha no proximo acesso"
               Height          =   195
               Left            =   240
               TabIndex        =   5
               Top             =   2340
               Width           =   3735
            End
            Begin VB.CheckBox chkSenhaPadrao 
               Caption         =   "Senha padrão ( 123 )"
               Height          =   255
               Left            =   240
               TabIndex        =   4
               Top             =   2040
               Width           =   3375
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Grupo:"
               Height          =   195
               Left            =   120
               TabIndex        =   19
               Top             =   1500
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "Nome do Usuario:"
               Height          =   195
               Left            =   120
               TabIndex        =   9
               Top             =   300
               Width           =   1455
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Login:"
               Height          =   195
               Left            =   180
               TabIndex        =   8
               Top             =   1020
               Width           =   495
            End
         End
         Begin MSFlexGridLib.MSFlexGrid msfgUsuarios 
            Height          =   5415
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   9551
            _Version        =   393216
            Cols            =   3
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   "^id   |<Nome                                |<Login                "
         End
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir Usuario"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alterar Usuario"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excluir Usuario"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir Grupo"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alterar Grupo"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excluir Grupo"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Manutenção da Tabela"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5400
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   18
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":0038
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":048A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":07A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":1036
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":2288
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":2B62
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":33F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":3C86
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":4ED8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":51F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":550C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":5903
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":5E9D
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":6437
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":69D1
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":6F6B
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":7505
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuGerenciador.frx":7BFF
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formUsuGerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdUsu       As Integer
Dim idGrupo     As Integer
Dim strTabela   As String
Private Sub LstMenus(Optional iUsu As Integer)
    Dim i       As Integer
    Dim Menu    As String
    With lstUsuMenu
        .Clear
        .AddItem "01 - Comercial"
        .AddItem "02 - Faturamento"
        .AddItem "03 - Financeiro"
        If iUsu = 0 Then
                Exit Sub
            Else
                Menu = PgDadosUsuario(iUsu).Menus
                For i = 1 To Len(Trim(Menu))
                    If Mid(Menu, i, 1) = 1 Then
                            .Selected(i - 1) = True
                        Else
                            .Selected(i - 1) = False
                    End If
                Next
        End If
    End With
End Sub
Private Sub LimpFormUsu()
    cboUsuNome.Clear
    txtUsuLogin.Text = ""
    cboGrupo.Clear
    chkSenhaPadrao.Value = 0
    chkTrocarSenha.Value = 0
    chkSenhaNuncaExpira.Value = 0
    chkSuperUsuario.Value = 0
End Sub
Private Sub LimpFormGrupo()
    txtGrupoNome.Text = ""
    txtGrupoDescricao.Text = ""
End Sub
Private Sub LstUsuarios()
    Dim rst     As Recordset
    Dim sSQL    As String
    msfgUsuarios.Rows = 1
    sSQL = "SELECT * FROM UsuGerenciador WHERE ID_Empresa = " & ID_Empresa
    Set rst = RegistroBuscar(sSQL)
      If rst.BOF And rst.EOF Then
        Else
            rst.MoveFirst
            Do Until rst.EOF
                msfgUsuarios.Rows = msfgUsuarios.Rows + 1
                msfgUsuarios.TextMatrix(msfgUsuarios.Rows - 1, 0) = rst.Fields("id")
                msfgUsuarios.TextMatrix(msfgUsuarios.Rows - 1, 1) = IIf(IsNull(rst.Fields("Usu_Nome")), "", rst.Fields("Usu_Nome"))
                msfgUsuarios.TextMatrix(msfgUsuarios.Rows - 1, 2) = IIf(IsNull(rst.Fields("Usu_Login")), "", rst.Fields("Usu_Login"))
                rst.MoveNext
            Loop
    End If
    rst.Close
End Sub
Private Sub LstGrupos()
    On Error Resume Next
    Dim rst     As Recordset
    Dim sSQL    As String
    msfgGrupos.Rows = 1
    sSQL = "SELECT * FROM UsuGerenciadorGrupo WHERE ID_Empresa = " & ID_Empresa
    Set rst = RegistroBuscar(sSQL)
    If rst.BOF And rst.EOF Then
        Else
            rst.MoveFirst
            Do Until rst.EOF
                msfgGrupos.Rows = msfgGrupos.Rows + 1
                msfgGrupos.TextMatrix(msfgGrupos.Rows - 1, 0) = rst.Fields("id")
                msfgGrupos.TextMatrix(msfgGrupos.Rows - 1, 1) = rst.Fields("grupo_Nome")
                msfgGrupos.TextMatrix(msfgGrupos.Rows - 1, 2) = rst.Fields("grupo_Descricao")
                rst.MoveNext
            Loop
    End If
    rst.Close
End Sub

Private Sub cboGrupo_DropDown()
    Dim rst     As Recordset
    Dim sSQL    As String
    cboGrupo.Clear
    sSQL = "SELECT * FROM UsuGerenciadorGrupo WHERE ID_Empresa = " & ID_Empresa
    Set rst = RegistroBuscar(sSQL)
    If rst.BOF And rst.EOF Then
        Else
            rst.MoveFirst
            Do Until rst.EOF
                cboGrupo.AddItem Left("0000", 4 - Len(rst.Fields("Id"))) & rst.Fields("id") & " - " & rst.Fields("grupo_nome")
                rst.MoveNext
            Loop
    End If
    rst.Close
End Sub

Private Sub cboUsuNome_DropDown()
    Dim rst     As Recordset
    Dim sSQL    As String
    cboUsuNome.Clear
    sSQL = "SELECT * FROM RHFuncionarioCadastro WHERE ID_Empresa = " & ID_Empresa
    Set rst = RegistroBuscar(sSQL)
    If rst.BOF And rst.EOF Then
        Else
            rst.MoveFirst
            Do Until rst.EOF
                cboUsuNome.AddItem Left("000", 3 - Len(rst.Fields("id"))) & rst.Fields("id") & _
                                   " - " & rst.Fields("xNome")
                rst.MoveNext
            Loop
    End If
    rst.Close
End Sub

Private Sub chkSenhaPadrao_Click()
    If chkSenhaPadrao.Value = 1 Then
            chkTrocarSenha.Value = 1
            chkTrocarSenha.Enabled = False
        Else
        chkTrocarSenha.Value = 0
        chkTrocarSenha.Enabled = True
    End If
End Sub
Private Sub MontarNiveis()
    Dim rst     As Recordset
    Dim sSQL    As String
    trvGrupos.Nodes.Clear
    'sSQL = "SELECT * FROM UsuGerenciadorFormularios WHERE ID_Empresa = " & ID_Empresa & " ORDER BY Descricao"
    sSQL = "SELECT * FROM UsuGerenciadorFormularios ORDER BY Descricao"
    Set rst = RegistroBuscar(sSQL)
    If rst.BOF And rst.EOF Then
            trvGrupos.Nodes.Clear
        Else
            rst.MoveFirst
            Do Until rst.EOF
                Niveis IIf(IsNull(rst.Fields("Descricao")), rst.Fields("Formulario"), rst.Fields("Descricao")), rst.Fields("Formulario")
                rst.MoveNext
            Loop
    End If
    
    'Niveis "Cliente", "Teste", "formClientesRp"
End Sub
Private Sub Niveis(nomeGrupo As String, nmForm As String)
    On Error Resume Next
'Private Sub Niveis(nomeGrupo As String, subGrupo As String, nmForm As String)
    Dim i       As Integer
    Dim Existe  As Boolean
'Grupo
    Existe = False
    For i = 1 To trvGrupos.Nodes.Count
        If trvGrupos.Nodes.item(i).Text = nomeGrupo Then
            Existe = True
            
        End If
    Next
    If Existe = False Then
        trvGrupos.Nodes.Add , , nmForm, nomeGrupo ', 1
    End If
'SubGrupos
    'trvGrupos.Nodes.Add nomeGrupo & "1", tvwChild, nmForm, subGrupo ', 2
'Opcoes
    'trvGrupos.Nodes.Add nmForm, tvwChild, subGrupo & "1", "Incluir"  ', 4
    DoEvents
    trvGrupos.Nodes.Add nmForm, tvwChild, nmForm & "1", "Incluir" ', 4
    trvGrupos.Nodes.item(nmForm & "1").Checked = FormPermissao(nmForm, "n", idGrupo)
    If FormPermissao(nmForm, "n", idGrupo) = True Then trvGrupos.Nodes.item(nmForm).Checked = True

    trvGrupos.Nodes.Add nmForm, tvwChild, nmForm & "2", "Alterar"  ', 4
    trvGrupos.Nodes.item(nmForm & "2").Checked = FormPermissao(nmForm, "a", idGrupo)
    If FormPermissao(nmForm, "a", idGrupo) = True Then trvGrupos.Nodes.item(nmForm).Checked = True
    
    trvGrupos.Nodes.Add nmForm, tvwChild, nmForm & "3", "Excluir"  ', 4
    trvGrupos.Nodes.item(nmForm & "3").Checked = FormPermissao(nmForm, "e", idGrupo)
    If FormPermissao(nmForm, "e", idGrupo) = True Then trvGrupos.Nodes.item(nmForm).Checked = True
    
    trvGrupos.Nodes.Add nmForm, tvwChild, nmForm & "4", "Imprimir"   ', 4
    trvGrupos.Nodes.item(nmForm & "4").Checked = FormPermissao(nmForm, "i", idGrupo)
    If FormPermissao(nmForm, "i", idGrupo) = True Then trvGrupos.Nodes.item(nmForm).Checked = True
    
    trvGrupos.Nodes.Add nmForm, tvwChild, nmForm & "5", "Consultar"   ', 4
    trvGrupos.Nodes.item(nmForm & "5").Checked = FormPermissao(nmForm, "c", idGrupo)
    If FormPermissao(nmForm, "c", idGrupo) = True Then trvGrupos.Nodes.item(nmForm).Checked = True
    
    
End Sub



Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    HDForm Me, False
    HDMenu Me, True
    msfgUsuarios.Enabled = True
    msfgGrupos.Enabled = True
    LimpaFormulario Me
    sstGerenciador.Tab = 0
    LstUsuarios
    LstGrupos
End Sub


Private Sub msfgGrupos_Click()
    With msfgGrupos
        idGrupo = .TextMatrix(.Row, 0)
        
        txtGrupoNome.Text = IIf(Trim(.TextMatrix(.Row, 1)) = "", "", .TextMatrix(.Row, 1))
        txtGrupoDescricao.Text = IIf(Trim(.TextMatrix(.Row, 2)) = "", "", .TextMatrix(.Row, 2))
        MontarNiveis

    End With
End Sub

Private Sub msfgUsuarios_Click()
    With msfgUsuarios
        IdUsu = .TextMatrix(.Row, 0)
        cboUsuNome.Clear
        cboUsuNome.AddItem IIf(Trim(.TextMatrix(.Row, 1)) = "", " ", .TextMatrix(.Row, 1))
        cboUsuNome.Text = cboUsuNome.List(0)
        txtUsuLogin.Text = .TextMatrix(.Row, 2)
        cboGrupo.Clear
        If PgDadosUsuario(IdUsu).Grupo <> 0 Then
            cboGrupo.AddItem Left("0000", 4 - Len(PgDadosUsuario(IdUsu).Grupo)) & PgDadosUsuario(IdUsu).Grupo _
                             & " - " & PgDadosUsuGrupo(PgDadosUsuario(IdUsu).Grupo).Nome
            cboGrupo.Text = cboGrupo.List(0)
        End If
        chkTrocarSenha.Value = PgDadosUsuario(IdUsu).TrocarSenha
        chkSenhaNuncaExpira.Value = PgDadosUsuario(IdUsu).SenhaNuncaExp
        chkSuperUsuario.Value = PgDadosUsuario(IdUsu).SuperUsuario
        LstMenus (IdUsu)
    End With
End Sub




Private Sub IncluirUsuario()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    sstGerenciador.Tab = 0
    IdUsu = 0
    HDForm Me, True
    HDMenu Me, False
    msfgUsuarios.Enabled = False
    LimpFormUsu
    strTabela = "UsuGerenciador"
End Sub
Private Sub AlterarUsuario()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If IdUsu = 0 Then
        MsgBox "Selecione um Usuario.", vbInformation, "Aviso"
        Exit Sub
    End If
    sstGerenciador.Tab = 0
    HDForm Me, True
    HDMenu Me, False
    cboUsuNome.Enabled = False
    msfgUsuarios.Enabled = False
    strTabela = "UsuGerenciador"
End Sub
Private Sub ExcluirUsuario()
    If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If
    sstGerenciador.Tab = 0
    If IdUsu = 0 Then
        MsgBox "Selecione um Usuario.", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir o Usuario " & cboUsuNome.Text & ".", vbInformation + vbYesNo, "Excluir") = vbYes Then
        If RegistroExcluir("UsuGerenciador", "id = " & IdUsu) = False Then
                MsgBox "Erro ao excluir"
            Else
                LstUsuarios
                LimpFormUsu
        End If
    End If
End Sub
Private Sub IncluirGrupo()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    sstGerenciador.Tab = 1
    idGrupo = 0
    HDForm Me, True
    HDMenu Me, False
    msfgGrupos.Enabled = False
    LimpFormGrupo
    strTabela = "UsuGerenciadorGrupo"
End Sub
Private Sub AlterarGrupo()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    sstGerenciador.Tab = 1
    If idGrupo = 0 Then
        MsgBox "Selecione um Grupo.", vbInformation, "Aviso"
        Exit Sub
    End If
    HDForm Me, True
    HDMenu Me, False
    msfgGrupos.Enabled = False
    strTabela = "UsuGerenciadorGrupo"
End Sub
Private Sub ExcluirGrupo()
    If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If
    sstGerenciador.Tab = 1
    If idGrupo = 0 Then
        MsgBox "Selecione um Grupo.", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir o grupo " & txtGrupoNome.Text & ".", vbInformation + vbYesNo, "Excluir") = vbYes Then
        If RegistroExcluir("UsuGerenciadorGrupo", "id = " & idGrupo) = False Then
                MsgBox "Erro ao excluir"
            Else
                LstGrupos
                LimpFormGrupo
        End If
    End If
End Sub
Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir Usuario"
            IncluirUsuario
        Case "Alterar Usuario"
            AlterarUsuario
        Case "Excluir Usuario"
            ExcluirUsuario
        Case "Incluir Grupo"
            IncluirGrupo

        Case "Alterar Grupo"
            AlterarGrupo
        Case "Excluir Grupo"
            ExcluirGrupo
        Case "Salvar"
            If strTabela = "UsuGerenciador" Then
                    If grvRegistroUsu = True Then
                        HDForm Me, False
                        HDMenu Me, True
                        msfgUsuarios.Enabled = True
                        msfgGrupos.Enabled = True
                        IdUsu = 0
                        idGrupo = 0
                        LimpFormUsu
                        LstUsuarios
                        Exit Sub
                    End If
                ElseIf strTabela = "UsuGerenciadorGrupo" Then
                    If grvRegistroGrupo = True Then
                        HDForm Me, False
                        HDMenu Me, True
                        msfgGrupos.Enabled = True
                        msfgUsuarios.Enabled = True
                        IdUsu = 0
                        idGrupo = 0
                        LimpFormGrupo
                        LstGrupos
                        Exit Sub
                    End If
            End If
        Case "Cancelar"
            HDForm Me, False
            HDMenu Me, True
            msfgGrupos.Enabled = True
            msfgUsuarios.Enabled = True
            IdUsu = 0
            idGrupo = 0
            strTabela = ""
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    
    End Select
End Sub


Private Sub trvGrupos_NodeCheck(ByVal Node As MSComctlLib.Node)
       
   trvGrupos.SelectedItem = Node
    
    Dim idx     As Integer
    Dim Idx2    As Integer
    Dim idx3    As Integer
    On Error GoTo ErrStat
    '***************************  Quando marcar os filhos marca o pai  **********************************
    Dim m As Integer
    
    If trvGrupos.SelectedItem.Children Then
        For idx = trvGrupos.SelectedItem.Child.FirstSibling.Index To trvGrupos.SelectedItem.Child.LastSibling.Index
           trvGrupos.Nodes.item(idx).Checked = trvGrupos.SelectedItem.Checked
            If trvGrupos.Nodes.item(idx).Children Then
                For Idx2 = trvGrupos.Nodes.item(idx).Child.FirstSibling.Index To trvGrupos.Nodes.item(idx).Child.LastSibling.Index
                   trvGrupos.Nodes.item(Idx2).Checked = trvGrupos.Nodes.item(idx).Checked
                   '***********************
                   If trvGrupos.Nodes.item(Idx2).Children Then
                        For idx3 = trvGrupos.Nodes.item(Idx2).Child.FirstSibling.Index To trvGrupos.Nodes.item(Idx2).Child.LastSibling.Index
                            trvGrupos.Nodes.item(idx3).Checked = trvGrupos.Nodes.item(Idx2).Checked
                        Next
                    End If
                   '*************************
                Next
            End If
        Next
    End If

    '**************  Quando eu marcar o pai, os filhos daquele pai serão marcados também  **************
    If Node.Checked Then
        If Not Node.Parent Is Nothing Then
            Node.Parent.Checked = True
            If Not Node.Parent.Parent Is Nothing Then
                Node.Parent.Parent.Checked = True
                'If Not Node.Parent.Parent Is Nothing Then
                '    Node.Parent.Parent.Checked = True
                'End If
            End If
        End If
    End If
    
    Exit Sub
    
ErrStat:
    If Err.Number = 91 And InStr(1, trvGrupos.SelectedItem, "Store") Then
         MsgBox trvGrupos.SelectedItem & " não possui filho(s) !", vbCritical, "Erro"
    ElseIf Err.Number = 91 Then
         MsgBox "O nodo não está selecionado !", vbCritical, "Erro"
    End If

End Sub

Public Sub MontarBaseDeDados()
    Dim vDados(1000)    As Variant
    Dim contReg         As Integer

    'Usuarios
    contReg = 0
    vDados(contReg) = Array("Usu_Nome", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Usu_Login", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Usu_Senha", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Usu_Grupo", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("Usu_TrocaSenha", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("Usu_SenhaNuncaExpira", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("Usu_SuperUsuario", "1", "N"): contReg = contReg + 1
    vDados(contReg) = Array("Usu_Menus", "10", "S") ': contReg = contReg + 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg
    'Grupos
    contReg = 0
    vDados(contReg) = Array("grupo_Nome", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("grupo_Descricao", "200", "S") ': contReg = contReg + 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "Grupo"
    'Formularios
    contReg = 0
    vDados(contReg) = Array("Formulario", "200", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Descricao", "200", "S") ': contReg = contReg + 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "Formularios"
    'Acessos
    contReg = 0
    vDados(contReg) = Array("GrupoId", "200", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Formulario", "200", "S"): contReg = contReg + 1
    vDados(contReg) = Array("Permissao", "50", "S") ': contReg = contReg + 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "Acessos"
End Sub

Private Function grvRegistroGrupo() As Boolean
    Dim vReg(100)   As Variant
    Dim cReg        As Integer 'Contador de Registros
    Dim i           As Integer
    Dim X           As Integer
    Dim Permissao   As String
    Dim Formulario  As String
    cReg = 0
    vReg(cReg) = Array("grupo_Nome", txtGrupoNome.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("grupo_Descricao", txtGrupoDescricao.Text, "S") ': cReg = cReg + 1
    If idGrupo = 0 Then
    idGrupo = RegistroIncluir("UsuGerenciadorGrupo", vReg, cReg)
            If idGrupo = 0 Then
                    MsgBox "Erro ao Incluir"
                    grvRegistroGrupo = False
                Else
                    grvRegistroGrupo = True
            End If
        Else
            If RegistroAlterar("UsuGerenciadorGrupo", vReg, cReg, "Id = " & idGrupo) = False Then
                    MsgBox "Erro ao Alterar."
                    grvRegistroGrupo = False
                Else
                    grvRegistroGrupo = True
            End If
    End If
    'Gravar a permissao do Grupo
    If RegistroExcluir("UsuGerenciadorAcessos", "GrupoId=" & idGrupo) = False Then
        MsgBox "Erro ao excluir permissões antigas.", vbInformation, "Aviso"
        Exit Function
    End If
    
    With trvGrupos.Nodes
        For i = 1 To .Count
            
          
            If .item(i).Children <> 0 Then
                    For X = i + 1 To .item(i).Children + i
                        Formulario = .item(i).Key
                        'Debug.Print .Item(x).Text
                        Permissao = Permissao & IIf(.item(X).Checked = True, 1, 0)
                    Next
                    cReg = 0
                    vReg(cReg) = Array("GrupoID", idGrupo, "N"): cReg = cReg + 1
                    vReg(cReg) = Array("Formulario", Formulario, "S"): cReg = cReg + 1
                    vReg(cReg) = Array("Permissao", Permissao, "S") ': cReg = cReg + 1
                    RegistroIncluir "UsuGerenciadorAcessos", vReg, cReg
                    'Debug.Print Formulario
                Else
                    Permissao = ""
                    
                    'Debug.Print .Item(i).Child.Text
                    'Debug.Print .Item(i).Child.Checked
          
            End If
            
        Next
    End With
    
    
    
    
    
End Function
Private Function grvRegistroUsu() As Boolean
    Dim vReg(1000)   As Variant
    Dim cReg         As Integer 'Contador de Registros
    Dim i            As Integer
    Dim sMenu        As String
    
    If Trim(cboGrupo.Text) = "" Then
        MsgBox "Selecione um grupo.", vbInformation, "Aviso"
        grvRegistroUsu = False
        Exit Function
    End If
    
    cReg = 0
    vReg(cReg) = Array("Usu_Nome", cboUsuNome.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Usu_Login", txtUsuLogin.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Usu_Grupo", Left(cboGrupo.Text, 4), "N"): cReg = cReg + 1
    
    If chkSenhaPadrao.Value = 1 Then
        vReg(cReg) = Array("Usu_Senha", "123", "S"): cReg = cReg + 1
    End If
    
    vReg(cReg) = Array("Usu_TrocaSenha", chkTrocarSenha.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("Usu_SenhaNuncaExpira", chkSenhaNuncaExpira.Value, "N"): cReg = cReg + 1
    vReg(cReg) = Array("Usu_SuperUsuario", chkSuperUsuario.Value, "N"): cReg = cReg + 1
    
    
    'Grava as preferencias de menu respeitando a ordem da listagem em tela
    sMenu = ""
    For i = 0 To lstUsuMenu.ListCount - 1
        If lstUsuMenu.Selected(i) = True Then
                sMenu = sMenu & "1"
            Else
                sMenu = sMenu & "0"
        End If
    Next

    vReg(cReg) = Array("Usu_Menus", sMenu, "S"): cReg = cReg + 1
    
    cReg = cReg - 1
    If IdUsu = 0 Then
            If RegistroIncluir(strTabela, vReg, cReg) = 0 Then
                    MsgBox "Erro ao Incluir"
                    grvRegistroUsu = False
                Else
                    grvRegistroUsu = True
            End If
        Else
            If RegistroAlterar(strTabela, vReg, cReg, "Id = " & IdUsu) = False Then
                    MsgBox "Erro ao Alterar."
                    grvRegistroUsu = False
                Else
                    grvRegistroUsu = True
                
            End If
    End If
End Function



