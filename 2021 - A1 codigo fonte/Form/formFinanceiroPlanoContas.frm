VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFinanceiroPlanoContas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financeiro - Plano de Contas"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6600
   Begin VB.Frame Frame2 
      Height          =   1755
      Left            =   60
      TabIndex        =   3
      Top             =   4200
      Width           =   6435
      Begin VB.CheckBox chkTotalizador 
         Caption         =   "Totalizador"
         Height          =   195
         Left            =   3300
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cboCD 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1260
         Width           =   1755
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1140
         MaxLength       =   250
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   900
         Width           =   3915
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   540
         Width           =   3915
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1140
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Cred./Deb.:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ID:"
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   180
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3675
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   6435
      Begin MSFlexGridLib.MSFlexGrid msfgPC 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   5
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "^ID|<Codigo                |<Descrição                                               |^CD  |^Totalizador "
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
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
         Left            =   4920
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
               Picture         =   "formFinanceiroPlanoContas.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroPlanoContas.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroPlanoContas.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroPlanoContas.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroPlanoContas.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroPlanoContas.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroPlanoContas.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroPlanoContas.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroPlanoContas.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroPlanoContas.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroPlanoContas.frx":54D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFinanceiroPlanoContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IdReg     As Integer
Dim sTabela   As String

Private Sub HDF(op As Boolean)
    HDForm Me, op
    txtID.Enabled = IIf(op = True, False, True)
    msfgPC.Enabled = IIf(op = True, False, True)
End Sub

Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    
    IdReg = 0
    HDF True
    HDMenu Me, False
    LF
    
End Sub
Private Sub Alterar()
    
    
    If IdReg = 0 Then
        MsgBox "Selecione um registro!", vbInformation, App.EXEName
        Exit Sub
    End If
    HDF True
    HDMenu Me, False
    'LimpaFormulario Me
    
End Sub
Private Sub LF()
    txtID.Text = ""
    txtCodigo.Text = ""
    txtDescricao.Text = ""
    cboCD.Clear
End Sub

Private Sub ListarPlanos()
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim totalizador As String
    
    sSQL = "SELECT * FROM " & sTabela & " WHERE id_empresa=" & ID_Empresa & " ORDER BY Codigo"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            With msfgPC
                Rst.MoveFirst
                .Rows = 1
                Do Until Rst.EOF
                    If cNull(Rst.Fields("totalizador")) = "1" Then
                            totalizador = "S"
                        Else
                            totalizador = "N"
                        End If
                
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("id")
                    .TextMatrix(.Rows - 1, 1) = cNull(Rst.Fields("Codigo"))
                    .TextMatrix(.Rows - 1, 2) = cNull(Rst.Fields("Descricao"))
                    .TextMatrix(.Rows - 1, 3) = cNull(Rst.Fields("cd"))
                    .TextMatrix(.Rows - 1, 4) = totalizador
                    Rst.MoveNext
                Loop
            End With
    End If
    Rst.Close
End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    sTabela = Mid(Me.Name, 5, Len(Me.Name))
    HDF False
    HDMenu Me, True
    ListarPlanos
    
End Sub
Private Sub cboCD_DropDown()
    With cboCD
        .Clear
        .AddItem "C - Crédito"
        .AddItem "D - Débito"
    End With
End Sub

Private Sub msfgPC_Click()
    With msfgPC
        IdReg = .TextMatrix(.Row, 0)
        txtID.Text = .TextMatrix(.Row, 0)
        txtCodigo.Text = .TextMatrix(.Row, 1)
        txtDescricao.Text = .TextMatrix(.Row, 2)
        cboCD.Clear
        Select Case Trim(.TextMatrix(.Row, 3))
            Case "C"
                cboCD.AddItem "C - Crédito"
            Case "D"
                cboCD.AddItem "D - Débito"
            Case Else
                cboCD.AddItem " "
        End Select
        cboCD.Text = cboCD.List(0)
        
        If Trim(.TextMatrix(.Row, 4)) = "S" Then
                chkTotalizador.Value = 1
            Else
                chkTotalizador.Value = 0
        End If
        
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
        Case "Imprimir"
            Imprimir
            
        Case "Pesquisar"
            IdReg = 0
            PesquisarRegistro
            
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDF False
                ListarPlanos
            End If
            
        
        Case "Cancelar"
            'HDMenu Me, True
            'HDForm Me, False
            'LimpaFormulario Me
            
        Case "Manutenção da Tabela"
            formManutencaoTabelas.IniciarManutencao Me
    End Select
End Sub
Private Sub Excluir()
    If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
            MsgBox "Selecione um Registro!", vbInformation, App.EXEName
            Exit Sub
        Else
            If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                        vbCrLf & _
                        "Código: " & txtCodigo.Text & vbCrLf & _
                        "Descrição.: " & txtDescricao.Text, vbYesNo + vbQuestion) = vbYes Then
                               
                If RegistroExcluir(sTabela, "Id = " & IdReg) = True Then
                    LimpaFormulario Me
                End If
            End If
    End If
End Sub

Private Sub Imprimir()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    sSQL = "SELECT * FROM " & sTabela & " ORDER BY Codigo"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado!", vbInformation, App.EXEName
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Set rptFinanceiroListPlanoContas.DataSource = Rst.DataSource
            rptFinanceiroListPlanoContas.Show 1
            Rst.Close
    End If
End Sub
Private Sub PesquisarRegistro()
    
    If IdReg = 0 Then
        IdReg = formBuscar.IniciarBusca(sTabela, , , , "ORDER BY Codigo")
    End If
    If IdReg = 0 Then
        LF
        Exit Sub
    End If
    
    txtID.Text = PgDadosPlanoContas("ID", CStr(IdReg)).Id
    txtCodigo.Text = PgDadosPlanoContas("ID", CStr(IdReg)).Codigo
    txtDescricao.Text = PgDadosPlanoContas("ID", CStr(IdReg)).Descricao
    cboCD.Clear
    cboCD.AddItem PgDadosPlanoContas("ID", CStr(IdReg)).cd
    cboCD.Text = cboCD.List(0)
    
End Sub
Private Function grvRegistro() As Boolean
    On Error GoTo TrtErroGrv
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    If ValidarDados = False Then
        grvRegistro = False
    End If
    grvRegistro = False
    cReg = 0
    vReg(cReg) = Array("Codigo", Trim(txtCodigo.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Descricao", Trim(txtDescricao.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("CD", Left(Trim(cboCD.Text), 1), "S"): cReg = cReg + 1
    vReg(cReg) = Array("totalizador", chkTotalizador.Value, "N"): cReg = cReg + 1
    cReg = cReg - 1
    If IdReg = 0 Then
            IdReg = RegistroIncluir(sTabela, vReg, cReg)
        Else
            RegistroAlterar sTabela, vReg, cReg, "id = " & IdReg
    End If
    txtID.Text = IdReg
    grvRegistro = True
    Exit Function
TrtErroGrv:
    grvRegistro = False
End Function
Private Function ValidarDados() As Boolean
    If Trim(txtCodigo.Text) = "" Then
        MsgBox "O campo CODIGO não pode ser um valor vazio!", vbInformation, App.EXEName
        ValidarDados = False
        Exit Function
    End If
    If Trim(txtDescricao.Text) = "" Then
        MsgBox "O campo DESCRIÇÃO não pode ser um valor vazio!", vbInformation, App.EXEName
        ValidarDados = False
        Exit Function
    End If
    If Trim(cboCD.Text) = "" Then
        MsgBox "O campo CEDITO/DEBITO não pode ser um valor vazio!", vbInformation, App.EXEName
        ValidarDados = False
        Exit Function
    End If
End Function
