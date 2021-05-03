VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formRHFuncionarioFolhadePagamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RH - Folha de Pagamento"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   10830
   Begin VB.Frame Frame2 
      Caption         =   "Movimento:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   10635
      Begin VB.ComboBox cboCD 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   2355
      End
      Begin VB.TextBox txtValor 
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1380
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   180
         Width           =   7935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Credito/Debito:"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor:"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   420
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Extrato:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4035
      Left            =   60
      TabIndex        =   4
      Top             =   1260
      Width           =   10635
      Begin MSFlexGridLib.MSFlexGrid msfgMov 
         Height          =   3195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5636
         _Version        =   393216
         Cols            =   5
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"formRHFuncionarioFolhadePagamento.frx":0000
      End
      Begin VB.Label lblvSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
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
         Height          =   315
         Left            =   8520
         TabIndex        =   19
         Top             =   3540
         Width           =   1995
      End
      Begin VB.Label lblvDeb 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5880
         TabIndex        =   18
         Top             =   3600
         Width           =   1995
      End
      Begin VB.Label lblvCred 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3180
         TabIndex        =   17
         Top             =   3600
         Width           =   1995
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo:"
         Height          =   255
         Left            =   7740
         TabIndex        =   16
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Débito:"
         Height          =   195
         Left            =   5160
         TabIndex        =   15
         Top             =   3660
         Width           =   675
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Crédito:"
         Height          =   195
         Left            =   2280
         TabIndex        =   14
         Top             =   3660
         Width           =   855
      End
   End
   Begin VB.TextBox txtMesAno 
      Height          =   315
      Left            =   7620
      MaxLength       =   7
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   780
      Width           =   975
   End
   Begin VB.ComboBox cboFuncionario 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   5895
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10830
      _ExtentX        =   19103
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
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Enviar para Financeiro"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
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
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":00C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":0518
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":0832
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":10C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":2316
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":2BF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":3482
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":3D14
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":4F66
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":5280
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":559A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhadePagamento.frx":5991
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Mes/Ano:"
      Height          =   195
      Left            =   6720
      TabIndex        =   2
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Funcionário:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   975
   End
End
Attribute VB_Name = "formRHFuncionarioFolhadePagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim idFunc  As Integer
Dim IdReg   As Integer


Private Sub Excluir()
     If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If
    If MsgBox("Deseja realmente EXCLUIR esse registro?", vbYesNo + vbQuestion, App.EXEName) = vbYes Then
    'RegistroAlterar "RHFuncionarioFolhadePagamento", vReg, cReg, "id=" & idReg
        If RegistroExcluir("RHFuncionarioFolhadePagamento", "id=" & IdReg) = True Then
                LimpForm
                Atualiza
            Else
                MsgBox "Erro ao excluir registro!", vbCritical, App.EXEName
                IdReg = 0
        End If
        
    End If
End Sub

Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    If idFunc = 0 Then
        MsgBox "Selecione um Funcionário!", vbInformation, App.EXEName
        Exit Sub
    End If
    LimpForm
    HDMenu Me, False
    HDForm Me, True
    
    cboFuncionario.Enabled = False
    txtMesAno.Enabled = False
    msfgMov.Enabled = False
End Sub
Private Sub LimpForm()
    txtDescricao.Text = ""
    txtValor.Text = ""
    cboCD.Clear
    
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If idFunc = 0 Or IdReg = 0 Then
        MsgBox "Selecione um Funcionário ou Registro!", vbInformation, App.EXEName
        Exit Sub
    End If
    'LimpForm
    HDMenu Me, False
    HDForm Me, True
    
    cboFuncionario.Enabled = False
    txtMesAno.Enabled = False
    msfgMov.Enabled = False
End Sub
Private Sub cboCD_DropDown()
    With cboCD
        .Clear
        .AddItem "C - Crédito"
        .AddItem "D - Débito"
    End With
End Sub

Private Sub cboFuncionario_Click()
    If Trim(cboFuncionario.Text) = "" Then
        idFunc = 0
        Exit Sub
    End If
    idFunc = Left(Trim(cboFuncionario.Text), 4)
    Atualiza
End Sub

Private Sub cboFuncionario_DropDown()
    Dim Rst As Recordset
    cboFuncionario.Clear
    Set Rst = RegistroBuscar("SELECT * FROM RHFuncionarioCadastro WHERE ID_Empresa = " & ID_Empresa & " ORDER BY xNome")
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboFuncionario.AddItem Left(String(4, "0"), 4 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & " - " & Rst.Fields("xNome")
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
    idFunc = 0
    Cancelar
    txtMesAno.Text = Format(Date, "MM/YYYY")
End Sub

Private Sub msfgMov_Click()
    With msfgMov
        If Trim(.TextMatrix(.Row, 0)) = "" Then Exit Sub
        IdReg = .TextMatrix(.Row, 0)
        txtDescricao.Text = .TextMatrix(.Row, 2)
        txtValor.Text = .TextMatrix(.Row, 3)
        cboCD.Clear
        If .TextMatrix(.Row, 4) = "C" Then
                cboCD.AddItem "C - Crédito"
                cboCD.Text = cboCD.List(0)
            ElseIf .TextMatrix(.Row, 4) = "D" Then
                cboCD.AddItem "D - Débito"
                cboCD.Text = cboCD.List(0)
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
        Case "Atualizar"
            Atualiza
        
        Case "Enviar para Financeiro"
            MovFinanceiro
        Case "Salvar"
            If grvRegistro = True Then
                'HDMenu Me, True
                'HDForm Me, False
                Cancelar
                Atualiza
            End If
            
        
        Case "Cancelar"
            Cancelar
        
        Case "Manutenção da Tabela"
            'formManutencaoTabelas.IniciarManutencao Me
            MontarBaseDeDados
    End Select
End Sub
Private Sub MovFinanceiro()
    Dim nFat    As String
    Dim vSaldo  As String
    Dim MesAno  As String
    nFat = "000"
    vSaldo = ChkVal(lblvSaldo.Caption, 0, cDecMoeda)
    MesAno = Trim(txtMesAno.Text)
    
    
    Call MovimentarContasPagarReceber("P", Date, nFat, vSaldo, "RHFuncionarioCadastro", idFunc, PgDadosRhFuncionario(idFunc).Nome, _
                                    PgDadosRhFuncionario(idFunc).CPF, PgDadosConfig.RHConta, PgDadosConfig.RHCentroCustos, PgDadosConfig.RHDocumento, PgDadosConfig.RHPlanoContas, "", "", Date, nFat & "-1/1", "0", _
                                    "0", "0", "0", "0", "0", "0", vSaldo, "Ref.:" & MesAno)
                                    
    MsgBox "Lançamento gerado com sucesso!", vbInformation, App.EXEName
End Sub
Private Sub Cancelar()
    HDMenu Me, True
    HDForm Me, False
    LimpForm
    cboFuncionario.Enabled = True
    txtMesAno.Enabled = True
    msfgMov.Enabled = True
    IdReg = 0
    
End Sub
Private Sub MontarBaseDeDados()
    Dim vReg(10)      As Variant
    Dim cReg            As Integer
    
    
    cReg = 0
 
    
    vReg(cReg) = Array("IDFunc", "10", "N"): cReg = cReg + 1
    vReg(cReg) = Array("MesAno", "10", "S"): cReg = cReg + 1
    vReg(cReg) = Array("Doc", "50", "S"): cReg = cReg + 1
    vReg(cReg) = Array("descricao", "255", "S"): cReg = cReg + 1
    vReg(cReg) = Array("Valor", "20", "S"): cReg = cReg + 1
    vReg(cReg) = Array("Cd", "1", "S"): cReg = cReg + 1
    cReg = cReg - 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vReg, cReg
    
End Sub


Private Sub txtMesAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 47 Then
        If InStr(txtMesAno.Text, "/") = 0 Then
                Exit Sub
            Else
                KeyAscii = 0
                Exit Sub
        End If
    End If
    KeyAscii = SoNumeros(KeyAscii)
    
End Sub
Private Sub Atualiza()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    Dim vC      As String
    Dim vD      As String
    
    
    If idFunc = 0 Then Exit Sub
    msfgMov.Rows = 1
    LancarSalarioBase
    sSQL = "SELECT * FROM RHFuncionarioFolhadePagamento WHERE IdFunc=" & idFunc & " AND MesAno='" & Trim(txtMesAno.Text) & "' ORDER BY DOC"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                With msfgMov
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("ID")
                    .TextMatrix(.Rows - 1, 1) = cNull(Rst.Fields("Doc"))
                    .TextMatrix(.Rows - 1, 2) = cNull(Rst.Fields("Descricao"))
                    .TextMatrix(.Rows - 1, 3) = cNull(Rst.Fields("Valor"))
                    .TextMatrix(.Rows - 1, 4) = cNull(Rst.Fields("cd"))
                    Select Case UCase(cNull(Rst.Fields("cd")))
                        Case "C"
                            vC = Val(ChkVal(vC, 0, cDecMoeda)) + Val(ChkVal(cNull(Rst.Fields("Valor")), 0, cDecMoeda))
                        Case "D"
                            vD = Val(ChkVal(vD, 0, cDecMoeda)) + Val(ChkVal(cNull(Rst.Fields("Valor")), 0, cDecMoeda))
                    End Select
                    Rst.MoveNext
                End With
            Loop
    End If
    Rst.Close
    
    lblvCred.Caption = ConvMoeda(vC)
    lblvDeb.Caption = ConvMoeda(vD)
    lblvSaldo.Caption = ConvMoeda(Val(ChkVal(vC, 0, cDecMoeda)) - Val(ChkVal(vD, 0, cDecMoeda)))
End Sub
Private Function grvRegistro() As Boolean
    Dim vReg(10)    As Variant
    Dim cReg        As Integer
    
    Dim sDoc As String
    
    sDoc = Format(Date, "YYMMDD") & Format(Time, "HHMMSS")
    
    cReg = 0
    vReg(cReg) = Array("IDFunc", idFunc, "N"): cReg = cReg + 1
    vReg(cReg) = Array("MesAno", txtMesAno.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("DOC", sDoc, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Descricao", Trim(txtDescricao.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("Valor", Trim(txtValor.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("CD", Left(cboCD.Text, 1), "S"): cReg = cReg + 1
    cReg = cReg - 1
    If IdReg = 0 Then
            IdReg = RegistroIncluir("RHFuncionarioFolhadePagamento", vReg, cReg)
        Else
            RegistroAlterar "RHFuncionarioFolhadePagamento", vReg, cReg, "id=" & IdReg
    End If
    IdReg = 0
    LimpForm
    grvRegistro = True
End Function
Private Sub LancarSalarioBase()
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    Dim vReg(10)    As Variant
    Dim cReg        As Integer
    
    If idFunc = 0 Then Exit Sub
    
    
    
    msfgMov.Rows = 1
   
    sSQL = "SELECT * FROM RHFuncionarioFolhadePagamento WHERE IdFunc=" & idFunc & " AND MesAno='" & Trim(txtMesAno.Text) & "' AND DOC ='000001'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            cReg = 0
            vReg(cReg) = Array("idFunc", idFunc, "S"): cReg = cReg + 1
            vReg(cReg) = Array("MesAno", Trim(txtMesAno.Text), "S"): cReg = cReg + 1
            vReg(cReg) = Array("Doc", "000001", "S"): cReg = cReg + 1
            vReg(cReg) = Array("Descricao", "Salário Base", "S"): cReg = cReg + 1
            vReg(cReg) = Array("Valor", ChkVal(PgDadosRhFuncionario(idFunc).Salario, 0, cDecMoeda), "S"): cReg = cReg + 1
            vReg(cReg) = Array("CD", "C", "S"): cReg = cReg + 1
            cReg = cReg - 1
            RegistroIncluir "RHFuncionarioFolhadePagamento", vReg, cReg
            
    End If
    Rst.Close
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtValor.Text, KeyAscii, cDecMoeda)
End Sub
