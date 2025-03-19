VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A1 - Sistema de Gerenciamento Comercial"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8190
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdSair 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5520
      TabIndex        =   7
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton cmdEntrar 
      Appearance      =   0  'Flat
      Caption         =   "&Entrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2820
      TabIndex        =   6
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dados"
      Height          =   2655
      Left            =   1920
      TabIndex        =   0
      Top             =   45
      Width           =   6180
      Begin VB.ComboBox cboEmpresa 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   540
         Width           =   5835
      End
      Begin VB.TextBox txtSenha 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   225
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1935
         Width           =   5745
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   225
         TabIndex        =   1
         Top             =   1230
         Width           =   5745
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Empresa:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblVersao 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Versão: 2016.0.52"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3840
         TabIndex        =   5
         Top             =   2355
         Width           =   2130
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   4
         Top             =   1665
         Width           =   1455
      End
      Begin VB.Label lnlNome 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   3795
      Left            =   0
      Picture         =   "frmlogin.frx":0442
      Stretch         =   -1  'True
      Top             =   -90
      Width           =   1905
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sLogin As Boolean

Private Sub cboEmpresa_Click()
    If Len(Trim(cboEmpresa.Text)) = 0 Then Exit Sub
    
    ID_Empresa = Left(cboEmpresa.Text, 2)
End Sub



Private Sub Command1_Click()
Shell App.Path & "\bbCobranca\BBConnectAPI.exe list", vbNormalFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub
Private Sub cmdSair_Click()
    sLogin = False
    Unload Me
End Sub

Private Sub txtnome_GotFocus()
    txtNome.BackColor = &HC0FFFF
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSenha.SetFocus
End Sub

Private Sub txtnome_LostFocus()
    txtNome.BackColor = &HFFFFFF
End Sub

Private Sub txtSenha_GotFocus()
    txtSenha.BackColor = &HC0FFFF
End Sub

Private Sub txtSenha_LostFocus()
    txtSenha.BackColor = &HFFFFFF
End Sub

'**********************************
'**********************************

Private Sub cmdEntrar_Click()
    sLogin = AutenticaUsuario(txtNome.Text, txtSenha.Text)
    If sLogin = True Then
            Unload Me
        Else
            MsgBox "LOGIN ou SENHA invalido.", vbInformation, "Login"
    End If
End Sub


Public Function EfetuarLogin() As Boolean
    Me.Show 1
    EfetuarLogin = sLogin
End Function


Private Sub Form_Load()
    LimpaFormulario Me
    Me.Caption = "A1 - Login " '& "[ versão: " & sVersao & " rev." & cVersao & "]"
    lblVersao.Caption = "Versão: " & sVersao & " rev." & cVersao
    listarEmpresas
End Sub

Private Function AutenticaUsuario(Nome As String, senha As String) As Boolean
    On Error GoTo trtErrAutUsu
    Dim sSQL    As String
    Dim Rst     As Recordset
    sSQL = "SELECT * FROM UsuGerenciador WHERE ID_Empresa = " & ID_Empresa & " AND Usu_Login = '" & Nome & "' AND Usu_Senha = '" & senha & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            AutenticaUsuario = False
        Else
            AutenticaUsuario = True
            Rst.MoveFirst
            ID_Usuario = Rst.Fields("id")
            If Rst.Fields("Usu_TrocaSenha") = 1 Then
                AutenticaUsuario = False
                Unload Me
                formUsuTrocarSenha.Show 1
            End If
    End If
    Rst.Close
    Exit Function
trtErrAutUsu:
    RegLogDataBase 0, "", Err.Number, Err.Description
End Function

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdEntrar_Click
End Sub

Private Sub listarEmpresas()
    cboEmpresa.Clear
'    cboEmpresa.Text = ""
    Dim Rst As Recordset
    Dim sSQL As String
    Dim Item As String
    cboEmpresa.Clear
    sSQL = "SELECT * FROM empresas"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            ID_Empresa = 0
            Exit Sub
        Else
            Rst.MoveLast
            'Caso tenha apenas uma empresa cadastrada
            If Rst.RecordCount = 1 Then
                    cboEmpresa.Enabled = False
                    Rst.MoveFirst
                    ID_Empresa = Rst.Fields("id")
                    Item = ZE(cNull(Rst.Fields("id")), 2) & " - " & Rst.Fields("CNPJ") & " - " & Rst.Fields("nome")
                    cboEmpresa.AddItem (Item)
                    cboEmpresa.Text = cboEmpresa.List(0)
                    Rst.Close
                    Exit Sub
                Else
                    cboEmpresa.Enabled = True
                    ID_Empresa = 0
                    Rst.MoveFirst
                    
                    Do Until Rst.EOF
                        Item = ZE(cNull(Rst.Fields("id")), 2) & " - " & Rst.Fields("CNPJ") & " - " & Rst.Fields("nome")
                        cboEmpresa.AddItem (Item)
                        
                        
                        Rst.MoveNext
                    Loop
                
            End If
            
     End If
    

End Sub
