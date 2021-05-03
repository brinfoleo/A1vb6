VERSION 5.00
Begin VB.Form formLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "A1 - Login"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Height          =   555
      Left            =   3780
      MaskColor       =   &H00404040&
      Picture         =   "formLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4740
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Cancel          =   -1  'True
      Height          =   555
      Left            =   5520
      MaskColor       =   &H00404040&
      Picture         =   "formLogin.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4740
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4035
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   6915
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   2460
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   2160
         Width           =   2715
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   2460
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1620
         Width           =   2715
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Top             =   2220
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   1680
         Width           =   915
      End
   End
End
Attribute VB_Name = "formLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sLogin As Boolean



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Command1_Click()
    sLogin = AutenticaUsuario(Text1.Text, Text2.Text)
    If sLogin = True Then
            Unload Me
        Else
            MsgBox "LOGIN ou SENHA invalido.", vbInformation, "Login"
    End If
End Sub

Private Sub Command2_Click()
    sLogin = False
    Unload Me
End Sub
Public Function EfetuarLogin() As Boolean
    Me.Show 1
    EfetuarLogin = sLogin
End Function


Private Sub Form_Load()
    LimpaFormulario Me
    Me.Caption = "A1 - Login " & "[ versão: " & sVersao & " rev." & cVersao & "]"
End Sub

Private Function AutenticaUsuario(Nome As String, senha As String) As Boolean
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
End Function



Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Command1_Click
End Sub
