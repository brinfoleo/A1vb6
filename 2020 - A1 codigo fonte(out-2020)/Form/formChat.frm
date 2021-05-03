VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formChat 
   Caption         =   "Chat"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   6930
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1620
      Top             =   2700
   End
   Begin MSFlexGridLib.MSFlexGrid msfgChat 
      Height          =   4335
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   3
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "^id|^Usuario                      |^IP            "
   End
End
Attribute VB_Name = "formChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UsuDestino  As String
Dim usuOrigem   As String
Dim ipdestino   As String

Private Sub Form_Load()
    On Error Resume Next
    UsuariosConectados
    Atualizar
End Sub
Private Sub Atualizar()
    On Error Resume Next
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    msfgChat.Rows = 1
    apagarUsuariosOffline
    sSQL = "SELECT DISTINCT Usuario FROM ConexaoGerenciador WHERE ID_Empresa = " & ID_Empresa & " order BY IP"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                With msfgChat
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = Rst.fields("id")
                '.TextMatrix(.Rows - 1, 1) = Rst.fields("Nome")
                .TextMatrix(.Rows - 1, 2) = Rst.fields("IP")
                .TextMatrix(.Rows - 1, 1) = Trim(Mid(Rst.fields("Usuario"), 4, Len(Rst.fields("Usuario"))))
                '.TextMatrix(.Rows - 1, 4) = Rst.Fields("Data")
                '.TextMatrix(.Rows - 1, 5) = Rst.Fields("Hora")
                '.TextMatrix(.Rows - 1, 6) = Rst.Fields("Status")
                End With
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub
Private Sub msfgChat_Click()
    UsuDestino = msfgChat.TextMatrix(msfgChat.Row, 1)
    ipdestino = msfgChat.TextMatrix(msfgChat.Row, 2)
End Sub
Private Sub msfgChat_DblClick()
    Dim nForm As New formChatConversa
    nForm.IniciarChat (ipdestino)
End Sub
Private Sub Timer1_Timer()
    Atualizar
End Sub
