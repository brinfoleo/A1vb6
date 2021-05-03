VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form formChatConversa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   4215
   Begin MSWinsockLib.Winsock Ws 
      Left            =   3720
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMsg 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5520
      Width           =   3975
   End
   Begin VB.TextBox txtMsgs 
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "formChatConversa.frx":0000
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "formChatConversa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Sit=0 - Chat solicitado por terceiros
'Sit=1 - Chat solicitado pelo usuario
Dim Sit As Integer

Private Sub Form_Load()
    'Ws.LocalPort = 1720
    txtMsgs.Text = ""
    txtMsg.Text = ""
    
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnviarMensagem (">>: " & txtMsg.Text)
        txtMsg.Text = ""
    End If
End Sub
Private Sub EnviarMensagem(msg As String)
    On Error GoTo MsgTrat
    Ws.SendData msg
    txtMsgs.Text = txtMsgs.Text & _
                   "(" & Time & ")" & _
                   " " & _
                   msg & _
                   vbCrLf
    Exit Sub
MsgTrat:
 txtMsgs.Text = txtMsgs.Text & _
                   "(" & Time & ")" & _
                   " " & _
                   "Erro: " & Err.Number & " - Erro no envio da mensagem." & _
                   vbCrLf
    
End Sub
Public Sub IniciarChat(ipDest As String)
    'Sit=0 - Chat solicitado por terceiros
    'Sit=1 - Chat solicitado pelo usuario
    Sit = 1
    Me.Caption = "Chat - " & ipDest
    txtMsg.Enabled = False
    Me.Show
    Ws.Connect ipDest, 1270
End Sub
Public Sub ResponderChat(ipDest As String)
    'Sit=0 - Chat solicitado por terceiros
    'Sit=1 - Chat solicitado pelo usuario
    Sit = 0
    Me.Caption = "Chat - " & ipDest
    Me.Show
End Sub

Private Sub txtMsgs_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Ws_Connect()
    txtMsg.Enabled = True
End Sub

Private Sub Ws_ConnectionRequest(ByVal requestID As Long)
    Ws.Accept requestID
End Sub

Private Sub Ws_DataArrival(ByVal bytesTotal As Long)
    Dim MsgRecebida As String
    Ws.GetData MsgRecebida
    EnviarMensagem ("<<: " & MsgRecebida)
End Sub

Private Sub Ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    EnviarMensagem ("ERRO DE COMUNICAÇÃO COM DESTINO...")
End Sub
