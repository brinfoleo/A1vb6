VERSION 5.00
Begin VB.Form formConexao 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A1 - Conexão"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   4095
      Begin VB.TextBox txtNomeBD 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1200
         Width           =   2475
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   1440
         MaxLength       =   60
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   2475
      End
      Begin VB.TextBox txtPorta 
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   780
         Width           =   2475
      End
      Begin VB.TextBox txtdbUsu 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1620
         Width           =   2475
      End
      Begin VB.TextBox txtdbSenha 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2040
         Width           =   2475
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Banco de Dados:"
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Endereço de IP:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Porta:"
         Height          =   195
         Left            =   420
         TabIndex        =   9
         Top             =   825
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   1695
         Width           =   915
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Senha:"
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   2100
         Width           =   1095
      End
   End
   Begin VB.CommandButton botGravar 
      Caption         =   "&Gravar"
      Height          =   555
      Left            =   4140
      TabIndex        =   6
      Top             =   2820
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   120
      Picture         =   "formConexao.frx":0000
      Stretch         =   -1  'True
      Top             =   300
      Width           =   1725
   End
End
Attribute VB_Name = "formConexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub botGravar_Click()
    Dim sTexto As String
    Dim caminho As String

    
    caminho = App.Path & "\" & App.EXEName & ".cfg"
    
    If Dir(caminho) <> "" Then
        Kill caminho
    End If
    'sTexto = txtIP.Text & "|" & txtPorta.Text & "|" & txtdbUsu.Text & "|" & txtdbSenha.Text
    sTexto = "# Registro criado/modificado em: " & Now
    grvFile caminho, sTexto
    grvFile caminho, "IP=" & Trim(txtIP.Text)
    grvFile caminho, "PORT=" & Trim(txtPorta.Text)
    grvFile caminho, "nmDatabase=" & Trim(txtNomeBD.Text)
    grvFile caminho, "USU=" & Trim(txtdbUsu.Text)
    grvFile caminho, "Senha=" & Trim(txtdbSenha.Text)
    

    MsgBox "Registro armazenado com sucesso!" & vbCrLf & "Favor reiniciar o sistema!", vbInformation, "Aviso"
    Unload Me
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    'optConexao_Click (0)
    pegarDadosGravados
End Sub
Private Sub pegarDadosGravados()
    On Error Resume Next
    txtNomeBD.Text = nmDatabase
    txtIP.Text = srv_IP
    txtPorta.Text = srv_Porta
    txtdbUsu.Text = dbUsu
    txtdbSenha.Text = dbSenha
End Sub

'Private Sub optConexao_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            txtIP.Enabled = False
'            txtPorta.Enabled = False
'           txtIP.Text = "127.0.0.1"
'             txtNomeBD.Enabled = False
'            txtPorta.Text = "3306"
'            txtNomeBD.Text = "a1_padrao"
'        Case 1
'            txtIP.Enabled = True
'            txtPorta.Enabled = True
'            txtNomeBD.Enabled = True
'    End Select
'End Sub
Private Sub grvReg(Dados As String)
    On Error GoTo TrtErro
    
    'define o ObjPreview filesystem e demais variaveis
    Dim fso As New FileSystemObject
    Dim Arquivo As File
    Dim arquivoLog As TextStream
    Dim msg As String
    Dim caminho As String

    
    caminho = App.Path & "\" & App.EXEName & ".cfg"
    
    If Dir(caminho) <> "" Then
        Kill caminho
    End If
    'se o arquivo não existir então cria
    If fso.FileExists(caminho) Then
            Set Arquivo = fso.GetFile(caminho)
        Else
            Set arquivoLog = fso.CreateTextFile(caminho)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(caminho)
    End If
    'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
    
    'monta informações para gerar a linha da mensagem
    msg = Dados

    'inclui linhas no arquivo texto
    arquivoLog.WriteLine msg
    
    'escreve uma linha em branco no arquivo - se voce quiser
    'arquivoLog.WriteBlankLines (1)
    'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing
    Exit Sub
TrtErro:
        MsgBox "Erro ao gerar registro da NFe em Texto .                                   " & _
           vbCrLf & vbCrLf & _
           "Erro n.: " & Err.Number & _
           vbCrLf & vbCrLf & _
           "Descrição: " & Err.Description & _
           vbCrLf

End Sub



'Private Sub txtIP_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 8 Then Exit Sub
'    If KeyAscii = 46 Then Exit Sub
'    If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
'End Sub

Private Sub txtPorta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub
