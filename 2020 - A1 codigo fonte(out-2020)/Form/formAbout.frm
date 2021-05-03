VERSION 5.00
Begin VB.Form formAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre o Sistema"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9315
   Begin VB.Frame Frame2 
      Caption         =   "Dados do Aplicativo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   4860
      TabIndex        =   7
      Top             =   4080
      Width           =   4395
      Begin VB.Label lblSerial 
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   4155
      End
      Begin VB.Label lblVersao 
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   4155
      End
   End
   Begin VB.TextBox Text2 
      Height          =   1875
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "formAbout.frx":0000
      Top             =   360
      Width           =   9195
   End
   Begin VB.Frame Frame1 
      Caption         =   "Base de Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   60
      TabIndex        =   1
      Top             =   4080
      Width           =   4695
      Begin VB.Label lblNomeBD 
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   4395
      End
      Begin VB.Label lblPorta 
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   510
         Width           =   4395
      End
      Begin VB.Label lblIP 
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   4395
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1395
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "formAbout.frx":11EA
      Top             =   2580
      Width           =   9195
   End
   Begin VB.Label Label2 
      Caption         =   "Historico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2340
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Licença de Uso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "formAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




'Private Sub Command1_Click()
'    On Error Resume Next
'    Dim sSQL As String
'    Dim Rst As Recordset
'    Dim Mail    As String
'
'    sSQL = "SELECT * FROM Clientes WHERE EmailNFe IS NOT NULL"
'    Set Rst = RegistroBuscar(sSQL)
'    If Rst.BOF And Rst.EOF Then
'
'        Else
'            Rst.MoveFirst
'            Do Until Rst.EOF
'                Mail = Rst.Fields("EmailNFe")
'                If Trim(Mail) <> "" Then
'                    If InStr(Mail, "@") = 0 Then
'                        Mail = Mid(Mail, 1, InStr(Mail, ".") - 1) & "@" & Mid(Mail, InStr(Mail, ".") + 1, Len(Mail))
'                        'Rst.Edit
'                        'Rst.Fields("EmailNFe") = LCase(Mail)
'                        BD.Execute "UPDATE Clientes SET EmailNFe = '" & LCase(Mail) & "', eMail = '" & LCase(Mail) & "' WHERE id=" & Rst.Fields("ID")
'                        'MsgBox Mail
'                    End If
'                End If
'                Rst.MoveNext
'            Loop
'    End If
'
'End Sub

Private Sub Form_Load()
    lblSerial.Caption = Numero_Serial
    lblIP.Caption = "IP: " & srv_IP
    lblPorta.Caption = "Porta: " & srv_Porta
    lblVersao = "Versão: " & sVersao & " rev." & cVersao
    lblNomeBD.Caption = "Database: " & nmDatabase
End Sub
