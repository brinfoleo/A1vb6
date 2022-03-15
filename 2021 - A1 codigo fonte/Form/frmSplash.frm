VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5220
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblVersao 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000.00.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   6540
      TabIndex        =   4
      Top             =   60
      Width           =   1920
   End
   Begin VB.Label lblMensagem 
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciando"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   60
      TabIndex        =   2
      Top             =   4920
      Width           =   6780
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NF-e 4.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   5340
      TabIndex        =   3
      Top             =   4740
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema para Gerenciamento de Empresas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1200
      Left            =   15
      TabIndex        =   1
      Top             =   1980
      Width           =   6210
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "© A1 – Gerenciamento de empresas 2010 - Todos os direitos reservados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   3330
      Width           =   8040
   End
   Begin VB.Image Image1 
      Height          =   5280
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   -45
      Width           =   8655
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Sub CarregarFormulario(msg As String)
    DoEvents
    lblVersao.Caption = App.Major & "." & App.Minor & "." & App.Revision
    'lblRevisao.Caption = "Revisão: " & cVersao
    lblMensagem.Caption = msg
    Sleep 250
    Me.Show
End Sub
Public Sub FecharFormulario()
    Unload Me
End Sub

