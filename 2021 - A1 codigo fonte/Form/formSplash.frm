VERSION 5.00
Begin VB.Form formSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9675
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Homologado NF-e 3.10"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   7920
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblRevisao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6300
      TabIndex        =   2
      Top             =   600
      Width           =   3075
   End
   Begin VB.Label lblMensagem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2100
      TabIndex        =   1
      Top             =   3240
      Width           =   7455
   End
   Begin VB.Label lblVersao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6300
      TabIndex        =   0
      Top             =   300
      Width           =   3075
   End
   Begin VB.Image Image1 
      Height          =   4380
      Left            =   0
      Picture         =   "formSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   -60
      Width           =   9690
   End
End
Attribute VB_Name = "formSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub CarregarFormulario(msg As String)
    DoEvents
    lblVersao.Caption = "Versão: " & sVersao
    lblRevisao.Caption = "Revisão: " & cVersao
    lblMensagem.Caption = msg
    
    Me.Show
End Sub
Public Sub FecharFormulario()
    Unload Me
End Sub
