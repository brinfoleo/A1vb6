VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form formImpressoraSelecionar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecionar Impressora"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   3360
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selecionar"
      Height          =   555
      Left            =   4980
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   7035
   End
   Begin VB.Label Label1 
      Caption         =   "Selecione a Impressora:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1995
   End
End
Attribute VB_Name = "formImpressoraSelecionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function SelecionarImpressora() As Boolean
    On Error GoTo TrataErro
    SelecionarImpressora = False
    cd.CancelError = True
    cd.ShowPrinter
    SelecionarImpressora = True
    Unload Me
    Exit Function
TrataErro:
    SelecionarImpressora = False
    Unload Me
End Function
