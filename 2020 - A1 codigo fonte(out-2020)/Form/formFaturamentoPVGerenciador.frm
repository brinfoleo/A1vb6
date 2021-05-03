VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formFaturamentoPVGerenciador 
   Caption         =   "A1 - Gerenciador de Pre-Vendas"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   13935
   Begin VB.Frame frmTitulosAtrazados 
      Caption         =   "Titulos em atrazo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   8400
      TabIndex        =   2
      Top             =   3420
      Width           =   4695
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2355
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4154
         _Version        =   393216
         Cols            =   4
         FormatString    =   "^ID |>Titulo                       |>Valor             |^Vencimento    "
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   13215
      Begin MSFlexGridLib.MSFlexGrid msfgPV 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12915
         _ExtentX        =   22781
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   8
         FormatString    =   $"formFaturamentoPVGerenciador.frx":0000
      End
   End
End
Attribute VB_Name = "formFaturamentoPVGerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

