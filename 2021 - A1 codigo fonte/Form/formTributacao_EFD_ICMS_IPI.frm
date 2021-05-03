VERSION 5.00
Begin VB.Form formTributacao_EFD_ICMS_IPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tributação - EFD ICMS/IPI"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboLayoutDocumento 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2475
   End
   Begin VB.Label Label1 
      Caption         =   "Layout do Documento:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   1635
   End
End
Attribute VB_Name = "formTributacao_EFD_ICMS_IPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

