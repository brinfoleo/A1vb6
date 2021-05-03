VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formCalendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurações - Calendario"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2715
      Left            =   3240
      TabIndex        =   2
      Top             =   180
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4789
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3015
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   4180
         _Version        =   393216
         Appearance      =   1
         StartOfWeek     =   55574529
         CurrentDate     =   40707
      End
   End
End
Attribute VB_Name = "formCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

