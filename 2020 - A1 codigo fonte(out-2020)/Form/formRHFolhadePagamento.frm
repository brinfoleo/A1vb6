VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formRHFolhadePagamento 
   Caption         =   "RH - Folha de Pagamento"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   13410
   Begin VB.Frame Frame2 
      Caption         =   "Movimento:"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   3540
      Width           =   12795
      Begin VB.ComboBox cboCD 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   2355
      End
      Begin VB.TextBox txtValor 
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1380
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   180
         Width           =   7755
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Credito/Debito:"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor:"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   420
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Extrato:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   60
      TabIndex        =   4
      Top             =   1260
      Width           =   12855
      Begin MSFlexGridLib.MSFlexGrid msfgMov 
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   12675
         _ExtentX        =   22357
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   3
         FormatString    =   $"formRHFolhadePagamento.frx":0000
      End
   End
   Begin VB.TextBox txtMesAno 
      Height          =   315
      Left            =   7620
      MaxLength       =   7
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   780
      Width           =   975
   End
   Begin VB.ComboBox cboFuncionario 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   5895
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13410
      _ExtentX        =   23654
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alterar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pesquisar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Manutenção da Tabela"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5280
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFolhadePagamento.frx":00AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFolhadePagamento.frx":04FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFolhadePagamento.frx":0818
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFolhadePagamento.frx":10AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFolhadePagamento.frx":22FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFolhadePagamento.frx":2BD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFolhadePagamento.frx":3468
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFolhadePagamento.frx":3CFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFolhadePagamento.frx":4F4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFolhadePagamento.frx":5266
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFolhadePagamento.frx":5580
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Mes/Ano:"
      Height          =   195
      Left            =   6720
      TabIndex        =   2
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Funcionário:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   975
   End
End
Attribute VB_Name = "formRHFolhadePagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCD_Change()

End Sub

Private Sub cboCD_DropDown()
    With cboCD
        .Clear
        .AddItem "C - Crédito"
        .AddItem "D - Débito"
    End With
End Sub
