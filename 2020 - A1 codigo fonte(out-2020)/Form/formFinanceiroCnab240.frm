VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formFinanceiroCnab240 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A1 - CNAB 240"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboConta 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1620
      Width           =   3915
   End
   Begin VB.Frame frmPeriodo 
      Caption         =   "Periodo Emissão:"
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
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4275
      Begin MSComCtl2.DTPicker dtpDtInicio 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   117309441
         CurrentDate     =   40557
      End
      Begin MSComCtl2.DTPicker dtpDtFinal 
         Height          =   315
         Left            =   2460
         TabIndex        =   2
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   117309441
         CurrentDate     =   40557
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Até:"
         Height          =   195
         Left            =   2100
         TabIndex        =   4
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "De:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Conta:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "formFinanceiroCnab240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Obs:
'
' No tipo de documento existe a opcao de impressao
' Filtrar tds os q forem boleto para ser gerado o cnab240
'
'
'

'Private Sub cboConta_Click()
'    On Error GoTo TrtErro
'    Dim Rst     As Recordset
'    Dim sSQL    As String
'    Dim idConta As Integer
'
'    If Trim(RS(cboConta.Text)) = "" Then Exit Sub
'    idConta = Trim(Left(cboConta.Text, 3))
'    sSQL = "SELECT * FROM FinanceiroConta WHERE ID_Empresa = " & ID_Empresa & " AND id = " & idConta
'
'    Set Rst = RegistroBuscar(sSQL)
'    If Rst.BOF And Rst.EOF Then
'            MsgBox "Erro ao localizar conta"
'            Exit Sub
'        Else
'            Rst.MoveFirst
'            'txtMulta.Text = IIf(IsNull(Rst.Fields("Multa")), "0", Rst.Fields("Multa"))
'            'txtJuros.Text = IIf(IsNull(Rst.Fields("Juros")), "0", Rst.Fields("Juros"))
'            '22.02.2017 - if removido pois nao atualizava quando mudava a conta
'            'cboconta
'            'If Trim(cboBanco.Text) = "" Then
'                cboBanco.Clear
'                cboBanco.AddItem IIf(IsNull(Rst.Fields("banco")), " ", Left("000", 3 - Len(Rst.Fields("banco"))) & Rst.Fields("banco")) & " - " & pgDadosBanco(Rst.Fields("banco")).Nome
'                cboBanco.Text = cboBanco.List(0)
'            'End If
'            txtDiasProtesto.Text = IIf(IsNull(Rst.Fields("DiasProtesto")), "0", Rst.Fields("DiasProtesto"))
'    End If
'    Rst.Close
'    Exit Sub
'TrtErro:
'    Exit Sub
'End Sub
Private Sub Form_Load()

End Sub
