VERSION 5.00
Begin VB.Form formClientesRelatorio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes - Relatório"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4185
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4035
      Begin VB.OptionButton OptionBt 
         Caption         =   "&Todos"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton OptionBt 
         Caption         =   "&Por &UF"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   900
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Listar"
         Height          =   555
         Left            =   2280
         TabIndex        =   2
         Top             =   660
         Width           =   1515
      End
      Begin VB.ComboBox cboUF 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   900
         Width           =   675
      End
   End
End
Attribute VB_Name = "formClientesRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If OptionBt(1).Value = True Then
        If Len(cboUF.Text) = 0 Then
            MsgBox "Selecione uma UF!", vbInformation, App.EXEName
            Exit Sub
        End If
        ImprimirListaClientes cboUF.Text
        Else
            ImprimirListaClientes
    End If
End Sub

Private Sub cboUF_DropDown()
    Dim Rst As Recordset
    cboUF.Clear
    Set Rst = RegistroBuscar("SELECT * FROM TributacaoUF ORDER BY sigla")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboUF.AddItem Rst.Fields("sigla")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub Form_Load()
    cboUF.Enabled = False
    OptionBt(0).Value = True
    cboUF.Clear
    End Sub

Private Sub OptionBt_Click(Index As Integer)
    If Index = 1 Then
            cboUF.Enabled = True
        Else
            cboUF.Enabled = False
    End If
            
End Sub
