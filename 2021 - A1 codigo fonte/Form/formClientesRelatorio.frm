VERSION 5.00
Begin VB.Form formClientesRelatorio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes - Relatório"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6270
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6075
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   2895
         Begin VB.ComboBox cboUF 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   660
            Width           =   675
         End
         Begin VB.OptionButton OptionBt 
            Caption         =   "&Por &UF"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton OptionBt 
            Caption         =   "&Todos"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1035
         End
      End
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Listar"
         Height          =   675
         Left            =   3840
         TabIndex        =   1
         Top             =   1380
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor:"
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   420
         Width           =   735
      End
   End
End
Attribute VB_Name = "formClientesRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboVendedor_DropDown()
    Dim Rst As Recordset
    cboVendedor.Clear
    Set Rst = RegistroBuscar("SELECT * FROM RHFuncionarioCadastro WHERE ID_Empresa = " & ID_Empresa & " ORDER BY xNome")
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            cboVendedor.AddItem "9999 - Todos"
            Do Until Rst.EOF
                cboVendedor.AddItem Left(String(4, "0"), 4 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & " - " & Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If
End Sub


Private Sub Command1_Click()
    Dim numVend As Integer
    numVend = 0
    If Len(cboVendedor.Text) <> 0 Then
            numVend = Left(cboVendedor.Text, 4)
            numVend = IIf(numVend = 9999, 0, numVend)
    End If
    If OptionBt(1).Value = True Then
        If Len(cboUF.Text) = 0 Then
            MsgBox "Selecione uma UF!", vbInformation, App.EXEName
            Exit Sub
        End If
        
        ImprimirListaClientes numVend, cboUF.Text
        Else
            ImprimirListaClientes numVend
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
