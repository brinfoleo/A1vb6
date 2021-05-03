VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formFaturamentoPVPagamento 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8115
      Begin VB.CommandButton btoConcluido 
         Cancel          =   -1  'True
         Caption         =   "&Concluido"
         Height          =   435
         Left            =   6120
         TabIndex        =   5
         Top             =   2700
         Width           =   1815
      End
      Begin VB.CommandButton btoAplicar 
         Height          =   375
         Left            =   2400
         Picture         =   "formFaturamentoPVPagamento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aplicar valor..."
         Top             =   2700
         Width           =   375
      End
      Begin VB.TextBox txtNovoValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   660
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2760
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid msfgParc 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   3
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "^Parcela    |^Vencimento            |>Valor                           "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor:"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   2820
         Width           =   435
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Está sendo considerada a data de emissão da pré-venda para o(s) calculo(s) do(s) vencimento(s)."
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   3420
      Width           =   6915
   End
End
Attribute VB_Name = "formFaturamentoPVPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sValor      As String
'Dim vTotal      As String
Dim aCob(100)   As Variant
Dim cCob        As Integer


Public Sub CarregarFormulario(ArrayParcelas As Variant, numParcelas As Integer)
    On Error GoTo TrtErro
    'aCob = ArrayParcelas
    cCob = numParcelas
    CopyArray ArrayParcelas, aCob
    '###################################################################################
    '### Mostrar Parcelas
    '###################################################################################
    Dim i As Integer
    With msfgParc
            .Rows = 1
            sValor = 0
        For i = 0 To cCob
            .Rows = .Rows + 1
            sValor = Val(ChkVal(sValor, 0, cDecMoeda)) + Val(ChkVal(CStr(aCob(i)(1)), 0, cDecMoeda))
            .TextMatrix(.Rows - 1, 0) = i + 1 & "/" & cCob + 1
            .TextMatrix(.Rows - 1, 1) = aCob(i)(0) 'Date + Rst.Fields("DiasCorridos")
            .TextMatrix(.Rows - 1, 2) = ConvMoeda(CStr(aCob(i)(1))) 'Val(ChkVal(IIf(IsNull(Rst.Fields("Percentual")), 0, Rst.Fields("Percentual")), 0, 3)) * Val(ChkVal(sValor, 0, 2)) / 100
            'cCob = cCob + 1
        Next
    End With

    Me.Show 1
    CopyArray aCob, ArrayParcelas
    Exit Sub
TrtErro:
    Unload Me
End Sub

Private Sub btoAplicar_Click()
    Dim tParc   As Integer 'Total de parcelas
    Dim lParc   As Integer 'parcela sendo modificada
    Dim i       As Integer
    If ChkVal(txtNovoValor.Text, 0, cDecMoeda) = ChkVal("0", 0, cDecMoeda) Then
        MsgBox "O valor da parcela não pode ser zerado!", vbInformation, App.EXEName
        Exit Sub
    End If
    With msfgParc
        .TextMatrix(.Row, 2) = ConvMoeda(txtNovoValor.Text)
        txtNovoValor.Text = ""
    End With
   
End Sub

Private Sub btoConcluido_Click()
    Dim i       As Integer
    Dim Soma    As String
    Soma = 0
    With msfgParc
        For i = 0 To .Rows - 2
            aCob(i)(1) = ChkVal(.TextMatrix(i + 1, 2), 0, cDecMoeda)
            Soma = Val(ChkVal(Soma, 0, cDecMoeda)) + Val(ChkVal(CStr(aCob(i)(1)), 0, cDecMoeda))
        Next
    End With
    If ChkVal(Soma, 0, cDecMoeda) <> ChkVal(sValor, 0, cDecMoeda) Then
            MsgBox "Valor divergente do total da Pre-Venda!", vbInformation, App.EXEName
        Else
            Unload Me
    End If
End Sub

Private Sub Form_Load()
    txtNovoValor.Text = ""
End Sub

Private Sub msfgParc_Click()
    txtNovoValor.Text = ChkVal(msfgParc.TextMatrix(msfgParc.Row, 2), 0, cDecMoeda)
    txtNovoValor.SetFocus
End Sub
Private Sub txtNovoValor_GotFocus()
     With txtNovoValor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtNovoValor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        btoAplicar_Click
    End If
    
    If txtNovoValor.SelLength = Len(txtNovoValor.Text) Then
        txtNovoValor.Text = ""
    End If
    KeyAscii = ChkVal(txtNovoValor.Text, KeyAscii, cDecMoeda)
End Sub
