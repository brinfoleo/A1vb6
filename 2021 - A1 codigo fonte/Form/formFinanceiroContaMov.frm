VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formFinanceiroContaMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimento de Conta"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   3750
   Begin VB.CommandButton btoGravar 
      Caption         =   "&Gravar"
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   3420
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Conta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      Begin VB.ComboBox cboConta 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Width           =   3615
      Begin VB.TextBox txtDocDesc 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1140
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   104398849
         CurrentDate     =   42524
      End
      Begin VB.TextBox txtDoc 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   780
         Width           =   1275
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   900
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1980
         Width           =   1335
      End
      Begin VB.OptionButton optAcaoCD 
         Caption         =   "Débito"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton optAcaoCD 
         Caption         =   "Crédito"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1980
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Data:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Documento:"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   780
         Width           =   915
      End
   End
End
Attribute VB_Name = "formFinanceiroContaMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private idConta As Integer
Private opcCD   As String 'Opcao Cred / Deb

Private Sub btoGravar_Click()
    If gravar = True Then
        LimpForm
        Exit Sub
    End If
End Sub

Private Sub cboConta_Click()
    If Trim(cboConta.Text) = "" Then Exit Sub
    idConta = Left(Trim(cboConta.Text), 3)
    
End Sub

Private Sub cboConta_DropDown()
    Dim sSQL    As String
    Dim Rst     As Recordset
    LimpForm
    sSQL = "SELECT * FROM FinanceiroConta"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma conta cadastrada!", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboConta.AddItem ZE(Rst.Fields("ID"), 3) & " - " & Rst.Fields("Agencia") & "/" & Rst.Fields("Conta")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub
Private Sub LimpForm()
    cboConta.Clear
    'cboConta.Text = ""
    dtpData.Value = Date
    txtDoc.Text = ""
    txtDocDesc.Text = ""
    optAcaoCD(0).Value = False
    optAcaoCD(1).Value = False
    txtValor.Text = ""
    
    
    idConta = 0
    opcCD = ""
End Sub

Private Sub Form_Load()
    LimpForm
End Sub

Private Sub optAcaoCD_Click(Index As Integer)
    If optAcaoCD(0).Value = True Then
            txtValor.ForeColor = vbBlue
            opcCD = "C"
        ElseIf optAcaoCD(1).Value = True Then
            txtValor.ForeColor = vbRed
            opcCD = "D"
        Else
            txtValor.ForeColor = vbWhite
            opcCD = ""
    End If
       
End Sub
Private Function gravar() As Boolean
On Error GoTo TrtErrGrav
    If idConta = 0 Then
        MsgBox "Selecione uma conta!", vbInformation, App.EXEName
        gravar = False
        Exit Function
    End If
    If opcCD = "" Then
        MsgBox "Selecione opção: CRÉDITO ou DÉBITO", vbInformation, App.EXEName
        gravar = False
        Exit Function
    End If
    If Val(ChkVal(txtValor.Text, 0, cDecMoeda)) <= 0 Then
        MsgBox "Informe um VALOR valido!", vbInformation, App.EXEName
        gravar = False
        Exit Function
    End If
    MovimentarConta idConta, opcCD, _
                        "0", _
                        dtpData.Value, _
                        txtDoc.Text, _
                        0, _
                        txtDocDesc.Text, txtValor.Text
                        
    MsgBox "Registro gravado com sucesso!", vbInformation, App.EXEName
    gravar = True
    Exit Function
TrtErrGrav:
    gravar = False
    MsgBox "Erro ao gravar o registro!" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, App.EXEName
End Function

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtValor.Text, KeyAscii, cDecMoeda)
End Sub
