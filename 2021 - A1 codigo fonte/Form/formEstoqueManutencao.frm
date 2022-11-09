VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formEstoqueManutencao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estoque - Manutenção"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   6480
   Begin VB.Frame Frame3 
      Height          =   1755
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   6255
      Begin VB.TextBox txtDocumento 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1380
         Width           =   1995
      End
      Begin VB.TextBox txtCNPJ 
         Height          =   285
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   660
         Width           =   2055
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1020
         Width           =   4395
      End
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   127336449
         CurrentDate     =   40540
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Documento:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ/CPF:"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   195
         Left            =   300
         TabIndex        =   24
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Data:"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.TextBox txtSaldoFin 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4140
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   7200
      Width           =   2235
   End
   Begin VB.Frame Frame2 
      Height          =   4815
      Left            =   60
      TabIndex        =   14
      Top             =   2340
      Width           =   6315
      Begin VB.Frame Frame1 
         Height          =   2295
         Left            =   120
         TabIndex        =   28
         Top             =   2400
         Width           =   6075
         Begin VB.TextBox txtObs 
            Height          =   795
            Left            =   1380
            MaxLength       =   65000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Text            =   "formEstoqueManutencao.frx":0000
            Top             =   1380
            Width           =   4455
         End
         Begin VB.TextBox txtQuantidade 
            Height          =   285
            Left            =   1380
            MaxLength       =   20
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   300
            Width           =   1995
         End
         Begin VB.TextBox txtVlUnitario 
            Height          =   285
            Left            =   1380
            MaxLength       =   20
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   660
            Width           =   1995
         End
         Begin VB.TextBox txtVlTotal 
            Height          =   285
            Left            =   1380
            MaxLength       =   20
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1020
            Width           =   1995
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Quantidade:"
            Height          =   195
            Left            =   360
            TabIndex        =   32
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor Unitário:"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Total:"
            Height          =   195
            Left            =   360
            TabIndex        =   30
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Obs:"
            Height          =   195
            Left            =   720
            TabIndex        =   29
            Top             =   1380
            Width           =   555
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   17
         Top             =   540
         Width           =   6075
         Begin VB.TextBox txtReferencia 
            Height          =   285
            Left            =   960
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   600
            Width           =   1755
         End
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   960
            MaxLength       =   120
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   960
            Width           =   4995
         End
         Begin VB.TextBox txtSaldoIni 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   1380
            Width           =   1635
         End
         Begin VB.TextBox txtID 
            Height          =   285
            Left            =   960
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label lblUnidade 
            Caption         =   "Label15"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   33
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Referencia:"
            Height          =   195
            Left            =   60
            TabIndex        =   22
            Top             =   660
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   60
            TabIndex        =   21
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Saldo:"
            Height          =   195
            Left            =   420
            TabIndex        =   20
            Top             =   1440
            Width           =   435
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "ID:"
            Height          =   195
            Left            =   60
            TabIndex        =   19
            Top             =   300
            Width           =   855
         End
      End
      Begin VB.ComboBox cboMovimento 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   5175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Movimento:"
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pesquisar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
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
               Picture         =   "formEstoqueManutencao.frx":0006
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueManutencao.frx":0458
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueManutencao.frx":0772
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueManutencao.frx":1004
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueManutencao.frx":2256
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueManutencao.frx":2B30
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueManutencao.frx":33C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueManutencao.frx":3C54
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueManutencao.frx":4EA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueManutencao.frx":51C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueManutencao.frx":54DA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Saldo Final:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1380
      TabIndex        =   16
      Top             =   7320
      Width           =   2655
   End
End
Attribute VB_Name = "formEstoqueManutencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim IdReg     As Long
Dim strTabela   As String

Private Sub cboMovimento_Click()
    CalcSaldoTotalEstoque
End Sub

Private Sub cboMovimento_DropDown()
    Dim Rst As Recordset
    Dim sSQL As String
    sSQL = "SELECT * FROM EstoqueMovimento"
    cboMovimento.Clear
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboMovimento.AddItem Left(String(5, "0"), 5 - Len(Rst.Fields("ID"))) & Trim(Rst.Fields("ID")) & " - " & _
                Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
End Sub


Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    lblUnidade.Caption = ""
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    dtpData.Value = Date
    'HDForm Me, False
    HDMenu Me, False
    tbMenu.Buttons(1).Enabled = True
    
    
End Sub




Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Salvar"
            If IdReg = 0 Then
                MsgBox "Favor selecionar um Produto!", vbInformation, "Aviso"
                Exit Sub
            End If
            If grvRegistro = True Then
                    MsgBox "Registro gravado com Sucesso!", vbInformation, "Aviso"
                    IdReg = 0
                    LpForm
                    txtID.SetFocus
                    Exit Sub
                Else
                    MsgBox "Erro ao gravar, verificar.", vbInformation, "Aviso"
            End If
        Case "Cancelar"
            IdReg = 0
            LpForm
        Case "Pesquisar"
            PesquisarProduto
    End Select

End Sub
Private Sub LpForm()
    txtID.Text = ""
    txtReferencia.Text = ""
    txtDescricao.Text = ""
    txtSaldoIni.Text = "0"
    lblUnidade.Caption = ""
    
    txtQuantidade.Text = ""
    txtVlUnitario.Text = ""
    txtVlTotal.Text = ""
    txtObs.Text = ""
    txtSaldoFin.Text = "0"
End Sub
Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IdReg = 0
        PesquisarProduto
    End If
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IdReg = 0
        PesquisarProduto
    End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        IdReg = txtID.Text
        PesquisarProduto
    End If
    KeyAscii = IIf(IsNumeric(Chr(KeyAscii)), KeyAscii, 0)
End Sub

Private Sub txtQuantidade_Change()
    CalcValorTotalItem
    CalcSaldoTotalEstoque
End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtQuantidade.Text, KeyAscii, cDecQtd)
End Sub


Private Sub txtreferencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IdReg = 0
        PesquisarProduto
    End If
End Sub
            


Private Function validaFormulario() As Boolean
    If Trim(txtDocumento.Text) = "" Then
        MsgBox "Favor informar o numero do documento!", vbInformation, App.EXEName
        validaFormulario = False
        Exit Function
    End If
    If Trim(txtNome.Text) = "" Then
        MsgBox "Favor informar nome ou razão social!", vbInformation, App.EXEName
        validaFormulario = False
        Exit Function
    End If
    If Trim(cboMovimento.Text) = "" Then
        MsgBox "Favor selecionar o tipo de movimento!", vbInformation, App.EXEName
        validaFormulario = False
        Exit Function
    End If
    If Trim(txtID.Text) = "" Then
        MsgBox "Favor informar um produto!", vbInformation, App.EXEName
        validaFormulario = False
        Exit Function
    End If
    validaFormulario = True
End Function
Private Function grvRegistro() As Boolean
    If validaFormulario = False Then Exit Function
    Dim acao As String
    acao = IIf(pgMovEst(Left(cboMovimento.Text, 5)).AcaoDescr = "SOMAR (+)", "e", "s")
    If MovimentarEstoque(acao, IdReg, dtpData.Value, Trim(txtDocumento.Text), Trim(txtQuantidade.Text), Trim(txtVlUnitario.Text), _
                        Trim(ChkVal(txtVlTotal.Text, 0, 2)), Trim(txtObs.Text), Trim(txtNome.Text), , , Trim(txtCNPJ.Text)) = True Then
            grvRegistro = True
        Else
            grvRegistro = False
    End If

    
'    Dim vReg(199)   As Variant
'    Dim cReg        As Integer
'    cReg = 0
'    vReg(cReg) = Array("Deposito", Trim(Left(cboDeposito.Text, 5)), "S"): cReg = cReg + 1
'    vReg(cReg) = Array("IdProduto", IdReg, "S"): cReg = cReg + 1
'    vReg(cReg) = Array("Data", dtpData.Value, "D"): cReg = cReg + 1
'    vReg(cReg) = Array("Documento", Trim(txtDocumento.Text), "S"): cReg = cReg + 1
'    vReg(cReg) = Array("Movimento", pgMovEst(Left(cboMovimento.Text, 5)).AcaoDescr, "S"): cReg = cReg + 1
'    vReg(cReg) = Array("Descricao", pgMovEst(Left(cboMovimento.Text, 5)).Descricao, "S"): cReg = cReg + 1
'    vReg(cReg) = Array("Quantidade", Trim(txtQuantidade.Text), "S"): cReg = cReg + 1
'    vReg(cReg) = Array("Saldo", Trim(txtSaldoFin.Text), "S"): cReg = cReg + 1
'    vReg(cReg) = Array("ValorUnitario", Trim(txtVlUnitario.Text), "S"): cReg = cReg + 1
'    vReg(cReg) = Array("ValorTotal", Trim(ChkVal(txtVlTotal.Text, 0, 2)), "S"): cReg = cReg + 1
'    vReg(cReg) = Array("Obs", Trim(txtObs.Text), "S") ': cReg = cReg + 1
'
'    If RegistroIncluir("estoquekardex", vReg, cReg) = 0 Then
'            MsgBox "Erro ao Incluir"
'            grvRegistro = False
'            Exit Function
'        Else
 '           grvRegistro = True
'    End If
'
'    'Alterrar Saldo em produto
'    vReg(0) = Array("Saldo", Trim(txtSaldoFin.Text), "S")
'    If RegistroAlterar("estoqueproduto", vReg, 0, "Id = " & IdReg) = False Then
'            MsgBox "Erro ao Alterar."
 '           grvRegistro = False
 ''       Else
 '           grvRegistro = True
 '   End If


End Function


Private Sub MostrarDados()
    Dim sSQL As String
    sSQL = "SELECT * FROM " & strTabela & " WHERE Id = " & IdReg

    ExibirDados Me, sSQL


End Sub

Private Sub PesquisarProduto()
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    If IdReg = 0 Then
        IdReg = formBuscar.IniciarBusca("EstoqueProduto")
    End If
    
    If Trim(IdReg) = 0 Then
        LpForm
        Exit Sub
    End If
    
    sSQL = "SELECT * FROM EstoqueProduto WHERE Id_Empresa = " & ID_Empresa & " AND " & _
           " Deposito = " & ID_Deposito & " AND " & _
           "Id = " & IdReg
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado!", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            txtID.Text = IdReg
            txtReferencia.Text = IIf(IsNull(Rst.Fields("Referencia")), "", Rst.Fields("Referencia"))
            txtDescricao.Text = IIf(IsNull(Rst.Fields("Descricao")), "", Rst.Fields("Descricao"))
            txtSaldoIni.Text = IIf(IsNull(Rst.Fields("Saldo")), 0, Rst.Fields("Saldo"))
            lblUnidade.Caption = IIf(IsNull(Rst.Fields("Unidade")), "", "/" & Rst.Fields("Unidade"))
    End If
    Rst.Close
    CalcSaldoTotalEstoque
    
End Sub

Private Sub txtSaldoIni_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub CalcValorTotalItem()
    Dim vTotI As String
    vTotI = Val(ChkVal(txtQuantidade.Text, 0, 2)) * Val(ChkVal(txtVlUnitario.Text, 0, 2))
    If Val(vTotI) > 0 Then
        txtVlTotal.Text = ConvMoeda(vTotI)
    End If
End Sub
Private Sub CalcSaldoTotalEstoque()
    If Trim(cboMovimento.Text) = "" Then Exit Sub
    Select Case pgMovEst(Left(Trim(cboMovimento.Text), 5)).acao
        Case "+"
            txtSaldoFin.Text = Val(ChkVal(txtSaldoIni.Text, 0, 3)) + Val(ChkVal(txtQuantidade.Text, 0, 3))
        Case "-"
            txtSaldoFin.Text = Val(ChkVal(txtSaldoIni.Text, 0, 3)) - Val(ChkVal(txtQuantidade.Text, 0, 3))
        Case "N"
            'Fara somente registro
            txtSaldoFin.Text = txtSaldoIni.Text
        Case Else
            MsgBox "Favor informar o tipo de movimento.", vbInformation, "Aviso"
    End Select
End Sub

Private Sub txtVlTotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtVlUnitario.Text = Val(ChkVal(txtVlTotal.Text, 0, cDecMoeda)) / Val(ChkVal(txtQuantidade.Text, 0, cDecQtd))
        txtVlUnitario.Text = ConvMoeda(ChkVal(txtVlUnitario.Text, 0, cDecMoeda))
    End If
End Sub

Private Sub txtVlUnitario_Change()
    CalcValorTotalItem
End Sub


Private Sub txtVlUnitario_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtVlUnitario.Text, KeyAscii, cDecMoeda)
End Sub


