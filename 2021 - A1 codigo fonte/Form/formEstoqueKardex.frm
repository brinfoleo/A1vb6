VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formEstoqueKardex 
   Caption         =   "Estoque - Kardex"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   13680
   Begin VB.Frame frmProduto 
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
      Height          =   1575
      Left            =   60
      TabIndex        =   9
      Top             =   480
      Width           =   7755
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1020
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton btoPesq 
         Height          =   315
         Left            =   5400
         Picture         =   "formEstoqueKardex.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txtReferencia 
         Height          =   285
         Left            =   3600
         TabIndex        =   12
         Text            =   "Text1"
         ToolTipText     =   "Pressione <F3> para consultar Produto..."
         Top             =   660
         Width           =   1755
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1020
         MaxLength       =   120
         TabIndex        =   11
         Text            =   "Text1"
         ToolTipText     =   "Pressione <F3> para consultar Produto..."
         Top             =   1080
         Width           =   5835
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "ID:"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Referencia:"
         Height          =   195
         Left            =   2700
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Depósito:"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame frmSaldo 
      Caption         =   "Saldo Atual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   11280
      TabIndex        =   7
      Top             =   7020
      Width           =   2295
      Begin VB.TextBox txtSaldoAtual 
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
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Text            =   "10000"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame frmKardex 
      Caption         =   "Kardex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   60
      TabIndex        =   5
      Top             =   2100
      Width           =   13575
      Begin MSFlexGridLib.MSFlexGrid fgrKardex 
         Height          =   4515
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   7964
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7920
      TabIndex        =   0
      Top             =   480
      Width           =   3075
      Begin MSComCtl2.DTPicker dtpPeriodoFinal 
         Height          =   315
         Left            =   660
         TabIndex        =   2
         Top             =   660
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   104792065
         CurrentDate     =   40525
      End
      Begin MSComCtl2.DTPicker dtpPeriodoInicio 
         Height          =   315
         Left            =   660
         TabIndex        =   1
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   104792065
         CurrentDate     =   40525
      End
      Begin VB.Label Label2 
         Caption         =   "Final:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   435
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   13680
      _ExtentX        =   24130
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":07DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":0AF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":1388
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":25DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":2EB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":3746
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":3FD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":522A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":5544
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":585E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueKardex.frx":5C55
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formEstoqueKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg   As Integer


Private Sub btoPesq_Click()
    IdReg = 0
    PesquisarProduto
End Sub



Private Sub fgrKardex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
    Dim i As Integer
    Dim ii As Integer
    With fgrKardex
        If .Rows = 1 Then Exit Sub
        If Trim(.TextMatrix(1, 0)) = "" Then Exit Sub

        i = IIf(.MouseRow = 0, 1, .MouseRow)
        If .MouseCol <= 0 Then
                ii = IIf(Trim(.TextMatrix(i, 0)) = "", 0, Trim(.TextMatrix(i, 0)))
                .ToolTipText = ""
            Else
                .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        End If
    End With

End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    IdReg = 0
    
    HDMenu Me, True
    MontarGrid
    dtpPeriodoInicio.Value = Date - 30
    dtpPeriodoFinal.Value = Date
    'Verifica se existe um deposito unico no sistema
    If Trim(ID_Deposito) = "" Then
            cboDeposito.Enabled = True
        Else
            cboDeposito.AddItem ID_Deposito & " - " & pgDescrDeposito(ID_Deposito)
            cboDeposito.Text = cboDeposito.List(0)
            cboDeposito.Enabled = False
    End If
    
    AtualizarGrid
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    frmKardex.Width = Me.ScaleWidth - 200
    frmKardex.Height = Me.ScaleHeight - (frmProduto.Height + frmSaldo.Height + tbMenu.Height + 300)
    
    fgrKardex.Width = Me.ScaleWidth - 400
    fgrKardex.Height = frmKardex.Height - tbMenu.Height - 300
    
    
    frmSaldo.Left = Me.Width - (frmSaldo.Width + 250)
    frmSaldo.Top = (tbMenu.Height + frmKardex.Height + frmProduto.Height + 200)
    
    
    
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualizar"
            AtualizarGrid
        Case "Manutenção da Tabela"
            ManutencaoTabela
    End Select
End Sub



Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IdReg = 0
        PesquisarProduto
    End If


End Sub



Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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

Private Sub txtreferencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IdReg = 0
        PesquisarProduto
    End If

End Sub


Private Sub txtSaldoAtual_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub ManutencaoTabela()

    Dim vReg(1000)  As Variant
    Dim cReg        As Integer 'Cota os registros
    
    'Dim sSQL As String
    cReg = 0
    vReg(cReg) = Array("Deposito", 10, "N"): cReg = cReg + 1
    vReg(cReg) = Array("IdReg", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("DataMov", 10, "D"): cReg = cReg + 1
    vReg(cReg) = Array("Data", 10, "D"): cReg = cReg + 1
    vReg(cReg) = Array("Documento", 30, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Movimento", 30, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Quantidade", 15, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Unidade", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Saldo", 15, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Descricao", 250, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ValorUnitario", 15, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ValorTotal", 15, "S"): cReg = cReg + 1
    vReg(cReg) = Array("IdNome", 10, "N"): cReg = cReg + 1
    vReg(cReg) = Array("Nome", 100, "S"): cReg = cReg + 1
    vReg(cReg) = Array("docNome", 100, "S"): cReg = cReg + 1
    vReg(cReg) = Array("NFe", 100, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Obs", 1000, "S") ': cReg = cReg + 1
    
       
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vReg, cReg
    
End Sub
Private Sub MontarGrid()
    With fgrKardex
        .Rows = 1
        .FormatString = "^ID|^Data             |^Documento   |^Movimento    " & _
                        "|<Descrição                                        " & _
                        "|^Quantidade    " & _
                        "|>Valor Unitario   " & _
                        "|>Valor Total      " & _
                        "|>Saldo            " & _
                        "|<Nome/Razão Social                                   " & _
                        "|<Obs                                                             "
    End With
End Sub

Private Sub AtualizarGrid()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    If dtpPeriodoInicio.Value > dtpPeriodoFinal.Value Then
        MsgBox "A data inicial não pode ser superior a final!", vbInformation, "Aviso"
        Exit Sub
    End If

    If Trim(IdReg) = 0 Then Exit Sub
    'sSQL = "SELECT * FROM EstoqueProduto " & _
           "WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & _
           sFiltro
    
    sSQL = "SELECT * FROM EstoqueKardex " & _
           "WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND " & _
           "IdProduto = " & IdReg & " AND " & _
           "Data >= '" & Format(dtpPeriodoInicio.Value, "yyyy-mm-dd") & "' AND Data <= '" & Format(dtpPeriodoFinal.Value, "yyyy-mm-dd") & "' " & _
           "ORDER BY Id"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            fgrKardex.Rows = 1
        Else
            fgrKardex.Rows = 1
            Rst.MoveFirst
            Do Until Rst.EOF
                fgrKardex.Rows = fgrKardex.Rows + 1
                fgrKardex.TextMatrix(fgrKardex.Rows - 1, 0) = Rst.Fields("Id")
                fgrKardex.TextMatrix(fgrKardex.Rows - 1, 1) = Rst.Fields("Data")
                fgrKardex.TextMatrix(fgrKardex.Rows - 1, 2) = IIf(IsNull(Rst.Fields("Documento")), "", Rst.Fields("Documento"))
                fgrKardex.TextMatrix(fgrKardex.Rows - 1, 3) = Rst.Fields("Movimento")
                fgrKardex.TextMatrix(fgrKardex.Rows - 1, 4) = IIf(IsNull(Rst.Fields("Descricao")), "", Trim(Rst.Fields("Descricao")))
                fgrKardex.TextMatrix(fgrKardex.Rows - 1, 5) = ChkVal(Rst.Fields("Quantidade"), 0, 3)
                fgrKardex.TextMatrix(fgrKardex.Rows - 1, 6) = ConvMoeda(Rst.Fields("ValorUnitario"))
                fgrKardex.TextMatrix(fgrKardex.Rows - 1, 7) = ConvMoeda(Rst.Fields("ValorTotal"))
                fgrKardex.TextMatrix(fgrKardex.Rows - 1, 8) = ChkVal(Rst.Fields("Saldo"), 0, 3)
                fgrKardex.TextMatrix(fgrKardex.Rows - 1, 9) = IIf(IsNull(Rst.Fields("Nome")), "", Rst.Fields("Nome"))
                fgrKardex.TextMatrix(fgrKardex.Rows - 1, 10) = IIf(IsNull(Rst.Fields("Obs")), "", Trim(Rst.Fields("Obs")))
                Rst.MoveNext
                
            Loop
           ' txtSaldoAtual.Text = fgrKardex.TextMatrix(fgrKardex.Rows - 1, 8)
    End If
    Rst.Close
End Sub
Public Sub ReceberConsultaExterna(codProd As Integer)
    
    Me.Show
    IdReg = codProd
    PesquisarProduto
    AtualizarGrid
End Sub

Private Sub cboDeposito_DropDown()
    Dim Rst As Recordset
    cboDeposito.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueDeposito ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
            
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboDeposito.AddItem Left(String(5, "0"), 5 - Len(Trim(Rst.Fields("id")))) & Trim(Rst.Fields("id")) & " - " & Rst.Fields("descricao")
                Rst.MoveNext
            Loop
    End If

End Sub
Private Sub PesquisarProduto()
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    fgrKardex.Rows = 1
    If Trim(IdReg) = 0 Then
        IdReg = formBuscar.IniciarBusca("EstoqueProduto") ', "Descricao, Referencia, CodigoBarras, NCM,IPIAliquota,ICMSCST")
        If Trim(IdReg) = 0 Then Exit Sub
    End If
    
    sSQL = "SELECT * FROM EstoqueProduto WHERE Status='ATIVO' AND Id = " & IdReg
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado"
            txtSaldoAtual.Text = "0"
        Else
            Rst.MoveFirst
            txtID.Text = IIf(IsNull(Rst.Fields("ID")), "", Rst.Fields("ID"))
            txtReferencia.Text = IIf(IsNull(Rst.Fields("Referencia")), "", Rst.Fields("Referencia"))
            txtDescricao.Text = IIf(IsNull(Rst.Fields("Descricao")), "", Rst.Fields("Descricao"))
            cboDeposito.AddItem Rst.Fields("Deposito")
            cboDeposito.Text = cboDeposito.List(0)
            txtSaldoAtual.Text = ChkVal(IIf(IsNull(Rst.Fields("Saldo")), "0", Rst.Fields("Saldo")), 0, 3)
            AtualizarGrid
    End If
    Rst.Close
    
    
End Sub
