VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formEstoqueGerenciador 
   Caption         =   "Estoque - Gerenciador"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   12255
   Begin VB.Frame frmFiltro2 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   60
      TabIndex        =   11
      Top             =   420
      Width           =   11475
      Begin VB.ComboBox cboSubGrupo 
         Height          =   315
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   300
         Width           =   3195
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   3195
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Subgrupo:"
         Height          =   195
         Left            =   3960
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame frmGrade 
      Height          =   6615
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   11895
      Begin MSFlexGridLib.MSFlexGrid msfgGrade 
         Height          =   6195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   10927
         _Version        =   393216
         Cols            =   7
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"formEstoqueGerenciador.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12255
      _ExtentX        =   21616
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
            Object.ToolTipText     =   "Imprimir Grade"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame frmFiltro 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   0
         Width           =   8775
         Begin VB.OptionButton optBusca 
            Caption         =   "Qualquer parte"
            Height          =   195
            Index           =   1
            Left            =   7320
            TabIndex        =   10
            Top             =   120
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "Inicia com"
            Height          =   195
            Index           =   0
            Left            =   6240
            TabIndex        =   9
            Top             =   120
            Width           =   1035
         End
         Begin VB.ComboBox cboListQtd 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "formEstoqueGerenciador.frx":00EE
            Left            =   900
            List            =   "formEstoqueGerenciador.frx":010A
            TabIndex        =   6
            Text            =   "Combo1"
            Top             =   60
            Width           =   975
         End
         Begin VB.TextBox txtFiltro 
            Height          =   285
            Left            =   4140
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   60
            Width           =   1935
         End
         Begin VB.ComboBox cboFiltro 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2580
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Filtro:"
            Height          =   195
            Left            =   2100
            TabIndex        =   8
            Top             =   120
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Registros:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   735
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1320
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
               Picture         =   "formEstoqueGerenciador.frx":0139
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueGerenciador.frx":058B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueGerenciador.frx":08A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueGerenciador.frx":1137
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueGerenciador.frx":2389
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueGerenciador.frx":2C63
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueGerenciador.frx":34F5
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueGerenciador.frx":3D87
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueGerenciador.frx":4FD9
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueGerenciador.frx":52F3
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueGerenciador.frx":560D
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueGerenciador.frx":5A04
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   7320
      TabIndex        =   16
      Top             =   7920
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblTexto 
      Caption         =   "Duplo click para ficha de cadastro do produto..."
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   7920
      Width           =   6795
   End
End
Attribute VB_Name = "formEstoqueGerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboGrupo_DropDown()
    lstCbo cboGrupo, "EstoqueGrupos"
End Sub



Private Sub cbosubGrupo_DropDown()
    lstCbo cboSubGrupo, "EstoqueSubGrupo"
End Sub





Private Sub msfgGrade_DblClick()
    formEstoqueProduto.pesqLoadForm msfgGrade.TextMatrix(msfgGrade.Row, 0)
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualizar"
            LoadGrid
        Case "Imprimir Grade"
            ImprimirGrade
    End Select
End Sub
Private Sub ImprimirGrade()
    Dim Rst     As Recordset
    Dim sSQL    As String
    'Dim vCusto  As String
    'Dim vSaldo  As String
    'Dim qSaldo  As String
    
    MontarTabelaTemporaria
    
    sSQL = "SELECT * FROM tmp_EstoqueGerenciador WHERE ID_Empresa = " & ID_Empresa
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenum registro encontrado.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            'Rst.MoveFirst
            'vSaldo = 0
            'qSaldo = 0
            'Do Until Rst.EOF
            '    Status (Rst.RecordCount)
            '    vCusto = Val(ChkVal(vCusto, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("custo"), 0, cDecMoeda))
            '    qSaldo = Val(ChkVal(qSaldo, 0, cDecQtd)) + Val(ChkVal(Rst.Fields("Saldo"), 0, cDecQtd))
            '    vSaldo = Val(ChkVal(vSaldo, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("Preco"), 0, cDecMoeda))
            '    Rst.MoveNext
            'Loop
            
            Set rptListaEstoque.DataSource = Rst.DataSource
            
            rptListaEstoque.Sections("Section1").Controls("txtID").DataField = "IDOrig"
            rptListaEstoque.Sections("Section2").Controls.Item("lbltitulo").Caption = "Listagem de Estoque"
            rptListaEstoque.Sections("Section2").Controls.Item("lblCusto").Visible = True
            rptListaEstoque.Sections("Section1").Controls.Item("txtCusto").Visible = True
            rptListaEstoque.Sections("Section2").Controls.Item("lblPreco").Visible = True
            rptListaEstoque.Sections("Section1").Controls.Item("txtPreco").Visible = True
            
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot1").Visible = False
            'rptListaEstoque.Sections("Section5").Controls.Item("lblTot1").Caption = ChkVal(qSaldo, 0, cDecQtd)
            
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot2").Visible = False
            'rptListaEstoque.Sections("Section5").Controls.Item("lblTot2").Caption = ConvMoeda(vCusto)
            
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot3").Visible = False
            'rptListaEstoque.Sections("Section5").Controls.Item("lblTot3").Caption = ConvMoeda(vSaldo)
            
            rptListaEstoque.Show 1
            Rst.Close
    End If
End Sub
Private Sub cboFiltro_DropDown()
    Dim i As Integer
    cboFiltro.Clear
    For i = 0 To msfgGrade.Cols - 1
        cboFiltro.AddItem msfgGrade.TextMatrix(0, i)
    Next
End Sub

Private Sub cboListQtd_Click()
    LoadGrid
End Sub

Private Sub cboListQtd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        LoadGrid
    End If
    KeyAscii = IIf(IsNumeric(Chr(KeyAscii)), KeyAscii, 0)
    
    
End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LimpForm
    LoadGrid
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With frmFiltro2
        .Top = 540
        .Left = 80
        .Width = formEstoqueGerenciador.Width - 300
    End With
    
    With frmGrade
        '.Top = 540
        .Top = frmFiltro2.Height + frmFiltro2.Top
        .Left = 80
        .Width = formEstoqueGerenciador.Width - 300
        '.Height = formEstoqueGerenciador.Height - (1140 + pb.Height)
        .Height = formEstoqueGerenciador.Height - (1140 + pb.Height + frmFiltro2.Height)
    End With
    With msfgGrade
        .Height = frmGrade.Height - 350
        .Width = frmGrade.Width - 250
    End With
    With frmFiltro
        .Left = formEstoqueGerenciador.Width - (.Width + 100)
    End With
    With pb
        .Top = frmGrade.Height + frmGrade.Top + 50
        .Left = formEstoqueGerenciador.Width - (.Width + 200)
    End With
     With lblTexto
        .Top = frmGrade.Height + frmGrade.Top + 50
        .Left = 80
    End With

End Sub
Private Sub LoadGrid()
    On Error GoTo ErrLoadGrid
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    
    '*************************************************************************************************************
    'Quantidade de registros apresentados
    Dim nmReg   As String
    Select Case Trim(cboListQtd.Text)
        Case ""
            cboListQtd.Text = cboListQtd.List(0)
            nmReg = " LIMIT " & Trim(cboListQtd.Text)
        Case "(Todos)"
            nmReg = ""
        Case Else
            nmReg = " LIMIT " & Trim(cboListQtd.Text)
    End Select
    '*************************************************************************************************************
    '*************************************************************************************************************
    'Filtra por Grupo e SubGrupo
    Dim fGrupo      As String
    Dim fSubGrupo   As String
    Dim Class       As String
    If Trim(cboGrupo.Text) = "" Then
            fGrupo = ""
        Else
            fGrupo = "Grupo = " & Left(Trim(cboGrupo.Text), 5)
    End If
    If Trim(cboSubGrupo.Text) = "" Then
            fSubGrupo = ""
        Else
            fSubGrupo = "SubGrupo = " & Left(Trim(cboSubGrupo.Text), 5)
    End If
    
    
    Class = fGrupo
    
    If Class = "" Then
            Class = fSubGrupo
        Else
            Class = IIf(Trim(fSubGrupo) <> "", Class & " AND " & fSubGrupo, Class)
    End If
    '*************************************************************************************************************
    
    
    '*************************************************************************************************************
    'Filtra por texto digitado
    Dim lFiltro As String
    Dim sFiltro As String 'String com o filtro aplicado
    Dim sBusca  As String
    Dim parte   As String
    Dim sBtmp   As String
    Dim sOrdem  As String
    
    If Trim(cboFiltro.Text) = "" Then
            sFiltro = ""
        Else
            lFiltro = cboFiltro.Text
            
            If optBusca(0).Value = True Then
                    '*********************************************************
                    '*** Efetua abusca com o texto digitado
                    '*********************************************************
                    Select Case lFiltro
                        Case "Preço Venda"
                            lFiltro = "Preco"
                            sBtmp = lFiltro & " = " & ChkVal(txtFiltro.Text, 0, cDecMoeda) '& "%"
                        Case Else
                            sBtmp = lFiltro & " LIKE '" & txtFiltro.Text & "%'"
                    End Select
                Else
                    '*********************************************************
                    '*** Efetua a   busca com qualquer parte do texto digitado
                    '*********************************************************
                    sBusca = Replace(Trim(txtFiltro.Text), " ", "|") & "|"
                    Do Until InStr(sBusca, "|") = 0
                        parte = Trim(Mid(sBusca, 1, InStr(sBusca, "|") - 1))
                        parte = Replace(parte, "'", "''")
                        Select Case lFiltro
                            Case "Preço Venda"
                                lFiltro = "Preco"
                                If Trim(parte) <> "" Then 'Verifica se nao esta buscando um valor vazio
                                    parte = ChkVal(parte, 0, cDecMoeda)
                                    sBtmp = IIf(Trim(sBtmp) = "", "", sBtmp & " AND ") & lFiltro & "=" & Trim(parte) '& " OR " & lFiltro & "<=" & Trim(parte)
                                End If
                            Case Else
                                sBtmp = IIf(Trim(sBtmp) = "", "", sBtmp & " AND ") & lFiltro & " LIKE '%" & Trim(parte) & "%'"
                        End Select
                        sBusca = Mid(sBusca, InStr(sBusca, "|") + 1, Len(sBusca))
                    Loop
            End If
            
            sOrdem = " ORDER BY " & lFiltro
    End If
    
    sFiltro = IIf(Trim(sBtmp) = "", "", " AND " & sBtmp) & _
              IIf(Trim(Class) = "", "", " AND " & Class) & _
              sOrdem & _
              nmReg

    '*************************************************************************************************************
    '*************************************************************************************************************
    'Pega os dados no BD
    msfgGrade.Rows = 1
    sSQL = "SELECT * FROM estoqueproduto " & _
           "WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND Status = 'ATIVO' " & _
           sFiltro
    Set Rst = RegistroBuscar(sSQL)
    
    If Rst.BOF And Rst.EOF Then
            Rst.Close
            Exit Sub
        Else
            'Status (Rst.RecordCount)
            Rst.MoveFirst
    End If
    'Lista na grid
    With msfgGrade
        Do Until Rst.EOF
            DoEvents
            status (Rst.RecordCount)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Left(String(8, "0"), 8 - Len(Trim(Rst.fields("ID")))) & Trim(Rst.fields("ID"))
            .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rst.fields("Referencia")), "", Rst.fields("Referencia"))
            .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rst.fields("Descricao")), "", Rst.fields("Descricao"))
            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.fields("Unidade")), "", Rst.fields("Unidade"))
            
            '********************************************************************************
            'Alteração somente para saber se sai no balanco
            'If cNull(Rst.Fields("IncluirBalanco")) = 0 Or Trim(cNull(Rst.Fields("IncluirBalanco"))) = "" Then
            '        .TextMatrix(.Rows - 1, 3) = .TextMatrix(.Rows - 1, 3)
            '    Else
            '        .TextMatrix(.Rows - 1, 3) = .TextMatrix(.Rows - 1, 3) & "*"
            'End If
            '********************************************************************************
            
            .TextMatrix(.Rows - 1, 4) = ChkVal(IIf(IsNull(Rst.fields("Saldo")), "0", Rst.fields("Saldo")), 0, cDecQtd)
            .TextMatrix(.Rows - 1, 5) = ConvMoeda(ChkVal(IIf(IsNull(Rst.fields("Custo")), "0", Rst.fields("Custo")), 0, cDecMoeda))
            .TextMatrix(.Rows - 1, 6) = ConvMoeda(ChkVal(IIf(IsNull(Rst.fields("Preco")), "0", Rst.fields("Preco")), 0, cDecMoeda))
            
            'Muda a cor de acordo com o saldo
            If ChkVal(IIf(IsNull(Rst.fields("Saldo")), "0", Rst.fields("Saldo")), 0, cDecQtd) < 0 Then
                .FillStyle = flexFillRepeat
                .Row = .Rows - 1
                .Col = 0
                .ColSel = .Cols - 1
                '.CellFontBold = True
                .CellForeColor = vbRed
            End If
            
            Rst.MoveNext
        Loop
    End With
    Exit Sub
ErrLoadGrid:
    msfgGrade.Rows = 1
    Resume Next
    
End Sub
Private Sub LimpForm()
    cboListQtd.Text = ""
    txtFiltro.Text = ""
    msfgGrade.Rows = 1
End Sub
Private Sub status(Max As Long)
    pb.min = 0
    pb.Max = Max
    DoEvents
    pb.Value = pb.Value + 1
    If pb.Value > 0 And pb.Value < Max Then
            pb.Visible = True
            Me.Enabled = False
        Else
            pb.Visible = False
            pb.Value = 0
            Me.Enabled = True
    End If
End Sub

Private Sub txtFiltro_Change()
    'LoadGrid
End Sub
Private Sub MontarTabelaTemporaria()
    Dim sCampos     As String
    Dim i           As Integer
    Dim ii          As Integer
    Dim vReg(1000)  As Variant
    Dim cReg        As Integer
    
    BD.Execute "DROP TABLE IF EXISTS tmp_estoquegerenciador"
    
    sCampos = ""
    For i = 1 To msfgGrade.Cols - 1
        sCampos = sCampos & RS(rc(msfgGrade.TextMatrix(0, i))) & " VARCHAR(100) default Null,"
    Next
    sCampos = "CREATE TABLE IF NOT EXISTS tmp_estoquegerenciador " & _
              "(Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "UsuID VARCHAR(10) default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "IDOrig VARCHAR(30) default Null," & _
               sCampos & " PRIMARY KEY (" & RS(rc(msfgGrade.TextMatrix(0, 0))) & "))"
    BD.Execute sCampos
    
    With msfgGrade
        For i = 1 To .Rows - 1
            cReg = 0
            vReg(cReg) = Array("IDOrig", .TextMatrix(i, 0), "S"): cReg = cReg + 1
            For ii = 1 To .Cols - 1
                status (.Rows - 1)
                vReg(cReg) = Array(RS(rc(.TextMatrix(0, ii))), .TextMatrix(i, ii), "S"): cReg = cReg + 1
            Next
            cReg = cReg - 1
            RegistroIncluir "tmp_EstoqueGerenciador", vReg, cReg
        Next
    End With
    
    
End Sub

Private Sub txtFiltro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then LoadGrid
End Sub

Private Sub lstCbo(cbo As Control, sTabela As String)
    Dim Rst As Recordset
    cbo.Clear
    Set Rst = RegistroBuscar("SELECT * FROM " & sTabela & " ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cbo.AddItem Left(String(5, "0"), 5 - Len(Rst.fields("Id"))) & Rst.fields("Id") & _
                                 " - " & Rst.fields("descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub
