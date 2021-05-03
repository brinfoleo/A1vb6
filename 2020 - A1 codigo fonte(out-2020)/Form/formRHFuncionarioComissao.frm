VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formRHFuncionarioComissao 
   Caption         =   "RH - Comissão"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Notas Fiscais"
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
      Left            =   6420
      TabIndex        =   11
      Top             =   540
      Width           =   2835
      Begin VB.CheckBox chkMovFinanceiro 
         Caption         =   "Mov. Financeiro"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   540
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox chkMovFisco 
         Caption         =   "Mov. Fisco"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listagem por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      TabIndex        =   9
      Top             =   480
      Width           =   6255
      Begin VB.TextBox txtNFIni 
         Height          =   285
         Left            =   1500
         MaxLength       =   9
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtNFFim 
         Height          =   285
         Left            =   1500
         MaxLength       =   9
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   540
         Width           =   1335
      End
      Begin VB.OptionButton optListagem 
         Caption         =   "Num. Nota:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   1155
      End
      Begin VB.OptionButton optListagem 
         Caption         =   "Data Emissão:"
         Height          =   195
         Index           =   1
         Left            =   3300
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpNFIni 
         Height          =   285
         Left            =   4740
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   92995585
         CurrentDate     =   40584
      End
      Begin MSComCtl2.DTPicker dtpNFFim 
         Height          =   285
         Left            =   4740
         TabIndex        =   5
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   92995585
         CurrentDate     =   40584
      End
   End
   Begin VB.Frame frmVendedores 
      Caption         =   "Vendedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   60
      TabIndex        =   7
      Top             =   1560
      Width           =   13935
      Begin MSComctlLib.ListView lstV 
         Height          =   2655
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msfgNotas 
      Height          =   5355
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Duplo click para mudar a comissão..."
      Top             =   4740
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   9446
      _Version        =   393216
      Cols            =   9
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"formRHFuncionarioComissao.frx":0000
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Gerar Comissão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Relatorio"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir DANFe"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Trocar Comissionado"
            ImageIndex      =   14
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
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":00D3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":0525
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":083F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":10D1
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":2323
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":2BFD
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":348F
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":3D21
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":4F73
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":528D
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":55A7
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":599E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":7150
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioComissao.frx":76EA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formRHFuncionarioComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstNotas    As String
Dim vTotal      As String
Dim vComissao   As String
Dim NFe         As String
Dim lnVend      As Integer
Dim idFunc      As Integer
Private Function MountConsulta(iFun As Integer) As String
    Dim sSQL As String
    sSQL = "SELECT * FROM FaturamentoNFe " & _
           "WHERE ID_Empresa = " & ID_Empresa & " " & _
           "AND ide_tpNF=1 " & _
           "AND movFisco=" & chkMovFisco.Value & " " & _
           "AND movFinanceiro=" & chkMovFinanceiro.Value & " " & _
           "AND ide_natOP = 'VENDA' " & _
           "AND ger_Vendedor=" & iFun & " " & _
           "AND canc_nProt IS NULL"
    
    If optListagem(0).Value = True Then
            sSQL = sSQL & " AND ide_nNF >=" & IIf(Trim(txtNFIni.Text) = "", "0", txtNFIni.Text) & " AND ide_nNF <= " & IIf(Trim(txtNFFim.Text) = "", "0", txtNFFim.Text)
        ElseIf optListagem(1).Value = True Then
            sSQL = sSQL & " AND ide_dEmi >= '" & Format(dtpNFIni.Value, "yyyy-mm-dd") & "' AND ide_dEmi <= '" & Format(dtpNFFim.Value, "yyyy-mm-dd") & "'"
        Else
            MsgBox "Selecione uma opcao de listagem!", vbInformation, App.EXEName
            sSQL = ""
    End If
    MountConsulta = sSQL
End Function

Private Sub calcTotais(idFunc As Integer)
    Dim Rst         As Recordset
    Dim Rst1        As Recordset
    Dim sSQL        As String
    Dim vTot        As String
    Dim vCom        As String
    
    
    lstNotas = ""
    
    If idFunc = 0 Then
        MsgBox "Selecione um funcionario"
        Exit Sub
    End If
    
    sSQL = MountConsulta(idFunc)
    
    If Trim(sSQL) = "" Then Exit Sub
    
    vTotal = 0
    vComissao = 0
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            vTot = 0
            vCom = 0
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                vTot = Val(ChkVal(vTot, 0, cDecMoeda)) + Val(Rst.Fields("total_vProd"))
                sSQL = "SELECT * FROM FaturamentoNFeItens WHERE ID_Empresa = " & ID_Empresa & " AND IdNFe='" & Rst.Fields("IdNFe") & "'"
                Set Rst1 = RegistroBuscar(sSQL)
                If Rst1.BOF And Rst1.EOF Then
                        MsgBox "Erro ao localizar os itens da NFe n." & Rst.Fields("idNFe")
                    Else
                        lstNotas = lstNotas & Rst.Fields("ide_nNF") & "/"
                        Rst1.MoveFirst
                        Do Until Rst1.EOF
                            vCom = Val(ChkVal(vCom, 0, cDecMoeda)) + Val(IIf(IsNull(Rst1.Fields("comissao_vComissao")), 0, Rst1.Fields("comissao_vComissao")))
                            Rst1.MoveNext
                        Loop
                End If
                
                Rst.MoveNext
            Loop
    End If
    vTotal = vTot
    vComissao = vCom
    Rst.Close
End Sub



Private Sub LancarComissao(idFunc As Integer, sDescricao As String, sValor As String, sMesAno As String)
    Dim vReg(10)    As Variant
    Dim cReg        As Integer
    Dim sDoc        As String
    sDoc = Format(Date, "YYMMDD") & Format(Time, "HHMMSS")
    
    cReg = 0
    vReg(cReg) = Array("idFunc", idFunc, "N"): cReg = cReg + 1
    vReg(cReg) = Array("MesAno", sMesAno, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Doc", sDoc, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Descricao", sDescricao, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Valor", sValor, "S"): cReg = cReg + 1
    vReg(cReg) = Array("CD", "C", "S"): cReg = cReg + 1
    cReg = cReg - 1
    RegistroIncluir "RHFuncionarioFolhadePagamento", vReg, cReg
End Sub

Private Sub LoadNotasFiscais() '(idFunc As Integer)
    Dim Rst     As Recordset
    Dim Rst1    As Recordset
    Dim sSQL    As String
    Dim iSQL    As String
    Dim sCor    As Integer
    
    
    sSQL = MountConsulta(idFunc)
    If Trim(sSQL) = "" Then Exit Sub
    'sSQL = "SELECT * FROM FaturamentoNFe " & _
           "WHERE ID_Empresa = " & ID_Empresa & " " & _
           "AND ide_tpNF=1 " & _
           "AND movFisco=" & IIf(chkMovFisco.Value = True, 1, 0) & " " & _
           "AND movFinanceiro=" & IIf(chkMovFinanceiro.Value = True, 1, 0) & " " & _
           "AND ide_natOP = 'VENDA' " & _
           "AND ger_Vendedor=" & idFunc & " " & _
           "AND canc_nProt IS NULL"
    
    'If optListagem(0).Value = True Then
    '        iSQL = " AND FaturamentoNFe.ide_nNF >=" & IIf(Trim(txtNFIni.Text) = "", "0", txtNFIni.Text) & " AND FaturamentoNFe.ide_nNF <= " & IIf(Trim(txtNFFim.Text) = "", "0", txtNFFim.Text)
    '    ElseIf optListagem(1).Value = True Then
    '        iSQL = " AND FaturamentoNFe.ide_dEmi >= '" & Format(dtpNFIni.Value, "yyyy-mm-dd") & "' AND FaturamentoNFe.ide_dEmi <= '" & Format(dtpNFFim.Value, "yyyy-mm-dd") & "'"
    '    Else
    '        MsgBox "Selecione uma opcao de listagem"
    '        Exit Sub
    'End If
    'sSQL = sSQL & iSQL
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            msfgNotas.Rows = 1
            Exit Sub
        Else
            Rst.MoveFirst
    End If
    msfgNotas.FillStyle = flexFillRepeat
    msfgNotas.Rows = 1
    Do Until Rst.EOF
        sSQL = "SELECT * FROM FaturamentoNFeItens WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & Rst.Fields("IdNFe") & "'"
        Set Rst1 = RegistroBuscar(sSQL)
        If Rst1.BOF And Rst1.EOF Then
                MsgBox "Erro ao localizar os itens da NFe n." & Rst.Fields("idNFe")
                'Exit Do
            Else
                Rst1.MoveFirst
                With msfgNotas
                    sCor = IIf(sCor = 7, 15, 7)
                    Do Until Rst1.EOF
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = Rst1.Fields("id")
                        .TextMatrix(.Rows - 1, 1) = Rst1.Fields("idNFe")
                        .TextMatrix(.Rows - 1, 2) = Rst.Fields("ide_nNF")
                        .TextMatrix(.Rows - 1, 3) = Rst.Fields("ide_dEmi")
                        .TextMatrix(.Rows - 1, 4) = Rst.Fields("dest_xNome")
                        .TextMatrix(.Rows - 1, 5) = Rst1.Fields("det_xProd")
                        .TextMatrix(.Rows - 1, 6) = Rst1.Fields("det_vProd")
                        .TextMatrix(.Rows - 1, 7) = ChkVal(IIf(IsNull(Rst1.Fields("comissao_pComissao")), "0", Rst1.Fields("comissao_pComissao")), 0, 3)
                        .TextMatrix(.Rows - 1, 8) = ConvMoeda(IIf(IsNull(Rst1.Fields("comissao_vComissao")), "0", Rst1.Fields("comissao_vComissao"))) 'ConvMoeda(Val(PgDadosRhFuncionario(idFunc).Comissao) * Val(Rst1.Fields("det_vProd")) / 100)
                        DoEvents
                        .Row = .Rows - 1
                        '.RowSel = 1
                        .Col = 1
                        .ColSel = .Cols - 1
                        .CellBackColor = QBColor(sCor)
                        Rst1.MoveNext
                    Loop
                End With
        End If
        Rst.MoveNext
    Loop
End Sub


Private Sub MontarGrid()
    Dim ListH   As ColumnHeader
    
    
    
    lstV.View = lvwReport
    lstV.ColumnHeaders.Clear
    lstV.Checkboxes = True
    lstV.FullRowSelect = True
    lstV.GridLines = True
    lstV.LabelEdit = lvwManual
    
    Set ListH = lstV.ColumnHeaders.Add(1, , "Codigo")
    Set ListH = lstV.ColumnHeaders.Add(2, , "Nome")
    Set ListH = lstV.ColumnHeaders.Add(3, , "Salario")
    Set ListH = lstV.ColumnHeaders.Add(4, , "Total Vendido")
    Set ListH = lstV.ColumnHeaders.Add(5, , "Valor da Comissão")
    Set ListH = lstV.ColumnHeaders.Add(6, , "Num. Documento")
    Set ListH = lstV.ColumnHeaders.Add(7, , "Nota(s) Fiscal(is)")
End Sub
Private Sub LoadFuncionarios()
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim Lst     As ListItem
        
    lstV.ListItems.Clear
    sSQL = "SELECT * FROM RHFuncionarioCadastro WHERE id_empresa=" & ID_Empresa & " AND Comissao IS NOT NULL AND Comissao <> 0 ORDER BY xNome"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
            idFunc = Left("0000", 4 - Len(Rst.Fields("Id"))) & Rst.Fields("Id")
            calcTotais (idFunc)
            Set Lst = lstV.ListItems.Add(, , Left("0000", 4 - Len(idFunc)) & idFunc)
                Lst.SubItems(1) = IIf(IsNull(Rst.Fields("xNome")), "", Rst.Fields("xNome"))
                Lst.SubItems(2) = ConvMoeda(IIf(IsNull(Rst.Fields("Salario")), 0, Rst.Fields("Salario")))
                Lst.SubItems(3) = ConvMoeda(vTotal) 'ConvMoeda(calcTotalVendido(idFunc))
                Lst.SubItems(4) = ConvMoeda(vComissao)
                Lst.SubItems(5) = Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & Left("0000", 4 - Len(idFunc)) & idFunc
                Lst.SubItems(6) = lstNotas
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    
    
    
End Sub


Private Sub MontarTabelaTemporaria()

    BD.Execute "DROP TABLE IF EXISTS tmp_comissao"
    BD.Execute "CREATE TABLE IF NOT EXISTS tmp_comissao" & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "Func VARCHAR(120) default Null," & _
               "Nome VARCHAR(100) default Null," & _
               "Dt DATE default Null," & _
               "nNF VARCHAR(100) default Null," & _
               "vTotal VARCHAR(100) default Null," & _
               "vProd VARCHAR(100) default Null," & _
               "pComi VARCHAR(100) default Null," & _
               "vComi VARCHAR(100) default Null," & _
               "PRIMARY KEY (Id))"
End Sub

Private Sub ImprimirRelatorio()
    Dim Rst         As Recordset
    Dim Rst1        As Recordset
    Dim Rst2        As Recordset
    Dim sSQL        As String
    Dim iSQL        As String
    Dim vReg(1000)  As Variant
    Dim cReg        As Integer
    Dim pComi       As String
    Dim vComi       As String
    
    Dim TotvProd    As String
    Dim TotvComi    As String
    Dim Periodo     As String
    
    Dim idFunc      As Integer
    Dim i           As Integer
    
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If
    
    
    
    MontarTabelaTemporaria
    
    For i = 1 To lstV.ListItems.Count
        If lstV.ListItems.Item(i).Checked = True Then
            idFunc = lstV.ListItems(i).Text
            
            'sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND ide_tpNF = 1 AND ide_natOP = 'VENDA' AND ger_Vendedor=" & idFunc & " AND canc_nProt IS NULL"
            sSQL = MountConsulta(lstV.ListItems(i).Text)
            If Trim(sSQL) = "" Then Exit Sub
            
            If optListagem(0).Value = True Then
                    iSQL = " AND FaturamentoNFe.ide_nNF >=" & IIf(Trim(txtNFIni.Text) = "", "0", txtNFIni.Text) & " AND FaturamentoNFe.ide_nNF <= " & IIf(Trim(txtNFFim.Text) = "", "0", txtNFFim.Text)
                    Periodo = "De: " & txtNFIni.Text & "   Até: " & txtNFFim.Text
                ElseIf optListagem(1).Value = True Then
                    iSQL = " AND FaturamentoNFe.ide_dEmi >= '" & Format(dtpNFIni.Value, "yyyy-mm-dd") & "' AND FaturamentoNFe.ide_dEmi <= '" & Format(dtpNFFim.Value, "yyyy-mm-dd") & "'"
                    Periodo = "De: " & dtpNFIni.Value & "   Até: " & dtpNFFim.Value
                Else
                    MsgBox "Selecione uma opcao de listagem"
                    Periodo = ""
                    Exit Sub
            End If
            sSQL = sSQL & iSQL
    
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                MsgBox "Nenhuma NOTA FISCAL registrada no intervalo selecionado.", vbInformation, "Aviso"
                Exit Sub
            End If
        TotvProd = 0
        TotvComi = 0
        Do Until Rst.EOF
            sSQL = "SELECT * FROM FaturamentoNFeItens WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & Rst.Fields("IdNFe") & "'"
            Set Rst1 = RegistroBuscar(sSQL)
            If Rst1.BOF And Rst1.EOF Then
                    'MsgBox "Erro ao localizar itens da NF-e"
                    'Exit Do
                Else
                    Rst1.MoveFirst
                    vComi = 0
                    pComi = 0
                    Do Until Rst1.EOF
                        vComi = ChkVal(Val(vComi) + Val(ChkVal(IIf(IsNull(Rst1.Fields("comissao_vComissao")), 0, Rst1.Fields("comissao_vComissao")), 0, cDecMoeda)), 0, cDecMoeda)
                        pComi = Val(ChkVal(pComi, 0, 3)) + Val(ChkVal(Rst1.Fields("comissao_pComissao"), 0, 3))
                        Rst1.MoveNext
                    Loop
                    pComi = ChkVal(Val(ChkVal(pComi, 0, 3)) / Rst1.RecordCount, 0, 3)
            End If
            cReg = 0
            'pComi = ChkVal((Val(vComi) * 100) / Val(Rst.Fields("total_vprod")), 0, 3)
            vReg(cReg) = Array("Func", PgDadosRhFuncionario(idFunc).Nome, "S"): cReg = cReg + 1
            vReg(cReg) = Array("Nome", Rst.Fields("dest_xNome"), "S"): cReg = cReg + 1
            vReg(cReg) = Array("Dt", Rst.Fields("ide_dEmi"), "D"): cReg = cReg + 1
            vReg(cReg) = Array("nNF", Rst.Fields("ide_nNF"), "S"): cReg = cReg + 1
            vReg(cReg) = Array("vProd", Rst.Fields("total_vProd"), "S"): cReg = cReg + 1
            vReg(cReg) = Array("pComi", pComi, "S"): cReg = cReg + 1
            vReg(cReg) = Array("vComi", vComi, "S") ': cReg = cReg + 1
            
            TotvComi = Val(ChkVal(TotvComi, 0, cDecMoeda)) + Val(ChkVal(vComi, 0, cDecMoeda))
            TotvProd = Val(ChkVal(TotvProd, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vProd"), 0, cDecMoeda))
            
            If RegistroIncluir("tmp_comissao", vReg, cReg) = 0 Then
                MsgBox "Erro ao incluir comissao", vbInformation, "Aviso"
            End If
            Rst.MoveNext
        Loop
        Rst.Close
        Rst1.Close
        
        'Impressao de relatorio
        sSQL = "SELECT * FROM tmp_comissao WHERE ID_Empresa = " & ID_Empresa & " AND Func = '" & PgDadosRhFuncionario(idFunc).Nome & "'"
        Set Rst2 = RegistroBuscar(sSQL)
        If Rst2.BOF And Rst2.EOF Then
                MsgBox "Nenhum Registro "
            Else
                Rst2.MoveFirst
                Set rptListaComissao.DataSource = Rst2.DataSource
                rptListaComissao.Sections("Section2").Controls.Item("lblFunc").Caption = PgDadosRhFuncionario(idFunc).Nome
                rptListaComissao.Sections("Section2").Controls.Item("lblPeriodo").Caption = Periodo
                rptListaComissao.Sections("Section5").Controls.Item("lblTotProd").Caption = ConvMoeda(TotvProd)
                rptListaComissao.Sections("Section5").Controls.Item("lblTotComi").Caption = ConvMoeda(TotvComi)
                rptListaComissao.Show 1
        End If
        Rst2.Close
    End If
   Next
   
   '------------- LISTAR TODAS AS COMISSOES POR VENDEDOR ---------
    'DEU ERRO POIS ACJO QUE MYSQL NAO ACEITA O COMANDO SHAPE
    'Dim Rst2 As Recordset
    'Dim cmd As New ADODB.Command
    'sSQL = "SHAPE {SELECT Nome, dt " & _
           "FROM tmp_Comissao} " & _
           "AS Command1 COMPUTE Command1 BY Nome"
    'With cmd
    '    .ActiveConnection = BD
    '    .CommandType = adCmdText
    '    .CommandText = sSQL
    '    Set Rst2 = .Execute
    'End With
End Sub
Private Sub GerarComissao()
    Dim i       As Integer
    Dim func    As Integer
    Dim nFat    As String
    Dim vlComi  As String
    Dim Nts     As String
    Dim MesAno  As String
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    If MsgBox("Deseja gerar comissão no contas a pagar agora?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    MesAno = Format(Date, "MM/YYYY")
    
    MesAno = InputBox("Informe o Mês/Ano (MM/YYYY) referente as comissões.", App.EXEName, MesAno)
    
    If ValidarMesAno(MesAno) = False Then
        'MsgBox "Mês/Ano invalido! Comissões não geradas!", vbCritical, App.EXEName
        Exit Sub
    End If
    
    
    For i = 1 To lstV.ListItems.Count
        If lstV.ListItems.Item(i).Checked = True Then
            func = lstV.ListItems(i).Text
            nFat = lstV.ListItems(i).SubItems(5)
            vlComi = ChkVal(lstV.ListItems(i).SubItems(4), 0, 2)
            Nts = lstV.ListItems(i).SubItems(6)
            
            If vlComi = 0 Then
                    MsgBox "Valor da comissão do funcionario " & PgDadosRhFuncionario(func).Nome & " não pode ser igual a zero (R$ 0,00). Favor Verificar!", vbInformation, "Aviso"
                Else
                    LancarComissao func, "Comissão Ref.: " & Nts, vlComi, MesAno
                    'Call MovimentarContasPagarReceber("P", Date, nFat, vlComi, "RHFuncionarioCadastro", func, PgDadosRhFuncionario(func).Nome, _
                                    PgDadosRhFuncionario(func).CPF, PgDadosConfig.RHConta, PgDadosConfig.RHCentroCustos, PgDadosConfig.RHDocumento, PgDadosConfig.RHPlanoContas, "", "", Date, nFat & "-1/1", "0", _
                                    "0", "0", "0", "0", "0", "0", vlComi, "NFs:" & Nts)
            End If
            
        End If
    Next
    MsgBox "Comissão gerada com sucesso!", vbInformation, "Aviso"
End Sub
Private Function ValidarMesAno(sTexto As String) As Boolean
    On Error GoTo TrtErrMA
    If Trim(sTexto) < 7 Then
        MsgBox "Favor informar MM/YYYY.", vbCritical, App.EXEName
        ValidarMesAno = False
        Exit Function
    End If
    If Mid(sTexto, 1, InStr(sTexto, "/") - 1) > 12 Then
        MsgBox "Favor informar MM/YYYY.", vbCritical, App.EXEName
        ValidarMesAno = False
        Exit Function
    End If
    If Len(Mid(sTexto, InStr(sTexto, "/") + 1)) <> 4 Then
        MsgBox "Favor informar MM/YYYY.", vbCritical, App.EXEName
        ValidarMesAno = False
        Exit Function
    End If
    
    ValidarMesAno = True
    Exit Function
TrtErrMA:
    MsgBox "Favor informar MM/YYYY.", vbCritical, App.EXEName
    ValidarMesAno = False
End Function

Private Sub AtualizarLista()
    LoadFuncionarios
    msfgNotas.Rows = 1
End Sub






Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    dtpNFIni.Value = Date
    dtpNFFim.Value = Date
    txtNFIni.Text = 0
    txtNFFim.Text = 0
    chkMovFisco.Value = 1
    chkMovFinanceiro.Value = 1
    optListagem_Click (0)
    MontarGrid
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    frmVendedores.Width = Me.Width - 300
    lstV.Width = frmVendedores.Width - 300
    
    msfgNotas.Left = frmVendedores.Left
    msfgNotas.Width = frmVendedores.Width
    msfgNotas.Height = Me.Height - (msfgNotas.Top + 600)
End Sub
Private Sub msfgNotas_Click()
    If UCase(msfgNotas.TextMatrix(msfgNotas.Row, 1)) = "NFE" Then
        NFe = ""
        Exit Sub
    End If
    NFe = msfgNotas.TextMatrix(msfgNotas.Row, 1)
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualizar"
            AtualizarLista
        Case "Gerar Comissão"
            GerarComissao
        Case "Imprimir Relatorio"
            ImprimirRelatorio
        Case "Imprimir DANFe"
            ImprimirCopiaNota
        Case "Trocar Comissionado"
            TrocarComissionado
    End Select
End Sub
Private Sub ImprimirCopiaNota()
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If

    If Trim(NFe) = "" Then
        MsgBox "Selecione uma Nota Fiscal.", vbInformation, "Aviso"
        Exit Sub
    End If
    ImprimirDANFE (NFe)
    
End Sub
Private Sub TrocarComissionado()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If Trim(NFe) = "" Then
        MsgBox "Selecione uma Nota Fiscal.", vbInformation, "Aviso"
        Exit Sub
    End If
    formRHFuncionarioTrocarComissao.CarregarDadosNFe (NFe)
    AtualizarLista
End Sub
Private Sub lstV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    idFunc = lstV.SelectedItem
    LoadNotasFiscais '(lstV.SelectedItem)
End Sub


Private Sub msfgnotas_DblClick()
    Dim comi    As String
    Dim a(1)    As Variant
    Dim idFunc  As Integer
    
    idFunc = lstV.SelectedItem
    
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    
    comi = InputBox("Informe o novo percentual de comissão.", , msfgNotas.TextMatrix(msfgNotas.Row, 7))
    If Trim(comi) <> "" And IsNumeric(comi) = True Then
        'msfgNotas.TextMatrix(msfgNotas.Row, 7) = ChkVal(comi, 0, 3)
        a(0) = Array("comissao_pComissao", ChkVal(comi, 0, 3), "S")
        a(1) = Array("comissao_vComissao", ChkVal((Val(ChkVal(comi, 0, 3)) * Val(msfgNotas.TextMatrix(msfgNotas.Row, 6)) / 100), 0, cDecMoeda), "S")
        If RegistroAlterar("FaturamentoNFeItens", a, 1, "Id=" & msfgNotas.TextMatrix(msfgNotas.Row, 0)) = False Then
            MsgBox "Erro ao atualizar comissao. Favor verificar", vbInformation, "Aviso"
        End If
        'LoadFuncionarios
        LoadNotasFiscais '(idFunc)
    End If

End Sub



Private Sub optListagem_Click(Index As Integer)
    If optListagem(0).Value = True Then
            txtNFIni.Enabled = True
            txtNFFim.Enabled = True
            dtpNFIni.Enabled = False
            dtpNFFim.Enabled = False
        Else
            txtNFIni.Enabled = False
            txtNFFim.Enabled = False
            dtpNFIni.Enabled = True
            dtpNFFim.Enabled = True
    End If
End Sub

Private Sub txtNFFim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub


Private Sub txtNFIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub



