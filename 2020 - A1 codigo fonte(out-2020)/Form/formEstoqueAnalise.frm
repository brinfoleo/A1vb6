VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formEstoqueAnalise 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estoque - Analise"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7500
   Begin VB.Frame Frame3 
      Caption         =   "Grupo e Subgrupo"
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
      TabIndex        =   13
      Top             =   3480
      Width           =   7335
      Begin VB.ComboBox cboSubgrupo 
         Height          =   315
         Left            =   4500
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   300
         Width           =   2595
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   300
         Width           =   2595
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Subgrupo:"
         Height          =   195
         Left            =   3720
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Perido"
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
      TabIndex        =   7
      Top             =   2640
      Width           =   7335
      Begin VB.Frame frameAnalise 
         Caption         =   "Analise"
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
         Left            =   5100
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   2235
         Begin VB.OptionButton optAnalise 
            Caption         =   "Analítico"
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   21
            Top             =   480
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton optAnalise 
            Caption         =   "Sintético"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   20
            Top             =   240
            Width           =   1155
         End
      End
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   315
         Left            =   900
         TabIndex        =   8
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   114425857
         CurrentDate     =   40665
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   315
         Left            =   3120
         TabIndex        =   9
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   114753537
         CurrentDate     =   40665
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Final:"
         Height          =   195
         Left            =   2700
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   7335
      Begin VB.Frame Frame4 
         Caption         =   "Ordenar por:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4740
         TabIndex        =   22
         Top             =   0
         Width           =   2595
         Begin VB.ComboBox cboOrdenar 
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.OptionButton optTpRelatorio 
         Caption         =   "Analise de saldo no periodo"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   3735
      End
      Begin VB.OptionButton optTpRelatorio 
         Caption         =   "Listagem de Saldo por Grupo e/ou Subgrupo"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1380
         Width           =   3735
      End
      Begin VB.OptionButton optTpRelatorio 
         Caption         =   "Listagem de Saldo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.OptionButton optTpRelatorio 
         Caption         =   "Listagem de Saldo, custo e preço de tabela"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3735
      End
      Begin VB.OptionButton optTpRelatorio 
         Caption         =   "Listagem de saldo negativo"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   780
         Width           =   3735
      End
      Begin VB.OptionButton optTpRelatorio 
         Caption         =   "Balanço de estoque pelo custo do produto "
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   3735
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   3540
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2760
         Top             =   60
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
               Picture         =   "formEstoqueAnalise.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueAnalise.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueAnalise.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueAnalise.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueAnalise.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueAnalise.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueAnalise.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueAnalise.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueAnalise.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueAnalise.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueAnalise.frx":54D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueAnalise.frx":58CB
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formEstoqueAnalise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TpRelatorio As Integer

Private Sub cboGrupo_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    cboGrupo.Clear
    sSQL = "SELECT * FROM EstoqueGrupos ORDER BY Descricao"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboGrupo.AddItem ZE(Trim(Rst.Fields("Id")), 5) & " - " & Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub cboOrdenar_DropDown()
    With cboOrdenar
        .Clear
        .AddItem "ID"
        .AddItem "Referencia"
        .AddItem "Descrição"
        
    End With
    
End Sub

Private Sub cbosubGrupo_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    cboSubgrupo.Clear
    sSQL = "SELECT * FROM EstoqueSubGrupo ORDER BY Descricao"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboSubgrupo.AddItem ZE(Trim(Rst.Fields("Id")), 5) & " - " & Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub
Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FiltrosDoForm False
End Sub
Private Sub FiltrosDoForm(op As Boolean)
    dtpIni.Value = Date
    dtpFin.Value = Date
    dtpIni.Enabled = op
    dtpFin.Enabled = op
    cboGrupo.Enabled = op
    cboSubgrupo.Enabled = op
    
    frameAnalise.Visible = False
End Sub

Private Sub optAnalise_Click(Index As Integer)

    Select Case Index
        Case 0
            cboGrupo.Enabled = False
            cboSubgrupo.Enabled = False
        Case 1

            cboGrupo.Enabled = True
            cboSubgrupo.Enabled = True

        Case Else

    End Select
End Sub

Private Sub optTpRelatorio_Click(Index As Integer)
    FiltrosDoForm False
    TpRelatorio = Index
    
    Select Case TpRelatorio
        Case 4
            cboGrupo.Enabled = True
            cboSubgrupo.Enabled = True
        Case 5
            dtpIni.Enabled = True
            dtpFin.Enabled = True
            cboGrupo.Enabled = True
            cboSubgrupo.Enabled = True
            frameAnalise.Visible = True
    End Select
    
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case TpRelatorio
        Case 0
            Rpt_000
        Case 1
            Rpt_001
        Case 2
            Rpt_002
        Case 3
            Rpt_003
        Case 4
            Rpt_004
        Case 5
            Rpt_005
    End Select
End Sub
Private Function OrderBy() As String
    Select Case LCase(cboOrdenar.Text)
        Case "id"
            OrderBy = "ORDER BY id"
        Case "descrição"
            OrderBy = "ORDER BY descricao"
        Case "referencia"
            OrderBy = "ORDER BY referencia"
        Case Else
            OrderBy = ""
    End Select
    
End Function
Private Sub Rpt_000()
    'Listagem para conferencia de Saldo
    'Data: 29/06/2011
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    sSQL = "SELECT * FROM EstoqueProduto " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND Deposito = " & ID_Deposito & _
           " AND status = 'ATIVO' " & _
           OrderBy
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Set rptListaEstoque.DataSource = Rst.DataSource
            rptListaEstoque.Sections("Section2").Controls.Item("lbltitulo").Caption = "Listagem de Estoque"
            rptListaEstoque.Sections("Section2").Controls.Item("lblCusto").Visible = False
            rptListaEstoque.Sections("Section1").Controls.Item("txtCusto").Visible = False
            rptListaEstoque.Sections("Section2").Controls.Item("lblPreco").Visible = False
            rptListaEstoque.Sections("Section1").Controls.Item("txtPreco").Visible = False
            
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot1").Visible = False
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot2").Visible = False
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot3").Visible = False
            
            rptListaEstoque.Show 1
    End If
End Sub
Private Sub Rpt_001()
    'Listagem para conferencia de Saldo, preço de custo e tabela
    'Data: 04/07/2011
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    sSQL = "SELECT * FROM EstoqueProduto " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND Deposito = " & ID_Deposito & _
           " AND status = 'ATIVO' " & _
           OrderBy
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            '******************************************************************************************************************
            '**** Data de Criacao: 04/07/2011
            '**** Objetivo: Atualizar as casas decimais de todo o estoque
            '******************************************************************************************************************
            'Dim vDados(100) As Variant
            'Dim cDados      As Integer
            'cDados = 0
            'Do Until Rst.EOF
            '    vDados(cDados) = Array("Custo", ChkVal(IIf(IsNull(Rst.Fields("Custo")), "0", Rst.Fields("Custo")), 0, cDecMoeda), "S"): cDados = cDados + 1
            '    vDados(cDados) = Array("Preco", ChkVal(IIf(IsNull(Rst.Fields("Preco")), "0", Rst.Fields("Preco")), 0, cDecMoeda), "S"): cDados = cDados + 1
            '    vDados(cDados) = Array("Saldo", ChkVal(IIf(IsNull(Rst.Fields("Saldo")), "0", Rst.Fields("Saldo")), 0, cDecQtd), "S") ': cDados = cDados + 1
            '    RegistroAlterar "EstoqueProduto", vDados, cDados, "ID = " & Rst.Fields("ID")
            '    cDados = 0
            '    Rst.MoveNext
            'Loop
            '******************************************************************************************************************
            Set rptListaEstoque.DataSource = Rst.DataSource
            rptListaEstoque.Sections("Section2").Controls.Item("lbltitulo").Caption = "Listagem de Estoque"
            rptListaEstoque.Sections("Section2").Controls.Item("lblCusto").Visible = True
            rptListaEstoque.Sections("Section1").Controls.Item("txtCusto").Visible = True
            rptListaEstoque.Sections("Section2").Controls.Item("lblPreco").Visible = True
            rptListaEstoque.Sections("Section1").Controls.Item("txtPreco").Visible = True
            
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot1").Visible = False
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot2").Visible = False
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot3").Visible = False
            
            rptListaEstoque.Show 1
    End If
End Sub
Private Sub Rpt_002()
    'Listagem de Saldo NEGATIVO
    'Data: 04/07/2011
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    sSQL = "SELECT * FROM EstoqueProduto " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND Deposito = " & ID_Deposito & _
           " AND status = 'ATIVO' " & _
           " AND Saldo < 0 " & OrderBy
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Set rptListaEstoque.DataSource = Rst.DataSource
            rptListaEstoque.Sections("Section2").Controls.Item("lbltitulo").Caption = "Listagem de Estoque com saldo NEGATIVO"
            rptListaEstoque.Sections("Section2").Controls.Item("lblCusto").Visible = True
            rptListaEstoque.Sections("Section1").Controls.Item("txtCusto").Visible = True
            rptListaEstoque.Sections("Section2").Controls.Item("lblPreco").Visible = True
            rptListaEstoque.Sections("Section1").Controls.Item("txtPreco").Visible = True
            
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot1").Visible = False
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot2").Visible = False
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot3").Visible = False
            
            rptListaEstoque.Show 1
    End If
End Sub
Private Sub Rpt_003()
    
    '*************************************************************
    '*** Obj.: Listagem de BALANCO DO ESTOQUE PELO CUSTO
    '*************************************************************
    On Error Resume Next
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim vCusto  As String
    Dim vSaldo  As String
    Dim qSaldo  As String
    
    
    If MontarTabelaTemporaria_003 = False Then Exit Sub
    
    sSQL = "SELECT * FROM tmp_estoqueanalise ORDER BY SubGrupo, Grupo, ref"
    Set Rst = RegistroBuscar(sSQL)
'*****************************************************************************************************************
'*** Funcao modificada
'*****************************************************************************************************************

    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            'Variaveis
            Dim nmGrupo     As String 'Nome grupo

            Dim qGrupo      As String 'Quantidade do GRUPO
            Dim vGrupo      As String 'Quantidade do GRUPO

            Dim qTotal      As String 'Quantidade total no estoque
            Dim vTotal      As String 'Valor total no estoque

            Dim c           As Integer 'Conta a quant reg por folha
            Dim pg          As Integer 'Pagina

            Rst.MoveFirst
            'Seleciona o tipo de impressora
            'formImpressoraSelecionar.SelecionarImpressora
            If formImpressoraSelecionar.SelecionarImpressora = False Then Exit Sub
            
            nmGrupo = ""
            vSaldo = 0
            qSaldo = 0
            qTotal = 0
            vTotal = 0
            pg = 1
            c = 80
            nmGrupo = Trim(Rst.Fields("SubGrupo")) & " " & Trim(Rst.Fields("Grupo"))
            Do Until Rst.EOF
                '*** Gera um cab para a nova pagina ********
                If nmGrupo <> Trim(Rst.Fields("SubGrupo")) & " " & Trim(Rst.Fields("Grupo")) Or c = 80 Then

                    'ZERA OS SOMATORIOS TOTAIS QUANDO O GRUPO MUDAR
                    If nmGrupo <> Trim(Rst.Fields("SubGrupo")) & " " & Trim(Rst.Fields("Grupo")) Then
                        'qTotal = ChkVal(Val(qTotal) + Val(qGrupo), 0, cDecQtd)
                        Printer.Print Tab(5); String(180, "-")
                        Printer.Print " "
                        Printer.Print Tab(10); "Total de " & nmGrupo & "  "; Tab(80); qGrupo; Tab(120); ConvMoeda(vGrupo)
                        pg = pg + 1
                        qGrupo = 0
                        vGrupo = 0
                    End If
                    'Proxima pagina
                    If pg <> 1 Then
                        Printer.NewPage
                        'pg = pg + 1
                    End If

                    'Cab
                    Printer.FontSize = 8
                    Printer.Print Tab(5); " "
                    Printer.Print Tab(5); String(102, "=")
                    Printer.FontBold = True
                    Printer.Print Tab(7); Rst.Fields("SubGrupo") & " " & Rst.Fields("Grupo"); Tab(108); Date & " pag.: " & Left("000", 3 - Len(Trim(pg))) & Trim(pg) 'Printer.Page
                    Printer.FontBold = False
                    Printer.Print Tab(5); String(102, "=")
                    Printer.Print Tab(10); "Descrição"; Tab(80); "Saldo"; Tab(100); "Custo", Tab(120); "Total"
                    Printer.Print Tab(5); String(180, "-")
                    c = 0

                    nmGrupo = Trim(Rst.Fields("SubGrupo")) & " " & Trim(Rst.Fields("Grupo"))
                End If
                '*******************************************
                c = c + 1
                Printer.FontSize = 8

                qSaldo = ChkVal(Rst.Fields("Saldo"), 0, cDecQtd)
                vCusto = ChkVal(Rst.Fields("Custo"), 0, cDecMoeda)
                vSaldo = ChkVal(Val(qSaldo) * Val(vCusto), 0, cDecMoeda)
                vSaldo = ChkVal(IIf(Val(vSaldo) < 0, 0, vSaldo), 0, cDecMoeda)

                qTotal = ChkVal(Val(qTotal) + Val(qSaldo), 0, cDecQtd)
                vTotal = ChkVal(Val(vTotal) + Val(vSaldo), 0, cDecMoeda)
                
                qGrupo = ChkVal(Val(qGrupo) + Val(qSaldo), 0, cDecQtd)
                vGrupo = ChkVal(Val(vGrupo) + Val(vSaldo), 0, cDecMoeda)
            '*****************************************************************************
            '*** IMPRIME OS ITENS DA NF
            '*** Printer.Print Tab(5); ZE(Rst.fields("idProd"), 4); Tab(12); Rst.fields("Descricao") & IIf(Rst.fields("balanco") <> 0, " (b)", ""); Tab(80); qSaldo & "/" & Rst.fields("Unidade"); Tab(100); vCusto; Tab(120); Trim(vSaldo)
            '*****************************************************************************
                Printer.Print Tab(5); ZE(Rst.Fields("idProd"), 4); Tab(12); Rst.Fields("Descricao"); Tab(80); qSaldo & "/" & Rst.Fields("Unidade"); Tab(100); vCusto; Tab(120); Trim(vSaldo)
                'grvReg "estoque.txt", nmGrupo & ";" & Rst.Fields("Descricao") & ";" & Val(qSaldo) & ";" & Rst.Fields("Unidade") & ";" & Val(vCusto) & ";" & Val(Trim(vSaldo))
               Rst.MoveNext
            Loop
            Printer.Print Tab(5); String(180, "-")
            Printer.Print " "
            Printer.Print Tab(10); "Total de " & nmGrupo & "  "; Tab(80); qGrupo; Tab(120); ConvMoeda(vGrupo)
            Printer.Print Tab(5); " "
            Printer.Print Tab(5); String(102, "=")
            Printer.Print Tab(10); "Quantidade Total: " & qTotal
            Printer.Print Tab(10); "Valor Total..........: " & ConvMoeda(vTotal)
            Printer.Print Tab(5); String(102, "=")
            'Printer.EndDoc
    End If
    '######################################################################################################################
    '### Totalizar por grupo e sub grupo
    '######################################################################################################################
    Rst.MoveFirst
    'nmGrupo = ""
    Printer.NewPage
    Printer.FontSize = 8
    Printer.Print Tab(5); " "
    Printer.Print Tab(5); String(102, "=")
    Printer.FontBold = True
    Printer.Print Tab(7); "TOTAL POR GRUPOS"; Tab(108); Date & " pag.: " & Left("000", 3 - Len(Trim("1"))) & Trim("1") 'Printer.Page
    Printer.FontBold = False
    Printer.Print Tab(5); String(102, "=")

    Printer.Print Tab(5); "Descrição"; Tab(50); "Saldo"; Tab(70); "Total"
    Printer.Print Tab(5); String(180, "-")
    nmGrupo = Trim(Rst.Fields("SubGrupo")) & " " & Trim(Rst.Fields("Grupo"))
    qSaldo = "0"
    vCusto = "0"
    vSaldo = "0"
    qGrupo = "0"
    vGrupo = "0"
    qTotal = "0"
    vTotal = "0"
    Do Until Rst.EOF
        qSaldo = ChkVal(Rst.Fields("Saldo"), 0, cDecQtd)
        vCusto = ChkVal(Rst.Fields("Custo"), 0, cDecMoeda)
        vSaldo = ChkVal(Val(qSaldo) * Val(vCusto), 0, cDecMoeda)

        qTotal = ChkVal(Val(qTotal) + Val(qSaldo), 0, cDecQtd)
        vTotal = ChkVal(Val(vTotal) + Val(vSaldo), 0, cDecMoeda)
        qGrupo = ChkVal(Val(qGrupo) + Val(qSaldo), 0, cDecQtd)
        vGrupo = ChkVal(Val(vGrupo) + Val(vSaldo), 0, cDecMoeda)
        '************************************************************************************
        
        Rst.MoveNext
        
        If nmGrupo <> Trim(Rst.Fields("SubGrupo")) & " " & Trim(Rst.Fields("Grupo")) Then
            Printer.Print Tab(5); nmGrupo; Tab(45); "|"; Tab(50); qGrupo; Tab(65); "|"; Tab(70); ConvMoeda(vGrupo)
            qGrupo = "0"
            vGrupo = "0"
            nmGrupo = Trim(Rst.Fields("SubGrupo")) & " " & Trim(Rst.Fields("Grupo"))
        End If
    Loop
    Printer.Print " "
    Printer.FontBold = True
    Printer.Print Tab(50); qTotal; Tab(70); ConvMoeda(vTotal)
    Printer.FontBold = False
    Printer.EndDoc
'*************************************************************************************************************************************
'*** Data: 27.07.2011
'*** Obj.: Criara uma função que imprima de forma hierarquica
'*************************************************************************************************************************************
'
'    If Rst.BOF And Rst.EOF Then
'            MsgBox "Nenhum registro encontrado.", vbInformation, "Aviso"
'            Rst.Close
'            Exit Sub
'        Else
'            Rst.MoveFirst
'            vSaldo = 0
'            qSaldo = 0
'            Do Until Rst.EOF
'                Status (Rst.RecordCount)
'                vCusto = Val(ChkVal(vCusto, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("custo"), 0, cDecMoeda))
'                qSaldo = Val(ChkVal(qSaldo, 0, cDecQtd)) + Val(ChkVal(Rst.Fields("Saldo"), 0, cDecQtd))
'                vSaldo = Val(ChkVal(vSaldo, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("Preco"), 0, cDecMoeda))
'                Rst.MoveNext
'            Loop
'            Set rptListaEstoque.DataSource = Rst.DataSource
'            rptListaEstoque.Sections("Section2").Controls.Item("lbltitulo").Caption = "Listagem de Estoque para BALANÇO"
'            rptListaEstoque.Sections("Section2").Controls.Item("lblCusto").Visible = True
'            rptListaEstoque.Sections("Section1").Controls.Item("txtCusto").Visible = True
'            rptListaEstoque.Sections("Section2").Controls.Item("lblPreco").Visible = True
'            rptListaEstoque.Sections("Section1").Controls.Item("txtPreco").Visible = True
'
'            rptListaEstoque.Sections("Section5").Controls.Item("lblTot1").Visible = True
'            rptListaEstoque.Sections("Section5").Controls.Item("lblTot1").Caption = ChkVal(qSaldo, 0, cDecQtd)
'
'            rptListaEstoque.Sections("Section5").Controls.Item("lblTot2").Visible = True
'            rptListaEstoque.Sections("Section5").Controls.Item("lblTot2").Caption = ConvMoeda(vCusto)
'
'            rptListaEstoque.Sections("Section5").Controls.Item("lblTot3").Visible = True
'            rptListaEstoque.Sections("Section5").Controls.Item("lblTot3").Caption = ConvMoeda(vSaldo)
'
'            rptListaEstoque.Show 1
'
'    End If
End Sub
Private Sub Rpt_004()
    'Listagem para conferencia de Saldo por grupo e sub grupo
    'Data: 13/01/2012
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim iG      As String
    Dim iSG     As String
    Dim saldoF  As String
    
    If Trim(cboGrupo.Text) = "" And Trim(cboSubgrupo.Text) = "" Then
        MsgBox "Favor selecionar um Grupo / Subgrupo!", vbInformation, "Aviso"
        Exit Sub
    End If
    
    iG = IIf(Trim(cboGrupo.Text) = "", "", "grupo=" & Left(cboGrupo.Text, 5))
    iSG = IIf(Trim(cboSubgrupo.Text) = "", "", "subgrupo=" & Left(cboSubgrupo.Text, 5))
    
    sSQL = "SELECT * FROM EstoqueProduto " & _
           "WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & _
           " AND status = 'ATIVO' " & _
           IIf(Trim(iG) = "", "", " AND " & iG) & _
           IIf(Trim(iSG) = "", "", " AND " & iSG) & _
           " " & OrderBy
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            
            Rst.MoveFirst
            Do Until Rst.EOF
                If Val(ChkVal(Rst.Fields("Saldo"), 0, cDecQtd)) >= 0 Then
                    saldoF = Val(ChkVal(saldoF, 0, cDecQtd)) + Val(ChkVal(Rst.Fields("Saldo"), 0, cDecQtd))
                End If
                Rst.MoveNext
            Loop
            Rst.MoveFirst
            'Set rptListaEstoque.Sections("Section1").Controls.Item("Text3").DataFormat = "#.##0,000"
            Set rptListaEstoque.DataSource = Rst.DataSource
            rptListaEstoque.Sections("Section2").Controls.Item("lbltitulo").Caption = "Listagem de Estoque: " & Trim(Mid(cboGrupo.Text, 8, Len(cboGrupo.Text))) & " / " & Trim(Mid(cboSubgrupo.Text, 8, Len(cboSubgrupo.Text)))
            
            
            rptListaEstoque.Sections("Section2").Controls.Item("lblCusto").Visible = True
            rptListaEstoque.Sections("Section2").Controls.Item("lblCusto").Caption = "Balanço"
            
            rptListaEstoque.Sections("Section1").Controls.Item("txtCusto").Visible = True
            rptListaEstoque.Sections("Section1").Controls.Item("txtCusto").DataField = "IncluirBalanco"
            
            rptListaEstoque.Sections("Section2").Controls.Item("lblPreco").Visible = False
            
            rptListaEstoque.Sections("Section1").Controls.Item("txtPreco").Visible = False
            
            
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot1").Visible = True
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot2").Visible = False
            rptListaEstoque.Sections("Section5").Controls.Item("lblTot3").Visible = False
            
               rptListaEstoque.Sections("Section5").Controls.Item("lblTot1").Caption = ChkVal(saldoF, 0, cDecQtd)
            
            rptListaEstoque.Show 1
    End If
End Sub
Private Sub Rpt_005()
    '*
    '* Data: 13/01/2012
    '* Analise de saldo/custo no periodo
    '* Objetivo: Informar o 3 saldos/custos o inicial o final e o atual
    '*
    Dim sTabela As String
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim Rst1    As Recordset
    Dim saldoF  As String
    Dim sTable  As String
    
    Dim SaldoIT     As String
    Dim SaldoFT     As String
    Dim SaldoDif    As String
    
    sTabela = "tmp_estoqueanaliserpt_005"
    
    If optAnalise.Item(0).Value = True Then
            'Sintetico
            MontarTabelaTemporaria_005_001 sTabela
        Else
            'Analitico
            MontarTabelaTemporaria_005 sTabela
    End If
    
    sSQL = "SELECT SUM(SaldoI) As SaldoIT, SUM(SaldoF) As SaldoFT, SUM(Diferenca) As SaldoDif " & _
           "FROM " & sTabela & " " & _
           "ORDER BY Descricao"
        
    Set Rst1 = RegistroBuscar(sSQL)
    If Rst1.BOF And Rst1.EOF Then
            SaldoIT = ChkVal("0.00", 0, cDecQtd)
            SaldoFT = ChkVal("0.00", 0, cDecQtd)
            SaldoDif = ChkVal("0.00", 0, cDecQtd)
        Else
            Rst1.MoveFirst
            
            SaldoIT = ChkVal(cNull(Rst1.Fields("SaldoIT")), 0, cDecQtd)
            SaldoFT = ChkVal(cNull(Rst1.Fields("SaldoFT")), 0, cDecQtd)
            SaldoDif = ChkVal(cNull(Rst1.Fields("SaldoDif")), 0, cDecQtd)
            

    End If
    Rst1.Close
    
    sSQL = "SELECT * " & _
           "FROM " & sTabela & " " & _
            IIf(optAnalise(0).Value = True, "ORDER BY Descricao", OrderBy)
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            
            
            Rst.MoveFirst
            
            Set rptListaEstoquePeriodo.DataSource = Rst.DataSource
            
            rptListaEstoquePeriodo.Sections("Section2").Controls.Item("lblSaldoI").Caption = "Saldo em " & dtpIni.Value
            rptListaEstoquePeriodo.Sections("Section2").Controls.Item("lblSaldoF").Caption = "Saldo em " & dtpFin.Value
            rptListaEstoquePeriodo.Sections("Section2").Controls.Item("lblSaldoA").Caption = "Movimento "
            
            rptListaEstoquePeriodo.Sections("Section5").Controls.Item("lblTot1").Caption = SaldoIT
            rptListaEstoquePeriodo.Sections("Section5").Controls.Item("lblTot2").Caption = SaldoFT
            rptListaEstoquePeriodo.Sections("Section5").Controls.Item("lblTot3").Caption = SaldoDif
            
            rptListaEstoquePeriodo.Show 1
    End If
End Sub
Private Function MontarTabelaTemporaria_003() As Boolean
    On Error GoTo TrtErro_003
    Dim sSQL        As String
    Dim tabela      As String
    Dim Rst         As Recordset
    Dim vReg(1000)  As Variant
    Dim cReg        As Integer
    Dim vTotal      As String
    
    
    
    tabela = "tmp_estoqueanalise"
    MontarTabelaTemporaria_003 = False
    
    BD.Execute "DROP TABLE IF EXISTS " & tabela
    
    BD.Execute "CREATE TABLE IF NOT EXISTS " & tabela & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "idProd INT default Null," & _
               "ref VARCHAR(100) default Null," & _
               "balanco INT default 0," & _
               "Descricao VARCHAR(200) default Null," & _
               "Unidade VARCHAR(100) default Null," & _
               "Saldo VARCHAR(100) default Null," & _
               "Custo VARCHAR(100) default Null," & _
               "Preco VARCHAR(100) default Null," & _
               "Grupo VARCHAR(100) default Null," & _
               "SubGrupo VARCHAR(100) default Null," & _
               "PRIMARY KEY (Id))"
               
    sSQL = "SELECT * FROM estoqueproduto" & _
           " WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & _
           " AND STATUS = 'ATIVO'" & _
           " AND IncluirBalanco = 1 " & OrderBy

   ' MsgBox "Removido " & " AND IncluirBalanco = 1 " & OrderBy

    '**************************************************************************
    '*** VARIAVEIS QUE MODIFICAO A SITUACAO FINAL DO BALANCO
    Dim tmp_Saldo       As String 'Variavel para zerar estoque com saldo negativo
    Dim tmp_vCusto      As String 'Aletra o valor do custo
    '**************************************************************************
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado.", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            cReg = 0
            Do Until Rst.EOF
'                If Rst.Fields("id") = 3266 Then
'                    MsgBox "ww"
'                End If
                status (Rst.RecordCount)
                vReg(cReg) = Array("idProd", Rst.Fields("id"), "N"): cReg = cReg + 1
                vReg(cReg) = Array("ref", cNull(Rst.Fields("referencia")), "S"): cReg = cReg + 1
                vReg(cReg) = Array("balanco", IIf(Trim(cNull(Rst.Fields("incluirbalanco"))) <> "1", "0", "1"), "N"): cReg = cReg + 1
                vReg(cReg) = Array("Descricao", Rst.Fields("Descricao"), "S"): cReg = cReg + 1
                vReg(cReg) = Array("Unidade", UCase(Rst.Fields("unidade")), "S"): cReg = cReg + 1
                
                '>>>> Custo
                tmp_vCusto = ChkVal(IIf(IsNull(Rst.Fields("Custo")), "0", Rst.Fields("Custo")), 0, cDecMoeda)
                'tmp_vCusto = Val(tmp_vCusto) - Val(ChkVal(10 * Val(tmp_vCusto) / 100, 0, cDecMoeda))
                tmp_vCusto = ChkVal(tmp_vCusto, 0, cDecMoeda)
                '* 27.12.2012
                '* Caso haja equivalencia de grupo e sub grupo
                '* alterar valor de custo e registrar o mesmo.
                '*
                'DoEvents
                'Debug.Print "***"
                'If Rst.fields("subGrupo") & "/" & Rst.fields("Grupo") = "3/3" Then
                'Debug.Print pgDescrSubGrupo(Rst.fields("subGrupo")) & "/" & pgDescrGrupo(Rst.fields("Grupo"))
                '    tmp_vCusto = ChkVal(Val(tmp_vCusto) * 1.16, 0, cDecMoeda)
               '
               '     BD.Execute "UPDATE estoqueproduto SET Custo=" & tmp_vCusto & " WHERE ID = " & Rst.fields("ID")
               '     BD.Execute "UPDATE estoqueproduto SET IncluirBalanco=1 WHERE ID = " & Rst.fields("ID")
               '     'Debug.Print "Alterado " & Rst.fields("id") & " - " & tmp_vCusto
               ' End If
                vReg(cReg) = Array("Custo", ChkVal(tmp_vCusto, 0, cDecMoeda), "S"): cReg = cReg + 1
                
                '>>>> Saldo
                tmp_Saldo = IIf(IsNull(Rst.Fields("Saldo")) Or Rst.Fields("Saldo") < 0, "0", Rst.Fields("Saldo"))
                'tmp_Saldo = ChkVal(IIf(tmp_Saldo < 0, 0, tmp_Saldo), 0, cDecQtd)
                vReg(cReg) = Array("saldo", tmp_Saldo, "S"): cReg = cReg + 1
                
                vReg(cReg) = Array("Grupo", pgDescrGrupo(IIf(IsNull(Rst.Fields("Grupo")), 0, Rst.Fields("Grupo"))), "S"): cReg = cReg + 1
                vReg(cReg) = Array("SubGrupo", pgDescrSubGrupo(IIf(IsNull(Rst.Fields("SubGrupo")), 0, Rst.Fields("SubGrupo"))), "S"): cReg = cReg + 1
                
                'vTotal = Val(ChkVal(IIf(IsNull(Rst.Fields("Saldo")), "0", Rst.Fields("Saldo")), 0, cDecQtd)) * Val(ChkVal(IIf(IsNull(Rst.Fields("Custo")), "0", Rst.Fields("Custo")), 0, cDecMoeda))
                
                'Caso o saldo seja negativo o valor do total de ser zero
                If Val(tmp_Saldo) < 0 Then
                        vTotal = 0
                    Else
                        vTotal = Val(ChkVal(tmp_Saldo, 0, cDecQtd)) * Val(ChkVal(tmp_vCusto, 0, cDecMoeda))
                End If
                
                vTotal = ChkVal(vTotal, 0, cDecMoeda)
                
                vReg(cReg) = Array("Preco", vTotal, "S")
                '
                'RJ, 15.12.2017
                'Codigo comentado pois material mesmo com total zerado
                'pode conter itens no estoque
                '
                'Se os valores forem zero nao gravar
                'If vTotal <= 0 Then
                '        Dim mmm As String
                '        mmm = "expurgado"
                '    Else
                        RegistroIncluir tabela, vReg, cReg
                'End If
                
                cReg = 0
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    MontarTabelaTemporaria_003 = True
    Exit Function
TrtErro_003:
    MsgBox "Erro ao montar base de dados para relatorio!" & vbCrLf & Err.Description, vbInformation, "Aviso" & Err.Number
    MontarTabelaTemporaria_003 = False
    Exit Function
End Function
Private Function MontarTabelaTemporaria_005(tabela As String) As Boolean
'* RJ, 22.10.2012
'* Monta tabela temporaria para armazenar os saldos inicial, final e atual
'*

    On Error GoTo TrtErro_005
    Dim sSQL        As String
    'Dim Tabela      As String
    Dim Rst         As Recordset
    Dim Rst1        As Recordset
    Dim vReg(1000)  As Variant
    Dim cReg        As Integer
    Dim vTotal      As String
    Dim iG          As String
    Dim iSG         As String
    
    
    If Trim(cboGrupo.Text) = "" Then
            iG = ""
        Else
            iG = "Grupo = " & Left(cboGrupo.Text, 5)
    End If
    
    If Trim(cboSubgrupo.Text) = "" Then
            iSG = ""
        Else
            iSG = "subGrupo = " & Left(cboSubgrupo.Text, 5)
    End If
    
    
    
    
    
    MontarTabelaTemporaria_005 = False
    
    BD.Execute "DROP TABLE IF EXISTS " & LCase(tabela)
    
    sSQL = "CREATE TABLE IF NOT EXISTS " & tabela & _
           " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
           "Id_Empresa INT default Null," & _
           "DtHr VARCHAR(20) default Null," & _
           "UsuID INT default Null," & _
           "prodID NUMERIC default Null," & _
           "referencia VARCHAR(250) default Null," & _
           "grupo NUMERIC default Null," & _
           "sgrupo NUMERIC default Null," & _
           "Descricao VARCHAR(200) default Null," & _
           "Unidade VARCHAR(100) default Null," & _
           "SaldoI VARCHAR(100) default Null," & _
           "SaldoF VARCHAR(100) default Null," & _
           "Diferenca VARCHAR(100) default Null," & _
           "PRIMARY KEY (Id))"
    
    BD.Execute sSQL
    
    sSQL = "SELECT * FROM EstoqueProduto " & _
           "WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND status ='ATIVO' AND IncluirBalanco=1 " & _
           IIf(iG = "", "", " AND " & iG) & _
           IIf(iSG = "", "", " AND " & iSG) & _
           " " & OrderBy
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado.", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            
            Do Until Rst.EOF
                cReg = 0
                status (Rst.RecordCount)
                vReg(cReg) = Array("prodID", Rst.Fields("id"), "N"): cReg = cReg + 1
                vReg(cReg) = Array("referencia", cNull(Rst.Fields("referencia")), "S"): cReg = cReg + 1
                vReg(cReg) = Array("Descricao", Rst.Fields("Descricao"), "S"): cReg = cReg + 1
                vReg(cReg) = Array("Unidade", UCase(Rst.Fields("unidade")), "S"): cReg = cReg + 1
                
                vReg(cReg) = Array("grupo", UCase(Rst.Fields("grupo")), "N"): cReg = cReg + 1
                vReg(cReg) = Array("sgrupo", UCase(Rst.Fields("subgrupo")), "N"): cReg = cReg + 1
                
                '*
                '* Buscar no Kardex mov da mercadoria
                '*
                Dim vIni    As String
                Dim vFin    As String
                Dim vDif    As String
                
                'sSQL = "SELECT * FROM estoquekardex " & _
                        "WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " " & _
                        "AND IDProduto = " & Rst.fields("ID") & " " & _
                        "AND DataMov BETWEEN '" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpFin.Value, "YYYY-MM-DD") & "' " & _
                        "ORDER BY DataMov"
                        
                '*Buscar data Inicial
                sSQL = "SELECT * FROM estoquekardex " & _
                        "WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " " & _
                        "AND IDProduto = " & Rst.Fields("ID") & " " & _
                        "AND DataMov <='" & Format(dtpIni.Value, "YYYY-MM-DD") & "'" & _
                        "ORDER BY DataMov"
                Set Rst1 = RegistroBuscar(sSQL)
                If Rst1.BOF And Rst1.EOF Then
                        vIni = 0
                        
                    Else
                        
                        Rst1.MoveLast
                        vIni = cNull(Rst1.Fields("Saldo"))
                End If
                Rst1.Close
                
                
                '*Buscar data Final
                sSQL = "SELECT * FROM estoquekardex " & _
                        "WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " " & _
                        "AND IDProduto = " & Rst.Fields("ID") & " " & _
                        "AND DataMov >='" & Format(dtpFin.Value, "YYYY-MM-DD") & "'" & _
                        "ORDER BY DataMov"
                Set Rst1 = RegistroBuscar(sSQL)
                If Rst1.BOF And Rst1.EOF Then
                        
                        vFin = IIf(dtpFin.Value >= Date, cNull(Rst.Fields("Saldo")), 0)
                    Else
                        Rst1.MoveFirst
                        vFin = cNull(Rst1.Fields("Saldo"))
                End If
                Rst1.Close

                '        '* Inicial
                '        vReg(cReg) = Array("SaldoI", ChkVal("0,00", 0, cDecQtd), "S"): cReg = cReg + 1
                '        'vReg(cReg) = Array("CustoI", ChkVal("0,00", 0, cDecMoeda), "S"): cReg = cReg + 1
                '        '* Final
                '        vReg(cReg) = Array("SaldoF", ChkVal("0,00", 0, cDecQtd), "S"): cReg = cReg + 1
                '        'vReg(cReg) = Array("CustoF", ChkVal("0,00", 0, cDecMoeda), "S"): cReg = cReg + 1
                '    Else
                '
                '        '* Inicial
                '        Rst1.MoveFirst
                '        vReg(cReg) = Array("SaldoI", ChkVal(Rst1.fields("Saldo"), 0, cDecQtd), "S"): cReg = cReg + 1
                '        'vReg(cReg) = Array("CustoI", ChkVal(Rst1.fields("Custo"), 0, cDecMoeda), "S"): cReg = cReg + 1
                '        '* Final
                '        Rst1.MoveLast
                '        vReg(cReg) = Array("SaldoF", ChkVal(Rst1.fields("Saldo"), 0, cDecQtd), "S"): cReg = cReg + 1
                '        'vReg(cReg) = Array("CustoF", ChkVal(Rst1.fields("Custo"), 0, cDecMoeda), "S"): cReg = cReg + 1
                'End If
                'vReg(cReg) = Array("SaldoI", ChkVal(Rst1.fields("Saldo"), 0, cDecQtd), "S"): cReg = cReg + 1
                        '* Atual
                
                
                
                vReg(cReg) = Array("SaldoI", ChkVal(vIni, 0, cDecQtd), "S"): cReg = cReg + 1
                vReg(cReg) = Array("SaldoF", ChkVal(vFin, 0, cDecQtd), "S"): cReg = cReg + 1
                
                vDif = Val(ChkVal(vFin, 0, cDecQtd)) - Val(ChkVal(vIni, 0, cDecQtd))
                vDif = ChkVal(vDif, 0, cDecQtd)
                
                vReg(cReg) = Array("Diferenca", ChkVal(vDif, 0, cDecQtd), "S"): cReg = cReg + 1
                
                'Rst1.Close
                cReg = cReg - 1
                RegistroIncluir tabela, vReg, cReg
                cReg = 0
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    MontarTabelaTemporaria_005 = True
    Exit Function
TrtErro_005:
    MsgBox "Erro ao montar base de dados para relatorio!" & vbCrLf & Err.Description, vbInformation, "Aviso" & Err.Number
    MontarTabelaTemporaria_005 = False
    Exit Function
End Function

Private Function MontarTabelaTemporaria_005_001(tabela As String) As Boolean
'* RJ, 22.10.2012
'* Monta tabela temporaria para armazenar os saldos inicial, final e atual
'* SINTETICO
'*
    'On Error GoTo TrtErro_005
    
    Dim a(1000, 1000) As Variant
    
    
    Dim sSQL        As String
    'Dim Tabela      As String
    Dim Rst         As Recordset
    Dim Rst1        As Recordset
    Dim vReg(1000)  As Variant
    Dim cReg        As Integer
    Dim vTotal      As String
    
    Dim iG          As String
    Dim iSG         As String
    
    Dim vIni    As String
    Dim vFin    As String
    Dim vDif    As String
    
    'If Trim(cboGrupo.Text) = "" Then
    '        iG = ""
    '    Else
    '        iG = "Grupo = " & Left(cboGrupo.Text, 5)
    'End If
    '
    'If Trim(cboSubgrupo.Text) = "" Then
    '        iSG = ""
    '    Else
    '        iSG = "subGrupo = " & Left(cboSubgrupo.Text, 5)
    'End If
    
    
    
    
    
    MontarTabelaTemporaria_005_001 = False
    
    BD.Execute "DROP TABLE IF EXISTS " & LCase(tabela)
    
    sSQL = "CREATE TABLE IF NOT EXISTS " & tabela & _
           " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
           "Id_Empresa INT default Null," & _
           "DtHr VARCHAR(20) default Null," & _
           "UsuID INT default Null," & _
           "prodID NUMERIC default Null," & _
           "referencia VARCHAR(250) default Null," & _
           "grupo NUMERIC default Null," & _
           "sgrupo NUMERIC default Null," & _
           "Descricao VARCHAR(200) default Null," & _
           "Unidade VARCHAR(100) default Null," & _
           "SaldoI VARCHAR(100) default Null," & _
           "SaldoF VARCHAR(100) default Null," & _
           "Diferenca VARCHAR(100) default Null," & _
           "PRIMARY KEY (Id))"
    
    BD.Execute sSQL
    
    sSQL = "SELECT * FROM EstoqueProduto " & _
           "WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND status ='ATIVO' AND IncluirBalanco=1 " & _
           IIf(iG = "", "", " AND " & iG) & _
           IIf(iSG = "", "", " AND " & iSG) & _
           " ORDER BY grupo, subgrupo, Descricao"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado.", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            vIni = 0
            vFin = 0
            vDif = 0
            iG = Rst.Fields("Grupo")
            iSG = Rst.Fields("SubGrupo")
            
            Do Until Rst.EOF
              
                status (Rst.RecordCount)

                '*
                '* Buscar no Kardex movimento da mercadoria
                '*
                Dim vTMP_I As String
                Dim vTMP_F As String
                Dim vTMP_D As String
                '***************************************************************************
                '*Buscar data Inicial
                '***************************************************************************
                sSQL = "SELECT * FROM estoquekardex " & _
                        "WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " " & _
                        "AND IDProduto = " & Rst.Fields("ID") & " " & _
                        "AND DataMov <='" & Format(dtpIni.Value, "YYYY-MM-DD") & "'" & _
                        "ORDER BY DataMov"
                Set Rst1 = RegistroBuscar(sSQL)
                If Rst1.BOF And Rst1.EOF Then
                        vTMP_I = 0
                    Else
                        Rst1.MoveLast
                        vTMP_I = ChkVal(cNull(Rst1.Fields("Saldo")), 0, cDecQtd)
                End If
                Rst1.Close
                vIni = Val(ChkVal(vIni, 0, cDecQtd)) + Val(ChkVal(vTMP_I, 0, cDecQtd))
                '***************************************************************************
                
                
                '***************************************************************************
                '*Buscar data Final
                '***************************************************************************
                sSQL = "SELECT * FROM estoquekardex " & _
                        "WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " " & _
                        "AND IDProduto = " & Rst.Fields("ID") & " " & _
                        "AND DataMov >='" & Format(dtpFin.Value, "YYYY-MM-DD") & "'" & _
                        "ORDER BY DataMov"
                Set Rst1 = RegistroBuscar(sSQL)
                If Rst1.BOF And Rst1.EOF Then
                        vTMP_F = ChkVal(cNull(Rst.Fields("Saldo")), 0, cDecQtd)
                    Else
                        Rst1.MoveFirst
                        vTMP_F = ChkVal(cNull(Rst1.Fields("Saldo")), 0, cDecQtd)
                End If
                Rst1.Close
                vFin = Val(ChkVal(vFin, 0, cDecQtd)) + Val(ChkVal(vTMP_F, 0, cDecQtd))
                vTMP_D = Val(ChkVal(vTMP_F, 0, cDecQtd)) - Val(ChkVal(vTMP_I, 0, cDecQtd))
                vDif = Val(ChkVal(vDif, 0, cDecQtd)) + Val(ChkVal(vTMP_D, 0, cDecQtd))
                '***************************************************************************
                
                
                'Dim v As Variant
                'v = IIf(IsMissing(a(Rst.fields("Grupo"), Rst.fields("SubGrupo"))(0)), 0, 1)
                'a(Rst.fields("Grupo"), Rst.fields("SubGrupo")) = Array(v)
                
                

                Rst.MoveNext
                If Rst.EOF Then
                        Rst.MovePrevious
                        cReg = 0
                        vReg(cReg) = Array("prodID", Rst.Fields("id"), "N"): cReg = cReg + 1
                        vReg(cReg) = Array("referencia", cNull(Rst.Fields("referencia")), "S"): cReg = cReg + 1
                        vReg(cReg) = Array("Descricao", pgDescrSubGrupo(iSG) & " " & pgDescrGrupo(iG), "S"): cReg = cReg + 1
                        vReg(cReg) = Array("Unidade", UCase(Rst.Fields("unidade")), "S"): cReg = cReg + 1
                
                        vReg(cReg) = Array("grupo", iG, "N"): cReg = cReg + 1
                        vReg(cReg) = Array("sgrupo", iSG, "N"): cReg = cReg + 1

                        vReg(cReg) = Array("SaldoI", ChkVal(vIni, 0, cDecQtd), "S"): cReg = cReg + 1
                        vReg(cReg) = Array("SaldoF", ChkVal(vFin, 0, cDecQtd), "S"): cReg = cReg + 1
                        vReg(cReg) = Array("Diferenca", ChkVal(vDif, 0, cDecQtd), "S"): cReg = cReg + 1
                
                        cReg = cReg - 1
                        RegistroIncluir tabela, vReg, cReg
                        Exit Do
                    
                    Else
                        If ZE(CInt(Rst.Fields("Grupo")), 6) & ZE(CInt(Rst.Fields("SubGrupo")), 6) <> ZE(CInt(iG), 6) & ZE(CInt(iSG), 6) Then
                            cReg = 0
                            Rst.MovePrevious
                            vReg(cReg) = Array("prodID", Rst.Fields("id"), "N"): cReg = cReg + 1
                            vReg(cReg) = Array("referencia", cNull(Rst.Fields("referencia")), "S"): cReg = cReg + 1
                            vReg(cReg) = Array("Descricao", pgDescrSubGrupo(iSG) & " " & pgDescrGrupo(iG), "S"): cReg = cReg + 1
                            vReg(cReg) = Array("Unidade", UCase(Rst.Fields("unidade")), "S"): cReg = cReg + 1
                
                            vReg(cReg) = Array("grupo", iG, "N"): cReg = cReg + 1
                            vReg(cReg) = Array("sgrupo", iSG, "N"): cReg = cReg + 1

                            vReg(cReg) = Array("SaldoI", ChkVal(vIni, 0, cDecQtd), "S"): cReg = cReg + 1
                            vReg(cReg) = Array("SaldoF", ChkVal(vFin, 0, cDecQtd), "S"): cReg = cReg + 1
                            vReg(cReg) = Array("Diferenca", ChkVal(vDif, 0, cDecQtd), "S"): cReg = cReg + 1
                
                            cReg = cReg - 1
                            RegistroIncluir tabela, vReg, cReg
                    
                            Rst.MoveNext
                    
                                        
                            iG = Rst.Fields("Grupo")
                            iSG = Rst.Fields("SubGrupo")
                            vIni = 0
                            vFin = 0
                            vDif = 0
                    End If
                End If
                 
            Loop
                'vReg(cReg) = Array("prodID", Rst.fields("id"), "N"): cReg = cReg + 1
                'vReg(cReg) = Array("Descricao", "Grupo/Sub", "S"): cReg = cReg + 1
                'vReg(cReg) = Array("Unidade", UCase(Rst.fields("unidade")), "S"): cReg = cReg + 1
               '
               ' vReg(cReg) = Array("grupo", UCase(Rst.fields("grupo")), "N"): cReg = cReg + 1
               ' vReg(cReg) = Array("sgrupo", UCase(Rst.fields("subgrupo")), "N"): cReg = cReg + 1
'
 '               vReg(cReg) = Array("SaldoI", ChkVal(vIni, 0, cDecQtd), "S"): cReg = cReg + 1
 '               vReg(cReg) = Array("SaldoF", ChkVal(vFin, 0, cDecQtd), "S"): cReg = cReg + 1
 '               vReg(cReg) = Array("Diferenca", ChkVal(vDif, 0, cDecQtd), "S"): cReg = cReg + 1
 '
 '               cReg = cReg - 1
  '              RegistroIncluir Tabela, vReg, cReg
 '               cReg = 0
                
    End If
    Rst.Close
    MontarTabelaTemporaria_005_001 = True
    Exit Function
TrtErro_005:
    MsgBox "Erro ao montar base de dados para relatorio!" & vbCrLf & Err.Description, vbInformation, "Aviso" & Err.Number
    MontarTabelaTemporaria_005_001 = False
    Exit Function
End Function


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

