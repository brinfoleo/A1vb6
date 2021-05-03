VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formFinanceiroDRE 
   Caption         =   "DRE - Demonstração do resultado em exercício"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   10140
   Begin MSFlexGridLib.MSFlexGrid msfgDRE 
      Height          =   6135
      Left            =   60
      TabIndex        =   6
      Top             =   1500
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   4
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"formFinanceiroDRE.frx":0000
   End
   Begin VB.Frame frmMenu 
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   10035
      Begin VB.CommandButton btAtualiza 
         Caption         =   "&Atualizar"
         Height          =   495
         Left            =   4560
         TabIndex        =   7
         Top             =   420
         Width           =   975
      End
      Begin VB.Frame frmPeriodo 
         Caption         =   "Periodo:"
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
         TabIndex        =   1
         Top             =   180
         Width           =   4275
         Begin MSComCtl2.DTPicker dtpDtInicio 
            Height          =   315
            Left            =   480
            TabIndex        =   2
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   110297089
            CurrentDate     =   40557
         End
         Begin MSComCtl2.DTPicker dtpDtFinal 
            Height          =   315
            Left            =   2460
            TabIndex        =   3
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   110297089
            CurrentDate     =   40557
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Até:"
            Height          =   195
            Left            =   2100
            TabIndex        =   5
            Top             =   420
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "De:"
            Height          =   195
            Left            =   180
            TabIndex        =   4
            Top             =   420
            Width           =   255
         End
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   1260
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "formFinanceiroDRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListagemPlanoContas()
    'If chkAcesso(Me, "i") = False Then Exit Sub
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    Dim i           As Integer
    Dim ii          As Integer
    Dim Rst         As Recordset
    Dim sSQL        As String
    
 
    
    'Dados do campo
    Dim dc As String
    
    
'    Dim rst0 As Recordset
'
'    sSQL = "SELECT * FROM FinanceiroContasPRCadastro"
'    sSQL = sSQL & " WHERE ID_Empresa = 1 AND"
'    sSQL = sSQL & " Vencimento >= '2016-05-01' AND Vencimento <= '2016-07-04'"
'    sSQL = sSQL & " ORDER BY PlanocontasCodigo"
'    Set rst0 = RegistroBuscar(sSQL)
'
'    If rst0.BOF And rst0.EOF Then
'            Exit Sub
'        Else
'            rst0.MoveFirst
'    End If
'
'    With msfgDRE
'        .Enabled = False
'        For i = 1 To .Rows - 1
'            'status (.Rows - 1)
'            cReg = 0
'            For ii = 1 To .Cols - 1
'                dc = msfgDRE.TextMatrix(i, ii)
'                If Left(dc, 2) = "R$" Then
'                    dc = ChkVal(Mid(dc, 3, Len(dc)), 0, cDecMoeda)
'                End If
'                vReg(cReg) = Array(RS(.TextMatrix(0, ii)), dc, "S"): cReg = cReg + 1
'            Next
'            Dim idPC As String
'            idPC = PgDadosFinanceiroFatura(msfgDRE.TextMatrix(i, 0)).idPlanoContas
'            vReg(cReg) = Array(RS("codPlanContas"), PgDadosPlanoContas("id", idPC).Codigo, "S"): cReg = cReg + 1
'            cReg = cReg - 1
'            RegistroIncluir "tmp_Titulos", vReg, cReg
'        Next
'
'        .Enabled = True
'    End With
'
'
'    sSQL = "SELECT * FROM FinanceiroContasPRCadastro"
'    sSQL = sSQL & " WHERE ID_Empresa = 1 AND"
'    sSQL = sSQL & " Vencimento >= '" & Format(dtpDtInicio.Value, "YYYY-MM-dd") & "' AND Vencimento <= '" & Format(dtpDtFinal.Value, "YYYY-MM-dd") & "'"
'    sSQL = sSQL & " ORDER BY PlanocontasCodigo"
    MontarTabelaTemporariaPC
    
    Dim sSQL2 As String
    Dim Rst2 As Recordset
    Dim vl  As String
    Dim porcentagem  As String 'Percentual aplicado
    
    Dim totGrupo As Boolean 'informa se é um totalizador ou nao
    Dim vlGrupo As String 'armazena o valor global do grupo
    
    '29.06.2016
    'Agrupa todos os valores dos codigos do plano de contas
    'sSQL = "SELECT DISTINCT codPlanContas FROM tmp_titulos ORDER BY codPlanContas"
    sSQL = "SELECT * FROM tmp_titulospc ORDER BY codigo"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                
                totGrupo = IIf(PgDadosPlanoContas("codigo", "'" & Rst.Fields("codigo") & "'").totalizador = 1, True, False)
                '29.06.2016 - Se o item for um totalizador usar like
                If totGrupo = True Then
                        'sSQL2 = "select codplancontas, sum(ValorAtualizado) AS tValor FROM tmp_titulos WHERE codplancontas LIKE '" & Rst.Fields("PlanocontasCodigo") & ".%'" 'group by codplancontas"
                        
                        
                        sSQL2 = "SELECT PlanocontasCodigo, sum(vlDuplicata) AS tValor"
                        sSQL2 = sSQL2 & " FROM FinanceiroContasPRCadastro"
                        sSQL2 = sSQL2 & " WHERE ID_Empresa = 1 AND"
                        sSQL2 = sSQL2 & " Vencimento >= '" & Format(dtpDtInicio.Value, "YYYY-MM-dd") & "' AND Vencimento <= '" & Format(dtpDtFinal.Value, "YYYY-MM-dd") & "'"
                        sSQL2 = sSQL2 & " AND (PlanocontasCodigo = '" & Rst.Fields("codigo") & "'"
                        sSQL2 = sSQL2 & " OR PlanocontasCodigo LIKE '" & Rst.Fields("codigo") & ".%')"
                        sSQL2 = sSQL2 & " ORDER BY PlanocontasCodigo"
                        
                        
                    Else
                        'sSQL2 = "select codplancontas, sum(ValorAtualizado) AS tValor FROM tmp_titulos WHERE codplancontas = '" & Rst.Fields("PlanocontasCodigo") & "'" 'group by codplancontas"
                        sSQL2 = "SELECT PlanocontasCodigo, sum(vlDuplicata) AS tValor"
                        sSQL2 = sSQL2 & " FROM FinanceiroContasPRCadastro"
                        sSQL2 = sSQL2 & " WHERE ID_Empresa = 1 AND"
                        sSQL2 = sSQL2 & " Vencimento >= '" & Format(dtpDtInicio.Value, "YYYY-MM-dd") & "' AND Vencimento <= '" & Format(dtpDtFinal.Value, "YYYY-MM-dd") & "'"
                        sSQL2 = sSQL2 & " AND PlanocontasCodigo = '" & Rst.Fields("codigo") & "'"
                        sSQL2 = sSQL2 & " ORDER BY PlanocontasCodigo"
                        
                End If
                Set Rst2 = RegistroBuscar(sSQL2)
                If Rst2.BOF And Rst2.EOF Then
                        vl = "0"
                    Else
                        Rst2.MoveFirst
                        vl = IIf(cNull(Rst2.Fields("tValor")) = "", 0, cNull(Rst2.Fields("tValor")))
                End If
                If totGrupo = True Then
                        vlGrupo = vl
                        porcentagem = IIf(Val(vl) <= 0, "0", "100")
                    Else
                        If vl = 0 Then
                                porcentagem = "0"
                            Else
                                porcentagem = (Val(vl) * 100) / Val(vlGrupo)
                        End If
                End If
                porcentagem = ChkVal(porcentagem, 0, 2) & "%"
                Rst2.Close
                
                'Dim descrPC As String
                'descrPC = PgDadosPlanoContas("codigo", "'" & Rst.Fields("codigo") & "'").Descricao
        
                cReg = 0
                'vReg(cReg) = Array("codigo", Rst.Fields("codplancontas"), "S"): cReg = cReg + 1
                '
                vReg(cReg) = Array("valor", Replace(ChkVal(vl, 0, cDecMoeda), ".", ","), "S"): cReg = cReg + 1
                vReg(cReg) = Array("porcentagem ", porcentagem, "S"): cReg = cReg + 1
                cReg = cReg - 1
                'RegistroIncluir "tmp_titulospc", vReg, cReg
                RegistroAlterar "tmp_titulospc", vReg, cReg, "codigo='" & Rst.Fields("codigo") & "'"
                Rst.MoveNext
            Loop
            
    End If
    
    '31/05/2016 - Gerar o relatorio
    Dim Rst3 As Recordset
    sSQL = "SELECT * FROM tmp_titulospc ORDER BY codigo"
    Set Rst3 = RegistroBuscar(sSQL)
    If Rst3.BOF And Rst3.EOF Then
            Exit Sub
        Else
        'Set rptListaTitulosPC.DataSource = Rst3.DataSource
        'rptListaTitulosPC.Show 1
        msfgDRE.Rows = 1
        Do Until Rst3.EOF
            With msfgDRE
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = cNull(Rst3.Fields("codigo"))
                .TextMatrix(.Rows - 1, 1) = cNull(Rst3.Fields("Descricao"))
                .TextMatrix(.Rows - 1, 2) = ConvMoeda(ChkVal(cNull(Rst3.Fields("valor")), 0, cDecMoeda))
                .TextMatrix(.Rows - 1, 3) = cNull(Rst3.Fields("porcentagem"))
                If PgDadosPlanoContas("codigo", "'" & Rst3.Fields("codigo") & "'").totalizador = 1 Then
                    .FillStyle = flexFillRepeat
                    .Row = .Rows - 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    .CellFontBold = True
                    '.CellForeColor = vbRed
                    
                    
                End If
            End With
            Rst3.MoveNext
        Loop
    End If
    Rst3.Close
    
End Sub

'Private Sub MontarTabelaTemporaria()
'
'    Dim sCampos     As String
'    Dim i           As Integer
'
'    BD.Execute "DROP TABLE IF EXISTS tmp_titulos"
'    sCampos = ""
'    For i = 1 To msfgContas.Cols - 1
'        sCampos = sCampos & RS(msfgContas.TextMatrix(0, i)) & " VARCHAR(100) default Null,"
'    Next
'
'    sCampos = sCampos & " codPlanContas VARCHAR(100) default Null,"
'
'    sCampos = "CREATE TABLE IF NOT EXISTS tmp_titulos " & _
'              "(Id INT(11) NOT NULL AUTO_INCREMENT," & _
'               "Id_Empresa INT default Null," & _
'               "UsuID VARCHAR(10) default Null," & _
'               "DtHr VARCHAR(20) default Null," & _
'               sCampos & " PRIMARY KEY (" & msfgContas.TextMatrix(0, 0) & "))"
'    BD.Execute sCampos
'End Sub



Private Sub MontarTabelaTemporariaPC()
    '24/05/2016
    'Tabela temporaria plano de contas
    Dim sCampos     As String
    Dim i           As Integer
    
    BD.Execute "DROP TABLE IF EXISTS tmp_titulospc"
    sCampos = "CREATE TABLE IF NOT EXISTS tmp_titulospc " & _
              "(Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "UsuID VARCHAR(10) default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "codigo VARCHAR(20) default Null," & _
               "Descricao VARCHAR(100) default Null," & _
               "totalizador VARCHAR(100) default Null," & _
               "valor VARCHAR(20) default Null," & _
               "porcentagem VARCHAR(20) default Null," & _
               " PRIMARY KEY (Id))"

    BD.Execute sCampos
    
    '29.06.2016 - Copia os dados da tabela de plano de contas
    BD.Execute "INSERT INTO tmp_titulospc (Id_Empresa, UsuID, codigo, descricao, totalizador) SELECT Id_Empresa, " & ID_Usuario & ", codigo, descricao, totalizador FROM financeiroplanocontas"
End Sub

Private Sub btAtualiza_Click()
    ListagemPlanoContas
    
End Sub

Private Sub Form_Load()
    dtpDtInicio.Value = Date
    dtpDtFinal.Value = Date
End Sub
Private Sub status(Max As Long)
    
    pb.min = 0
    pb.Max = Max
    DoEvents
    pb.Value = pb.Value + 1
    If pb.Value > 0 And pb.Value < Max Then
            Me.Enabled = False
            pb.Visible = True
            Me.Enabled = False
        Else
            Me.Enabled = True
            pb.Visible = False
            pb.Value = 0
            Me.Enabled = True
    End If
End Sub

Private Sub Form_Resize()

    frmMenu.Width = Me.Width - 300
    
    msfgDRE.Height = Me.ScaleHeight - (frmMenu.Height + 300)
    msfgDRE.Width = Me.Width - 300
    pb.Width = msfgDRE.Width
End Sub
