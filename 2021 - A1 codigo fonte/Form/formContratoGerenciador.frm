VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formContratoGerenciador 
   Caption         =   "Contratos"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11310
   Begin VB.Frame frmMenu 
      Height          =   1275
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   10035
      Begin VB.Frame Frame1 
         Caption         =   "Cliente:"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   180
         Width           =   3915
         Begin VB.ComboBox cboCliente 
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Text            =   "Combo1"
            Top             =   360
            Width           =   3735
         End
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
         TabIndex        =   5
         Top             =   180
         Width           =   4275
         Begin VB.OptionButton optPerido 
            Caption         =   "Nenhum"
            Height          =   195
            Index           =   2
            Left            =   900
            TabIndex        =   12
            Top             =   0
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optPerido 
            Caption         =   "Termino"
            Height          =   195
            Index           =   1
            Left            =   2760
            TabIndex        =   7
            Top             =   0
            Width           =   915
         End
         Begin VB.OptionButton optPerido 
            Caption         =   "Inicio"
            Height          =   195
            Index           =   0
            Left            =   1920
            TabIndex        =   6
            Top             =   0
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpDtInicio 
            Height          =   315
            Left            =   480
            TabIndex        =   8
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   97583105
            CurrentDate     =   40557
         End
         Begin MSComCtl2.DTPicker dtpDtFinal 
            Height          =   315
            Left            =   2460
            TabIndex        =   9
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   97583105
            CurrentDate     =   40557
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Até:"
            Height          =   195
            Left            =   2100
            TabIndex        =   11
            Top             =   420
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "De:"
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   420
            Width           =   255
         End
      End
   End
   Begin VB.Frame frmContratos 
      Height          =   4815
      Left            =   180
      TabIndex        =   0
      Top             =   1980
      Width           =   7515
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   60
         Visible         =   0   'False
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid msfgContratos 
         Height          =   3855
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6800
         _Version        =   393216
         Cols            =   5
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"formContratoGerenciador.frx":0000
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualiza"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alterar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Pesquisar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
               Picture         =   "formContratoGerenciador.frx":0095
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratoGerenciador.frx":04E7
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratoGerenciador.frx":0801
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratoGerenciador.frx":1093
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratoGerenciador.frx":22E5
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratoGerenciador.frx":2BBF
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratoGerenciador.frx":3451
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratoGerenciador.frx":3CE3
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratoGerenciador.frx":4F35
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratoGerenciador.frx":524F
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratoGerenciador.frx":5569
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formContratoGerenciador.frx":5960
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formContratoGerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private sSQLPesq    As String 'String com pesquisa a ser executada
Private OrderBy     As String 'Ordena de acordo com a coluna selecionada
Private idContrato  As Integer 'ID do contrato
Private idCliente   As Integer 'ID do Cliente
Private Function filtro() As String
    Dim tpPeriodo   As String
    Dim sqlFiltro   As String
    Dim tpCliente   As String 'Informa o cliente


    sqlFiltro = "SELECT * FROM contrato WHERE ID_Empresa = " & ID_Empresa


    
    '***** Filtro por data ******
    Dim sPeriodo As String
    If optPerido.Item(0).Value = True Then
            tpPeriodo = "dtIni"
        ElseIf optPerido.Item(1).Value = True Then
            tpPeriodo = "dtFin"
        ElseIf optPerido.Item(2).Value = True Then
            tpPeriodo = ""
    End If
 
   
    If Trim(tpPeriodo) <> "" Then
            sPeriodo = tpPeriodo & " >= '" & Format(dtpDtInicio.Value, "YYYY-MM-DD") & _
                        "' AND " & tpPeriodo & " <= '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & "'"
'        Else
'            sPeriodo = " AND dtini >= '" & Format(Date, "YYYY-MM-DD") & _
'                        "' AND dtini <= '" & Format(Date, "YYYY-MM-DD") & "'"
    End If
    '****************************************
    '********* Cliente ***********************
    
    If idCliente <> 0 Then
        tpCliente = "idCliente=" & idCliente
    End If
    
    
    
    sqlFiltro = sqlFiltro & _
                IIf(sPeriodo = "", "", " AND " & sPeriodo) & _
                IIf(tpCliente = "", "", " AND " & tpCliente)


'
'            IIf(ContaPR = "", "", " AND " & ContaPR) & _
'            IIf(ContaFV = "", "", " AND " & ContaFV) & _
'            IIf(DtQuitacao = "", "", " AND " & DtQuitacao) & _
'            IIf(tpDoc = "", "", " AND " & tpDoc) & _
'            IIf(cCustos = "", "", " AND " & cCustos) & _
'            IIf(Sacado = "", "", " AND " & Sacado) & _
'            IIf(vDuplicata = "", "", " AND " & vDuplicata) & _
'            IIf(vNossoNumero = "", "", " AND " & vNossoNumero) & _
'            IIf(nDupl = "", "", " AND " & nDupl)
'
    filtro = sqlFiltro
End Function


Private Sub Form_Activate()
    AtualizarLista
End Sub

Private Sub Form_Load()
    OrderBy = ""
    Me.Top = 0
    Me.Left = 0
    
    dtpDtInicio.Value = Date
    dtpDtFinal.Value = Date
    optPerido_Click (2)
    cboCliente.Text = ""
    filtro
    
End Sub
Private Sub AtualizarLista()
    On Error Resume Next
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim sqlEx       As String
    
    'If Trim(sSQLPesq) = "" Then Exit Sub

    DoEvents
    msfgContratos.Rows = 1
    
    sqlEx = filtro & IIf(Trim(OrderBy) = "", " ORDER BY dtini", OrderBy)
            
    Set Rst = RegistroBuscar(sqlEx)
    If Rst.BOF And Rst.EOF Then
        Else
            
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                With msfgContratos
                    DoEvents
                    .Rows = .Rows + 1
                    
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("ID")
                    .TextMatrix(.Rows - 1, 1) = cNull(Rst.Fields("numcontrato"))
                    .TextMatrix(.Rows - 1, 2) = cNull(Rst.Fields("dtini"))
                    .TextMatrix(.Rows - 1, 3) = cNull(Rst.Fields("dtFin"))

                    .TextMatrix(.Rows - 1, 4) = PgDadosCliente(cNull(Rst.Fields("idcliente"))).Nome
'                    .TextMatrix(.Rows - 1, 4) = rst.Fields("Emissao")
'                    .TextMatrix(.Rows - 1, 5) = IIf(IsNull(rst.Fields("NumDuplicata")), " ", rst.Fields("NumDuplicata"))
'                    .TextMatrix(.Rows - 1, 6) = ConvMoeda(ChkVal(IIf(IsNull(rst.Fields("vlDuplicata")), "0", rst.Fields("vlDuplicata")), 0, cDecMoeda)) & IIf(rst.Fields("ContaPR") = "P", "D", "C")
'                    .TextMatrix(.Rows - 1, 7) = rst.Fields("Vencimento")
'                    .TextMatrix(.Rows - 1, 8) = Date
'
                    'dtCalc = IIf(IsNull(rst.Fields("dataquitacao")), Date, rst.Fields("dataquitacao")) 'Checa a data de quitacao do documento
                    '.TextMatrix(.Rows - 1, 9) = IIf(CDate(dtCalc) - rst.Fields("Vencimento") < 0, 0, CDate(dtCalc) - rst.Fields("Vencimento"))
                    Rst.MoveNext
                End With
            Loop
    End If
    Rst.Close
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    With frmMenu
        .Width = Me.Width - 450
    End With
    
    With frmContratos
        
        .Left = frmMenu.Left
        .Top = frmMenu.Height + 450
        .Width = frmMenu.Width
        .Height = Me.ScaleHeight - (frmMenu.Height + 700)
    
        
        msfgContratos.Height = .Height - 700 '- (lblAviso.Height + 300)
        msfgContratos.Width = .Width - 100
        
    End With
    'lblAviso.Top = msfgContas.Top + msfgContas.Height
'
'    frmGrafico.Top = frmContas.Top + frmContas.Height
'
'    frmtDuplicata.Top = frmGrafico.Top
'
'
'    txtObs.Top = frmGrafico.Top + 80
'    txtObs.Width = Me.Width - (300 + txtObs.Left)
    pb.Top = frmContratos.Top + 100
    pb.Left = msfgContratos.Width - (pb.Width + 200)

End Sub

Private Sub msfgContratos_Click()
On Error GoTo ZeraId
    idContrato = msfgContratos.TextMatrix(msfgContratos.RowSel, 0)
    Exit Sub
ZeraId:
    idContrato = 0
End Sub

Private Sub msfgContratos_DblClick()
    If msfgContratos.MouseRow = 0 Then
            OrderBy = Trim(msfgContratos.TextMatrix(0, msfgContratos.MouseCol))
            Select Case LCase(OrderBy)
                Case "id"
                    OrderBy = " ORDER BY id"
                Case "Num.Contrato"
                    OrderBy = " ORDER BY numcontrato"
                Case "cliente"
                    OrderBy = " ORDER BY nmCliente"
                Case Else
                    OrderBy = " ORDER BY dtini"
            End Select
            AtualizarLista
        Else
            'Edita Contrato
             
             Alterar
    End If
End Sub
Private Sub Incluir()
    idContrato = 0
    formContrato.loadForm 0
End Sub
Private Sub Alterar()
    If idContrato = 0 Then
        MsgBox "Selecione um contrato!", vbInformation, App.EXEName
        Exit Sub
    End If
    formContrato.loadForm idContrato
End Sub
Private Sub Excluir()
    
    If idContrato = 0 Then
        MsgBox "Selecione um contrato!", vbInformation, App.EXEName
        Exit Sub
    End If
    
    Dim msgExcluir As String
    msgExcluir = "Deseja realmente excluir o contrato " & idContrato & "?"
    
    If MsgBox(msgExcluir, vbYesNo + vbInformation, App.EXEName & " - Excluir") = vbYes Then
        RegistroExcluir "contrato", "id=" & idContrato
        RegistroExcluir "contratomateriais", "id=" & idContrato
        RegistroExcluir "contratofuncionaros", "id=" & idContrato
        AtualizarLista
    End If
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

Private Sub optPerido_Click(Index As Integer)
    If Index = 2 Then
            dtpDtInicio.Enabled = False
            dtpDtFinal.Enabled = False
        Else
            dtpDtInicio.Enabled = True
            dtpDtFinal.Enabled = True
    End If
End Sub
Private Sub cboCliente_Click()
    If Trim(cboCliente.Text) = "" Then
        idCliente = 0
        Exit Sub
    End If
    PesquisarCliente "ID", Trim(Left(Trim(cboCliente.Text), 6)), "N"
End Sub
Private Sub cboCliente_KeyPress(KeyAscii As Integer)
    idCliente = 0
End Sub

Private Sub cboCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarCliente
    End If
End Sub
Private Sub cboCliente_DropDown()
    Dim Rst As Recordset
    idCliente = 0
    Set Rst = RegistroBuscar("SELECT * FROM Clientes WHERE ID_Empresa = " & ID_Empresa & " AND xNome LIKE '" & cboCliente.Text & "%'")
    If Rst.BOF And Rst.EOF Then
            cboCliente.Clear
            Exit Sub
        Else
            cboCliente.Clear
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCliente.AddItem Left(String(6, "0"), 6 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & _
                                   " - " & _
                                   Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If
End Sub
Private Sub PesquisarCliente(Optional sCampo As String, Optional sBusca As String, Optional SN As String)
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    If Trim(sCampo) = "" Then
        sBusca = formBuscar.IniciarBusca("Clientes", , , , , "Status='Ativo'") ', "xNome,xlgr,nro,xcpl,xbairro,xmun,uf,fone")
        sCampo = "Id"
        SN = "N"
        If Trim(sBusca) = 0 Then Exit Sub
    End If
    If SN = "N" Then
            sSQL = "SELECT * FROM Clientes WHERE ID_Empresa = " & ID_Empresa & " AND Status='Ativo' AND " & sCampo & " = '" & sBusca & "'"
        Else
            sSQL = "SELECT * FROM Clientes WHERE ID_Empresa = " & ID_Empresa & " AND Status='Ativo' AND " & sCampo & " = " & sBusca
    End If
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado"
        Else
            Rst.MoveFirst
            idCliente = Rst.Fields("Id")
            cboCliente.Text = Trim(Rst.Fields("xNome"))
           
            
    End If
    Rst.Close
End Sub
Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualiza"
            AtualizarLista
        Case "Incluir"
            Incluir
        Case "Alterar"
            Alterar
        Case "Excluir"
            Excluir
'        Case "Pesquisar"
'            PesquisarRegistro
'        Case "Salvar"
'            Salvar
'        Case "Cancelar"
'            Cancelar
'        Case "Manutenção da Tabela"
'            formManutencaoTabelas.IniciarManutencao Me, "SELECT * FROM Clientes"
    End Select
    
End Sub

