VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFinanceiroContasPRGerenciador 
   Caption         =   "Financeiro - Contas a Pagar e Receber"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame frmContas 
      Height          =   5055
      Left            =   60
      TabIndex        =   9
      Top             =   420
      Width           =   14895
      Begin MSComCtl2.DTPicker dtpCalc 
         Height          =   315
         Left            =   7620
         TabIndex        =   12
         Top             =   3120
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   165085185
         CurrentDate     =   40658
      End
      Begin MSFlexGridLib.MSFlexGrid msfgContas 
         Height          =   4755
         Left            =   60
         TabIndex        =   10
         ToolTipText     =   "Duplo clique na coluna para ordenar..."
         Top             =   180
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8387
         _Version        =   393216
         Cols            =   15
         AllowUserResizing=   1
         FormatString    =   $"formFinanceiroContasPR.frx":0000
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
   Begin VB.Frame frmGrafico 
      Caption         =   "Grafico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   3795
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   1875
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   3675
         Begin MSChart20Lib.MSChart msc 
            Height          =   2715
            Left            =   60
            OleObjectBlob   =   "formFinanceiroContasPR.frx":0106
            TabIndex        =   8
            ToolTipText     =   "Click no Grafico para alterar o modo de visualização..."
            Top             =   -540
            Width           =   3675
         End
      End
   End
   Begin VB.TextBox txtObs 
      Height          =   1875
      Left            =   8340
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "formFinanceiroContasPR.frx":20E5
      Top             =   6420
      Width           =   6435
   End
   Begin VB.Frame frmtDuplicata 
      Caption         =   "Duplicata"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   4080
      TabIndex        =   0
      Top             =   6360
      Width           =   3735
      Begin VB.TextBox txtSomaDuplicataC 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Text            =   "0,00"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtSomaDuplicataD 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Text            =   "0,00"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Credito Nominal:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Debito Nominal:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Width           =   1155
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filtro"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alterar Documento"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Documento"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Documento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Lista Simplificada"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Lista Completa"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Lista Plano de Contas"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Registrar Fatura"
            ImageIndex      =   15
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   12060
         TabIndex        =   13
         Top             =   60
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
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
            NumListImages   =   16
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":20EB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":253D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":2857
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":30E9
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":433B
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":4C15
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":54A7
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":5D39
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":6F8B
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":72A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":75BF
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":79B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":7A14
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":88EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":894C
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPR.frx":96D5
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Caption         =   "# - cobranca registrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   7680
      Width           =   3735
   End
End
Attribute VB_Name = "formFinanceiroContasPRGerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg       As Long
Dim sSQL        As String
Dim sSQLExec    As String 'Armazena ultima instrucao sql do filtro
Dim sDuplTc     As String 'Total Credito
Dim sDuplTd     As String 'Total Debito
Dim sMMc        As String 'Multa Mora Credito
Dim sMMd        As String 'Multa Mora Debito
Dim sDuplAc     As String 'Atualizada Credito
Dim sDuplAd     As String 'Atualizada Debito
Dim sDuplQc     As String 'Quitadas Credito
Dim sDuplQd     As String 'Quitadas Debito
Dim OrderBy     As String 'Ordena a listagem de acordo com a coluna clicada



'Dim sDuplC  As String 'Credito
'Dim sDuplD  As String 'Debito

Private Sub CalcularSomas()
    Dim i   As Integer
    Dim inf As String
    
    sDuplTc = "0"
    sMMc = "0"
    sDuplAc = "0"
    sDuplQc = "0"
    
    sDuplTd = "0"
    sMMd = "0"
    sDuplAd = "0"
    sDuplQd = "0"
    
    With msfgContas
        For i = 1 To .Rows - 1
            status (.Rows - 1)
            inf = Right(.TextMatrix(i, 6), 1)
            If inf = "C" Then
                    sDuplTc = Val(ChkVal(sDuplTc, 0, cDecMoeda)) + Val(ChkVal(Mid(.TextMatrix(i, 6), 1, Len(.TextMatrix(i, 6)) - 1), 0, cDecMoeda))
                    sMMc = Val(ChkVal(sMMc, 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 10), 0, cDecMoeda))
                    sDuplAc = Val(ChkVal(sDuplAc, 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 11), 0, cDecMoeda))
                    sDuplQc = Val(ChkVal(sDuplQc, 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 13), 0, cDecMoeda))
                Else
                    sDuplTd = Val(ChkVal(sDuplTd, 0, cDecMoeda)) + Val(ChkVal(Mid(.TextMatrix(i, 6), 1, Len(.TextMatrix(i, 6)) - 1), 0, cDecMoeda))
                    sMMd = Val(ChkVal(sMMd, 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 10), 0, cDecMoeda))
                    sDuplAd = Val(ChkVal(sDuplAd, 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 11), 0, cDecMoeda))
                    sDuplQd = Val(ChkVal(sDuplQd, 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(i, 13), 0, cDecMoeda))
             End If
        Next
    End With
End Sub


Private Sub AtualizarLista()
    'Lista os boletos no GRID
    On Error Resume Next
    Dim Rst         As Recordset
    Dim dtCalc      As String
    
    
    If Trim(sSQL) = "" Then Exit Sub

    DoEvents
    msfgContas.Rows = 1
    dtpCalc.Visible = False
    txtObs.Text = ""
    txtSomaDuplicataC.Text = ""
    txtSomaDuplicataD.Text = ""
    
    sSQLExec = sSQL & IIf(Trim(OrderBy) = "", " ORDER BY emissao", OrderBy)
            
    Set Rst = RegistroBuscar(sSQLExec)
    If Rst.BOF And Rst.EOF Then
            'MsgBox "Nenum documento Localizado"
            txtSomaDuplicataC.Text = ConvMoeda("0")
            txtSomaDuplicataD.Text = ConvMoeda("0")
            Exit Sub
        Else
            If Rst.RecordCount >= 500 Then
                'Verifica se o resultado tem mais de 500 reg
                If MsgBox("Esta consulta gerou mais de 500 registros. Deseja realmente visualizar?", vbYesNo, App.EXEName) = vbNo Then
                    Rst.Close
                    Exit Sub
                End If
            End If
                
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                With msfgContas
                    DoEvents
                    .Rows = .Rows + 1
                    
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("ID")
                    .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rst.Fields("Nome")), " ", Rst.Fields("Nome"))
                    
                    'Tipo
                    .TextMatrix(.Rows - 1, 2) = pgDadosTipoDocumento(cNull(Rst.Fields("tpDocumento"))).Sigla
                    'FV
                    .TextMatrix(.Rows - 1, 3) = IIf(cNull(Rst.Fields("FixoVariavel")) = "", "", cNull(Rst.Fields("FixoVariavel")))
                    .TextMatrix(.Rows - 1, 4) = Rst.Fields("Emissao")
                    
                    'NUMERO DA DUPLICATA
                    Dim nDup As String
                    nDup = IIf(IsNull(Rst.Fields("NumDuplicata")), " ", Rst.Fields("NumDuplicata"))
                    nDup = nDup & IIf(Rst.Fields("gerarCNAB240") <> 0, "#", "")
                    .TextMatrix(.Rows - 1, 5) = nDup
                    
                    
                    .TextMatrix(.Rows - 1, 6) = ConvMoeda(ChkVal(IIf(IsNull(Rst.Fields("vlDuplicata")), "0", Rst.Fields("vlDuplicata")), 0, cDecMoeda)) & IIf(Rst.Fields("ContaPR") = "P", "D", "C")
                    .TextMatrix(.Rows - 1, 7) = Rst.Fields("Vencimento")
                    '.TextMatrix(.Rows - 1, 8) = Date
                    
                    dtCalc = IIf(IsNull(Rst.Fields("dataquitacao")), Date, Rst.Fields("dataquitacao")) 'Checa a data de quitacao do documento
                    .TextMatrix(.Rows - 1, 9) = IIf(CDate(dtCalc) - Rst.Fields("Vencimento") < 0, 0, CDate(dtCalc) - Rst.Fields("Vencimento"))
                    
                    .TextMatrix(.Rows - 1, 10) = ChkVal(AtualizaCobranca(Rst.Fields("ID"), dtCalc).vCalcFin, 0, cDecMoeda) 'ConvMoeda(Val(AtualizaCobranca(Rst.fields("ID"), dtCalc).vMulta) + Val(AtualizaCobranca(Rst.fields("ID"), dtCalc).vMora))
                    '.TextMatrix(.Rows - 1, 10) = ConvMoeda(.TextMatrix(.Rows - 1, 10))
                    
                    .TextMatrix(.Rows - 1, 11) = ConvMoeda(Val(ChkVal(Left(.TextMatrix(.Rows - 1, 6), Len(.TextMatrix(.Rows - 1, 6)) - 1), 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(.Rows - 1, 10), 0, cDecMoeda)))

                    .TextMatrix(.Rows - 1, 12) = IIf(IsNull(Rst.Fields("DataQuitacao")), " ", Rst.Fields("DataQuitacao"))
                    
                    If Trim(.TextMatrix(.Rows - 1, 12)) <> "" Then
                            .TextMatrix(.Rows - 1, 13) = ConvMoeda(ChkVal(IIf(IsNull(Rst.Fields("VlCobrado")), "0", Rst.Fields("VlCobrado")), 0, cDecMoeda))
                            .TextMatrix(.Rows - 1, 8) = .TextMatrix(.Rows - 1, 12)
                        Else
                            .TextMatrix(.Rows - 1, 8) = Date
                    End If
                    .TextMatrix(.Rows - 1, 14) = IIf(IsNull(Rst.Fields("Obs")), " ", Rst.Fields("Obs"))
                    
         
                 
                    'If Right(.TextMatrix(.Rows - 1, 4), 1) = "C" Then
                    '    sDuplC = Val(ChkVal(sDuplC, 0, 2)) + Val(ChkVal(Left(.TextMatrix(.Rows - 1, 4), Len(.TextMatrix(.Rows - 1, 4)) - 1), 0, 2))
                    'End If
                    'If Right(.TextMatrix(.Rows - 1, 4), 1) = "D" Then
                    '    sDuplD = Val(ChkVal(sDuplD, 0, 2)) + Val(ChkVal(Left(.TextMatrix(.Rows - 1, 4), Len(.TextMatrix(.Rows - 1, 4)) - 1), 0, 2))
                    'End If
                    '##################################################################################
                    '### MUDAR A COR DO TEXTO DE ACORDO COM O DOCUMENTO
                    '##################################################################################
                    .Row = .Rows - 1
                    .Col = 1
                    .ColSel = .Cols - 1
                    .FillStyle = flexFillRepeat
                    .CellForeColor = IIf(Rst.Fields("ContaPR") = "P", vbRed, vbBlue)
                    
                    Rst.MoveNext
                End With
            Loop
    End If
    CalcularSomas
    txtSomaDuplicataC.Text = ConvMoeda(sDuplTc)
    txtSomaDuplicataD.Text = ConvMoeda(sDuplTd)
    Rst.Close
    MontarGrafico
    
    
    

End Sub

Private Sub MontarGrafico()
'Row - especifica a linha corrente
'RowCount - Determina o número de sequência de dados
'RowLabel - define o rótulo de dados da linha corrente
'Data - Permite a leitura e a atribuição dos valores de dados ao gráfico
'Column - Define a coluna ativa.
'ColumnCount - Define o número de colunas ativas
'ColumnLabel - Define a legenda para a coluna ativa
'BorderStyle - Define a borda do gráfico
'ChartData - Permite atribuir valores ás sequências de dados a partir de uma matriz (array) de duas dimensões
    With msc
        .RowCount = 1
        .ColumnCount = 2
        .Column = 1
        .Data = ChkVal(txtSomaDuplicataD.Text, 0, 0)
        .Column = 2
        .Data = ChkVal(txtSomaDuplicataC.Text, 0, 0)
    End With
End Sub




Private Sub registrarBoletoBBCobranca(idFatura As Long)
    If IdReg = 0 Then Exit Sub
    
    'Verifica o valor da coluna
    'If (Len(Trim(PgDadosFinanceiroFatura(idFatura).FixoVariavel)) > 0) Then
    '    MsgBox "Fatura nao pode ser registrada!", vbExclamation, "A1 - Aviso"
    '    Exit Sub
    'End If
    'Verifica se e um boleto BB
    If pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).banco <> 1 Then
        MsgBox "Fatura nao pode ser registrada! Somenta fatura do Banco do Brasil.", vbExclamation, "A1 - Aviso"
        Exit Sub
    End If
    
    If MsgBox("Confirma registrar boleto?", vbYesNo + vbQuestion, "Registro de boleto") = vbNo Then Exit Sub
    '--------------------------------------------------------------------------------------
    'Enviar dados para registrar Boleto
    '--------------------------------------------------------------------------------------
    
    Dim bbCob As New BBCobranca
    Dim jsonBoleto As String
    
    Dim cnpjBeneficiario As String
    Dim nomeBeneficiario As String
    
    cnpjBeneficiario = PgDadosEmpresa(ID_Empresa).CNPJ
    nomeBeneficiario = PgDadosEmpresa(ID_Empresa).Nome
    
    Dim Convenio As String
    Dim carteira As String
    'Dim convenio As String
    'Dim idFatura As Long
    Dim DiasProtesto As String
    Dim carteiraVariacao As String
    Dim tipoConta As String
    Dim Valor As String
    Dim vDeducao As String
    Dim emissao As String
    Dim Vencimento As String
    Dim nFatura As String
    Dim nDuplicata As String
   

    Convenio = pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).Convenio
    carteira = pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).carteira
    carteiraVariacao = pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).Variacao
    tipoConta = pgDadosConta(PgDadosFinanceiroFatura(idFatura).idConta).Tipo
    Valor = PgDadosFinanceiroFatura(idFatura).vlCobrado
    vDeducao = PgDadosFinanceiroFatura(idFatura).Deducoes
    emissao = PgDadosFinanceiroFatura(idFatura).emissao
    Vencimento = PgDadosFinanceiroFatura(idFatura).Vencimento
    nFatura = PgDadosFinanceiroFatura(idFatura).NumFatura
    nDuplicata = PgDadosFinanceiroFatura(idFatura).NumDuplicata
    DiasProtesto = PgDadosFinanceiroFatura(idFatura).DiasProtesto
    
    
    
    Dim vJurosMora As String
    vJurosMora = cobCalcMora(Valor, 1, 2, "D")
     
    Dim vMulta As String
    vMulta = cobCalcMulta(Valor, 0, 1)
    
    Dim Sacado As String
    Dim sacID As Integer
    Dim sacTpInscricao As String
    Dim sacCNPJ As String
    Dim sacNome As String
    Dim sacLgr As String
    Dim sacCep As String
    Dim sacBairro As String
    Dim sacMun As String
    Dim sacUF As String
    Dim sacFone As String
    Dim sacMail As String
    
    sacID = PgDadosFinanceiroFatura(idFatura).IDSacado
    sacTpInscricao = "2"
    sacNome = PgDadosCliente(sacID).Nome
    sacCNPJ = PgDadosCliente(sacID).Doc
    sacLgr = PgDadosCliente(sacID).Lgr & " " & PgDadosCliente(sacID).Nro
    sacBairro = PgDadosCliente(sacID).Bairro
    sacMun = PgDadosCliente(sacID).Mun
    sacUF = PgDadosCliente(sacID).uf
    sacFone = PgDadosCliente(sacID).Fone
    sacMail = PgDadosCliente(sacID).Mail
    
    Sacado = bbCob.jsonSacado(sacTpInscricao, _
                              sacCNPJ, _
                              sacNome, _
                              sacLgr, _
                              sacCep, _
                              sacMun, _
                              sacBairro, _
                              sacUF, _
                              sacFone, _
                              sacMail)

     
    jsonBoleto = bbCob.GerarBoletoBB(Convenio:=Convenio, _
                                    carteira:=carteira, _
                                    carteiraVariacao:=carteiraVariacao, _
                                    tipoConta:=tipoConta, _
                                    dataEmissao:=emissao, _
                                    DataVencimento:=Vencimento, _
                                    nFatura:=nFatura, _
                                    nDuplicata:=nDuplicata, _
                                    Valor:=Valor, _
                                    vDeducao:=vDeducao, _
                                    vMulta:=vMulta, _
                                    vJuros:=vJurosMora, _
                                    DiasProtesto:="5", _
                                    Sacado:=Sacado, _
                                    cnpjBeneficiario:=cnpjBeneficiario, _
                                    nomeBeneficiario:=nomeBeneficiario, _
                                    NossoNumero:=bbCob.GerarNossoNumero(Convenio, idFatura), _
                                    smsg:="MENSAGEM")
            
    'Debug.Print jsonBoleto

  

  'Gravar log tmp checagem
    ExcluirFile App.Path & "\bbCobranca\boleto.txt"
    'ExcluirFile PgDadosConfig.pFileArmazenamento & "\boleto-001.txt"
    
    grvFile App.Path & "\bbCobranca\boleto.txt", jsonBoleto
    'grvFile PgDadosConfig.pFileArmazenamento & "\boleto-001.txt", "|Id:" & Id & vbCrLf & _
                                "|NN:" & NossoNumero & vbCrLf & _
                                "|LD:" & LinhaDigitavel & vbCrLf & _
                                "|CB:" & CodigoBarras
    Dim sStatus As String
    sStatus = bbCobrancaEnviarBoletoAPI
    '--------------------------------------------------------------------------------------
    'Grava boleto como registrado
    '--------------------------------------------------------------------------------------
  
    If MsgBox("Confirma marcar esta cobranca como registrada junto ao banco?", vbQuestion + vbYesNo, "A1 - Aviso") = vbYes Then
        Dim vDados(1)   As Variant
        Dim cReg        As Integer
        
        Dim criterio As String
        cReg = 0
        vDados(cReg) = Array("FixoVariavel", "R", "S"): cReg = cReg + 1
        cReg = cReg - 1
        criterio = "id=" & idFatura
        RegistroAlterar "financeirocontasprcadastro", vDados, cReg, criterio
    End If
    AtualizarLista
End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
            Unload Me
        Else
            AtualizarLista
    End If
End Sub



Private Sub msfgContas_DblClick()
    
    If msfgContas.MouseRow = 0 Then
        OrderBy = msfgContas.TextMatrix(0, msfgContas.MouseCol)
        Select Case LCase(OrderBy)
            Case "nome"
                OrderBy = " ORDER BY nome"
            Case "tipo"
                OrderBy = " ORDER BY tpdocumento"
            Case "emissao"
                OrderBy = " ORDER BY emissao"
            Case "duplicata"
                OrderBy = " ORDER BY numduplicata"
            Case "valor duplicata"
                OrderBy = " ORDER BY vlduplicata"
            Case "vencimento"
                OrderBy = " ORDER BY vencimento"
            Case "data quitação"
                OrderBy = " ORDER BY dataquitacao"
            Case Else
                OrderBy = " ORDER BY emissao"
        End Select
        AtualizarLista
        
    End If
    
End Sub

Private Sub msfgContas_EnterCell()

    With msfgContas
        If .Enabled = False Then Exit Sub
        If .MouseCol <> 8 Or .MouseRow = 0 Then
            dtpCalc.Visible = False
            Exit Sub
        End If
        dtpCalc.Top = .Top + .CellTop
        dtpCalc.Left = .Left + .CellLeft
        dtpCalc.Width = .CellWidth
        dtpCalc.Height = .CellHeight
        dtpCalc.Value = IIf(Trim(.TextMatrix(.Row, 8)) = "", Date, .TextMatrix(.Row, 8))
        '22.02.2017 - não habilita edição caso o titulo esteja quitado
        If Len(Trim(.TextMatrix(.Row, 12))) = 0 Then
                dtpCalc.Visible = True
            Else
                dtpCalc.Visible = False
        End If
    End With
End Sub

Private Sub msfgContas_LeaveCell()
    Dim dtCalc      As String 'Registra a data para o qual o boleto ta sendo calculado

    
    With msfgContas
        If .Enabled = False Then Exit Sub
        If .Col <> 8 And .Row = 1 Or dtpCalc.Visible = False Then
            Exit Sub
        End If
        .TextMatrix(.Row, 8) = dtpCalc.Value
        dtCalc = .TextMatrix(.Row, 8) 'IIf(IsNull(Rst.Fields("dataquitacao")), Date, Rst.Fields("dataquitacao")) 'Checa a data de quitacao do documento
        .TextMatrix(.Row, 9) = AtualizaCobranca(IdReg, dtCalc).DiasVencidos  'IIf(CDate(dtCalc) - CDate(.TextMatrix(.Row, 7)) < 0, 0, CDate(dtCalc) - CDate(.TextMatrix(.Row, 7)))

        .TextMatrix(.Row, 10) = ConvMoeda(AtualizaCobranca(IdReg, dtCalc).vCalcFin) '
        '.TextMatrix(.Row, 10) = ConvMoeda(Val(AtualizaCobranca(IdReg, dtCalc).vMulta) + Val(AtualizaCobranca(IdReg, dtCalc).vMora))
        '.TextMatrix(.Row, 11) = ConvMoeda(.TextMatrix(.Row, 8))

        .TextMatrix(.Row, 11) = ConvMoeda(Val(ChkVal(Left(.TextMatrix(.Row, 6), Len(.TextMatrix(.Row, 6)) - 1), 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(.Row, 10), 0, cDecMoeda)))
    End With
End Sub


Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualizar"
            AtualizarLista
            
        Case "Filtro"
            filtro
            
        Case "Alterar Documento"
            AlterarDocumento
            
        Case "Imprimir Documento"
            ImprimirDocumento
            
        Case "Registrar Fatura"
            registrarBoletoBBCobranca (IdReg)
            'mntArqCNAB240
'
'        Case "Imprimir Listagem Completa"
'            ImprimirListagemCompleta
            
    End Select
End Sub
Private Sub tbMenu_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
   Select Case ButtonMenu.Text
        Case "Imprimir Documento"
            ImprimirDocumento
        Case "Lista Simplificada"
            ImprimirListagem
        Case "Lista Completa"
            ImprimirListagemCompleta
         Case "Lista Plano de Contas"
            ImprimirListagemPlanoContas
            
    End Select
End Sub
Private Sub mntArqCNAB240()

    If MsgBox("Sistema entrara em modo operante para registrar os boletos junto ao banco. Isso pode levar alguns minutos dependendo do numero de boletos solicitado. Deseja continuar?", vbQuestion + vbYesNo, "A - Aviso") = vbNo Then
        MsgBox "Operacao cancelada.", vbInformation, "Aviso"
        Exit Sub
    End If
    'Dim vDados(10)   As Variant
    'Dim cReg        As Integer
    'Dim i As Integer
    'Dim criterio As String
    'Dim lote As String
    
    'lote = "100"
    
    'With msfgContas
    '    .Enabled = False
    '    For i = 1 To .Rows - 1
    '        cReg = 0
    '        vDados(cReg) = Array("gerarcnab240", lote, "N"): cReg = cReg + 1
    '        cReg = cReg - 1
    '        criterio = "id=" & .TextMatrix(i, 0)
    '        RegistroAlterar "financeirocontasprcadastro", vDados, cReg, criterio
    '    Next
    '    .Enabled = True
    'End With
    
    
    'Variaveis necessarias
    'Dim DtIni As Date
    'Dim DtFin As Date
    'Dim conta As Integer
    'DtIni = "01/03/2017"
    'DtFin = "30/03/2017"
    'conta = 1
    'cnab240 DtIni, DtFin, conta, lote
End Sub
Private Sub ImprimirListagemPlanoContas()
    If chkAcesso(Me, "i") = False Then Exit Sub
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    Dim i           As Integer
    Dim ii          As Integer
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    MontarTabelaTemporaria
    
    'Dados do campo
    Dim dc As String
    
    With msfgContas
        .Enabled = False
        For i = 1 To .Rows - 1
            status (.Rows - 1)
            cReg = 0
            For ii = 1 To .Cols - 1
                dc = msfgContas.TextMatrix(i, ii)
                If Left(dc, 2) = "R$" Then
                    dc = ChkVal(Mid(dc, 3, Len(dc)), 0, cDecMoeda)
                End If
                vReg(cReg) = Array(RS(.TextMatrix(0, ii)), dc, "S"): cReg = cReg + 1
            Next
            Dim idPC As String
            idPC = PgDadosFinanceiroFatura(msfgContas.TextMatrix(i, 0)).idPlanoContas
            vReg(cReg) = Array(RS("codPlanContas"), PgDadosPlanoContas("id", idPC).Codigo, "S"): cReg = cReg + 1
            cReg = cReg - 1
            RegistroIncluir "tmp_Titulos", vReg, cReg
        Next
        
        .Enabled = True
    End With
    
    
    
    MontarTabelaTemporariaPC
    
    Dim sSQL2 As String
    Dim Rst2 As Recordset
    Dim vl  As String
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
                
                '29.06.2016 - Se o item for um totalizador usar like
                If PgDadosPlanoContas("codigo", "'" & Rst.Fields("codigo") & "'").totalizador = 1 Then
                        sSQL2 = "select codplancontas, sum(ValorAtualizado) AS tValor FROM tmp_titulos"
                        sSQL2 = sSQL2 & " WHERE codplancontas = '" & Rst.Fields("codigo") & "'"
                        sSQL2 = sSQL2 & " OR codplancontas LIKE '" & Rst.Fields("codigo") & ".%'"
                    Else
                        sSQL2 = "select codplancontas, sum(ValorAtualizado) AS tValor FROM tmp_titulos WHERE codplancontas = '" & Rst.Fields("codigo") & "'" 'group by codplancontas"
                End If
                Set Rst2 = RegistroBuscar(sSQL2)
                If Rst2.BOF And Rst2.EOF Then
                        vl = "0"
                    Else
                        Rst2.MoveFirst
                        vl = IIf(cNull(Rst2.Fields("tValor")) = "", 0, cNull(Rst2.Fields("tValor")))
                End If
                Rst2.Close
                Dim descrPC As String
                descrPC = cNull(Rst.Fields("descricao")) 'PgDadosPlanoContas("codigo", "'" & Rst.Fields("descricao") & "'").Descricao
        
                cReg = 0
                'vReg(cReg) = Array("codigo", Rst.Fields("codplancontas"), "S"): cReg = cReg + 1
                'vReg(cReg) = Array("descricao", descrPC, "S"): cReg = cReg + 1
                vReg(cReg) = Array("valor", Replace(ChkVal(vl, 0, cDecMoeda), ".", ","), "S"): cReg = cReg + 1
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
        Else
        Set rptListaTitulosPC.DataSource = Rst3.DataSource
        'rptListaTitulos.Sections("Section5").Controls.Item("lblCred").Caption = ConvMoeda(sDuplAc)
        'rptListaTitulos.Sections("Section5").Controls.Item("lblDeb").Caption = ConvMoeda(sDuplAd)
        
        'Credito
        'rptListaTitulos.Sections("Section5").Controls.Item("lblValorDuplicataC").Caption = ConvMoeda(sDuplAc)
        'rptListaTitulos.Sections("Section5").Controls.Item("lblValorDuplicataD").Caption = ConvMoeda(sDuplAd)
        
        rptListaTitulosPC.Show 1
    End If
    Rst3.Close
    
End Sub
Private Sub ImprimirListagem()
    If chkAcesso(Me, "i") = False Then Exit Sub
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    Dim i           As Integer
    Dim ii          As Integer
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    MontarTabelaTemporaria
    
    With msfgContas
        For i = 1 To .Rows - 1
            cReg = 0
            For ii = 1 To .Cols - 1
                
                vReg(cReg) = Array(RS(.TextMatrix(0, ii)), msfgContas.TextMatrix(i, ii), "S"): cReg = cReg + 1
                
            Next
            cReg = cReg - 1
            RegistroIncluir "tmp_titulos", vReg, cReg
        Next
        'Solicitante: Jorge Marques
        'Alterado por: Leonardo Aquino 20.02.2015
        'Inclusao de update para apagar os campos calc para e dias venc caso o
        'Titulo esteja em dia
        BD.Execute "UPDATE tmp_titulos SET Calcpara='', DiasVencidos='' WHERE DiasVencidos='0'"
        'FIM DA ALTERACAO
        
    End With
    CalcularSomas
    
    'sSQL = "SELECT *,CONCAT(valorDuplicata , CD) as vDuplicata FROM tmp_Titulos"
    sSQL = "SELECT * FROM tmp_titulos"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
        Set rptListaTitulos.DataSource = Rst.DataSource
        rptListaTitulos.Sections("Section5").Controls.Item("lblCred").Caption = ConvMoeda(sDuplAc)
        rptListaTitulos.Sections("Section5").Controls.Item("lblDeb").Caption = ConvMoeda(sDuplAd)
        
        'Credito
        rptListaTitulos.Sections("Section5").Controls.Item("lblValorDuplicataC").Caption = ConvMoeda(sDuplAc)
        rptListaTitulos.Sections("Section5").Controls.Item("lblValorDuplicataD").Caption = ConvMoeda(sDuplAd)
        
        rptListaTitulos.Show 1
    End If
    Rst.Close
        
    
End Sub
Private Sub ImprimirListagemCompleta()
    If chkAcesso(Me, "i") = False Then Exit Sub
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    Dim i           As Integer
    Dim ii          As Integer
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    MontarTabelaTemporaria
    
    With msfgContas
        .Enabled = False
        For i = 1 To .Rows - 1
            cReg = 0
            For ii = 1 To .Cols - 1
                vReg(cReg) = Array(RS(.TextMatrix(0, ii)), msfgContas.TextMatrix(i, ii), "S"): cReg = cReg + 1
            Next
            cReg = cReg - 1
            RegistroIncluir "tmp_Titulos", vReg, cReg
        Next
        .Enabled = True
    End With
    CalcularSomas
    
    sSQL = "SELECT * FROM tmp_titulos"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
        Set rptListaTitulosGeral.DataSource = Rst.DataSource
        
        rptListaTitulosGeral.Orientation = 2 ' rptOrientLandscape
        
        rptListaTitulosGeral.Sections("Section5").Controls.Item("lblCred").Caption = txtSomaDuplicataC.Text
        rptListaTitulosGeral.Sections("Section5").Controls.Item("lblDeb").Caption = txtSomaDuplicataD.Text
        'CREDITO
        rptListaTitulosGeral.Sections("Section5").Controls.Item("lblValorDuplicataC").Caption = ConvMoeda(Val(ChkVal(sDuplTc, 0, cDecMoeda))) ' + Val(ChkVal(sDuplTd, 0, cDecMoeda)))
        rptListaTitulosGeral.Sections("Section5").Controls.Item("lblMultaMoraC").Caption = ConvMoeda(Val(ChkVal(sMMc, 0, cDecMoeda))) ' + Val(ChkVal(sMMd, 0, cDecMoeda)))
        rptListaTitulosGeral.Sections("Section5").Controls.Item("lblValorAtualizadoC").Caption = ConvMoeda(Val(ChkVal(sDuplAc, 0, cDecMoeda))) ' + Val(ChkVal(sDuplAd, 0, cDecMoeda)))
        rptListaTitulosGeral.Sections("Section5").Controls.Item("lblValorQuitadoC").Caption = ConvMoeda(Val(ChkVal(sDuplQc, 0, cDecMoeda))) ' + Val(ChkVal(sDuplQd, 0, cDecMoeda)))
        'DEBITO
        rptListaTitulosGeral.Sections("Section5").Controls.Item("lblValorDuplicataD").Caption = ConvMoeda(Val(ChkVal(sDuplTd, 0, cDecMoeda)))
        rptListaTitulosGeral.Sections("Section5").Controls.Item("lblMultaMoraD").Caption = ConvMoeda(Val(ChkVal(sMMd, 0, cDecMoeda)))
        rptListaTitulosGeral.Sections("Section5").Controls.Item("lblValorAtualizadoD").Caption = ConvMoeda(Val(ChkVal(sDuplAd, 0, cDecMoeda)))
        rptListaTitulosGeral.Sections("Section5").Controls.Item("lblValorQuitadoD").Caption = ConvMoeda(Val(ChkVal(sDuplQd, 0, cDecMoeda)))
        
        
        rptListaTitulosGeral.Show 1
    End If
    Rst.Close
        
    
End Sub
Private Sub MontarTabelaTemporaria()

    Dim sCampos     As String
    Dim i           As Integer
    
    BD.Execute "DROP TABLE IF EXISTS tmp_titulos"
    sCampos = ""
    For i = 1 To msfgContas.Cols - 1
        sCampos = sCampos & RS(msfgContas.TextMatrix(0, i)) & " VARCHAR(100) default Null,"
    Next
    
    sCampos = sCampos & " codPlanContas VARCHAR(100) default Null,"
    
    sCampos = "CREATE TABLE IF NOT EXISTS tmp_titulos " & _
              "(Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "UsuID VARCHAR(10) default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               sCampos & " PRIMARY KEY (" & msfgContas.TextMatrix(0, 0) & "))"
    BD.Execute sCampos
End Sub

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
               " PRIMARY KEY (Id))"

    BD.Execute sCampos
    
    '29.06.2016 - Copia os dados da tabela de plano de contas
    BD.Execute "INSERT INTO tmp_titulospc (Id_Empresa, UsuID, codigo, descricao, totalizador) SELECT Id_Empresa, " & ID_Usuario & ", codigo, descricao, totalizador FROM financeiroplanocontas"
End Sub


Private Sub ImprimirDocumento()
    Dim docPrint    As String
    Dim idDoc       As Long
    Dim difDatas        As Integer
    
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If
    
    If msfgContas.Row = 0 Then Exit Sub
    idDoc = msfgContas.TextMatrix(msfgContas.Row, 0)
    docPrint = pgDadosTipoDocumento(PgDadosFinanceiroFatura(idDoc).idTpDoc).Impressao
    
    Select Case Left(docPrint, 2)
        Case "01"
            If PgDadosConfig.ImpBoleto = 1 Then
                    ImprBB_Pre (idDoc)
                Else
                    difDatas = DateDiff("d", PgDadosFinanceiroFatura(idDoc).Vencimento, msfgContas.TextMatrix(msfgContas.Row, 8))
                    If difDatas > 0 And Trim(PgDadosFinanceiroFatura(idDoc).DataQuitacao) = "00:00:00" Then
                        If MsgBox("Deseja imprimir o boleto com os valores atualizado?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                    AtualizarBoleto idDoc, msfgContas.TextMatrix(msfgContas.Row, 8)
                                Else
                                    BoletoBancario (idDoc)
                            End If
                        Else
                            BoletoBancario (idDoc)
                    End If
            End If

        Case "02"
            impDuplicata (idDoc)
        Case "03"
        Case Else
            MsgBox "Erro ao localizar documento de impressão.", vbInformation, "Aviso"
    End Select
    
End Sub

Private Sub AlterarDocumento()
    If IdReg = 0 Then Exit Sub
    formFinanceiroContasPRCadastro.LoadDocumento (IdReg)
End Sub
Private Sub Form_Load()
    
    OrderBy = ""
    Me.Top = 0
    Me.Left = 0
    txtObs.Text = ""
    dtpCalc.Visible = False
    
    sSQL = "SELECT * FROM FinanceiroContasPRCadastro WHERE ID_Empresa = " & ID_Empresa & _
           " AND Emissao >= '" & Format(Date, "YYYY-MM-DD") & "' AND Emissao <= '" & Format(Date, "YYYY-MM-DD") & "'" '& "' AND DataQuitacao IS NULL"
    
End Sub
Private Sub filtro()
    OrderBy = ""
    sSQL = ""
    sSQL = formFinanceiroContasPRFiltro.filtro
    AtualizarLista
End Sub
Public Sub filtroExterno(instSQL As String)
    Me.Show
    sSQL = instSQL
    
    
    'AtualizarLista
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    frmContas.Width = Me.Width - 300
    frmContas.Height = Me.ScaleHeight - (frmGrafico.Height + tbMenu.Height + 200)
    
    msfgContas.Height = frmContas.Height - 300 '- (lblAviso.Height + 300)
    msfgContas.Width = frmContas.Width - 100
    'lblAviso.Top = msfgContas.Top + msfgContas.Height
    
    frmGrafico.Top = frmContas.Top + frmContas.Height

    frmtDuplicata.Top = frmGrafico.Top
    

    txtObs.Top = frmGrafico.Top + 80
    txtObs.Width = Me.Width - (300 + txtObs.Left)
    
    pb.Left = Me.Width - (pb.Width + 200)
    
End Sub


Private Sub msc_Click()
    If msc.chartType = VtChChartType2dPie Then
            msc.chartType = VtChChartType2dCombination
        Else
            msc.chartType = VtChChartType2dPie
    End If
End Sub

Private Sub msfgContas_Click()


    If msfgContas.Row = 0 Then Exit Sub
    IdReg = msfgContas.TextMatrix(msfgContas.Row, 0)

    txtObs.Text = IIf(Trim(msfgContas.TextMatrix(msfgContas.Row, 14)) = "", "", msfgContas.TextMatrix(msfgContas.Row, 14))
    '22.02.2017
    'Caso o titulo esteja quitado nao calcular
    
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub status(Max As Long)
    
    pb.Min = 0
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

Private Sub txtSomaDuplicataC_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtSomaDuplicataD_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
