VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFaturamentoAnalise 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Analise"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   9390
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   60
      TabIndex        =   5
      Top             =   1380
      Width           =   9255
      Begin VB.OptionButton optOpcao 
         Caption         =   " Entrada e Saida  por grupo"
         Height          =   255
         Index           =   12
         Left            =   4920
         TabIndex        =   22
         Top             =   1320
         Width           =   4035
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Analise de rentabilidade das VENDAS"
         Height          =   195
         Index           =   11
         Left            =   4920
         TabIndex        =   21
         Top             =   780
         Width           =   3195
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Listagem de VENDAS por vendedor acumulado mensal"
         Height          =   375
         Index           =   10
         Left            =   4920
         TabIndex        =   18
         Top             =   360
         Width           =   4275
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   " Entrada e Saida  por grupo e sub-grupo no periodo"
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   16
         Top             =   1620
         Width           =   4035
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Listagem de VENDAS por cliente no periodo"
         Height          =   255
         Index           =   8
         Left            =   180
         TabIndex        =   15
         Top             =   1380
         Width           =   3615
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Listagem sintética de VENDAS por vendedor no periodo"
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   14
         Top             =   1095
         Width           =   4575
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Analise de COMPRAS por grupo e sub-grupo no periodo"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   13
         Top             =   2520
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Listagem de COMPRAS no periodo"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   12
         Top             =   2235
         Width           =   5295
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Resumo MENSAL de COMPRAS"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   11
         Top             =   1950
         Width           =   5295
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Resumo DIARIO de COMPRAS"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   10
         Top             =   1425
         Width           =   4155
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Listagem analítica de VENDAS no periodo"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   810
         Width           =   3735
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Resumo MENSAL de VENDAS"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   525
         Width           =   5295
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Resumo DIARIO de VENDAS"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   5295
      End
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   0
      Top             =   480
      Width           =   9255
      Begin VB.ComboBox cboFuncionario 
         Height          =   315
         Left            =   5700
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   300
         Width           =   3435
      End
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   315
         Left            =   900
         TabIndex        =   3
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   117899265
         CurrentDate     =   40665
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   117899265
         CurrentDate     =   40665
      End
      Begin VB.Label Label3 
         Caption         =   "Funcionario:"
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Final:"
         Height          =   195
         Left            =   2700
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   420
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9390
      _ExtentX        =   16563
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
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   60
         Visible         =   0   'False
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1560
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
               Picture         =   "formFaturamentoAnalise.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoAnalise.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoAnalise.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoAnalise.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoAnalise.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoAnalise.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoAnalise.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoAnalise.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoAnalise.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoAnalise.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoAnalise.frx":54D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoAnalise.frx":58CB
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFaturamentoAnalise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TpRelatorio As Integer
Private Function CalcVenda_000(vVenda As String, vNF As String) As String
    CalcVenda_000 = Val(ChkVal(vVenda, 0, cDecMoeda)) + Val(ChkVal(vNF, 0, cDecMoeda))
End Function



Private Sub MontarTabelaTemporaria_010_a(vend As Integer)

    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim Rst1        As Recordset
    Dim vAcumComi   As String
    Dim vAcumVenda  As String
    Dim MesAno      As String
    Dim vReg(10)    As Variant
    Dim cReg        As Integer
    
    sSQL = "SELECT * " & _
            "FROM faturamentonfe " & _
            "WHERE ID_Empresa = " & ID_Empresa & " AND ide_tpNF = 1 AND ide_natOP = 'VENDA' AND canc_nProt IS NULL " & _
            "AND ide_dEmi BETWEEN '" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpFin.Value, "YYYY-MM-DD") & "' " & _
            "AND ger_vendedor=" & vend & " "
            
    'sSQL = "SELECT * " & _
           "FROM faturamentonfe " & _
           "WHERE ID_Empresa = " & ID_Empresa & " AND ide_tpNF = 1 AND ide_natOP = 'VENDA' AND canc_nProt IS NULL " & _
           "AND ide_dEmi BETWEEN '" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpFin.Value, "YYYY-MM-DD") & "' " & _
           "AND ger_vendedor=" & vend & " " & _
           "ORDER BY ger_vendedor"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            MesAno = Format(Rst.Fields("ide_demi"), "MM/YYYY")
            Do Until Rst.EOF
                status Rst.RecordCount
                
                If MesAno <> Format(Rst.Fields("ide_demi"), "MM/YYYY") Then
                    cReg = 0
                    vReg(cReg) = Array("Vendedor", Rst.Fields("Ger_Vendedor"), "N"): cReg = cReg + 1
                    vReg(cReg) = Array("Mes", "01/" & MesAno, "D"): cReg = cReg + 1
                    vReg(cReg) = Array("vVenda", vAcumVenda, "S"): cReg = cReg + 1
                    vReg(cReg) = Array("vComissao", vAcumComi, "S"): cReg = cReg + 1
                    cReg = cReg - 1
                    RegistroIncluir "tmp_faturamentoanalise", vReg, cReg
                    MesAno = Format(Rst.Fields("ide_demi"), "MM/YYYY")
                    vAcumComi = 0
                    vAcumVenda = 0
                End If
                
                
                sSQL = "SELECT * FROM faturamentonfeitens " & _
                       "WHERE idnfe='" & Rst.Fields("idnfe") & "'"
                Set Rst1 = RegistroBuscar(sSQL)
                If Rst1.BOF And Rst1.EOF Then
                    Else
                        Rst1.MoveFirst
                        Do Until Rst1.EOF
                            vAcumVenda = Val(ChkVal(vAcumVenda, 0, cDecMoeda)) + Val(ChkVal(Rst1.Fields("det_vProd"), 0, cDecMoeda))
                            vAcumComi = Val(ChkVal(vAcumComi, 0, cDecMoeda)) + Val(ChkVal(Rst1.Fields("comissao_vComissao"), 0, cDecMoeda))
                            Rst1.MoveNext
                        Loop
                End If
                Rst1.Close
                'vAcumVenda = Val(ChkVal(vAcumVenda, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("det_vProd"), 0, cDecMoeda))
                'vAcumComi = Val(ChkVal(vAcumComi, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("comissao_vComissao"), 0, cDecMoeda))
                Rst.MoveNext
                
            Loop
            Rst.MovePrevious
            'If MesAno <> Format(Rst.Fields("ide_demi"), "MM/YYYY") Then
                    cReg = 0
                    vReg(cReg) = Array("Vendedor", Rst.Fields("Ger_Vendedor"), "N"): cReg = cReg + 1
                    vReg(cReg) = Array("Mes", "01/" & MesAno, "D"): cReg = cReg + 1
                    vReg(cReg) = Array("vVenda", vAcumVenda, "S"): cReg = cReg + 1
                    vReg(cReg) = Array("vComissao", vAcumComi, "S"): cReg = cReg + 1
                    cReg = cReg - 1
                    RegistroIncluir "tmp_faturamentoanalise", vReg, cReg
                    MesAno = Format(Rst.Fields("ide_demi"), "MM/YYYY")
             '   End If
             Rst.Close
    End If
End Sub

Private Sub cboFuncionario_DropDown()
    Dim Rst As Recordset
    cboFuncionario.Clear
    Set Rst = RegistroBuscar("SELECT * FROM RHFuncionarioCadastro WHERE Id_Empresa=" & ID_Empresa & " ORDER BY xNome")
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboFuncionario.AddItem Left(String(4, "0"), 4 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & " - " & Rst.Fields("xNome")
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
    dtpIni.Value = Date
    dtpFin.Value = Date
End Sub

Private Sub optOpcao_Click(Index As Integer)
    TpRelatorio = Index
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case TpRelatorio
        Case 0 'Resumo DIARIO de VENDAS
            Rpt_000
        Case 1 'Resumo MENSAL de VENDAS
            Rpt_001
        Case 2 'Listagem de VENDAS no periodo
            Rpt_002
        Case 3
            Rpt_003
        Case 4
            Rpt_004
        Case 5
            Rpt_005
        Case 6 'Analise de compras por grupo e sub grupo
            Rpt_006
        Case 7 'Listagem de vendas por vendedor no periodo
            Rpt_007
        Case 8 'Listagem de vendas por cliente no periodo
            Rpt_008
        Case 9 'Entrada e Saida  por grupo e sub grupo
            Rpt_009
        Case 10 'Listagem de VENDAS por vendedor acumulado mensal
            Rpt_010
        Case 11 'Analise da rentabilidade das vendas
            Rpt_011
        Case 12 ' Entrada e Saida  por grupo
            Rpt_012
    End Select
End Sub
Private Sub Rpt_000()
'****************************************************************
'****************************************************************
'*** Relatorio do Resumo Diario de vendas
'****************************************************************
'****************************************************************
    Dim Rst         As Recordset
    Dim Rst1        As Recordset
    Dim sSQL        As String
    Dim dt          As Date
    
    Dim vVenda      As String
    Dim vCanc       As String
    Dim vDif        As String
    Dim vTotal      As String
    'Soma somente o valor do produto sem impostos
    
    
    
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_tpNF = 1" & _
           " AND ide_natOP = 'VENDA'" & _
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & _
           "' ORDER BY ide_dEmi"
           
    '25/04/2012 - Removido pois nao pega somente NFe de saida
    'sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & _
           "' ORDER BY ide_dEmi"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma nota emitida no periodo.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            MontarTabelaTemporaria_000
            dt = Rst.Fields("ide_dEmi")
            vVenda = 0
            vCanc = 0
            vDif = 0
            vTotal = 0
            Do Until Rst.EOF
                If dt <> Rst.Fields("ide_dEmi") Then
                        vDif = Val(ChkVal(vVenda, 0, cDecMoeda)) - Val(ChkVal(vCanc, 0, cDecMoeda))
                        vTotal = Val(ChkVal(vDif, 0, cDecMoeda)) + Val(ChkVal(vTotal, 0, cDecMoeda))
                        Grv_000 dt, vVenda, vCanc, vDif
                        dt = Rst.Fields("ide_dEmi")
                        vVenda = 0
                        vCanc = 0
                        vDif = 0
                        
                        vVenda = CalcVenda_000(vVenda, Rst.Fields("total_vNF"))
                        If Not IsNull(Rst.Fields("Canc_nProt")) Or Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                        If IsNull(Rst.Fields("Canc_nProt")) And Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                    Else
                        vVenda = CalcVenda_000(vVenda, Rst.Fields("total_vNF"))
                        If Not IsNull(Rst.Fields("Canc_nProt")) Or Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                        If IsNull(Rst.Fields("Canc_nProt")) And Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                End If
                Rst.MoveNext
            Loop
            vDif = Val(ChkVal(vVenda, 0, cDecMoeda)) - Val(ChkVal(vCanc, 0, cDecMoeda))
            vTotal = Val(ChkVal(vDif, 0, cDecMoeda)) + Val(ChkVal(vTotal, 0, cDecMoeda))
            Grv_000 dt, vVenda, vCanc, vDif
    End If
    Rst.Close
    sSQL = "SELECT * FROM tmp_faturamentoanalise"
    Set Rst = RegistroBuscar(sSQL)
    Set rptListaAnaliseFaturamentoDiario.DataSource = Rst.DataSource
    rptListaAnaliseFaturamentoDiario.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
    rptListaAnaliseFaturamentoDiario.Sections("Section2").Controls.Item("lblTitulo").Caption = "Analise Faturamento Diario (De: " & dtpIni.Value & " até " & dtpFin.Value & ")"
    rptListaAnaliseFaturamentoDiario.Show 1
    Rst.Close
End Sub
Private Sub Rpt_001()
'****************************************************************
'****************************************************************
'*** Relatorio do Resumo Mensal de vendas
'****************************************************************
'****************************************************************
    Dim Rst         As Recordset
    Dim Rst1        As Recordset
    Dim sSQL        As String
    Dim dt          As String
    
    Dim vVenda      As String
    Dim vCanc       As String
    Dim vDif        As String
    Dim vTotal      As String
    'Soma somente o valor do produto sem impostos
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_tpNF=1" & _
           " AND ide_natOP = 'VENDA'" & _
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & _
           "' ORDER BY ide_dEmi"
    
    
    '25/04/2012 - Retirado pois nao estava listando somente as NFe de Saida
    'sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & _
           "' ORDER BY ide_dEmi"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma nota emitida no periodo.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            MontarTabelaTemporaria_001
            dt = Format(Rst.Fields("ide_dEmi"), "MM/YYYY")
            vVenda = 0
            vCanc = 0
            vDif = 0
            vTotal = 0
            Do Until Rst.EOF
                If dt <> Format(Rst.Fields("ide_dEmi"), "MM/YYYY") Then
                        vDif = Val(ChkVal(vVenda, 0, cDecMoeda)) - Val(ChkVal(vCanc, 0, cDecMoeda))
                        vTotal = Val(ChkVal(vDif, 0, cDecMoeda)) + Val(ChkVal(vTotal, 0, cDecMoeda))
                        Grv_001 dt, vVenda, vCanc, vDif
                        dt = Format(Rst.Fields("ide_dEmi"), "MM/YYYY")
                        vVenda = 0
                        vCanc = 0
                        vDif = 0
                        
                        vVenda = CalcVenda_000(vVenda, Rst.Fields("total_vNF"))
                        If Not IsNull(Rst.Fields("Canc_nProt")) Or Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                        If IsNull(Rst.Fields("Canc_nProt")) And Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                    Else
                        vVenda = CalcVenda_000(vVenda, Rst.Fields("total_vNF"))
                        If Not IsNull(Rst.Fields("Canc_nProt")) Or Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                        If IsNull(Rst.Fields("Canc_nProt")) And Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                End If
                Rst.MoveNext
            Loop
            vDif = Val(ChkVal(vVenda, 0, cDecMoeda)) - Val(ChkVal(vCanc, 0, cDecMoeda))
            vTotal = Val(ChkVal(vDif, 0, cDecMoeda)) + Val(ChkVal(vTotal, 0, cDecMoeda))
            Grv_001 dt, vVenda, vCanc, vDif
    End If
    Rst.Close
    sSQL = "SELECT * FROM tmp_faturamentoanalise"
    Set Rst = RegistroBuscar(sSQL)
    Set rptListaAnaliseFaturamentoDiario.DataSource = Rst.DataSource
    rptListaAnaliseFaturamentoDiario.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
    rptListaAnaliseFaturamentoDiario.Sections("Section2").Controls.Item("lblTitulo").Caption = "Analise Faturamento Mensal (De: " & dtpIni.Value & " até " & dtpFin.Value & ")"
    rptListaAnaliseFaturamentoDiario.Show 1
    Rst.Close
End Sub
Private Sub Rpt_002()
'****************************************************************
'****************************************************************
'*** Relatorio de Listagem de vendas no periodo
'****************************************************************
'****************************************************************
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim vProd   As String
    Dim vTotal  As String
    Dim idFunc  As Integer
    vTotal = "0"
    vProd = "0"
    
    
     If Len(Trim(cboFuncionario.Text)) <> 0 Then
        idFunc = Left(cboFuncionario.Text, 4)
    End If
      
    
    sSQL = "SELECT ide_dEmi, ide_tpNF, ide_natOP, ide_nNF, dest_xNome, (total_vProd * 1) AS vProd, (total_vNF * 1) AS vNF " & _
           "FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_tpNF=1" & _
           " AND ide_natOP = 'VENDA'"
           
            'Adiciona informação de 1 ou todos vendedores
            sSQL = sSQL & IIf(idFunc = 0, "", " AND ger_Vendedor = " & idFunc)
     
           sSQL = sSQL & " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & _
           "' AND canc_nProt IS NULL ORDER BY ide_dEmi, ide_nNF"
    
    '25/04/2012 - Retirado pois nao listava somente as NFe de Saida
    'sSQL = "SELECT ide_dEmi, ide_tpNF, ide_nNF, dest_xNome, IF(ide_tpNF <> 1,- total_vNF, total_vNF*1) AS vNF " & _
           "FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & _
           "' AND canc_nProt IS NULL ORDER BY ide_dEmi, ide_nNF"
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma registro encontrado no periodo!", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                'If Rst.Fields("ide_tpNF") = "1" Then
                        vTotal = Val(ChkVal(vTotal, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("vNF"), 0, cDecMoeda))
                        vProd = Val(ChkVal(vProd, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("vProd"), 0, cDecMoeda))
                '    Else
                '        vTotal = Val(ChkVal(vTotal, 0, cDecMoeda)) - Val(ChkVal(Rst.Fields("vNF"), 0, cDecMoeda))
                'End If
                Rst.MoveNext
            Loop
            Rst.MoveFirst
            Set rptListaVendasPeriodo.DataSource = Rst.DataSource
            rptListaVendasPeriodo.Sections("Section2").Controls.Item("lblTitulo").Caption = "RELATORIO DE NOTAS FISCAIS DE VENDA"
            rptListaVendasPeriodo.Sections("Section1").Controls.Item("txtNome").DataField = "dest_xNome"
            rptListaVendasPeriodo.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
            rptListaVendasPeriodo.Sections("Section5").Controls.Item("lblvProd").Visible = True
            rptListaVendasPeriodo.Sections("Section5").Controls.Item("lblvProd").Caption = ConvMoeda(ChkVal(vProd, 0, cDecMoeda))
            rptListaVendasPeriodo.Show 1
    End If
    Rst.Close
End Sub
Private Sub Rpt_003()
'****************************************************************
'****************************************************************
'*** Relatorio do Resumo Diario de COMPRAS
'*** Data: 25/07/2011
'****************************************************************
'****************************************************************
    Dim Rst         As Recordset
    Dim Rst1        As Recordset
    Dim sSQL        As String
    Dim dt          As Date
    
    Dim vVenda      As String
    Dim vCanc       As String
    Dim vDif        As String
    Dim vTotal      As String
    'Soma somente o valor do produto sem impostos
    
    sSQL = "SELECT * FROM FaturamentoNFeEntrada WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & "'" & _
           IIf(PgDadosConfig.NFDevolucaoCompra = 1, "", " AND NFDevolucao = 0") & _
           " ORDER BY ide_dEmi"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma nota cadastrada no periodo.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            MontarTabelaTemporaria_000
            dt = Rst.Fields("ide_dEmi")
            vVenda = 0
            vCanc = 0
            vDif = 0
            vTotal = 0
            Do Until Rst.EOF
                If dt <> Rst.Fields("ide_dEmi") Then
                        vDif = Val(ChkVal(vVenda, 0, cDecMoeda)) - Val(ChkVal(vCanc, 0, cDecMoeda))
                        vTotal = Val(ChkVal(vDif, 0, cDecMoeda)) + Val(ChkVal(vTotal, 0, cDecMoeda))
                        Grv_000 dt, vVenda, vCanc, vDif
                        dt = Rst.Fields("ide_dEmi")
                        vVenda = 0
                        vCanc = 0
                        vDif = 0
                        
                        vVenda = CalcVenda_000(vVenda, Rst.Fields("total_vNF"))
                        'If Not IsNull(Rst.Fields("Canc_nProt")) Or Rst.Fields("ide_tpNF") <> 1 Then
                        If Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                        'If IsNull(Rst.Fields("Canc_nProt")) And Rst.Fields("ide_tpNF") <> 1 Then
                        If Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                    Else
                        vVenda = CalcVenda_000(vVenda, Rst.Fields("total_vNF"))
                        'If Not IsNull(Rst.Fields("Canc_nProt")) Or Rst.Fields("ide_tpNF") <> 1 Then
                        If Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                        'If IsNull(Rst.Fields("Canc_nProt")) And Rst.Fields("ide_tpNF") <> 1 Then
                        If Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                End If
                Rst.MoveNext
            Loop
            vDif = Val(ChkVal(vVenda, 0, cDecMoeda)) - Val(ChkVal(vCanc, 0, cDecMoeda))
            vTotal = Val(ChkVal(vDif, 0, cDecMoeda)) + Val(ChkVal(vTotal, 0, cDecMoeda))
            Grv_000 dt, vVenda, vCanc, vDif
    End If
    Rst.Close
    sSQL = "SELECT * FROM tmp_FaturamentoAnalise"
    Set Rst = RegistroBuscar(sSQL)
    Set rptListaAnaliseFaturamentoDiario.DataSource = Rst.DataSource
    rptListaAnaliseFaturamentoDiario.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
    rptListaAnaliseFaturamentoDiario.Sections("Section2").Controls.Item("lblTitulo").Caption = "Analise de Compra Diaria (De: " & dtpIni.Value & " até " & dtpFin.Value & ")"
    rptListaAnaliseFaturamentoDiario.Show 1
    Rst.Close
End Sub

Private Sub Rpt_004()
'****************************************************************
'****************************************************************
'*** Relatorio do Resumo Mensal de COMPRAS
'*** Data: 25/07/2011
'****************************************************************
'****************************************************************
    Dim Rst         As Recordset
    Dim Rst1        As Recordset
    Dim sSQL        As String
    Dim dt          As String
    
    Dim vVenda      As String
    Dim vCanc       As String
    Dim vDif        As String
    Dim vTotal      As String
    'Soma somente o valor do produto sem impostos
    
    sSQL = "SELECT * FROM FaturamentoNFeEntrada WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & "'" & _
           IIf(PgDadosConfig.NFDevolucaoCompra = 1, "", " AND NFDevolucao = 0") & _
           " ORDER BY ide_dEmi"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma nota emitida no periodo.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            MontarTabelaTemporaria_001
            dt = Format(Rst.Fields("ide_dEmi"), "MM/YYYY")
            vVenda = 0
            vCanc = 0
            vDif = 0
            vTotal = 0
            Do Until Rst.EOF
                If dt <> Format(Rst.Fields("ide_dEmi"), "MM/YYYY") Then
                        vDif = Val(ChkVal(vVenda, 0, cDecMoeda)) - Val(ChkVal(vCanc, 0, cDecMoeda))
                        vTotal = Val(ChkVal(vDif, 0, cDecMoeda)) + Val(ChkVal(vTotal, 0, cDecMoeda))
                        Grv_001 dt, vVenda, vCanc, vDif
                        dt = Format(Rst.Fields("ide_dEmi"), "MM/YYYY")
                        vVenda = 0
                        vCanc = 0
                        vDif = 0
                        
                        vVenda = CalcVenda_000(vVenda, Rst.Fields("total_vNF"))
                        'If Not IsNull(Rst.Fields("Canc_nProt")) Or Rst.Fields("ide_tpNF") <> 1 Then
                        If Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                        'If IsNull(Rst.Fields("Canc_nProt")) And Rst.Fields("ide_tpNF") <> 1 Then
                        If Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                    Else
                        vVenda = CalcVenda_000(vVenda, Rst.Fields("total_vNF"))
                        'If Not IsNull(Rst.Fields("Canc_nProt")) Or Rst.Fields("ide_tpNF") <> 1 Then
                        If Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                        'If IsNull(Rst.Fields("Canc_nProt")) And Rst.Fields("ide_tpNF") <> 1 Then
                        If Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                End If
                Rst.MoveNext
            Loop
            vDif = Val(ChkVal(vVenda, 0, cDecMoeda)) - Val(ChkVal(vCanc, 0, cDecMoeda))
            vTotal = Val(ChkVal(vDif, 0, cDecMoeda)) + Val(ChkVal(vTotal, 0, cDecMoeda))
            Grv_001 dt, vVenda, vCanc, vDif
    End If
    Rst.Close
    sSQL = "SELECT * FROM tmp_FaturamentoAnalise"
    Set Rst = RegistroBuscar(sSQL)
    Set rptListaAnaliseFaturamentoDiario.DataSource = Rst.DataSource
    rptListaAnaliseFaturamentoDiario.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
    rptListaAnaliseFaturamentoDiario.Sections("Section2").Controls.Item("lblTitulo").Caption = "Analise Compra Mensal (De: " & dtpIni.Value & " até " & dtpFin.Value & ")"
    rptListaAnaliseFaturamentoDiario.Show 1
    Rst.Close
End Sub

Private Sub Rpt_005()
'****************************************************************
'****************************************************************
'*** Relatorio de Listagem de Compras No periodo
'*** Data: 25/07/2011
'****************************************************************
'****************************************************************
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim vTotal  As String
    Dim vProd   As String
    vTotal = "0"
    vProd = "0"
    'sSQL = "SELECT * FROM FaturamentoNFeEntrada WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & _
           "' AND canc_nProt IS NULL ORDER BY ide_dEmi"
           
           
    sSQL = "SELECT ide_dEmi, ide_tpNF, ide_nNF, emit_xNome, dest_xNome, (total_vProd*1) as vProd, (total_vNF*1) as vNF " & _
           "FROM FaturamentoNFeEntrada " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & "'" & _
           IIf(PgDadosConfig.NFDevolucaoCompra = 1, "", " AND NFDevolucao = 0") & _
           " ORDER BY ide_dEmi"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma registro encontrado no periodo!", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                If Rst.Fields("ide_tpNF") = "1" Then
                    'MsgBox Rst.Fields("ide_nnf") & " - " & Rst.Fields("NFDevolucao")
                    vProd = Val(ChkVal(vProd, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("vProd"), 0, cDecMoeda))
                    vTotal = Val(ChkVal(vTotal, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("vNF"), 0, cDecMoeda))
                    Debug.Print Rst.Fields("vNF")
                End If
                Rst.MoveNext
            Loop
            Rst.MoveFirst
            Set rptListaVendasPeriodo.DataSource = Rst.DataSource
            rptListaVendasPeriodo.Sections("Section2").Controls.Item("lblTitulo").Caption = "RELATORIO DE NOTAS FISCAIS DE COMPRA"
            rptListaVendasPeriodo.Sections("Section1").Controls.Item("txtNome").DataField = "emit_xNome"
            rptListaVendasPeriodo.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
            rptListaVendasPeriodo.Sections("Section5").Controls.Item("lblvProd").Caption = ConvMoeda(ChkVal(vProd, 0, cDecMoeda))
            rptListaVendasPeriodo.Show 1
    End If
    Rst.Close
End Sub
Private Sub Rpt_006()
'****************************************************************
'****************************************************************
'*** Analise de compras por grupo e subgrupo por periodo
'*** Data: 08/11/2011
'****************************************************************
'****************************************************************
    Dim Rst         As Recordset
    Dim Rst1        As Recordset
    Dim sSQL        As String
    Dim vReg(10)    As Variant
    Dim cReg        As Integer
    Dim Qtd         As String
    Dim vUnit       As String
    Dim vTotal      As String
    Dim vProd       As String
    Dim c           As Integer 'contador
    MontarTabelaTemporaria_006
    sSQL = "SELECT * FROM FaturamentoNFeEntrada WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & "'" & _
           IIf(PgDadosConfig.NFDevolucaoCompra = 1, "", " AND NFDevolucao = 0") & _
           " ORDER BY ide_dEmi"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma registro encontrado no periodo!", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                sSQL = "SELECT * FROM FaturamentoNFeEntradaItens WHERE IdNFe= '" & Rst.Fields("IdNFe") & "'"
                Set Rst1 = RegistroBuscar(sSQL)
                If Rst1.BOF And Rst1.EOF Then
                    Else
                        Rst1.MoveFirst
                        Do Until Rst1.EOF
                            vTotal = 0
                            cReg = 0
                            'vReg(cReg) = Array("Dt", Rst.Fields("ide_dEmi"), "D"): cReg = cReg + 1
                            
                            
                            vTotal = Val(ChkVal(cNull(Rst1.Fields("Estoque_Qtd")), 0, cDecQtd)) * Val(ChkVal(cNull(Rst1.Fields("Estoque_vUnit")), 0, cDecQtd))
                            vTotal = ChkVal(vTotal, 0, cDecMoeda)
                            Qtd = Val(ChkVal(Qtd, 0, cDecQtd)) * Val(ChkVal(cNull(Rst1.Fields("Estoque_Qtd")), 0, cDecQtd))
                            vUnit = Val(ChkVal(vUnit, 0, cDecMoeda)) * Val(ChkVal(cNull(Rst1.Fields("Estoque_vUnit")), 0, cDecMoeda))
                            c = c + 1
                            vReg(cReg) = Array("Qtd", Qtd, "S"): cReg = cReg + 1
                            vReg(cReg) = Array("vUnit", vUnit, "S"): cReg = cReg + 1
                            vReg(cReg) = Array("vTotal", vTotal, "S"): cReg = cReg + 1
                            cReg = cReg - 1
                            RegistroIncluir "tmp_FaturamentoAnalise", vReg, cReg
                            Rst1.MoveNext
                        Loop
                        Rst1.Close
                End If
                Rst.MoveNext
            Loop
            'Rst.Close
            
            Set rptListaVendasPeriodo.DataSource = Rst.DataSource
            rptListaVendasPeriodo.Sections("Section2").Controls.Item("lblTitulo").Caption = "RELATORIO DE NOTAS FISCAIS DE COMPRA"
            rptListaVendasPeriodo.Sections("Section1").Controls.Item("txtNome").DataField = "emit_xNome"
            rptListaVendasPeriodo.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
            'rptListaVendasPeriodo.Sections("Section5").Controls.Item("lblvProd").Caption = ConvMoeda(ChkVal(vProd, 0, cDecMoeda))
            rptListaVendasPeriodo.Show 1
    End If
    Rst.Close
End Sub
Private Sub Rpt_007()
'****************************************************************
'****************************************************************
'*** Listagem de vendas por vendedor no periodo
'****************************************************************
'****************************************************************
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim vTotal      As String
    Dim vTotalIPI   As String
    Dim idFunc      As Integer
    
    If Len(Trim(cboFuncionario.Text)) <> 0 Then
        idFunc = Left(cboFuncionario.Text, 4)
    End If
    
    sSQL = "SELECT rhfuncionariocadastro.Id,rhfuncionariocadastro.xnome,rhfuncionariocadastro.comissao, ger_Vendedor, ide_tpNF, ide_natOP, ide_dEmi, SUM(total_vProd) as vProd,SUM(total_vIPI) as vIPI, COUNT(ide_nNF) as Contador " & _
           "FROM FaturamentoNFe,rhFuncionarioCadastro " & _
           "WHERE FaturamentoNFe.ID_Empresa = " & ID_Empresa & _
           " AND FaturamentoNFe.ide_tpNF = 1"
           
           'Adiciona informação de 1 ou todos vendedores
           sSQL = sSQL & IIf(idFunc = 0, "", " AND FaturamentoNFe.ger_Vendedor = " & idFunc)
           
           sSQL = sSQL & " AND FaturamentoNFe.ide_natOP = 'VENDA'" & _
           " AND rhFuncionarioCadastro.Id=FaturamentoNFe.ger_Vendedor" & _
           " AND FaturamentoNFe.ide_dEmi BETWEEN '" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpFin.Value, "YYYY-MM-DD") & "' " & _
           " AND canc_nProt IS NULL " & _
           " AND rhfuncionariocadastro.Comissao IS NOT NULL AND rhfuncionariocadastro.Comissao <> 0 " & _
           "GROUP BY ger_Vendedor " & _
           "ORDER BY rhfuncionariocadastro.xNome"
    
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst Is Nothing Then
        MsgBox "Nenhuma nota fiscal emitida no periodo.", vbInformation, "Aviso"
            'Rst.Close
            Exit Sub
    End If
    
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma nota fiscal emitida no periodo.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            vTotal = 0
            vTotalIPI = 0
            Rst.MoveFirst
            Do Until Rst.EOF
                vTotal = Val(ChkVal(vTotal, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("vProd"), 0, cDecMoeda))
                vTotalIPI = Val(ChkVal(vTotalIPI, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("vipi"), 0, cDecMoeda))
                Rst.MoveNext
            Loop
            Rst.MoveFirst
    End If
    
    Set rptListaAnaliseFaturamentoVendedor.DataSource = Rst.DataSource
    rptListaAnaliseFaturamentoVendedor.Sections("Section2").Controls.Item("lblTitulo").Caption = "Analise Faturamento por Vendedor de " & dtpIni.Value & " ate " & dtpFin.Value
    rptListaAnaliseFaturamentoVendedor.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
    rptListaAnaliseFaturamentoVendedor.Sections("Section5").Controls.Item("lblTotalIPI").Caption = ConvMoeda(ChkVal(vTotalIPI, 0, cDecMoeda))
    rptListaAnaliseFaturamentoVendedor.Show 1
    Rst.Close
End Sub
Private Sub Rpt_010()

'****************************************************************
'*** Listagem de VENDAS por vendedor acumulado mensal
'*** 18/06/2012
'****************************************************************
    If Trim(cboFuncionario.Text) = "" Then
        MsgBox "Selecione um funcionario!", vbInformation, App.EXEName
        Exit Sub
    End If

    MontarTabelaTemporaria_010

    Dim sSQL    As String
    Dim Rst     As Recordset
    
    sSQL = "SELECT tmp_faturamentoanalise.*, rhfuncionariocadastro.* " & _
            "FROM rhfuncionariocadastro, tmp_faturamentoanalise " & _
            "WHERE rhfuncionariocadastro.id=tmp_faturamentoanalise.vendedor " & _
            "GROUP BY vendedor, Mes " & _
            "ORDER BY tmp_faturamentoanalise.vendedor, tmp_faturamentoanalise.mes"

            
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma nota fiscal emitida no periodo.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
    End If

    Set rptListaAnaliseFaturamentoVendedorMensal.DataSource = Rst.DataSource
    rptListaAnaliseFaturamentoVendedorMensal.Sections("Section2").Controls.Item("lblVendedor").Caption = Rst.Fields("xNome")
    'rptListaAnaliseFaturamentoVendedorMensal.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
    rptListaAnaliseFaturamentoVendedorMensal.Show 1
    Rst.Close
End Sub
Private Sub Rpt_008()
'****************************************************************
'****************************************************************
'*** Listagem de vendas por cliente no periodo
'****************************************************************
'****************************************************************
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim vTotal  As String
    
    sSQL = "SELECT ide_tpNF, ide_natOP, dest_IdDest, " & _
                  "CONCAT(CAST(dest_IdDest AS CHAR), ' - ', dest_xNome) AS xNome, " & _
                  "ger_Vendedor, ide_demi, SUM( IF(ide_tpNF <> 1,- total_vNF,total_vNF)) AS vProd, " & _
                  "COUNT(ide_nNF) AS Contador , total_vIPI as vIPI " & _
           "FROM FaturamentoNFe " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND Ide_tpNf = 1" & _
           " AND ide_natOP = 'VENDA'" & _
           " AND ide_dEmi BETWEEN '" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpFin.Value, "YYYY-MM-DD") & "'" & _
           " AND canc_nProt IS NULL " & _
           "GROUP BY dest_IdDest " & _
           "ORDER BY vProd DESC"
           
    '25/04/2012 - Removido pois nao Listava somente as saidas
    ' sSQL = "SELECT ide_tpNF, dest_IdDest, CONCAT(CAST(dest_IdDest AS CHAR), ' - ', dest_xNome) AS xNome, ger_Vendedor, ide_demi, SUM( IF(ide_tpNF <> 1,- total_vNF,total_vNF)) AS vProd, COUNT(ide_nNF) AS Contador " & _
           "FROM FaturamentoNFe " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi BETWEEN '" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpFin.Value, "YYYY-MM-DD") & "'" & _
           "" & _
           " AND canc_nProt IS NULL " & _
           "GROUP BY dest_IdDest " & _
           "ORDER BY vProd DESC"
 
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma nota fiscal emitida no periodo.", vbInformation, "Aviso"
            Rst.Close
            Exit Sub
        Else
            vTotal = 0
            Rst.MoveFirst
            Do Until Rst.EOF
                vTotal = Val(ChkVal(vTotal, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("vProd"), 0, cDecMoeda))
                Rst.MoveNext
            Loop
            Rst.MoveFirst
    End If
    
    Set rptListaAnaliseFaturamentoVendedor.DataSource = Rst.DataSource
    rptListaAnaliseFaturamentoVendedor.Sections("Section2").Controls.Item("lblTitulo").Caption = "Analise Faturamento por CLIENTE de " & dtpIni.Value & " ate " & dtpFin.Value
    rptListaAnaliseFaturamentoVendedor.Sections("Section2").Controls.Item("label1").Caption = "Valor da NF"
    rptListaAnaliseFaturamentoVendedor.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
    rptListaAnaliseFaturamentoVendedor.Sections("Section1").Controls.Item("Text3").DataField = "xNome"
    rptListaAnaliseFaturamentoVendedor.Show 1
    Rst.Close
End Sub
Private Sub Rpt_009()
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim SomaE   As String
    Dim SomaS   As String
    MontarTabelaTemporaria_009
    
    sSQL = "SELECT tmp_EstoqueAnalise.grupo, tmp_EstoqueAnalise.sgrupo,tmp_EstoqueAnalise.e, tmp_EstoqueAnalise.s, EstoqueGrupos.Id, EstoqueGrupos.Descricao AS grpDesc, EstoqueSubGrupo.id, EstoqueSubGrupo.Descricao  AS sgrpDesc " & _
           "FROM tmp_EstoqueAnalise,EstoqueGrupos, EstoqueSubGrupo " & _
           "WHERE tmp_EstoqueAnalise.grupo=estoqueGrupos.id AND tmp_EstoqueAnalise.sgrupo=estoqueSubGrupo.id " & _
           "ORDER BY grpDesc, sgrpDesc"
           
    Set Rst = RegistroBuscar(sSQL)
    
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado!", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            '***************************************************
            'Soma os Totais
            SomaE = 0: SomaS = 0
            Do Until Rst.EOF
                SomaE = Val(ChkVal(SomaE, 0, cDecQtd)) + Val(ChkVal(Rst.Fields("E"), 0, cDecQtd))
                SomaS = Val(ChkVal(SomaS, 0, cDecQtd)) + Val(ChkVal(Rst.Fields("S"), 0, cDecQtd))
                Rst.MoveNext
            Loop
            Rst.MoveFirst
            '***************************************************
            Set rptListaEstoqueAnaliseEntSai.DataSource = Rst.DataSource
            rptListaEstoqueAnaliseEntSai.Sections("Section2").Controls.Item("lblTitulo").Caption = "Listagem de Faturamento por Grupo e Subgrupo - " & dtpIni.Value & " até " & dtpFin.Value
            rptListaEstoqueAnaliseEntSai.Sections("Section1").Controls.Item("Text1").DataField = "grpDesc"
            rptListaEstoqueAnaliseEntSai.Sections("Section1").Controls.Item("Text3").DataField = "sgrpDesc"
            
            rptListaEstoqueAnaliseEntSai.Sections("Section5").Controls.Item("lblTotE").Caption = ChkVal(SomaE, 0, cDecQtd)
            rptListaEstoqueAnaliseEntSai.Sections("Section5").Controls.Item("lblTotS").Caption = ChkVal(SomaS, 0, cDecQtd)
            rptListaEstoqueAnaliseEntSai.Show 1
    End If
End Sub
Private Sub Rpt_012()
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim SomaE   As String
    Dim SomaS   As String
    MontarTabelaTemporaria_009
    
    'sSQL = "SELECT tmp_EstoqueAnalise.grupo, tmp_EstoqueAnalise.sgrupo, " & _
           "tmp_EstoqueAnalise.e, tmp_EstoqueAnalise.s, EstoqueGrupos.Id, " & _
           "EstoqueGrupos.Descricao AS grpDesc, EstoqueSubGrupo.id, " & _
           "EstoqueSubGrupo.Descricao  AS sgrpDesc " & _
           "FROM tmp_EstoqueAnalise,EstoqueGrupos, EstoqueSubGrupo " & _
           "WHERE tmp_EstoqueAnalise.grupo=estoqueGrupos.id AND tmp_EstoqueAnalise.sgrupo=estoqueSubGrupo.id " & _
           "GROUP BY tmp_EstoqueAnalise.grupo " & _
           "ORDER BY grpDesc"
     sSQL = "select estoquegrupos.Descricao AS grpdesc,'' AS sgrpdesc, SUM(e) AS e, SUM(s) as s FROM tmp_estoqueanalise,estoquegrupos WHERE estoquegrupos.id=tmp_estoqueanalise.grupo GROUP by grupo order by grpdesc"
    Set Rst = RegistroBuscar(sSQL)
    
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado!", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            '***************************************************
            'Soma os Totais
            SomaE = 0: SomaS = 0
            Do Until Rst.EOF
                SomaE = Val(ChkVal(SomaE, 0, cDecQtd)) + Val(ChkVal(Rst.Fields("E"), 0, cDecQtd))
                SomaS = Val(ChkVal(SomaS, 0, cDecQtd)) + Val(ChkVal(Rst.Fields("S"), 0, cDecQtd))
                Rst.MoveNext
            Loop
            Rst.MoveFirst
            '***************************************************
            Set rptListaEstoqueAnaliseEntSai.DataSource = Rst.DataSource
            rptListaEstoqueAnaliseEntSai.Sections("Section2").Controls.Item("lblTitulo").Caption = "Listagem de Faturamento por Grupo - " & dtpIni.Value & " até " & dtpFin.Value
            rptListaEstoqueAnaliseEntSai.Sections("Section1").Controls.Item("Text1").DataField = "grpDesc"
            rptListaEstoqueAnaliseEntSai.Sections("Section1").Controls.Item("Text3").DataField = "sgrpDesc"
            
            rptListaEstoqueAnaliseEntSai.Sections("Section5").Controls.Item("lblTotE").Caption = ChkVal(SomaE, 0, cDecQtd)
            rptListaEstoqueAnaliseEntSai.Sections("Section5").Controls.Item("lblTotS").Caption = ChkVal(SomaS, 0, cDecQtd)
            rptListaEstoqueAnaliseEntSai.Show 1
    End If
End Sub
Private Sub Rpt_011()
    
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim sTabela As String
    
    sTabela = "tmp_faturamentoanalise"
    
    MontarTabelaTemporaria_011 (sTabela)
    
    sSQL = "SELECT * FROM " & sTabela
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum registro encontrado", vbInformation, App.EXEName
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            
            Set rptListaAnaliseFaturamentoMarkup.DataSource = Rst.DataSource
            rptListaAnaliseFaturamentoMarkup.Show 1
            Rst.Close
    End If
    
    
End Sub


Private Sub Grv_000(dt As Date, vVendas As String, vCanc As String, vDif As String)
    Dim vReg(1000)  As Variant
    Dim cReg        As Integer
    
    cReg = 0
    
    
    vVendas = ConvMoeda(ChkVal(vVendas, 0, cDecMoeda))
    vCanc = ConvMoeda(ChkVal(vCanc, 0, cDecMoeda))
    vDif = ConvMoeda(ChkVal(vDif, 0, cDecMoeda))
    
    
    vReg(cReg) = Array("Dt", dt, "D"): cReg = cReg + 1
    vReg(cReg) = Array("DtSemana", UCase(Left(Format(dt, "Long Date"), InStr(Format(dt, "Long Date"), ",") - 1)), "S"): cReg = cReg + 1
    vReg(cReg) = Array("vVenda", vVendas, "S"): cReg = cReg + 1
    vReg(cReg) = Array("vCanc", vCanc, "S"): cReg = cReg + 1
    vReg(cReg) = Array("vDif", vDif, "S") ': cReg = cReg + 1
    RegistroIncluir "tmp_faturamentoAnalise", vReg, cReg
End Sub
Private Sub Grv_001(dt As String, vVendas As String, vCanc As String, vDif As String)
    Dim vReg(1000)  As Variant
    Dim cReg        As Integer
    
    cReg = 0
    
    
    vVendas = ConvMoeda(ChkVal(vVendas, 0, cDecMoeda))
    vCanc = ConvMoeda(ChkVal(vCanc, 0, cDecMoeda))
    vDif = ConvMoeda(ChkVal(vDif, 0, cDecMoeda))
    
    
    vReg(cReg) = Array("Dt", dt, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("DtSemana", UCase(Left(Format(Dt, "Long Date"), InStr(Format(Dt, "Long Date"), ",") - 1)), "S"): cReg = cReg + 1
    vReg(cReg) = Array("vVenda", vVendas, "S"): cReg = cReg + 1
    vReg(cReg) = Array("vCanc", vCanc, "S"): cReg = cReg + 1
    vReg(cReg) = Array("vDif", vDif, "S") ': cReg = cReg + 1
    RegistroIncluir "tmp_faturamentoAnalise", vReg, cReg
End Sub

Private Sub MontarTabelaTemporaria_000()

    BD.Execute "DROP TABLE IF EXISTS tmp_faturamentoanalise"
    BD.Execute "CREATE TABLE IF NOT EXISTS tmp_faturamentoanalise" & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "Dt DATE default Null," & _
               "dtSemana VARCHAR(100) default Null," & _
               "vVenda VARCHAR(100) default Null," & _
               "vCanc VARCHAR(100) default Null," & _
               "vDif VARCHAR(100) default Null," & _
               "PRIMARY KEY (Id))"
End Sub
Private Sub MontarTabelaTemporaria_001()

    BD.Execute "DROP TABLE IF EXISTS tmp_faturamentoanalise"
    BD.Execute "CREATE TABLE IF NOT EXISTS tmp_faturamentoanalise" & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "Dt VARCHAR(10) default Null," & _
               "dtSemana VARCHAR(100) default Null," & _
               "vVenda VARCHAR(100) default Null," & _
               "vCanc VARCHAR(100) default Null," & _
               "vDif VARCHAR(100) default Null," & _
               "PRIMARY KEY (Id))"
End Sub
Private Sub MontarTabelaTemporaria_006()

    BD.Execute "DROP TABLE IF EXISTS tmp_FaturamentoAnalise"
    BD.Execute "CREATE TABLE IF NOT EXISTS tmp_FaturamentoAnalise" & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "Dt VARCHAR(10) default Null," & _
               "idGrupo VARCHAR(100) default Null," & _
               "idSubGrupo VARCHAR(100) default Null," & _
               "Qtd VARCHAR(100) default Null," & _
               "vUnit VARCHAR(100) default Null," & _
               "vTotal VARCHAR(100) default Null," & _
               "PRIMARY KEY (Id))"
End Sub

Private Sub MontarTabelaTemporaria_010()
    '18/06/2012 - Monta as tabelas TMP se separa por vendedor
    
    BD.Execute "DROP TABLE IF EXISTS tmp_faturamentoanalise"
    BD.Execute "CREATE TABLE IF NOT EXISTS tmp_faturamentoanalise" & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "Vendedor VARCHAR(100) default Null," & _
               "Mes DATE default Null," & _
               "vComissao VARCHAR(50) default Null," & _
               "vVenda VARCHAR(50) default Null," & _
               "PRIMARY KEY (Id))"
               
               
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim vend        As Integer
    vend = Left(Trim(cboFuncionario.Text), 4)
    
    sSQL = "SELECT * " & _
           "FROM faturamentonfe " & _
           "WHERE ID_Empresa = " & ID_Empresa & " AND ide_tpNF = 1 AND ide_natOP = 'VENDA' AND canc_nProt IS NULL " & _
           "AND ide_dEmi BETWEEN '" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpFin.Value, "YYYY-MM-DD") & "' " & _
           "AND ger_vendedor=" & vend & " " & _
           "ORDER BY ger_vendedor"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            
        Else
            Rst.MoveFirst
            vend = "0"
            Do Until Rst.EOF
                If vend <> Rst.Fields("ger_Vendedor") Then
                    MontarTabelaTemporaria_010_a Rst.Fields("ger_Vendedor")
                    vend = Rst.Fields("ger_Vendedor")
                End If
                Rst.MoveNext
            Loop
    End If
End Sub
Private Sub MontarTabelaTemporaria_009()
    On Error Resume Next
    '########################################################################################################
    '### Analise de Entrada/Saida por grupo
    '########################################################################################################
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim Rst1        As Recordset
    Dim tabela      As String
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    '########################################################################################################
    '### Montando a tabela temporaria
    '########################################################################################################
    tabela = LCase("tmp_EstoqueAnalise")
    
    BD.Execute "DROP TABLE IF EXISTS " & LCase(tabela)
    
    BD.Execute "CREATE TABLE IF NOT EXISTS " & tabela & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "Grupo NUMERIC(11) default Null," & _
               "sGrupo NUMERIC(11) default Null," & _
               "E VARCHAR(100) default Null," & _
               "S VARCHAR(100) default Null," & _
               "PRIMARY KEY (Id))"
    '########################################################################################################
    '### Capturando as Entradas
    '########################################################################################################
    sSQL = "SELECT FatEnt.*, FatEntItens.*, Est.*, SUM(FatEntItens.Estoque_Qtd) as saldoEst " & _
           "FROM FaturamentoNFeEntrada AS FatEnt, FaturamentoNFeEntradaItens AS FatEntItens, EstoqueProduto AS Est " & _
           "WHERE FatEnt.ide_dEmi BETWEEN '" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpFin.Value, "YYYY-MM-DD") & "' " & _
           "AND FatEnt.IdNFe = FatEntItens.IdNFe " & _
           "AND Est.Id = FatEntItens.det_IdProduto " & _
           "GROUP BY Est.Grupo, Est.SubGrupo"
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                cReg = 0
                vReg(cReg) = Array("Grupo", Rst.Fields("Grupo"), "N"): cReg = cReg + 1
                vReg(cReg) = Array("sGrupo", Rst.Fields("SubGrupo"), "N"): cReg = cReg + 1
                vReg(cReg) = Array("E", ChkVal(Rst.Fields("saldoEst"), 0, cDecQtd), "S"): cReg = cReg + 1
                vReg(cReg) = Array("S", ChkVal("0", 0, cDecQtd), "S"): cReg = cReg + 1
                cReg = cReg - 1
                RegistroIncluir tabela, vReg, cReg
                'MsgBox Rst.Fields("ID") & " - " & Rst.Fields("Descricao") & vbCrLf & _
                   Rst.Fields("emit_xNome") & vbCrLf & _
                   Rst.Fields("ide_nnf") & vbCrLf & _
                   Rst.Fields("Grupo") & vbCrLf & _
                   "Saldo " & Rst.Fields("saldoEst")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    '########################################################################################################
    '### Capturando as Saidas
    '########################################################################################################
     sSQL = "SELECT FatEnt.*, FatEntItens.*, Est.*, SUM(FatEntItens.Estoque_Qtd) as saldoEst " & _
           "FROM FaturamentoNFe AS FatEnt, FaturamentoNFeItens AS FatEntItens, EstoqueProduto AS Est " & _
           "WHERE FatEnt.ide_dEmi BETWEEN '" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpFin.Value, "YYYY-MM-DD") & "' " & _
           "AND FatEnt.IdNFe = FatEntItens.IdNFe " & _
           "AND Est.Id = FatEntItens.det_IdProduto " & _
           "AND FatEnt.canc_nProt IS NULL " & _
           "GROUP BY Est.Grupo, Est.SubGrupo"
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'Exit Sub
        Else
            Rst.MoveFirst
            'MsgBox Rst.RecordCount
            Do Until Rst.EOF
                status (Rst.RecordCount)
                cReg = 0
                vReg(cReg) = Array("Grupo", Rst.Fields("Grupo"), "N"): cReg = cReg + 1
                vReg(cReg) = Array("sGrupo", Rst.Fields("SubGrupo"), "N"): cReg = cReg + 1
                vReg(cReg) = Array("S", ChkVal(Rst.Fields("saldoEst"), 0, cDecQtd), "S"): cReg = cReg + 1
                cReg = cReg - 1
                
                sSQL = "SELECT * FROM " & tabela & " WHERE Grupo = " & Rst.Fields("Grupo") & " AND sGrupo = " & Rst.Fields("SubGrupo") & ""
                Set Rst1 = RegistroBuscar(sSQL)
                If Rst1.BOF And Rst1.EOF Then
                        cReg = cReg + 1: vReg(cReg) = Array("E", ChkVal("0", 0, cDecQtd), "S")
                        RegistroIncluir tabela, vReg, cReg
                    Else
                        RegistroAlterar tabela, vReg, cReg, "Grupo = " & Rst.Fields("Grupo") & " AND sGrupo = " & Rst.Fields("SubGrupo") & ""
                End If
                Rst1.Close
                Rst.MoveNext
            Loop
    End If
    Rst.Close

End Sub
Private Sub MontarTabelaTemporaria_011(sTabela As String)
    ' 30/08/2012 - Leonardo Aquino
    'Avalia a o markup da venda
    
    
    
    BD.Execute "DROP TABLE IF EXISTS " & sTabela
    BD.Execute "CREATE TABLE IF NOT EXISTS " & sTabela & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "Dt VARCHAR(10) default Null," & _
               "Nome VARCHAR(250) default Null," & _
               "Doc VARCHAR(100) default Null," & _
               "Produto VARCHAR(300) default Null," & _
               "unid VARCHAR(10) default Null," & _
               "qtd VARCHAR(100) default Null," & _
               "vUnCusto VARCHAR(100) default Null," & _
               "vUnVenda VARCHAR(100) default Null," & _
               "vTotCusto VARCHAR(100) default Null," & _
               "vTotVenda VARCHAR(100) default Null," & _
               "vMarkup VARCHAR(100) default Null," & _
               "PRIMARY KEY (Id))"
               
               
    '***************************************************************
    Dim sSQL        As String
    Dim Rst         As Recordset
    
    Dim qv          As String 'Quantidade vendida
    Dim un          As String 'Unidade de Armazenamento
                
    Dim vcu         As String 'Valor de custo unitario
    Dim vct         As String 'Valor de custo total
                
    Dim vvu         As String 'Valor de venda unitario
    Dim vvt         As String 'Valor de venda total
                
    Dim vmkp        As String 'Valor do markup
    
    Dim cReg        As Integer
    Dim vReg(10)    As Variant
    
    
    sSQL = "SELECT * " & _
           "FROM estoquekardex " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND deposito = " & ID_Deposito & _
           " AND movimento = 'SUBTRAIR (-)'" & _
           " AND datamov BETWEEN '" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpFin.Value, "YYYY-MM-DD") & "'" & _
           " AND nfe IS NOT NULL "
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                'Calculo do custo
                qv = ChkVal(Rst.Fields("Quantidade"), 0, cDecQtd)
                un = pgDadosEstoqueProduto(Rst.Fields("idProduto")).Unidade
                
                'Calculo do custo
                vcu = ChkVal(pgDadosEstoqueProduto(Rst.Fields("idProduto")).VlCusto, 0, cDecMoeda)
                
                vct = Val(qv) * Val(vcu)
                
                'Calculo da venda
                vvu = ChkVal(Rst.Fields("ValorUnitario"), 0, cDecMoeda)
                vvt = Val(qv) * Val(vvu)
                
                'vct = IIf(CInt(vct) = 0, vvt, vct)
                vmkp = Val(ChkVal(vvt, 0, cDecMoeda)) / Val(ChkVal(IIf(CInt(vct) = 0, vvt, vct), 0, cDecMoeda))
                vmkp = Format(vmkp * 100, "0.000") & "%"
                
                cReg = 0
                vReg(cReg) = Array("Dt", Format(Rst.Fields("datamov"), "dd/mm/yyyy"), "S"): cReg = cReg + 1
                vReg(cReg) = Array("Doc", Rst.Fields("documento"), "S"): cReg = cReg + 1
                vReg(cReg) = Array("Nome", Rst.Fields("nome"), "S"): cReg = cReg + 1
                
                vReg(cReg) = Array("Produto", ZE(Rst.Fields("idProduto"), 6) & "-" & pgDadosEstoqueProduto(Rst.Fields("idProduto")).Descricao, "S"): cReg = cReg + 1
                vReg(cReg) = Array("Unid", un, "S"): cReg = cReg + 1
                vReg(cReg) = Array("qtd", qv, "S"): cReg = cReg + 1
                
                vReg(cReg) = Array("vUnCusto", ConvMoeda(vcu), "S"): cReg = cReg + 1
                vReg(cReg) = Array("vTotCusto", ConvMoeda(vct), "S"): cReg = cReg + 1
                
                vReg(cReg) = Array("vUnVenda", ConvMoeda(vvu), "S"): cReg = cReg + 1
                vReg(cReg) = Array("vTotVenda", ConvMoeda(vvt), "S"): cReg = cReg + 1
                
                vReg(cReg) = Array("vMarkup", vmkp, "S"): cReg = cReg + 1
                
                cReg = cReg - 1
                
                RegistroIncluir sTabela, vReg, cReg
                Rst.MoveNext
            Loop
            
    End If
    Rst.Close

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
