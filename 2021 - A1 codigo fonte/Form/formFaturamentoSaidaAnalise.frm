VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formFaturamentoSaidaAnalise 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Analise Saida"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4380
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   60
      TabIndex        =   5
      Top             =   1560
      Width           =   4215
      Begin VB.OptionButton optOpcao 
         Caption         =   "Listagem de Vendas no periodo"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   720
         Width           =   2595
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Resumo MENSAL de vendas"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   480
         Width           =   2595
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Resumo DIARIO de vendas"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   2355
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Perido"
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   4215
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   315
         Left            =   420
         TabIndex        =   3
         Top             =   540
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55443457
         CurrentDate     =   40665
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   315
         Left            =   2460
         TabIndex        =   4
         Top             =   540
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55443457
         CurrentDate     =   40665
      End
      Begin VB.Label Label2 
         Caption         =   "Final:"
         Height          =   195
         Left            =   2460
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   420
         TabIndex        =   1
         Top             =   300
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4380
      _ExtentX        =   7726
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
         Left            =   3600
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
               Picture         =   "formFaturamentoSaidaAnalise.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoSaidaAnalise.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoSaidaAnalise.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoSaidaAnalise.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoSaidaAnalise.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoSaidaAnalise.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoSaidaAnalise.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoSaidaAnalise.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoSaidaAnalise.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoSaidaAnalise.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoSaidaAnalise.frx":54D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoSaidaAnalise.frx":58CB
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFaturamentoSaidaAnalise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TpRelatorio As Integer
Private Function CalcVenda_001(vVenda As String, vNF As String) As String
    CalcVenda_001 = Val(ChkVal(vVenda, 0, cDecMoeda)) + Val(ChkVal(vNF, 0, cDecMoeda))
End Function

Private Sub Form_Load()
    dtpIni.Value = Date
    dtpFin.Value = Date
End Sub

Private Sub optOpcao_Click(Index As Integer)
    TpRelatorio = Index
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case TpRelatorio
        Case 0
            Rpt_001
        Case 1
            Rpt_002
        Case 2
            Rpt_003
    End Select
End Sub
Private Sub Rpt_001()
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
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & _
           "' ORDER BY ide_dEmi"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma nota emitida no periodo.", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            MontarTabelaTemporaria_001
            dt = Rst.Fields("ide_dEmi")
            vVenda = 0
            vCanc = 0
            vDif = 0
            vTotal = 0
            Do Until Rst.EOF
                If dt <> Rst.Fields("ide_dEmi") Then
                        vDif = Val(ChkVal(vVenda, 0, cDecMoeda)) - Val(ChkVal(vCanc, 0, cDecMoeda))
                        vTotal = Val(ChkVal(vDif, 0, cDecMoeda)) + Val(ChkVal(vTotal, 0, cDecMoeda))
                        Grv_001 dt, vVenda, vCanc, vDif
                        dt = Rst.Fields("ide_dEmi")
                        vVenda = 0
                        vCanc = 0
                        vDif = 0
                        
                        vVenda = CalcVenda_001(vVenda, Rst.Fields("total_vNF"))
                        If Not IsNull(Rst.Fields("Canc_nProt")) Or Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                        If IsNull(Rst.Fields("Canc_nProt")) And Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                    Else
                        vVenda = CalcVenda_001(vVenda, Rst.Fields("total_vNF"))
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
    sSQL = "SELECT * FROM tmp_FaturamentoAnalise"
    Set Rst = RegistroBuscar(sSQL)
    Set rptListaanaliseFaturamentoDiario.DataSource = Rst.DataSource
    rptListaanaliseFaturamentoDiario.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
    rptListaanaliseFaturamentoDiario.Sections("Section2").Controls.Item("lblTitulo").Caption = "Analise Faturamento Diario (De: " & dtpIni.Value & " até " & dtpFin.Value & ")"
    rptListaanaliseFaturamentoDiario.Show 1
    Rst.Close
End Sub
Private Sub Rpt_002()
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
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & _
           "' ORDER BY ide_dEmi"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma nota emitida no periodo.", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            MontarTabelaTemporaria_002
            dt = Format(Rst.Fields("ide_dEmi"), "MM/YYYY")
            vVenda = 0
            vCanc = 0
            vDif = 0
            vTotal = 0
            Do Until Rst.EOF
                If dt <> Format(Rst.Fields("ide_dEmi"), "MM/YYYY") Then
                        vDif = Val(ChkVal(vVenda, 0, cDecMoeda)) - Val(ChkVal(vCanc, 0, cDecMoeda))
                        vTotal = Val(ChkVal(vDif, 0, cDecMoeda)) + Val(ChkVal(vTotal, 0, cDecMoeda))
                        Grv_002 dt, vVenda, vCanc, vDif
                        dt = Format(Rst.Fields("ide_dEmi"), "MM/YYYY")
                        vVenda = 0
                        vCanc = 0
                        vDif = 0
                        
                        vVenda = CalcVenda_001(vVenda, Rst.Fields("total_vNF"))
                        If Not IsNull(Rst.Fields("Canc_nProt")) Or Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                        If IsNull(Rst.Fields("Canc_nProt")) And Rst.Fields("ide_tpNF") <> 1 Then
                            vCanc = Val(ChkVal(vCanc, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                        End If
                    Else
                        vVenda = CalcVenda_001(vVenda, Rst.Fields("total_vNF"))
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
            Grv_002 dt, vVenda, vCanc, vDif
    End If
    Rst.Close
    sSQL = "SELECT * FROM tmp_FaturamentoAnalise"
    Set Rst = RegistroBuscar(sSQL)
    Set rptListaanaliseFaturamentoDiario.DataSource = Rst.DataSource
    rptListaanaliseFaturamentoDiario.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
    rptListaanaliseFaturamentoDiario.Sections("Section2").Controls.Item("lblTitulo").Caption = "Analise Faturamento Diario (De: " & dtpIni.Value & " até " & dtpFin.Value & ")"
    rptListaanaliseFaturamentoDiario.Show 1
    Rst.Close
End Sub
Private Sub Rpt_003()
'****************************************************************
'****************************************************************
'*** Relatorio de Analise de vendas No periodo
'****************************************************************
'****************************************************************
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim vTotal As String
    vTotal = "0"
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpIni.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpFin.Value, "YYYY-MM-DD") & _
           "' AND canc_nProt IS NULL ORDER BY ide_dEmi"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma registro encontrado no periodo!", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                If Rst.Fields("ide_tpNF") = "1" Then
                    vTotal = Val(ChkVal(vTotal, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("total_vNF"), 0, cDecMoeda))
                End If
                Rst.MoveNext
            Loop
            Rst.MoveFirst
            Set rptListaVendasPeriodo.DataSource = Rst.DataSource
            rptListaVendasPeriodo.Sections("Section5").Controls.Item("lblTotal").Caption = ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
            rptListaVendasPeriodo.Show 1
    End If
    Rst.Close
End Sub
Private Sub Grv_001(dt As Date, vVendas As String, vCanc As String, vDif As String)
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
Private Sub Grv_002(dt As String, vVendas As String, vCanc As String, vDif As String)
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

Private Sub MontarTabelaTemporaria_001()

    BD.Execute "DROP TABLE IF EXISTS tmp_FaturamentoAnalise"
    BD.Execute "CREATE TABLE IF NOT EXISTS tmp_FaturamentoAnalise" & _
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
Private Sub MontarTabelaTemporaria_002()

    BD.Execute "DROP TABLE IF EXISTS tmp_FaturamentoAnalise"
    BD.Execute "CREATE TABLE IF NOT EXISTS tmp_FaturamentoAnalise" & _
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


