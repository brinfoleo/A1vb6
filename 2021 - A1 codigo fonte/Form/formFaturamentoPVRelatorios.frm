VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFaturamentoPVRelatorios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pré-Venda - Relatórios"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6435
   Begin VB.Frame Frame1 
      Caption         =   "Critério"
      Height          =   795
      Left            =   60
      TabIndex        =   8
      Top             =   1740
      Width           =   6255
      Begin VB.OptionButton optCriterio 
         Caption         =   "&SEM nota fiscal vinculada"
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   11
         Top             =   300
         Width           =   2235
      End
      Begin VB.OptionButton optCriterio 
         Caption         =   "&COM nota fiscal vinculada"
         Height          =   375
         Index           =   1
         Left            =   1260
         TabIndex        =   10
         Top             =   300
         Width           =   2235
      End
      Begin VB.OptionButton optCriterio 
         Caption         =   "&Todos"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   975
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
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      Begin VB.OptionButton optListagem 
         Caption         =   "Data Emissão:"
         Height          =   195
         Index           =   1
         Left            =   3300
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optListagem 
         Caption         =   "Num. PV:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   420
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.TextBox txtNFFim 
         Height          =   285
         Left            =   1500
         MaxLength       =   9
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   540
         Width           =   1335
      End
      Begin VB.TextBox txtNFIni 
         Height          =   285
         Left            =   1500
         MaxLength       =   9
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpNFIni 
         Height          =   285
         Left            =   4740
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   113180673
         CurrentDate     =   40584
      End
      Begin MSComCtl2.DTPicker dtpNFFim 
         Height          =   285
         Left            =   4740
         TabIndex        =   6
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   117112833
         CurrentDate     =   40584
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
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
               Picture         =   "formFaturamentoPVRelatorios.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":54D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":58CB
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":707D
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoPVRelatorios.frx":7617
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFaturamentoPVRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private criterio As Integer

Private Sub LstPV()
    Dim Rst     As Recordset
    Dim sSQL    As String
   
    Select Case criterio
        Case 0 'Todos
            'sSQL = "SELECT faturamentonfe.ide_nnf, FaturamentoPV.*, RHFuncionarioCadastro.*" & _
                   " FROM FaturamentoPV, RHFuncionarioCadastro" & _
                   " WHERE FaturamentoPV.ID_Empresa = " & ID_Empresa & _
                   " AND FaturamentoPV.vendedor = RHFuncionarioCadastro.id"


            '22.02.2018 - Testado e funcionando
            sSQL = "SELECT" & _
                " faturamentopv.Emissao, faturamentonfe.ide_nNF, faturamentopv.Id," & _
                " faturamentopv.IdCliente, faturamentopv.Cliente," & _
                " FaturamentoPV.VlTotalPV , rhfuncionariocadastro.xNome" & _
                " from" & _
                " faturamentonfe RIGHT JOIN" & _
                " faturamentopv ON faturamentopv.Id = faturamentonfe.ger_idPV" & _
                " Inner Join" & _
                " rhfuncionariocadastro ON faturamentopv.Vendedor = rhfuncionariocadastro.Id" & _
                " Where" & _
                " faturamentopv.ID_Empresa = " & ID_Empresa

        Case 1 '1 - Com NF
                
            'sSQL = "SELECT *" & _
                " FROM FaturamentoPV AS fpv, RHFuncionarioCadastro AS rh" & _
                " WHERE fpv.id IN (SELECT ger_idpv FROM faturamentonfe WHERE ger_idpv is not null)" & _
                " AND fpv.ID_Empresa = " & ID_Empresa & " AND fpv.vendedor = rh.id"
        '21.02.18 - Ajustar para que possa pegar o num da nfe
            sSQL = "SELECT * " & _
            " from" & _
            " faturamentonfe INNER JOIN" & _
            " rhfuncionariocadastro ON faturamentonfe.ger_Vendedor =" & _
            " rhfuncionariocadastro.Id INNER JOIN" & _
            " faturamentopv ON faturamentonfe.ger_idPV = faturamentopv.Id" & _
            " Where" & _
            " faturamentopv.Id_Empresa = " & ID_Empresa '& _
            " AND faturamentopv.Emissao BETWEEN '2018-02-21' AND '2018-02-21';"
            
        
        Case 2 '2 - Sem NF
            'sSQL = "SELECT * FROM faturamentopv WHERE id not IN (SELECT ger_idpv FROM faturamentonfe)"
             sSQL = "SELECT '00000' AS ide_nnf, FaturamentoPV.*, RHFuncionarioCadastro.*" & _
                " FROM FaturamentoPV, RHFuncionarioCadastro" & _
                " WHERE FaturamentoPV.id not IN (SELECT ger_idpv FROM faturamentonfe WHERE ger_idpv is not null)" & _
                " AND FaturamentoPV.ID_Empresa = " & ID_Empresa & " AND FaturamentoPV.vendedor = RHFuncionarioCadastro.id"
        
    End Select
                   
           
           
    'Seleciona o tipo de listagem, numero ou data
    If optListagem(0).Value = True Then
            sSQL = sSQL & " AND FaturamentoPV.id >=" & IIf(Trim(txtNFIni.Text) = "", "0", txtNFIni.Text) & " AND FaturamentoPV.id <= " & IIf(Trim(txtNFFim.Text) = "", "0", txtNFFim.Text)
        ElseIf optListagem(1).Value = True Then
            sSQL = sSQL & " AND FaturamentoPV.emissao >= '" & Format(dtpNFIni.Value, "yyyy-mm-dd") & "' AND FaturamentoPV.emissao <= '" & Format(dtpNFFim.Value, "yyyy-mm-dd") & "'"
        Else
            MsgBox "Selecione uma opção de listagem!", vbInformation, App.EXEName
            Exit Sub
    End If
    
   sSQL = sSQL & " ORDER BY FaturamentoPV.emissao, FaturamentoPV.vendedor, FaturamentoPV.id"
   
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum documento encontrado!", vbInformation, App.EXEName
        Else
            Rst.MoveFirst
            Set rptListaPVPeriodo.DataSource = Rst.DataSource
            rptListaPVPeriodo.Sections("Section2").Controls.Item("lblTitulo").Caption = "RELATORIO DE PRE-VENDAS NO PERIDO (" & Replace(optCriterio(criterio).Caption, "&", "") & ")"
            'rptListaVendasPeriodo.Sections("Section1").Controls.Item("txtNome").DataField = "dest_xNome"
            rptListaPVPeriodo.Sections("Section5").Controls.Item("lblTotal").Visible = False
            rptListaPVPeriodo.Sections("Section5").Controls.Item("lblTotal").Caption = "0.00" 'ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
            rptListaPVPeriodo.Sections("Section5").Controls.Item("lblvProd").Visible = False
            rptListaPVPeriodo.Sections("Section5").Controls.Item("lblvProd").Caption = "0.00" 'ConvMoeda(ChkVal(vProd, 0, cDecMoeda))
            rptListaPVPeriodo.Show 1
            
    End If
    Rst.Close
End Sub

Private Sub Form_Load()
    criterio = 0
    limpaform
End Sub
Private Sub limpaform()
    txtNFIni.Text = ""
    txtNFFim.Text = ""
    dtpNFIni.Value = Date
    dtpNFFim.Value = Date
    
    optListagem_Click (0)
    
End Sub

Private Sub optCriterio_Click(Index As Integer)
    '0 - Todos
    '1 - Com NF
    '2 - Sem NF

    criterio = Index
End Sub

Private Sub optListagem_Click(Index As Integer)
    If Index = 0 Then
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

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualizar"
            LstPV
    End Select
End Sub


Private Sub txtNFFim_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub txtNFIni_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
End Sub
