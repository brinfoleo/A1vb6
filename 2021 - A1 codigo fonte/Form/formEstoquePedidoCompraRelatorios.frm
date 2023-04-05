VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formEstoquePedidoCompraRelatorios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedido de Compra - Relatórios"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6435
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
         Caption         =   "Num. OC:"
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
         Format          =   119603201
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
         Format          =   119603201
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
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":54D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":58CB
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":707D
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoquePedidoCompraRelatorios.frx":7617
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formEstoquePedidoCompraRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub LstOC()

    Dim Rst     As Recordset
    Dim sSQL    As String
   
    
    sSQL = "select " & _
        "estoquepedidocompra.emissao, estoquepedidocompra.vendedor, rhfuncionariocadastro.xNome, " & _
        "estoquepedidocompra.id, estoquepedidocompra.idcliente, estoquepedidocompra.cliente, estoquepedidocompra.vltotalpv " & _
        "from estoquepedidocompra Inner Join rhfuncionariocadastro ON estoquepedidocompra.Vendedor = rhfuncionariocadastro.Id " & _
        " where" & _
                " estoquepedidocompra.ID_Empresa = " & ID_Empresa
    
                   
           
           
    'Seleciona o tipo de listagem, numero ou data
    If optListagem(0).Value = True Then
            sSQL = sSQL & " AND estoquepedidocompra.id >=" & IIf(Trim(txtNFIni.Text) = "", "0", txtNFIni.Text) & " AND estoquepedidocompra.id <= " & IIf(Trim(txtNFFim.Text) = "", "0", txtNFFim.Text)
        ElseIf optListagem(1).Value = True Then
            sSQL = sSQL & " AND estoquepedidocompra.emissao >= '" & Format(dtpNFIni.Value, "yyyy-mm-dd") & "' AND estoquepedidocompra.emissao <= '" & Format(dtpNFFim.Value, "yyyy-mm-dd") & "'"
        Else
            MsgBox "Selecione uma opção de listagem!", vbInformation, App.EXEName
            Exit Sub
    End If
    
   sSQL = sSQL & " ORDER BY estoquepedidocompra.emissao, estoquepedidocompra.vendedor, estoquepedidocompra.id"
   
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum documento encontrado!", vbInformation, App.EXEName
        Else
            Rst.MoveFirst
            Set rptListaOCPeriodo.DataSource = Rst.DataSource
            rptListaOCPeriodo.Sections("Section2").Controls.Item("lblTitulo").Caption = "RELATORIO DE ORDEM DE COMPRAS NO PERIDO "
            'rptListaVendasPeriodo.Sections("Section1").Controls.Item("txtNome").DataField = "dest_xNome"
            rptListaOCPeriodo.Sections("Section5").Controls.Item("lblTotal").Visible = False
            rptListaOCPeriodo.Sections("Section5").Controls.Item("lblTotal").Caption = "0.00" 'ConvMoeda(ChkVal(vTotal, 0, cDecMoeda))
            rptListaOCPeriodo.Sections("Section5").Controls.Item("lblvProd").Visible = False
            rptListaOCPeriodo.Sections("Section5").Controls.Item("lblvProd").Caption = "0.00" 'ConvMoeda(ChkVal(vProd, 0, cDecMoeda))
            rptListaOCPeriodo.Show 1
            
    End If
    Rst.Close
End Sub
Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    
    limpaform
End Sub
Private Sub limpaform()
    txtNFIni.Text = ""
    txtNFFim.Text = ""
    dtpNFIni.Value = Date
    dtpNFFim.Value = Date
    
    optListagem_Click (0)
    
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
            LstOC
    End Select
End Sub


Private Sub txtNFFim_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub txtNFIni_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
End Sub
