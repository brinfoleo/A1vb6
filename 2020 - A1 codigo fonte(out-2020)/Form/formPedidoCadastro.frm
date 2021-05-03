VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form formPedidoCadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedido"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   13380
   Begin VB.Frame Frame2 
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   120
      TabIndex        =   18
      Top             =   2940
      Width           =   13155
      Begin VB.Frame Frame3 
         Caption         =   "Totais"
         Height          =   975
         Left            =   120
         TabIndex        =   42
         Top             =   2820
         Width           =   7635
         Begin VB.Frame Frame4 
            Caption         =   "Itens"
            Height          =   675
            Left            =   120
            TabIndex        =   49
            Top             =   180
            Width           =   1815
            Begin VB.Label lblItens 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0000"
               Height          =   315
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Valor da Mercadoria"
            Height          =   675
            Left            =   1980
            TabIndex        =   47
            Top             =   180
            Width           =   1815
            Begin VB.Label lblMercadoria 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "R$ 0,00"
               Height          =   315
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Valor do IPI"
            Height          =   675
            Left            =   3840
            TabIndex        =   45
            Top             =   180
            Width           =   1815
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "R$ 0,00"
               Height          =   315
               Left            =   120
               TabIndex        =   46
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Total do Pedido"
            Height          =   675
            Left            =   5700
            TabIndex        =   43
            Top             =   180
            Width           =   1815
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "R$ 0,00"
               Height          =   315
               Left            =   120
               TabIndex        =   44
               Top             =   240
               Width           =   1575
            End
         End
      End
      Begin VB.CommandButton btoReomoverItem 
         Caption         =   "Remover Item"
         Height          =   315
         Left            =   11400
         TabIndex        =   39
         Top             =   1320
         Width           =   1395
      End
      Begin VB.CommandButton btoAdicionarItem 
         Caption         =   "Adicionar Item"
         Height          =   375
         Left            =   11400
         TabIndex        =   38
         Top             =   900
         Width           =   1395
      End
      Begin MSFlexGridLib.MSFlexGrid msfgItens 
         Height          =   1095
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   1931
         _Version        =   393216
         Cols            =   10
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"formPedidoCadastro.frx":0000
      End
      Begin VB.TextBox txtSubTotalProduto 
         Height          =   285
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox txtValorIPI 
         Height          =   285
         Left            =   7560
         MaxLength       =   15
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1200
         Width           =   1035
      End
      Begin VB.TextBox txtAliquotaIPI 
         Height          =   285
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTotalProduto 
         Height          =   285
         Left            =   8820
         MaxLength       =   15
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtValorUnitario 
         Height          =   285
         Left            =   2760
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtQuantidade 
         Height          =   315
         Left            =   1260
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox cboUnidade 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   2100
         MaxLength       =   120
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   540
         Width           =   10755
      End
      Begin VB.TextBox txtProdutoID 
         Height          =   285
         Left            =   180
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Sub Total Produto"
         Height          =   195
         Left            =   4800
         TabIndex        =   36
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label16 
         Caption         =   "IPI (Valor):"
         Height          =   255
         Left            =   7620
         TabIndex        =   32
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label15 
         Caption         =   "IPI (%):"
         Height          =   195
         Left            =   6600
         TabIndex        =   31
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Total do Produto"
         Height          =   195
         Left            =   8820
         TabIndex        =   29
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label13 
         Caption         =   "Unidade:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label12 
         Caption         =   "Preço Unitário:"
         Height          =   195
         Left            =   2760
         TabIndex        =   24
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   1260
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   2100
         TabIndex        =   20
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Código do Produto:"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   13155
      Begin VB.ComboBox cboTransportadora 
         Height          =   315
         Left            =   1260
         TabIndex        =   41
         Text            =   "Combo1"
         Top             =   1080
         Width           =   6435
      End
      Begin VB.ComboBox cboFormaPagamento 
         Height          =   315
         Left            =   10260
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1380
         Width           =   2775
      End
      Begin VB.TextBox txtObs 
         Height          =   555
         Left            =   1260
         MaxLength       =   65000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "formPedidoCadastro.frx":00A6
         Top             =   1560
         Width           =   3315
      End
      Begin VB.ComboBox cboCondicoesPagamento 
         Height          =   315
         Left            =   10260
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1020
         Width           =   2775
      End
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   10260
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   1260
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   660
         Width           =   6435
      End
      Begin VB.TextBox txtStatus 
         Height          =   285
         Left            =   8640
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   240
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpEmissao 
         Height          =   315
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   40517
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1260
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Transportadora:"
         Height          =   195
         Left            =   60
         TabIndex        =   40
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "Forma de Pagamento:"
         Height          =   195
         Left            =   8580
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Observações:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Condições de Pagamento:"
         Height          =   195
         Left            =   8280
         TabIndex        =   7
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Vendedor:"
         Height          =   255
         Left            =   9420
         TabIndex        =   6
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Status do Pedido"
         Height          =   195
         Left            =   7320
         TabIndex        =   4
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Data Emissão:"
         Height          =   195
         Left            =   2520
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pedido:"
         Height          =   255
         Left            =   660
         TabIndex        =   2
         Top             =   300
         Width           =   555
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alterar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pesquisar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formPedidoCadastro.frx":00AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formPedidoCadastro.frx":04FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formPedidoCadastro.frx":0818
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formPedidoCadastro.frx":10AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formPedidoCadastro.frx":22FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formPedidoCadastro.frx":2BD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formPedidoCadastro.frx":3468
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formPedidoCadastro.frx":3CFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formPedidoCadastro.frx":4F4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formPedidoCadastro.frx":5266
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formPedidoCadastro.frx":5580
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formPedidoCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim IdReg     As Integer 'ID do Pedido
Dim IdItem    As Integer 'Id dos itens do pedido
Dim strTabela   As String


Private Sub PesquisarRegistro()
    Dim psqTMP  As String
    psqTMP = FormBusca.IniciarBusca(strTabela)
    IdReg = IIf(psqTMP = "", 0, psqTMP)
    
    If IdReg = 0 Then
            LimpaFormulario Me 'me
        Else
            MostrarDados
    End If
End Sub




Private Sub btoAdicionarItem_Click()
    With msfgItens
        .TextMatrix(.Row, 0) = IdItem
        .TextMatrix(.Row, 1) = txtProdutoID.Text
        .TextMatrix(.Row, 2) = txtDescricao.Text
        .TextMatrix(.Row, 3) = cboUnidade.Text
        .TextMatrix(.Row, 4) = txtQuantidade.Text
        .TextMatrix(.Row, 5) = txtValorUnitario.Text
        .TextMatrix(.Row, 6) = txtSubTotalProduto.Text
        .TextMatrix(.Row, 7) = txtAliquotaIPI.Text
        .TextMatrix(.Row, 8) = txtValorIPI.Text
        .TextMatrix(.Row, 9) = txtTotalProduto.Text
    End With
End Sub

Private Sub btoReomoverItem_Click()
    If MsgBox("Deseja realmente remover este item?", vbYesNo, "Removendo Item do Pedido") = vbYes Then
        msfgItens.RemoveItem msfgItens.Row
    End If
End Sub


Private Sub cboCliente_DropDown()
    Dim rst As Recordset
    
    Set rst = RegistroBuscar("SELECT * FROM Clientes WHERE Nome LIKE '" & cboCliente.Text & "%'")
    If rst.BOF And rst.EOF Then
            cboCliente.Clear
            Exit Sub
        Else
            cboCliente.Clear
            rst.MoveFirst
            Do Until rst.EOF
                cboCliente.AddItem Left(String(6, "0"), 6 - Len(Trim(rst.Fields("ID")))) & rst.Fields("ID") & _
                " - " & _
                rst.Fields("Nome")
                rst.MoveNext
            Loop
    End If

End Sub
Private Sub cbotransportadora_DropDown()
    Dim rst As Recordset
    
    Set rst = RegistroBuscar("SELECT * FROM Transportadoras WHERE Nome LIKE '" & cboTransportadora.Text & "%'")
    If rst.BOF And rst.EOF Then
            cboTransportadora.Clear
            Exit Sub
        Else
            cboTransportadora.Clear
            rst.MoveFirst
            Do Until rst.EOF
                cboTransportadora.AddItem Left(String(6, "0"), 6 - Len(Trim(rst.Fields("ID")))) & rst.Fields("ID") & _
                " - " & _
                rst.Fields("Nome")
                rst.MoveNext
            Loop
    End If

End Sub

'Private Sub cboCliente_KeyPress(KeyAscii As Integer)
'    If KeyCode = 114 Then
'        PesquisarRegistro
'    End If
'End Sub

Private Sub cboCondicoesPagamento_DropDown()
    Dim rst As Recordset
    cboCondicoesPagamento.Clear
    Set rst = RegistroBuscar("SELECT * FROM FinanceiroCondicoespagamento")
    If rst.BOF And rst.EOF Then
            Exit Sub
        Else
            rst.MoveFirst
            Do Until rst.EOF
                cboCondicoesPagamento.AddItem rst.Fields("Descricao")
                rst.MoveNext
            Loop
    End If
            
End Sub


Private Sub cboFormaPagamento_DropDown()
    Dim rst As Recordset
    cboFormaPagamento.Clear
    Set rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento")
    If rst.BOF And rst.EOF Then
            Exit Sub
        Else
            rst.MoveFirst
            Do Until rst.EOF
                cboFormaPagamento.AddItem rst.Fields("Descricao")
                rst.MoveNext
            Loop
    End If

End Sub



Private Sub cboUnidade_DropDown()
  Dim rst As Recordset
    cboUnidade.Clear
    Set rst = RegistroBuscar("SELECT * FROM EstoqueUnidadeMedida")
    If rst.BOF And rst.EOF Then
            Exit Sub
        Else
            rst.MoveFirst
            Do Until rst.EOF
                cboUnidade.AddItem rst.Fields("Sigla")
                rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboVendedor_DropDown()
    Dim rst As Recordset
    cboVendedor.Clear
    Set rst = RegistroBuscar("SELECT * FROM RHFuncionarioCadastro")
    If rst.BOF And rst.EOF Then
            Exit Sub
        Else
            rst.MoveFirst
            Do Until rst.EOF
                cboVendedor.AddItem Left(String(4, "0"), 4 - Len(Trim(rst.Fields("ID")))) & rst.Fields("ID") & " - " & rst.Fields("Nome")
                rst.MoveNext
            Loop
    End If
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    HDForm Me, False
    HDMenu Me, True
    
    
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
            IdReg = 0
            HDMenu Me, False
            HDForm Me, True
           LimpaFormulario Me
        Case "Alterar"
            If IdReg = 0 Then
                MsgBox "Selecione uma Grupo"
                Exit Sub
            End If
            HDForm Me, True
            HDMenu Me, False
        Case "Excluir"
            If IdReg = 0 Then
                    MsgBox "Selecione um Registro"
                    Exit Sub
                Else
                    If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                               vbCrLf & _
                               "Descrição.: " & txtDescricao.Text, vbYesNo + vbCritical) = vbYes Then
                               
                        If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                            LimpaFormulario Me
                        End If
                    End If
            End If
        Case "Pesquisar"
            PesquisarRegistro
            
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDForm Me, False
                'LimpaFormulario me
                'txtCNPJ.Enabled = True
            End If
            
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            
        Case "Manutenção da Tabela"
            formManutencaoTabelas.IniciarManutencao Me
    End Select
End Sub

Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim I           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer 'Contador de Registros
    cReg = 0
    For I = 0 To Me.Controls.Count - 1
        Set Controle = Me.Controls(I)
        
        If TypeOf Controle Is TextBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Text, "S")
            cReg = cReg + 1
        End If
        If TypeOf Controle Is ComboBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Text, "S")
            cReg = cReg + 1
        End If
        If TypeOf Controle Is CheckBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Value, "S")
            cReg = cReg + 1
        End If
    Next
    
     
    If IdReg = 0 Then
            If RegistroIncluir(strTabela, vReg, cReg) = False Then
                    MsgBox "Erro ao Incluir"
                    grvRegistro = False
                Else
                    grvRegistro = True
            End If
        Else
            If RegistroAlterar(strTabela, vReg, cReg, "Id = " & IdReg) = False Then
                    MsgBox "Erro ao Alterar."
                    grvRegistro = False
                Else
                    grvRegistro = True
                
            End If
    End If



End Function


Private Sub MostrarDados()
    Dim sSQL As String
    sSQL = "SELECT * FROM " & strTabela & " WHERE Id = " & IdReg

    ExibirDados Me, sSQL


End Sub





Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
End Sub

Private Sub txtSigla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If

End Sub



Private Sub txtPrecoUnitario_Change()

End Sub

Private Sub txtStatus_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
