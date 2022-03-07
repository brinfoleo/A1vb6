VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form formFaturamentoPVItem 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   12330
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstItens 
      Height          =   255
      Left            =   6480
      TabIndex        =   51
      ToolTipText     =   "De um duplo click ou <ENTER> para selecionar o item."
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      Caption         =   "|  Fisco  |"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   60
      TabIndex        =   35
      Top             =   2280
      Width           =   10395
      Begin VB.ComboBox cboCSTICMS 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox txtCSTDescricao 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   2220
         MultiLine       =   -1  'True
         TabIndex        =   49
         Text            =   "formFaturamentoPVItem.frx":0000
         Top             =   1020
         Width           =   3975
      End
      Begin VB.TextBox txtNCMDescricao 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   2220
         MultiLine       =   -1  'True
         TabIndex        =   48
         Text            =   "formFaturamentoPVItem.frx":0006
         Top             =   300
         Width           =   3975
      End
      Begin VB.Frame Frame7 
         Caption         =   "|  Impostos  |"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   6420
         TabIndex        =   37
         Top             =   180
         Width           =   3795
         Begin VB.TextBox txtpFCP 
            Height          =   285
            Left            =   840
            TabIndex        =   58
            Text            =   "Text1"
            Top             =   1125
            Width           =   735
         End
         Begin VB.TextBox txtvBCICMSST 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            TabIndex        =   55
            Text            =   "0,00"
            Top             =   1500
            Width           =   1935
         End
         Begin VB.TextBox txtvICMSST 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            TabIndex        =   54
            Text            =   "0,00"
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txtpICMS 
            Height          =   285
            Left            =   840
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   660
            Width           =   735
         End
         Begin VB.TextBox txtAliquotaIPI 
            Height          =   285
            Left            =   840
            MaxLength       =   6
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "FCP(%):"
            Height          =   195
            Left            =   180
            TabIndex        =   60
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label lblvICMSFCP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1740
            TabIndex        =   59
            Top             =   1140
            Width           =   1935
         End
         Begin VB.Label Label18 
            Caption         =   "Base Calc. ICMS-ST:"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "ICMS-ST:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   1980
            Width           =   675
         End
         Begin VB.Label lblvICMS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1740
            TabIndex        =   43
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblvIPI 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1740
            TabIndex        =   42
            Top             =   300
            Width           =   1935
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "ICMS(%):"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   675
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "IPI (%):"
            Height          =   195
            Left            =   180
            TabIndex        =   40
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.TextBox txtNCM 
         Height          =   285
         Left            =   720
         MaxLength       =   8
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "NCM:"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "CST:"
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   1080
         Width           =   435
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "|  TOTAIS  |"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6120
      TabIndex        =   15
      Top             =   5160
      Width           =   4335
      Begin VB.TextBox txtDescItem 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1740
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1080
         Width           =   2475
      End
      Begin VB.TextBox txtSubTotalProduto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   240
         Width           =   2475
      End
      Begin VB.TextBox txtTotalProduto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1860
         Width           =   2475
      End
      Begin VB.TextBox txtValorIPI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   660
         Width           =   2475
      End
      Begin VB.Label lblvICMSSTtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label8"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1740
         TabIndex        =   46
         Top             =   1440
         Width           =   2475
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "ICMS-ST:"
         Height          =   195
         Left            =   840
         TabIndex        =   24
         Top             =   1500
         Width           =   795
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Desconto:"
         Height          =   195
         Left            =   960
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Sub Total Produto:"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Total do Produto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1980
         Width           =   1515
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "IPI (Valor):"
         Height          =   195
         Left            =   900
         TabIndex        =   17
         Top             =   660
         Width           =   795
      End
   End
   Begin VB.CommandButton btoAdicionarItem 
      Caption         =   "&Adicionar Item"
      Height          =   735
      Left            =   10560
      Picture         =   "formFaturamentoPVItem.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton btoCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   10560
      Picture         =   "formFaturamentoPVItem.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   900
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "|  Produto  |"
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
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10395
      Begin VB.TextBox txtUltPreco 
         Appearance      =   0  'Flat
         Height          =   795
         Left            =   3660
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Text            =   "formFaturamentoPVItem.frx":0620
         Top             =   180
         Width           =   6495
      End
      Begin VB.TextBox txtProdutoID 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1560
         MaxLength       =   120
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1020
         Width           =   8535
      End
      Begin VB.ComboBox cboUnidade 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtQuantidade 
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1710
         Width           =   1095
      End
      Begin VB.TextBox txtValorUnitario 
         Height          =   285
         Left            =   4380
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   1710
         Width           =   1935
      End
      Begin VB.TextBox txtItemID 
         Height          =   285
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   47
         Top             =   1740
         Width           =   135
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Referencia:"
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   540
         TabIndex        =   13
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label11 
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   2760
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Preço Unitário:"
         Height          =   195
         Left            =   4380
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Unidade:"
         Height          =   195
         Left            =   1560
         TabIndex        =   10
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "ID:"
         Height          =   195
         Left            =   1140
         TabIndex        =   9
         Top             =   420
         Width           =   195
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   2415
      Left            =   60
      TabIndex        =   25
      Top             =   5100
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   4260
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pré-Venda"
      TabPicture(0)   =   "formFaturamentoPVItem.frx":0626
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "NFe"
      TabPicture(1)   =   "formFaturamentoPVItem.frx":0642
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Ordem de Compra"
      TabPicture(2)   =   "formFaturamentoPVItem.frx":065E
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Pedido:"
         Height          =   1875
         Left            =   120
         TabIndex        =   30
         Top             =   420
         Width           =   5715
         Begin VB.TextBox txtnPedido 
            Height          =   285
            Left            =   300
            MaxLength       =   15
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txtiPedido 
            Height          =   315
            Left            =   300
            MaxLength       =   6
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   1140
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Numero do Pedido:"
            Height          =   195
            Left            =   300
            TabIndex        =   34
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Item do Pedido:"
            Height          =   195
            Left            =   300
            TabIndex        =   33
            Top             =   900
            Width           =   1395
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Complemento de Descrição NFe"
         Height          =   1875
         Left            =   -74880
         TabIndex        =   28
         Top             =   420
         Width           =   5715
         Begin VB.CheckBox chkIndTot 
            Caption         =   "Item NÃO compõe Total da NFe"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   1560
            Width           =   3075
         End
         Begin VB.TextBox txtComplDescricaoNFe 
            Height          =   1215
            Left            =   120
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Text            =   "formFaturamentoPVItem.frx":067A
            Top             =   240
            Width           =   5475
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Complemento de Descrição Pré-Venda"
         Height          =   1875
         Left            =   -74880
         TabIndex        =   26
         Top             =   420
         Width           =   5715
         Begin VB.TextBox txtComplDescricaoPV 
            Height          =   1515
            Left            =   120
            MaxLength       =   65000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Text            =   "formFaturamentoPVItem.frx":0680
            Top             =   240
            Width           =   5475
         End
      End
   End
End
Attribute VB_Name = "formFaturamentoPVItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Retorno     As Variant
Dim IdReg       As Integer
Dim idCli       As Integer
Dim UFCli       As String
Dim bcICMS      As Integer '0=Mercadoria / 1 - Total da nota
Dim vBCICMSST   As String
Dim flag        As Boolean


Private Function pgDadosLinhaZero(sTexto As String) As Integer
    Dim i       As Integer
    Dim sobra   As String
    Dim Item    As Integer
    i = 0
    sobra = sTexto & "|"
    Do Until InStr(sobra, "|") = 0
        Select Case i
            Case 0 'Id do Cliente
                idCli = CInt(Mid(sobra, 1, InStr(sobra, "|") - 1))
                sobra = Mid(sobra, InStr(sobra, "|") + 1, Len(sobra))
            Case 1 ' 0= item novo
                Item = CInt(Mid(sobra, 1, InStr(sobra, "|") - 1))
                sobra = Mid(sobra, InStr(sobra, "|") + 1, Len(sobra))
            Case 2 'UF de Destino
                UFCli = Mid(sobra, 1, InStr(sobra, "|") - 1)
                sobra = Mid(sobra, InStr(sobra, "|") + 1, Len(sobra))
            Case 3 'Base de Calculo do ICMS 0/1
                bcICMS = CInt(Mid(sobra, 1, InStr(sobra, "|") - 1))
                sobra = Mid(sobra, InStr(sobra, "|") + 1, Len(sobra))
        End Select
        i = i + 1
    Loop
    pgDadosLinhaZero = Item
End Function
Public Function CarregarFormulario(vRec As Variant, cReg As Integer, nQtd As String) As Variant
    '08.12.2014- nQtd inclusa para que o sistema faca novos calculos e retorne
    'sem a necessidade de pressionar nenhum bt
    LimpaFormulario Me
    lstItens.Visible = False
    SSTab.Tab = 0
    Retorno = ""
    If pgDadosLinhaZero(CStr(vRec(0))) <> 0 Then
            
            IdReg = IIf(Trim(vRec(1)) = "", 0, vRec(1))
            txtItemID.Text = vRec(1)
            txtProdutoID.Text = vRec(2)
            txtDescricao.Text = vRec(3)
            txtNCM.Text = vRec(4)
            
            
            cboCSTICMS.Clear
            cboCSTICMS.AddItem vRec(5)
            If Len(Trim(vRec(5))) <> 0 Then
                cboCSTICMS.Text = cboCSTICMS.List(0)
            End If
            'txtCST.Text = vRec(5) 'pgDadosEstoqueProduto(CInt(vRec(1))).ICMSCST
            'pgCSTcomDescricao (IIf(Trim(vRec(1)) = "", 0, vRec(1)))
            
            
            cboUnidade.Clear
            cboUnidade.AddItem IIf(Trim(vRec(6)) = "", " ", vRec(6))
            cboUnidade.Text = cboUnidade.List(0)
            '08.12.2014  - Altecacao
            txtQuantidade.Text = IIf(Trim(nQtd) = "", vRec(7), Trim(nQtd))
            txtValorUnitario.Text = vRec(8)
            txtSubTotalProduto.Text = vRec(9)
            txtValorIPI.Text = vRec(10)
            txtDescItem.Text = vRec(11)
            txtTotalProduto.Text = vRec(12)
            txtAliquotaIPI.Text = vRec(13)
            txtpICMS.Text = Replace(vRec(14), "%", "")
            vBCICMSST = vRec(15)
            txtvBCICMSST.Text = vRec(15)
            
            txtvICMSST.Text = vRec(16)
            lblvICMSSTtotal.Caption = vRec(16)
            
            txtpFCP.Text = vRec(17)
            
            txtnPedido.Text = vRec(18)
        
            txtiPedido.Text = vRec(19)
            txtComplDescricaoPV.Text = vRec(20)
            txtComplDescricaoNFe.Text = vRec(21)
            'chkIndustrializacao.Value = IIf(vRec(21) = "I", 1, 0)
            chkIndTot.Value = IIf(vRec(22) = "S", 0, 1)
            PgUltimoVenda
        Else
            idCli = Left(vRec(0), InStr(vRec(0), "|") - 1)
            'UFCli = Right(vRec(0), 2)
    End If
    '08.12.2014 - Atualizado
    If Trim(nQtd) = "" Then
            Me.Show 1
        Else
            CalcVlItem
            btoAdicionarItem_Click
    End If
    CarregarFormulario = Retorno
End Function



Private Sub btoAdicionarItem_Click()
    
    'ID|Referencia|Descricao|NCM|Unid|Qtd|vUnit|SubTotal|vIPI|vDesc|vTotal|pIPI|pICMS|N.Ped|item ped|ComplDescricaoPV|ComplDescricaoNFe|pFCP
    
    Retorno = Array(txtItemID.Text, _
                    txtProdutoID.Text, _
                    txtDescricao.Text, _
                    txtNCM.Text, _
                    cboCSTICMS.Text, _
                    cboUnidade.Text, _
                    txtQuantidade.Text, _
                    txtValorUnitario.Text, _
                    txtSubTotalProduto.Text, _
                    txtValorIPI.Text, _
                    txtDescItem.Text, _
                    txtTotalProduto.Text, _
                    txtAliquotaIPI.Text, _
                    txtpICMS.Text, _
                    txtvICMSST.Text, _
                    txtnPedido.Text, _
                    txtiPedido.Text, _
                    txtComplDescricaoPV.Text, _
                    txtComplDescricaoNFe.Text, _
                    vBCICMSST, _
                    IIf(chkIndTot.Value = 0, "S", "N"), _
                    txtpFCP.Text)
    Unload Me
    
End Sub

Private Sub btoCancelar_Click()
    Retorno = ""
    Unload Me
End Sub





Private Sub cboCSTICMS_Click()
    If Trim(cboCSTICMS.Text) = "" Then
        txtCSTDescricao.Text = ""
        Exit Sub
    End If
    txtCSTDescricao.Text = UCase(PgDadosCST(cboCSTICMS.Text, "ICMS").Descricao)
End Sub

Private Sub cboCSTICMS_DropDown()
    Dim sSQL As String
    Dim Rst As Recordset
    Dim Tabela As String
    cboCSTICMS.Clear
    
    
    
    If PgDadosEmpresa(ID_Empresa).RegimeTrib = "3" Then
            Tabela = "B" ' ORDER BY cst"
        Else
            Tabela = "C" ' ORDER BY cst"
    End If
    
    sSQL = "SELECT * FROM TributacaoCST WHERE Tabela = '" & Tabela & "' ORDER BY CST"
    
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCSTICMS.AddItem UCase(Rst.Fields("CST"))
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboUnidade_DropDown()
  Dim Rst As Recordset
    cboUnidade.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueUnidadeMedida WHERE ID_Empresa = " & ID_Empresa)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboUnidade.AddItem UCase(Rst.Fields("Sigla"))
                Rst.MoveNext
            Loop
    End If
End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    'LimpaFormulario Me
    flag = False
    txtUltPreco.Visible = False
End Sub



Private Sub lstItens_DblClick()
    CarregarProduto
End Sub


Private Sub lstItens_GotFocus()
    lstItens.Visible = True
End Sub

Private Sub lstItens_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CarregarProduto
    End If
End Sub


Private Sub lstItens_LostFocus()
    lstItens.Visible = False
End Sub





Private Sub txtAliquotaIPI_Change()
    CalcVlItem
End Sub

Private Sub txtAliquotaIPI_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtAliquotaIPI.Text, KeyAscii, 3)
End Sub





Private Sub txtCSTDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub txtDescItem_Change()
    CalcVlItem
End Sub

Private Sub txtDescItem_GotFocus()
        txtDescItem.Text = ChkVal(txtDescItem.Text, 0, 2)
End Sub

Private Sub txtDescItem_KeyPress(KeyAscii As Integer)
    If txtDescItem.SelLength = Len(txtDescItem.Text) Then
        txtDescItem.Text = ""
    End If

    KeyAscii = ChkVal(txtDescItem.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtDescItem_LostFocus()
    txtDescItem.Text = ConvMoeda(txtDescItem.Text)
End Sub

Private Sub CaixaPesquisa()
    Dim sSQL As String
    Dim Rst As Recordset
    If flag = False Then
        lstItens.Visible = False
        Exit Sub
    End If
    
    With lstItens
        If Trim(txtDescricao.Text) = "" And flag = True Then
            .Visible = False
            Exit Sub
        End If
        .Left = txtDescricao.Left
        .Top = txtDescricao.Top + txtDescricao.Height
        .Width = txtDescricao.Width
        .Height = 5 * txtDescricao.Height
        .Visible = False
    
        .Clear
        sSQL = "SELECT * FROM estoqueproduto " & _
               "WHERE status='ATIVO' AND ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND " & _
               "Descricao LIKE '" & rc(txtDescricao.Text) & "%' ORDER BY Descricao LIMIT 200"
        Set Rst = RegistroBuscar(sSQL)
        If Rst.BOF And Rst.EOF Then
                .Visible = False
            Else
                If Rst.RecordCount = 1 Then
                        .Visible = False
                    Else
                        .Visible = True
                        Do Until Rst.EOF
                            .AddItem Left(String(6, "0"), 6 - Len(Rst.Fields("id"))) & Rst.Fields("id") & "  " & Rst.Fields("Descricao")
                            Rst.MoveNext
                        Loop
                End If
        End If
    End With
End Sub

Private Sub txtDescricao_Change()
    Me.Caption = flag
    If flag = True Then
        CaixaPesquisa
    End If
End Sub

Private Sub txtDescricao_GotFocus()
    flag = True
End Sub

Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        If lstItens.Visible = True Then
            lstItens.SetFocus
        End If
    End If
    
    If KeyCode = 114 Then
        PesquisarProduto
    End If
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim txt As String
    txt = txtDescricao.Text & Chr(KeyAscii)
    KeyAscii = IIf(numCaracter(txt) > 120, 0, KeyAscii)
End Sub
Private Function numCaracter(sTexto As String) As Integer
    Dim xTexto As String
    xTexto = sTexto
    xTexto = Replace(xTexto, "&", "&amp;")
    xTexto = Replace(xTexto, "<", "&lt;")
    xTexto = Replace(xTexto, ">", "&gt;")
    xTexto = Replace(xTexto, """", "&quot;")
    xTexto = Replace(xTexto, "'", "&#39;")
    
    numCaracter = Len(xTexto)
End Function
Private Sub txtDescricao_LostFocus()
    flag = False
    'lstItens.Visible = False
End Sub

Private Sub txtItemID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarProduto
    End If

End Sub


Private Sub txtItemID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PesquisarProduto (IIf(Trim(txtItemID.Text) = "", 0, Trim(txtItemID.Text)))
    End If
    KeyAscii = SoNumeros(KeyAscii)
End Sub







Private Sub txtNCM_Change()
    If Trim(txtNCM.Text = "") Then
        txtNCMDescricao.Text = ""
        Exit Sub
    End If
    txtNCMDescricao.Text = PgDadosNCM("NCM", Trim(txtNCM.Text), "S").Descricao & " [" & PgDadosNCM("NCM", Trim(txtNCM.Text), "S").pIPI & "%]"
End Sub

Private Sub txtNCMDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtpFCP_Change()
    CalcVlItem
End Sub

Private Sub txtpFCP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtpFCP.Text, KeyAscii, 2)
End Sub

Private Sub txtpICMS_Change()
    CalcVlItem
End Sub

Private Sub txtpICMS_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtpICMS.Text, KeyAscii, 2)
End Sub

Private Sub txtProdutoID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarProduto
    End If
End Sub
Private Sub txtQuantidade_Change()
    CalcVlItem
End Sub

Private Sub txtQuantidade_GotFocus()
    With txtQuantidade
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    If txtValorUnitario.SelLength = Len(txtValorUnitario.Text) Then
        txtValorUnitario.Text = ""
    End If
    KeyAscii = ChkVal(txtQuantidade.Text, KeyAscii, cDecQtd)
End Sub



Private Sub txtValorUnitario_LostFocus()
    txtValorUnitario.Text = ConvMoeda(txtValorUnitario.Text)
End Sub

Private Sub txtSubTotalProduto_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtTotalProduto_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub




Private Sub txtValorIPI_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub txtValorUnitario_Change()
    CalcVlItem
End Sub

Private Sub txtValorUnitario_GotFocus()
    With txtValorUnitario
        .Text = ChkVal(.Text, 0, cDecMoeda)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtValorUnitario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If txtValorUnitario.SelLength = Len(txtValorUnitario.Text) Then
        txtValorUnitario.Text = ""
    End If
    KeyAscii = ChkVal(txtValorUnitario.Text, KeyAscii, cDecMoeda)
End Sub
Private Sub PesquisarProduto(Optional Id As Integer)
    On Error GoTo TratarErro
    Dim pICMS   As String
    Dim cst     As String
    If Trim(Id) = "0" Then
            Id = formBuscar.IniciarBusca("EstoqueProduto") ', , , , , "status='ATIVO'")
            If Trim(Id) = 0 Then Exit Sub
    End If
    txtItemID.Text = Id
    IdReg = Id
    PgUltimoVenda
    If Trim(txtDescricao.Text) <> "" Then
        If MsgBox("Dados do produto já preenchido deseja substituir?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    
    If Trim(pgDadosEstoqueProduto(Id).Descricao) = "" Then Exit Sub
    
    txtProdutoID.Text = pgDadosEstoqueProduto(Id).Referencia ' IIf(IsNull(Rst.Fields("Referencia")), "", Rst.Fields("Referencia"))
    txtDescricao.Text = pgDadosEstoqueProduto(Id).Descricao ' IIf(IsNull(Rst.Fields("Descricao")), "", Rst.Fields("Descricao"))
    cboUnidade.Clear
    cboUnidade.AddItem pgDadosEstoqueProduto(Id).Unidade  'IIf(IsNull(Rst.Fields("Unidade")), " ", Rst.Fields("Unidade"))
    cboUnidade.Text = cboUnidade.List(0)
    txtNCM.Text = pgDadosEstoqueProduto(Id).NCM  'IIf(IsNull(Rst.Fields("NCM")), "", Rst.Fields("NCM"))
    txtNCMDescricao.Text = PgDadosNCM("NCM", Trim(txtNCM.Text), "S").Descricao & " [" & PgDadosNCM("NCM", Trim(txtNCM.Text), "S").pIPI & "%]"
    txtQuantidade.Text = "0"
    
    cboCSTICMS.Clear
    cboCSTICMS.AddItem pgDadosEstoqueProduto(Id).ICMSCST
    cboCSTICMS.Text = cboCSTICMS.List(0)
    'txtCST.Text = pgDadosEstoqueProduto(Id).ICMSCST
    'pgCSTcomDescricao (Id)
    
    txtValorUnitario.Text = IIf(Trim(pgDadosEstoqueProduto(Id).VlTabela) = "", "0.00", pgDadosEstoqueProduto(Id).VlTabela) 'ConvMoeda(Rst.Fields("preco"))
    txtAliquotaIPI.Text = pgDadosEstoqueProduto(Id).IPIAliquota  'IIf(IsNull(Rst.Fields("ipialiquota")), "0.00", Rst.Fields("ipialiquota"))
    txtDescItem.Text = ConvMoeda("0")
    'ICMS

   'Material cadastrado no Estoque
    'If idCli = 0 Then
    If Trim(UFCli) = "" Then
            pICMS = "0"
        Else
            cst = Trim(cboCSTICMS.Text)
            
            'pICMS = pgDadosICMS(PgDadosCliente(idCli).UF, 0).ICMS
            
            If Trim(UFCli) <> Trim(PgDadosEmpresa(ID_Empresa).uf) Then
                    'Pega o ICMS Interestadual
                    pICMS = pgDadosICMS(UFCli, 0).ICMS
                Else
                    'Pega o ICMS Interno
                    pICMS = pgDadosICMS(UFCli, 0).ICMSInt
            End If
            
            pICMS = IIf(Trim(pgAliqDifICMS(txtNCM.Text, PgDadosCliente(idCli).uf)) = "", pICMS, pgAliqDifICMS(txtNCM.Text, PgDadosCliente(idCli).uf))
            pICMS = IIf(cst = "60", "0", pICMS)
            'pICMS = pICMS
    End If
    txtpICMS.Text = pICMS
    'ICMS zerado nao disponibiliza FCP
    If pICMS = 0 Then
            txtpFCP.Text = "0"
        Else
            txtpFCP.Text = pgDadosICMS(UFCli, 0).ICMSFECP
    End If
    Exit Sub
TratarErro:
    RegLog "0", "0", "[formFaturamentoPVItem.pesquisarProduto] - " & Err.Number & " - " & Err.Description
End Sub

Private Sub CalcVlItem()
    Dim SubTotalItem    As String
    Dim IPIItem         As String
    Dim TotalItem       As String
    Dim vICMSST         As String
    Dim vICMS           As String
    Dim sMVA            As String
    Dim vICMSFCP        As String
    
    
    SubTotalItem = Val(ChkVal(txtQuantidade.Text, 0, cDecQtd)) * Val(ChkVal(txtValorUnitario.Text, 0, cDecMoeda))
    IPIItem = (Val(ChkVal(SubTotalItem, 0, cDecMoeda)) * Val(ChkVal(txtAliquotaIPI.Text, 0, cDecMoeda))) / 100
    TotalItem = (Val(ChkVal(SubTotalItem, 0, cDecMoeda)) + Val(ChkVal(IPIItem, 0, cDecMoeda))) - Val(ChkVal(txtDescItem.Text, 0, cDecMoeda))
    
'    '10/07/2017 - Todo o item foi comentado para que o usuario
'    '             coloque manualmente o valor do ICMS ST
'
'    If IdReg = 0 Then
'            sMVA = 0
'        Else
'            sMVA = pgDadosEstoqueProduto(IdReg).MVA
'    End If
'
'
'    If IdReg <> 0 Then
'        If UFCli <> PgDadosEmpresa(ID_Empresa).UF And pgDadosEstoqueProduto(IdReg).ICMSCST = "60" Then
'                '##############################################################################################
'                '18/04/2012
'                'Calculo_ICMSST = ERRADO pois o valor deve ser calculado com base no valor do produto e valor total
'                '                 do produto e nao apenas em um unico valor.
'                '##############################################################################################
'                vBCICMSST = Calculo_ICMSST(PgDadosEmpresa(ID_Empresa).UF, UFCli, sMVA, TotalItem).vBCICMSST
'                vICMSST = Calculo_ICMSST(PgDadosEmpresa(ID_Empresa).UF, UFCli, sMVA, TotalItem).vICMSST
'            Else
'                vICMSST = 0
'        End If
'    End If
'    '########################################################################################################################
'    '### 12/01/2012 - Os paremetros abaixo forao inclusos devido haver a necessidade de saber se o estado de destino tem
'    '###              convenio de ICMS ST para efetuar o Calculo da ST.
'    '###              Projeto: Gerar um modulo que faça todo o tipo de analise no pedido/NF emitida para avaliar se deve ou
'    '###                       nao incluir a ST.
'            vBCICMSST = 0
'            vICMSST = 0
'    '########################################################################################################################
'
    vBCICMSST = ChkVal(txtvBCICMSST.Text, 0, cDecMoeda)
    vICMSST = ChkVal(txtvICMSST.Text, 0, cDecMoeda)
'#####################################################################################
    
    'Calculo do ICMS
    If bcICMS = 0 Then
            vICMS = Val(ChkVal(txtpICMS.Text, 0, cDecMoeda)) * Val(ChkVal(SubTotalItem, 0, cDecMoeda))
            vICMS = Val(ChkVal(vICMS, 0, cDecMoeda)) / 100
            
            'FCP
            vICMSFCP = Val(ChkVal(txtpFCP.Text, 0, cDecMoeda)) * Val(ChkVal(SubTotalItem, 0, cDecMoeda))
            vICMSFCP = Val(ChkVal(vICMSFCP, 0, cDecMoeda)) / 100
    
        Else
            'ICMS
            vICMS = Val(ChkVal(txtpICMS.Text, 0, cDecMoeda)) * Val(ChkVal(TotalItem, 0, cDecMoeda))
            vICMS = Val(ChkVal(vICMS, 0, cDecMoeda)) / 100
    
            'FCP
            vICMSFCP = Val(ChkVal(txtpFCP.Text, 0, cDecMoeda)) * Val(ChkVal(TotalItem, 0, cDecMoeda))
            vICMSFCP = Val(ChkVal(vICMSFCP, 0, cDecMoeda)) / 100
    
    End If
    
    TotalItem = ((Val(ChkVal(SubTotalItem, 0, cDecMoeda)) + Val(ChkVal(IPIItem, 0, cDecMoeda))) + Val(ChkVal(vICMSST, 0, cDecMoeda))) - Val(ChkVal(txtDescItem.Text, 0, cDecMoeda))

    txtSubTotalProduto.Text = ConvMoeda(SubTotalItem)
    
    lblvIPI.Caption = ConvMoeda(IPIItem)
    txtValorIPI.Text = ConvMoeda(IPIItem)
    
    lblvICMS.Caption = ConvMoeda(vICMS)
    lblvICMSFCP.Caption = ConvMoeda(vICMSFCP)
    
    txtTotalProduto.Text = ConvMoeda(TotalItem)
    'txtvICMSST.Text = ConvMoeda(vICMSST)
    lblvICMSSTtotal.Caption = ConvMoeda(vICMSST)
End Sub

Private Sub CarregarProduto()
    If Trim(lstItens.Text) = "" Then
        Exit Sub
    End If
    PesquisarProduto Left(lstItens.Text, 6)
    lstItens.Visible = False
End Sub


'Private Sub pgCSTcomDescricao(idMaterial As Integer)
'    Dim sSQL    As String
'    Dim Rst     As Recordset
'    If Trim(idMaterial) = "" Or Trim(idMaterial) = 0 Then Exit Sub
'    sSQL = "SELECT * FROM TributacaoCST WHERE ID_Empresa = " & ID_Empresa & _
'           " AND Tabela = 'B' AND CST=" & pgDadosEstoqueProduto(idMaterial).ICMSCST
'    Set Rst = RegistroBuscar(sSQL)
'    If Rst Is Nothing Then Exit Sub
'    If Rst.BOF And Rst.EOF Then
'            txtCST.Text = ""
'            txtCSTDescricao.Text = ""
'        Else
'            txtCST.Text = Rst.Fields("CST")
'            txtCSTDescricao.Text = Rst.Fields("descricao")
'    End If
'    Rst.Close
'End Sub

Private Sub PgUltimoVenda()
    On Error GoTo TrtErroPgUltPreco
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    sSQL = "SELECT FNFe.ide_dEmi, FNFe.idNFe, FNFe.Dest_IdDest, FNFeI.idNFe, FNFeI.det_xProd, FNFeI.det_vUnCom, FNFeI.det_uCom, FNFeI.det_qCom, FNFeI.det_idProduto " & _
           "FROM FaturamentoNFe AS FNFe, FaturamentoNFeItens AS FNFeI " & _
           "WHERE FNFe.dest_idDest = " & idCli & " AND FNFeI.idNFe = FNFe.idNFe AND FNFeI.det_idProduto = " & IdReg
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            txtUltPreco.Text = ""
            txtUltPreco.Visible = False
        Else
            Rst.MoveLast
            txtUltPreco.Text = "Venda em : " & Rst.Fields("ide_dEmi") & vbCrLf & _
                               "Descrição: " & Rst.Fields("det_xProd") & vbCrLf & _
                               "Qtd/Unid.: " & ChkVal(Rst.Fields("det_qCom"), 0, cDecQtd) & "/" & Rst.Fields("det_uCom") & _
                               "              Val.Unit.: " & ConvMoeda(Rst.Fields("det_vUnCom"))
            txtUltPreco.Visible = True
            'MsgBox Rst.Fields("idnfe")
    End If
    Rst.Close
    Exit Sub
TrtErroPgUltPreco:
    txtUltPreco.Visible = False
End Sub

Private Sub txtvBCICMSST_Change()
    CalcVlItem
End Sub

Private Sub txtvBCICMSST_GotFocus()
    With txtvBCICMSST
        .Text = ChkVal(.Text, 0, 2)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtvBCICMSST_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If txtvBCICMSST.SelLength = Len(txtvBCICMSST.Text) Then
        txtvBCICMSST.Text = ""
    End If
    KeyAscii = ChkVal(txtvBCICMSST.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvBCICMSST_LostFocus()
    txtvBCICMSST.Text = ConvMoeda(txtvBCICMSST.Text)
End Sub

Private Sub txtvICMSST_Change()
    CalcVlItem
End Sub

Private Sub txtvICMSST_GotFocus()
    With txtvICMSST
        .Text = ChkVal(.Text, 0, 2)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtvICMSST_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If txtvICMSST.SelLength = Len(txtvICMSST.Text) Then
        txtvICMSST.Text = ""
    End If
    KeyAscii = ChkVal(txtvICMSST.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvICMSST_LostFocus()
    txtvICMSST.Text = ConvMoeda(txtvICMSST.Text)
End Sub
