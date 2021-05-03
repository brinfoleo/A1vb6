VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formFaturamentoNFeCancelada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Cancelamento de NF-e"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9930
   Begin VB.Frame Frame2 
      Caption         =   "Protocolo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   9735
      Begin VB.TextBox txtProtocolo 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   240
         Width           =   9255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   9795
      Begin VB.Frame Frame3 
         Caption         =   "Autorização de Uso"
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
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   9555
         Begin VB.TextBox txtHrAutorizacao 
            Height          =   285
            Left            =   6000
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   540
            Width           =   1695
         End
         Begin VB.TextBox txtDtAutorizacao 
            Height          =   285
            Left            =   4140
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   540
            Width           =   1695
         End
         Begin VB.TextBox txtnProt 
            Height          =   315
            Left            =   120
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   540
            Width           =   3855
         End
         Begin VB.Label Label5 
            Caption         =   "Hora:"
            Height          =   195
            Left            =   6000
            TabIndex        =   14
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Data:"
            Height          =   195
            Left            =   4140
            TabIndex        =   11
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Protocolo:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   300
            Width           =   2415
         End
      End
      Begin VB.TextBox txtchNFe 
         Height          =   285
         Left            =   120
         MaxLength       =   60
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   420
         Width           =   7575
      End
      Begin VB.TextBox txtxJust 
         Height          =   1335
         Left            =   120
         MaxLength       =   255
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2100
         Width           =   9555
      End
      Begin VB.Label Label3 
         Caption         =   "Justificativa:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Chave de acesso:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "NF-e"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Protocolo de Cancelamento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pesquisar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Picture         =   "formFaturamentoNFeCancelada.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCancelada.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCancelada.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCancelada.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCancelada.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCancelada.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCancelada.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCancelada.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCancelada.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCancelada.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCancelada.frx":54D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFaturamentoNFeCancelada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg    As Integer
Public Sub CancelamentoNfe(idNFe As Integer)
    If idNFe = 0 Then Exit Sub
  
    HDMenu Me, False
    HDForm Me, True
    LimpaFormulario Me
    tbMenu.Buttons.Item(4).Enabled = True
    IdReg = idNFe
    PesquisarRegistro
    formFaturamentoNFeCancelada.Show
End Sub
Private Function grvRegistro() As Boolean
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    'Dim nmArq       As String
    Dim chvnfe      As String
    
    If ValidarDados = False Then
        grvRegistro = False
        Exit Function
    End If
    
    
    cReg = 0
    vReg(cReg) = Array("canc_xJust", txtxJust.Text, "S") ': cReg = cReg + 1

     IdReg = RegistroAlterar("FaturamentoNFe", vReg, cReg, "id=" & IdReg)
            If IdReg = 0 Then
                    MsgBox "Erro ao incluir Cancelamento de NF-e.", vbInformation, "Aviso"
                    grvRegistro = False
                Else
                    grvRegistro = True
                    chvnfe = Trim(txtchNFe.Text)
                    'nmArq = Trim(chvnfe) & "-ped-can.txt"
                    Cancelar_NFe (chvnfe)
                    estornoDeEstoque (chvnfe)
                    MsgBox "Solicitação de Cancelamento de NF-e em andamento.", vbInformation, "Aviso"
                    
            End If

End Function
Private Sub estornoDeEstoque(NFe As String)
'#
'# Estorna os materiais da NFe Cancelada
'# 25/10/2013
'#

    Dim sSQL    As String
    Dim Rst     As Recordset
    'Dar entrada no Estoque do produto
    sSQL = "SELECT * FROM EstoqueKardex WHERE NFe = '" & NFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'Rst.Close
            'Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                MovimentarEstoque "e", Rst.Fields("IdProduto"), Date, Rst.Fields("Documento"), Rst.Fields("Quantidade"), Rst.Fields("ValorUnitario"), Rst.Fields("ValorTotal"), "Estorno devido Nota Fiscal Cancelada", _
                                  Rst.Fields("Nome"), NFe, Rst.Fields("IDNome"), Rst.Fields("DocNome")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    
End Sub
Private Function ValidarDados() As Boolean
    Dim difDatas    As Integer 'Armazena a dif entre datas para saber se a NF pode ser cancelada

    If Len(Trim(txtProtocolo.Text)) > 5 Then
        MsgBox "NFe já cancelada!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    If Len(Trim(txtxJust.Text)) < 15 Then
        MsgBox "A justificativa deve conter no minimo 15 caracteres.", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    If Len(Trim(txtnProt.Text)) < 1 Then
        MsgBox "Não é possível cancelar NF-e que não tenha sido autorizada.", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    
    '*****************************************************************************************
    '*** Data: 11/07/2011
    '*** Obj: Impedir o cancelamento de NFe fora do prazo previsto por lei
    '*****************************************************************************************
    
    '02/01/2012 - Modificado devido ATO COTEPE ICMS N. 35, DE 24 DE NOVEMBRO 2010 para prazo de 24 hrs
    'difDatas = Date - CDate(txtDtAutorizacao.Text)
    
    difDatas = DateDiff("h", Trim(txtDtAutorizacao.Text) & " " & Trim(Left(Trim(txtHrAutorizacao.Text), 8)), Now)
    
    
    If difDatas > PgDadosConfig.NFePrazoCancelamento Then
        MsgBox "Prazo limite para cancelamento de Nota Fiscal ultrapassado!", vbCritical, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    '*****************************************************************************************
    ValidarDados = True
End Function

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    IdReg = 0
    LimpaFormulario Me
    HDForm Me, False
    HDMenu Me, True
End Sub

Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    IdReg = 0
    HDMenu Me, False
    HDForm Me, True
    LimpaFormulario Me
    tbMenu.Buttons.Item(4).Enabled = True
            
End Sub
Private Sub Imprimir(opcao As Integer)
'##########################################################
'### Opcao = 1 - NFe
'###         2 - Protocolo de Cancelamento
'##########################################################
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If
    If Trim(txtchNFe.Text) = "" Then
        MsgBox "Selecione uma Nota Fiscal.", vbInformation, "Aviso"
        Exit Sub
    End If
    Select Case opcao
        Case 1
            ImprimirDANFE (Trim(txtchNFe.Text))
        Case 2
            ImprimirProtCanc (Trim(txtchNFe.Text))
    End Select
    
End Sub



Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
            Incluir
        Case "Imprimir"
            Imprimir 1
        Case "Pesquisar"
            IdReg = 0
            PesquisarRegistro
            
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDForm Me, False
                RegistroExcluir "financeirocontasprcadastro", "ide_NFe = '" & txtchNFe.Text & "'"
            End If
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            'txtID.Enabled = True
        Case "Manutenção da Tabela"
            'formManutencaoTabelas.IniciarManutencao Me
            'MontarBaseDeDados
    End Select
End Sub





Private Sub PesquisarRegistro()
    Dim sSQL        As String
    Dim Rst         As Recordset
    
    
    If Trim(IdReg) = 0 Then
        IdReg = formBuscar.IniciarBusca("FaturamentoNFe")
    End If
    
    If IdReg = 0 Then
        LimpaFormulario Me
        Exit Sub
    End If
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado."
            LimpaFormulario Me
        Else
            Rst.MoveFirst
            txtchNFe.Text = Rst.Fields("idNFe")
            txtnProt.Text = IIf(IsNull(Rst.Fields("nProt")), "", Rst.Fields("nProt"))
            If IsNull(Rst.Fields("nProt")) Then
                    txtDtAutorizacao.Text = ""
                    txtHrAutorizacao.Text = ""
                Else
                    txtDtAutorizacao.Text = IIf(IsNull(Rst.Fields("dhProt")), Rst.Fields("ide_demi"), Mid(Rst.Fields("dhProt"), 1, InStr(Rst.Fields("dhProt"), " ")))
                    txtHrAutorizacao.Text = IIf(IsNull(Rst.Fields("dhProt")), Rst.Fields("ide_demi"), Mid(Rst.Fields("dhProt"), InStr(Rst.Fields("dhProt"), " "), Len(Rst.Fields("dhProt"))))
            End If
            txtxJust.Text = IIf(IsNull(Rst.Fields("canc_xJust")), "", Rst.Fields("canc_xJust"))
            txtProtocolo.Text = Rst.Fields("canc_nProt") & " - " & Rst.Fields("canc_Status")
    End If
    Rst.Close
End Sub


Private Sub tbMenu_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Index
        Case 1
            Imprimir 1
        Case 2
            Imprimir 2
    End Select

End Sub

Private Sub txtchNFe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
End Sub


Private Sub txtchNFe_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub



Private Sub txtDtAutorizacao_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txthrAutorizacao_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtnProt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
End Sub


Private Sub txtnProt_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub txtProtocolo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
End Sub

Private Sub txtProtocolo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub txtxJust_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
End Sub


