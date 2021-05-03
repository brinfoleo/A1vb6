VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formFaturamentoNFeInutilizarNumero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Inutilizar Número de NF-e"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   8100
   Begin VB.Frame Frame1 
      Height          =   3315
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   7935
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   315
         Left            =   1260
         TabIndex        =   16
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58523649
         CurrentDate     =   40816
      End
      Begin VB.Frame Frame2 
         Height          =   1995
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   4815
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo NF:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblTpNF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1020
            TabIndex        =   22
            Top             =   210
            Width           =   3495
         End
         Begin VB.Label lblSerie 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1020
            TabIndex        =   21
            Top             =   900
            Width           =   3495
         End
         Begin VB.Label lblModelo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1020
            TabIndex        =   20
            Top             =   1230
            Width           =   3495
         End
         Begin VB.Label lblAno 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1020
            TabIndex        =   19
            Top             =   1590
            Width           =   3495
         End
         Begin VB.Label lblCodUF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1020
            TabIndex        =   18
            Top             =   540
            Width           =   3495
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Ano:"
            Height          =   195
            Left            =   420
            TabIndex        =   17
            Top             =   1620
            Width           =   495
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Modelo:"
            Height          =   195
            Left            =   300
            TabIndex        =   14
            Top             =   1260
            Width           =   615
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Série:"
            Height          =   195
            Left            =   360
            TabIndex        =   13
            Top             =   930
            Width           =   555
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Cod. UF:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   570
            Width           =   795
         End
      End
      Begin VB.TextBox txtJust 
         Height          =   795
         Left            =   1260
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "formFaturamentoNFeInutilizarNumero.frx":0000
         Top             =   2340
         Width           =   6435
      End
      Begin VB.TextBox txtnFin 
         Height          =   285
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1560
         Width           =   1395
      End
      Begin VB.TextBox txtnIni 
         Height          =   285
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1200
         Width           =   1395
      End
      Begin VB.CommandButton btoPesqTipoNF 
         Height          =   315
         Left            =   2400
         Picture         =   "formFaturamentoNFeInutilizarNumero.frx":0006
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtTpNF 
         Height          =   285
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Data:"
         Height          =   195
         Left            =   540
         TabIndex        =   15
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Justificativa:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2340
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero Final:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo NF:"
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   420
         Width           =   915
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
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
               Picture         =   "formFaturamentoNFeInutilizarNumero.frx":0390
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeInutilizarNumero.frx":07E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeInutilizarNumero.frx":0AFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeInutilizarNumero.frx":138E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeInutilizarNumero.frx":25E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeInutilizarNumero.frx":2EBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeInutilizarNumero.frx":374C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeInutilizarNumero.frx":3FDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeInutilizarNumero.frx":5230
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeInutilizarNumero.frx":554A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeInutilizarNumero.frx":5864
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFaturamentoNFeInutilizarNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Function grvRegistro() As Boolean
    Dim i           As Integer
    Dim chv         As String
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    Dim ide_verProc As String
    Dim IdReg       As Integer
    
    If ValidarDados = False Then Exit Function
    If MsgBox("Após essa solicitação não havera possibilidade de utilizar esse(s) numero(s)!" & vbCrLf & vbCrLf & _
              "Deseja continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
              grvRegistro = False
              Exit Function
    End If
    
    
    chv = Trim(Left(lblCodUF.Caption, 2)) & _
        Trim(lblAno.Caption) & _
        Trim(PgDadosEmpresa(ID_Empresa).CNPJ) & _
        ZE(lblModelo.Caption, 2) & _
        ZE(lblSerie.Caption, 3) & _
        ZE(txtnIni.Text, 9) & _
        ZE(txtnFin.Text, 9)
    ide_verProc = sVersao & "." & cVersao
    
    For i = txtnIni.Text To txtnFin.Text
        cReg = 0

        vReg(cReg) = Array("IdNFe", chv, "S"): cReg = cReg + 1
        vReg(cReg) = Array("Versao", VersaoNFe, "S"): cReg = cReg + 1
        vReg(cReg) = Array("ide_cUF", Left(lblCodUF.Caption, 2), "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_cNF", ide_cNF, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_natOp", ide_natOp, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_indPag", ide_indPag, "S"): cReg = cReg + 1
        vReg(cReg) = Array("ide_mod", lblModelo.Caption, "S"): cReg = cReg + 1
        vReg(cReg) = Array("ide_Serie", lblSerie.Caption, "S"): cReg = cReg + 1
        vReg(cReg) = Array("ide_nNF", ZE(i, 9), "S"): cReg = cReg + 1
        vReg(cReg) = Array("ide_dEmi", dtpData.Value, "D"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_dSaiEnt", ide_dSaiEnt, "D"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_hSaiEnt", ide_hSaiEnt, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_tpNf", ide_tpNF, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_cMunFG", ide_cMunFG, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_refNFe", ide_refNFe, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_tpImp", ide_tpImp, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_tpEmis", ide_tpEmis, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_cDV", ide_cDV, "S"): cReg = cReg + 1
        vReg(cReg) = Array("ide_tpAmb", PgDadosConfig.Ambiente, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_finNFe", ide_finNFe, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ide_procEmi", ide_procEmi, "S"): cReg = cReg + 1
        vReg(cReg) = Array("ide_verProc", ide_verProc, "S"): cReg = cReg + 1
      'TOTAIS
        vReg(cReg) = Array("total_vBC", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("total_vICMS", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("total_vBCST", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("total_vICMSST", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
    
        vReg(cReg) = Array("total_vProd", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("total_vFrete", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("total_vSeg", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("total_vDesc", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("total_vOutro", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("total_vIPI", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("total_vPIS", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("total_vCOFINS", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
        vReg(cReg) = Array("total_vNF", ChkVal("0", 0, cDecMoeda), "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ger_Vendedor", ger_Vendedor, "N"): cReg = cReg + 1
    'vReg(cReg) = Array("ger_idPV", ger_idPV, "N") ': cReg = cReg + 1
    
    'Numero de NFe Inutilizado
    'vDados(contReg) = Array("inut_nProt", "250", "S"): contReg = contReg + 1
    'vDados(contReg) = Array("inut_dhRecbto", "250", "S"): contReg = contReg + 1
        vReg(cReg) = Array("inut_xJust", Trim(UCase(txtJust.Text)), "S"): cReg = cReg + 1
    'vDados(contReg) = Array("inut_Status", "250", "S"): contReg = contReg + 1
    'DADOS DESTINATARIO
        vReg(cReg) = Array("dest_xNome", "INUTILIZADO", "S"): cReg = cReg + 1
        vReg(cReg) = Array("emit_CNPJ", PgDadosEmpresa(ID_Empresa).CNPJ, "S"): cReg = cReg + 1
        cReg = cReg - 1
        IdReg = RegistroIncluir("FaturamentoNFe", vReg, cReg)
        If IdReg = 0 Then
                MsgBox "Erro ao gravar solicitação de inutilização!", vbInformation, "Aviso"
                grvRegistro = False
            Else
                grvRegistro = True
        End If
    
        '---------------------------------------------------------------------------------------------------------
    Next
    If grvRegistro = True Then
        Inutilizar_NFe (chv)
    End If
End Function

Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Unload Me
    End If

    LimpForm
    HDForm Me, True
    HDMenu Me, False
End Sub

Private Sub LimpForm()
    LimpaFormulario Me
    lblTpNF.Caption = ""
    lblAno.Caption = ""
    lblSerie.Caption = ""
    lblModelo.Caption = ""
    lblCodUF.Caption = ""
    dtpData.Value = Date
End Sub

Private Sub dtpData_Click()
    lblAno.Caption = Format(dtpData.Value, "YY")
End Sub


Private Sub Form_Load()
    LimpForm
    HDForm Me, False
    HDMenu Me, True
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
            Incluir
            
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDForm Me, False
            '    RegistroExcluir "financeirocontasprcadastro", "ide_NFe = '" & txtchNFe.Text & "'"
            End If
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpForm
    End Select
End Sub


Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If

End Sub

Private Sub txtnIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub
Private Sub txtnfin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtTpNF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarTpNF
    End If
End Sub
Private Sub txtTpNF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        PesquisarTpNF (txtTpNF.Text)
    End If
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub
Private Sub btoPesqTipoNF_Click()
    PesquisarTpNF
End Sub

Private Sub PesquisarTpNF(Optional idNF As Integer)
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    If idNF = 0 Then
        idNF = formBuscar.IniciarBusca("FaturamentoTipoNotaFiscal")
        If idNF = 0 Then
            Exit Sub
        End If
    End If
    
    sSQL = "SELECT * FROM FaturamentoTipoNotaFiscal WHERE ID_Empresa = " & ID_Empresa & " AND ID = " & idNF
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'LimpForm
            'DesMontarVariaveisFiscal
            'idPedido = 0
            'idTpNF = 0
            'HDMenu Me, True
            'HDForm Me, False
            'txtTpNF.Enabled = True
            'btoPesqTipoNF.Enabled = True
            'tbMenu.Buttons(3).Enabled = False
            LimpForm
            MsgBox "Nenhum Registro encontrado."
        Else
            'HDForm Me, True
            Rst.MoveFirst
            
            txtTpNF.Text = Rst.Fields("ID")
            lblTpNF.Caption = Rst.Fields("Descricao")
            lblModelo.Caption = Rst.Fields("Modelo")
            lblSerie.Caption = Rst.Fields("Serie")
            lblAno.Caption = Format(dtpData.Value, "YY")
            
            lblCodUF.Caption = pgDadosICMS(PgDadosEmpresa(ID_Empresa).UF, 0).codUF & " - " & PgDadosEmpresa(ID_Empresa).UF
    End If
    Rst.Close
    
    
End Sub
Private Function ValidarDados() As Boolean
    Dim Rst     As Recordset
    Dim sSQL    As String
    ValidarDados = False
    If Trim(lblTpNF.Caption) = "" Or Trim(lblCodUF.Caption) = "" Or _
       Trim(lblSerie.Caption) = "" Or Trim(lblModelo.Caption) = "" Or _
       Trim(lblAno.Caption) = "" Then
       MsgBox "Selecione o Tipo de Nota Fiscal!", vbInformation, "Aviso"
    End If
    
    If txtnIni.Text > txtnFin.Text Then
        MsgBox "Numero inicial não pode ser maior que o numero final!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    If Len(Trim(txtJust.Text)) < 15 Then
        MsgBox "A justificativa deve ter no minimo 15 caracteres!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ide_nNF >= " & txtnIni.Text & " AND ide_nNF <= " & txtnFin.Text
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            MsgBox "Impossível inutilizar o(s) número(s) devido já ter (em) sido usado(s) para emissão de documento(s) fiscal (is)!", vbInformation, "Aviso"
            ValidarDados = False
            Exit Function
    End If
    Rst.Close
     
    ValidarDados = True
End Function
