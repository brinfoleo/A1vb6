VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFaturamentoNFeCartaCorrecao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Carta de Correção"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   10725
   Begin VB.Frame Frame2 
      Caption         =   "Correção"
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
      Left            =   60
      TabIndex        =   12
      Top             =   3360
      Width           =   10455
      Begin VB.TextBox txtCorrecao 
         Height          =   1935
         Left            =   120
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "formFaturamentoNFeCartaCorrecao.frx":0000
         Top             =   240
         Width           =   10155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da Carta de Correção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   10455
      Begin VB.TextBox txtDoc 
         Height          =   285
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   2160
         MaxLength       =   120
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox txtnNF 
         Height          =   285
         Left            =   2100
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtnProt 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   960
         Width           =   4755
      End
      Begin VB.TextBox txtChvNFe 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   600
         Width           =   8115
      End
      Begin VB.TextBox txtRegID 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   1155
      End
      Begin MSComCtl2.DTPicker dtpdEmi 
         Height          =   315
         Left            =   2100
         TabIndex        =   9
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   99549185
         CurrentDate     =   40668
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ/CPF:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1875
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   195
         Left            =   1140
         TabIndex        =   13
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Num. Nota Fiscal:"
         Height          =   195
         Left            =   540
         TabIndex        =   10
         Top             =   2460
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Protocolo de Autorização:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Data:"
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   2100
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Chave de Acesso:"
         Height          =   195
         Left            =   420
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número:"
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   675
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10725
      _ExtentX        =   18918
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
               Picture         =   "formFaturamentoNFeCartaCorrecao.frx":0006
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCartaCorrecao.frx":0458
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCartaCorrecao.frx":0772
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCartaCorrecao.frx":1004
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCartaCorrecao.frx":2256
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCartaCorrecao.frx":2B30
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCartaCorrecao.frx":33C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCartaCorrecao.frx":3C54
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCartaCorrecao.frx":4EA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCartaCorrecao.frx":51C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeCartaCorrecao.frx":54DA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFaturamentoNFeCartaCorrecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg       As Integer
Dim idNF        As Integer
Dim CampoID     As Integer
Dim strTabela   As String

'##################################################################################
'### Funcao para remover todos os itens da tab. FaturamentoNFeCartaCorrecaoItens
'### para a tab. FaturamentoNFeCartaCorrecao
'##################################################################################
'Private Sub Command1_Click()
'    Dim Rst1    As Recordset
'    Dim sSQL1   As String
'    Dim Rst2    As Recordset
'    Dim sSQL2   As String
'    Dim X       As String
'    Dim vReg(1) As Variant
'
'    sSQL1 = "SELECT * FROM FaturamentoNFeCartaCorrecao ORDER BY Id"
'    Set Rst1 = RegistroBuscar(sSQL1)
'    If Rst1.BOF And Rst1.EOF Then
'        Else
'            Rst1.MoveFirst
'            Do Until Rst1.EOF
'                sSQL2 = "SELECT * FROM FaturamentoNFeCartaCorrecaoItens WHERE idReg=" & Rst1.Fields("id") & " ORDER BY Id"
'                Set Rst2 = RegistroBuscar(sSQL2)
'                If Rst2.BOF And Rst2.EOF Then
'                        Rst1.MoveNext
'                    Else
'                        Rst2.MoveFirst
'                        X = ""
'                        Do Until Rst2.EOF
'                            X = X & Rst2.Fields("cIrregular") & ": " & Rst2.Fields("Correcao") & "; "
'                            Rst2.MoveNext
'                        Loop
'                        Debug.Print X
'                        vReg(0) = Array("Correcao", X, "S")
'                        RegistroAlterar "FaturamentoNFeCartaCorrecao", vReg, 0, "id=" & Rst1.Fields("id")
'                        Rst1.MoveNext
'                End If
'            Loop
'    End If
'    Rst1.Close
'End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then Unload Me
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    IdReg = 0
    HDForm Me, False
    HDMenu Me, True
    txtRegID.Enabled = True
End Sub


Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
            Incluir
        Case "Alterar"
            Alterar
        Case "Excluir"
            Excluir
        Case "Imprimir"
            ImprimirDocumento
        Case "Pesquisar"
            IdReg = 0
            PesquisarRegistro
            
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDForm Me, False
                txtRegID.Enabled = False
                ExportarCCe
                'LimpaFormulario me
            End If
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            txtRegID.Enabled = True
        
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select
End Sub
Private Sub ExportarCCe()
    On Error Resume Next
    Dim tmp As String
    If Trim(txtChvNFe.Text) = "" Then Exit Sub
    tmp = Exportar_CCe_v200_TXT(txtChvNFe.Text)
    
    RegLog "CCe", "0", "CCe: " & tmp
End Sub
Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then Exit Sub
    LimpaFormulario Me
    IdReg = 0
    HDForm Me, True
    HDMenu Me, False
    txtRegID.Enabled = False
    'msfgCC.Rows = 1
End Sub
Private Sub ImprimirDocumento()
    Dim Rst     As Recordset
    Dim sSQL    As String
    'Dim Rst1    As Recordset
    'Dim sSQL1   As String
    
    If chkAcesso(Me, "i") = False Then Exit Sub
    
    If IdReg = 0 Then
        MsgBox "Selecione um Registro."
        Exit Sub
    End If
    
    sSQL = "SELECT * FROM FaturamentoNFeCartaCorrecao WHERE id =" & IdReg
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar Carta de Correção.", vbInformation, "Aviso"
            Exit Sub
        Else
            Rst.MoveFirst
            'sSQL1 = "SELECT * FROM FaturamentoNFeCartaCorrecaoItens WHERE IDReg =" & IdReg
            'Set Rst1 = RegistroBuscar(sSQL1)
            'If Rst1.BOF And Rst1.EOF Then
            '    Else
            '        Rst1.MoveFirst
            'End If
    End If
    
    Set rptNFeCC.DataSource = Rst.DataSource
    
    With rptNFeCC.Sections("tit1").Controls
        .Item("lblData").Caption = UCase(PgDadosEmpresa(ID_Empresa).Mun & ", " & Mid(Format(Date, "Long date"), InStr(Format(Date, "Long date"), ",") + 1, Len(Format(Date, "Long date"))))
        .Item("lblTitulo").Caption = UCase("CARTA DE CORREÇÃO EM DOCUMENTOS FISCAIS N°." & Left(String(5, "0"), 5 - Len(Trim(IdReg))) & IdReg)
        .Item("lblNome").Caption = Rst.Fields("Nome")
        If Len(Rst.Fields("Doc")) > 11 Then
                .Item("lblCNPJ").Caption = Replace(Format(Rst.Fields("Doc"), "00 000 000/0000-00"), " ", ".")
            Else
                .Item("lblCNPJ").Caption = Replace(Format(Rst.Fields("Doc"), "000 000 000-00"), " ", ".")
        End If
        
        .Item("lblChvNFe").Caption = Rst.Fields("ChvNFe")
        .Item("lblProtocolo").Caption = Rst.Fields("nProt")
        .Item("lblEmissao").Caption = Rst.Fields("dEmi")
        .Item("lblNF").Caption = Rst.Fields("nNF")
    End With
    rptNFeCC.Sections("Section3").Controls.Item("lblEmpresa").Caption = UCase(PgDadosEmpresa(ID_Empresa).Nome)
    rptNFeCC.Sections("Section3").Controls.Item("lblEmpresaDoc").Caption = Replace(Format(PgDadosEmpresa(ID_Empresa).CNPJ, "## ### ###/####-##"), " ", ".")
    
    rptNFeCC.Show 1
    Rst.Close
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then Exit Sub
    
    If IdReg = 0 Then
        MsgBox "Selecione um Registro."
        Exit Sub
    End If
    HDForm Me, True
    HDMenu Me, False
    txtRegID.Enabled = False
End Sub
Private Sub Excluir()
    If chkAcesso(Me, "e") = False Then Exit Sub
    
     If IdReg = 0 Then
            MsgBox "Selecione um Registro"
            Exit Sub
        Else
            If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                        vbCrLf & _
                        "Carta de Correção: " & txtRegID.Text & vbCrLf & _
                        "Nota Fiscal: " & txtnNF.Text & vbCrLf & _
                        "Nome: " & txtNome.Text, vbYesNo + vbQuestion) = vbYes Then
                               
                If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                    RegistroExcluir strTabela & "Itens", "IdReg = " & IdReg
                    LimpaFormulario Me
                    IdReg = 0
                    End If
                End If
    End If
End Sub
Private Sub MontarBaseDeDados()
    Dim vDados(100) As Variant
    Dim cReg        As Integer
    cReg = 0
    vDados(cReg) = Array("Nome", "120", "S"): cReg = cReg + 1
    vDados(cReg) = Array("Doc", "20", "S"): cReg = cReg + 1
    vDados(cReg) = Array("ChvNFe", "50", "S"): cReg = cReg + 1
    vDados(cReg) = Array("nProt", "50", "S"): cReg = cReg + 1
    vDados(cReg) = Array("dEmi", "20", "D"): cReg = cReg + 1
    vDados(cReg) = Array("nNF", "15", "S"): cReg = cReg + 1
    vDados(cReg) = Array("Correcao", "1000", "S") ': cReg = cReg + 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, cReg
    
End Sub
Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim i           As Integer
    Dim cReg        As Integer 'Contador de Registros
    cReg = 0
    If Len(Trim(txtCorrecao.Text)) < 15 Then
        MsgBox "Tamanho minimo de 15 caracteres.", vbInformation, "Aviso"
        grvRegistro = False
        Exit Function
    End If
    If Len(Trim(txtChvNFe.Text)) <= 0 Then
        MsgBox "Fvor informar chave de acesso da NFe.", vbInformation, "Aviso"
        grvRegistro = False
        Exit Function
    End If
    vReg(cReg) = Array("Nome", txtNome.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Doc", txtDoc.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ChvNFe", txtChvNFe.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("nProt", txtnProt.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("dEmi", dtpdEmi.Value, "D"): cReg = cReg + 1
    vReg(cReg) = Array("nNF", txtnNF.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Correcao", txtCorrecao.Text, "S") ': cReg = cReg + 1
    If IdReg = 0 Then
            IdReg = RegistroIncluir(strTabela, vReg, cReg)
            If IdReg = 0 Then
                    MsgBox "Erro ao Incluir"
                    grvRegistro = False
                Else
                    grvRegistro = True
                    txtRegID.Text = IdReg
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

Private Sub PesquisarRegistro()
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    If IdReg = 0 Then
        IdReg = formBuscar.IniciarBusca(strTabela)
        If IdReg = 0 Then Exit Sub
    End If
    sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar o registro n." & IdReg, vbInformation, "Aviso"
            LimpaFormulario Me
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
    End If
    
    txtRegID.Text = IdReg
    txtNome.Text = IIf(IsNull(Rst.Fields("Nome")), "", Rst.Fields("Nome"))
    txtDoc.Text = IIf(IsNull(Rst.Fields("Doc")), "", Rst.Fields("Doc"))
    
    txtChvNFe.Text = IIf(IsNull(Rst.Fields("ChvNfe")), "", Rst.Fields("ChvNfe"))
    txtnProt.Text = IIf(IsNull(Rst.Fields("nProt")), "", Rst.Fields("nProt"))
    dtpdEmi.Value = IIf(IsNull(Rst.Fields("dEmi")), Date, Rst.Fields("dEmi"))
    txtnNF.Text = IIf(IsNull(Rst.Fields("nNF")), "", Rst.Fields("nNF"))
    txtCorrecao.Text = IIf(IsNull(Rst.Fields("Correcao")), "", Rst.Fields("Correcao"))
    
    
End Sub

Private Sub txtChvNFe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        idNF = 0
        PesquisarNF
    End If
    
End Sub

Private Sub PesquisarNF()
    Dim sSQL    As String
    Dim Rst     As Recordset
    If Trim(IdReg) = 0 Then
        idNF = formBuscar.IniciarBusca("FaturamentoNFe")
    End If
    
    If idNF = 0 Then
            LimpaFormulario Me
            Exit Sub
    End If

    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & idNF
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Registro nao encontrado."
            LimpaFormulario Me
        Else
            Rst.MoveFirst
            txtChvNFe.Text = Rst.Fields("idNFe")
            dtpdEmi.Value = IIf(IsNull(Rst.Fields("ide_dEmi")), Date, Rst.Fields("ide_dEmi"))
            txtnProt.Text = IIf(IsNull(Rst.Fields("nProt")), "", Rst.Fields("nProt"))
            txtNome.Text = IIf(IsNull(Rst.Fields("dest_xNome")), "", Rst.Fields("dest_xNome"))
            txtDoc.Text = IIf(IsNull(Rst.Fields("dest_CNPJ")), "", Rst.Fields("dest_CNPJ"))
            txtnNF.Text = IIf(IsNull(Rst.Fields("ide_nNF")), "", Rst.Fields("ide_nNF"))
    End If
    Rst.Close

End Sub

Private Sub txtChvNFe_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub



Private Sub txtnNF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        idNF = 0
        PesquisarNF
    End If
End Sub

Private Sub txtnProt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        idNF = 0
        PesquisarNF
    End If
End Sub

Private Sub txtRegID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IdReg = 0
        PesquisarRegistro
    End If
End Sub

Private Sub txtRegID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtRegID.Text) = "" Then
                Exit Sub
            Else
                IdReg = txtRegID.Text
                PesquisarRegistro
        End If
    End If
    
End Sub


