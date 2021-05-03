VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formFinanceiroCondicoesPagamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Condições de Pagamento"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6840
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   6735
      Begin VB.Frame Frame2 
         Height          =   2115
         Left            =   180
         TabIndex        =   6
         Top             =   900
         Width           =   6435
         Begin VB.CommandButton btoAdicionar 
            Caption         =   "Adicionar >>"
            Height          =   375
            Left            =   1020
            TabIndex        =   14
            Top             =   1620
            Width           =   1275
         End
         Begin VB.TextBox txtPercentual 
            Height          =   285
            Left            =   1260
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   1140
            Width           =   855
         End
         Begin VB.TextBox txtDC 
            Height          =   285
            Left            =   1260
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   720
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid msfgCondicoes 
            Height          =   1875
            Left            =   2460
            TabIndex        =   13
            Top             =   180
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3307
            _Version        =   393216
            Cols            =   3
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   "^Parcela  |^Dias Corridos   |^Percentual "
         End
         Begin VB.Label lblParcela 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000"
            Height          =   315
            Left            =   1260
            TabIndex        =   12
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Parcela:"
            Height          =   195
            Left            =   180
            TabIndex        =   11
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Percentual:"
            Height          =   195
            Left            =   180
            TabIndex        =   9
            Top             =   1140
            Width           =   915
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Dias Corridos:"
            Height          =   195
            Left            =   60
            TabIndex        =   8
            Top             =   780
            Width           =   1035
         End
      End
      Begin VB.TextBox txtParcelas 
         Height          =   285
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "2"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   4995
      End
      Begin VB.Label Label3 
         Caption         =   "Num. de Parcelas"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   780
         TabIndex        =   2
         Top             =   300
         Width           =   795
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6840
      _ExtentX        =   12065
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
               Picture         =   "formFinanceiroCondicoesPagamento.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroCondicoesPagamento.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroCondicoesPagamento.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroCondicoesPagamento.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroCondicoesPagamento.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroCondicoesPagamento.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroCondicoesPagamento.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroCondicoesPagamento.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroCondicoesPagamento.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroCondicoesPagamento.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroCondicoesPagamento.frx":54D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFinanceiroCondicoesPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg     As Integer
Dim strTabela   As String


Private Function grvRegistroGrade() As Boolean
    Dim vReg(199)    As Variant
    Dim i           As Integer
    Dim cReg        As Integer 'Contador de Registros
    
    cReg = 0
    RegistroExcluir "financeirocondicoespagamentoparcelas", "IDCondicoes=" & IdReg
    For i = 1 To msfgCondicoes.Rows - 1
            vReg(cReg) = Array("IDCondicoes", IdReg, "S"): cReg = cReg + 1
            vReg(cReg) = Array("Parcela", msfgCondicoes.TextMatrix(i, 0), "S"): cReg = cReg + 1
            vReg(cReg) = Array("DiasCorridos", msfgCondicoes.TextMatrix(i, 1), "S"): cReg = cReg + 1
            vReg(cReg) = Array("Percentual", msfgCondicoes.TextMatrix(i, 2), "S") ':cReg = cReg + 1
            
        If RegistroIncluir("financeirocondicoespagamentoparcelas", vReg, cReg) = 0 Then
                MsgBox "Erro ao Incluir"
                grvRegistroGrade = False
            Else
                grvRegistroGrade = True
                cReg = 0
        End If
    Next
End Function

Private Sub PesquisarRegistro()
    ''Dim idreg  As String
    IdReg = formBuscar.IniciarBusca(strTabela)
    ''IdReg = IIf(idreg = "", 0, idreg)
    
    If IdReg = 0 Then
            LimpaFormulario Me 'me
        Else
            MostrarDados
    End If
End Sub




Private Sub btoAdicionar_Click()
    If msfgCondicoes.Rows = 1 Then Exit Sub
    If Val(lblParcela.Caption) <= 0 Then
        MsgBox "Favor selecionar uma condicao de pagamento!", vbInformation, App.EXEName
        Exit Sub
    End If
    msfgCondicoes.TextMatrix(msfgCondicoes.Row, 0) = lblParcela.Caption
    msfgCondicoes.TextMatrix(msfgCondicoes.Row, 1) = txtDC.Text
    msfgCondicoes.TextMatrix(msfgCondicoes.Row, 2) = txtPercentual.Text
End Sub



Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    HDForm Me, False
    HDMenu Me, True
    
   
End Sub

Private Sub msfgCondicoes_Click()
    If msfgCondicoes.Row = 0 Then Exit Sub
    lblParcela.Caption = msfgCondicoes.TextMatrix(msfgCondicoes.Row, 0)
    txtDC.Text = msfgCondicoes.TextMatrix(msfgCondicoes.Row, 1)
    txtPercentual.Text = msfgCondicoes.TextMatrix(msfgCondicoes.Row, 2)
End Sub
Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    IdReg = 0
    HDMenu Me, False
    HDForm Me, True
    LimpaFormulario Me
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Selecione uma Grupo"
        Exit Sub
    End If
    HDForm Me, True
    HDMenu Me, False
End Sub
Private Sub Excluir()
    If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
            MsgBox "Selecione um Registro"
            Exit Sub
        Else
            If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                        vbCrLf & _
                        "Descrição.: " & txtDescricao.Text, vbYesNo + vbQuestion) = vbYes Then
                               
                If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                    LimpaFormulario Me
                End If
            End If
    End If
End Sub
Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
            Incluir
        Case "Alterar"
            Alterar
        Case "Excluir"
            Excluir
        Case "Pesquisar"
            PesquisarRegistro
            
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDForm Me, False
                If grvRegistroGrade = True Then
                    HDMenu Me, True
                    HDForm Me, False
                End If
            End If
 
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            
        Case "Manutenção da Tabela"
            ManutencaoTabela
    End Select
End Sub
Private Sub ManutencaoTabela()

    Dim vReg(1000)  As Variant
    Dim cReg        As Integer 'Cota os registros
    
    
    formManutencaoTabelas.IniciarManutencao Me
            'MsgBox "O sistema vai criar uma base de dadis simples para registro das parcelas. Favor ver o codigo fonte"
            'BD.Execute "DROP TABLE financeirocondicoespagamentoparcelas"
            'BD.Execute "CREATE TABLE IF NOT EXISTS financeirocondicoespagamentoparcelas " & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _

               '"IDCondicoes VARCHAR(20) default Null," & _
               "Parcela VARCHAR(20) default Null," & _
               "DiasCorridos VARCHAR(20) default Null," & _
               "Percentual VARCHAR(20) default Null," & _
               "PRIMARY KEY (Id))"
    
    
    
    
    'Dim sSQL As String
    cReg = 0
    vReg(cReg) = Array("IDCondicoes", 20, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Parcelas", 20, "S"): cReg = cReg + 1
    vReg(cReg) = Array("DiasCorridos", 20, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Percentual", 20, "S") ': cReg = cReg + 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vReg, cReg, "parcelas"
    
End Sub
Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim i           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer 'Contador de Registros
    cReg = 0
    If ValidarDados = False Then Exit Function
    
    For i = 0 To Me.Controls.Count - 1
        Set Controle = Me.Controls(i)
        
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
    
     cReg = cReg - 1
    If IdReg = 0 Then
            IdReg = RegistroIncluir(strTabela, vReg, cReg)
            If IdReg = 0 Then
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
Private Function ValidarDados() As Boolean
    Dim sP  As String
    Dim i   As Integer
    ValidarDados = False
    If Trim(txtDescricao.Text) = "" Then
        MsgBox "Favor informar a DESCRIÇÃO das condicoes de Pagamento", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    sP = 0
    For i = 1 To msfgCondicoes.Rows - 1
        sP = Val(ChkVal(sP, 0, 3)) + Val(ChkVal(msfgCondicoes.TextMatrix(i, 2), 0, 3))
    Next
    If Trim(sP) <> 100 Then
        MsgBox "Somatorio do percentual de pagamentos é de " & sP & " não sendo igual a 100%. Favor Verificar", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    
    ValidarDados = True
End Function



Private Sub txtDC_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then btoAdicionar_Click
    KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub txtParcelas_Change()
    Dim i As Integer
    If Trim(txtParcelas.Text) = "" Then Exit Sub
    msfgCondicoes.Rows = IIf(Trim(txtParcelas.Text) = "", 2, Trim(txtParcelas.Text) + 1)
    For i = 1 To msfgCondicoes.Rows - 1
        msfgCondicoes.TextMatrix(i, 0) = Left("000", 3 - Len(Trim(i))) & Trim(i)
        msfgCondicoes.TextMatrix(i, 2) = ChkVal(100 / (msfgCondicoes.Rows - 1), 0, 3)
    Next
End Sub
Private Sub txtPercentual_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtPercentual.Text, KeyAscii, 3)
End Sub
Private Sub MostrarDados()
    Dim sSQL As String
    Dim Rst As Recordset
    
    sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg

    ExibirDados Me, sSQL
    
    
    txtDC.Text = ""
    txtPercentual.Text = ""
    'Carrega grade
    
    sSQL = "SELECT * FROM financeirocondicoespagamentoparcelas WHERE ID_Empresa = " & ID_Empresa & " AND IDCondicoes = " & IdReg
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            msfgCondicoes.Rows = 1
        Else
            msfgCondicoes.Rows = 1
            Rst.MoveFirst
            Do Until Rst.EOF
                With msfgCondicoes
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = IIf(IsNull(Rst.fields("Parcela")), 0, Rst.fields("Parcela"))
                    .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rst.fields("DiasCorridos")), "0", Rst.fields("DiasCorridos"))
                    .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rst.fields("Percentual")), "0", Rst.fields("Percentual"))
                    Rst.MoveNext
                End With
            Loop
    End If
    
    
End Sub







