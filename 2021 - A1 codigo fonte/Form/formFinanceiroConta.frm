VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFinanceiroConta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financeiro - Cadastro de Conta"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6045
   Begin VB.Frame Frame1 
      Height          =   4995
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   5895
      Begin VB.Frame Frame2 
         Caption         =   "Cobrança:"
         Height          =   2115
         Left            =   180
         TabIndex        =   16
         Top             =   2760
         Width           =   5595
         Begin VB.TextBox txtTipo 
            Height          =   285
            Left            =   1440
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   1680
            Width           =   1395
         End
         Begin VB.TextBox txtConvenioLider 
            Height          =   285
            Left            =   1440
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   1380
            Width           =   1395
         End
         Begin VB.TextBox txtVariacao 
            Height          =   285
            Left            =   1440
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   780
            Width           =   1395
         End
         Begin VB.TextBox txtContrato 
            Height          =   285
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   180
            Width           =   1395
         End
         Begin VB.TextBox txtCarteira 
            Height          =   285
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   480
            Width           =   1395
         End
         Begin VB.TextBox txtConvenio 
            Height          =   285
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   1080
            Width           =   1395
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Convenio:"
            Height          =   195
            Left            =   600
            TabIndex        =   28
            Top             =   1740
            Width           =   795
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Convenio Lider:"
            Height          =   195
            Left            =   180
            TabIndex        =   26
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Variação:"
            Height          =   195
            Left            =   300
            TabIndex        =   24
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Contrato:"
            Height          =   195
            Left            =   660
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Carteira:"
            Height          =   195
            Left            =   600
            TabIndex        =   20
            Top             =   540
            Width           =   795
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Convenio:"
            Height          =   195
            Left            =   660
            TabIndex        =   19
            Top             =   1140
            Width           =   735
         End
      End
      Begin VB.TextBox txtContaDV 
         Height          =   315
         Left            =   3780
         MaxLength       =   1
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1140
         Width           =   315
      End
      Begin VB.TextBox txtAgenciaDV 
         Height          =   315
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   780
         Width           =   315
      End
      Begin VB.TextBox txtDiasProtesto 
         Height          =   285
         Left            =   1860
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2280
         Width           =   1035
      End
      Begin VB.TextBox txtJuros 
         Height          =   285
         Left            =   1860
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtMulta 
         Height          =   285
         Left            =   1860
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtConta 
         Height          =   285
         Left            =   1860
         MaxLength       =   20
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1140
         Width           =   1755
      End
      Begin VB.TextBox txtAgencia 
         Height          =   285
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   780
         Width           =   915
      End
      Begin VB.ComboBox cboBanco 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   3555
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Dias para protesto:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2340
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Juros por dia (%):"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Multa (%):"
         Height          =   255
         Left            =   780
         TabIndex        =   6
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Conta:"
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Agencia (s/DV):"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome do Banco:"
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   420
         Width           =   1635
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6045
      _ExtentX        =   10663
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
               Picture         =   "formFinanceiroConta.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroConta.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroConta.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroConta.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroConta.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroConta.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroConta.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroConta.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroConta.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroConta.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroConta.frx":54D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFinanceiroConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IdReg     As Integer
Dim strTabela   As String


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


Private Sub cboBanco_DropDown()
    Dim Rst As Recordset
    cboBanco.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroBancoCadastro WHERE id_empresa=" & ID_Empresa & " ORDER BY Nome")
    If Rst.BOF And Rst.EOF Then
            
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboBanco.AddItem Left(String(3, "0"), 3 - Len(Trim(Rst.Fields("id")))) & Trim(Rst.Fields("id")) & " - " & Rst.Fields("Nome")
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
    LimpaFormulario Me
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
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
                        "Descrição.: " & cboBanco.Text, vbYesNo + vbQuestion) = vbYes Then
                               
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
                'LimpaFormulario me
                'txtCNPJ.Enabled = True
            End If
            
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            
            
        Case "Manutenção da Tabela"
            MontarBaseDados
    End Select
End Sub
Private Sub MontarBaseDados()
    'formManutencaoTabelas.IniciarManutencao Me
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    cReg = 0
    vReg(cReg) = Array("Banco", 10, "N"): cReg = cReg + 1
    vReg(cReg) = Array("Agencia", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("AgenciaDV", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Conta", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ContaDV", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Multa", 50, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Juros", 50, "S"): cReg = cReg + 1
    vReg(cReg) = Array("DiasProtesto", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Saldo", 30, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("Contrato", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Carteira", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Variacao", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Convenio", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ConvenioLider", 10, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Tipo", 10, "S") ': cReg = cReg + 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vReg, cReg
    
    cReg = 0
    vReg(cReg) = Array("IDConta", 10, "N"): cReg = cReg + 1
    vReg(cReg) = Array("IdRegDoc", 10, "N"): cReg = cReg + 1
    vReg(cReg) = Array("Data", 10, "D"): cReg = cReg + 1
    vReg(cReg) = Array("Documento", 30, "S"): cReg = cReg + 1
    vReg(cReg) = Array("tpDoc", 10, "N"): cReg = cReg + 1
    vReg(cReg) = Array("Descricao", 250, "S"): cReg = cReg + 1
    vReg(cReg) = Array("CD", 1, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Valor", 30, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Saldo", 30, "S") ': cReg = cReg + 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vReg, cReg, "Historico"
    
    
End Sub
Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim i           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer 'Contador de Registros
    cReg = 0
    For i = 0 To Me.Controls.Count - 1
        Set Controle = Me.Controls(i)
        
        If TypeOf Controle Is TextBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Text, "S")
            cReg = cReg + 1
        End If
        If TypeOf Controle Is ComboBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Trim(Left(Controle.Text, 3)), "N")
            cReg = cReg + 1
        End If
        If TypeOf Controle Is CheckBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Value, "S")
            cReg = cReg + 1
        End If
    Next
    vReg(cReg) = Array("Saldo", ChkVal("0.00", 0, cDecMoeda), "S")
    'cReg = cReg - 1
     
    If IdReg = 0 Then
            If RegistroIncluir(strTabela, vReg, cReg) = 0 Then
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
    Dim sSQL    As String
    Dim sTMP    As String
    sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg

    ExibirDados Me, sSQL
    
    With cboBanco
        sTMP = Trim(.Text)
        .Clear
        If sTMP <> "" Then
            .AddItem Left(String(3, "0"), 3 - Len(sTMP)) & sTMP & " - " & pgDadosBanco(CInt(sTMP)).Nome
            .Text = .List(0)
        End If
    End With

End Sub

Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub


Private Sub txtConta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub


Private Sub txtDiasProtesto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub


Private Sub txtJuros_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtJuros.Text, KeyAscii, 3)
End Sub


Private Sub txtMulta_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtMulta.Text, KeyAscii, 3)
End Sub


