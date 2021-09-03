VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formFinanceiroTipoDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financeiro - Tipo de Documento"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8220
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   8115
      Begin VB.ComboBox cboFormaPgto 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   4995
      End
      Begin VB.ComboBox cboImpressao 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   4995
      End
      Begin VB.TextBox txtSigla 
         Height          =   285
         Left            =   1620
         MaxLength       =   30
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtTipo 
         Height          =   285
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   600
         Width           =   4995
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Forma de Pagamento (NFe):"
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Impressão:"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sigla:"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de documento:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   780
         TabIndex        =   1
         Top             =   660
         Width           =   795
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8220
      _ExtentX        =   14499
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
               Picture         =   "formFinanceiroTipoDocumento.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroTipoDocumento.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroTipoDocumento.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroTipoDocumento.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroTipoDocumento.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroTipoDocumento.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroTipoDocumento.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroTipoDocumento.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroTipoDocumento.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroTipoDocumento.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroTipoDocumento.frx":54D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFinanceiroTipoDocumento"
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





Private Sub cboFormaPgto_DropDown()
'07.05.18 - Colocar um espaco a mais para que a gravacao capture apenas o numero
    With cboFormaPgto
        .Clear
        .AddItem "01  - Dinheiro"
        .AddItem "02  - Cheque"
        .AddItem "03  - Cartão de Crédito"
        .AddItem "04  - Cartão de Débito"
        .AddItem "05  - Crédito Loja"
        .AddItem "10  - Vale Alimentação"
        .AddItem "11  - Vale Refeição"
        .AddItem "12  - Vale Presente"
        .AddItem "13  - Vale Combustível"
        .AddItem "14  - Duplicata Mercantil"
        .AddItem "15  - Boleto Bancário"
        .AddItem "90  - Sem Pagamento"
        .AddItem "99  - Outros"
    End With
End Sub

Private Sub cboImpressao_DropDown()
    cboImpressao.Clear
    cboImpressao.AddItem "01 - Boleto Bancario"
    cboImpressao.AddItem "02 - Duplicata"
    cboImpressao.AddItem "03 - Recibo"
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
    txtTipo.Enabled = True
    
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
                        "Referencia: " & txtTipo.Text & vbCrLf & _
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
                'LimpaFormulario me
                'txtCNPJ.Enabled = True
            End If
            
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            txtTipo.Enabled = True
            
        Case "Manutenção da Tabela"
            formManutencaoTabelas.IniciarManutencao Me
    End Select
End Sub

Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim i           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer 'Contador de Registros
    
    
    If Trim(cboImpressao.Text) = "" Then
        MsgBox "Selecione o tipo de impressão.", vbInformation, "Aviso"
        grvRegistro = False
        Exit Function
    End If
    
     If Trim(txtSigla.Text) = "" Then
        MsgBox "Informe uma sigla.", vbInformation, "Aviso"
        grvRegistro = False
        Exit Function
    End If
    cReg = 0
    For i = 0 To Me.Controls.Count - 1
        Set Controle = Me.Controls(i)
        
        If TypeOf Controle Is TextBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Text, "S")
            cReg = cReg + 1
        End If
        If TypeOf Controle Is ComboBox Then
            vReg(cReg) = Array(Trim(Mid(Controle.Name, 4, Len(Controle.Name))), Controle.Text, "S")
            cReg = cReg + 1
        End If
        If TypeOf Controle Is CheckBox Then
            vReg(cReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), Controle.Value, "S")
            cReg = cReg + 1
        End If
    Next
    
     cReg = cReg - 1
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
Private Sub txtTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
End Sub

Private Sub MostrarDados()
    Dim sSQL As String
    sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg

    ExibirDados Me, sSQL


End Sub




Private Sub txtTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub
