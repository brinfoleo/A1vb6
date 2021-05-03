VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formTransportadoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transportadora"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8400
   Begin VB.Frame Frame3 
      Caption         =   "Observações"
      Height          =   1155
      Left            =   120
      TabIndex        =   27
      Top             =   4380
      Width           =   8175
      Begin VB.TextBox txtObs 
         Height          =   795
         Left            =   120
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "formTransportadoras.frx":0000
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   60
      TabIndex        =   15
      Top             =   2160
      Width           =   8235
      Begin VB.TextBox txtxEnder 
         Height          =   285
         Left            =   1080
         MaxLength       =   60
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   6735
      End
      Begin VB.TextBox txtBairro 
         Height          =   285
         Left            =   1080
         MaxLength       =   60
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   600
         Width           =   2955
      End
      Begin VB.ComboBox cboUF 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1020
         Width           =   915
      End
      Begin VB.ComboBox cboMun 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtCEP 
         Height          =   285
         Left            =   6300
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   900
         Width           =   1515
      End
      Begin VB.TextBox txtMail 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1440
         Width           =   3915
      End
      Begin VB.TextBox txtFone 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Endereço:"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Bairro:"
         Height          =   255
         Left            =   300
         TabIndex        =   21
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Municipio:"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "UF:"
         Height          =   195
         Left            =   540
         TabIndex        =   19
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "CEP:"
         Height          =   195
         Left            =   5760
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "e-mail:"
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1635
      Left            =   60
      TabIndex        =   11
      Top             =   480
      Width           =   8235
      Begin VB.ComboBox cboPessoa 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   180
         Width           =   1935
      End
      Begin VB.TextBox txtIE 
         Height          =   315
         Left            =   4680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   540
         Width           =   1935
      End
      Begin VB.TextBox txtCNPJ 
         Height          =   285
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   555
         Width           =   1935
      End
      Begin VB.TextBox txtxNome 
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   900
         Width           =   5295
      End
      Begin VB.TextBox txtFant 
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1260
         Width           =   5295
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Pessoa:"
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Insc.Estadual:"
         Height          =   255
         Left            =   3420
         TabIndex        =   23
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Razão Social:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "CPF/CNPJ:"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome Fantasia:"
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   1320
         Width           =   1155
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
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
               Picture         =   "formTransportadoras.frx":0006
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTransportadoras.frx":0458
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTransportadoras.frx":0772
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTransportadoras.frx":1004
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTransportadoras.frx":2256
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTransportadoras.frx":2B30
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTransportadoras.frx":33C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTransportadoras.frx":3C54
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTransportadoras.frx":4EA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTransportadoras.frx":51C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTransportadoras.frx":54DA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formTransportadoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IdReg    As Integer
Dim strTabela           As String


Private Sub PesquisarRegistro()
    ''Dim idreg  As String
    IdReg = formBuscar.IniciarBusca(strTabela) ', "xNome,xEnder,bairro,mun,uf,fone")
    ''IdReg = IIf(idreg = "", 0, idreg)
    
    If IdReg = 0 Then
            LimpaFormulario Me
        Else
            MostrarDados
    End If
End Sub



Private Sub cboMun_DropDown()
    Dim rst     As Recordset
    Dim sSQL    As String
    If Trim(cboUF.Text) = "" Then
        MsgBox "Selecione uma Unidade Federal (UF)."
        Exit Sub
    End If
    sSQL = "SELECT * FROM TributacaoMunicipio WHERE UF = '" & Trim(UCase(cboUF.Text)) & "' ORDER BY Descricao"
    cboMun.Clear
    Set rst = RegistroBuscar(sSQL)
    If rst.BOF And rst.EOF Then
        Else
            rst.MoveFirst
            Do Until rst.EOF
                cboMun.AddItem UCase(rst.Fields("Descricao"))
                rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboPessoa_DropDown()
    With cboPessoa
        .Clear
        .AddItem UCase("Fisica")
        .AddItem UCase("Juridica")
    End With
End Sub

Private Sub cboUF_DropDown()
    Dim rst As Recordset
    cboUF.Clear
    Set rst = RegistroBuscar("SELECT * FROM TributacaoUF ORDER BY sigla")
    If rst.BOF And rst.EOF Then
        Else
            rst.MoveFirst
            Do Until rst.EOF
                cboUF.AddItem rst.Fields("sigla")
                rst.MoveNext
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
    txtCNPJ.Enabled = True
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
        MsgBox "Selecione uma Transportadora"
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
            MsgBox "Selecione uma Transportadora"
            Exit Sub
        Else
            If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                        vbCrLf & _
                        "CNPJ: " & txtCNPJ.Text & vbCrLf & _
                        "Nome: " & txtxNome.Text, vbYesNo + vbQuestion) = vbYes Then
                               
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
                txtCNPJ.Enabled = True
            End If
            
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            txtCNPJ.Enabled = True
        
        Case "Manutenção da Tabela"
            formManutencaoTabelas.IniciarManutencao Me
    End Select
End Sub

Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim i           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer 'Contador de Registros
            
    
    If ValidarDados = False Then
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
Private Sub txtCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
    
End Sub

Private Sub txtCNPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        BuscarDados (txtCNPJ.Text)
    End If
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub
Private Sub BuscarDados(strCNPJ As String)
    Dim rst     As ADODB.Recordset
    Dim strSQL  As String
    
    'sstTransportadora.Tab = 0
    
    strSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND CNPJ = '" & strCNPJ & "'"

    Set rst = RegistroBuscar(strSQL)
    If rst.BOF And rst.EOF Then
            MsgBox "Nenhum Registro encontrado"
            rst.Close
            Exit Sub
        Else
            rst.MoveFirst
            IdReg = rst.Fields("Id")
            rst.Close
            MostrarDados
    End If
End Sub
Private Sub MostrarDados()
    Dim sSQL As String
    sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg
    ExibirDados Me, sSQL
End Sub
Private Function ValidarDados() As Boolean
    ValidarDados = False
    
    If Trim(cboPessoa.Text) = "" Then
        MsgBox "Selecione a pessoa FISICA ou JURIDICA!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    
    If Trim(txtCNPJ.Text) = "" Then
        MsgBox "Digite um CNPJ ou CPF valido!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    
    If Trim(txtxNome.Text) = "" Then
        MsgBox "Digite um Nome ou Razão Social valido!", vbInformation, "Aviso"
        ValidarDados = False
        Exit Function
    End If
    
    If Trim(cboUF.Text) = "" Then
        MsgBox "O campo UF esta com valor invalido. Favor verificar!"
        ValidarDados = False
        Exit Function
    End If
    If Validar_CNPJ_CPF(Trim(txtCNPJ.Text)) = False Then
        MsgBox "O campo CNPJ esta com valor invalido. Favor verificar!"
        ValidarDados = False
        Exit Function
    End If
    If Validar_IE(Trim(txtIE.Text), cboUF.Text) = False Then
        MsgBox "O campo Inscrição Estadual esta com valor invalido. Favor verificar!"
        ValidarDados = False
        Exit Function
    End If


    
    
    ValidarDados = True
End Function
Public Sub ReceberDadosTransportadora(CNPJ As String, _
                                    Optional Nome As String, _
                                    Optional IE As String, _
                                    Optional Ender As String, _
                                    Optional UF As String, _
                                    Optional Mun As String)

                                    
                                    
                                    

    txtCNPJ.Text = CNPJ
    txtxNome.Text = Nome
    txtIE.Text = IE
    txtxEnder.Text = Ender
    If Trim(UF) <> "" Then
        cboUF.AddItem UF
        cboUF.Text = cboUF.List(0)
    End If
    If Trim(Mun) <> "" Then
        cboMun.AddItem Mun
        cboMun.Text = cboMun.List(0)
    End If
    
    IdReg = 0
    HDMenu Me, False
    HDForm Me, True
End Sub

