VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formTributacaoCFOP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tributação - CFOP"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9060
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   60
      TabIndex        =   3
      Top             =   3840
      Width           =   8895
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1440
         Width           =   3255
      End
      Begin VB.ComboBox cboSituacao 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   300
         Width           =   2415
      End
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   1080
         MaxLength       =   500
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1080
         Width           =   6855
      End
      Begin VB.TextBox txtcCFOP 
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Situação:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   1140
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "CFOP:"
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   780
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   8895
      Begin MSFlexGridLib.MSFlexGrid msfgGrid 
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   5
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"formTributacaoCFOP.frx":0000
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
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
               Picture         =   "formTributacaoCFOP.frx":00C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoCFOP.frx":051A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoCFOP.frx":0834
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoCFOP.frx":10C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoCFOP.frx":2318
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoCFOP.frx":2BF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoCFOP.frx":3484
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoCFOP.frx":3D16
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoCFOP.frx":4F68
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoCFOP.frx":5282
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoCFOP.frx":559C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formTributacaoCFOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTabela     As String
Dim IdReg       As Integer
Private Sub LoadGrid()
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    msfgGrid.Rows = 1
    
    'sSQL = "SELECT * FROM " & sTabela & " WHERE ID_Empresa = " & ID_Empresa & " ORDER BY cCFOP"
    sSQL = "SELECT * FROM " & sTabela & " ORDER BY cCFOP"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            With msfgGrid
                Do Until Rst.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = cNull(Rst.Fields("id"))
                    .TextMatrix(.Rows - 1, 1) = cNull(Rst.Fields("Situacao"))
                    .TextMatrix(.Rows - 1, 2) = cNull(Rst.Fields("Tipo"))
                    .TextMatrix(.Rows - 1, 3) = cNull(Rst.Fields("cCFOP"))
                    .TextMatrix(.Rows - 1, 4) = cNull(Rst.Fields("Descricao"))
                    Rst.MoveNext
                Loop
                
            End With
    End If
    Rst.Close
    
End Sub
'########################################################################################################




'Dim sTabela   As String


Private Sub PesquisarRegistro()
    ''Dim idreg  As String
    IdReg = formBuscar.IniciarBusca(sTabela)
    ''IdReg = IIf(idreg = "", 0, idreg)
    
    If IdReg = 0 Then
            LimpaFormulario Me 'me
        Else
            MostrarDados
    End If
End Sub


Private Sub cboSituacao_DropDown()
    With cboSituacao
        .Clear
        .AddItem "0 - Entrada"
        .AddItem "1 - Saída"
    End With
End Sub

Private Sub cboTipo_DropDown()
    With cboTipo
        .Clear
        .AddItem "0 - Interna"
        .AddItem "1 - Interestadual"
        .AddItem "2 - Importação"
    End With
End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    sTabela = "TributacaoCFOP"
    HDForm Me, False
    HDMenu Me, True
    LoadGrid
    msfgGrid.Enabled = True
End Sub

Private Sub msfgGrid_DblClick()
    If Not IsNumeric(msfgGrid.TextMatrix(msfgGrid.Row, 0)) Then Exit Sub
    IdReg = msfgGrid.TextMatrix(msfgGrid.Row, 0)
    MostrarDados
End Sub
Private Sub MostrarDados()
    With msfgGrid
        cboSituacao.Clear
        cboSituacao.AddItem .TextMatrix(.Row, 1)
        cboSituacao.Text = cboSituacao.List(0)
        
        cboTipo.Clear
        cboTipo.AddItem .TextMatrix(.Row, 2)
        cboTipo.Text = cboTipo.List(0)
        
        txtcCFOP.Text = .TextMatrix(.Row, 3)
        txtDescricao.Text = .TextMatrix(.Row, 4)
    End With
End Sub
Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    IdReg = 0
    HDMenu Me, False
    HDForm Me, True
    'LimpaFormulario Me
    txtcCFOP.Text = ""
    txtDescricao.Text = ""
    msfgGrid.Enabled = False
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Selecione uma Registro.", vbInformation, App.EXEName
        Exit Sub
    End If
    HDForm Me, True
    HDMenu Me, False
    msfgGrid.Enabled = False
End Sub
Private Sub Excluir()
    If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
            MsgBox "Selecione um Registro!", vbInformation, App.EXEName
            Exit Sub
        Else
            If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                        vbCrLf & _
                        "Descrição.: " & txtDescricao.Text, vbYesNo + vbQuestion) = vbYes Then
                               
                If RegistroExcluir(sTabela, "Id = " & IdReg) = True Then
                    LimpaFormulario Me
                End If
            End If
    End If
    LoadGrid
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
            LoadGrid
            msfgGrid.Enabled = True
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            LoadGrid
            msfgGrid.Enabled = True
        Case "Manutenção da Tabela"
            formManutencaoTabelas.IniciarManutencao Me
    End Select
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
            If RegistroIncluir(sTabela, vReg, cReg) = 0 Then
                    MsgBox "Erro ao Incluir"
                    grvRegistro = False
                Else
                    grvRegistro = True
            End If
        Else
            If RegistroAlterar(sTabela, vReg, cReg, "Id = " & IdReg) = False Then
                    MsgBox "Erro ao Alterar."
                    grvRegistro = False
                Else
                    grvRegistro = True
                
            End If
    End If



End Function


