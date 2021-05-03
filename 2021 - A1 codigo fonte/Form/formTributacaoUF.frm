VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formTributacaoUF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tributa��o - Tabela de UF"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8490
   Begin VB.Frame Frame2 
      Height          =   3795
      Left            =   60
      TabIndex        =   7
      Top             =   420
      Width           =   8355
      Begin MSFlexGridLib.MSFlexGrid msfgTabICMS 
         Height          =   3495
         Left            =   60
         TabIndex        =   8
         ToolTipText     =   "Duplo click para selecionar..."
         Top             =   180
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   7
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "^Id |<Descri��o                                    |^Sigla   |^ICMS Interestadual |^ICMS Interna  |^FECP   |^Cod.UF "
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2475
      Left            =   60
      TabIndex        =   3
      Top             =   4260
      Width           =   8355
      Begin VB.TextBox txtICMSFECP 
         Height          =   285
         Left            =   3240
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1470
         Width           =   975
      End
      Begin VB.TextBox txtICMSInt 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox txtCodUF 
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1920
         Width           =   915
      End
      Begin VB.TextBox txtICMS 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtSigla 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "FECP:"
         Height          =   195
         Left            =   2700
         TabIndex        =   14
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "ICMS Interna:"
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Cod. UF: "
         Height          =   195
         Left            =   660
         TabIndex        =   10
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ICMS Interestadual:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1140
         Width           =   1395
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Descri��o:"
         Height          =   195
         Left            =   540
         TabIndex        =   5
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sigla:"
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   795
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8490
      _ExtentX        =   14975
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
            Object.ToolTipText     =   "Manuten��o da Tabela"
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
               Picture         =   "formTributacaoUF.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoUF.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoUF.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoUF.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoUF.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoUF.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoUF.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoUF.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoUF.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoUF.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoUF.frx":54D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formTributacaoUF"
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
    MostrarGrid
    msfgTabICMS.Enabled = True
End Sub

Private Sub msfgTabICMS_DblClick()
    If Not IsNumeric(msfgTabICMS.TextMatrix(msfgTabICMS.Row, 0)) Then Exit Sub
    IdReg = msfgTabICMS.TextMatrix(msfgTabICMS.Row, 0)
    MostrarDados
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
        MsgBox "Selecione uma Registro."
        Exit Sub
    End If
    HDForm Me, True
    HDMenu Me, False
    msfgTabICMS.Enabled = False
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
                        "Descri��o.: " & txtDescricao.Text, vbYesNo + vbQuestion) = vbYes Then
                               
                If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                    LimpaFormulario Me
                End If
            End If
    End If
    MostrarGrid
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
            MostrarGrid
            msfgTabICMS.Enabled = True
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpaFormulario Me
            MostrarGrid
            msfgTabICMS.Enabled = True
        Case "Manuten��o da Tabela"
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
    Dim sSQL As String
    'sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg
    sSQL = "SELECT * FROM " & strTabela & " WHERE Id = " & IdReg

    ExibirDados Me, sSQL


End Sub


Private Sub MostrarGrid()
    On Error Resume Next
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    msfgTabICMS.Rows = 1
   sSQL = "SELECT * FROM " & strTabela & " ORDER BY Sigla"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                With msfgTabICMS
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("ID")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("Descricao")
                    .TextMatrix(.Rows - 1, 2) = Rst.Fields("Sigla")
                    .TextMatrix(.Rows - 1, 3) = Rst.Fields("ICMS")
                    .TextMatrix(.Rows - 1, 4) = Rst.Fields("ICMSInt")
                    .TextMatrix(.Rows - 1, 5) = Rst.Fields("ICMSFECP")
                    .TextMatrix(.Rows - 1, 6) = IIf(IsNull(Rst.Fields("CodUF")), "", Rst.Fields("CodUF"))
                End With
                Rst.MoveNext
            Loop
    End If
End Sub


Private Sub txtICMS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    KeyAscii = ChkVal(txtICMS.Text, KeyAscii, 2)
End Sub



Private Sub txtICMSFECP_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
    KeyAscii = ChkVal(txtICMSFECP.Text, KeyAscii, 2)
End Sub

Private Sub txtICMSInt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    KeyAscii = ChkVal(txtICMSInt.Text, KeyAscii, 2)
End Sub
