VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formRHFuncionarioTrocarComissao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RH - Trocar Comissionado"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   9435
      Begin VB.TextBox txtChvAcesso 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   960
         Width           =   7335
      End
      Begin VB.TextBox txtnNF 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   5835
      End
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Chave de Acesso:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Num. Nota:"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Novo Comissionado:"
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   1380
         Width           =   1455
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Picture         =   "formRHFuncionarioTrocarComissao.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioTrocarComissao.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioTrocarComissao.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioTrocarComissao.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioTrocarComissao.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioTrocarComissao.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioTrocarComissao.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioTrocarComissao.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioTrocarComissao.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioTrocarComissao.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioTrocarComissao.frx":54D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formRHFuncionarioTrocarComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboVendedor_DropDown()
    Dim Rst As Recordset
    cboVendedor.Clear
    Set Rst = RegistroBuscar("SELECT * FROM RHFuncionarioCadastro WHERE ID_Empresa = " & ID_Empresa & " ORDER BY xNome")
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboVendedor.AddItem Left(String(4, "0"), 4 - Len(Trim(Rst.Fields("ID")))) & Rst.Fields("ID") & " - " & Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If
End Sub


Public Function CarregarDadosNFe(chvnfe As String)
    Dim Rst     As Recordset
    Dim sSQL    As String
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND idNFe= '" & chvnfe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao carregar dados da Nfe " & chvnfe, vbInformation, "Aviso"
            HDForm Me, False
            HDMenu Me, True
            tbMenu.Buttons(2).Enabled = True
            Exit Function
        Else
            Rst.MoveFirst
            txtNome.Text = Rst.Fields("Dest_xNome")
            txtnNF.Text = Rst.Fields("ide_nNF")
            txtChvAcesso.Text = chvnfe
            cboVendedor.Clear
            cboVendedor.AddItem Left("0000", 4 - Len(Rst.Fields("ger_vendedor"))) & Rst.Fields("ger_vendedor") & " - " & _
                                PgDadosRhFuncionario(Rst.Fields("ger_Vendedor")).Nome
            cboVendedor.Text = cboVendedor.List(0)
            HDForm Me, True
            HDMenu Me, False
    End If
    Rst.Close
    Me.Show 1
End Function



Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Salvar"
            If grvRegistro = True Then
                Unload Me
            End If
        
        Case "Cancelar"
            Unload Me
    End Select
End Sub

Private Function grvRegistro() As Boolean
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    cReg = 0
    
    If Trim(cboVendedor.Text) = "" Then
        MsgBox "Selecione um Vendedor!", vbInformation, "Aviso"
        grvRegistro = False
        Exit Function
    End If
    
    vReg(cReg) = Array("Ger_Vendedor", Left(cboVendedor.Text, 4), "N")
    
    If RegistroAlterar("FaturamentoNFe", vReg, cReg, "idNFe = '" & txtChvAcesso.Text & "'") = True Then
            MsgBox "Registro alterado com sucesso!", vbInformation, "Aviso"
        Else
            MsgBox "Falha ao alterar registro!", vbInformation, "Aviso"
    End If
    
    
    Unload Me
End Function

Private Sub txtChvAcesso_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtnNF_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
