VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formEstoqueItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estoque - Cadastro de Item"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11025
   Begin VB.Frame Frame5 
      Caption         =   "Informações Complementares:"
      Height          =   2235
      Left            =   60
      TabIndex        =   15
      Top             =   4320
      Width           =   10395
      Begin VB.TextBox txtInformacoesComplementares 
         Height          =   1875
         Left            =   120
         MaxLength       =   65000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Text            =   "formEstoqueItem.frx":0000
         Top             =   240
         Width           =   10095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fisco/Tributos:"
      Height          =   1155
      Left            =   60
      TabIndex        =   10
      Top             =   3060
      Width           =   4635
      Begin VB.TextBox txtAliquotaIPI 
         Height          =   285
         Left            =   900
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   720
         Width           =   555
      End
      Begin VB.TextBox txtMVA 
         Height          =   285
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtNCM 
         Height          =   315
         Left            =   600
         MaxLength       =   8
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "MVA:"
         Height          =   255
         Left            =   2340
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Aliq. IPI:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Label7 
         Caption         =   "NCM:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quantidades:"
      Height          =   1155
      Left            =   4920
      TabIndex        =   7
      Top             =   3060
      Width           =   1935
      Begin VB.TextBox txtQtdMinima 
         Height          =   285
         Left            =   780
         MaxLength       =   5
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txtQtdMedia 
         Height          =   285
         Left            =   780
         MaxLength       =   5
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Minima:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Média:"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agrupamento:"
      Height          =   675
      Left            =   60
      TabIndex        =   5
      Top             =   2280
      Width           =   6135
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "Grupo:"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   10755
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtCodigoBarras 
         Height          =   315
         Left            =   7320
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1140
         Width           =   2235
      End
      Begin VB.ComboBox cboUnidade 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1260
         Width           =   795
      End
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   1020
         MaxLength       =   120
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   780
         Width           =   6795
      End
      Begin VB.TextBox txtReferencia 
         Height          =   315
         Left            =   1020
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Status:"
         Height          =   195
         Left            =   7440
         TabIndex        =   27
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label10 
         Caption         =   "Código de Barras:"
         Height          =   255
         Left            =   6000
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Unidade:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Referencia:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11025
      _ExtentX        =   19447
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
               Picture         =   "formEstoqueItem.frx":0006
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueItem.frx":0458
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueItem.frx":0772
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueItem.frx":1004
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueItem.frx":2256
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueItem.frx":2B30
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueItem.frx":33C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueItem.frx":3C54
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueItem.frx":4EA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueItem.frx":51C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEstoqueItem.frx":54DA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formEstoqueItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim IdReg     As Integer
Dim strTabela   As String


Private Sub PesquisarRegistro()
    Dim psqTMP  As String
    psqTMP = FormBusca.IniciarBusca(strTabela)
    IdReg = IIf(psqTMP = "", 0, psqTMP)
    
    If IdReg = 0 Then
            LimpaFormulario Me 'me
        Else
            MostrarDados
    End If
End Sub


Private Sub cboAcao_DropDown()
    cboAcao.Clear
    cboAcao.AddItem "SOMAR (+)"
    cboAcao.AddItem "SUBTRAIR (-)"
    cboAcao.AddItem "NENHUM"

End Sub

Private Sub cboGrupo_DropDown()
    Dim Rst As Recordset
    cboGrupo.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueGrupos ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
            
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboGrupo.AddItem Rst.Fields("Referencia") & " - " & Rst.Fields("descricao")
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboStatus_DropDown()
    cboStatus.Clear
    cboStatus.AddItem "Ativo"
    cboStatus.AddItem "Inativo"
End Sub

Private Sub cboUnidade_DropDown()
    Dim Rst As Recordset
    cboUnidade.Clear
    Set Rst = RegistroBuscar("SELECT * FROM EstoqueUnidadeMedida ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboUnidade.AddItem Rst.Fields("sigla")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    HDForm Me, False
    HDMenu Me, True
    
  
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
            IdReg = 0
            HDMenu Me, False
            HDForm Me, True
           LimpaFormulario Me
        Case "Alterar"
            If IdReg = 0 Then
                MsgBox "Selecione uma Grupo"
                Exit Sub
            End If
            HDForm Me, True
            HDMenu Me, False
        Case "Excluir"
            If IdReg = 0 Then
                    MsgBox "Selecione um Registro"
                    Exit Sub
                Else
                    If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                               vbCrLf & _
                               "Descrição: " & txtDescricao.Text & vbCrLf & _
                                vbYesNo + vbCritical) = vbYes Then
                               
                        If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                            LimpaFormulario Me
                        End If
                    End If
            End If
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
            formManutencaoTabelas.IniciarManutencao Me
            
    End Select
End Sub

Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim I           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer 'Contador de Registros
    cReg = 0
    For I = 0 To Me.Controls.Count - 1
        Set Controle = Me.Controls(I)
        
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
    
     
    If IdReg = 0 Then
            If RegistroIncluir(strTabela, vReg, cReg) = False Then
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
    sSQL = "SELECT * FROM " & strTabela & " WHERE Id = " & IdReg

    ExibirDados Me, sSQL


End Sub





