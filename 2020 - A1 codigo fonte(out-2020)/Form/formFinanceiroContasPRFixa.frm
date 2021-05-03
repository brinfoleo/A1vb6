VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formFinanceiroContasPRFixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financeiro - Contas Pagar/Receber Fixas"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   12525
   Begin VB.Frame Frame1 
      Caption         =   "Meses de Vencimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6300
      TabIndex        =   35
      Top             =   3300
      Width           =   6135
      Begin VB.CheckBox chkMes 
         Caption         =   "Dezembro"
         Height          =   195
         Index           =   11
         Left            =   4140
         TabIndex        =   47
         Top             =   1020
         Width           =   1155
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Novembro"
         Height          =   195
         Index           =   10
         Left            =   4140
         TabIndex        =   46
         Top             =   780
         Width           =   1155
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Outubro"
         Height          =   195
         Index           =   9
         Left            =   4140
         TabIndex        =   45
         Top             =   540
         Width           =   1155
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Setembro"
         Height          =   195
         Index           =   8
         Left            =   4140
         TabIndex        =   44
         Top             =   300
         Width           =   1155
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Agosto"
         Height          =   195
         Index           =   7
         Left            =   2280
         TabIndex        =   43
         Top             =   1020
         Width           =   1035
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Julho"
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   42
         Top             =   780
         Width           =   1035
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Junho"
         Height          =   195
         Index           =   5
         Left            =   2280
         TabIndex        =   41
         Top             =   540
         Width           =   1035
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Maio"
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   40
         Top             =   300
         Width           =   1035
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Abril"
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   39
         Top             =   1020
         Width           =   1035
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Março"
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   38
         Top             =   780
         Width           =   1035
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Fevereiro"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   37
         Top             =   540
         Width           =   1035
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Janeiro"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   36
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Periodo de abrangência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   6300
      TabIndex        =   29
      Top             =   2160
      Width           =   6135
      Begin VB.TextBox txtDia 
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   540
         Width           =   435
      End
      Begin VB.CheckBox chkAntSabDom 
         Caption         =   "Antecipar venc. Sab./Dom."
         Height          =   375
         Left            =   4620
         TabIndex        =   30
         Top             =   600
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpvInicial 
         Height          =   315
         Left            =   300
         TabIndex        =   31
         Top             =   540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   136314883
         CurrentDate     =   40994
      End
      Begin MSComCtl2.DTPicker dtpvFinal 
         Height          =   315
         Left            =   1860
         TabIndex        =   32
         Top             =   540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   136314883
         CurrentDate     =   40994
      End
      Begin VB.Label Label9 
         Caption         =   "Dia do Vencimento"
         Height          =   195
         Left            =   3240
         TabIndex        =   48
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Ultimo Vencimento"
         Height          =   195
         Left            =   1680
         TabIndex        =   34
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Primeiro Vencimento"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Observações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   27
      Top             =   4740
      Width           =   12315
      Begin VB.TextBox txtObs 
         Height          =   735
         Left            =   120
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   28
         Text            =   "formFinanceiroContasPRFixa.frx":0000
         Top             =   240
         Width           =   12075
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Controle Interno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   6300
      TabIndex        =   20
      Top             =   540
      Width           =   6135
      Begin VB.ComboBox cboCentroCustos 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   4575
      End
      Begin VB.ComboBox cboDocumento 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   660
         Width           =   4575
      End
      Begin VB.ComboBox cboPlanoContas 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Centro de Custos:"
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Documento:"
         Height          =   195
         Left            =   300
         TabIndex        =   25
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Plano de Contas:"
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   1140
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Sacado/Cedente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   60
      TabIndex        =   11
      Top             =   3000
      Width           =   6135
      Begin VB.TextBox txtDoc 
         Height          =   285
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   720
         Width           =   2595
      End
      Begin VB.ComboBox cboNome 
         Height          =   315
         Left            =   1140
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   1080
         Width           =   4935
      End
      Begin VB.ComboBox cboCadastro 
         Height          =   315
         ItemData        =   "formFinanceiroContasPRFixa.frx":0006
         Left            =   1140
         List            =   "formFinanceiroContasPRFixa.frx":0008
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome/Razão:"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ/CFP:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Cadastro:"
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Identificação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   60
      TabIndex        =   4
      Top             =   540
      Width           =   6135
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   660
         Width           =   4995
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "ID:"
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de lançamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   1800
      Width           =   6135
      Begin VB.TextBox txtvFatura 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtnFatura 
         Height          =   285
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optTpLanc 
         Caption         =   "A Receber"
         Height          =   195
         Index           =   1
         Left            =   1380
         TabIndex        =   2
         Top             =   465
         Width           =   1095
      End
      Begin VB.OptionButton optTpLanc 
         Caption         =   "A Pagar"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   465
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2820
         TabIndex        =   18
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Fatura:"
         Height          =   195
         Left            =   3000
         TabIndex        =   9
         Top             =   300
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
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
            Object.ToolTipText     =   "Imprimir Listagem Geral"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pesquisar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Manutenção da Tabela"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4380
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":000A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":045C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":0776
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":1008
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":225A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":2B34
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":33C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":3C58
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":4EAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":51C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":54DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContasPRFixa.frx":58D5
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFinanceiroContasPRFixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg       As Integer
Dim IDSacado    As Integer
Dim sTabSacado  As String
Dim strTabela   As String
Private Sub cboCadastro_DropDown()
    With cboCadastro
        .Clear
        .AddItem "00 - Outros"
        .AddItem "01 - Clientes"
        .AddItem "02 - Fornecedores"
        .AddItem "03 - Transportadora"
        .AddItem "04 - Funcionario"
    End With
End Sub
Private Sub cboCadastro_Click()
    If cboCadastro.Text = "" Then Exit Sub
    IDSacado = 0
    cboNome.Clear
    txtDoc.Text = ""
    Select Case Left(Trim(cboCadastro.Text), 2)
        Case 1 'Cliente
            sTabSacado = "Clientes"
        Case 2 'Fornecedor
            sTabSacado = "Fornecedores"
        Case 3 ' Transportadora
            sTabSacado = "Transportadoras"
        Case 4 ' Transportadora
            sTabSacado = "RHFuncionarioCadastro"
        Case Else
            sTabSacado = ""
    End Select
End Sub

Private Sub cboCentroCustos_DropDown()
    Dim Rst As Recordset
    cboCentroCustos.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCentroCustos")
    If Rst.BOF And Rst.EOF Then
            'Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCentroCustos.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub cboDocumento_DropDown()
    Dim Rst As Recordset
    cboDocumento.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento")
    If Rst.BOF And Rst.EOF Then
            'Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboDocumento.AddItem Left(String(3, "0"), 3 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & _
                                 Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub



Private Sub cboPlanoContas_DropDown()
    Dim Rst As Recordset
    cboPlanoContas.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroPlanoContas ORDER BY Codigo")
    If Rst.BOF And Rst.EOF Then
            'Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboPlanoContas.AddItem ZE(Rst.Fields("id"), 3) & " - (" & Rst.Fields("Codigo") & ") " & Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
    HDMenu Me, True
    HDFormulario False
    strTabela = Mid(Me.Name, 5, Len(Me.Name))

    dtpvInicial.Value = Date
    dtpvFinal.Value = Date + 30
    Me.Top = 0
    Me.Left = 0
End Sub
Private Sub HDFormulario(op As Boolean)
    HDForm Me, op
    txtID.Enabled = IIf(op = False, True, False)
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
            IdReg = 0
            PesquisarRegistro
        Case "Imprimir Listagem Geral"
            ImprimirListagemGeral
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDFormulario False
            End If
        Case "Cancelar"
            HDMenu Me, True
            HDFormulario False
            LimpaFormulario Me
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select
End Sub
Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    IdReg = 0
    IDSacado = 0
    sTabSacado = ""
    HDFormulario True
    HDMenu Me, False
    LimpaFormulario Me
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    'IDReg = 0
    'IDSacado = 0
    'sTabSacado = ""
    HDFormulario True
    HDMenu Me, False
    'LimpaFormulario Me
End Sub
Private Sub Excluir()
    If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
            MsgBox "Selecione um Registro!", vbCritical, App.EXEName
            Exit Sub
        Else
            If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                        vbCrLf & _
                        "ID: " & IdReg & vbCrLf & _
                        "Descrição: " & txtDescricao.Text, vbYesNo + vbQuestion) = vbYes Then
                              
                If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                    IdReg = 0
                    IDSacado = 0
                    sTabSacado = ""
                    LimpaFormulario Me
                End If
            End If
    End If

End Sub
Private Function grvRegistro() As Boolean
    Dim vReg(199)   As Variant
    Dim cReg        As Integer 'Contador de Registros
    'Dim l           As Integer
    Dim i           As Integer
    Dim tmp         As Integer
    Dim ContaTp     As String
    Dim Meses       As String
    Dim DtIni       As Date
    Dim DtFin       As Date
    
    If Validar = False Then
        grvRegistro = False
        Exit Function
    End If
    
    If optTpLanc.Item(0).Value = True Then
            ContaTp = "P"
        Else
            ContaTp = "R"
    End If
    
    cReg = 0
    vReg(cReg) = Array("Descricao", Trim(txtDescricao.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("contaPR", ContaTp, "S"): cReg = cReg + 1
    'vReg(cReg) = Array("Emissao", dtpEmissao.Value, "D"): cReg = cReg + 1
    vReg(cReg) = Array("nFatura", txtnFatura.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("vFatura", ChkVal(Trim(txtvFatura.Text), 0, cDecMoeda), "S"): cReg = cReg + 1
       
    vReg(cReg) = Array("Tabela", sTabSacado, "S"): cReg = cReg + 1
    vReg(cReg) = Array("IdSacado", IDSacado, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Nome", cboNome.Text, "S"): cReg = cReg + 1
    vReg(cReg) = Array("CNPJ", txtDoc.Text, "S"): cReg = cReg + 1
    
    'vReg(cReg) = Array("Conta", Left(cboConta.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("CentroCusto", Left(cboCentroCustos.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("TpDocumento", Left(cboDocumento.Text, 3), "S"): cReg = cReg + 1
    vReg(cReg) = Array("PlanoContas", Left(cboPlanoContas.Text, 3), "S"): cReg = cReg + 1
    
    Meses = ""
    For i = 0 To 11
        If chkMes(i).Value = 1 Then
            Meses = Meses & "/" & Left("00", 2 - Len(Trim(i) + 1)) & i + 1
        End If
    Next
    
     vReg(cReg) = Array("Meses", Meses, "S"): cReg = cReg + 1
    
    
    DtIni = DiadoMes(dtpvInicial.Value, 0)
    DtFin = DiadoMes(dtpvFinal.Value, 1)
    
    
    vReg(cReg) = Array("VencDia", txtDia.Text, "N"): cReg = cReg + 1
    vReg(cReg) = Array("VencInicial", DtIni, "D"): cReg = cReg + 1
    vReg(cReg) = Array("VencFinal", DtFin, "D"): cReg = cReg + 1
    
    vReg(cReg) = Array("AntSabDom", chkAntSabDom.Value, "N"): cReg = cReg + 1
    
    
    vReg(cReg) = Array("Obs", Trim(txtObs.Text), "S"): cReg = cReg + 1
    
    cReg = cReg - 1
    If IdReg = 0 Then
            IdReg = RegistroIncluir(strTabela, vReg, cReg)
            If IdReg = 0 Then
                    MsgBox "Erro ao Incluir"
                    grvRegistro = False
                Else
                    txtID.Text = IdReg
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

Private Sub MontarBaseDeDados()
    Dim vReg(199)  As Variant
    Dim cReg     As Integer
    Dim i           As Integer
    
    cReg = 0
    
    vReg(cReg) = Array("Descricao", txtDescricao.MaxLength, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("ContaPR", "10", "S"): cReg = cReg + 1
    
    'vReg(cReg) = Array("Emissao", "10", "D"): cReg = cReg + 1
    vReg(cReg) = Array("nFatura", "100", "S"): cReg = cReg + 1
    vReg(cReg) = Array("vFatura", "100", "S"): cReg = cReg + 1
    
    'vReg(cReg) = Array("Conta", "10", "N"): cReg = cReg + 1
    vReg(cReg) = Array("CentroCusto", "10", "N"): cReg = cReg + 1
    vReg(cReg) = Array("TpDocumento", "10", "N"): cReg = cReg + 1
    vReg(cReg) = Array("PlanoContas", "10", "N"): cReg = cReg + 1
    
    vReg(cReg) = Array("Tabela", "50", "S"): cReg = cReg + 1
    vReg(cReg) = Array("idSacado", "50", "N"): cReg = cReg + 1
    vReg(cReg) = Array("CNPJ", "30", "S"): cReg = cReg + 1
    vReg(cReg) = Array("Nome", "120", "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("Meses", "100", "S"): cReg = cReg + 1
    'vReg(cReg) = Array("LinhaDigitavel", "100", "S"): cReg = cReg + 1
    'vReg(cReg) = Array("NossoNumero", "30", "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("VencInicial", "10", "D"): cReg = cReg + 1
    vReg(cReg) = Array("VencFinal", "10", "D"): cReg = cReg + 1
    vReg(cReg) = Array("VencDia", "10", "N"): cReg = cReg + 1
    vReg(cReg) = Array("AntSabDom", "1", "N"): cReg = cReg + 1
    
    
    'vReg(cReg) = Array("NumDuplicata", "30", "S"): cReg = cReg + 1
    'vReg(cReg) = Array("VlDuplicata", "30", "DC"): cReg = cReg + 1
    
    'vReg(cReg) = Array("Multa", "10", "S"): cReg = cReg + 1
    'vReg(cReg) = Array("Juros", "10", "S"): cReg = cReg + 1
    'vReg(cReg) = Array("DiasProtesto", "10", "N"): cReg = cReg + 1
    
    
    'vReg(cReg) = Array("IdBanco", "10", "N"): cReg = cReg + 1
    
    'vReg(cReg) = Array("Acrescimo", "30", "S"): cReg = cReg + 1
    'vReg(cReg) = Array("Abatimento", "30", "S"): cReg = cReg + 1
    'vReg(cReg) = Array("Deducoes", "30", "S"): cReg = cReg + 1
    'vReg(cReg) = Array("MultaMora", "30", "S"): cReg = cReg + 1
    'vReg(cReg) = Array("VlCobrado", "30", "S"): cReg = cReg + 1
    
    'vReg(cReg) = Array("DataQuitacao", "30", "D"): cReg = cReg + 1
    'vReg(cReg) = Array("IdContaQuitacao", "10", "N"): cReg = cReg + 1
    
    vReg(cReg) = Array("Obs", txtObs.MaxLength, "S"): cReg = cReg + 1
    
    'vReg(cReg) = Array("ObsBol1", "2000", "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ObsBol2", "2000", "S"): cReg = cReg + 1
    'vReg(cReg) = Array("ObsBol3", "2000", "S"): cReg = cReg + 1
    
    'vReg(cReg) = Array("ide_NFe", "60", "S"): cReg = cReg + 1
    
    'vReg(cReg) = Array("IdFixa", "50", "S"): cReg = cReg + 1
    
    cReg = cReg - 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vReg, cReg
    
End Sub
Private Sub cboNome_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    If Trim(cboCadastro.Text) = "" Or sTabSacado = "" Then
        MsgBox "Selecione um tipo de cadastro."
        Exit Sub
    End If
    'cboNome.Clear
    If Left(cboCadastro.Text, 2) = "00" Then
        Exit Sub
    End If
    sSQL = "SELECT * FROM " & sTabSacado & _
           " WHERE ID_Empresa = " & ID_Empresa & _
           " AND xNome LIKE '" & Trim(cboNome.Text) & "%'" & _
           " ORDER BY xNome"
    
    If sSQL = "" Then Exit Sub
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboNome.AddItem Left(String(5, "0"), 5 - Len(Trim(Rst.Fields("id")))) & Trim(Rst.Fields("id")) & " - " & Rst.Fields("xNome")
                Rst.MoveNext
            Loop
    End If
    Rst.Close

End Sub
Private Sub cboNome_Click()
    If Trim(cboNome.Text) = "" Then Exit Sub
    txtDoc.Text = ""
    IDSacado = Trim(Left(cboNome.Text, 5))
    PesquisarSacado
End Sub
Private Sub cboNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IDSacado = 0
        PesquisarSacado
    End If
End Sub



Private Sub txtDia_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
    
End Sub



Private Sub txtDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IDSacado = 0
        PesquisarSacado
    End If
End Sub
Private Sub txtDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PesquisarSacado Trim(txtDoc.Text)
    End If
    KeyAscii = SoNumeros(KeyAscii)
End Sub


Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        IdReg = 0
        PesquisarRegistro
    End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IdReg = Trim(txtID.Text)
        PesquisarRegistro
    End If
    KeyAscii = SoNumeros(KeyAscii)
End Sub


Private Sub txtvFatura_GotFocus()
    txtvFatura.Text = ChkVal(txtvFatura.Text, 0, cDecMoeda)
    txtvFatura.SelStart = 0
    txtvFatura.SelLength = Len(txtvFatura.Text)
End Sub

Private Sub txtvFatura_KeyPress(KeyAscii As Integer)
    
    If txtvFatura.SelLength = Len(txtvFatura.Text) Then
        txtvFatura.Text = ""
    End If
    KeyAscii = ChkVal(txtvFatura.Text, KeyAscii, cDecMoeda)
End Sub


Private Sub txtvFatura_LostFocus()
    txtvFatura.Text = ConvMoeda(txtvFatura.Text)
End Sub

Private Sub PesquisarSacado(Optional sCNPJ As String)
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim Campo   As String
    
    If Trim(sTabSacado) = "" Then Exit Sub
    
    
    Select Case sTabSacado
        Case "Clientes"
            Campo = "Doc"
        Case "Fornecedores"
            Campo = "Doc"
        Case "Transportador"
            Campo = "CNPJ"
        Case "RHFuncionarioCadastro"
            Campo = "CPF"
        Case Else
            Exit Sub
    End Select
    
    
    If Trim(sCNPJ) <> "" Then
            sSQL = "SELECT * FROM " & sTabSacado & " WHERE " & Campo & "= '" & sCNPJ & "'"
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                    'Exit Sub
                Else
                    Rst.MoveFirst
                    IDSacado = Rst.Fields("ID")
                    'PesquisarSacado
            End If
            Rst.Close
        Else
            If Left(cboCadastro.Text, 2) = "00" Then
                Exit Sub
            End If
    
            If IDSacado = 0 Then
                IDSacado = formBuscar.IniciarBusca(sTabSacado)
            End If
    End If
   
    Select Case sTabSacado
        Case "Fornecedores"
            txtDoc.Text = PgDadosFornecedor(IDSacado).Doc
            'cboNome.Clear
            cboNome.Text = PgDadosFornecedor(IDSacado).Nome
                    
        Case "Clientes"
            txtDoc.Text = PgDadosCliente(IDSacado).Doc
            'cboNome.Clear
            cboNome.Text = PgDadosCliente(IDSacado).Nome
                    
        Case "Transportadoras"
            txtDoc.Text = pgDadosTransportadora(IDSacado).CNPJ
            'cboNome.Clear
            cboNome.Text = pgDadosTransportadora(IDSacado).Nome
        Case "RHFuncionarioCadastro"
            txtDoc.Text = PgDadosRhFuncionario(IDSacado).CPF
            'cboNome.Clear
            cboNome.Text = PgDadosRhFuncionario(IDSacado).Nome
        Case Else
            MsgBox "Selecione um tipo de Cadastro"
    End Select
End Sub

Private Sub PesquisarRegistro()
    On Error Resume Next
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim i       As Integer
    LimpaFormulario Me
    

    If IdReg = 0 Then
        IdReg = formBuscar.IniciarBusca(strTabela)
    End If
    If IdReg = 0 Then Exit Sub
    sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND ID=" & IdReg
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar Registro", vbCritical, App.EXEName
            Exit Sub
        Else
            Rst.MoveFirst
            DoEvents
            txtID.Text = IdReg
            txtDescricao.Text = cNull(Rst.Fields("Descricao"))
            If Rst.Fields("ContaPR") = "P" Then
                    optTpLanc.Item(0).Value = True
                    optTpLanc.Item(1).Value = False
                Else
                    optTpLanc.Item(0).Value = False
                    optTpLanc.Item(1).Value = True
            End If
            
            cboCadastro.Clear
            cboCadastro.AddItem nomeTabela(cNull(Rst.Fields("Tabela")))
            cboCadastro.Text = cboCadastro.List(0)
            
            
            IDSacado = Rst.Fields("idSacado")
            cboNome.Text = Rst.Fields("Nome")
            txtDoc.Text = Rst.Fields("CNPJ")
            
            txtDia.Text = Rst.Fields("vencDia")
            dtpvInicial.Value = Rst.Fields("vencInicial")
            dtpvFinal.Value = Rst.Fields("vencFinal")
            chkAntSabDom.Value = Rst.Fields("AntSabDom")
            
            For i = 1 To 12
                If InStr(Rst.Fields("Meses"), Left("00", 2 - Len(Trim(i))) & Trim(i)) <> 0 Then
                    chkMes(i - 1).Value = 1
                End If
            Next
            
            
            txtnFatura.Text = Rst.Fields("nFatura")
            txtvFatura.Text = ConvMoeda(IIf(IsNull(Rst.Fields("vFatura")), "0", Rst.Fields("vFatura")))
            
            
            cboCentroCustos.Clear
            cboCentroCustos.AddItem pgDadosCentroCustos(Rst.Fields("CentroCusto")).Id & " - " & pgDadosCentroCustos(Rst.Fields("CentroCusto")).Descricao
            cboCentroCustos.Text = cboCentroCustos.List(0)
            
            cboPlanoContas.Clear
            cboPlanoContas.AddItem ZE(PgDadosPlanoContas("ID", Rst.Fields("PlanoContas")).Id, 3) & " - (" & PgDadosPlanoContas("ID", Rst.Fields("PlanoContas")).Codigo & ") " & PgDadosPlanoContas("ID", Rst.Fields("PlanoContas")).Descricao
            cboPlanoContas.Text = cboPlanoContas.List(0)
            
            cboDocumento.AddItem pgDadosTipoDocumento(Rst.Fields("tpDocumento")).Id & " - " & pgDadosTipoDocumento(Rst.Fields("tpDocumento")).Descricao
            cboDocumento.Text = cboDocumento.List(0)
            
            
            txtObs.Text = cNull(Rst.Fields("OBS"))

            
    End If
    Rst.Close
End Sub

Private Function nomeTabela(sTabela As String) As String
        Select Case LCase(sTabela)
            Case "clientes"
                nomeTabela = "01 - Clientes"
            Case "fornecedores"
                nomeTabela = "02 - Fornecedores"
            Case "trasportadoras"
                nomeTabela = "03 - Transportadora"
            Case "rhfuncionariocadastro"
                nomeTabela = "04 - Funcionario"
            Case Else
                nomeTabela = "00 - Outros"
        End Select
End Function
Public Function DiadoMes(strData As String, op As Integer) As Date
    '##########################################################################################################
    '### 26/03/2012
    '### op= 0 primeiro / 1 - ultimo dia
    '###
    '##########################################################################################################
    Dim strAno  As String
    Dim strMes  As String
    Dim strDia  As String
    Dim DataF   As Date
    strData = Format(strData, "yyyymmdd")

    strAno = Mid$(strData, 1, 4)
    strMes = Mid$(strData, 5, 2)

    Select Case strMes
        Case "04", "06", "09", "11"
            strDia = "30"

        Case "02"
            If Bissexto(Val(strAno)) Then
                    strDia = "29"
                Else
                    strDia = "28"
            End If

        Case Else
            strDia = "31"
    End Select
    If op = 0 Then
        strDia = "01"
    End If
    
    DataF = strDia & "/" & strMes & "/" & strAno
    DiadoMes = Format(DataF, "DD/MM/YYYY")
    'If (DiadoMes = strData) And SaltaMesAtual Then
    '    strProximoDia = ProximoDia(strData)
    '    DiadoMes = DiadoMes(strProximoDia, False)
    'End If

End Function

Private Function Validar() As Boolean
    
    If Trim(txtDescricao.Text) = "" Then
        MsgBox "O campo DESCRICAO não é valido!", vbInformation, App.EXEName
        Validar = False
        Exit Function
    End If
    
    If optTpLanc(0).Value = 0 And optTpLanc(1).Value = 0 Then
        MsgBox "O campo A PAGAR / A RECEBER não é valido!", vbInformation, App.EXEName
        Validar = False
        Exit Function
    End If
    If Trim(txtnFatura.Text) = "" Then
        MsgBox "O campo FATURA não é valido!", vbInformation, App.EXEName
        Validar = False
        Exit Function
    End If
    
    If Trim(txtvFatura.Text) = "" Then
        MsgBox "O campo VALOR não é valido!", vbInformation, App.EXEName
        Validar = False
        Exit Function
    End If
    If Trim(cboNome.Text) = "" Then
        MsgBox "O campo NOME não é valido!", vbInformation, App.EXEName
        Validar = False
        Exit Function
    End If
    If txtDia.Text <= 0 Or txtDia.Text > 31 Then
        MsgBox "Dia do vencimento invalido!", vbInformation, App.EXEName
        Validar = False
        Exit Function
    End If
    Validar = True
    
End Function
Private Sub ImprimirListagemGeral()
    '###################################################################################
    '### 27/03/2012
    '### Imprime a listagem geral dos titulos fixos
    '###################################################################################
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    sSQL = "SELECT * FROM FinanceiroContasPRFixa"
    Set Rst = RegistroBuscar(sSQL)
    
    If Rst.BOF And Rst.EOF Then
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            Set rptListaTitulosFixos.DataSource = Rst.DataSource
            rptListaTitulosFixos.Orientation = 2 'rptOrientLandscape
            'rptListaTitulosFixos.Sections("Section5").Controls.Item("lblCred").Caption = ConvMoeda(sDuplAc)
            'rptListaTitulosFixos.Sections("Section5").Controls.Item("lblDeb").Caption = ConvMoeda(sDuplAd)
            rptListaTitulosFixos.Show 1 '.ExportReport
    End If
    
    
End Sub
