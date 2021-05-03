VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formFinanceiroBaixaAutomaticaTitulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Baixa Automatica de Titulos"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7020
   Begin VB.CheckBox chkConfirmacao 
      Caption         =   "Confirmar titulo antes de baixar."
      Height          =   435
      Left            =   4680
      TabIndex        =   13
      Top             =   1260
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Quitação:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   6735
      Begin VB.ComboBox cboConta 
         Height          =   315
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   3795
      End
      Begin MSComCtl2.DTPicker dtpQuitacao 
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55771137
         CurrentDate     =   40696
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Conta:"
         Height          =   195
         Left            =   2220
         TabIndex        =   15
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Data:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   420
         Width           =   375
      End
   End
   Begin VB.OptionButton optTipoDoc 
      Caption         =   "AMBOS"
      Height          =   195
      Index           =   2
      Left            =   5340
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton optTipoDoc 
      Caption         =   "A RECEBER"
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   600
      Width           =   1515
   End
   Begin VB.OptionButton optTipoDoc 
      Caption         =   "A PAGAR"
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Value           =   -1  'True
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo do Vencimento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   4395
      Begin MSComCtl2.DTPicker dtpAte 
         Height          =   315
         Left            =   2700
         TabIndex        =   7
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55771137
         CurrentDate     =   40696
      End
      Begin MSComCtl2.DTPicker dtpDe 
         Height          =   315
         Left            =   540
         TabIndex        =   6
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55771137
         CurrentDate     =   40696
      End
      Begin VB.Label Label2 
         Caption         =   "Até:"
         Height          =   195
         Left            =   2340
         TabIndex        =   2
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "De:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   5
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
               Picture         =   "formFinanceiroBaixaAutomaticaTitulo.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroBaixaAutomaticaTitulo.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroBaixaAutomaticaTitulo.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroBaixaAutomaticaTitulo.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroBaixaAutomaticaTitulo.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroBaixaAutomaticaTitulo.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroBaixaAutomaticaTitulo.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroBaixaAutomaticaTitulo.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroBaixaAutomaticaTitulo.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroBaixaAutomaticaTitulo.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroBaixaAutomaticaTitulo.frx":54D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo de Documento:"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "formFinanceiroBaixaAutomaticaTitulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'#############################################################################################
'### Data: 28/09/2011
'### Objetivo deste form e baixar todos os documentos com seu valor nominal
'###
'#############################################################################################
Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub
Private Sub cboConta_DropDown()
    Dim sSQL    As String
    Dim Rst     As Recordset
    cboConta.Clear
    sSQL = "SELECT * FROM FinanceiroConta"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma conta cadastrada!", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboConta.AddItem ZE(Rst.Fields("ID"), 3) & " - " & Rst.Fields("Agencia") & "/" & Rst.Fields("Conta")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub

Private Sub Form_Load()
    dtpDe.Value = Date
    dtpAte.Value = Date
    dtpQuitacao.Value = Date
    
End Sub


Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Salvar"
           Quitar
    End Select
End Sub
Private Sub Quitar()
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim tpDoc       As String
    Dim vDados(1)   As Variant
    Dim cReg        As Integer
    Dim tReg        As Integer
    Dim IdConta     As Integer
    'Verifica se tem permissao de acesso
    If chkAcesso(Me, "a") = False Then Exit Sub
    
    
    'pega o Id da Conta para Mov. Banco
    If Trim(cboConta.Text) = "" Then
        MsgBox "Favor selecionar uma conta!", vbInformation, "Aviso"
        Exit Sub
    End If
    IdConta = Left(Trim(cboConta.Text), 3)
    
    'Pega o tipo de documento
    If optTipoDoc(0).Value = True Then
            tpDoc = " AND ContaPR = 'P'"
        ElseIf optTipoDoc(1).Value = True Then
            tpDoc = " AND ContaPR = 'R'"
        ElseIf optTipoDoc(2).Value = True Then
        tpDoc = ""
    End If
    
    sSQL = "SELECT * FROM FinanceiroContasPRCadastro " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND Vencimento >= '" & Format(dtpDe.Value, "YYYY-MM-DD") & "' AND Vencimento <= '" & Format(dtpAte.Value, "YYYY-MM-DD") & "'" & _
           tpDoc & _
           " AND DataQuitacao IS NULL ORDER BY Vencimento"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum titulo a ser baixado!", vbInformation, "Aviso"
        Else
            Rst.MoveLast
            tReg = Rst.RecordCount
            cReg = 0
            Rst.MoveFirst
            
            Do Until Rst.EOF
                cReg = cReg + 1
                Status (tReg)
                If chkConfirmacao.Value = 1 Then
                        If MsgBox("Baixar o titulo: " & vbCrLf & vbCrLf & _
                                  "Sacado/Cedente: " & Rst.Fields("Nome") & vbCrLf & _
                                  "Titulo: " & Rst.Fields("NumDuplicata") & vbCrLf & _
                                  "Valor Nominal: " & ConvMoeda(Rst.Fields("vlDuplicata")) & vbCrLf & _
                                  "Vencimento: " & Rst.Fields("Vencimento"), vbInformation + vbYesNo, "Baixa Automatica - Registro:" & cReg & "/" & tReg) = vbYes Then
                                        
                                        vDados(0) = Array("DataQuitacao", dtpQuitacao.Value, "D")
                                        vDados(1) = Array("IdContaQuitacao", IdConta, "N")

                                        RegistroAlterar "FinanceiroContasPRCadastro", vDados, 1, "id = " & Rst.Fields("Id")
                                        'Movimenta conta
                                        MovimentarConta IdConta, IIf(Rst.Fields("ContaPR") = "R", "C", "D"), Rst.Fields("ID"), dtpQuitacao.Value, Rst.Fields("numDuplicata"), _
                                                        Rst.Fields("tpDocumento"), Rst.Fields("Nome"), Rst.Fields("VlDuplicata")
                        End If
                    Else
                        vDados(0) = Array("DataQuitacao", dtpQuitacao.Value, "D")
                        vDados(1) = Array("IdContaQuitacao", IdConta, "N")
                        RegistroAlterar "FinanceiroContasPRCadastro", vDados, 1, "id = " & Rst.Fields("Id")
                        
                        'Movimenta conta
                        MovimentarConta IdConta, IIf(Rst.Fields("ContaPR") = "R", "C", "D"), Rst.Fields("ID"), dtpQuitacao.Value, Rst.Fields("numDuplicata"), _
                                                        Rst.Fields("tpDocumento"), Rst.Fields("Nome"), Rst.Fields("VlDuplicata")
                End If
                
                Rst.MoveNext
            Loop
            'MsgBox "Titulos baixados com sucesso!", vbInformation, "Aviso"
    End If
    
End Sub
Private Sub Status(Max As Long)
    pb.Min = 0
    pb.Max = Max
    DoEvents
    pb.Value = pb.Value + 1
    If pb.Value > 0 And pb.Value < Max Then
            'pb.Visible = True
            Me.Enabled = False
        Else
            'pb.Visible = False
            pb.Value = 0
            Me.Enabled = True
    End If
End Sub

