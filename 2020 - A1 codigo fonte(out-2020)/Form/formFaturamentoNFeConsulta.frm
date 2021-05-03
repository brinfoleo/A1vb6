VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form formFaturamentoNFeConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Consulta NFe"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3570
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   3315
      Begin VB.TextBox txtnNF 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero da NFe:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3570
      _ExtentX        =   6297
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
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2400
         Top             =   -60
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
               Picture         =   "formFaturamentoNFeConsulta.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeConsulta.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeConsulta.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeConsulta.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeConsulta.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeConsulta.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeConsulta.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeConsulta.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeConsulta.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeConsulta.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeConsulta.frx":54D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeConsulta.frx":58CB
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFaturamentoNFeConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    LimpaFormulario Me
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Select Case tbMenu.Buttons(Button.Index).ToolTipText
    '    Case "Atualizar"
            BuscarNFe
    '    Case "Manutenção da Tabela"
    '        ManutencaoTabela
    'End Select
End Sub
Private Sub BuscarNFe()
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim nNF     As String
    
    If Trim(txtnNF.Text) = "" Then Exit Sub
    
    nNF = Left(String(9, "0"), 9 - Len(Trim(txtnNF.Text))) & Trim(txtnNF.Text)
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ide_nNF ='" & nNF & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Or Rst Is Nothing Then
            MsgBox "Nota Fiscal não localizada!", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            '**************************************************
            '* Filtra se outros usuarios podem ver as vendas de outro vendedor
            '*
            '* Checa se e super usuario
            If PgDadosUsuario(ID_Usuario).SuperUsuario = 0 Then
                If PgDadosConfig.VisualizarOutrosFunc = 0 Then
                    If CInt(Rst.fields("ger_Vendedor")) <> CInt(Left(PgDadosUsuario(ID_Usuario).Nome, 3)) Then
                        MsgBox "Somente sera permitido visualizar os dados da NF-e pelo seu emissor (" & PgDadosRhFuncionario(Rst.fields("ger_Vendedor")).Nome & ")!", vbInformation, "Aviso"
                        Rst.Close
                        Exit Sub
                    End If
                End If
            End If
    
    
            
            
            
            
            '*****************************************
            ImprimirDANFE2 (Rst.fields("idNFe"))
    End If
    Rst.Close
End Sub

Private Sub txtnNF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        BuscarNFe
    End If
    KeyAscii = IIf(IsNumeric(Chr(KeyAscii)), KeyAscii, 0)
End Sub
