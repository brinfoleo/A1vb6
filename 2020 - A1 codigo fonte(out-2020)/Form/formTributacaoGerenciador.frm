VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formTributacaoGerenciador 
   Caption         =   "Tributos"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   15120
   Begin VB.Frame Frame2 
      Caption         =   "Listar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   13
      Top             =   480
      Width           =   3975
      Begin VB.OptionButton optList 
         Caption         =   "Saida"
         Height          =   195
         Index           =   2
         Left            =   2220
         TabIndex        =   16
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optList 
         Caption         =   "Entrada"
         Height          =   195
         Index           =   1
         Left            =   1140
         TabIndex        =   15
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optList 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame frmRodape 
      Height          =   1035
      Left            =   180
      TabIndex        =   7
      Top             =   4320
      Width           =   11535
      Begin VB.Frame Frame1 
         Caption         =   "Legenda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   180
         TabIndex        =   8
         Top             =   180
         Width           =   6615
         Begin VB.Label Label3 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FF0000&
            Height          =   255
            Left            =   3300
            TabIndex        =   11
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Debito (Notas Fiscais de Venda)"
            Height          =   195
            Left            =   840
            TabIndex        =   10
            Top             =   330
            Width           =   2355
         End
         Begin VB.Label Label6 
            Caption         =   "Credito (Notas Fiscais de Compra)"
            Height          =   195
            Left            =   4080
            TabIndex        =   9
            Top             =   330
            Width           =   2415
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msfgTributos 
      Height          =   2835
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   5001
      _Version        =   393216
      Cols            =   15
      AllowUserResizing=   1
      FormatString    =   $"formTributacaoGerenciador.frx":0000
   End
   Begin VB.Frame frmPeriodo 
      Caption         =   "Periodo de Emissão:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      Begin MSComCtl2.DTPicker dtpDtInicio 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   112394241
         CurrentDate     =   40557
      End
      Begin MSComCtl2.DTPicker dtpDtFinal 
         Height          =   315
         Left            =   2460
         TabIndex        =   2
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   112394241
         CurrentDate     =   40557
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Até:"
         Height          =   195
         Left            =   2100
         TabIndex        =   4
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "De:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Registros de Saida"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Registros de Entrada"
            ImageIndex      =   9
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
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":0130
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":0582
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":089C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":112E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":2380
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":2C5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":34EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":3D7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":4FD0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":52EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":5604
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":59FB
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoGerenciador.frx":71AD
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formTributacaoGerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nNF As String

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    msfgTributos.Width = Me.Width - 400
    msfgTributos.Height = Me.Height - (msfgTributos.Top + frmRodape.Height + 600)
    frmRodape.Top = msfgTributos.Top + msfgTributos.Height
    frmRodape.Width = msfgTributos.Width
End Sub

Private Sub msfgTributos_Click()
    If msfgTributos.TextMatrix(msfgTributos.Row, 3) = "Num.Nota" Then Exit Sub
    If msfgTributos.Rows = 1 Then Exit Sub
    nNF = msfgTributos.TextMatrix(msfgTributos.Row, 3)
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualizar"
            AtualizarLista
        Case "Imprimir Registros de Saida"
            ImprimirSaida
        Case "Imprimir Registros de Entrada"
            ImprimirEntrada
'        Case "Imprimir DANFe"
'            ImprimirDANFe
    End Select
End Sub
Private Sub Form_Load()
    dtpDtInicio.Value = Date
    dtpDtFinal.Value = Date
    
End Sub

Private Sub AtualizarLista()
    msfgTributos.Rows = 1
    If optList(0).Value = True Then
            LstNotasFiscaisSaida
            LstNotasFiscaisEntrada
        ElseIf optList(1).Value = True Then
            
            LstNotasFiscaisEntrada
        ElseIf optList(2).Value = True Then
            LstNotasFiscaisSaida
    End If
            
End Sub
Private Sub ImprimirEntrada()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    'Nominal
    Dim totNF       As String
    Dim vBC         As String
    Dim vICMS       As String
    Dim vBCST       As String
    Dim vICMSST     As String
    Dim vIPI        As String
    Dim vPIS        As String
    Dim vCOFINS     As String
    
    'Deducoes
    Dim DtotNF       As String
    Dim DvBC         As String
    Dim DvICMS       As String
    Dim DvBCST       As String
    Dim DvICMSST     As String
    Dim DvIPI        As String
    Dim DvPIS        As String
    Dim DvCOFINS     As String
    
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If
    
     sSQL = MntSQLListEntrada
     
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            
        Else
            Rst.MoveFirst
            Set rptListaTributosSaida.DataSource = Rst.DataSource
            '******************************************
            If Rst.BOF And Rst.EOF Then
                Else
                    Rst.MoveFirst
                    Do Until Rst.EOF
                    '******************************************************************
                    totNF = ChkVal(Val(totNF) + Val(ChkVal(Rst.Fields("vNF"), 0, cDecMoeda)), 0, cDecMoeda)
                    vBC = ChkVal(Val(vBC) + Val(ChkVal(Rst.Fields("vBC"), 0, cDecMoeda)), 0, cDecMoeda)
                    
                    'vICMS = ChkVal(Val(vICMS) + Val(ChkVal(Rst.fields("vICMS"), 0, cDecMoeda)), 0, cDecMoeda)
                    vICMS = Val(ChkVal(vICMS, 0, cDecMoeda)) + Val(ChkVal(cNull(Rst.Fields("vICMS")), 0, cDecMoeda)) + Val(ChkVal(cNull(Rst.Fields("vCredICMSSN")), 0, cDecMoeda))
                    If cNull(Rst.Fields("retICMSST")) = "1" Then
                        DvICMS = ChkVal(Val(DvICMS) + Val(ChkVal(Rst.Fields("vICMS"), 0, cDecMoeda)), 0, cDecMoeda)
                    End If
                    
                    vBCST = ChkVal(Val(vBCST) + Val(ChkVal(cNull(Rst.Fields("vBCST")), 0, cDecMoeda)), 0, cDecMoeda)
                    vICMSST = ChkVal(Val(vICMSST) + Val(ChkVal(cNull(Rst.Fields("vICMSST")), 0, cDecMoeda)), 0, cDecMoeda)
                    vIPI = ChkVal(Val(vIPI) + Val(ChkVal(Rst.Fields("vIPI"), 0, cDecMoeda)), 0, cDecMoeda)
                    vPIS = ChkVal(Val(vPIS) + Val(ChkVal(Rst.Fields("vPIS"), 0, cDecMoeda)), 0, cDecMoeda)
                    vCOFINS = ChkVal(Val(vCOFINS) + Val(ChkVal(Rst.Fields("vCOFINS"), 0, cDecMoeda)), 0, cDecMoeda)
                    '*********************************************************************
                    '    totNF = ChkVal(Val(totNF) + Val(Rst.Fields("vNF")), 0, cDecMoeda)
                    '    vBC = ChkVal(Val(vBC) + Val(Rst.Fields("vBC")), 0, cDecMoeda)
                    '    vICMS = ChkVal(Val(vICMS) + Val(Rst.Fields("vICMS")), 0, cDecMoeda)
                    '    vBCST = ChkVal(Val(vBCST) + Val(IIf(IsNull(Rst.Fields("vBCST")), 0, Rst.Fields("vBCST"))), 0, cDecMoeda)
                    '    vICMSST = ChkVal(Val(vICMSST) + Val(IIf(IsNull(Rst.Fields("vICMSST")), 0, Rst.Fields("vICMSST"))), 0, cDecMoeda)
                    '    vIPI = ChkVal(Val(vIPI) + Val(Rst.Fields("vIPI")), 0, cDecMoeda)
                    '    vPIS = ChkVal(Val(vPIS) + Val(Rst.Fields("vPIS")), 0, cDecMoeda)
                    '    vCOFINS = ChkVal(Val(vCOFINS) + Val(Rst.Fields("vCOFINS")), 0, cDecMoeda)
                        Rst.MoveNext
                    Loop
            End If
            '***********************************************
            
            rptListaTributosSaida.Sections("Section2").Controls("lblTitulo").Caption = "Registro de ENTRADA"
            rptListaTributosSaida.Sections("Section2").Controls("lblPeriodo").Caption = "Perido: " & dtpDtInicio.Value & " até " & dtpDtFinal.Value
            
            'Valores Nominais
            rptListaTributosSaida.Sections("Section5").Controls("lblnf").Caption = totNF
            rptListaTributosSaida.Sections("Section5").Controls("lblBC").Caption = vBC
            rptListaTributosSaida.Sections("Section5").Controls("lblvICMS").Caption = vICMS
            rptListaTributosSaida.Sections("Section5").Controls("lblBCST").Caption = vBCST
            rptListaTributosSaida.Sections("Section5").Controls("lblvICMSST").Caption = vICMSST
            rptListaTributosSaida.Sections("Section5").Controls("lblvIPI").Caption = vIPI
            rptListaTributosSaida.Sections("Section5").Controls("lblvPIS").Caption = vPIS
            rptListaTributosSaida.Sections("Section5").Controls("lblvCOFINS").Caption = vCOFINS
            
            'Deducoes
            rptListaTributosSaida.Sections("Section5").Controls("lblDnf").Caption = ChkVal(DtotNF, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDBC").Caption = ChkVal(DvBC, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDvICMS").Caption = ChkVal(DvICMS, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDBCST").Caption = ChkVal(DvBCST, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDvICMSST").Caption = ChkVal(DvICMSST, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDvIPI").Caption = ChkVal(DvIPI, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDvPIS").Caption = ChkVal(DvPIS, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDvCOFINS").Caption = ChkVal(DvCOFINS, 0, cDecMoeda)
            
            'Valor Final
            rptListaTributosSaida.Sections("Section5").Controls("lblFnf").Caption = ChkVal(Val(totNF) - Val(DtotNF), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFBC").Caption = ChkVal(Val(vBC) - Val(DvBC), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFvICMS").Caption = ChkVal(Val(vICMS) - Val(DvICMS), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFBCST").Caption = ChkVal(Val(vBCST) - Val(DvBCST), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFvICMSST").Caption = ChkVal(Val(vICMSST) - Val(DvICMSST), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFvIPI").Caption = ChkVal(Val(vIPI) - Val(DvIPI), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFvPIS").Caption = ChkVal(Val(vPIS) - Val(DvPIS), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFvCOFINS").Caption = ChkVal(Val(vCOFINS) - Val(DvCOFINS), 0, cDecMoeda)
            
            rptListaTributosSaida.Show 1
    End If
End Sub

Private Sub ImprimirSaida()
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    Dim totNF       As String
    Dim vBC         As String
    Dim vICMS       As String
    Dim vBCST       As String
    Dim vICMSST     As String
    Dim vIPI        As String
    Dim vPIS        As String
    Dim vCOFINS     As String
    
    'Deducoes
    Dim DtotNF       As String
    Dim DvBC         As String
    Dim DvICMS       As String
    Dim DvBCST       As String
    Dim DvICMSST     As String
    Dim DvIPI        As String
    Dim DvPIS        As String
    Dim DvCOFINS     As String
    
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If
    
    
    'sSQL = "SELECT " & _
           "ide_dEmi, ide_nNF, dest_uf, " & _
           "total_vNF ," & _
           " total_vBC, total_vICMS, total_vIPI,total_vBCST, total_vICMSST, total_vPIS, total_vCOFINS " & _
           " FROM FaturamentoNFe WHERE ide_dEmi >='" & Format(dtpDtInicio.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & "' AND canc_nProt IS NULL ORDER BY ide_nNF"
    sSQL = MntSQLListSaida
    'sSQL = "SELECT " & _
           "ide_dEmi, ide_tpNF, ide_nNF, dest_uf as uf, " & _
           "IF(canc_nProt IS NULL,total_vNF,'cancelada')AS vNF , " & _
           "IF(canc_nProt IS NULL,total_vBC,'canc.')AS vBC, " & _
           "IF(canc_nProt IS NULL,total_vICMS,'canc.')AS vICMS, " & _
           "IF(canc_nProt IS NULL,total_vIPI,'canc.')AS vIPI, " & _
           "IF(canc_nProt IS NULL,total_vBCST,'canc.')AS vBCST, " & _
           "IF(canc_nProt IS NULL,total_vICMSST,'canc.')AS vICMSST, " & _
           "IF(canc_nProt IS NULL,total_vPIS,'canc.')AS vPIS, " & _
           "IF(canc_nProt IS NULL,total_vCOFINS,'canc.')AS vCOFINS " & _
           " FROM FaturamentoNFe " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpDtInicio.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & _
           "' ORDER BY ide_nNF"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            
        Else
            Rst.MoveFirst
            Set rptListaTributosSaida.DataSource = Rst.DataSource
            '******************************************
            If Rst.BOF And Rst.EOF Then
                Else
                    Rst.MoveFirst
                    Do Until Rst.EOF
                        totNF = ChkVal(Val(totNF) + Val(ChkVal(Rst.Fields("vNF"), 0, cDecMoeda)), 0, cDecMoeda)
                        vBC = ChkVal(Val(vBC) + Val(ChkVal(Rst.Fields("vBC"), 0, cDecMoeda)), 0, cDecMoeda)
                        vICMS = ChkVal(Val(vICMS) + Val(ChkVal(Rst.Fields("vICMS"), 0, cDecMoeda)), 0, cDecMoeda)
                        vBCST = ChkVal(Val(vBCST) + Val(ChkVal(Rst.Fields("vBCST"), 0, cDecMoeda)), 0, cDecMoeda)
                        vICMSST = ChkVal(Val(vICMSST) + Val(ChkVal(Rst.Fields("vICMSST"), 0, cDecMoeda)), 0, cDecMoeda)
                        vIPI = ChkVal(Val(vIPI) + Val(ChkVal(Rst.Fields("vIPI"), 0, cDecMoeda)), 0, cDecMoeda)
                        vPIS = ChkVal(Val(vPIS) + Val(ChkVal(Rst.Fields("vPIS"), 0, cDecMoeda)), 0, cDecMoeda)
                        vCOFINS = ChkVal(Val(vCOFINS) + Val(ChkVal(Rst.Fields("vCOFINS"), 0, cDecMoeda)), 0, cDecMoeda)
                        Rst.MoveNext
                    Loop
            End If
            '***********************************************
            rptListaTributosSaida.Sections("Section2").Controls("lblTitulo").Caption = "Registro de SAIDA"
            rptListaTributosSaida.Sections("Section2").Controls("lblPeriodo").Caption = "Perido: " & dtpDtInicio.Value & " até " & dtpDtFinal.Value
            
            rptListaTributosSaida.Sections("Section5").Controls("lblnf").Caption = totNF
            rptListaTributosSaida.Sections("Section5").Controls("lblBC").Caption = vBC
            rptListaTributosSaida.Sections("Section5").Controls("lblvICMS").Caption = vICMS
            rptListaTributosSaida.Sections("Section5").Controls("lblBCST").Caption = vBCST
            rptListaTributosSaida.Sections("Section5").Controls("lblvICMSST").Caption = vICMSST
            rptListaTributosSaida.Sections("Section5").Controls("lblvIPI").Caption = vIPI
            rptListaTributosSaida.Sections("Section5").Controls("lblvPIS").Caption = vPIS
            rptListaTributosSaida.Sections("Section5").Controls("lblvCOFINS").Caption = vCOFINS
            
            'Deducoes
            rptListaTributosSaida.Sections("Section5").Controls("lblDnf").Caption = ChkVal(DtotNF, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDBC").Caption = ChkVal(DvBC, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDvICMS").Caption = ChkVal(DvICMS, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDBCST").Caption = ChkVal(DvBCST, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDvICMSST").Caption = ChkVal(DvICMSST, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDvIPI").Caption = ChkVal(DvIPI, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDvPIS").Caption = ChkVal(DvPIS, 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblDvCOFINS").Caption = ChkVal(DvCOFINS, 0, cDecMoeda)
            
            'Valor Final
            rptListaTributosSaida.Sections("Section5").Controls("lblFnf").Caption = ChkVal(Val(totNF) - Val(DtotNF), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFBC").Caption = ChkVal(Val(vBC) - Val(DvBC), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFvICMS").Caption = ChkVal(Val(vICMS) - Val(DvICMS), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFBCST").Caption = ChkVal(Val(vBCST) - Val(DvBCST), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFvICMSST").Caption = ChkVal(Val(vICMSST) - Val(DvICMSST), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFvIPI").Caption = ChkVal(Val(vIPI) - Val(DvIPI), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFvPIS").Caption = ChkVal(Val(vPIS) - Val(DvPIS), 0, cDecMoeda)
            rptListaTributosSaida.Sections("Section5").Controls("lblFvCOFINS").Caption = ChkVal(Val(vCOFINS) - Val(DvCOFINS), 0, cDecMoeda)
            
            rptListaTributosSaida.Show 1
    End If
End Sub
Private Sub LstNotasFiscaisSaida()
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    Dim totNF       As String
    Dim vBC         As String
    Dim vICMS       As String
    Dim vBCST       As String
    Dim vICMSST     As String
    Dim vIPI        As String
    Dim vPIS        As String
    Dim vCOFINS     As String
    Dim total_vFCP  As String
    
  
    'Exibe NFe na tela
    'msfgTributos.Rows = 1
    
    sSQL = MntSQLListSaida
    'sSQL = "SELECT * FROM FaturamentoNFe" & _
           " WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpDtInicio.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & _
           "' AND canc_nProt IS NULL"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            With msfgTributos
                Do Until Rst.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("Id")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("ide_dEmi")
                    .TextMatrix(.Rows - 1, 2) = cNull(Rst.Fields("ide_tpNF"))
                    .TextMatrix(.Rows - 1, 3) = Rst.Fields("ide_Serie")
                    .TextMatrix(.Rows - 1, 4) = Rst.Fields("Ide_nNF")
                    .TextMatrix(.Rows - 1, 5) = cNull(Rst.Fields("UF"))
                    .TextMatrix(.Rows - 1, 6) = ChkVal(Rst.Fields("vNF"), 0, cDecMoeda): totNF = ChkVal(Val(totNF) + Val(ChkVal(Rst.Fields("vNF"), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 7) = ChkVal(Rst.Fields("vBC"), 0, cDecMoeda): vBC = ChkVal(Val(vBC) + Val(ChkVal(Rst.Fields("vBC"), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 8) = ChkVal(Rst.Fields("vICMS"), 0, cDecMoeda): vICMS = ChkVal(Val(vICMS) + Val(ChkVal(Rst.Fields("vICMS"), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 9) = ChkVal(Rst.Fields("vBCST"), 0, cDecMoeda): vBCST = ChkVal(Val(vBCST) + Val(ChkVal(Rst.Fields("vBCST"), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 10) = ChkVal(Rst.Fields("vICMSST"), 0, cDecMoeda): vICMSST = ChkVal(Val(vICMSST) + Val(ChkVal(Rst.Fields("vICMSST"), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 11) = ChkVal(Rst.Fields("vIPI"), 0, cDecMoeda): vIPI = ChkVal(Val(vIPI) + Val(ChkVal(Rst.Fields("vIPI"), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 12) = ChkVal(Rst.Fields("vPIS"), 0, cDecMoeda): vPIS = ChkVal(Val(vPIS) + Val(ChkVal(Rst.Fields("vPIS"), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 13) = ChkVal(Rst.Fields("vCOFINS"), 0, cDecMoeda): vCOFINS = ChkVal(Val(vCOFINS) + Val(ChkVal(Rst.Fields("vCOFINS"), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 14) = ChkVal(cNull(Rst.Fields("total_vFCP")), 0, cDecMoeda): total_vFCP = ChkVal(Val(total_vFCP) + Val(ChkVal(cNull(Rst.Fields("total_vFCP")), 0, cDecMoeda)), 0, cDecMoeda)
                    Rst.MoveNext
                Loop
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 6) = ConvMoeda(totNF)
                .TextMatrix(.Rows - 1, 7) = ConvMoeda(vBC)
                .TextMatrix(.Rows - 1, 8) = ConvMoeda(vICMS)
                .TextMatrix(.Rows - 1, 9) = ConvMoeda(vBCST)
                .TextMatrix(.Rows - 1, 10) = ConvMoeda(vICMSST)
                .TextMatrix(.Rows - 1, 11) = ConvMoeda(vIPI)
                .TextMatrix(.Rows - 1, 12) = ConvMoeda(vPIS)
                .TextMatrix(.Rows - 1, 13) = ConvMoeda(vCOFINS)
                .TextMatrix(.Rows - 1, 14) = ConvMoeda(total_vFCP)
                .FillStyle = flexFillRepeat
                .Row = .Rows - 1
                .Col = 0
                .ColSel = .Cols - 1
                .CellFontBold = True
                .CellForeColor = vbRed
            End With
    End If
End Sub
Private Sub LstNotasFiscaisEntrada()
    Dim Rst         As Recordset
    Dim sSQL        As String
    
    Dim totNF       As String
    Dim vBC         As String
    Dim vICMS       As String
    Dim vBCST       As String
    Dim vICMSST     As String
    Dim vIPI        As String
    Dim vPIS        As String
    Dim vCOFINS     As String
    
    'Deducoes
    Dim DtotNF       As String
    Dim DvBC         As String
    Dim DvICMS       As String
    Dim DvBCST       As String
    Dim DvICMSST     As String
    Dim DvIPI        As String
    Dim DvPIS        As String
    Dim DvCOFINS     As String
  
    'Exibe NFe na tela
    'msfgTributos.Rows = 1
    
    
    sSQL = MntSQLListEntrada
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            With msfgTributos
                Do Until Rst.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("Id")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("ide_dEmi")
                    .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rst.Fields("Ide_tpNF")), "", Rst.Fields("Ide_tpNF"))
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Ide_Serie")), "", Rst.Fields("Ide_Serie"))
                    .TextMatrix(.Rows - 1, 4) = Rst.Fields("Ide_nNF")
                    .TextMatrix(.Rows - 1, 5) = Rst.Fields("UF")
                    .TextMatrix(.Rows - 1, 6) = ChkVal(Rst.Fields("vNF"), 0, cDecMoeda): totNF = ChkVal(Val(totNF) + Val(ChkVal(Rst.Fields("vNF"), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 7) = ChkVal(Rst.Fields("vBC"), 0, cDecMoeda): vBC = ChkVal(Val(vBC) + Val(ChkVal(Rst.Fields("vBC"), 0, cDecMoeda)), 0, cDecMoeda)
                    
                    vICMS = Val(ChkVal(vICMS, 0, cDecMoeda)) + Val(ChkVal(cNull(Rst.Fields("vICMS")), 0, cDecMoeda)) '+ Val(ChkVal(cNull(Rst.fields("vCredICMSSN")), 0, cDecMoeda))
                    If cNull(Rst.Fields("retICMSST")) = "1" Then
                        DvICMS = Val(ChkVal(DvICMS, 0, cDecMoeda)) + Val(ChkVal(cNull(Rst.Fields("vICMS")), 0, cDecMoeda))
                        'DvICMS = Val(ChkVal(DvICMS, 0, cDecMoeda))
                        'vICMS = ChkVal(Val(vICMS) - Val(ChkVal(cNull(Rst.fields("vICMS")), 0, cDecMoeda)), 0, cDecMoeda)
                    End If
                    .TextMatrix(.Rows - 1, 8) = ChkVal(Val(ChkVal(cNull(Rst.Fields("vICMS")), 0, cDecMoeda)) + Val(ChkVal(cNull(Rst.Fields("vCredICMSSN")), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 9) = ChkVal(IIf(IsNull(Rst.Fields("vBCST")), "0.00", Rst.Fields("vBCST")), 0, cDecMoeda): vBCST = ChkVal(Val(vBCST) + Val(ChkVal(cNull(Rst.Fields("vBCST")), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 10) = ChkVal(IIf(IsNull(Rst.Fields("vICMSST")), "0.00", Rst.Fields("vICMSST")), 0, cDecMoeda): vICMSST = ChkVal(Val(vICMSST) + Val(ChkVal(cNull(Rst.Fields("vICMSST")), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 11) = ChkVal(Rst.Fields("vIPI"), 0, cDecMoeda): vIPI = ChkVal(Val(vIPI) + Val(ChkVal(Rst.Fields("vIPI"), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 12) = ChkVal(Rst.Fields("vPIS"), 0, cDecMoeda): vPIS = ChkVal(Val(vPIS) + Val(ChkVal(Rst.Fields("vPIS"), 0, cDecMoeda)), 0, cDecMoeda)
                    .TextMatrix(.Rows - 1, 13) = ChkVal(Rst.Fields("vCOFINS"), 0, cDecMoeda): vCOFINS = ChkVal(Val(vCOFINS) + Val(ChkVal(Rst.Fields("vCOFINS"), 0, cDecMoeda)), 0, cDecMoeda)
                    Rst.MoveNext
                Loop
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 6) = ConvMoeda(totNF)
                .TextMatrix(.Rows - 1, 7) = ConvMoeda(vBC)
                .TextMatrix(.Rows - 1, 8) = ConvMoeda(vICMS)
                .TextMatrix(.Rows - 1, 9) = ConvMoeda(vBCST)
                .TextMatrix(.Rows - 1, 10) = ConvMoeda(vICMSST)
                .TextMatrix(.Rows - 1, 11) = ConvMoeda(vIPI)
                .TextMatrix(.Rows - 1, 12) = ConvMoeda(vPIS)
                .TextMatrix(.Rows - 1, 13) = ConvMoeda(vCOFINS)
                .FillStyle = flexFillRepeat
                .Row = .Rows - 1
                .Col = 0
                .ColSel = .Cols - 1
                .CellFontBold = True
                .CellForeColor = vbBlue
            End With
    End If
End Sub

Private Function MntSQLListEntrada() As String
    '##########################################################
    '### Monta a Consulta em SQL para listagem de entradas
    '##########################################################
    Dim sSQL As String
    
    'sSQL = "SELECT " & _
           "id," & _
           "MovFisco," & _
           "ide_dEmi, ide_tpNF, ide_nNF, emit_uf as uf,ide_Serie, " & _
           "total_vNF * 1 AS vNF , " & _
           "total_vBC * 1 AS vBC, " & _
           "total_vICMS * 1 AS vICMS, " & _
           "total_vCredICMSSN * 1 AS vICMSSN, " & _
           "total_vIPI * 1 AS vIPI, " & _
           "total_vBCST * 1 AS vBCST, " & _
           "total_vICMSST * 1 AS vICMSST, " & _
           "total_vPIS * 1 AS vPIS, " & _
           "total_vCOFINS * 1 AS vCOFINS " & _
           " FROM FaturamentoNFeEntrada " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpDtInicio.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & _
           "' AND MovFisco=1 " & _
           "ORDER BY ide_dEmi, ide_nNF"
    sSQL = "SELECT " & _
           "id," & _
           "MovFisco, " & _
           "retICMSST," & _
           "ide_dEmi, ide_tpNF, ide_nNF, emit_uf as uf,ide_Serie, " & _
           "total_vNF * 1 AS vNF , " & _
           "total_vBC * 1 AS vBC, " & _
           "((total_vICMS * 1) + (total_vCredICMSSN * 1)) AS vICMS, " & _
           "total_vIPI * 1 AS vIPI, " & _
           "total_vBCST * 1 AS vBCST, " & _
           "total_vICMSST * 1 AS vICMSST, " & _
           "total_vPIS * 1 AS vPIS, " & _
           "total_vCOFINS * 1 AS vCOFINS, " & _
           "total_vCredICMSSN AS vCredICMSSN " & _
           " FROM FaturamentoNFeEntrada " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpDtInicio.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & _
           "' AND MovFisco=1 " & _
           "ORDER BY ide_dEmi, ide_nNF"
    MntSQLListEntrada = sSQL
End Function
Private Function MntSQLListSaida() As String
    Dim sSQL As String
           
    sSQL = "SELECT " & _
           "id, ide_dEmi, ide_tpNF, ide_nNF, dest_uf as uf, ide_Serie, MovFisco, total_vFCP, " & _
           "IF(canc_nProt IS NULL,IF(ide_tpNF<>1,- total_vNF, total_vNF*1),0) AS vNF , " & _
           "IF(canc_nProt IS NULL,IF(ide_tpNF<>1,- total_vBC,total_vBC*1),0) AS vBC, " & _
           "IF(canc_nProt IS NULL,IF(ide_tpNF<>1,- total_vICMS, total_vICMS*1),0) AS vICMS, " & _
           "IF(canc_nProt IS NULL,IF(ide_tpNF<>1,- total_vIPI,total_vIPI*1),0) AS vIPI, " & _
           "IF(canc_nProt IS NULL,IF(ide_tpNF<>1,- total_vBCST, total_vBCST*1),0) AS vBCST, " & _
           "IF(canc_nProt IS NULL,IF(ide_tpNF<>1,- total_vICMSST, total_vICMSST*1),0) AS vICMSST, " & _
           "IF(canc_nProt IS NULL,IF(ide_tpNF<>1,- total_vPIS, total_vPIS*1),0) AS vPIS, " & _
           "IF(canc_nProt IS NULL,IF(ide_tpNF<>1,- total_vCOFINS, total_vCOFINS*1),0) AS vCOFINS " & _
           "FROM FaturamentoNFe " & _
           "WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_dEmi >='" & Format(dtpDtInicio.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & "' " & _
           "AND MovFisco = 1 " & _
           "ORDER BY ide_nNF"
    
    MntSQLListSaida = sSQL
    
End Function
