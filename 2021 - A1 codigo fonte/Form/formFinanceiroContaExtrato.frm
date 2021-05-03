VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formFinanceiroContaExtrato 
   Caption         =   "Financeiro - Extrato"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   14340
   Begin VB.CheckBox chkAgruparMovimento 
      Caption         =   "Agrupar Movimento"
      Height          =   255
      Left            =   12060
      TabIndex        =   14
      Top             =   720
      Width           =   1755
   End
   Begin VB.Frame frmSaldo 
      Caption         =   "Saldo Atual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10500
      TabIndex        =   11
      Top             =   6960
      Width           =   3735
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   12
         Text            =   "0,00"
         Top             =   240
         Width           =   3315
      End
   End
   Begin VB.Frame frmExtrato 
      Height          =   5475
      Left            =   120
      TabIndex        =   9
      Top             =   1380
      Width           =   14115
      Begin MSFlexGridLib.MSFlexGrid msfgExtrato 
         Height          =   5175
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   9128
         _Version        =   393216
         Cols            =   7
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"formFinanceiroContaExtrato.frx":0000
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Conta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   7335
      Begin VB.ComboBox cboConta 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   5715
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Dados da Conta:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   7560
      TabIndex        =   0
      Top             =   480
      Width           =   4275
      Begin MSComCtl2.DTPicker dtpPeriodoFinal 
         Height          =   315
         Left            =   2640
         TabIndex        =   1
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57933825
         CurrentDate     =   40525
      End
      Begin MSComCtl2.DTPicker dtpPeriodoInicio 
         Height          =   315
         Left            =   660
         TabIndex        =   2
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57933825
         CurrentDate     =   40525
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Final:"
         Height          =   195
         Left            =   2220
         TabIndex        =   3
         Top             =   360
         Width           =   435
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14340
      _ExtentX        =   25294
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
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recalcular Saldo"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   10380
         TabIndex        =   13
         Top             =   60
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
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
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":00DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":052F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":0849
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":10DB
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":232D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":2C07
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":3499
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":3D2B
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":4F7D
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":5297
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":55B1
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFinanceiroContaExtrato.frx":59A8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFinanceiroContaExtrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdReg As Integer

Private Sub cboConta_Click()
    If Trim(cboConta.Text) = "" Then Exit Sub
    IdReg = Left(Trim(cboConta.Text), 3)
    
End Sub

Private Sub cboConta_DropDown()
    Dim sSQL    As String
    Dim Rst     As Recordset
    LimpForm
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
Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub
Private Sub status(Max As Long)

    pb.min = 0
    pb.Max = Max
    DoEvents
    pb.Value = pb.Value + 1
    If pb.Value > 0 And pb.Value < Max Then
            pb.Visible = True
            Me.Enabled = False
        Else
            pb.Visible = False
            pb.Value = 0
            Me.Enabled = True
    End If
End Sub
Private Sub LimpForm()
    cboConta.Clear
    pb.Visible = False
    dtpPeriodoInicio.Value = Date
    dtpPeriodoFinal.Value = Date
    msfgExtrato.Rows = 1
    txtSaldo.Text = ConvMoeda("0,00")
End Sub

Private Sub Form_Load()
    LimpForm
End Sub
Private Sub AtualizarGrid()
    On Error Resume Next
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    If Trim(cboConta.Text) = "" Or IdReg = 0 Then Exit Sub
    With msfgExtrato
        .Rows = 1
        txtSaldo.Text = ConvMoeda(pgDadosConta(IdReg).Saldo)
        
        If chkAgruparMovimento.Value = 0 Then
                'Não Agrupa resultados
                sSQL = "SELECT id, CD, Data, Documento as xDoc, Descricao as xDesc,valor as xValor, Saldo, tpDoc" & _
                        " FROM FinanceiroContaHistorico" & _
                        " WHERE ID_Empresa = " & ID_Empresa & " AND IdConta = " & IdReg & " AND" & _
                        " Data >= '" & Format(dtpPeriodoInicio.Value, "yyyy-mm-dd") & "' AND Data <= '" & Format(dtpPeriodoFinal.Value, "yyyy-mm-dd") & "'" & _
                        " ORDER BY Data,Id"
                '.Cols = 7
                '.FormatString = .FormatString & "|>Saldo                     "
            Else
                'Agrupar Resultados
                sSQL = "SELECT id, CD, Data, Documento as xDoc, Descricao as xDesc,sum(valor) as xValor, Saldo,tpDoc" & _
                        " FROM FinanceiroContaHistorico" & _
                        " WHERE ID_Empresa = " & ID_Empresa & " AND IdConta = " & IdReg & " AND" & _
                        " Data >= '" & Format(dtpPeriodoInicio.Value, "yyyy-mm-dd") & "' AND Data <= '" & Format(dtpPeriodoFinal.Value, "yyyy-mm-dd") & "'" & _
                        " GROUP BY data,tpDoc, CD ORDER BY Data,Id,Saldo ASC"
                '.Cols = 6
                '.FormatString = "^id |^Data              |^Documento           |<Descrição                                                                         |>Credito                     |>Debito                     "
        End If
           
           
        'sSQL = "SELECT * FROM FinanceiroContaHistorico WHERE IdConta=" & IdReg & " ORDER BY ID"
        Set Rst = RegistroBuscar(sSQL)
        If Rst.BOF And Rst.EOF Then
                MsgBox "Nenhum historico encontrado!", vbInformation, "Aviso"
            Else
                Rst.MoveFirst
                Do Until Rst.EOF
                    status (Rst.RecordCount)
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = ZE(Rst.Fields("id"), 5)
                    .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rst.Fields("Data")), "", Rst.Fields("Data"))
                    If chkAgruparMovimento.Value = 0 Then
                            'Não Agrupado
                            .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rst.Fields("xDoc")), "", Rst.Fields("xDoc")) & " (" & pgDadosTipoDocumento(Rst.Fields("tpDoc")).Sigla & ")"
                            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("xDesc")), "", Rst.Fields("xDesc"))
                            .TextMatrix(.Rows - 1, 6) = ConvMoeda(ChkVal(IIf(IsNull(Rst.Fields("Saldo")), "0", Rst.Fields("Saldo")), 0, cDecMoeda))
                        Else
                            'Agrupado
                            .TextMatrix(.Rows - 1, 2) = "" 'IIf(IsNull(Rst.Fields("xDoc")), "", Rst.Fields("xDoc"))
                            .TextMatrix(.Rows - 1, 3) = pgDescrTipoDoc(IIf(IsNull(Rst.Fields("tpDoc")), "0", Rst.Fields("tpDoc")))
                    End If
                    
                    If Rst.Fields("CD") = "C" Then
                            .TextMatrix(.Rows - 1, 4) = ConvMoeda(ChkVal(IIf(IsNull(Rst.Fields("xValor")), "0", Rst.Fields("xValor")), 0, cDecMoeda))
                        Else
                            .TextMatrix(.Rows - 1, 5) = ConvMoeda(ChkVal(IIf(IsNull(Rst.Fields("xValor")), "0", Rst.Fields("xValor")), 0, cDecMoeda))
                    End If
                    
                    Rst.MoveNext
                Loop
        End If
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    frmExtrato.Width = Me.ScaleWidth - 250
    frmExtrato.Height = Me.ScaleHeight - (frmExtrato.Top + frmSaldo.Height + 200)
    
    msfgExtrato.Width = frmExtrato.Width - 250
    msfgExtrato.Height = frmExtrato.Height - 250
    
    frmSaldo.Top = frmExtrato.Height + frmExtrato.Top + 100
    frmSaldo.Left = (frmExtrato.Width + frmExtrato.Left) - frmSaldo.Width
    
     pb.Left = Me.Width - (pb.Width + 200)
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualizar"
            AtualizarGrid
        Case "Recalcular Saldo"
            RecalcularSaldo
    End Select
End Sub
Private Sub RecalcularSaldo()
    If IdReg = 0 Then
        MsgBox "Selecione uma Conta!", vbInformation, App.EXEName
        Exit Sub
    End If
    If MsgBox("Este procedimento levará alguns minutos! " & vbCrLf & "Deseja continuar?", vbQuestion + vbYesNo, App.EXEName) = vbNo Then
        Exit Sub
    End If
    
    Dim vReg(10)    As Variant
    Dim cReg        As Integer
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim vSaldo      As String
    
    
    
    
    
    sSQL = "SELECT * FROM FinanceiroContaHistorico ORDER BY Data, id"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            vSaldo = 0
        Else
            vSaldo = 0
            Rst.MoveFirst
            Do Until Rst.EOF
                status Rst.RecordCount
                If Rst.Fields("CD") = "C" Then
                        vSaldo = Val(ChkVal(vSaldo, 0, cDecMoeda)) + Val(ChkVal(Rst.Fields("Valor"), 0, cDecMoeda))
                    Else
                        vSaldo = Val(ChkVal(vSaldo, 0, cDecMoeda)) - Val(ChkVal(Rst.Fields("Valor"), 0, cDecMoeda))
                End If
                vSaldo = ChkVal(vSaldo, 0, cDecMoeda)
                
                cReg = 0
                vReg(cReg) = Array("Saldo", vSaldo, "S")
                RegistroAlterar "FinanceiroConta", vReg, cReg, "id=" & Rst.Fields("id")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    
    'Zerar o Salda Conta
    cReg = 0
    vReg(cReg) = Array("Saldo", vSaldo, "S")
    If RegistroAlterar("FinanceiroConta", vReg, cReg, "id=" & IdReg) = True Then
        MsgBox "Saldo recalculado com sucesso!", vbInformation, App.EXEName
        AtualizarGrid
    End If
    
End Sub
Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


