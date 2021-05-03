VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formRHFuncionarioFolhaPonto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RH - Folha de Ponto"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5775
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   5595
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   1155
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ano:"
         Height          =   195
         Left            =   2700
         TabIndex        =   3
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mês:"
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   555
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
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
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4680
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":076C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":0FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":2250
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":2B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":33BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":3C4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":4EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":51BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":54D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":58CB
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":707D
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formRHFuncionarioFolhaPonto.frx":7617
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formRHFuncionarioFolhaPonto"
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
 Dim ano As Integer
    ano = 2010
    For ano = 2010 To 2030
        cboAno.AddItem ano
    Next
    With cboMes
        .Clear
        .AddItem "JANEIRO"
        .AddItem "FEVEREIRO"
        .AddItem "MARÇO"
        .AddItem "ABRIL"
        .AddItem "MAIO"
        .AddItem "JUNHO"
        .AddItem "JULHO"
        .AddItem "AGOSTO"
        .AddItem "SETEMBRO"
        .AddItem "OUTUBRO"
        .AddItem "NOVEMBRO"
        .AddItem "DEZEMBRO"
    End With
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Imprimir"
            Imprimir
    End Select

End Sub
Private Sub Imprimir()

    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If
    
     If Trim(cboMes.Text) = "" Or Trim(cboAno.Text) = "" Then
        MsgBox "Selecione MES / ANO!", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim Dia     As Integer
    Dim Mes     As Integer
    Dim ano     As Integer
    Dim dt      As String
    Dim i       As Integer
    Dim dSemana As String
    Dim cFonte  As String
    
    ano = cboAno.Text
    
    Select Case cboMes.Text
        Case "JANEIRO"
            Dia = 31
            Mes = 1
        Case "FEVEREIRO"
            Mes = 2
            If (ano Mod 4 = 0 And ano Mod 100 <> 0) Or (ano Mod 400 = 0) Then
                Dia = 29
            Else
                Dia = 28
            End If
            
        Case "MARÇO"
            Dia = 31
            Mes = 3
        Case "ABRIL"
            Dia = 30
            Mes = 4
        Case "MAIO"
            Dia = 31
            Mes = 5
        Case "JUNHO"
            Dia = 30
            Mes = 6
        Case "JULHO"
            Dia = 31
            Mes = 7
        Case "AGOSTO"
            Dia = 31
            Mes = 8
        Case "SETEMBRO"
            Dia = 31
            Mes = 9
        Case "OUTUBRO"
            Dia = 31
            Mes = 10
        Case "NOVEMBRO"
            Dia = 30
            Mes = 11
        Case "DEZEMBRO"
            Dia = 31
            Mes = 12
    End Select
    
    
    
        
    
    sSQL = "SELECT * FROM RHFuncionarioCadastro WHERE ID_Empresa = " & ID_Empresa & " AND FolhaPonto = 1"
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma folha de ponto selecionada.", vbInformation, "Aviso"
        Else
            Set rptRHFolhaPonto.DataSource = Rst.DataSource
            '================================
            For i = 1 To 31
                dt = i & "/" & Mes & "/" & ano
                If IsDate(dt) = True Then
                        If Weekday(dt) = 1 Then
                                dSemana = "DOMINGO"
                                cFonte = vbBlack
                            ElseIf Weekday(dt) = 7 Then
                                dSemana = "SABADO"
                                cFonte = vbBlack
                            ElseIf i > Dia Then
                                dSemana = "***********"
                                cFonte = vbBlack
                            Else
                                dSemana = "."
                                cFonte = vbWhite
                        End If
                    Else
                        dSemana = "***********"
                        cFonte = vbBlack
                End If
                rptRHFolhaPonto.Sections("Section1").Controls.Item("lbl" & i).ForeColor = cFonte
                rptRHFolhaPonto.Sections("Section1").Controls.Item("lbl" & i).Caption = dSemana
                'Debug.Print Dt & "  " & Weekday(Dt)
            Next
            '====================
            rptRHFolhaPonto.Sections("Section2").Controls.Item("lblCab").Caption = "FOLHA DE PONTO - " & cboMes.Text & "/" & cboAno.Text
            rptRHFolhaPonto.Show 1
    End If
    Rst.Close
End Sub
