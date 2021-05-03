VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formUsuConexaoGerenciador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de Conexão de Usuarios"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   10845
   Begin MSFlexGridLib.MSFlexGrid msfgConec 
      Height          =   3915
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6906
      _Version        =   393216
      Cols            =   8
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"formUsuConexaoGerenciador.frx":0000
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Checar Conexão"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Manutenção da Tabela"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkAtualizar 
         Caption         =   "Atualizar lista"
         Height          =   195
         Left            =   2400
         TabIndex        =   2
         Top             =   120
         Width           =   1755
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   6300
         Top             =   0
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
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":00C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":0512
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":082C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":10BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":2310
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":2BEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":347C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":3D0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":4F60
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":527A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":5594
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":598B
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":713D
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":76D7
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":7DD1
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":84CB
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formUsuConexaoGerenciador.frx":8BC5
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formUsuConexaoGerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lin     As Integer
Dim IdReg   As Integer

Private Sub Atualizar()
    On Error Resume Next
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    chkAtualizar.Value = 0
    chkAtualizar.Enabled = False
    
    msfgConec.Rows = 1
    
    sSQL = "SELECT * FROM ConexaoGerenciador WHERE ID_Empresa = " & ID_Empresa & " ORDER BY IP"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                With msfgConec
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = Rst.fields("id")
                .TextMatrix(.Rows - 1, 1) = Rst.fields("IdPrg")
                .TextMatrix(.Rows - 1, 2) = Rst.fields("Nome")
                .TextMatrix(.Rows - 1, 3) = Rst.fields("IP")
                .TextMatrix(.Rows - 1, 4) = Rst.fields("Usuario")
                .TextMatrix(.Rows - 1, 5) = Rst.fields("Data")
                .TextMatrix(.Rows - 1, 6) = Rst.fields("Hora")
                .TextMatrix(.Rows - 1, 7) = Rst.fields("Status")
                verificarDadosStatus Rst.fields("id"), Rst.fields("Usuario"), Rst.fields("Status")
                End With
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    chkAtualizar.Enabled = True
    chkAtualizar.Value = 1
End Sub
Private Sub verificarDadosStatus(idL As Integer, usu As String, status As String)
    
    If UCase(status) = "CONECTADO" Then
        'verificarDadosStatus = "CONECTADO"
        Exit Sub
    End If
    
    Dim Hr      As String
    Dim min     As Integer
    
    If InStr(status, "CHECAR") <> 0 Then
        Hr = Trim(Mid(status, InStr(status, "-") + 1, Len(status)))
        min = DateDiff("n", Hr, Time)
        If min > 1 Then
            If MsgBox("Usuario " & usu & " não responde a " & min & " minutos." & vbCrLf & _
                   "Deseja excluir o usuario desta lista?", vbYesNo + vbQuestion, App.EXEName) = vbYes Then
                   RegistroExcluir "ConexaoGerenciador", "id=" & idL
                   Atualizar
            End If
        End If
    End If

End Sub
Private Sub MontarBaseDeDados()
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    cReg = 0
    vReg(cReg) = Array("idPrg", 100, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Nome", 50, "S"): cReg = cReg + 1
    vReg(cReg) = Array("IP", 30, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Usuario", 30, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("ID_Usuario", 30, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("Data", 50, "D"): cReg = cReg + 1
    vReg(cReg) = Array("Hora", 50, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Status", 250, "S") ': cReg = cReg + 1
    
    formManutencaoTabelas.Gerar_BD_com_Array Me, vReg, cReg
    'Dim vDados(1000)    As Variant
    'Dim contReg         As Integer
    'Dim i               As Integer
    
    'contReg = 0
    'vDados(contReg) = Array("DtHrReg", "100", "S"): contReg = contReg + 1
    'vDados(contReg) = Array("idCliente", "50", "N"): contReg = contReg + 1
    'vDados(contReg) = Array("Descricao", "5000", "S") ': contReg = contReg + 1
    
   
    
    
End Sub
Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Atualizar
End Sub

Private Sub msfgConec_Click()
    lin = msfgConec.Row
    IdReg = msfgConec.TextMatrix(msfgConec.Row, 0)
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualizar"
            Atualizar
        Case "Checar Conexão"
            ChecarConexao
            'Atualizar
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select

End Sub
Private Sub ChecarConexao()
      Dim sSQL    As String
    Dim Rst     As Recordset
    Dim vReg(1) As Variant
    Dim sTexto  As String
    

    '#
    '# Leonardo Aquino
    '# 27/09/2012
    '#
    '# Mudanca de codigo para avalir todos os usuarios conectado ao
    '# invez de um por vez.
    '# Diminuição tempo de resposta para de 6 para 1 minuto
    '#
    '#
    
    Dim l As Integer
    sSQL = "SELECT * FROM ConexaoGerenciador WHERE ID_Empresa = " & ID_Empresa
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar Usuario na tabela. Feche o aplicativo e abra novamente!", vbInformation, App.EXEName
        Else
            Rst.MoveFirst
            l = 1
            Do Until Rst.EOF
                sTexto = "CHECAR - " & Time
                msfgConec.TextMatrix(l, 7) = sTexto
                vReg(0) = Array("Status", sTexto, "S")
                RegistroAlterar "ConexaoGerenciador", vReg, 0, "id=" & Rst.fields("id")
                Rst.MoveNext
                l = IIf(msfgConec.Rows < l, l, l + 1)
                
            Loop
    End If
    Rst.Close
End Sub

Private Sub Timer1_Timer()
    If chkAtualizar.Value = 1 Then Atualizar
End Sub
