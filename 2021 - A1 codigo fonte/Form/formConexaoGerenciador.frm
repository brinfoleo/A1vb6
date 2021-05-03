VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formConexaoGerenciador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de Conexão"
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
      AllowUserResizing=   1
      FormatString    =   $"formConexaoGerenciador.frx":0000
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
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Manutenção da Tabela"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkAtualizar 
         Caption         =   "Atualizar lista"
         Height          =   195
         Left            =   840
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
            NumListImages   =   16
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":00C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":0512
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":082C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":10BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":2310
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":2BEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":347C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":3D0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":4F60
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":527A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":5594
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":598B
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":713D
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":76D7
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":7DD1
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formConexaoGerenciador.frx":84CB
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formConexaoGerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Atualizar()
    On Error Resume Next
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    msfgConec.Rows = 1
    
    sSQL = "SELECT * FROM ConexaoGerenciador WHERE ID_Empresa = " & ID_Empresa & " ORDER BY IP"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                With msfgConec
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = Rst.Fields("id")
                .TextMatrix(.Rows - 1, 1) = Rst.Fields("IdPrg")
                .TextMatrix(.Rows - 1, 2) = Rst.Fields("Nome")
                .TextMatrix(.Rows - 1, 3) = Rst.Fields("IP")
                .TextMatrix(.Rows - 1, 4) = Rst.Fields("Usuario")
                .TextMatrix(.Rows - 1, 5) = Rst.Fields("Data")
                .TextMatrix(.Rows - 1, 6) = Rst.Fields("Hora")
                .TextMatrix(.Rows - 1, 7) = Rst.Fields("Status")
                End With
                Rst.MoveNext
            Loop
    End If
    Rst.Close
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


Private Sub Form_Load()
    Atualizar
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualizar"
            Atualizar
        Case "Manutenção da Tabela"
            MontarBaseDeDados
    End Select

End Sub


Private Sub Timer1_Timer()
    If chkAtualizar.Value = 1 Then Atualizar
End Sub
