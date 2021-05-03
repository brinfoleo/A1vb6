VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formTributacaoNCM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tributação - Tabela Nomeclatura Comum no Mercosul (NCM)"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   7005
   Begin VB.Frame Frame2 
      Height          =   3795
      Left            =   60
      TabIndex        =   6
      Top             =   420
      Width           =   6855
      Begin MSFlexGridLib.MSFlexGrid msfgTabICMS 
         Height          =   3495
         Left            =   60
         TabIndex        =   7
         ToolTipText     =   "Duplo click para selecionar..."
         Top             =   180
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   4
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "^Id |<Descrição                                                            |^NCM                              |^IPI    "
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   60
      TabIndex        =   2
      Top             =   4260
      Width           =   6855
      Begin VB.TextBox txtCEST 
         Height          =   285
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox txtIPI 
         Height          =   285
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton btoExcluir 
         Height          =   495
         Left            =   5040
         Picture         =   "formTributacaoNCM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir UF/ % ICMS"
         Top             =   2820
         Width           =   1155
      End
      Begin VB.CommandButton btoIncluir 
         Height          =   495
         Left            =   3840
         Picture         =   "formTributacaoNCM.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Incluir UF/ % ICMS"
         Top             =   2820
         Width           =   1155
      End
      Begin VB.TextBox txtICMS 
         Height          =   285
         Left            =   2460
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2970
         Width           =   795
      End
      Begin VB.ComboBox cboUF 
         Height          =   315
         Left            =   540
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2940
         Width           =   915
      End
      Begin MSFlexGridLib.MSFlexGrid msfgICMS 
         Height          =   1275
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Click para selecionar..."
         Top             =   1440
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   2249
         _Version        =   393216
         Cols            =   3
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "^id |^UF               |^ICMS (%)   "
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1260
         MaxLength       =   100
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtNCM 
         Height          =   285
         Left            =   1260
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "CEST:"
         Height          =   195
         Left            =   3300
         TabIndex        =   17
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "IPI:"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "ICMS(%):"
         Height          =   195
         Left            =   1680
         TabIndex        =   10
         Top             =   3000
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "UF:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "NCM:"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   660
         Width           =   795
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
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
               Picture         =   "formTributacaoNCM.frx":0614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoNCM.frx":0A66
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoNCM.frx":0D80
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoNCM.frx":1612
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoNCM.frx":2864
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoNCM.frx":313E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoNCM.frx":39D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoNCM.frx":4262
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoNCM.frx":54B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoNCM.frx":57CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formTributacaoNCM.frx":5AE8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formTributacaoNCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IdReg       As Integer
Dim strTabela   As String
Dim IdICMS      As Integer
Private Sub MontarBaseDados()
    'formManutencaoTabelas.IniciarManutencao Me
    Dim vDados(1000)    As Variant
    Dim contReg         As Integer
    Dim i               As Integer
    
    contReg = 0
    
    vDados(contReg) = Array("Descricao", "100", "S"): contReg = contReg + 1
    vDados(contReg) = Array("NCM", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("CEST", "20", "S"): contReg = contReg + 1
    vDados(contReg) = Array("IPI", "10", "S") ': contReg = contReg + 1

    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg
    
    'TABELA NCM / ICMS *******************************************
    contReg = 0
    vDados(contReg) = Array("idNCM", "10", "N"): contReg = contReg + 1
    vDados(contReg) = Array("UF", "5", "S"): contReg = contReg + 1
    vDados(contReg) = Array("ICMS", "10", "S") ': contReg = contReg + 1
    formManutencaoTabelas.Gerar_BD_com_Array Me, vDados, contReg, "ICMS"
    
End Sub

Private Sub btoExcluir_Click()
    If msfgICMS.Rows <= 2 Then
            msfgICMS.Rows = 1
            'Exit Sub
        Else
            msfgICMS.RemoveItem msfgICMS.Row
    End If
    IdICMS = 0
End Sub

Private Sub btoIncluir_Click()
    With msfgICMS
        If IdICMS = 0 Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = "999"
                .TextMatrix(.Rows - 1, 1) = cboUF.Text
                .TextMatrix(.Rows - 1, 2) = txtICMS.Text
            Else
                .TextMatrix(.Row, 1) = cboUF.Text
                .TextMatrix(.Row, 2) = txtICMS.Text
        End If
    End With
    IdICMS = 0
End Sub

Private Sub cboUF_DropDown()
    Dim Rst As Recordset
    cboUF.Clear
    Set Rst = RegistroBuscar("SELECT * FROM TributacaoUF ORDER BY sigla")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboUF.AddItem Rst.Fields("sigla")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub PesquisarRegistro()
    IdReg = formBuscar.IniciarBusca(strTabela, "NCM,Descricao,IPI")
    
    'If idReg = 0 Then
            LimpForm
    '    Else
            MostrarDados
            MostrarGrid (IdReg)
    'End If
End Sub


Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LimpForm
    strTabela = Mid(Me.Name, 5, Len(Me.Name))
    HDForm Me, False
    HDMenu Me, True
    MostrarGrid
    msfgTabICMS.Enabled = True
End Sub

Private Sub msfgICMS_Click()
    If msfgICMS.Rows = 1 Then Exit Sub
    IdICMS = msfgICMS.TextMatrix(msfgICMS.Row, 0)
    cboUF.Clear
    cboUF.AddItem IIf(Trim(msfgICMS.TextMatrix(msfgICMS.Row, 1)) = "", " ", msfgICMS.TextMatrix(msfgICMS.Row, 1))
    cboUF.Text = cboUF.List(0)
    txtICMS.Text = msfgICMS.TextMatrix(msfgICMS.Row, 2)
End Sub

Private Sub msfgTabICMS_DblClick()
    If Not IsNumeric(msfgTabICMS.TextMatrix(msfgTabICMS.Row, 0)) Then Exit Sub
    IdReg = msfgTabICMS.TextMatrix(msfgTabICMS.Row, 0)
    MostrarDados
    
    
End Sub
Private Sub LimpForm()
    LimpaFormulario Me
    msfgICMS.Rows = 1
End Sub
Private Sub Incluir()
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    IdReg = 0
    HDMenu Me, False
    HDForm Me, True
    LimpForm
End Sub
Private Sub Alterar()
    If chkAcesso(Me, "a") = False Then
        Exit Sub
    End If
    If IdReg = 0 Then
        MsgBox "Selecione uma Registro."
        Exit Sub
    End If
    HDForm Me, True
    HDMenu Me, False
    msfgTabICMS.Enabled = False
End Sub
Private Sub Excluir()
    If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If

    If IdReg = 0 Then
            MsgBox "Selecione um Registro."
            Exit Sub
        Else
            If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                        vbCrLf & _
                        "Descrição.: " & txtDescricao.Text, vbYesNo + vbQuestion) = vbYes Then
                               
                If RegistroExcluir(strTabela, "Id = " & IdReg) = True Then
                    LimpForm
                End If
            End If
    End If
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
            PesquisarRegistro
            
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu Me, True
                HDForm Me, False
                'LimpForm
                'txtCNPJ.Enabled = True
            End If
            MostrarGrid
            msfgTabICMS.Enabled = True
        
        Case "Cancelar"
            HDMenu Me, True
            HDForm Me, False
            LimpForm
            MostrarGrid
            msfgTabICMS.Enabled = True
        Case "Manutenção da Tabela"
            MontarBaseDados
    End Select
End Sub
Private Function grvRegistro() As Boolean
    Dim vReg(199)   As Variant
    Dim cReg        As Integer 'Contador de Registros
    Dim i           As Integer
    cReg = 0
    vReg(cReg) = Array("Descricao", Trim(txtDescricao.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("NCM", Trim(txtNCM.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("cest", Trim(txtCEST.Text), "S"): cReg = cReg + 1
    vReg(cReg) = Array("IPI", Trim(txtIPI.Text), "S") ': cReg = cReg + 1
    If IdReg = 0 Then
            If IdReg = RegistroIncluir(strTabela, vReg, cReg) Then
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

    'Incluir NCM/ICMS /////////////////////////////////////////////////////
    If RegistroExcluir(strTabela & "ICMS", "idNCM = " & IdReg) = False Then
        MsgBox "Erro ao excluir registro " & IdReg, vbInformation, "Aviso"
    End If
    With msfgICMS
        For i = 1 To .Rows - 1
            cReg = 0
            vReg(cReg) = Array("idNCM", IdReg, "N"): cReg = cReg + 1
            vReg(cReg) = Array("UF", .TextMatrix(i, 1), "S"): cReg = cReg + 1
            vReg(cReg) = Array("ICMS", .TextMatrix(i, 2), "S") ': cReg = cReg + 1

            If IdReg = RegistroIncluir(strTabela & "ICMS", vReg, cReg) Then
                    MsgBox "Erro ao Incluir NCM/ICMS"
                    grvRegistro = False
                Else
                    grvRegistro = True
            End If
        Next
    End With
    

End Function


Private Sub MostrarDados()
    Dim sSQL As String
    
    txtDescricao.Text = ""
    txtNCM.Text = ""
    txtIPI.Text = ""
    txtCEST.Text = ""
    
    
    sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & IdReg
    ExibirDados Me, sSQL
    
    
    MostrarGridICMS

End Sub


Private Sub MostrarGrid(Optional idBusca As Integer)
    On Error Resume Next
    Dim Rst     As Recordset
    Dim sSQL    As String
    If idBusca = 0 Then
            sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " LIMIT 200"
        Else
            sSQL = "SELECT * FROM " & strTabela & " WHERE ID_Empresa = " & ID_Empresa & " AND Id=" & idBusca & " LIMIT 200"
    End If
    
    msfgTabICMS.Rows = 1
   
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                With msfgTabICMS
                    DoEvents
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("ID")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("Descricao")
                    .TextMatrix(.Rows - 1, 2) = Rst.Fields("NCM")
                    .TextMatrix(.Rows - 1, 3) = Rst.Fields("IPI")
                End With
                Rst.MoveNext
            Loop
    End If
    MostrarGridICMS
End Sub
Private Sub MostrarGridICMS()
    On Error Resume Next
    Dim Rst As Recordset
    
    msfgICMS.Rows = 1
   
    Set Rst = RegistroBuscar("SELECT * FROM " & strTabela & "ICMS WHERE ID_Empresa = " & ID_Empresa & " AND idNCM =" & IdReg)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                With msfgICMS
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("ID")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("UF")
                    .TextMatrix(.Rows - 1, 2) = Rst.Fields("ICMS")
                End With
                Rst.MoveNext
            Loop
    End If
End Sub


Private Sub txtCEST_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim codCest As String
    If KeyCode = 114 Then
        codCest = formBuscar.IniciarBusca("TributacaoCEST", "NCM,CEST,Descricao")
        codCest = PgDadosCEST("id", codCest, "N").cest
        If Len(Trim(codCest)) <> 0 Then
            txtCEST.Text = codCest
        End If
    End If
End Sub

Private Sub txtCEST_KeyPress(KeyAscii As Integer)

    KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub txtIPI_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtIPI.Text, KeyAscii, 2)

End Sub

Private Sub txtNCM_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
End Sub
