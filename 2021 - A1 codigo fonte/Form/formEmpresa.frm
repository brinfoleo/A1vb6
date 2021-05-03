VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form formEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresas"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   9135
   Begin TabDlg.SSTab sstEmpresa 
      Height          =   5355
      Left            =   60
      TabIndex        =   5
      Top             =   2040
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9446
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Basico"
      TabPicture(0)   =   "formEmpresa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Documentos"
      TabPicture(1)   =   "formEmpresa.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "formEmpresa.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame3 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   27
         Top             =   540
         Width           =   8655
         Begin VB.TextBox txtCNAE 
            Height          =   285
            Left            =   1560
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   1560
            Width           =   2235
         End
         Begin VB.TextBox txtIM 
            Height          =   285
            Left            =   1560
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   1140
            Width           =   2235
         End
         Begin VB.TextBox txtIEST 
            Height          =   285
            Left            =   1560
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   720
            Width           =   2235
         End
         Begin VB.TextBox txtIE 
            Height          =   285
            Left            =   1560
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   300
            Width           =   2235
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "CNAE:"
            Height          =   195
            Left            =   540
            TabIndex        =   31
            Top             =   1620
            Width           =   975
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Insc. Municipal:"
            Height          =   195
            Left            =   300
            TabIndex        =   30
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Insc. Estadual ST:"
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Insc. Estadual:"
            Height          =   195
            Left            =   300
            TabIndex        =   28
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4815
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   8775
         Begin VB.TextBox txtFone 
            Height          =   315
            Left            =   1140
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   3900
            Width           =   2955
         End
         Begin VB.TextBox txtMail 
            Height          =   285
            Left            =   1140
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   3540
            Width           =   3915
         End
         Begin VB.TextBox txtCEP 
            Height          =   285
            Left            =   1140
            MaxLength       =   8
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   2820
            Width           =   2175
         End
         Begin VB.ComboBox cboMun 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   2400
            Width           =   2655
         End
         Begin VB.ComboBox cboUF 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1980
            Width           =   915
         End
         Begin VB.TextBox txtBairro 
            Height          =   285
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   1560
            Width           =   2955
         End
         Begin VB.ComboBox cboPais 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1080
            Width           =   2955
         End
         Begin VB.TextBox txtCpl 
            Height          =   285
            Left            =   3180
            MaxLength       =   60
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   720
            Width           =   4695
         End
         Begin VB.TextBox txtNro 
            Height          =   285
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtLgr 
            Height          =   285
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   360
            Width           =   6735
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Telefone:"
            Height          =   195
            Left            =   420
            TabIndex        =   37
            Top             =   3900
            Width           =   675
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "e-mail:"
            Height          =   195
            Left            =   660
            TabIndex        =   36
            Top             =   3600
            Width           =   435
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "CEP:"
            Height          =   195
            Left            =   600
            TabIndex        =   14
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "UF:"
            Height          =   195
            Left            =   600
            TabIndex        =   13
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "País:"
            Height          =   195
            Left            =   480
            TabIndex        =   12
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Municipio:"
            Height          =   255
            Left            =   300
            TabIndex        =   11
            Top             =   2460
            Width           =   795
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Bairro:"
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   1620
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Complemento:"
            Height          =   195
            Left            =   2100
            TabIndex        =   9
            Top             =   780
            Width           =   1035
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Número:"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   780
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Endereço:"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   420
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   8955
      Begin VB.TextBox txtFant 
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1020
         Width           =   5295
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   1320
         MaxLength       =   160
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   660
         Width           =   5295
      End
      Begin VB.TextBox txtCNPJ 
         Height          =   285
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   255
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome Fantasia:"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Razão Social:"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   1095
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
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
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresa.frx":0054
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresa.frx":04A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresa.frx":07C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresa.frx":1052
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresa.frx":22A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresa.frx":2B7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresa.frx":3410
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresa.frx":3CA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresa.frx":4EF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresa.frx":520E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Pressione <F3> para consulta..."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   7440
      Width           =   8955
   End
End
Attribute VB_Name = "FormEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdEmpresa As Integer


Private Sub PesquisarRegistro()
    Dim psqTMP  As String
    psqTMP = FormBusca.IniciarBusca("Empresas")
    IdEmpresa = IIf(psqTMP = "", 0, psqTMP)
    
    If IdEmpresa = 0 Then
            LimpaFormulario FormEmpresa
        Else
            MostrarDados
    End If
End Sub



Private Sub cboMun_DropDown()
    Dim Rst     As Recordset
    Dim sSQL    As String
    If Trim(cboUF.Text) = "" Then
        MsgBox "Selecione uma Unidade Federal (UF)."
        Exit Sub
    End If
    sSQL = "SELECT * FROM tbMunicipio WHERE IdUF = " & PgUF(cboUF.Text).id & " ORDER BY Descricao"
    cboMun.Clear
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboMun.AddItem Rst.Fields("Descricao")
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboPais_DropDown()
    Dim Rst As Recordset
    cboPais.Clear
    Set Rst = RegistroBuscar("SELECT * FROM tbPaises ORDER BY Pais")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboPais.AddItem Rst.Fields("Pais")
                Rst.MoveNext
            Loop
    End If
End Sub

Private Sub cboUF_DropDown()
    Dim Rst As Recordset
    cboUF.Clear
    Set Rst = RegistroBuscar("SELECT * FROM tbUF ORDER BY UF")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboUF.AddItem Rst.Fields("UF")
                Rst.MoveNext
            Loop
    End If

End Sub

Private Sub Form_Load()
    LimpaFormulario FormEmpresa
    sstEmpresa.Tab = 0
    HDForm FormEmpresa, False
    HDMenu FormEmpresa, True
    
    txtCNPJ.Enabled = True
    
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Incluir"
            IdEmpresa = 0
            HDMenu FormEmpresa, False
            HDForm FormEmpresa, True
           
        Case "Alterar"
            If IdEmpresa = 0 Then
                MsgBox "Selecione uma empresa"
                Exit Sub
            End If
            HDForm FormEmpresa, True
            HDMenu FormEmpresa, False
        Case "Excluir"
            If IdEmpresa = 0 Then
                    MsgBox "Selecione uma Empresa"
                    Exit Sub
                Else
                    If MsgBox("Deseja relamente EXCLUIR este registro?                 " & vbCrLf & _
                               vbCrLf & _
                               "CNPJ: " & txtCNPJ.Text & vbCrLf & _
                               "Nome: " & txtNome.Text, vbYesNo + vbCritical) = vbYes Then
                               
                        If RegistroExcluir("empresas", "Id = " & IdEmpresa) = True Then
                            LimpaFormulario FormEmpresa
                        End If
                    End If
            End If
        Case "Pesquisar"
            PesquisarRegistro
            
        Case "Salvar"
            If grvRegistro = True Then
                HDMenu FormEmpresa, True
                HDForm FormEmpresa, False
                'LimpaFormulario FormEmpresa
                'txtCNPJ.Enabled = True
            End If
            
        
        Case "Cancelar"
            HDMenu FormEmpresa, True
            HDForm FormEmpresa, False
            LimpaFormulario FormEmpresa
            txtCNPJ.Enabled = True
    End Select
End Sub

Private Function grvRegistro() As Boolean
    Dim vReg(199)    As Variant
    Dim I           As Integer
    Dim Controle    As Control
    Dim cReg        As Integer 'Contador de Registros
    cReg = 0
    For I = 0 To FormEmpresa.Controls.Count - 1
        Set Controle = FormEmpresa.Controls(I)
        
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
    
     
    If IdEmpresa = 0 Then
            If RegistroIncluir("Empresas", vReg, cReg) = False Then
                    MsgBox "Erro ao Incluir"
                    grvRegistro = False
                Else
                    grvRegistro = True
            End If
        Else
            If RegistroAlterar("Empresas", vReg, cReg, "Id = " & IdEmpresa) = False Then
                    MsgBox "Erro ao Alterar."
                    grvRegistro = False
                Else
                    grvRegistro = True
                
            End If
    End If



End Function
Private Sub txtCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        PesquisarRegistro
    End If
    
End Sub

Private Sub txtCNPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        BuscarDados (txtCNPJ.Text)
    End If
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub
Private Sub BuscarDados(strCNPJ As String)
    Dim Rst     As ADODB.Recordset
    Dim strSQL  As String
    
    sstEmpresa.Tab = 0
    
    strSQL = "SELECT * FROM Empresas WHERE CNPJ = '" & strCNPJ & "'"

    Set Rst = RegistroBuscar(strSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum Registro encontrado"
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            IdEmpresa = Rst.Fields("Id")
            Rst.Close
            MostrarDados
    End If
    
    
    
    
End Sub
Private Sub MostrarDados()
    Dim sSQL As String
    sSQL = "SELECT * FROM Empresas WHERE Id = " & IdEmpresa

    ExibirDados FormEmpresa, sSQL
'    txtCNPJ.Text = PgDadosEmpresa(IdEmpresa).cnpj
'    txtNome.Text = PgDadosEmpresa(IdEmpresa).nome
'    txtFant.Text = PgDadosEmpresa(IdEmpresa).fant
'    txtLgr.Text = PgDadosEmpresa(IdEmpresa).lgr
'    txtNro.Text = PgDadosEmpresa(IdEmpresa).nro
'    txtCpl.Text = PgDadosEmpresa(IdEmpresa).cpl
'    txtBairro.Text = PgDadosEmpresa(IdEmpresa).bairro
'    cboPais.Clear
'    cboPais.AddItem IIf(Trim(PgDadosEmpresa(IdEmpresa).pais) = "", "  ", PgDadosEmpresa(IdEmpresa).pais)
'    cboPais.Text = cboPais.List(0)
'    cboUF.Clear
'    cboUF.AddItem IIf(Trim(PgDadosEmpresa(IdEmpresa).uf) = "", "  ", PgDadosEmpresa(IdEmpresa).uf)
'    cboUF.Text = cboUF.List(0)
'    cboMun.Clear
'    cboMun.AddItem IIf(Trim(PgDadosEmpresa(IdEmpresa).mun) = "", "  ", PgDadosEmpresa(IdEmpresa).mun)
'    cboMun.Text = cboMun.List(0)
'    txtCEP.Text = PgDadosEmpresa(IdEmpresa).cep
'    txtIE.Text = PgDadosEmpresa(IdEmpresa).ie
'    txtIEST.Text = PgDadosEmpresa(IdEmpresa).iest
'    txtIM.Text = PgDadosEmpresa(IdEmpresa).im
'    txtCNAE.Text = PgDadosEmpresa(IdEmpresa).cnae
'
'    txtMail.Text = PgDadosEmpresa(IdEmpresa).mail
'    txtFone.Text = PgDadosEmpresa(IdEmpresa).fone

End Sub
