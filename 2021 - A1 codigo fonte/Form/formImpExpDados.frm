VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form formImpExpDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar e Exportar Dados"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7530
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   60
      TabIndex        =   8
      Top             =   2400
      Width           =   7395
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   7395
      Begin MSComDlg.CommonDialog cd 
         Left            =   6000
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton btoPesq 
         Height          =   315
         Left            =   6780
         Picture         =   "formImpExpDados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   930
         Width           =   375
      End
      Begin VB.TextBox txtCaminho 
         Height          =   285
         Left            =   1140
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   960
         Width           =   5595
      End
      Begin VB.ComboBox cboTabela 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3795
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Arquivo:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tabela:"
         Height          =   195
         Left            =   420
         TabIndex        =   2
         Top             =   540
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
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
            Object.ToolTipText     =   "Importar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar"
            ImageIndex      =   13
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
               Picture         =   "formImpExpDados.frx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":07DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":0AF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":1388
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":25DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":2EB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":3746
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":3FD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":522A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":5544
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":585E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":5C55
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formImpExpDados.frx":61EF
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formImpExpDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTabela As String

Private Sub ImportarDados()
    Dim f               As Long
    Dim linha           As String
    Dim vReg(1000)      As String
    Dim vDados(1000)    As Variant
    Dim cReg            As Integer
    Dim i               As Integer
    Dim X               As Integer
    Dim c               As Integer
    
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    
    List1.Clear
    f = FreeFile
    Open txtCaminho.Text For Input As f   'abre o arquivo texto
    
    pb.Max = 100
    pb.Min = 0
    
    cReg = 0
    i = 0
    X = 0
    c = 0
    Do While Not EOF(f)
        Line Input #f, linha 'lê uma linha do arquivo texto
                
        If i = 0 Then 'Retira o cabecalho
                Do Until InStr(linha, "|") = 0
                    vReg(c) = Mid(linha, 1, InStr(linha, "|") - 1): c = c + 1 'pega o dado
                    linha = Mid(linha, InStr(linha, "|") + 1, Len(linha)) 'diminui os caracteres da linha
                Loop
                c = c - 1
            Else
                
                For X = 0 To c
                    If UCase(vReg(X)) <> "ID" And UCase(vReg(X)) <> "DTHR" And UCase(vReg(X)) <> "ID_EMPRESA" Then
                        vDados(cReg) = Array(vReg(X), Mid(linha, 1, InStr(linha, "|") - 1), "S"): cReg = cReg + 1
                    End If
                    linha = Mid(linha, InStr(linha, "|") + 1, Len(linha)) 'diminui os caracteres da linha
                Next
                cReg = cReg - 1
                '********************************************************************
                'As linhas abaixo forao colocadas para atualizar a
                'tabela NCM com base em uma tabela feita e exportada no excel/acess
                '*** Colocar esta linha apos o comando INPUT ***********************
                'linha = linha & "|"
                '**********************************************************************
                'If vDados(6)(1) = "NT" Or Trim(vDados(6)(1)) = "" Then
                '    vDados(6)(1) = "0"
                '    Else
                '        vDados(6)(1) = Trim(vDados(6)(1))
                'End If
                '
                'BD.Execute "UPDATE " & strTabela & " " & _
                '                    "SET IPI = " & vDados(6)(1) & " " & _
                '                    "WHERE NCM = '" & vDados(4)(1) & "'"
                '*****************************************************************************
                RegistroIncluir strTabela, vDados, cReg  'c
                cReg = 0
        End If
        i = i + 1
        pb.Value = IIf(pb.Value = 100, 0, pb.Value + 1)
    Loop
    pb.Value = 100
    MsgBox i & " registros IMPORTADOS com suceso!", vbInformation, "Importação & Exportação de Dados"

    Close #f
End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Trim(cboTabela.Text) = "" Then
        MsgBox "Favor selecionar uma tabela!"
        Exit Sub
    End If
    If Trim(txtCaminho.Text) = "" Then
        MsgBox "Selecione o local do arquivo!"
        Exit Sub
    End If
    pb.Visible = True
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Importar"
            ImportarDados
        Case "Exportar"
            ExportarDados
    End Select
    pb.Visible = False
End Sub


Private Sub ExportarDados()
    Dim Rst         As Recordset
    Dim Coluna      As Field
    Dim strDados    As String
    Dim i           As Integer
    
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    
    
    If Dir(Trim(txtCaminho.Text)) <> "" Then
        If MsgBox("Arquivo já existe! Deseja substituir?", vbYesNo + vbExclamation, "Aviso") = vbYes Then
            Kill txtCaminho.Text
        End If
    End If
    
    
    If strTabela = "" Then Exit Sub
    
    Set Rst = RegistroBuscar("SELECT * FROM " & strTabela)
    If Rst.BOF And Rst.EOF Then
        MsgBox "Banco de Dados Vazio"
        Exit Sub
    End If
    pb.Max = Rst.RecordCount + 1
    pb.Min = 0
    
    If Rst.BOF And Rst.EOF Then
            MsgBox "Tabela " & strTabela & " sem registros!"
            Exit Sub
        Else
            Rst.MoveFirst
            'Imprime o cabecalho
            strDados = ""
            For Each Coluna In Rst.Fields
                strDados = strDados & Coluna.Name & "|"
            Next
            grvReg txtCaminho.Text, strDados
            
            
            'Imprime os dados
            strDados = ""
            For i = 1 To Rst.RecordCount
                For Each Coluna In Rst.Fields
                    strDados = strDados & IIf(IsNull(Rst.Fields(Coluna.Name)), "", Rst.Fields(Coluna.Name)) & "|"
                Next
                grvReg txtCaminho.Text, strDados
                strDados = ""
                Rst.MoveNext
                pb.Value = IIf(pb.Max = pb.Value, pb.Value, pb.Value + 1)
            Next
            
    End If
    Rst.Close
    
    MsgBox i & " registros EXPORTADOS com suceso!", vbInformation, "Importação & Exportação de Dados"
End Sub
Private Sub btoPesq_Click()
    cd.DialogTitle = "A1 - Importar & Exportar"
    cd.InitDir = SistemPath
    cd.Filter = "Texto|*.txt"
    cd.ShowOpen
    txtCaminho.Text = cd.FileName
End Sub

Private Sub cboTabela_Click()
    If Trim(cboTabela.Text) = "" Then
            strTabela = ""
            Exit Sub
        Else
            strTabela = Trim(cboTabela.Text)
    End If
End Sub

Private Sub cboTabela_DropDown()
    Dim Rst As Recordset
    cboTabela.Clear
    
    Set Rst = BD.OpenSchema(adSchemaTables)
    
    Rst.MoveFirst
    Do Until Rst.EOF
        'If Left(Rst("TABLE_NAME"), 4) <> "MSys" Then
            cboTabela.AddItem Rst("TABLE_NAME")
        'End If
        Rst.MoveNext
    Loop
    Rst.Close
End Sub





Private Sub Form_Load()
    LimpaFormulario Me
    HDMenu Me, True
End Sub
Private Sub grvReg(nmArquivo As String, Dados As String)
    On Error GoTo TrtErro
    
    'define o ObjPreview filesystem e demais variaveis
    Dim fso As New FileSystemObject
    Dim Arquivo As File
    Dim arquivoLog As TextStream
    Dim msg As String
    Dim caminho As String

    

    caminho = nmArquivo
    'se o arquivo não existir então cria
    If fso.FileExists(caminho) Then
            Set Arquivo = fso.GetFile(caminho)
        Else
            Set arquivoLog = fso.CreateTextFile(caminho)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(caminho)
    End If
    'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
    
    'monta informações para gerar a linha da mensagem
    msg = Dados

    'inclui linhas no arquivo texto
    arquivoLog.WriteLine msg
    
    'escreve uma linha em branco no arquivo - se voce quiser
    'arquivoLog.WriteBlankLines (1)
    'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing
    Exit Sub
TrtErro:
        MsgBox "Erro ao gerar registro exportacao.                                   " & _
           vbCrLf & vbCrLf & _
           "Erro n.: " & Err.Number & _
           vbCrLf & vbCrLf & _
           "Descrição: " & Err.Description & _
           vbCrLf

End Sub





