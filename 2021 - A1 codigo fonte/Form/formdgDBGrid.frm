VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form formdbDBGrid 
   Caption         =   "Exibindo Tabelas e Dados"
   ClientHeight    =   8100
   ClientLeft      =   1665
   ClientTop       =   1515
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8100
   ScaleWidth      =   15015
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   2580
      TabIndex        =   3
      Top             =   1200
      Width           =   5775
      Begin VB.CommandButton btoExec 
         Caption         =   "=>"
         Height          =   315
         Left            =   5400
         TabIndex        =   6
         Top             =   480
         Width           =   315
      End
      Begin VB.TextBox txtSQL 
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "Comando SQL"
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.ListBox lstTabelas 
      Height          =   3765
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3915
      Left            =   60
      TabIndex        =   1
      Top             =   4020
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   6906
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExibeTabelas 
      Appearance      =   0  'Flat
      Caption         =   "Exibir Tabelas"
      Height          =   375
      Left            =   2580
      TabIndex        =   0
      Top             =   180
      Width           =   2715
   End
End
Attribute VB_Name = "formdbDBGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btoExec_Click()
    On Error GoTo TrtExec
    LoadTable (Trim(txtSQL.Text))
    Exit Sub
TrtExec:
    MsgBox Err.Description, vbCritical, Err.Number
    Resume Next
End Sub

Private Sub cmdExibeTabelas_Click()
    
    lstTabelas.Clear

Dim rstSchema As ADODB.Recordset
Dim strCnn As String
Set rstSchema = BD.OpenSchema(adSchemaTables)
Do Until rstSchema.EOF
    lstTabelas.AddItem (rstSchema!TABLE_NAME)
rstSchema.MoveNext
Loop
rstSchema.Close

End Sub

Private Sub cmdProcuraArquivo_Click()
   
   On Error Resume Next
    
    'cmdlg1.ShowOpen
    
    If Err.Number = cdlCancel Then
        ' cancelado pelo usuario.
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ' erro desconhecido
        MsgBox "Erro " & Format$(Err.Number) & " ao selecionar o arquivo." & vbCrLf & Err.Description
        Exit Sub
    End If
    
    On Error GoTo 0

    'txtArquivo.Text = cmdlg1.FileName

End Sub



'Private Sub Command1_Click()
'    If MsgBox("Excluir tabela " & UCase(lstTabelas.Text) & "?", vbCritical + vbYesNo, "Exluir") = vbYes Then
'        BD.Execute "DROP TABLE " & lstTabelas.Text
'        cmdExibeTabelas_Click
'    End If
'End Sub



'Private Sub Command2_Click()
'14.11.2012 - Metodo para criar um novo deposito com base no antigo
'    Dim sSQL    As String
'    Dim Rst     As Recordset
'    Dim i       As Integer
'    Dim nCampo   As String
'    Dim vCampo As String
'
'    Dim sValues As String
'    Dim sFields As String
'
'    sSQL = "SELECT * FROM estoqueproduto WHERE Deposito = " & ID_Deposito
'    Set Rst = RegistroBuscar(sSQL)
'    If Rst.BOF And Rst.EOF Then Exit Sub
'    pb.Value = 0
'    pb.min = 0
'    pb.Max = Rst.RecordCount
'    List1.Clear
'    Rst.MoveFirst
'    Do Until Rst.EOF
'        pb1.Value = 0
'        pb1.min = 0
'        pb1.Max = Rst.fields.Count
'        For i = 0 To Rst.fields.Count - 1
'            DoEvents
'            pb1.Value = pb1.Value + 1
'            nCampo = Rst.fields(i).Name
'
'            If Rst.fields(i).Type = adInteger Then
'                    vCampo = IIf(cNull(Rst.fields(i)) = "", 0, Rst.fields(i))
'                Else
'                    vCampo = IIf(cNull(Rst.fields(i)) = "", "''", "'" & Rst.fields(i) & "'")
'            End If
'            If Trim(LCase(nCampo)) <> "id" Then
'                If nCampo = "Deposito" Then
'                    vCampo = 2
'                End If
'
'                If nCampo = "Saldo" Then
'                    vCampo = 0
'                End If
'                If vCampo = "''" Then vCampo = "Null"
'
'                sFields = IIf(Trim(sFields) = "", nCampo, sFields & "," & nCampo)
'                sValues = IIf(Trim(sValues) = "", vCampo, sValues & "," & vCampo)
'
'
'            End If
'        Next
'        DoEvents
'        List1.AddItem cNull(Rst.fields("Descricao"))
'        pb.Value = pb.Value + 1
'        BD.Execute "INSERT INTO estoqueproduto (" & sFields & ") VALUES(" & sValues & ")"
'        sFields = ""
'        sValues = ""
'        Rst.MoveNext
'    Loop
'
'End Sub
Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    'define o titulo do diálogo
    'cmdlg1.DialogTitle = "Procurar Arquivos .mdb"
    'define o caminho inicial
    'cmdlg1.InitDir = App.Path
    'define o filtro para exibir os arquivos
    'cmdlg1.Filter = "Arqs. MDB(*.mdb)|*.mdb|Todos " & "Arqs. (*.*)|*.*"
    'cmdlg1.FilterIndex = 1
    'define algumas variaveis
    'cmdlg1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNExplorer
    'dispara um erro se não for selecionado algo
    'cmdlg1.CancelError = True
    txtSQL.Text = ""
    cmdExibeTabelas_Click
End Sub

Private Sub Form_Resize()
    On Error GoTo sai
    With DataGrid1
        .Left = 50
        .Top = lstTabelas.Height + 300
        .Width = Me.Width - 300
        .Height = Me.ScaleHeight - (lstTabelas.Height + 600)
    End With
    Exit Sub
sai:
    
End Sub



'abre a tabela selecionada
Private Sub lstTabelas_Click()
    Dim Tabela As String
    'obtem o nome da tabela selecionada na lista
    Tabela = lstTabelas.List(lstTabelas.ListIndex)
    
    'monta instrução sql para selecionar todos os registros da tabela
    LoadTable ("SELECT * FROM " & Tabela)
  
End Sub

Private Sub LoadTable(sQL As String)
 Dim Rst As Recordset
 
 '**************************************************
 '***** 22/02/2013                            ******
 '***** REGISTRAR TODOS O COMANDOS EXECUTADOS ******
 '***** NA TELA  formdbDBGrid.                ******
    RegLog "0", "0", "formdbDBGrid: " & sQL
 '**************************************************
 
 
 'Dim nome_tabela As String
 '   Dim sql As String
'    Dim Tabela As String
    
'On Error GoTo trataerro

    
 
    'define a fonte de dados como sendo a tabela
'    Data1.Caption = nome_tabela
'    Data1.RecordSource = sql
'    Data1.Refresh
    
    ' torna o controle data e o dbgrid visiveis
'    Data1.Visible = True
'    DataGrid1.Visible = True
'    Exit Sub
'trataerro:
'    MsgBox Err.Number & vbCrLf & Err.Description
'**********************************************************************************************
    'Set Rst = RegistroBuscar("SELECT xNome, xFant FROM empresas")
    'On Error GoTo TrtErrGrid
    Set Rst = RegistroBuscar(sQL)
    If Rst Is Nothing Then
        'DataGrid1.Enabled = False
        'Text1.Enabled = False
        Me.Caption = "Busca - [ 00000 Registros]"
        Exit Sub
    End If
    
    If Rst.BOF And Rst.EOF Then
            'DataGrid1.Enabled = False
            'Text1.Enabled = False
            Me.Caption = "Busca - [ 00000 Registros]"
        Else
            DataGrid1.Enabled = True
            Rst.MoveLast
            Me.Caption = "Busca - [ " & Left(String(5, "0"), 5 - Len(Trim(Rst.RecordCount))) & Trim(Rst.RecordCount) & " Registros]"
            Rst.MoveFirst
    End If
    Set DataGrid1.DataSource = Rst.DataSource
    With DataGrid1
        .AllowUpdate = False
        '.EditActive = False
        '.Columns(0).Caption = "Razão Social"
        ''.Columns(0).DataField = Rst.Fields("xNome")
        '.Columns(0).Width = TextWidth("Razão Social") + 1000
        '.Columns(0).Alignment = dbgLeft

        '.Columns(1).Caption = "Razão Social"
        ''.Columns(1).DataField = Rst.Fields("xNome")
        '.Columns(1).Width = TextWidth("Razão Social") + 1000
        '.Columns(1).Alignment = dbgLeft
        
    End With
    Exit Sub
'TrtErrGrid:
'    MsgBox "Descrição: " & Err.Description, vbCritical, "Erro n. " & Err.Number
'    resultadoBusca = ""
'    Unload Me

End Sub
Private Sub txtSQL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btoExec_Click
    End If
        
End Sub
