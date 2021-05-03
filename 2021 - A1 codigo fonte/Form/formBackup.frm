VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Iniciar Backup"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Restaurar Backup"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Manutenção na Base de Dados"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin MSComDlg.CommonDialog cd 
         Left            =   6360
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5220
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formBackup.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formBackup.frx":06FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formBackup.frx":0DF4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.ListBox lstMSG 
      Height          =   1425
      Left            =   60
      TabIndex        =   19
      Top             =   6240
      Width           =   7575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   60
      TabIndex        =   3
      Top             =   1500
      Width           =   7635
      Begin VB.Frame Frame3 
         Caption         =   "Processamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   7395
         Begin MSComctlLib.ProgressBar pb2 
            Height          =   315
            Left            =   180
            TabIndex        =   9
            Top             =   2520
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar pb 
            Height          =   315
            Left            =   180
            TabIndex        =   10
            Top             =   1920
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label Label12 
            Caption         =   "Bytes processados:"
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   1260
            Width           =   2415
         End
         Begin VB.Label lblProcessoatual 
            Caption         =   "lblProcessoatual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   27
            Top             =   1260
            Width           =   3195
         End
         Begin VB.Label Label7 
            Caption         =   "Quantidade de registros processados:"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label Label6 
            Caption         =   "Nome da tabela em processamento:"
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   660
            Width           =   2655
         End
         Begin VB.Label Label5 
            Caption         =   "Quantidade de tabelas processadas:"
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   300
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "Tabelas processadas:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Registros processados:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label lblTabelaAtual 
            Caption         =   "lblTabelaAtual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3000
            TabIndex        =   13
            Top             =   300
            Width           =   3675
         End
         Begin VB.Label lblNomeTabela 
            Caption         =   "lblNomeTabela"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3000
            TabIndex        =   12
            Top             =   660
            Width           =   3855
         End
         Begin VB.Label lblQtdeRegistros 
            Caption         =   "lblQtdeRegistros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3000
            TabIndex        =   11
            Top             =   960
            Width           =   4095
         End
      End
      Begin VB.Label Label11 
         Caption         =   "Tamanho do Arquivo:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1260
         Width           =   1575
      End
      Begin VB.Label lbltamanhoArquivo 
         Caption         =   "lbltamanhoArquivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   25
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblLocalAramazenamento 
         Caption         =   "lblLocalAramazenamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   23
         Top             =   660
         Width           =   5415
      End
      Begin VB.Label Label10 
         Caption         =   "Local de armazenamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Label4 
         Caption         =   "Total de Tabelas:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblTotalTabelas 
         Caption         =   "lblTotalTabelas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   6
         Top             =   1020
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Arquivo:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label lblNomeBackup 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   4
         Top             =   300
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Base de Dados"
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
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   7635
      Begin VB.CheckBox chkIncDados 
         Caption         =   "Incluir dados"
         Height          =   195
         Left            =   6000
         TabIndex        =   29
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Porta:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "IP do Servidor:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblIP 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1380
         TabIndex        =   2
         Top             =   360
         Width           =   4155
      End
      Begin VB.Label lblPorta 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1380
         TabIndex        =   1
         Top             =   660
         Width           =   4215
      End
   End
End
Attribute VB_Name = "formBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Imprimir Grade
Option Explicit
Const MSG_01 = "Bakck UP Criado por: "
Const MSG_02 = "Base de Dados: "
Const MSG_03 = "Inicio/Hora: "
Const MSG_04 = "DD/MM/YY HH:MM:SS"
Const MSG_05 = "DBMS: MySQL v"
Const MSG_06 = "Estrutura da tabela "
Const MSG_07 = "Dados da tabela "
Const MSG_08 = "Fim do Backup: "
Dim sStop   As Boolean
Private Sub LimpForm()
    lblNomeBackup.Caption = ""
    lblTotalTabelas.Caption = "0"
    lblLocalAramazenamento.Caption = ""
    lblTabelaAtual.Caption = "0"
    lblNomeTabela.Caption = ""
    lblQtdeRegistros.Caption = "0"
    lbltamanhoArquivo.Caption = ""
    lblProcessoatual.Caption = ""
End Sub




Public Sub IniciarBackup()
    Dim nmArquivo   As String
    Dim storageFile As String
    
    storageFile = PgDadosConfig.pFileArmazenamento
    storageFile = IIf(Len(Trim(storageFile)) = 0, App.Path, storageFile)
    
    
    nmArquivo = "A1 " & Format(Date, "YYYY-MM-DD") & " " & Replace(Format(Time, "hh:mm:ss"), ":", ";") & ".sql"
    If PastaExiste(storageFile & "\Backup\") = False Then
        MkDir (storageFile & "\Backup\")
        'Call MySQLBackup(PgDadosConfig.pFileArmazenamento & "\Backup\Backup_" & Format(Date, "dd_MM_yy") & "_" & Format(Time, "hh_mm_ss") & ".sql", BD, lstLog)
    End If
    lblNomeBackup.Caption = nmArquivo
    lblLocalAramazenamento.Caption = storageFile & "\Backup\"
    Call MySQLBackup(storageFile & "\Backup\" & nmArquivo, BD) ', lstLog)
    
End Sub
Private Function PastaExiste(Dir As String) As Boolean
    Dim oDir As New Scripting.FileSystemObject
    PastaExiste = oDir.FolderExists(Dir)
End Function


Private Sub RestaurarBackup()
    Dim strArquivo As String
    If chkAcesso(Me, "n") = False Then
        Exit Sub
    End If
    cd.Filter = "Arquivo SQL (*.sql)| *.sql"
    cd.ShowOpen
    strArquivo = cd.filename
    If cd.filename <> "" Then
        Call MySQLRestore(strArquivo, BD, pb)
        'btnStart.Enabled = True
    Else
        MsgBox "Pasta ou arquivo não encontrado !!!"
    Exit Sub
 End If
Fim:

End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    'lblSerial.Caption = Numero_Serial
    lblIP.Caption = srv_IP
    lblPorta.Caption = srv_Porta
    HDMenu Me, True
    LimpForm
    
End Sub

Public Sub MySQLRestore(ByVal strNomeArquivo As String, cnn As ADODB.Connection, pb As ProgressBar) 'lst As ListView,
    'frmBackup.btnIniciarRestaura.Enabled = False
    'frmBackup.btnBackup.Enabled = False
    Dim smsg As String
    'Dim nodeX2 As ListItem
    'lst.View = lvwReport
    'lst.ListItems.Clear
    'lst.ColumnHeaders.Clear
    'lst.GridLines = False
    'lst.ColumnHeaders.Add , , "log"
    'lst.ColumnHeaders(1).Width = 5000
    
    Dim lngTotalBytes As Long, lngCurrentBytes As Long
    Dim X As Integer, strLinha As String, strAux As String
    Dim blnPassLines As Boolean
    Dim blnAnalizeIt As Boolean
    Dim sBaseReplace As String
    
    lblNomeBackup.Caption = strNomeArquivo
    lblLocalAramazenamento.Caption = strNomeArquivo
    X = FreeFile
    
    On Error GoTo ErrDrv
    
    Open strNomeArquivo For Input As #X
        lngTotalBytes = LOF(X)
        smsg = "Abrindo Arquivo de Backup"
        lstMSG.AddItem smsg

        'Set nodeX2 = lst.ListItems.Add(, , smsg)
        'nodeX2.EnsureVisible: DoEvents
    
    blnPassLines = False
    pb.Max = lngTotalBytes
    pb.Value = 0
    Dim xpb As Long
    Dim linha As String
    Dim IZ As Integer
    Dim sNomeBase As String
    Do While Not EOF(X)
    IZ = IZ + 1
    
        Line Input #X, strLinha
        
        If UCase(Left(strLinha, 16)) = UCase("# Base de Dados:") Then
            sNomeBase = Mid(strLinha, InStr(strLinha, "# Base de Dados: ") + 17)
            sNomeBase = "`" & Left(sNomeBase, InStr(sNomeBase, ";") - 1) & "`"
            
            'If frmBackup.optNovoBanco.Value = True Then
            '    sBaseReplace = " `" & frmBackup.txtNovoBanco & "`"
            'ElseIf frmBackup.optOriginal.Value = True Then
                sBaseReplace = sNomeBase
            'ElseIf frmBackup.optOutroBanco.Value = True Then
            '    sBaseReplace = " `" & frmBackup.cmbBanco & "`"
            'End If
            
        End If
        
       
        
        
        
        Select Case IZ
        
            Case 11
                If Left(UCase(strLinha), 4) = "DROP" Then
                    strLinha = Replace(strLinha, sNomeBase, sBaseReplace)
                End If
            
        
        
            Case 12
                If Left(UCase(strLinha), 6) = "CREATE" Then
                    strLinha = Replace(strLinha, sNomeBase, sBaseReplace)
                End If
                
        
        
            Case 13
                If Left(UCase(strLinha), 3) = "USE" Then
                    strLinha = Replace(strLinha, sNomeBase, sBaseReplace)
                End If
                        
        End Select
                    
        
        lngCurrentBytes = lngCurrentBytes + Len(strLinha)
        
        lbltamanhoArquivo.Caption = lngTotalBytes & " Byte(s)": DoEvents
        lblProcessoatual.Caption = lngCurrentBytes
        xpb = pb.Max / lngTotalBytes
        pb.Value = lngCurrentBytes: DoEvents
                
        blnAnalizeIt = True
        strLinha = Trim(strLinha)
        If Not blnPassLines Then
            If Left(strLinha, 1) = "#" Then
                blnAnalizeIt = False
            ElseIf Left(strLinha, 2) = "/*" Then
                blnAnalizeIt = False
                blnPassLines = True
            End If
        ElseIf Right(Trim(strLinha), 2) = "*/" Then
            blnPassLines = False
            blnAnalizeIt = False
        End If
         
        If blnAnalizeIt And strLinha <> "" Then

            While Mid(strLinha, Len(strLinha), 1) <> ";"
                strAux = strLinha
                Line Input #X, strLinha
                lngCurrentBytes = lngCurrentBytes + Len(strLinha)
                strLinha = Trim(strLinha)
                strLinha = strAux & strLinha
            Wend
            
            smsg = "Executando comando " & Left(strLinha, 255)
            lstMSG.AddItem smsg

            'Set nodeX2 = lst.ListItems.Add(, , smsg)
            'nodeX2.EnsureVisible: DoEvents
            'MsgBox "OPS"
            cnn.Execute strLinha
            
        End If
        
    Loop
    
    Close #X
    
    lblProcessoatual = lngTotalBytes
    smsg = "Processo concluído com sucesso !!!"
    lstMSG.AddItem smsg

    'Set nodeX2 = lst.ListItems.Add(, , smsg)
    'nodeX2.EnsureVisible: DoEvents
    'frmBackup.btnIniciarRestaura.Enabled = True
    'frmBackup.btnBackup.Enabled = True
    Exit Sub
ErrDrv:
    
    smsg = "ERROR:" & Err.Number & vbNewLine & Err.Description & vbNewLine
    'Set nodeX2 = lst.ListItems.Add(, , smsg)
    'nodeX2.EnsureVisible: DoEvents
    
    Err.Clear

End Sub
Public Sub MySQLBackup(ByVal strNomeArquivo As String, cnn As ADODB.Connection)
    Dim smsg                    As String
    Dim lngBytesProcessados     As Long
    
    On Error Resume Next
    
    Dim rss         As ADODB.Recordset
    Dim rssAux      As ADODB.Recordset
    
    Dim X As Long, i As Integer
    
    Dim strNomeTabela   As String
    Dim strLinha        As String
    Dim strBuffer       As String
    Dim strNomeBase     As String
    
    X = FreeFile
    Open strNomeArquivo For Output As X
    
    Print #X, ""
    Print #X, "#"
    
    Print #X, "# " & MSG_01 & App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
    smsg = "# " & MSG_01 & App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
    lstMSG.AddItem smsg
    lngBytesProcessados = Len(smsg) + lngBytesProcessados
    lblProcessoatual.Caption = lngBytesProcessados

    'strNomeBase = Mid(cnn.ConnectionString, InStr(cnn.ConnectionString, "database=") + 9, Len(cnn.ConnectionString))
    'strNomeBase = Left(strNomeBase, InStr(strNomeBase, ";") - 1)
    strNomeBase = Trim(cnn.DefaultDatabase)
    
    Print #X, "# " & MSG_02 & strNomeBase & ";"
    smsg = "# " & MSG_02 & strNomeBase
    lstMSG.AddItem smsg
    lngBytesProcessados = Len(smsg) + lngBytesProcessados
    lblProcessoatual.Caption = lngBytesProcessados
    
    Set rss = New ADODB.Recordset
    Set rssAux = New ADODB.Recordset

    Print #X, "# " & MSG_03 & Format(Now, MSG_04)
    smsg = "# " & MSG_03 & Format(Now, MSG_04)
    lstMSG.AddItem smsg
    lngBytesProcessados = Len(smsg) + lngBytesProcessados
    lblProcessoatual.Caption = lngBytesProcessados
    
    
    rss.Open "show variables like 'version';", cnn
    If Not rss.EOF Then
        Print #X, "# " & MSG_05 & rss.Fields(1)
        smsg = "# " & MSG_05 & rss.Fields(1)
        lstMSG.AddItem smsg
        lngBytesProcessados = Len(smsg) + lngBytesProcessados
        lblProcessoatual.Caption = lngBytesProcessados
    
    End If
    rss.Close

    Print #X, "#"
    Print #X, ""
    Print #X, "SET FOREIGN_KEY_CHECKS=0;"
    smsg = "Desativando a checagem de Constraint;"
    lstMSG.AddItem smsg
    lngBytesProcessados = Len(smsg) + lngBytesProcessados
    lblProcessoatual.Caption = lngBytesProcessados
    
    Print #X, ""
    Print #X, "DROP DATABASE IF EXISTS `" & strNomeBase & "`;"
    smsg = "Excluido o banco " & strNomeBase
    lstMSG.AddItem smsg
    lngBytesProcessados = Len(smsg) + lngBytesProcessados
    lblProcessoatual.Caption = lngBytesProcessados
    
    Print #X, "CREATE DATABASE `" & strNomeBase & "`;"
    smsg = "Criando o Banco " & strNomeBase
    lstMSG.AddItem smsg
    lngBytesProcessados = Len(smsg) + lngBytesProcessados
    lblProcessoatual.Caption = lngBytesProcessados
    
    Print #X, "USE `" & strNomeBase & "`;"
    smsg = "Executando o comando USE `" & strNomeBase & "`;"
    lstMSG.AddItem smsg
    lngBytesProcessados = Len(smsg) + lngBytesProcessados
    lblProcessoatual.Caption = lngBytesProcessados
    
    strNomeTabela = ""

    With rss
        .Open "SHOW TABLE STATUS", cnn
        lblTotalTabelas = rss.RecordCount: DoEvents
        
        smsg = rss.RecordCount & " Tabelas no Banco " & strNomeBase
        lstMSG.AddItem smsg
        lngBytesProcessados = Len(smsg) + lngBytesProcessados
        lblProcessoatual.Caption = lngBytesProcessados
    
        pb.Value = 0
        pb.Max = lblTotalTabelas.Caption
        Dim xpb2 As Integer
        Do While Not .EOF
        xpb2 = lblTabelaAtual.Caption
        pb.Value = xpb2
        
        If sStop = True Then
            smsg = "Processo Interrompido pelo usuário"
            lstMSG.AddItem smsg
            lngBytesProcessados = Len(smsg) + lngBytesProcessados
            lblProcessoatual.Caption = lngBytesProcessados
            sStop = False
            Exit Sub
        End If
            strNomeTabela = .Fields.Item("Name").Value
            
            lblNomeTabela.Caption = strNomeTabela: DoEvents
            lblTabelaAtual.Caption = rss.AbsolutePosition: DoEvents
            With rssAux

                .Open "SHOW CREATE TABLE " & strNomeTabela, cnn
                
                Print #X, ""
                Print #X, "#"
                Print #X, "# " & MSG_06 & strNomeTabela & ""
                smsg = "# " & MSG_06 & strNomeTabela & ""
                lstMSG.AddItem smsg
                lngBytesProcessados = Len(smsg) + lngBytesProcessados
                lblProcessoatual.Caption = lngBytesProcessados
                Print #X, "#"
                
                Print #X, .Fields.Item(1).Value & ";"
                
                
                .Close
                
            End With
            'Preenche com os dados da tabela
            If chkIncDados.Value = 1 Then
                With rssAux
                    .Open "SELECT * FROM " & strNomeTabela & "", cnn
                    lblQtdeRegistros.Caption = rssAux.RecordCount & " Registro(s)": DoEvents
                    smsg = "Selecionando a tabela " & strNomeTabela
                    lstMSG.AddItem smsg
                    lngBytesProcessados = Len(smsg) + lngBytesProcessados
                    lblProcessoatual.Caption = lngBytesProcessados
        
                    Print #X, ""
                    Print #X, "#"
                    Print #X, "# " & MSG_07 & strNomeTabela & ""
                    smsg = "# " & MSG_07 & strNomeTabela & ""
                    lstMSG.AddItem smsg
                    lngBytesProcessados = Len(smsg) + lngBytesProcessados
                    lblProcessoatual.Caption = lngBytesProcessados
        
                    Print #X, "#"
                    Print #X, "lock tables `" & strNomeTabela & "` write;"
                    smsg = "Bloqueando a tabela " & strNomeTabela & " contra gravação;"
                    lstMSG.AddItem smsg
                    lngBytesProcessados = Len(smsg) + lngBytesProcessados
                    lblProcessoatual.Caption = lngBytesProcessados
        
    
                    If Not .EOF Then
                        pb2.Max = .RecordCount
                        pb2.Value = 0
                        Dim xpb As Integer
                        xpb = pb2.Max / .RecordCount
                                                
                        smsg = "Inserindo os dados na tabela " & strNomeTabela
                        lstMSG.AddItem smsg
                        
                        Do While Not .EOF
    
                        On Error Resume Next
                            pb2.Value = pb2.Value + xpb: DoEvents
                        Err.Clear
                                                                   
                            strLinha = ""
                            For i = 0 To .Fields.Count - 1
                                strBuffer = IIf(IsNull(.Fields.Item(i).Value), "", .Fields.Item(i).Value)
                                
                                If .Fields.Item(i).Type = 5 Then
                                    strBuffer = Replace(Format(strBuffer, "0.00"), ",", ".")
                                End If
                                
                                If .Fields.Item(i).Type = 131 Then
                                    strBuffer = Replace(Format(strBuffer, "0.00"), ",", ".")
                                End If
                                If .Fields.Item(i).Type = 133 Then
                                    strBuffer = Format(strBuffer, "YYYY-MM-DD")
                                End If
                                
                                If .Fields.Item(i).Type = 135 Then
                                    strBuffer = Format(strBuffer, "yyyy-MM-dd hh:mm:ss")
                                End If
    
                                strBuffer = Replace(strBuffer, "\", "\\")
                                strBuffer = Replace(strBuffer, "'", "\'")
                                strBuffer = Replace(strBuffer, Chr(10), "")
                                strBuffer = Replace(strBuffer, Chr(13), "\r\n")
                                
                          
                                
                                If Trim(strBuffer) = "" Then
                                    strLinha = strLinha & ",NULL"
                                    Else
                                    strLinha = IIf(strLinha = "", strLinha, strLinha & ",") & "'" & strBuffer & "'"
                                End If
                                
                            Next i
                            
                            .MoveNext
                            
                            strLinha = "(" & strLinha & ")"
                        
                            Print #X, "INSERT INTO `" & strNomeTabela & "` VALUES " & strLinha & ";"
                           
                            lngBytesProcessados = Len(strLinha) + lngBytesProcessados
                            lblProcessoatual.Caption = lngBytesProcessados
                            
                        Loop
                        
                    End If
                    
                    .Close
                End With
            'Fim dos dados da tabela
            End If
            Print #X, "unlock tables;"
            smsg = "Desbloqueando a Tabela;"
            lstMSG.AddItem smsg
        
    
            Print #X, "#--------------------------------------------"
            smsg = "#--------------------------------------------"
            lstMSG.AddItem smsg
   
            .MoveNext
        Loop

        Print #X, ""
        Print #X, "SET FOREIGN_KEY_CHECKS=1;"
        smsg = "Ativando a checagem de Constraint;"
        lstMSG.AddItem smsg
   
        Print #X, ""
        Print #X, "# " & MSG_08 & Format(Now, MSG_04)
        smsg = "# " & MSG_08 & Format(Now, MSG_04)
        lstMSG.AddItem smsg
        
        .Close
    End With
    
    Close #X
    smsg = "Concluído !!!!!"
    lstMSG.AddItem smsg
    pb.Value = pb.Max
    pb2.Value = pb2.Max
    
End Sub
Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Iniciar Backup"
            If chkAcesso(Me, "n") = False Then
                Exit Sub
            End If
            IniciarBackup
            MsgBox "Processo concluído !!!"
        Case "Restaurar Backup"
            RestaurarBackup
        Case "Manutenção na Base de Dados"
            If RepararBD = True Then
                    MsgBox "Base de Dados reparada com sucesso!", vbInformation, "Aviso"
                Else
                    MsgBox "Erro ao reparar base de dados !", vbInformation, "Aviso"
            End If
    End Select
End Sub
