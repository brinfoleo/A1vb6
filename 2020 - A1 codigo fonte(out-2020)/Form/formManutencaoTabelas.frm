VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formManutencaoTabelas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manutenção"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "E&xcluir Tabela"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   5700
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   9675
   End
   Begin VB.TextBox Text1 
      Height          =   1305
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4080
      Width           =   9675
   End
   Begin VB.CommandButton btoExecutar 
      Caption         =   "&Executar"
      Height          =   435
      Left            =   8160
      TabIndex        =   0
      Top             =   5580
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Status:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9795
   End
End
Attribute VB_Name = "formManutencaoTabelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nReg            As Integer
Dim vDados(199)     As Variant
Dim Formulario      As Form
Dim sTabela         As String
Dim SQLextra        As String

Public Sub IniciarManutencao(recForm As Form, Optional cmdSQLextra As String)
    Set Formulario = recForm
    SQLextra = cmdSQLextra
    Me.Show 1
End Sub

Private Function MontarString() As String
    
    Dim sSQL       As String
    Dim sFields    As String
    Dim sValues    As String
    Dim i          As Integer
    
    
    pb.Value = Val(pb.Value) + 1
    
    BD.Execute "CREATE TABLE IF NOT EXISTS " & sTabela & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "PRIMARY KEY (Id))"
    
    ColocarUsuID sTabela
    
    '***** CRIAR TABELA ************************************************
    sSQL = "ALTER TABLE  " & sTabela & " ADD COLUMN Id_Empresa INT default Null"
    Aplicar (sSQL)
    
    sSQL = "ALTER TABLE  " & sTabela & " ADD COLUMN DtHr VARCHAR(20) default Null"
    Aplicar (sSQL)
    
    sSQL = "ALTER TABLE  " & sTabela & " ADD COLUMN UsuID INT default Null"
    Aplicar (sSQL)
    
    For i = 0 To nReg - 1 'UBound(vDados)
        Select Case UCase(vDados(i)(2))
            Case "S"
                If vDados(i)(1) > 250 Then
                        sFields = vDados(i)(0) & " TEXT default Null"
                    Else
                        sFields = vDados(i)(0) & " VARCHAR(" & vDados(i)(1) & ") default Null"
                End If
            Case "D"
                sFields = vDados(i)(0) & " DATE default Null"
        End Select
        
        sSQL = "ALTER TABLE " & sTabela & " ADD COLUMN " & _
                " " & sFields & " "
               ' Debug.Print sFields
        Aplicar (sSQL)
    
    Next i
    
    '***** MODIFICAR TABELA *********************************************
    For i = 0 To nReg - 1

        If vDados(i)(1) > 250 Then
                sFields = vDados(i)(0) & " TEXT default Null"
            Else
                sFields = vDados(i)(0) & " VARCHAR(" & vDados(i)(1) & ") default Null"
        End If
        sSQL = "ALTER TABLE " & sTabela & " MODIFY COLUMN" & _
                " " & sFields & " "
        Aplicar (sSQL)
    Next i
End Function
Private Sub ColocarUsuID(tabelaX)
    '******************************************************************************
    'Data: 07/06/2011
    'Obs.: Excluir linha, usada para corrigir erro superior
    Aplicar ("ALTER TABLE " & tabelaX & " ADD COLUMN  UsuID INT default Null")
    '******************************************************************************
End Sub
Private Sub Aplicar(sSQL As String)
    On Error GoTo TrtCriar
    pb.Value = pb.Value + 1
    BD.Execute sSQL
    Exit Sub
TrtCriar:
    DoEvents
    Select Case Err.Number
        Case "-2147217900"
        Case 380
            'Registro Duplicado
        Case Else
            List1.AddItem "[" & Err.Number & "] - " & Err.Description & "   - COMANDO: " & sSQL
'            MsgBox Err.Description, vbInformation, Err.Number
    End Select
    'Debug.Print sSQL
    RegLog "ManutencaoTabela", Err.Number, Err.Description & " [" & sSQL & "]"
    Exit Sub
End Sub
Private Sub pgDadosForm()
    Dim Controle    As Control
    Dim i           As Integer
    nReg = 0
    For i = 0 To Formulario.Controls.Count - 1
        Set Controle = Formulario.Controls(i)
        If TypeOf Controle Is TextBox Then
            vDados(nReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), IIf(Controle.MaxLength = 0, 250, Controle.MaxLength), "S")
            nReg = nReg + 1
        End If
        If TypeOf Controle Is ComboBox Then
            vDados(nReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), 250, "S")
            nReg = nReg + 1
        End If
        If TypeOf Controle Is CheckBox Then
            vDados(nReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), 1, "S")
            nReg = nReg + 1
        End If
        If TypeOf Controle Is DTPicker Then
            vDados(nReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), 10, "D")
            nReg = nReg + 1
        End If
        'Nao funconou registro duplicado
        'If TypeOf Controle Is OptionButton Then
        '    vDados(nReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), 1, "S")
        '    nReg = nReg + 1
        'End If
    Next
End Sub
Private Sub btoExecutar_Click()
    List1.Clear
    Text1.Text = ""
    pb.min = 0
    pb.Max = nReg * 2 + 1
    pb.Value = pb.min
    
    
    
    If Check1.Value = 0 Then
            pgDadosForm
            MontarString
            If SQLextra <> "" Then
                pb.Value = pb.Value - 1
                Aplicar (SQLextra)
            End If
            MsgBox "Criação/Manutenção no banco de dados concluido!", vbInformation
        Else
            If MsgBox("Deseja realmente excluir a Tabela: " & sTabela & " e todo seu conteudo?", vbYesNo + vbQuestion, "Aviso de Exclusão") = vbYes Then
                Aplicar ("DROP TABLE " & sTabela)
                If SQLextra <> "" Then
                    pb.Value = pb.Value - 1
                    Aplicar (SQLextra)
                End If
                MsgBox "Tabela " & sTabela & " excluida com sucesso! ", vbInformation
            End If
    End If
        
    
End Sub
Private Sub Form_Load()
    sTabela = LCase(Mid(Formulario.Name, 5, Len(Formulario.Name)))
    
    Label1.Caption = "Tabela: " & UCase(sTabela)
End Sub

Private Sub List1_Click()
    Text1.Text = List1.Text
End Sub

Public Function Gerar_BD_com_Array(ByVal strForm As Form, ByVal strDados As Variant, contReg As Integer, Optional ComplNome As String) As String
    
    Dim sSQL        As String
    Dim sFields     As String
    Dim sValues     As String
    Dim i           As Integer
    Dim strTabela   As String
    
    
    nReg = contReg
    
    Set Formulario = strForm
    sTabela = LCase(Mid(strForm.Name, 5, Len(strForm.Name)))
    strTabela = LCase(sTabela & ComplNome)
    
    btoExecutar.Enabled = False
    
    pb.min = 0
    pb.Value = 0
    
    BD.Execute "CREATE TABLE IF NOT EXISTS " & strTabela & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "UsuID INT default Null," & _
               "PRIMARY KEY (Id))"
    ColocarUsuID strTabela
    '***** CRIAR TABELA ************************************************
    For i = 0 To nReg 'UBound(strDados)
        Select Case UCase(strDados(i)(2))
            Case "S" 'STRING
                If strDados(i)(1) > 250 Then
                        sFields = strDados(i)(0) & " TEXT default Null"
                    Else
                        If strDados(i)(1) = 0 Then
                            MsgBox "Manutenção de Tabela" & vbCrLf & _
                                    "Campo: " & strDados(i)(0) & " com tamanho ZERO, avise ao suporte!", vbInformation, App.EXEName
                        End If
                        sFields = strDados(i)(0) & " VARCHAR(" & strDados(i)(1) & ") default Null"
                End If
            Case "D" 'DATA
                sFields = strDados(i)(0) & " DATE DEFAULT NULL"
            Case "DC" 'Decimal
                sFields = strDados(i)(0) & " DECIMAL(" & strDados(i)(1) & ",5) DEFAULT NULL"
            Case "N" 'Numero
                sFields = strDados(i)(0) & " INT DEFAULT NULL"
            Case "L" 'DATA
                sFields = strDados(i)(0) & " REAL DEFAULT NULL"
        End Select
        sSQL = "ALTER TABLE " & strTabela & " ADD COLUMN " & _
                " " & sFields & " "
        Aplicar (sSQL)
    Next i
    
    '***** MODIFICAR TABELA *********************************************
    For i = 0 To nReg
        Select Case UCase(strDados(i)(2))
            Case "S" 'STRING
                If strDados(i)(1) > 250 Then
                        sFields = strDados(i)(0) & " TEXT default Null"
                    Else
                        sFields = strDados(i)(0) & " VARCHAR(" & strDados(i)(1) & ") default Null"
                End If
            Case "D" 'DATA
                sFields = strDados(i)(0) & " DATE DEFAULT NULL"
            Case "DC" 'Decimal
                sFields = strDados(i)(0) & " DECIMAL(" & strDados(i)(1) & ",5) DEFAULT NULL"
        End Select
        sSQL = "ALTER TABLE " & strTabela & " MODIFY COLUMN" & _
                " " & sFields & " "
        Aplicar (sSQL)
        
    Next i
    
    MsgBox "Manutenção por Array na tabela " & UCase(strTabela) & " concluida!"
    Gerar_BD_com_Array = "OK"
    Unload Me
End Function
