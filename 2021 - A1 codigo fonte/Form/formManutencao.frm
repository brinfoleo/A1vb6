VERSION 5.00
Begin VB.Form formManutencao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manutenção"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   3555
      Begin VB.CheckBox Check1 
         Caption         =   "Excluir Tabela"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1875
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Modificar Tabela"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Criar Tabela"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.CommandButton btoExecutar 
      Caption         =   "Executar"
      Height          =   435
      Left            =   2100
      TabIndex        =   0
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3495
   End
End
Attribute VB_Name = "formManutencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCriterio     As String 'Criterio da tabela
Dim nReg            As Integer
Dim vDados(199)     As Variant
Dim formulario      As Form
Dim sTabela         As String
Private Sub ExcluirTabela()
    
    'sTabela = Mid(formulario.Name, 5, Len(formulario.Name))
    
    BD.Execute "DROP TABLE " & sTabela
End Sub

Public Sub IniciarManutencao(recForm As Form)
    Set formulario = recForm
    Me.Show
End Sub

Private Function MontarString() As String
    'sTabela - Nome da Tabela
    'vDados - Contem os dados como (1) Campos, (2) Dados , (3) Tipo de dados
    '          Obs. no (3) os dados podem ser s - string, i - inerger, n - numero,
    '          d - data, t - tempo e v variante
    
    'Dim sTabela    As String
    Dim sSQL       As String
    Dim sFields    As String
    Dim sValues    As String
    Dim I          As Integer
    
    'sTabela = Mid(formulario.Name, 5, Len(formulario.Name))
    
    For I = 0 To nReg - 1 'UBound(vDados)
        
        sFields = sFields & strCriterio & vDados(I)(0) & " VARCHAR(" & vDados(I)(1) & ") default Null " & ","
'        If vDados(I)(2) = "S" Then
'                '
'                'sValues = sValues + "'" + vDados(I)(1) + "'" + ","
'                sValues = sValues & vDados(I)(1) & ","
'                'sValues = sValues + IIf(Trim(vDados(I)(1)) = "", "Null,", "'" + vDados(I)(1) + "'" + ",")
'
'             ElseIf vDados(I)(2) = "N" Then
'                If vDados(I)(1) <> "" Then
'                        sValues = sValues + vDados(I)(1) + ","
'                    Else
'                        sValues = sValues + 0 + ","
'                End If
'
'            ElseIf vDados(I)(2) = "I" Then
'                If vDados(I)(1) <> "" Then
'                        sValues = sValues + CStr(Val(vDados(I)(1))) + ","
'                    Else
'                        sValues = sValues + "0" + ","
'                End If
'
'            ElseIf vDados(I)(2) = "D" Then
'                Dim sDt As String
'                sDt = vDados(I)(1)
'                sDt = IIf(sDt = "", "Null", "")
'
'                sValues = sValues + sDt + "," 'ConverteData(vDados(I)(1)) + ","
'
'            ElseIf vDados(I)(2) = "T" Then
'                sValues = sValues + vDados(I)(1) + "," 'ConverteTempo(vDados(I)(1)) + ","
'
'            ElseIf vDados(I)(2) = "V" Then
 '               sValues = sValues + vDados(I)(0) + "=" + vDados(I)(1) + ","
'
 '       End If
'
    Next I
    sFields = Left(sFields, Len(sFields) - 1)

    sSQL = "ALTER TABLE " & sTabela & _
           " " & sFields & " "
           
    BD.Execute "CREATE TABLE IF NOT EXISTS " & sTabela & _
               " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa VARCHAR(10) default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "PRIMARY KEY (Id))"
    
    MontarString = sSQL
    BD.Execute sSQL
    Exit Function

TrataErro:
    MsgBox "Erro ao MontarString registro.                                   " & _
           vbCrLf & vbCrLf & _
           "Erro n.: " & Err.Number & _
           vbCrLf & vbCrLf & _
           "Descrição: " & Err.Description & _
           vbCrLf
End Function
Private Sub pgDadosForm()
    Dim Controle    As Control
    Dim I           As Integer
    nReg = 0
    For I = 0 To formulario.Controls.Count - 1
        Set Controle = formulario.Controls(I)
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
        'Nao funconou registro duplicado
        'If TypeOf Controle Is OptionButton Then
        '    vDados(nReg) = Array(Mid(Controle.Name, 4, Len(Controle.Name)), 1, "S")
        '    nReg = nReg + 1
        'End If
    Next
End Sub
Private Sub btoExecutar_Click()
    
    
    If Check1.Value = 0 Then
            pgDadosForm
            MontarString
            MsgBox "Criação/Manutenção no banco de dados concluido"
        Else
            ExcluirTabela
            MsgBox "Tabela EXCLUIDA"
    End If
        
    
End Sub
Private Sub Check1_Click()
    If Check1.Value = 1 Then
            Option1(0).Enabled = False
            Option1(1).Enabled = False
        Else
            Option1(0).Enabled = True
            Option1(1).Enabled = True
    End If
End Sub

Private Sub Form_Load()
    sTabela = Mid(formulario.Name, 5, Len(formulario.Name))
    strCriterio = " ADD COLUMN "
    Label1.Caption = sTabela
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            strCriterio = " ADD COLUMN "
        Case 1
            strCriterio = " MODIFY "
        
    End Select
End Sub
