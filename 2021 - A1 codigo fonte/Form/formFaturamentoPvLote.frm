VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formFaturamentoPvLote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerar Prevenda em Lote"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btoGerarPreVenda 
      Caption         =   "&Gerar Pré-Venda"
      Height          =   555
      Left            =   6300
      TabIndex        =   4
      Top             =   5400
      Width           =   2115
   End
   Begin VB.CommandButton btoExcluirCliente 
      Caption         =   "&Excluir Cliente (Delete)"
      Height          =   555
      Left            =   2340
      TabIndex        =   3
      Top             =   5400
      Width           =   2115
   End
   Begin VB.CommandButton btoIncluirCliente 
      Caption         =   "&Incluir Cliente (Insert)"
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   2115
   End
   Begin VB.Frame Frame1 
      Height          =   5235
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8355
      Begin VB.TextBox txtQtd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2340
         Width           =   1155
      End
      Begin VB.TextBox txtPvModelo 
         Height          =   315
         Left            =   7080
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid msfgGrade 
         Height          =   4455
         Left            =   60
         TabIndex        =   1
         Top             =   720
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FormatString    =   "^id|<Cliente                                       |^PV            "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pré-Venda Modelo:"
         Height          =   195
         Left            =   5340
         TabIndex        =   5
         Top             =   300
         Width           =   1635
      End
   End
End
Attribute VB_Name = "formFaturamentoPvLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim idCli       As Integer
Dim idLote      As Long 'Lote de PV gerado

Dim pvOriginal  As Integer
Dim nCol        As Integer 'Numero de colunas
Dim qCol(100)   As Integer 'Quantidades de itens por coluna
Dim idPV        As Integer 'Numero da PV Selecionada



Private Sub lForm()
    LimpaFormulario Me
    txtQtd.Visible = False
    With msfgGrade
        .Rows = 1
        .Cols = 3
        .FormatString = "^id|<Cliente                                       |^PV            "
    End With
End Sub

Private Sub btoExcluirCliente_Click()
    RemoverItem
End Sub

Private Sub btoGerarPreVenda_Click()
    
    Dim idNPV       As Integer 'nova PV
    Dim sSQL        As String
    Dim cItem       As Integer
    Dim qItem(50)   As Variant
    Dim i           As Integer
    
    'idLote = Format(Now(), "YYYYMMDDHHMMSS")
    With msfgGrade
        For i = 1 To .Rows - 1
            idCli = .TextMatrix(i, 0)
            idPV = .TextMatrix(i, 2)
            For cItem = 3 To .Cols - 1
                qItem(cItem - 2) = IIf(Trim(.TextMatrix(i, cItem)) = "", 0, Trim(.TextMatrix(i, cItem)))
            Next
            cItem = cItem - 3
            idNPV = formFaturamentoPV.ClonarPV(pvOriginal, idPV, cItem, qItem)
            
            'So cadastra numero de lote para Novas PV
            'If idPV = 0 Then
            '    sSQL = "UPDATE FaturamentoPV " & _
            '            "SET idLote='" & idLote & "' " & _
            '            "WHERE id=" & idNPV
            '    BD.Execute sSQL
            'End If
            .TextMatrix(i, 2) = idNPV
        Next
    End With
    
End Sub

Private Sub btoIncluirCliente_Click()
    idCli = formBuscar.IniciarBusca("Clientes") ', "xNome,xlgr,nro,xcpl,xbairro,xmun,uf,fone")
    
    If idCli = 0 Then
            'LimpaFormulario Me
        Else
            incluirCliente 0
    End If

End Sub

Private Sub Form_Load()
  lForm
    
End Sub

Private Sub incluirCliente(Optional pv As Integer)
    Dim i As Integer
    Dim c As Integer 'Coluna
    
    If nCol = 0 Then Exit Sub
    
    With msfgGrade
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = idCli
        .TextMatrix(.Rows - 1, 1) = PgDadosCliente(idCli).Nome
        .TextMatrix(.Rows - 1, 2) = pv
        For i = 3 To .Cols - 1
            .TextMatrix(.Rows - 1, i) = qCol(i - 3)
        Next
    End With
    idCli = 0
    'pvOriginal = 0
    
End Sub
Private Sub msfggrade_EnterCell()

    With msfgGrade
        If .Enabled = False Then Exit Sub
        If .MouseCol < 3 Or .MouseRow = 0 Then
            txtQtd.Visible = False
            Exit Sub
        End If
        idPV = .TextMatrix(.Row, 2)
        txtQtd.Top = .Top + .CellTop
        txtQtd.Left = .Left + .CellLeft
        txtQtd.Width = .CellWidth
        txtQtd.Height = .CellHeight
        txtQtd.Text = Trim(.TextMatrix(.Row, .Col))
        txtQtd.Visible = True
    End With
End Sub

Private Sub msfggrade_LeaveCell()
    Dim dtCalc      As String 'Registra a data para o qual o boleto ta sendo calculado

    
    With msfgGrade
        If .Enabled = False Then Exit Sub
        If .Col < 3 And .Row = 1 Or txtQtd.Visible = False Then
            Exit Sub
        End If
        .TextMatrix(.Row, .Col) = txtQtd.Text
        'dtCalc = .TextMatrix(.Row, 8) 'IIf(IsNull(Rst.Fields("dataquitacao")), Date, Rst.Fields("dataquitacao")) 'Checa a data de quitacao do documento
        '.TextMatrix(.Row, 9) = IIf(CDate(dtCalc) - CDate(.TextMatrix(.Row, 7)) < 0, 0, CDate(dtCalc) - CDate(.TextMatrix(.Row, 7)))

        '.TextMatrix(.Row, 10) = ConvMoeda(Val(AtualizaCobranca(IdReg, dtCalc).vMulta) + Val(AtualizaCobranca(IdReg, dtCalc).vMora))
        '.TextMatrix(.Row, 11) = ConvMoeda(.TextMatrix(.Row, 8))

        '.TextMatrix(.Row, 11) = ConvMoeda(Val(ChkVal(Left(.TextMatrix(.Row, 6), Len(.TextMatrix(.Row, 6)) - 1), 0, cDecMoeda)) + Val(ChkVal(.TextMatrix(.Row, 10), 0, cDecMoeda)))
    End With
End Sub


Private Sub msfgGrade_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If msfgGrade.TextMatrix(msfgGrade.Row, 0) = "ID" Or msfgGrade.Rows = 1 Then Exit Sub 'Or Trim(msfgGrade.TextMatrix(msfgGrade.Row, 2)) = "" Then Exit Sub
    If KeyCode = 45 Then 'Adicionar linha em branco
        incluirCliente
    End If
    If KeyCode = 46 Then 'Tecla Delete
        RemoverItem
    End If
   
End Sub
Private Sub RemoverItem()

    If MsgBox("Deseja realmente remover este item?", vbYesNo, "Removendo Item do Pedido") = vbYes Then
        If msfgGrade.Rows = 2 Then
                msfgGrade.Rows = 1
            Else
                msfgGrade.RemoveItem msfgGrade.Row
        End If
    End If
End Sub

Private Sub txtPvModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
    If KeyAscii = 13 Then
            pvOriginal = Trim(txtPvModelo.Text)
            carregarPreVenda
        Else
            pvOriginal = 0
        
    End If
    
End Sub
Private Sub carregarPreVenda()
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    'Limpa o formulario
    lForm
    txtPvModelo.Text = pvOriginal
    
    'Carrega PreVenda Modelo
    sSQL = "SELECT * FROM FaturamentoPV WHERE ID_Empresa = " & ID_Empresa & " AND Id = " & pvOriginal
    
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar Pré-Venda", vbInformation, App.EXEName
            Rst.Close
            Exit Sub
        Else
            Rst.MoveFirst
            idCli = Rst.Fields("idCliente")
    End If
    Rst.Close
    
    
    sSQL = "SELECT * FROM faturamentopvitens WHERE ID_Empresa = " & ID_Empresa & " AND idpv= " & pvOriginal
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar o(s) item(ns) da Pré-Venda", vbInformation, App.EXEName
            Rst.Close
            Exit Sub
        Else
            Dim i As Integer
            Rst.MoveLast
            nCol = Rst.RecordCount
            msfgGrade.Cols = 3 + nCol
            Rst.MoveFirst
            i = 0
            Do Until Rst.EOF
              qCol(i) = Rst.Fields("quantidade")
              i = i + 1
              Rst.MoveNext
            Loop
            
            
    End If
    Rst.Close
    incluirCliente pvOriginal
    
End Sub

Private Sub txtQtd_KeyPress(KeyAscii As Integer)
    KeyAscii = SoNumeros(KeyAscii)
    If KeyAscii = 13 Then
        msfggrade_LeaveCell
        txtQtd.Visible = False
    End If
    
    
    
End Sub
