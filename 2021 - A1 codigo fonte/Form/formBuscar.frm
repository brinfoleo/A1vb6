VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form formBuscar 
   Caption         =   "Busca"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid grdBusca 
      Height          =   4455
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7858
      _Version        =   393216
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
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   60
      TabIndex        =   1
      Top             =   4800
      Width           =   8655
      Begin VB.OptionButton optBusca 
         Caption         =   "Inicia com"
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   0
         Top             =   1020
         Width           =   2115
      End
      Begin VB.OptionButton optBusca 
         Caption         =   "Qualquer parte do campo"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   540
         Width           =   6855
      End
      Begin VB.CommandButton btAplicar 
         Caption         =   "&Aplicar"
         Height          =   435
         Left            =   5160
         TabIndex        =   3
         Top             =   960
         Width           =   1635
      End
      Begin VB.CommandButton btCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   435
         Left            =   6900
         TabIndex        =   2
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "Campo de busca:"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Texto de Busca:"
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbCampoBusca 
         Caption         =   "lbCampoBusca"
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
         Left            =   1500
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "formBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst                 As Recordset 'ADODB.Recordset
Dim Tabela              As String
Dim CampoBusca          As String
Dim CamposBusca         As String
Dim Ordenar             As String
Dim resultadoBusca      As Integer
Dim sizeColuna(100)     As Integer
Dim criterio            As String

Public Function IniciarBusca(nomeTabela As String, _
                             Optional sColunas As String, _
                             Optional DefaultCampoBusca As String, _
                             Optional DefaultTextoBusca As String, _
                             Optional CampoOrdenar As String, _
                             Optional sCriterio As String) As Integer
                             
    If nomeTabela = "" Then Exit Function
    
    If Trim(sColunas) = "" Then
            Select Case LCase(nomeTabela)
                Case "estoqueproduto"
                    CamposBusca = "Id," & "Descricao, Referencia, CodigoBarras, NCM,IPIAliquota,ICMSCST,saldo,ID_Empresa,Deposito,InformacoesComplementares,status"
                    If PgDadosUsuario(ID_Usuario).SuperUsuario = 0 Or PgDadosConfig.EstoqueSUverDepositos = 0 Then
                        sCriterio = "status='ATIVO' AND Deposito=" & ID_Deposito
                    End If
                    
                Case "clientes"
                    CamposBusca = "Id," & "xNome,Doc,IE,xlgr,nro,xcpl,xbairro,xmun,uf,fone,email,emailnfe,ID_Empresa"
                Case "Transportadoras"
                    CamposBusca = "Id," & "xNome,Fant,xEnder,bairro,mun,uf,fone,ID_Empresa"
                'Case "Empresas"
                    'CamposBusca = "Id," & "Nome,Fant,CNPJ,Lgr,Nro,Cpl,bairro,mun,uf,fone,ID_Empresa"
                Case "fornecedores"
                    CamposBusca = "Id," & "xNome,Fant,Doc,Lgr,Nro,Cpl,bairro,mun,uf,fone,ID_Empresa"
                 Case "faturamentopv"
                    CamposBusca = "Id," & "Cliente,emissao,ID_Empresa"
                Case "faturamentonfe"
                    CamposBusca = "Id," & "IdNFe,ide_nNF,ide_dEmi,dest_xNome,canc_nProt,canc_xJust,ID_Empresa"
                Case "faturamentonfeentrada"
                    CamposBusca = "Id," & "IdNFe,ide_nNF,ide_dEmi,emit_xNome, total_vNF,ID_Empresa" ',canc_nProt,canc_xJust"
                Case "financeirocondicoespagamento"
                    CamposBusca = "Id," & "Descricao,ID_Empresa"
'                Case "tributacaocest"
'                    CamposBusca = "Id," & "Descricao,ID_Empresa"
                Case Else
                    CamposBusca = "*"
                    'MsgBox "sem paramentros"
            End Select
        Else
            CamposBusca = IIf(Trim(sColunas) = "", "*", "Id," & sColunas)
    End If
    Ordenar = CampoOrdenar
    Tabela = nomeTabela
    criterio = sCriterio
    resultadoBusca = 0
    'Busca Default
    If Trim(DefaultCampoBusca) <> "" Then
        CampoBusca = DefaultCampoBusca
        lbCampoBusca.Caption = UCase(DefaultCampoBusca)
        Text1.Text = DefaultTextoBusca
    End If
    
    
    formBuscar.Show 1
    IniciarBusca = resultadoBusca
    
End Function

Private Sub btAplicar_Click()
    Unload Me
End Sub

Private Sub btCancelar_Click()
    resultadoBusca = 0
    Unload Me
End Sub



Private Sub Form_Activate()
    On Error Resume Next
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    
    LimpaFormulario formBuscar
    PreencherGrid
    CampoBusca = grdBusca.Columns(1).Caption
    lbCampoBusca.Caption = UCase(CampoBusca)
End Sub

Private Sub PreencherGrid(Optional strSQL As String)
    Dim i           As Integer
    Dim largCol     As Long
    Dim sSQL        As String
    Select Case LCase(Tabela)
        Case "estoqueproduto"
            'sSQL = "SELECT " & CamposBusca & _
                   " FROM " & Tabela & _
                   " WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & _
                   IIf(Trim(strSQL) = "", "", " AND " & strSQL)
            sSQL = "SELECT " & CamposBusca & _
                   " FROM " & Tabela & _
                   " WHERE ID_Empresa = " & ID_Empresa & _
                   IIf(Trim(strSQL) = "", "", " AND " & strSQL)
            Ordenar = "ORDER BY Descricao"
        Case "tributacaocest"
            sSQL = "SELECT " & CamposBusca & _
                   " FROM " & Tabela & _
                   IIf(Trim(strSQL) = "", "", " WHERE " & strSQL)
            'Ordenar = "ORDER BY Descricao"
        Case Else
            sSQL = "SELECT " & CamposBusca & _
                   " FROM " & Tabela & _
                   " WHERE ID_Empresa = " & ID_Empresa & _
                   IIf(Trim(strSQL) = "", "", " AND " & strSQL)
    End Select
    
            If InStr(sSQL, "WHERE") <> 0 And Trim(criterio) <> "" Then
                sSQL = sSQL & " AND " & criterio
            End If
    sSQL = sSQL & IIf(Trim(Ordenar) = "", "", " " & Ordenar)
    Set Rst = RegistroBuscar(sSQL)
    If Rst Is Nothing Then
        grdBusca.Enabled = False
        Text1.Enabled = False
        Me.Caption = "Busca - [ 00000 Registros]"
        Exit Sub
    End If
    
    If Rst.BOF And Rst.EOF Then
            grdBusca.Enabled = False
            'Text1.Enabled = False
            Me.Caption = "Busca - [ 00000 Registros]"
        Else
            grdBusca.Enabled = True
            Rst.MoveLast
            Me.Caption = "Busca - [ " & Left(String(5, "0"), 5 - Len(Trim(Rst.RecordCount))) & Trim(Rst.RecordCount) & " Registros]"
            Rst.MoveFirst
    End If
    Set grdBusca.DataSource = Rst.DataSource
    'Dim i As Integer
    With grdBusca
        .AllowUpdate = False
        For i = 0 To .Columns.Count - 1
            If .Columns(i).Caption = "Id_Empresa" Then .Columns(i).Visible = False
            If .Columns(i).Caption = "DtHr" Then .Columns(i).Visible = False
        Next

        For i = 1 To .Columns.Count - 1
            largCol = .Columns(i).Width 'Largura da coluna atual
            .Columns(i).Width = IIf(sizeColuna(i) = 0, largCol, sizeColuna(i)) 'Dimenciona a coluna
            
            .Columns(0).Width = IIf(sizeColuna(0) = 0, TextWidth("99999"), sizeColuna(0))
            .Columns(1).Width = IIf(sizeColuna(1) = 0, .Width / 2, sizeColuna(1))

        Next
        
    End With
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    grdBusca.Width = Me.ScaleWidth - 150
    grdBusca.Height = Me.ScaleHeight - (150 + Frame1.Height)
    
    Frame1.Top = grdBusca.Height + 100
    
    Me.Width = IIf(Me.Width < 8925, 8925, Me.Width)
    Me.Height = IIf(Me.Height < 6900, 6900, Me.Height)
End Sub

Private Sub grdBusca_Click()
    selecionarRegistro
    Text1.SetFocus
End Sub
Private Sub selecionarRegistro()
    resultadoBusca = grdBusca.Columns(0).Text
End Sub
Private Sub grdBusca_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    sizeColuna(ColIndex) = grdBusca.Columns(ColIndex).Width
End Sub

Private Sub grdBusca_DblClick()
    btAplicar_Click
End Sub

Private Sub grdBusca_HeadClick(ByVal ColIndex As Integer)
    CampoBusca = grdBusca.Columns(ColIndex).Caption
    lbCampoBusca.Caption = UCase(CampoBusca)
    '******* Pega o Tipo de Campo *****************
    'Dim tipoCampo           As Integer
    'tipoCampo = Rst.Fields(campoBusca).Type
    '**********************************************
End Sub

Private Sub grdBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        selecionarRegistro
        Unload Me
    End If
End Sub

'Private Sub grdBusca_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        selecionarRegistro
'        Unload Me
'    End If
'End Sub

Private Sub Text1_Change()
    Dim sSQL    As String
    Dim sBusca  As String
    Dim sBtmp   As String
    Dim parte   As String
    If Trim(Text1.Text) = "" Then
            PreencherGrid
        Else
            If optBusca(0).Value = True Then
                    sBtmp = ""
                    sBusca = Replace(Trim(Text1.Text), " ", "|") & "|"
                    Do Until InStr(sBusca, "|") = 0
                        parte = Trim(Mid(sBusca, 1, InStr(sBusca, "|") - 1))
                        parte = Replace(parte, "'", "''")
                            
                        sBtmp = IIf(Trim(sBtmp) = "", "", sBtmp & " AND ") & CampoBusca & " LIKE '%" & Trim(parte) & "%'"
                        sBusca = Mid(sBusca, InStr(sBusca, "|") + 1, Len(sBusca))
                    Loop
                    sSQL = sBtmp '& " ORDER BY " & CampoBusca
                Else
                    sSQL = CampoBusca & " LIKE '" & Text1.Text & "%'" '&  ORDER BY " & CampoBusca
            End If
            '**********
            'sSQL = sSQL & " ORDER BY " & CampoBusca
            'If InStr(sSQL, "WHERE") <> 0 And Trim(Criterio) <> "" Then
            '    sSQL = sSQL & " AND " & Criterio
            'End If
            '******
            PreencherGrid (sSQL)
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then
        If grdBusca.Enabled = True Then
            grdBusca.SetFocus
        End If
    End If
End Sub
