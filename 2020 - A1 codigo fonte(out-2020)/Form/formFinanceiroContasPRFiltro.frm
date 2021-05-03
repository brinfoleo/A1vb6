VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formFinanceiroContasPRFiltro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contas a Pagar / Receber - Filtro"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   12795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Plano de Contas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8520
      TabIndex        =   34
      Top             =   1320
      Width           =   4155
      Begin VB.ListBox lstPlanoContas 
         Height          =   1635
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   35
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Digite o N. Duplicata ou parte dele:"
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
      Left            =   8520
      TabIndex        =   32
      Top             =   5640
      Width           =   4155
      Begin VB.TextBox txtnDup 
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   420
         Width           =   3855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Digite o N. Numero ou parte dele:"
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
      Left            =   8520
      TabIndex        =   26
      Top             =   4560
      Width           =   4155
      Begin VB.TextBox txtNossoNumero 
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   420
         Width           =   3915
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Valor Nominal da Duplicata"
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
      Left            =   8520
      TabIndex        =   21
      Top             =   3480
      Width           =   4155
      Begin VB.TextBox txtvDuplATE 
         Height          =   285
         Left            =   2580
         MaxLength       =   20
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   420
         Width           =   1395
      End
      Begin VB.TextBox txtvDuplDE 
         Height          =   285
         Left            =   660
         MaxLength       =   20
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   420
         Width           =   1395
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Até:"
         Height          =   195
         Left            =   2220
         TabIndex        =   23
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "De:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   435
      End
   End
   Begin VB.CommandButton btoCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8520
      TabIndex        =   18
      Top             =   7440
      Width           =   4155
   End
   Begin VB.CommandButton btoAplicar 
      Caption         =   "&Aplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8520
      TabIndex        =   17
      Top             =   6780
      Width           =   4155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cedente / Sacado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   8295
      Begin VB.OptionButton Option1 
         Caption         =   "contenha:"
         Height          =   195
         Index           =   1
         Left            =   1980
         TabIndex        =   37
         Top             =   2820
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "inicia com"
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   36
         Top             =   2820
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.TextBox txtNome 
         Height          =   315
         Left            =   3060
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2760
         Width           =   4995
      End
      Begin VB.ListBox lstNomes 
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   240
         Width           =   7995
      End
      Begin VB.Label Label3 
         Caption         =   "Nome que "
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   2820
         Width           =   855
      End
   End
   Begin VB.Frame frmPeriodo 
      Caption         =   "Periodo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   120
      TabIndex        =   5
      Top             =   180
      Width           =   4275
      Begin VB.OptionButton optPerido 
         Caption         =   "Emissão"
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.OptionButton optPerido 
         Caption         =   "Vencimento"
         Height          =   195
         Index           =   1
         Left            =   1860
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton optPerido 
         Caption         =   "Quitação"
         Height          =   195
         Index           =   2
         Left            =   3060
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpDtInicio 
         Height          =   315
         Left            =   480
         TabIndex        =   9
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   117964801
         CurrentDate     =   40557
      End
      Begin MSComCtl2.DTPicker dtpDtFinal 
         Height          =   315
         Left            =   2460
         TabIndex        =   10
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   115343361
         CurrentDate     =   40557
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "De:"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Até:"
         Height          =   195
         Left            =   2100
         TabIndex        =   11
         Top             =   480
         Width           =   315
      End
   End
   Begin VB.Frame frmExibir 
      Caption         =   "Exibir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   4500
      TabIndex        =   2
      Top             =   180
      Width           =   8175
      Begin VB.ComboBox cboContaFV 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   600
         Width           =   2715
      End
      Begin VB.ComboBox cboTipoConta 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   2715
      End
      Begin VB.ComboBox cboTipoContaCriterio 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2715
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Critério:"
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Quitação:"
         Height          =   195
         Left            =   4320
         TabIndex        =   29
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   420
         TabIndex        =   28
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame frmTipo 
      Caption         =   "Tipo de Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4095
      Begin VB.ListBox lstTipoDocumento 
         Height          =   1635
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Centro de Custos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4320
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
      Begin VB.ListBox lstCentroCustos 
         Height          =   1635
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "formFinanceiroContasPRFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQLFiltro As String

Private Sub btoAplicar_Click()
    MontarFiltroSQL
    Unload Me
End Sub

Private Sub btoCancelar_Click()
    sSQLFiltro = ""
    Unload Me
End Sub

Private Sub cboTipoConta_DropDown()
    cboTipoConta.Clear
    cboTipoConta.AddItem "(TODAS)"
    cboTipoConta.AddItem "A PAGAR"
    cboTipoConta.AddItem "A RECEBER"
End Sub

Private Sub cboTipoContaCriterio_DropDown()
    cboTipoContaCriterio.Clear
    cboTipoContaCriterio.AddItem "(TODAS)"
    cboTipoContaCriterio.AddItem "EM ABERTO"
    cboTipoContaCriterio.AddItem "QUITADAS"
End Sub
Private Sub cboContaFV_DropDown()
    cboContaFV.Clear
    cboContaFV.AddItem "(TODAS)"
    cboContaFV.AddItem "FIXAS"
    cboContaFV.AddItem "VARIAVEIS"
End Sub

Private Sub MontarFiltroSQL()
    Dim ContaPR         As String
    Dim DtQuitacao      As String
    Dim ContaFV         As String
    Dim tpPeriodo       As String 'Pega o tipo de periodo Emissao ou Vencimento
    Dim tpDoc           As String 'Tipo de documento apresentado
    Dim c               As Integer 'Contador TMP
    Dim cCustos         As String 'Centro de Custos
    Dim pContas         As String 'Plano de Contas
    Dim Sacado          As String 'Nome dos sacados/cedentes
    Dim vDuplicata      As String 'Intervalo do valor nominal da duplicata
    Dim vNossoNumero    As String 'Armazena Nosso Numero
    Dim nDupl           As String 'Armazena o numero da Duplicata ou parte dela
    
    '****** Periodo ******************************************
    If optPerido.Item(0).Value = True Then
            tpPeriodo = "Emissao"
        ElseIf optPerido.Item(1).Value = True Then
            tpPeriodo = "Vencimento"
        ElseIf optPerido.Item(2).Value = True Then
            tpPeriodo = "DataQuitacao"
    End If
    '*********************************************************************************
    '***** TIPO DE Periodo *****************************
    If Trim(tpPeriodo) <> "" Then
    tpPeriodo = tpPeriodo & " >= '" & Format(dtpDtInicio.Value, "YYYY-MM-DD") & _
            "' AND " & tpPeriodo & " <= '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & "'"
    End If
    '*********************************************************************************
    '***** TIPO DE CONTAS FIXAS / VARIAVEIS *******
    Select Case cboContaFV.Text
        Case "(TODAS)"
            ContaFV = ""
        Case "FIXAS"
            ContaFV = "FixoVariavel = 'F'"
        Case "VARIAVEIS"
            ContaFV = "FixoVariavel = 'V'"
        Case Else
            ContaFV = ""
    End Select
    
    '***** TIPO DE CONTAS PAGAR / RECEBER *******
    Select Case cboTipoConta.Text
        Case "(TODAS)"
            ContaPR = ""
        Case "A PAGAR"
            ContaPR = "ContaPR = 'P'"
        Case "A RECEBER"
            ContaPR = "ContaPR = 'R'"
        Case Else
            ContaPR = ""
    End Select
    '**************** QUITACAO *************
    Select Case cboTipoContaCriterio.Text
        Case "(TODAS)"
            DtQuitacao = ""
        Case "EM ABERTO"
            DtQuitacao = "DataQuitacao IS NULL"
        Case "QUITADAS"
            DtQuitacao = "dataQuitacao IS NOT NULL"
        Case Else
            DtQuitacao = ""
    End Select
    
   
    
    
    '********************* Tipo de Documento *****************************************
    tpDoc = ""
    For c = 0 To lstTipoDocumento.ListCount - 1
        With lstTipoDocumento
            If .Selected(c) = True Then
                tpDoc = tpDoc & IIf(Trim(tpDoc) = "", "", ",") & IIf(pgIdTipoDoc(.List(c)) = 0, "", pgIdTipoDoc(.List(c)))
            End If
        End With
    Next
    If Trim(tpDoc) <> "" Then
        tpDoc = " TpDocumento IN (" & tpDoc & ")"
    End If
    '**********************************************************************************
    '********************* Centro de Custos *****************************************
    cCustos = ""
    For c = 0 To lstCentroCustos.ListCount - 1
        With lstCentroCustos
            If .Selected(c) = True Then
                cCustos = cCustos & IIf(Trim(cCustos) = "", "", ",") & IIf(pgIdcCusto(.List(c)) = 0, "", pgIdcCusto(.List(c)))
            End If
        End With
    Next
    If Trim(cCustos) <> "" Then
        cCustos = " CentroCusto IN (" & cCustos & ")"
    End If
    '**********************************************************************************
    '********************* Plano de Contas *****************************************
    pContas = ""
    For c = 0 To lstPlanoContas.ListCount - 1
        With lstPlanoContas
            If .Selected(c) = True Then
                pContas = pContas & IIf(Trim(pContas) = "", "planoContasCodigo LIKE '", " OR planoContasCodigo LIKE '") & Trim(Mid(.List(c), 1, InStr(.List(c), "-") - 1)) & "%'"
            End If
        End With
    Next
    'If Trim(pContas) <> "" Then
    '    pContas = " planoContasCodigo IN (" & pContas & ")"
    'End If
    '**********************************************************************************
    '********************************* Sacado *****************************************
    Sacado = ""
    For c = 1 To lstNomes.ListCount - 1
        With lstNomes
            If .Selected(c) = True Then
                Sacado = Sacado & IIf(Trim(Sacado) = "", "", ",") & "'" & .List(c) & "'" ', "", pgIdcCusto(.List(c)))
            End If
        End With
    Next
    If Trim(Sacado) <> "" Then
        Sacado = " Nome IN (" & Sacado & ")"
    End If
    '**********************************************************************************
    '********************* Valor da duplicata *****************************************
    vDuplicata = ""
    If Trim(txtvDuplDE.Text) <> "" And Trim(txtvDuplATE.Text) <> "" Then
        vDuplicata = "vlDuplicata >= '" & txtvDuplDE.Text & "' AND vlDuplicata <= '" & txtvDuplATE.Text & "'"
    End If
    '**********************************************************************************
    '*************************** Nosso Numero *****************************************
    vNossoNumero = ""
    If Trim(txtNossoNumero.Text) <> "" Then
        vDuplicata = "NossoNumero LIKE '%" & Trim(txtNossoNumero.Text) & "%'"
    End If
    '**********************************************************************************
    '*************************** Numero Duplicata *************************************
    nDupl = ""
    If Trim(txtnDup.Text) <> "" Then
        nDupl = "NumDuplicata LIKE '%" & Trim(txtnDup.Text) & "%'"
    End If
    
    '**********************************************************************************
    
    sSQLFiltro = "SELECT * FROM FinanceiroContasPRCadastro " & _
            "WHERE ID_Empresa = " & ID_Empresa & _
            IIf(tpPeriodo = "", "", " AND " & tpPeriodo) & _
            IIf(ContaPR = "", "", " AND " & ContaPR) & _
            IIf(ContaFV = "", "", " AND " & ContaFV) & _
            IIf(DtQuitacao = "", "", " AND " & DtQuitacao) & _
            IIf(tpDoc = "", "", " AND " & tpDoc) & _
            IIf(cCustos = "", "", " AND " & cCustos) & _
            IIf(pContas = "", "", " AND " & pContas) & _
            IIf(Sacado = "", "", " AND " & Sacado) & _
            IIf(vDuplicata = "", "", " AND " & vDuplicata) & _
            IIf(vNossoNumero = "", "", " AND " & vNossoNumero) & _
            IIf(nDupl = "", "", " AND " & nDupl)
            '& " ORDER BY " & TpPeriodo
End Sub
Public Function filtro() As String
    Me.Show 1
    filtro = sSQLFiltro
    
End Function
Private Sub ListarDocumentos()
    Dim Rst As Recordset
    lstTipoDocumento.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroTipoDocumento WHERE ID_Empresa = " & ID_Empresa & " ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum documento cadastrado"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                lstTipoDocumento.AddItem Rst.Fields("Descricao")
                'lstTipoDocumento.Selected(lstTipoDocumento.ListCount - 1) = True
                Rst.MoveNext
            Loop
            
    End If
    
    
End Sub
Private Sub ListarCentroCustos()
    Dim Rst As Recordset
    lstCentroCustos.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroCentroCustos WHERE  ID_Empresa = " & ID_Empresa & " ORDER BY Descricao")
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum documento cadastrado"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                lstCentroCustos.AddItem Rst.Fields("Descricao")
                'lstCentroCustos.Selected(lstCentroCustos.ListCount - 1) = True
                Rst.MoveNext
            Loop
            
    End If
    
    
End Sub
Private Sub listarPlanoContas()
    Dim Rst As Recordset
    lstPlanoContas.Clear
    Set Rst = RegistroBuscar("SELECT * FROM FinanceiroPlanoContas WHERE ID_Empresa = " & ID_Empresa & " ORDER BY codigo")
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhum documento cadastrado"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                lstPlanoContas.AddItem Rst.Fields("codigo") & " - " & Rst.Fields("Descricao")
                'lstCentroCustos.Selected(lstCentroCustos.ListCount - 1) = True
                Rst.MoveNext
            Loop
            
    End If
    
    
End Sub
Private Sub TipoContaExibir()
    cboTipoConta.AddItem "(TODAS)"
    cboTipoConta.Text = cboTipoConta.List(0)
    
    cboTipoContaCriterio.AddItem "(TODAS)"
    cboTipoContaCriterio.Text = cboTipoContaCriterio.List(0)
    
    cboContaFV.AddItem "(TODAS)"
    cboContaFV.Text = cboContaFV.List(0)
End Sub

Private Sub Form_Load()
    optPerido(0).Value = False
    optPerido(1).Value = False
    optPerido(2).Value = False
    dtpDtInicio.Value = Date
    dtpDtFinal.Value = Date
    txtNome.Text = ""
    txtvDuplDE.Text = ""
    txtvDuplATE.Text = ""
    txtNossoNumero.Text = ""
    txtnDup.Text = ""
    
    ListarDocumentos
    ListarCentroCustos
    listarPlanoContas
    TipoContaExibir
    ListarNomes
End Sub
Private Sub ListarNomes(Optional criterio As String)
    Dim Rst     As Recordset
    Dim sSQL    As String
    If Trim(criterio) = "" Then
            sSQL = "SELECT DISTINCT Nome FROM FinanceiroContasPRCadastro WHERE ID_Empresa = " & ID_Empresa & " ORDER BY Nome"
        Else
            If Option1(0).Value = True Then
                    sSQL = "SELECT DISTINCT Nome FROM FinanceiroContasPRCadastro WHERE ID_Empresa = " & ID_Empresa & " AND Nome LIKE '" & criterio & "%' ORDER BY Nome"
                Else
                    sSQL = "SELECT DISTINCT Nome FROM FinanceiroContasPRCadastro WHERE ID_Empresa = " & ID_Empresa & " AND Nome LIKE '%" & criterio & "%' ORDER BY Nome"
            End If
    End If
    
    Set Rst = RegistroBuscar(sSQL)
    lstNomes.Clear
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            lstNomes.AddItem "(Todos)"
            Do Until Rst.EOF
                lstNomes.AddItem IIf(IsNull(Rst.Fields("Nome")), "", Rst.Fields("Nome"))
                Rst.MoveNext
            Loop
    End If
    
End Sub


Private Sub lstNomes_ItemCheck(Item As Integer)
 Dim opcao As Boolean
    Dim i As Integer
    If Item <> 0 Then Exit Sub
    If lstNomes.Selected(0) = True Then
            opcao = True
        Else
            opcao = False
    End If
    For i = 1 To lstNomes.ListCount - 1
        lstNomes.Selected(i) = opcao
    Next
End Sub

Private Sub Option1_Click(Index As Integer)
    txtNome_Change
End Sub

Private Sub txtNome_Change()
    ListarNomes (Trim(txtNome.Text))
End Sub


Private Sub txtvduplATE_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvDuplATE.Text, KeyAscii, cDecMoeda)
End Sub

Private Sub txtvduplDE_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(txtvDuplDE.Text, KeyAscii, cDecMoeda)
End Sub
