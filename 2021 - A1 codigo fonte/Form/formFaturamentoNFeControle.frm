VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formFaturamentoNFeGerenciador 
   Caption         =   "A1 - Controle de NF-e"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   14145
   Begin VB.TextBox txtObs 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "formFaturamentoNFeControle.frx":0000
      Top             =   4440
      Width           =   13755
   End
   Begin VB.Frame frmFiltro 
      Height          =   1395
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   13815
      Begin VB.Timer tmrAtualizacao 
         Interval        =   1000
         Left            =   2700
         Top             =   960
      End
      Begin VB.CheckBox chkAtualizar 
         Caption         =   "Atualizar automaticamente"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Value           =   1  'Checked
         Width           =   4395
      End
      Begin VB.Frame frmPeriodo 
         Caption         =   "Periodo de Emissão:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   4095
         Begin MSComCtl2.DTPicker dtpDtInicio 
            Height          =   315
            Left            =   480
            TabIndex        =   3
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55902209
            CurrentDate     =   40557
         End
         Begin MSComCtl2.DTPicker dtpDtFinal 
            Height          =   315
            Left            =   2460
            TabIndex        =   4
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55902209
            CurrentDate     =   40557
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "De:"
            Height          =   195
            Left            =   180
            TabIndex        =   6
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Até:"
            Height          =   195
            Left            =   2100
            TabIndex        =   5
            Top             =   360
            Width           =   315
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msfgNotas 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   1980
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   4154
      _Version        =   393216
      Cols            =   12
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"formFaturamentoNFeControle.frx":0008
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Danfe"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excluir NFe"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consultar Situação NF-e"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar NF-e"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Fatura(s)"
            ImageIndex      =   9
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
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":0181
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":05D3
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":08ED
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":117F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":23D1
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":2CAB
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":353D
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":3DCF
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":5021
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":533B
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":5655
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":5A4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":71FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeControle.frx":7798
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formFaturamentoNFeGerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cTempo  As Integer 'Armazena o tempo para atualizacao do grid
Dim NFe     As String ' Armazena a chave de acesso
Dim IdReg   As Integer 'Armazena o ID da nota fiscal
'Private Sub ChkStatusNfeGrid()
'    Dim i As Integer
'    With msfgNotas
'        For i = 1 To .Rows - 1
'            '.Rows = .Rows + 1
'            '.TextMatrix(.Rows - 1, 0) = Rst.Fields("Id")
'            '.TextMatrix(.Rows - 1, 1) = Rst.Fields("Ide_nNF") 'Num da nota
'            '.TextMatrix(.Rows - 1, 2) = Rst.Fields("IdNFe") 'Chave de acesso
'            '.TextMatrix(.Rows - 1, 3) = Rst.Fields("Ide_dEmi") 'Data de Emissao
'            '.TextMatrix(.Rows - 1, 4) = Rst.Fields("dest_xNome")
'            .TextMatrix(.Rows - 1, 6) = pgNumLoteEnvioNFe(.TextMatrix(i, 2))
'            .TextMatrix(.Rows - 1, 7) = pgNumReciboNFe(.TextMatrix(i, 6))
'            .TextMatrix(.Rows - 1, 8) = pgStatusNFe(.TextMatrix(i, 7))
'        Next
'    End With
'End Sub

Private Sub LoadStatusEnvio(chvNFe As String, DtEmissao As String)
    Dim Arquivo     As String 'Armazena o local do Arquivo
    Dim ConteudoXML As String 'Armazena as informacoes do arquivo xml
    Dim ConteudoTag As String 'Armazena as informacoes da Tag selecionada
    
    Dim nLote       As String 'Armazena o numero do Lote de envio gerado
    Dim nRecibo     As String 'Armazena o numero do Recibo de envio gerado
    Dim nProt       As String 'Armazena o numero do Protocolo
    Dim dhRecbto    As String 'Armazena a data e Hora do recebimento
    Dim cStat       As String 'Armazena o numero do Codigo de Status
    Dim xMotivo     As String 'Armazena o Motivo do CStat
    Dim StatusNFe   As String 'Armazena a situacao da NFe
    
    Dim nProtCanc   As String
    Dim StatusCanc  As String
    
    Dim aReg(1000)  As Variant
    Dim cReg        As Integer
    
    'Checa se houve erro na convercao do arquivo
    'Dim nNF As String
    Arquivo = PgDadosConfig.pRetorno & "\" & _
    Mid(chvNFe, 26, 9) & "_" & Mid(chvNFe, 7, 14) & "_" & _
    Format(DtEmissao, "DD") & "_" & Format(DtEmissao, "MM") & "_" & Format(DtEmissao, "YYYY") & "-nfe.err"
    ConteudoXML = LoadErroXML(Arquivo)
    If ConteudoXML = "" Then
            'StatusNFe = "Erro na conversão para XML."
        Else
            StatusNFe = ConteudoXML
            'Exit Sub
    End If
    
    
    
    'Carrega o numero do Lote que foi enviado
    Arquivo = PgDadosConfig.pRetorno & "\" & chvNFe & "-num-lot.xml"
    ConteudoXML = LoadXML(Arquivo)
    If ConteudoXML = "" Then
            'Checa se houve erro no arquivo
            Arquivo = PgDadosConfig.pRetorno & "\" & chvNFe & "-nfe.err"
            ConteudoXML = LoadErroXML(Arquivo)
            If ConteudoXML = "" Then
                    'StatusNFe = "Nenhum XML encontrado."
                    'exit sub
                Else
                    StatusNFe = ConteudoXML
                    'exit sub
            End If
        Else
            'Pega o Numero do Lote do Envio
            nLote = pgTagXML("<NumeroLoteGerado>", "</NumeroLoteGerado>", ConteudoXML)
    End If
    'Carrega NUMERO DO RECIBO de envio
    nLote = Mid(String(15, "0"), 1, 15 - Len(nLote)) & nLote
    Arquivo = PgDadosConfig.pRetorno & "\" & nLote & "-rec.xml"
    ConteudoXML = LoadXML(Arquivo)
    If ConteudoXML = "" Then
            'Checa se houve erro no arquivo
            Arquivo = PgDadosConfig.pRetorno & "\" & nLote & "-rec.err"
            ConteudoXML = LoadErroXML(Arquivo)
            If ConteudoXML = "" Then
                    'StatusNFe = ""
                    'exit sub
                Else
                    StatusNFe = ConteudoXML
                    'exit sub
            End If
        Else
            'Pega o Numero do Lote do Envio
            nRecibo = pgTagXML("<nRec>", "</nRec>", ConteudoXML)
    End If
    'Carrega o NUMERO DE PROTOCOLO Status e Motivo da NFe
     Arquivo = PgDadosConfig.pRetorno & "\" & nRecibo & "-pro-rec.xml"
    ConteudoXML = LoadXML(Arquivo)
    If ConteudoXML = "" Then
            'Checa se houve erro no arquivo
            Arquivo = PgDadosConfig.pRetorno & "\" & nRecibo & "-pro-rec.err"
            ConteudoXML = LoadErroXML(Arquivo)
            If ConteudoXML = "" Then
                    'StatusNFe = ""
                    'exit sub
                Else
                    StatusNFe = ConteudoXML
                    'exit sub
            End If
        Else
            'Pega o Numero do Protocolo Data e Hora do recebimento
            
            Dim tmp As String
            tmp = pgTagXML("<infProt Id", "</infProt>", ConteudoXML)
            ConteudoXML = tmp
            nProt = pgTagXML("<nProt>", "</nProt>", ConteudoXML)
            cStat = pgTagXML("<cStat>", "</cStat>", ConteudoXML)
            xMotivo = pgTagXML("<xMotivo>", "</xMotivo>", ConteudoXML)
            If Trim(nProt) <> "" Then
                dhRecbto = pgTagXML("<dhRecbto>", "</dhRecbto>", ConteudoXML)
                dhRecbto = Format(Mid(dhRecbto, 1, InStr(dhRecbto, "T") - 1), "DD/MM/YYYY") & " " & Mid(dhRecbto, InStr(dhRecbto, "T") + 1, Len(dhRecbto))
            End If
    End If
    
    'Private Function pgStatusCancNFe(chvNFe As String) As String
    
    'Checa se a NFe Foi Cancelada
    Arquivo = PgDadosConfig.pRetorno & "\" & chvNFe & "-can.xml"
    ConteudoXML = LoadXML(Arquivo)
    If ConteudoXML = "" Then
            'Checa se houve erro no arquivo
            Arquivo = PgDadosConfig.pRetorno & "\" & chvNFe & "-can.err"
            ConteudoXML = LoadErroXML(Arquivo)
            If ConteudoXML = "" Then
                    'StatusNFe = "Nenhum XML encontrado."
                    'exit sub
                Else
                    StatusNFe = ConteudoXML
                    'exit sub
            End If
        Else
            'Pega o Numero do Lote do Envio
            nProtCanc = pgTagXML("<nProt>", "</nProt>", ConteudoXML)
            StatusCanc = pgTagXML("<xMotivo>", "</xMotivo>", ConteudoXML)
            'nLote = pgTagXML("<NumeroLoteGerado>", "</NumeroLoteGerado>", ConteudoXML)
    End If

'    Dim docNFe      As DOMDocument
'    Dim Pasta       As String
'    'Dim campo       As String
'    If Trim(chvNFe) = "" Then Exit Function
'    Pasta = PgDadosConfig.pRetorno & "\" & chvNFe & "-can.xml"
'    If Dir(Pasta) = "" Then
'            pgStatusCancNFe = ""
'        Else
'            Set docNFe = New DOMDocument
'            docNFe.resolveExternals = True
'            docNFe.validateOnParse = True
'            docNFe.async = False
'
'            'Checa se houve algum erro ao carregar
'            If docNFe.parseError.reason <> "" Then
'                MsgBox "Erro ao ler XML : " & docNFe.parseError.reason
'                Exit Function
'            End If
'            Call docNFe.Load(Pasta)
'
'            'campo = pgTagXML("<infProt Id", "</infProt>", docNFe.xml)
'            pgStatusCancNFe = pgTagXML("<cStat>", "</cStat>", docNFe.xml) & " - " & pgTagXML("<xMotivo>", "</xMotivo>", docNFe.xml)
'    End If
'End Function

    
    
    
    If CInt(nLote) <> "0" Then aReg(cReg) = Array("Lote", nLote, "S"): cReg = cReg + 1
    If Trim(nRecibo) <> "" Then aReg(cReg) = Array("nRecibo", nRecibo, "S"): cReg = cReg + 1
    If Trim(xMotivo) <> "" Then aReg(cReg) = Array("xMotivo", xMotivo, "S"): cReg = cReg + 1
    If Trim(cStat) <> "" Then aReg(cReg) = Array("cStat", cStat, "S"): cReg = cReg + 1
    If Trim(nProt) <> "" Then aReg(cReg) = Array("nProt", nProt, "S"): cReg = cReg + 1
    If Trim(dhRecbto) <> "" Then aReg(cReg) = Array("dhRecbto", dhRecbto, "S") ': cReg = cReg + 1
    
    If Trim(nProtCanc) <> "" Then aReg(cReg) = Array("canc_nProt", nProtCanc, "S"): cReg = cReg + 1
    If Trim(StatusCanc) <> "" Then aReg(cReg) = Array("canc_Status", StatusCanc, "S"): cReg = cReg + 1
    
    aReg(cReg) = Array("StatusNFe", IIf(Trim(StatusNFe) = "", xMotivo, StatusNFe), "S"): cReg = cReg + 1
    'If Trim(dhRecbto) <> 0 Then aReg(cReg) = Array("canc_nProt", e, "S"): cReg = cReg + 1
    'If Trim(f) <> "" Then aReg(cReg) = Array("canc_Status", f, "S"): cReg = cReg + 1
    cReg = cReg - 1
    If cReg >= 0 Then
        If RegistroAlterar("FaturamentoNFe", aReg, cReg, "idNFe = '" & chvNFe & "'") = False Then
            MsgBox "Erro ao atualizar Registros", vbInformation, App.Title
        End If
    End If
    
    
    
    
End Sub
Private Sub LstNotasFiscais()
    Dim Rst As Recordset
    Dim sSQL As String
    
    'Atualiza status das NFe
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ide_dEmi >='" & Format(dtpDtInicio.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                LoadStatusEnvio Rst.Fields("idNFe"), Rst.Fields("ide_dEmi")
                Rst.MoveNext
            Loop
    End If
    
    
    txtObs.Text = ""
    'Exibe NFe na tela
    msfgNotas.Rows = 1
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ide_dEmi >='" & Format(dtpDtInicio.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Exit Sub
        Else
            Rst.MoveFirst
            With msfgNotas
                Do Until Rst.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("Id")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("Ide_nNF") 'Num da nota
                    .TextMatrix(.Rows - 1, 2) = Rst.Fields("IdNFe") 'Chave de acesso
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("nProt")), " ", Rst.Fields("nProt")) 'pgnProt(.TextMatrix(.Rows - 1, 7))
                    .TextMatrix(.Rows - 1, 4) = Rst.Fields("Ide_dEmi") 'Data de Emissao
                    .TextMatrix(.Rows - 1, 5) = Left(String(5, "0"), 5 - Len(Trim(Rst.Fields("dest_idDest")))) & Rst.Fields("dest_idDest") & " - " & Rst.Fields("dest_xNome")
                    .TextMatrix(.Rows - 1, 6) = IIf(IsNull(Rst.Fields("Lote")), " ", Rst.Fields("Lote")) 'pgNumLoteEnvioNFe(.TextMatrix(.Rows - 1, 2))
                    .TextMatrix(.Rows - 1, 7) = IIf(IsNull(Rst.Fields("nRecibo")), " ", Rst.Fields("nRecibo")) 'pgNumReciboNFe(.TextMatrix(.Rows - 1, 6))
                    .TextMatrix(.Rows - 1, 8) = IIf(IsNull(Rst.Fields("StatusNFe")), " ", Rst.Fields("StatusNFe")) 'pgStatusNFe(.TextMatrix(.Rows - 1, 7))
                    .TextMatrix(.Rows - 1, 9) = IIf(IsNull(Rst.Fields("canc_nProt")), " ", Rst.Fields("canc_nProt"))
                    .TextMatrix(.Rows - 1, 10) = IIf(IsNull(Rst.Fields("canc_xJust")), " ", Rst.Fields("canc_xJust"))
                    .TextMatrix(.Rows - 1, 11) = IIf(IsNull(Rst.Fields("canc_Status")), " ", Rst.Fields("canc_Status"))
                    Rst.MoveNext
                Loop
            End With
    End If
End Sub
Private Sub chkAtualizar_Click()
    If chkAtualizar.Value = 1 Then
            cTempo = 10
            chkAtualizar.FontBold = True
            chkAtualizar.Caption = "Atualizar automaticamente..." & cTempo
            'tmrAtualizacao.Interval = 1000
        Else
            chkAtualizar.FontBold = False
    End If
End Sub
Private Sub Form_Load()
    dtpDtInicio.Value = Date
    dtpDtFinal.Value = Date
    txtObs.Text = ""
    LstNotasFiscais
End Sub
'Private Function pgStatusNFe(NumRecibo As String) As String
'    Dim docNFe      As DOMDocument
'    Dim Pasta       As String
'    Dim campo       As String
'    If Trim(NumRecibo) = "" Then Exit Function
'    Pasta = PgDadosConfig.pRetorno & "\" & NumRecibo & "-pro-rec.xml"
'    If Dir(Pasta) = "" Then
'            pgStatusNFe = "0"
'        Else
'            Set docNFe = New DOMDocument
'            docNFe.resolveExternals = True
'            docNFe.validateOnParse = True
'            docNFe.async = False
'
'            'Checa se houve algum erro ao carregar
'            If docNFe.parseError.reason <> "" Then
'                MsgBox "Erro ao ler XML : " & docNFe.parseError.reason
'                Exit Function
'            End If
'            Call docNFe.Load(Pasta)
'
'            campo = pgTagXML("<infProt Id", "</infProt>", docNFe.xml)
'            pgStatusNFe = pgTagXML("<cStat>", "</cStat>", campo) & " - " & pgTagXML("<xMotivo>", "</xMotivo>", campo)
'    End If
'End Function

'Private Function pgStatusCancNFe(chvNFe As String) As String
'    Dim docNFe      As DOMDocument
'    Dim Pasta       As String
'    'Dim campo       As String
'    If Trim(chvNFe) = "" Then Exit Function
'    Pasta = PgDadosConfig.pRetorno & "\" & chvNFe & "-can.xml"
'    If Dir(Pasta) = "" Then
'            pgStatusCancNFe = ""
'        Else
'            Set docNFe = New DOMDocument
'            docNFe.resolveExternals = True
'            docNFe.validateOnParse = True
'            docNFe.async = False
'
'            'Checa se houve algum erro ao carregar
'            If docNFe.parseError.reason <> "" Then
'                MsgBox "Erro ao ler XML : " & docNFe.parseError.reason
'                Exit Function
'            End If
'            Call docNFe.Load(Pasta)
'
'            'campo = pgTagXML("<infProt Id", "</infProt>", docNFe.xml)
'            pgStatusCancNFe = pgTagXML("<cStat>", "</cStat>", docNFe.xml) & " - " & pgTagXML("<xMotivo>", "</xMotivo>", docNFe.xml)
'    End If
'End Function


'Private Function pgnProt(NumRecibo As String) As String
'    Dim docNFe      As DOMDocument
'    Dim Pasta       As String
'    Dim campo       As String
'    Dim Recibo      As String
'    Dim dhRec       As String
'
'    If Trim(NumRecibo) = "" Then Exit Function
'    Pasta = PgDadosConfig.pRetorno & "\" & NumRecibo & "-pro-rec.xml"
'    If Dir(Pasta) = "" Then
'            pgnProt = "0"
'        Else
'            Set docNFe = New DOMDocument
'            docNFe.resolveExternals = True
'            docNFe.validateOnParse = True
'            docNFe.async = False
'
'            'Checa se houve algum erro ao carregar
'            If docNFe.parseError.reason <> "" Then
'                MsgBox "Erro ao ler XML : " & docNFe.parseError.reason
'                Exit Function
'            End If
'            Call docNFe.Load(Pasta)
'
'            campo = pgTagXML("<infProt Id", "</infProt>", docNFe.xml)
'            Recibo = pgTagXML("<nProt>", "</nProt>", campo)
'            If Trim(Recibo) <> "" Then
'                dhRec = pgTagXML("<dhRecbto>", "</dhRecbto>", campo)
'                dhRec = Format(Mid(dhRec, 1, InStr(dhRec, "T") - 1), "DD/MM/YYYY") & " " & Mid(dhRec, InStr(dhRec, "T") + 1, Len(dhRec))
'            End If
'    End If
'    pgnProt = Recibo & " " & dhRec
'End Function

Private Function LoadErroXML(nmArquivo As String) As String
    Dim docNFe      As DOMDocument
    Dim Pasta       As String
    Dim f           As Long
    Dim linha       As String
    Dim strTexto    As String

    If Trim(nmArquivo) = "" Then
        LoadErroXML = ""
        Exit Function
    End If
    'Pasta = PgDadosConfig.pRetorno & "\" & nmarquivo
    If Dir(nmArquivo) = "" Then
            LoadErroXML = ""
        Else
            f = FreeFile
            Open nmArquivo For Input As f   'abre o arquivo texto
            Do While Not EOF(f)
                Line Input #f, linha 'lê uma linha do arquivo texto
                strTexto = strTexto & " " & linha
            Loop
            Close #f

            LoadErroXML = Trim(RC(strTexto))
    End If
End Function
'Private Function pgNumLoteEnvioNFe(idNFe As String) As String
'    Dim docNFe      As DOMDocument
'
'    Dim Pasta       As String
'    If Trim(idNFe) = "" Then Exit Function
'    Pasta = PgDadosConfig.pRetorno & "\" & idNFe & "-num-lot.xml"
'    If Dir(Pasta) = "" Then
'            pgNumLoteEnvioNFe = "0"
'        Else
'            Set docNFe = New DOMDocument
'            docNFe.resolveExternals = True
'            docNFe.validateOnParse = True
'            docNFe.async = False
'
'            'Checa se houve algum erro ao carregar
'            If docNFe.parseError.reason <> "" Then
'                MsgBox "Erro ao ler XML : " & docNFe.parseError.reason
'                Exit Function
'            End If
'            Call docNFe.Load(Pasta)
'
'            pgNumLoteEnvioNFe = pgTagXML("<NumeroLoteGerado>", "</NumeroLoteGerado>", docNFe.xml)
'
'    End If
'End Function
'Private Function pgNumProtCancNFe(idNFe As String) As String
'    Dim docNFe      As DOMDocument
'    Dim Pasta       As String
'    Dim a, b        As String
'
'
'    If Trim(idNFe) = "" Then Exit Function
'    Pasta = PgDadosConfig.pRetorno & "\" & idNFe & "-can.xml"
'    If Dir(Pasta) = "" Then
'            pgNumProtCancNFe = "0"
'        Else
'            Set docNFe = New DOMDocument
'            docNFe.resolveExternals = True
'            docNFe.validateOnParse = True
'            docNFe.async = False
'
'            'Checa se houve algum erro ao carregar
'            If docNFe.parseError.reason <> "" Then
'                MsgBox "Erro ao ler XML : " & docNFe.parseError.reason
'                Exit Function
'            End If
'            Call docNFe.Load(Pasta)
'            a = pgTagXML("<nProt>", "</nProt>", docNFe.xml)
'            If Trim(a) <> "" Then
'                b = pgTagXML("<dhRecbto>", "</dhRecbto>", docNFe.xml)
'                b = Format(Mid(b, 1, InStr(b, "T") - 1), "DD/MM/YYYY") & " " & Mid(b, InStr(b, "T") + 1, Len(b))
'            End If
'            pgNumProtCancNFe = a & " " & b
'    End If
'End Function
'Private Function pgNumReciboNFe(NumLoteEnvio As String) As String
'    Dim docNFe      As DOMDocument
'    Dim Pasta       As String
'    Dim f           As Long
'    Dim strTexto    As String
'    Dim linha       As String
'
'    If Trim(NumLoteEnvio) = "0" Then Exit Function
'    Pasta = PgDadosConfig.pRetorno & "\" & Left(String(15, "0"), 15 - Len(NumLoteEnvio)) & NumLoteEnvio & "-rec.xml"
'    If Dir(Pasta) = "" Then
'            Pasta = PgDadosConfig.pRetorno & "\" & Left(String(15, "0"), 15 - Len(NumLoteEnvio)) & NumLoteEnvio & "-rec.err"
'            If Dir(Pasta) = "" Then
'                    pgNumReciboNFe = "0"
'                Else
'                    f = FreeFile
'                    Open Pasta For Input As f   'abre o arquivo texto
'                    Do While Not EOF(f)
'                        Line Input #f, linha 'lê uma linha do arquivo texto
'                        strTexto = strTexto & " " & linha
'                    Loop
'                    Close #f
'                    pgNumReciboNFe = strTexto
'            End If
'        Else
'            Set docNFe = New DOMDocument
'            docNFe.resolveExternals = True
'            docNFe.validateOnParse = True
'            docNFe.async = False
'
'            'Checa se houve algum erro ao carregar
'            If docNFe.parseError.reason <> "" Then
'                MsgBox "Erro ao ler XML : " & docNFe.parseError.reason
'                Exit Function
'            End If
'            Call docNFe.Load(Pasta)
'
'            pgNumReciboNFe = pgTagXML("<nRec>", "</nRec>", docNFe.xml)
'    End If
'
'End Function
Private Function LoadXML(nmArquivo As String) As String
    Dim docNFe      As DOMDocument
    'Dim Pasta       As String
    
    'If Trim(chvNFe) = "" Then Exit Function
    'Pasta = PgDadosConfig.pRetorno & "\" & nmArquivo
    If Dir(nmArquivo) = "" Then
            LoadXML = ""
        Else
            Set docNFe = New DOMDocument
            docNFe.resolveExternals = True
            docNFe.validateOnParse = True
            docNFe.async = False
            
            'Checa se houve algum erro ao carregar
            If docNFe.parseError.reason <> "" Then
                MsgBox "Erro ao ler XML : " & docNFe.parseError.reason
                LoadXML = ""
                Exit Function
            End If
            Call docNFe.Load(nmArquivo)
            LoadXML = docNFe.xml
    End If
End Function
Private Function pgTagXML(tagI As String, tagF As String, sDoc As String) As String
    Dim str As String
    If InStr(sDoc, tagI) = 0 Then
        pgTagXML = ""
        Exit Function
    End If
    str = Mid(sDoc, InStr(sDoc, tagI) + Len(tagI), Len(sDoc))
    str = Left(str, InStr(str, tagF) - 1)
    pgTagXML = str
End Function

Private Sub Form_Resize()
    On Error Resume Next
    frmFiltro.Left = 120
    frmFiltro.Width = Me.Width - 300
    msfgNotas.Top = frmFiltro.Height + tbMenu.Height + 200
    msfgNotas.Left = frmFiltro.Left
    msfgNotas.Width = Me.Width - 350
    msfgNotas.Height = Me.Height - (frmFiltro.Height + tbMenu.Height + txtObs.Height + 800)
    txtObs.Top = msfgNotas.Top + msfgNotas.Height
    txtObs.Width = Me.Width - 350
End Sub



Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualizar"
            LstNotasFiscais
        Case "Imprimir Danfe"
            ImprimirDANFE (NFe)
        Case "Excluir NFe"
            ExcluirNFe (NFe)
        Case "Cancelar NF-e"
            CancelarNFe
        Case "Imprimir Fatura(s)"
            ImprimirFatura
    End Select
End Sub
Private Sub CancelarNFe()

    If IdReg = 0 Then
        MsgBox "Selecione uma nota fiscal.", vbInformation, "Aviso"
        Exit Sub
    End If

    If Trim(msfgNotas.TextMatrix(msfgNotas.Row, 3)) = "" Then
        MsgBox "Nota Fiscal NAO REGISTRADA NA RECEITA FEDERAL. Será impossivel cancelar!", vbInformation, "Aviso"
        Exit Sub
    End If

    formFaturamentoNFeCancelada.CancelamentoNfe (IdReg)
    LstNotasFiscais
End Sub
Private Sub ExcluirNFe(chvNFe As String)
    If Trim(chvNFe) = "" Then
        MsgBox "Selecione uma NFe.", vbInformation, "Aviso"
        Exit Sub
    End If
    If Trim(msfgNotas.TextMatrix(msfgNotas.Row, 3)) <> "" Then
        MsgBox "Nota Fiscal JA REGISTRADA NA RECEITA FEDERAL. Será impossivel excluir!", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Deseja EXCLUIR nota fiscal " & NFe & "?", vbInformation + vbYesNo, "Exclusão de NF-e") = vbYes Then
        RegistroExcluir "FaturamentoNFe", "idNFe = '" & NFe & "'"
        RegistroExcluir "FaturamentoNFeCobranca", "idNFe = '" & NFe & "'"
        RegistroExcluir "FaturamentoNFeItens", "idNFe = '" & NFe & "'"
        LstNotasFiscais
    End If
    
End Sub
Private Sub msfgNotas_Click()
    With msfgNotas
        If .TextMatrix(.Row, 0) = "ID" Or .TextMatrix(.Row, 0) = "" Then Exit Sub
        IdReg = .TextMatrix(.Row, 0)
        NFe = .TextMatrix(.Row, 2)
        txtObs.Text = .TextMatrix(.Row, 8)
        
    End With
End Sub
Private Sub ImprimirFatura()
    
    
    ImprBB_Pre_Cont NFe
    'Dim sSQL    As String
    'Dim Rst     As Recordset
    'Dim nFat    As String
    
    
    'sSQL = "SELECT * FROM FaturamentoNFeCobranca WHERE IdNFe = '" & NFe & "' ORDER BY id"
    'Set Rst = RegistroBuscar(sSQL)
    'If Rst.BOF And Rst.EOF Then
    '        MsgBox "Nenhuma fatura encontrada.", vbInformation, "Aviso"
    '    Else
    '        Rst.MoveFirst
    '        nFat = Rst.Fields("cobr_nFat")
    '
    'End If
    'Rst.Close
    'sSQL = "SELECT * FROM FinanceiroContasPRCadastro WHERE NumFatura = '" & nFat & "'"
    'Set Rst = RegistroBuscar(sSQL)
    'If Rst.BOF And Rst.EOF Then
    '        MsgBox "Erro ao localizar os boletos no modulo financeiro.", vbInformation, "Aviso"
    '    Else
    '        Rst.MoveFirst
    '        ImprBB_Pre_Cont Rst.Fields("id")
    'End If
End Sub
Private Sub tmrAtualizacao_Timer()
    If chkAtualizar.Value = 1 Then
            cTempo = cTempo - 1
            If cTempo <= 0 Then
                    chkAtualizar.Caption = "Atualizar automaticamente..." & cTempo
                    cTempo = 10
                    LstNotasFiscais
                Else
                    chkAtualizar.Caption = "Atualizar automaticamente..." & cTempo
                
            End If
        Else
            chkAtualizar.Caption = "Atualizar automaticamente..."
    End If
End Sub
'Private Sub AtualizarRegistro(chvNFe As String)
'    Dim a, b, c, d, e, f    As String
'    Dim XMLErro             As String
'    Dim aReg(1000)          As Variant
'    Dim cReg                As Integer
'    cReg = 0
'
'    a = pgNumLoteEnvioNFe(chvNFe)
'    b = pgNumReciboNFe(CStr(a))
'    c = pgStatusNFe(CStr(b))
'    d = pgnProt(CStr(b))
'    e = pgNumProtCancNFe(chvNFe)
'    f = pgStatusCancNFe(chvNFe)
'    XMLErro = pgErroXML(chvNFe)
'
'    If Trim(a) <> 0 Then aReg(cReg) = Array("Lote", a, "S"): cReg = cReg + 1
'    If Trim(b) <> "" Then aReg(cReg) = Array("nRecibo", b, "S"): cReg = cReg + 1
'    If Trim(c) <> "" Then aReg(cReg) = Array("xMotivo", c, "S"): cReg = cReg + 1
'    If Trim(XMLErro) <> "" Then aReg(cReg) = Array("xMotivo", XMLErro, "S"): cReg = cReg + 1
'    If Trim(d) <> "" Then aReg(cReg) = Array("nProt", d, "S"): cReg = cReg + 1
'    If Trim(e) <> 0 Then aReg(cReg) = Array("canc_nProt", e, "S"): cReg = cReg + 1
'    If Trim(f) <> "" Then aReg(cReg) = Array("canc_Status", f, "S"): cReg = cReg + 1
'    cReg = cReg - 1
'    If cReg >= 0 Then
'        If RegistroAlterar("FaturamentoNFe", aReg, cReg, "idNFe = '" & chvNFe & "'") = False Then
'            MsgBox "Erro ao atualizar Registros", vbInformation, App.Title
'        End If
'    End If
'End Sub
Private Sub msfgnotas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    With msfgNotas
        If .Rows = 1 Then Exit Sub
        If Trim(.TextMatrix(1, 0)) = "" Then Exit Sub

        i = IIf(.MouseRow = 0, 1, .MouseRow)
        .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
    End With
End Sub
