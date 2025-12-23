VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFaturamentoNFeGerenciador 
   Caption         =   "Faturamento - Controle de NF-e"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame frmEnvioEmail 
      Caption         =   "Envio de Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   8280
      TabIndex        =   10
      Top             =   4440
      Width           =   7095
      Begin VB.ListBox lstEnvMail 
         Height          =   1035
         ItemData        =   "formFaturamentoNFeGerenciador.frx":0000
         Left            =   180
         List            =   "formFaturamentoNFeGerenciador.frx":0002
         TabIndex        =   11
         Top             =   240
         Width           =   6795
      End
   End
   Begin VB.TextBox txtObs 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "formFaturamentoNFeGerenciador.frx":0004
      Top             =   4440
      Width           =   7575
   End
   Begin VB.Frame frmFiltro 
      Height          =   1395
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   15435
      Begin VB.CheckBox chkNFeEnviadasRF 
         Caption         =   "Somente NFe enviadas a RF"
         Height          =   195
         Left            =   3660
         TabIndex        =   24
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.OptionButton optListagem 
         Caption         =   "Cliente"
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
         Index           =   2
         Left            =   8580
         TabIndex        =   21
         Top             =   180
         Width           =   915
      End
      Begin VB.Frame frmCliente 
         Height          =   975
         Left            =   8460
         TabIndex        =   20
         Top             =   180
         Width           =   5835
         Begin VB.ComboBox cboCliente 
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Text            =   "Combo1"
            Top             =   300
            Width           =   5595
         End
         Begin VB.Label Label5 
            Caption         =   "Listar os ultimos 100 registros..."
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   3615
         End
      End
      Begin VB.OptionButton optListagem 
         Caption         =   "Numero de Nota"
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
         Index           =   1
         Left            =   4380
         TabIndex        =   15
         Top             =   180
         Width           =   1755
      End
      Begin VB.OptionButton optListagem 
         Caption         =   "Data de Emissão:"
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
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   180
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Frame frmNumNota 
         Height          =   735
         Left            =   4260
         TabIndex        =   13
         Top             =   180
         Width           =   3915
         Begin VB.TextBox txtnNFFin 
            Height          =   285
            Left            =   2460
            MaxLength       =   9
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   300
            Width           =   1275
         End
         Begin VB.TextBox txtnNFIni 
            Height          =   285
            Left            =   360
            MaxLength       =   9
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Até:"
            Height          =   195
            Left            =   1920
            TabIndex        =   17
            Top             =   360
            Width           =   315
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "De:"
            Height          =   195
            Left            =   60
            TabIndex        =   16
            Top             =   360
            Width           =   255
         End
      End
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
            Format          =   292225025
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
            Format          =   292225025
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
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   4154
      _Version        =   393216
      Cols            =   8
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"formFaturamentoNFeGerenciador.frx":000C
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Danfe"
            ImageIndex      =   13
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "DANFe"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Protocolo de Cancelamento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excluir NFe"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Solicitar Autorização da SEFAZ"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consultar Situação NF-e"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar NF-e"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar XML"
            ImageIndex      =   17
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Exportar XML Autorização"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Exportar XML Cancelamento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Fatura(s)"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Enviar XML por e-mail"
            ImageIndex      =   15
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Enviar XML da NFe por e-mail"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Enviar XML Cancelamento por e-mail"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog cd 
         Left            =   6900
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   11520
         TabIndex        =   12
         Top             =   60
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
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
            NumListImages   =   21
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":01A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":05F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":0910
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":11A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":23F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":2CCE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":3560
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":3DF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":5044
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":535E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":5678
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":5A6F
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":5ACD
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":6067
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":6761
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":6E5B
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":7555
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":822F
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":8F09
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":ABE3
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formFaturamentoNFeGerenciador.frx":B4BD
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
    
    Dim nProtCanc       As String
    Dim dhRecbtoCanc    As String 'Armazena a data e Hora do cancelamento
    Dim StatusCanc      As String
    
    Dim inut_dh     As String
    Dim inut_nProt  As String
    Dim inut_Status As String
    
    
    Dim aReg(1000)  As Variant
    Dim cReg        As Integer
    '*************************************************************
    '************ Checa se houve erro na convercao do arquivo
    '*************************************************************
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
    
    
    '*************************************************************
    '********* Carrega o numero do Lote que foi enviado
    '*************************************************************
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
    '*************************************************************
    '*************** Carrega NUMERO DO RECIBO de envio
    '*************************************************************
    nLote = Mid(String(15, "0"), 1, 15 - Len(nLote)) & nLote
    Arquivo = PgDadosConfig.pRetorno & "\" & nLote & "-pro-rec.xml" ' "-rec.xml"
    ConteudoXML = LoadXML(Arquivo)
    If ConteudoXML = "" Then
            'Checa se houve erro no arquivo
            Arquivo = PgDadosConfig.pRetorno & "\" & nLote & "-pro-rec.err"
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
            
            If Trim(nRecibo) = "" Then nRecibo = nLote 'Gambiarra:versao 2026 pois nao registra NFe autenticada
    End If
    '*************************************************************
    '**** Carrega o NUMERO DE PROTOCOLO Status e Motivo da NFe
    '*************************************************************
     Arquivo = PgDadosConfig.pRetorno & "\" & nRecibo & "-pro-rec.xml"
     'Arquivo = PgDadosConfig.pRetorno & "\" & nRecibo & "-pro-rec.xml"
     
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
            'tmp = pgTagXML("<infProt Id", "</infProt>", ConteudoXML)
            tmp = pgTagXML("<infProt>", "</infProt>", ConteudoXML)
            ConteudoXML = tmp
            nProt = pgTagXML("<nProt>", "</nProt>", ConteudoXML)
            cStat = pgTagXML("<cStat>", "</cStat>", ConteudoXML)
            xMotivo = pgTagXML("<xMotivo>", "</xMotivo>", ConteudoXML)
            If Trim(nProt) <> "" Then
                dhRecbto = pgTagXML("<dhRecbto>", "</dhRecbto>", ConteudoXML)
                dhRecbto = Format(Mid(dhRecbto, 1, InStr(dhRecbto, "T") - 1), "DD/MM/YYYY") & " " & Mid(dhRecbto, InStr(dhRecbto, "T") + 1, Len(dhRecbto))
            End If
    End If
    
    '*************************************************************
    '*********** Checa se a NFe Foi Cancelada
    '*************************************************************
    
    Arquivo = PgDadosConfig.pRetorno & "\" & chvNFe & "-ret-env-canc.xml"
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
            StatusCanc = "Cancelada - " & pgTagXML("<xMotivo>", "</xMotivo>", ConteudoXML)
            'nLote = pgTagXML("<NumeroLoteGerado>", "</NumeroLoteGerado>", ConteudoXML)
            If Trim(nProtCanc) <> "" Then
                dhRecbtoCanc = pgTagXML("<dhRegEvento>", "</dhRegEvento>", ConteudoXML)
                dhRecbtoCanc = Format(Mid(dhRecbtoCanc, 1, InStr(dhRecbtoCanc, "T") - 1), "DD/MM/YYYY") & " " & Mid(dhRecbtoCanc, InStr(dhRecbtoCanc, "T") + 1, Len(dhRecbtoCanc))
            End If
            'dhRecbtoCanc
    End If

    '*************************************************************
    '**************** Checa a Situacao da NFe
    '*************************************************************
    Arquivo = PgDadosConfig.pRetorno & "\" & chvNFe & "-sit.xml"
    ConteudoXML = LoadXML(Arquivo)
    If ConteudoXML = "" Then
            'Checa se houve erro no arquivo
            Arquivo = PgDadosConfig.pRetorno & "\" & chvNFe & "-sit.err"
            ConteudoXML = LoadErroXML(Arquivo)
            If ConteudoXML = "" Then
                Else
                    StatusNFe = ConteudoXML
            End If
        Else
            'Pega o Numero do Lote do Envio
            cStat = pgTagXML("<cStat>", "</cStat>", ConteudoXML)
            xMotivo = pgTagXML("<xMotivo>", "</xMotivo>", ConteudoXML)
            nProt = pgTagXML("<nProt>", "</nProt>", ConteudoXML)
            StatusNFe = xMotivo
            If Trim(nProt) <> "" Then
                dhRecbto = pgTagXML("<dhRecbto>", "</dhRecbto>", ConteudoXML)
                dhRecbto = Format(Mid(dhRecbto, 1, InStr(dhRecbto, "T") - 1), "DD/MM/YYYY") & " " & Mid(dhRecbto, InStr(dhRecbto, "T") + 1, Len(dhRecbto))
            End If
            
    End If
    'Apaga a situacao consultada
    ExcluirFile Arquivo
    RegLogDataBase 0, "", "", "Arquivo: " & Arquivo & " excluido"

    
    '*************************************************************
    '**************** Checa Inutilizacao da NFe
    '*************************************************************
    Arquivo = PgDadosConfig.pRetorno & "\" & chvNFe & "-procInutNFe.xml"
    ConteudoXML = LoadXML(Arquivo)
    If ConteudoXML = "" Then
            'Checa se houve erro no arquivo
            Arquivo = PgDadosConfig.pRetorno & "\" & chvNFe & "-procInutNFe.err"
            ConteudoXML = LoadErroXML(Arquivo)
            If ConteudoXML = "" Then
                Else
                    StatusNFe = ConteudoXML
            End If
        Else

            inut_Status = pgTagXML("<xMotivo>", "</xMotivo>", ConteudoXML)
            inut_nProt = pgTagXML("<nProt>", "</nProt>", ConteudoXML)
            inut_dh = pgTagXML("<dhRecbto>", "</dhRecbto>", ConteudoXML)
            
    End If
    
    
    If CInt(nLote) <> "0" Then aReg(cReg) = Array("Lote", nLote, "S"): cReg = cReg + 1
    If Trim(nRecibo) <> "" Then aReg(cReg) = Array("nRecibo", nRecibo, "S"): cReg = cReg + 1
    If Trim(xMotivo) <> "" Then aReg(cReg) = Array("xMotivo", xMotivo, "S"): cReg = cReg + 1
    If Trim(cStat) <> "" Then aReg(cReg) = Array("cStat", cStat, "S"): cReg = cReg + 1
    If Trim(nProt) <> "" Then aReg(cReg) = Array("nProt", nProt, "S"): cReg = cReg + 1
    If Trim(dhRecbto) <> "" Then aReg(cReg) = Array("dhProt", dhRecbto, "S"): cReg = cReg + 1
    'If Trim(dhRecbto) <> "" Then aReg(cReg) = Array("dhRecbto", dhRecbto, "S") ': cReg = cReg + 1
    
    If Trim(nProtCanc) <> "" Then aReg(cReg) = Array("canc_nProt", nProtCanc, "S"): cReg = cReg + 1
    If Trim(dhRecbtoCanc) <> "" Then aReg(cReg) = Array("canc_dhRecbto", dhRecbtoCanc, "S"): cReg = cReg + 1
    If Trim(StatusCanc) <> "" Then aReg(cReg) = Array("canc_Status", StatusCanc, "S"): cReg = cReg + 1
    
    If Trim(inut_nProt) <> "" Then aReg(cReg) = Array("inut_nProt", inut_nProt, "S"): cReg = cReg + 1
    If Trim(inut_dh) <> "" Then aReg(cReg) = Array("inut_dhRecbto", inut_dh, "S"): cReg = cReg + 1
    If Trim(inut_Status) <> "" Then aReg(cReg) = Array("inut_Status", inut_Status, "S"): cReg = cReg + 1
    
    
    '******************* Status da NFe ********************************************
    If Trim(StatusNFe) <> "" Then
            aReg(cReg) = Array("StatusNFe", StatusNFe, "S"): cReg = cReg + 1
        ElseIf nProtCanc <> "" Then
            StatusNFe = StatusCanc
        ElseIf inut_nProt <> "" Then
            StatusNFe = inut_Status
        Else
            'StatusNFe = xMotivo
            If Trim(xMotivo) <> "" Then aReg(cReg) = Array("StatusNFe", xMotivo, "S"): cReg = cReg + 1
    End If
    'aReg(cReg) = Array("StatusNFe", StatusNFe, "S"): cReg = cReg + 1
    'aReg(cReg) = Array("xMotivo", StatusNFe, "S"): cReg = cReg + 1
    '*******************************************************************************
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
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim sStatus As String
    
    
    Me.Enabled = False
    txtObs.Text = ""
    lstEnvMail.Clear

    'Exibe NFe na tela
    msfgNotas.Rows = 1
    If optListagem(0).Value = True Then
            sSQL = "SELECT * FROM FaturamentoNFe " & _
                   "WHERE ID_Empresa = " & ID_Empresa & " " & _
                   IIf(chkNFeEnviadasRF.Value = 1, "AND EnvioRF = " & chkNFeEnviadasRF.Value, "") & " " & _
                   "AND ide_dEmi >='" & Format(dtpDtInicio.Value, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & "' " & _
                   "ORDER BY ide_nNF"
        ElseIf optListagem(1).Value = True Then
            'Nao atualiza caso os dados de num. NF estajao em branco
            If Trim(txtnNFIni.Text) = "" Or Trim(txtnNFFin.Text) = "" Then
                Me.Enabled = True
                Exit Sub
            End If
            sSQL = "SELECT * FROM FaturamentoNFe " & _
                   "WHERE ID_Empresa = " & ID_Empresa & " " & _
                   "AND EnvioRF = " & chkNFeEnviadasRF.Value & " " & _
                   "AND ide_nNF >=" & txtnNFIni.Text & " AND ide_nNF <= " & txtnNFFin.Text '& "'"
        ElseIf optListagem(2).Value = True Then
            'Nao atualiza caso os dados de Nome do Cliente estajao em branco
            If Trim(cboCliente.Text) = "" Then
                Me.Enabled = True
                Exit Sub
            End If
            sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & _
                   " AND dest_xNome ='" & cboCliente.Text & "' LIMIT 100"
        Else
            sSQL = ""
            Me.Enabled = True
            Exit Sub
    End If
    Set Rst = RegistroBuscar(sSQL)
    If Rst Is Nothing Then
        Me.Enabled = True
        Exit Sub
    End If
    If Rst.BOF And Rst.EOF Then
            Rst.Close
            Me.Enabled = True
            Exit Sub
        Else
            
            Rst.MoveFirst
            With msfgNotas
                Do Until Rst.EOF
                    status (Rst.RecordCount)
                    'Le o Status da NFe *********************************************
                    If IsNull(Rst.Fields("nProt")) And Rst.Fields("EnvioRF") = 1 Then
                            'O sistema ira efetuar a leitura caso nao esteja computado o nProt
                            'na base de dados.
                            If IsNull(Rst.Fields("inut_nProt")) Then
                                IdReg = Rst.Fields("id")
                                NFe = Rst.Fields("idNFe")
                                'LoadStatusEnvio Rst.Fields("idNFe"), Rst.Fields("ide_dEmi")
                                LoadStatusEnvio NFe, Rst.Fields("ide_dEmi")
                            End If
                        Else
                            If Not IsNull(Rst.Fields("canc_xJust")) Then
                                If IsNull(Rst.Fields("canc_nProt")) Then
                                    LoadStatusEnvio Rst.Fields("idNFe"), Rst.Fields("ide_dEmi")
                                End If
                            End If
                            '17.01.18 - Efetua uma consulta caso cStat esteja vazio
                            'Exemplo: em caso de consulta de situacao
                            If IsNull(Rst.Fields("cStat")) Then
                                LoadStatusEnvio Rst.Fields("idNFe"), Rst.Fields("ide_dEmi")
                            End If
                            
                    End If
                    '*****************************************************************
                    
                    
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Rst.Fields("Id")
                    .TextMatrix(.Rows - 1, 1) = Rst.Fields("Ide_nNF") 'Num da nota
                    .TextMatrix(.Rows - 1, 2) = Rst.Fields("IdNFe") 'Chave de acesso
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("nProt")), " ", Rst.Fields("nProt"))
                    .TextMatrix(.Rows - 1, 4) = Rst.Fields("Ide_dEmi") 'Data de Emissao
                    .TextMatrix(.Rows - 1, 5) = ZE(IIf(IsNull(Rst.Fields("dest_idDest")), 0, Rst.Fields("dest_idDest")), 5) & " - " & IIf(IsNull(Rst.Fields("dest_xNome")), "", Rst.Fields("dest_xNome"))
                    If Rst.Fields("EnvioRF") = 0 Then
                            .TextMatrix(.Rows - 1, 6) = "*** NFe não enviada a Receita Federal ***"
                        Else
                            If IsNull(Rst.Fields("StatusNFe")) Then
                                    .TextMatrix(.Rows - 1, 6) = "Aguardando..."
                                Else
                                    .TextMatrix(.Rows - 1, 6) = Rst.Fields("cStat") & " - " & Rst.Fields("StatusNFe")
                                                                
                            End If
                    End If
                    .TextMatrix(.Rows - 1, 7) = UCase(PgDadosUsuario(cNull(Rst.Fields("UsuID"))).Login)
                                                                
                    
                    'Leo - 06.02.2017
                    'Checa o estatos e se for 100 envia o xml por email
                    '14.04.17 - Funcao retirada pois estava atrapalhando a
                    'atualizacao da tela, pois o usu nao enviava tds os emails
'                    If IsNull(Rst.Fields("nProt")) = False And qtdEmailsEnviados(Rst.Fields("IdNFe")) = 0 Then
'                        If MsgBox("Deseja enviar o e-mail?", vbYesNo, "E-mail automático") = vbYes Then
'                            IdReg = Rst.Fields("Id")
'                            enviarEmail 0, False
'                            IdReg = 0
'                            'MsgBox "xml enviado automaticamente"
'                        End If
'                    End If
                    
                    Rst.MoveNext
                Loop
            End With
            Rst.Close
    End If
    Me.Enabled = True
    IdReg = 0
    NFe = ""
End Sub
Private Sub LstStatisEnvioEmail()
    Dim Rst     As Recordset
    Dim sSQL    As String
    lstEnvMail.Clear
    
    If Trim(NFe) = "" Then Exit Sub
    
    sSQL = "SELECT * FROM FaturamentoNFeSendMail WHERE ID_Empresa = " & ID_Empresa & " AND IdNFe = '" & NFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                lstEnvMail.AddItem Rst.Fields("DtHr") & " - " & LCase(Rst.Fields("Status"))
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub


Private Function qtdEmailsEnviados(chvNFe As String) As Integer
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim nMail As Integer
    nMail = 0
    If Trim(chvNFe) = "" Then
        MsgBox "Erro ao localizar chave de acesso!", vbInformation, App.EXEName
        Exit Function
    End If
    sSQL = "SELECT * FROM FaturamentoNFeSendMail WHERE ID_Empresa = " & ID_Empresa & " AND IdNFe = '" & chvNFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveLast
            nMail = Rst.RecordCount
    End If
    Rst.Close
    qtdEmailsEnviados = nMail


End Function


Private Sub cboCliente_DropDown()
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim criterio    As String
    criterio = Trim(cboCliente.Text)
    
    If Trim(criterio) = "" Then
            sSQL = "SELECT DISTINCT dest_xNome FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " ORDER BY dest_xNome"
        Else
            sSQL = "SELECT DISTINCT dest_xNome FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND dest_xNome LIKE '" & criterio & "%' ORDER BY dest_xNome"
    End If
    
    Set Rst = RegistroBuscar(sSQL)
    cboCliente.Clear
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                cboCliente.AddItem cNull(Rst.Fields("dest_xNome"))
                Rst.MoveNext
            Loop
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

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    optListagem_Click (0)
    dtpDtInicio.Value = Date
    dtpDtFinal.Value = Date
    txtnNFIni.Text = ""
    txtnNFFin.Text = ""
    txtObs.Text = ""
    cboCliente.Clear
    If chkAcesso(Me, "c") = True Then
        LstNotasFiscais
    End If
    
End Sub

Private Sub status(Max As Long)
    pb.Min = 0
    pb.Max = Max
    DoEvents
    pb.Value = pb.Value + 1
    If pb.Value > 0 And pb.Value < Max Then
            pb.Visible = True
            Me.Enabled = False
        Else
            pb.Visible = False
            pb.Value = 0
            Me.Enabled = True
    End If
End Sub


Private Function LoadErroXML(nmArquivo As String) As String
    
    Dim docNFe      As DOMDocument60
    Dim Pasta       As String
    Dim F           As Long
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
            F = FreeFile
            Open nmArquivo For Input As F   'abre o arquivo texto
            Do While Not EOF(F)
                Line Input #F, linha 'lê uma linha do arquivo texto
                strTexto = strTexto & " " & linha
            Loop
            Close #F

            LoadErroXML = Trim(rc(strTexto))
    End If

End Function
Private Function LoadXML(nmArquivo As String) As String
    Dim docNFe      As DOMDocument60
    If Dir(nmArquivo) = "" Then
            LoadXML = ""
        Else
            Set docNFe = New DOMDocument60
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
    'frmFiltro.Width = Me.Width - 350 - frmEnvioEmail.Width
    frmFiltro.Width = Me.ScaleWidth - 250 'frmPeriodo.Width + frmNumNota.Width + frmCliente.Width
    
    
    
    msfgNotas.Top = frmFiltro.Height + tbMenu.Height + 200
    msfgNotas.Left = frmFiltro.Left
    msfgNotas.Width = Me.Width - 350
    msfgNotas.Height = Me.Height - (frmFiltro.Height + tbMenu.Height + txtObs.Height + 1000)
    
    txtObs.Top = msfgNotas.Top + msfgNotas.Height + 50
    txtObs.Width = (Me.Width / 2) - 250
    
    frmEnvioEmail.Top = txtObs.Top
    frmEnvioEmail.Left = txtObs.Left + txtObs.Width + 150
    frmEnvioEmail.Width = (Me.Width / 2) - 250
    frmEnvioEmail.Height = txtObs.Height
    
    lstEnvMail.Width = frmEnvioEmail.Width - 350
    lstEnvMail.Height = frmEnvioEmail.Height - 300
    
    pb.Left = Me.Width - (pb.Width + 200)
End Sub



Private Sub msfgNotas_SelChange()
    msfgNotas_Click
End Sub

Private Sub optListagem_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpDtInicio.Enabled = True
            dtpDtFinal.Enabled = True
            txtnNFIni.Enabled = False
            txtnNFFin.Enabled = False
            cboCliente.Enabled = False
        Case 1
            dtpDtInicio.Enabled = False
            dtpDtFinal.Enabled = False
            txtnNFIni.Enabled = True
            txtnNFFin.Enabled = True
            cboCliente.Enabled = False
        Case 2
            dtpDtInicio.Enabled = False
            dtpDtFinal.Enabled = False
            txtnNFIni.Enabled = False
            txtnNFFin.Enabled = False
            cboCliente.Enabled = True
        Case Else
            dtpDtInicio.Enabled = False
            dtpDtFinal.Enabled = False
            txtnNFIni.Enabled = False
            txtnNFFin.Enabled = False
            cboCliente.Enabled = False
    End Select
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Atualizar"
            LstNotasFiscais
        Case "Imprimir Danfe"
            ImpDANFE
        Case "Excluir NFe"
            ExcluirNFe (NFe)
        Case "Solicitar Autorização da SEFAZ"
            AutorizacaoSEFAZ
        Case "Cancelar NF-e"
            CancelarNFe
        Case "Imprimir Fatura(s)"
            ImprimirFatura
        Case "Enviar XML por e-mail"
            enviarEmail 0
        Case "Consultar Situação NF-e"
            ConsultarSituacao
        Case "Exportar XML"
            ExportarXML 1
    End Select
End Sub
Private Sub tbMenu_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case LCase(ButtonMenu.Text)
        Case "danfe"
            ImpDANFE
        Case "protocolo de cancelamento"
            ImprimirProtCanc Trim(NFe)
        Case "enviar xml da nfe por e-mail"
            enviarEmail 0
        Case "enviar xml cancelamento por e-mail"
            enviarEmail 1
        Case "exportar xml autorização"
            ExportarXML 1
        Case "exportar xml cancelamento"
            ExportarXML 2
    End Select

End Sub
Private Sub AutorizacaoSEFAZ()
    Dim nArqNfe     As String
    Dim cReg        As Integer
    Dim vReg(10)    As Variant
    Dim idNFe       As Integer
    
    If IdReg = 0 Then
        MsgBox "Selecione uma nota fiscal!", vbInformation, App.EXEName
        Exit Sub
    End If
    'nArqNfe = PgDadosConfig.pFileArmazenamento & "\" & NFe & ".txt"
    nArqNfe = PgDadosNotaFiscal(NFe).ide_nNF & "_" & PgDadosNotaFiscal(NFe).emit_CNPJ & "_" & Format(PgDadosNotaFiscal(NFe).ide_dEmi, "DD") & "_" & Format(PgDadosNotaFiscal(NFe).ide_dEmi, "MM") & "_" & Format(PgDadosNotaFiscal(NFe).ide_dEmi, "YYYY") & "-nfe.txt"
    'nmArq = Rst1.Fields("ide_nNF") & "_" & Rst1.Fields("emit_CNPJ") & "_" & Format(Rst1.Fields("ide_dEmi"), "DD") & "_" & Format(Rst1.Fields("ide_dEmi"), "MM") & "_" & Format(Rst1.Fields("ide_dEmi"), "YYYY") & "-nfe.txt"
    If PgDadosNotaFiscal(NFe).EnvioRF = 0 Then
        MsgBox "Essa NFe não poderá ser transmitida a Receita Federal.", vbInformation, App.EXEName
        Exit Sub
    End If
    If MsgBox("Deseja realmente retransmitir essa NFe a Receita Federal?", vbQuestion + vbYesNo, App.EXEName) = vbNo Then
        Exit Sub
    End If
    
    
    If Trim(Dir(PgDadosConfig.pFileArmazenamento & "\" & nArqNfe)) = "" Then
        MsgBox "Erro ao localizar arquivo matriz da NFe.", vbInformation, App.EXEName
        Exit Sub
    End If
    MoverPastaEnvio_UniNFe (nArqNfe)
     cReg = 0
    
    idNFe = PgDadosNotaFiscal(NFe).Id
    vReg(cReg) = Array("StatusNFe", "", "S")
    RegistroAlterar "FaturamentoNFe", vReg, cReg, "id=" & idNFe
    
    MsgBox "Solicitação de Autorização efetuada com sucesso!", vbInformation, App.EXEName
End Sub
Private Sub ExportarXML(op As Integer)
    '1 - Autorizacao
    '2 - Cancelamento
    On Error GoTo TrtErroCopyArq
    Dim Destino As String
    Dim Origem  As String
    If Trim(NFe) = "" Then
        MsgBox "Selecione uma Nota Fiscal!", vbInformation, App.EXEName
        Exit Sub
    End If
    cd.CancelError = True 'Forca um erro caso o usuario clique em cancelar/fechar
    cd.FileName = NFe
    cd.ShowSave
    'If cd.Object = "" Then Exit Sub
    Select Case op
        Case 1
            Origem = PgDadosConfig.pBackup & "\Autorizados\" & Format(msfgNotas.TextMatrix(msfgNotas.Row, 4), "YYYYMM") & "\" & NFe & "-procNFe.xml"
            Destino = cd.FileName & "-procNFe.xml"
        Case 2
            Origem = PgDadosConfig.pBackup & "\Autorizados\" & Format(msfgNotas.TextMatrix(msfgNotas.Row, 4), "YYYYMM") & "\" & NFe & "-procCancNFe.xml"
            Destino = cd.FileName & "-procCancNFe.xml"
    End Select
    
    If Dir(Origem) = "" Then
        MsgBox "XML não encontrado!", vbCritical, App.EXEName
        Exit Sub
    End If
    
    FileCopy Origem, Destino
    MsgBox "NFe exportada com sucesso!", vbInformation, "Aviso"
    Exit Sub
TrtErroCopyArq:
    'MsgBox "Erro ao exportar NFe! " & vbCrLf & Err.Description, vbInformation, "Erro n." & Err.Number
End Sub
Private Sub ConsultarSituacao()
    If Trim(NFe) = "" Then
        MsgBox "Selecione uma Nota Fiscal!", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Deseja consultar a SITUAÇÃO CADASTRAL da nota fiscal " & NFe & "?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Consultar_NFe (NFe)
        MsgBox "Consulta de Situação da Nf-e em processamento!", vbInformation, "Aviso"
    End If
    

End Sub
Private Sub ImpDANFE()
    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If
    If ImprimirDANFE(NFe) = False Then
        If MsgBox("Deseja visualizar o espelho da nota fiscal?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            ImprimirDANFE2 (NFe)
        End If
    End If
     
End Sub
Private Sub enviarEmail(tpDoc As Integer, Optional showForm = True)
    '### tpDoc
    '###      0 - XML da NFE
    '###      1 - XML do Cancelamento da NFE
    If IdReg = 0 Then
        MsgBox "Selecione uma nota fiscal.", vbInformation, App.EXEName
        Exit Sub
    End If
    If cNull(PgDadosNotaFiscal(NFe).nProt) = "" Then
        MsgBox "Nota Fiscal NAO REGISTRADA NA RECEITA FEDERAL. Será impossivel Enviar!", vbInformation, "Aviso"
        Exit Sub
    End If
    
'    If Trim(msfgNotas.TextMatrix(msfgNotas.Row, 3)) = "" Then
'        MsgBox "Nota Fiscal NAO REGISTRADA NA RECEITA FEDERAL. Será impossivel Enviar!", vbInformation, "Aviso"
'        Exit Sub
'    End If
    
    'Pega dados para envio da NFe
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim idCliente   As Integer
    Dim Anexo       As String
    Dim Assunto     As String
    Dim texto       As String
    Dim nNF         As String
    Dim dEmi        As String
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    Dim StatusEnvio As String
    'Dim msgGrvOK    As String
    'Dim msgGrvErr   As String
    
    sSQL = "SELECT * FROM faturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND idNFe = '" & NFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar a referencia da NFe.", vbInformation, "Aviso"
            Exit Sub
        Else
            Rst.MoveFirst
            idCliente = Rst.Fields("dest_idDest")
            dEmi = Rst.Fields("Ide_dEmi")
            nNF = Rst.Fields("ide_nNF")
    End If
    Rst.Close
    
    If Trim(PgDadosCliente(idCliente).emailnfe) = "" Then
        MsgBox "E-mail para envio da NF-e não cadastrado. Favor verificar.", vbInformation, "Aviso"
        'Exit Sub
    End If
    
    'nNF = msfgNotas.TextMatrix(msfgNotas.Row, 1)
    
    
    Select Case tpDoc
        Case 0 'XML  da Nota Fiscal
            Assunto = "NFe n." & nNF & " de " & dEmi & " - " & PgDadosCliente(idCliente).Nome
            
            texto = "Prezado cliente," & " " & vbCrLf & " " & vbCrLf
            texto = texto & "Você está recebendo a Nota Fiscal Eletrônica número " & nNF & ", de " & PgDadosEmpresa(ID_Empresa).Nome & ". Além disso, junto com a mercadoria seguirá o DANFE (Documento Auxiliar da Nota Fiscal Eletrônica), impresso em papel que acompanha o transporte das mesmas." & vbCrLf
            texto = texto & "Anexo à este e-mail você está recebendo também o arquivo XML da Nota Fiscal Eletrônica. Este arquivo deve ser armazenado eletronicamente por sua empresa pelo prazo de 05 (cinco) anos, conforme previsto na legislação tributária (Art. 173 do Código Tributário Nacional e § 4º da Lei 5.172 de 25/10/1966)." & vbCrLf
            texto = texto & "O DANFE em papel pode ser arquivado para apresentação ao fisco quando solicitado. Todavia, se sua empresa também for emitente de NF-e, o arquivamento eletrônico do XML de seus fornecedores é obrigatório, sendo passível de fiscalização." & vbCrLf
            texto = texto & "Para se certificar que esta NF-e é válida, queira por favor consultar sua autenticidade no site nacional do projeto NF-e (www.nfe.fazenda.gov.br), utilizando a chave de acesso contida no DANFE." & vbCrLf & " " & vbCrLf
            texto = texto & "Atenciosamente," & vbCrLf & " " & vbCrLf
            texto = texto & PgDadosEmpresa(ID_Empresa).Nome
            texto = texto & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf
            texto = texto & "Esse e-mail, bem como seu(s) anexo(s), foi gerado pelo sistema A1 - versão " & sVersao & "[rev." & cVersao & "]."

            Anexo = PgDadosConfig.pBackup & "\Autorizados\" & Format(dEmi, "YYYYMM") & "\" & NFe & "-procNFe.xml"
            
            'msgGrvOK = "Enviado XML da NFe com sucesso para " & PgDadosCliente(idCliente).emailnfe
            'msgGrvErr = "Falha no envio do XML da NFe para " & PgDadosCliente(idCliente).emailnfe
            
        Case 1 'XML do CANCELAMENTO da Nota Fiscal
            
            Assunto = "CANCELAMENTO da NFe n." & nNF & " de " & dEmi & " - " & PgDadosCliente(idCliente).Nome
            
            texto = "Prezado cliente," & " " & vbCrLf & " " & vbCrLf
            texto = texto & "Você está recebendo o CANCELAMENTO da Nota Fiscal Eletrônica número " & nNF & ", de " & PgDadosEmpresa(ID_Empresa).Nome & "." & vbCrLf
            texto = texto & "Anexo à este e-mail você está recebendo também o arquivo XML do cancelamento da Nota Fiscal Eletrônica. Este arquivo deve ser armazenado eletronicamente por sua empresa pelo prazo de 05 (cinco) anos, conforme previsto na legislação tributária (Art. 173 do Código Tributário Nacional e § 4º da Lei 5.172 de 25/10/1966)." & vbCrLf
            texto = texto & "O DANFE em papel pode ser arquivado para apresentação ao fisco quando solicitado. Todavia, se sua empresa também for emitente de NF-e, o arquivamento eletrônico do XML de seus fornecedores é obrigatório, sendo passível de fiscalização." & vbCrLf
            texto = texto & "Para se certificar que este CANCELAMENTO de NF-e é válida, queira por favor consultar sua autenticidade no site nacional do projeto NF-e (www.nfe.fazenda.gov.br), utilizando a chave de acesso." & vbCrLf & " " & vbCrLf
            texto = texto & "Atenciosamente," & vbCrLf & " " & vbCrLf
            texto = texto & PgDadosEmpresa(ID_Empresa).Nome
            texto = texto & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf
            texto = texto & "Esse e-mail, bem como seu(s) anexo(s), foi gerado pelo sistema A1 - versão " & sVersao & "[rev." & cVersao & "]."

            'Anexo = PgDadosConfig.pBackup & "\Autorizados\" & Format(dEmi, "YYYYMM") & "\" & NFe & "-procCancNFe.xml"
            Anexo = PgDadosConfig.pBackup & "\Autorizados\" & Format(dEmi, "YYYYMM") & "\" & NFe & "_110111_01-procEventoNFe.xml"
                                                                                       '_110111_01-procEventoNFe.xml
            
            'msgGrvOK = "Enviado XML de CANCELAMENTO com sucesso para " & PgDadosCliente(idCliente).emailnfe
            'msgGrvErr = "Falha no envio do XML de CANCELAMENTO para " & PgDadosCliente(idCliente).emailnfe
            
    End Select
    If Dir(Anexo) = "" Then
            MsgBox "Arquivo XML não encontrado. Envio cancelado!", vbInformation, "Aviso"
            Exit Sub
    End If
    
    '########################################################################################
    '### Enviar XML
    '########################################################################################
    'Dim mailRetorno As Integer
    'mailRetorno = formFaturamentoSendMail.ReceberDadosExternos(NFe, PgDadosCliente(idCliente).emailnfe, Anexo, Assunto, texto)
    If showForm = True Then
            formFaturamentoSendMail.ReceberDadosExternos idCliente, tpDoc, NFe, PgDadosCliente(idCliente).emailnfe, Anexo, Assunto, texto
        Else
            formFaturamentoSendMail.enviarEmail idCliente, tpDoc, NFe, PgDadosCliente(idCliente).emailnfe, Anexo, Assunto, texto
    End If
'    If mailRetorno = 1 Then
'            StatusEnvio = msgGrvOK
'        ElseIf mailRetorno = -1 Then
'            StatusEnvio = msgGrvErr
'        Else
'            StatusEnvio = ""
'    End If
'    'Grava o status do e-mail
'    If Trim(StatusEnvio) <> "" Then
'        cReg = 0
'        vReg(cReg) = Array("idNFe", NFe, "S"): cReg = cReg + 1
'        vReg(cReg) = Array("Status", StatusEnvio, "S") ': cReg = cReg + 1
'        If RegistroIncluir("FaturamentoNFeSendMail", vReg, cReg) = 0 Then
'            MsgBox "Erro ao incluir registro de envio de e-mail.", vbInformation, "Aviso"
'        End If
'    End If
    
    LstStatisEnvioEmail
    
End Sub

Private Sub CancelarNFe()

    If IdReg = 0 Then
        MsgBox "Selecione uma nota fiscal.", vbInformation, "Aviso"
        Exit Sub
    End If

    'If Trim(msfgNotas.TextMatrix(msfgNotas.Row, 3)) = "" Then
    If Trim(PgDadosNotaFiscal(NFe).nProt) = "" Then
        MsgBox "Nota Fiscal NAO REGISTRADA NA RECEITA FEDERAL. Será impossivel cancelar!", vbInformation, "Aviso"
        Exit Sub
    End If

    formFaturamentoNFeCancelada.CancelamentoNfe (IdReg)
    LstNotasFiscais
End Sub
Private Sub ExcluirNFe(chvNFe As String)
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    If chkAcesso(Me, "e") = False Then
        Exit Sub
    End If
    If Trim(chvNFe) = "" Then
        MsgBox "Selecione uma NFe.", vbInformation, "Aviso"
        Exit Sub
    End If
    'If Trim(msfgNotas.TextMatrix(msfgNotas.Row, 3)) <> "" Then
    If Trim(PgDadosNotaFiscal(chvNFe).nProt) <> "" Then
        MsgBox "Nota Fiscal JA REGISTRADA NA RECEITA FEDERAL. Será impossivel excluir!", vbInformation, "Aviso"
        Exit Sub
    End If
    '##############################################################################################################
    '### MOVIMENTAR ESTOQUE E EXCLUIR DADOS
    '##############################################################################################################
    If MsgBox("Deseja EXCLUIR nota fiscal " & NFe & "?", vbInformation + vbYesNo, "Exclusão de NF-e") = vbNo Then
        Exit Sub
    End If
    'Dar entrada no Estoque do produto
    sSQL = "SELECT * FROM EstoqueKardex WHERE NFe = '" & NFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'Rst.Close
            'Exit Sub
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                MovimentarEstoque "e", Rst.Fields("IdProduto"), Date, Rst.Fields("Documento"), Rst.Fields("Quantidade"), Rst.Fields("ValorUnitario"), Rst.Fields("ValorTotal"), "Entrada devido Nota Fiscal Excluida", _
                                  Rst.Fields("Nome"), NFe, Rst.Fields("IDNome"), Rst.Fields("DocNome")
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    
    '##############################################################
    '### Excluir Faturas
    '##############################################################
    RegistroExcluir "FaturamentoNFe", "idNFe = '" & NFe & "'"
    RegistroExcluir "FaturamentoNFeCobranca", "idNFe = '" & NFe & "'"
    RegistroExcluir "FaturamentoNFeItens", "idNFe = '" & NFe & "'"
    RegistroExcluir "financeirocontasprcadastro", "ide_NFe = '" & NFe & "'"
    'MovimentarEstoque e,
    
    
    RegLogDataBase 0, "ExcluirNFe", "0", "Excluiu NFe: " & chvNFe
    
    chvNFe = ""
    LstNotasFiscais
    
    
End Sub
Private Sub msfgNotas_Click()
    With msfgNotas
        If .TextMatrix(.Row, 0) = "ID" Or .TextMatrix(.Row, 0) = "" Then Exit Sub
        IdReg = .TextMatrix(.Row, 0)
        NFe = .TextMatrix(.Row, 2)
        txtObs.Text = .TextMatrix(.Row, 2) & vbCrLf & _
                      .TextMatrix(.Row, 6) & vbCrLf & _
                      IIf(Trim(ChkNFeTemCCe(NFe)) = 0, "", vbCrLf & _
                      "Carta de Correção: " & ChkNFeTemCCe(.TextMatrix(.Row, 2))) & vbCrLf & _
                      "Emissor: " & .TextMatrix(.Row, 7)
        chkAtualizar.Value = 0
        LstStatisEnvioEmail
    End With
End Sub
Private Sub ImprimirFatura()
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim docPrint    As String
    Dim idDoc       As Long

    If chkAcesso(Me, "i") = False Then
        Exit Sub
    End If

    sSQL = "SELECT * FROM financeiroContasPRCadastro WHERE ID_Empresa = " & ID_Empresa & " AND ide_NFe = '" & NFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar a referencia da NFe no Financeiro.", vbInformation, "Aviso"
            Exit Sub
        Else
            Rst.MoveFirst
            docPrint = pgDadosTipoDocumento(Rst.Fields("TpDocumento")).Impressao
    End If
    
    Select Case Left(docPrint, 2)
        Case "01" 'Boleto Bancario
            If PgDadosConfig.ImpBoleto = 1 Then
                    ImprBB_Pre_Cont NFe
                Else
                    If formImpressoraSelecionar.SelecionarImpressora = False Then Exit Sub
                    
                    Do Until Rst.EOF
                        idDoc = Rst.Fields("id")
                        BoletoBancario idDoc, False
                        Rst.MoveNext
                    Loop
            End If
            
        Case "02" 'Duplicata
            Do Until Rst.EOF
                idDoc = Rst.Fields("id")
                impDuplicata idDoc, False
                Rst.MoveNext
            Loop
        'Case "03" 'Falta colocar a opcao
        Case Else
            MsgBox "Erro ao localizar documento de impressão.", vbInformation, "Aviso"
    End Select
    
 
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

Private Sub msfgnotas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    With msfgNotas
        If .Rows = 1 Then Exit Sub
        If Trim(.TextMatrix(1, 0)) = "" Then Exit Sub

        i = IIf(.MouseRow = 0, 1, .MouseRow)
        .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
    End With
End Sub

Private Sub txtnNFFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtnNFIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub
