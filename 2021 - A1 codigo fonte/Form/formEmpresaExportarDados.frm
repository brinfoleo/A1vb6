VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formEmpresaExportarDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Dados"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   8535
   Begin VB.Frame frmEfdIcmsIpi 
      Caption         =   "EFD - ICMS/IPI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   60
      TabIndex        =   19
      Top             =   1380
      Width           =   8295
      Begin VB.CheckBox chkInventario 
         Caption         =   "&Inventário (Bloco H)"
         Height          =   255
         Left            =   3900
         TabIndex        =   26
         Top             =   1380
         Width           =   3315
      End
      Begin VB.ComboBox cbcodFinEFD 
         Height          =   315
         ItemData        =   "formEmpresaExportarDados.frx":0000
         Left            =   3900
         List            =   "formEmpresaExportarDados.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   840
         Width           =   3915
      End
      Begin MSComCtl2.DTPicker dtpDtFinEFD 
         Height          =   315
         Left            =   2160
         TabIndex        =   23
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62849025
         CurrentDate     =   41606
      End
      Begin MSComCtl2.DTPicker dtpDtIniEFD 
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62849025
         CurrentDate     =   41606
      End
      Begin VB.Label Label9 
         Caption         =   "Finalidade do Arquivo"
         Height          =   195
         Left            =   3900
         TabIndex        =   25
         Top             =   600
         Width           =   1755
      End
      Begin VB.Label Label8 
         Caption         =   "Data Final"
         Height          =   195
         Left            =   2160
         TabIndex        =   21
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Data Inicial"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   915
      End
   End
   Begin VB.Frame frmFORTES 
      Caption         =   "FORTES AC Fiscal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   60
      TabIndex        =   0
      Top             =   1380
      Width           =   8295
      Begin VB.TextBox txtComentario 
         Height          =   285
         Left            =   180
         MaxLength       =   40
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1080
         Width           =   7935
      End
      Begin MSComCtl2.DTPicker dtpDtIni 
         Height          =   315
         Left            =   600
         TabIndex        =   15
         Top             =   420
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62849025
         CurrentDate     =   40977
      End
      Begin MSComCtl2.DTPicker dtpDtFinal 
         Height          =   315
         Left            =   2760
         TabIndex        =   16
         Top             =   420
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62849025
         CurrentDate     =   40977
      End
      Begin VB.Label Label2 
         Caption         =   "De:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   450
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "Até:"
         Height          =   195
         Left            =   2400
         TabIndex        =   17
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label4 
         Caption         =   "Comentarios"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   1515
      End
   End
   Begin VB.Frame frmNFe 
      Caption         =   "Exportar XML da NF-e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   60
      TabIndex        =   5
      Top             =   1380
      Width           =   8295
      Begin MSComDlg.CommonDialog cd 
         Left            =   7740
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton btoDestinoXMLNFe 
         Caption         =   "..."
         Height          =   315
         Left            =   7080
         TabIndex        =   12
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox txtDestXML 
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1380
         Width           =   6135
      End
      Begin MSComCtl2.DTPicker dtpPeriodo 
         Height          =   315
         Left            =   1380
         TabIndex        =   8
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   62849027
         CurrentDate     =   40989
      End
      Begin VB.CheckBox chkNFeSaida 
         Caption         =   "NF-e Saida"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   1080
         Width           =   1755
      End
      Begin VB.CheckBox chkNFeEntrada 
         Caption         =   "NF-e Entrada"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   780
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Destino:"
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Periodo da NF-e:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboTpExportacao 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   900
      Width           =   8055
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   3300
         TabIndex        =   14
         Top             =   60
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
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
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":0004
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":0456
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":0770
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":1002
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":2254
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":2B2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":33C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":3C52
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":4EA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":51BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":54D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":58CF
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formEmpresaExportarDados.frx":5E69
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Selecione o tipo de Exportação"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "formEmpresaExportarDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tpExp As Integer
Private Sub Exportar()
    Select Case tpExp
        Case 1 'Arquivo Fortes Fiscal
            Fortes_GerarArquivo
        Case 2 'XML
            XML_Exportar dtpPeriodo.Value
        Case 3
            EmissaoEFD
        Case 4 'Sintegra
            Sintegra Format(dtpPeriodo.Value, "MM/YYYY")
            MsgBox "Registro gerado com sucesso!", vbInformation, App.EXEName
        Case Else
            MsgBox "Favor selecionar o tipo de Exportação", vbInformation, App.EXEName
    End Select
        
    
End Sub
Private Sub EmissaoEFD()
    Dim cod_fin As Integer
    
    If Trim(cbcodFinEFD.Text) = "" Then
            MsgBox "Favor selecionar a finalidade do arquivo.", vbInformation, App.EXEName
            Exit Sub
        Else
            cod_fin = Left(cbcodFinEFD.Text, 1)
    End If
    
    If dtpDtIniEFD.Value > dtpDtFinEFD.Value Then
        MsgBox "A data final não pode ser superior a data inicial.", vbInformation, App.EXEName
        Exit Sub
    End If
    
    MnFiscal_EFD cod_fin, dtpDtIniEFD.Value, dtpDtFinEFD.Value, IIf(chkInventario.Value = 1, True, False)
    
End Sub
Private Function Fortes_convCFOP_ES(sCFOP As String) As String
    Dim fCFOP As String
    
    sCFOP = RS(sCFOP)
    
    Select Case sCFOP
        Case "5102"
            fCFOP = "1102"
        Case "5101"
            fCFOP = "1102"
        Case "6102"
            fCFOP = "2102"
        Case "6403"
            fCFOP = "2403"
        Case "6123"
            fCFOP = "2102"
        Case "6101"
            fCFOP = "2102"
        Case "6902"
            fCFOP = "2102"
        Case "6124"
            fCFOP = "2102"
        Case "5405"
            fCFOP = "1403"
        Case "6401"
            fCFOP = "2403"
        Case "5403"
            fCFOP = "1403"
        Case Else
            'MsgBox "OPS!"
            fCFOP = sCFOP
    End Select
    Fortes_convCFOP_ES = fCFOP
End Function
Private Function Fortes_convCST_ES(sCSOSN As String) As String
    Dim fCST As String
    
    fCST = RS(sCSOSN)
    
    Select Case fCST
        Case "101"
            fCST = "00"
        Case "102"
            fCST = "40"
        Case "103"
            fCST = "40"
        Case "201"
            fCST = "10"
        Case "202"
            fCST = "30"
        Case "203"
            fCST = "30"
        Case "300"
            fCST = "51"
        Case "400"
            fCST = "41"
        Case "500"
            fCST = "60"
        Case "900"
            fCST = "90"
        Case Else
            'MsgBox "OPS!"
            fCST = fCST
    End Select
    Fortes_convCST_ES = fCST
End Function

Private Function Fortes_cvt(sTexto As String, tpDado As String, iTam As Integer, Optional cDecimais As Integer) As String
    'iTam - Tamanho
    'cDecimais - Casas decimais (maximo 2 casas quando nao declarado)
    If cDecimais = 0 Then cDecimais = 2
    
    Select Case UCase(tpDado)
        
        Case "D" 'Data
            If iTam <= 8 Then
                    Fortes_cvt = Format(sTexto, "YYMMDD")
                Else
                    Fortes_cvt = Format(sTexto, "YYYYMMDD")
            End If
        
        Case "C" 'Caracter
            If Trim(sTexto) = "" Then
                    Fortes_cvt = String(iTam, " ")
                Else
                    If Len(Trim(sTexto)) < iTam Then
                            Fortes_cvt = Trim(sTexto) & Mid(String(iTam, " "), 1, iTam - Len(Trim(sTexto)))
                        ElseIf Len(Trim(sTexto)) > iTam Then
                            Fortes_cvt = Mid(Trim(sTexto), 1, iTam)
                        Else
                            Fortes_cvt = Trim(sTexto)
                    End If
            End If
        
        Case "N" 'Numerico
            
            sTexto = RS(sTexto)
            sTexto = Replace(UCase(sTexto), "E", "")
            If Not IsNumeric(sTexto) Then
                Fortes_cvt = ""
                Exit Function
            End If
            If Len(sTexto) > iTam Then
                    Fortes_cvt = Mid(Trim(sTexto), 1, iTam)
                Else
                    Fortes_cvt = Trim(sTexto)
            End If
            
        Case "V" 'Valor
                    If Trim(sTexto) = "" Then
                            Fortes_cvt = ""
                        Else
                            Fortes_cvt = ChkVal(sTexto, 0, cDecimais)
                    End If
        
        Case "T" 'Texto
            Fortes_cvt = Trim(sTexto)
        
        Case Else 'Nenhuma das opcoes anteriores
            Fortes_cvt = sTexto
    End Select
End Function



'
Private Function Fortes_GerarArquivo() As Boolean
    'On Error GoTo trtErroFortes
    Me.Enabled = False
    '##############################################################
    '### 09/03/2012
    '### Layout de Importacao Fortes AC Fiscal versao 60
    '##############################################################
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim l           As Integer 'numero de linhas no arquivo
    Dim sTxt        As String
    Dim nmFile      As String
    
    
    
    nmFile = App.Path & "\" & RS(PgDadosEmpresa(ID_Empresa).CNPJ) & "-" & Format(Date, "YYYYMMDD") & ".fs"
    If Dir(nmFile) <> "" Then
        Kill nmFile
    End If
    l = 0
    '##############################################################
    '### Registro Tipo CAB - Cabecalho
    '##############################################################
    Dim idEmpresa   As String 'Identificador da empresa
    Dim AliquotaEspecifica As String
    Dim sComentario As String
    
    idEmpresa = ZE(CInt(PgDadosEmpresa(ID_Empresa).cCodID), 4)
    AliquotaEspecifica = "N"
    sComentario = IIf(Trim(txtComentario.Text) = "", "Arquivo gerado em " & Now(), Trim(txtComentario.Text))
    
    
    sTxt = "CAB"
    sTxt = sTxt & "|" & "60"
    sTxt = sTxt & "|" & App.ProductName '3 - sistema de Origem
    sTxt = sTxt & "|" & Fortes_cvt(Date, "D", 10)
    sTxt = sTxt & "|" & idEmpresa & "-" & Left(PgDadosEmpresa(ID_Empresa).Nome, 10)
    sTxt = sTxt & "|" & Fortes_cvt(dtpDtIni.Value, "D", 10)
    sTxt = sTxt & "|" & Fortes_cvt(dtpDtFinal.Value, "D", 10)
    sTxt = sTxt & "|" & Fortes_cvt(sComentario, "C", 40)
    sTxt = sTxt & "|" & AliquotaEspecifica
    l = l + 1
    grvFile nmFile, sTxt
    
    '##############################################################
    '### Registro Tipo PAR - Participantes dos Documentos Fiscais
    '##############################################################
    Dim cMun As String
    sSQL = "SELECT * FROM Clientes"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            'Rst.MoveLast
            
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                sTxt = "PAR"
                sTxt = sTxt & "|" & Fortes_cvt(Rst.Fields("ID"), "N", 9)
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("xNome")), "C", 60)
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("UF")), "C", 2)
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Doc")), "C", 14)
                sTxt = sTxt & "|" & Fortes_cvt(IIf(cNull(Rst.Fields("IE")) = "ISENTO", "", cNull(Rst.Fields("IE"))), "C", 14)
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("IM")), "C", 14)
                sTxt = sTxt & "|" & "N" '8 - ISS Digital
                sTxt = sTxt & "|" & "N" '9 - DIEF
                sTxt = sTxt & "|" & "N" '10 - DIC
                sTxt = sTxt & "|" & "N" '11 - DEMMS
                sTxt = sTxt & "|" & "N" '12 - Orgão Publico
                sTxt = sTxt & "|" & "N" '13 - Livro Eletronico
                sTxt = sTxt & "|" & "N" '14 - Fornecedor de Prod Primario
                sTxt = sTxt & "|" & "N" '15 - Sociedade Simples
                sTxt = sTxt & "|" & "35" '16 - Tipo Logradouro
                sTxt = sTxt & "|" & Rst.Fields("xLgr")
                  
                    
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Nro")), "N", 6) '18 - Numero
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("xCpl")), "C", 20)
                sTxt = sTxt & "|" & "01" '20 - Tipo de Bairro
                sTxt = sTxt & "|" & Rst.Fields("xBairro")
                sTxt = sTxt & "|" & Rst.Fields("CEP")
                cMun = PgDadosMunicipio(PgDadosCliente(Rst.Fields("id")).uf, PgDadosCliente(Rst.Fields("ID")).Mun).codMun
                sTxt = sTxt & "|" & Mid(cMun, 3, Len(cMun))
                sTxt = sTxt & "|" & "" '24 - DDD
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Fone")), "N", 8) '25 - Telefone
                sTxt = sTxt & "|" & Rst.Fields("Suframa")
                sTxt = sTxt & "|" & "" '27 - Substituto ISS
                sTxt = sTxt & "|" & "" '28 - Conta Remetente/Prestador
                sTxt = sTxt & "|" & "" '29 - Conta Dest/Tomador
                sTxt = sTxt & "|" & "1058"
                sTxt = sTxt & "|" & "N" '31 - Exterior
                sTxt = sTxt & "|" & IIf((Rst.Fields("IE")) = "", "C", "N") '32 - Isento de ICMS
                sTxt = sTxt & "|" & Rst.Fields("Email")
                l = l + 1
                grvFile nmFile, sTxt
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    'Fornecedores
    'cFor - Variavel criada pois nao pode haver ambiguidade nos cadastros de fornecedor e clientes para o fortes
    Dim cfor As String
    sSQL = "SELECT * FROM Fornecedores"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            'Rst.MoveLast
            
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                sTxt = "PAR"
                cfor = Fortes_cvt(Rst.Fields("ID"), "N", 7)
                cfor = "6" & Mid(String(4, "0"), 1, 4 - Len(Trim(cfor))) & cfor
                sTxt = sTxt & "|" & cfor
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("xNome")), "C", 60)
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("UF")), "C", 2)
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Doc")), "C", 14)
                sTxt = sTxt & "|" & Fortes_cvt(IIf(cNull(Rst.Fields("IE")) = "ISENTO", "", cNull(Rst.Fields("IE"))), "C", 14)
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("IM")), "C", 14)
                sTxt = sTxt & "|" & "N" '8 - ISS Digital
                sTxt = sTxt & "|" & "N" '9 - DIEF
                sTxt = sTxt & "|" & "N" '10 - DIC
                sTxt = sTxt & "|" & "N" '11 - DEMMS
                sTxt = sTxt & "|" & "N" '12 - Orgão Publico
                sTxt = sTxt & "|" & "N" '13 - Livro Eletronico
                sTxt = sTxt & "|" & "N" '14 - Fornecedor de Prod Primario
                sTxt = sTxt & "|" & "N" '15 - Sociedade Simples
                sTxt = sTxt & "|" & "35" '16 - Tipo Logradouro
                sTxt = sTxt & "|" & cNull(Rst.Fields("Lgr"))
                  
                    
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Nro")), "N", 6) '18 - Numero
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Cpl")), "C", 20)
                sTxt = sTxt & "|" & "01" '20 - Tipo de Bairro
                sTxt = sTxt & "|" & cNull(Rst.Fields("Bairro"))
                sTxt = sTxt & "|" & cNull(Rst.Fields("CEP"))
                cMun = PgDadosMunicipio(Rst.Fields("UF"), Rst.Fields("Mun")).codMun
                sTxt = sTxt & "|" & Mid(cMun, 3, Len(cMun))
                sTxt = sTxt & "|" & "" '24 - DDD
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Fone")), "N", 8) '25 - Telefone
                sTxt = sTxt & "|" & "" 'Rst.Fields("Suframa")'26 - Suframa
                sTxt = sTxt & "|" & "" '27 - Substituto ISS
                sTxt = sTxt & "|" & "" '28 - Conta Remetente/Prestador
                sTxt = sTxt & "|" & "" '29 - Conta Dest/Tomador
                sTxt = sTxt & "|" & "1058"
                sTxt = sTxt & "|" & "N" '31 - Exterior
                sTxt = sTxt & "|" & IIf((Rst.Fields("IE")) = "", "C", "N") '32 - Isento de ICMS
                sTxt = sTxt & "|" & cNull(Rst.Fields("mail"))
                l = l + 1
                grvFile nmFile, sTxt
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    
    '##############################################################
    '### Registro Tipo GRP - Grupo de Produtos
    '##############################################################
    sSQL = "SELECT * FROM EstoqueGrupos"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                sTxt = "GRP"
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("id")), "N", 8)
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Descricao")), "C", 40)
                sTxt = sTxt & "|" & Fortes_cvt("", "N", 1)
                sTxt = sTxt & "|" & Fortes_cvt("00", "N", 2) '5 - SPED Situacao
                sTxt = sTxt & "|" & Fortes_cvt("", "N", 7)
                l = l + 1
                grvFile nmFile, sTxt
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    
    '##############################################################
    '### Registro Tipo UND - Unidade de Medida
    '##############################################################
    sSQL = "SELECT * FROM EstoqueUnidadeMedida"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                sTxt = "UND"
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Sigla")), "C", 6)
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Descricao")), "C", 60)
                l = l + 1
                grvFile nmFile, sTxt
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    '##############################################################
    '### Registro Tipo PRO - Produtos
    '##############################################################
    sSQL = "SELECT * FROM EstoqueProduto " & _
           "WHERE id_Empresa = " & ID_Empresa & " AND Deposito=" & ID_Deposito & " AND Status = 'ATIVO' AND IncluirBalanco = 1"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                sTxt = "PRO"
                sTxt = sTxt & "|" & cNull(Rst.Fields("ID"))
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Descricao")), "C", 60)
                sTxt = sTxt & "|" & cNull(Rst.Fields("ID"))
                sTxt = sTxt & "|" & cNull(Rst.Fields("NCM"))
                sTxt = sTxt & "|" & cNull(Rst.Fields("Unidade"))
                sTxt = sTxt & "|" & "8" ' 07 - Unidade Medida DIEF
                sTxt = sTxt & "|" & "22" ' 08 - Unidade Medida CENFOP
                sTxt = sTxt & "|" & cNull(Rst.Fields("NCM"))  ' 09 - Classificacao Fiscal
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Grupo")), "N", 3) ' 10 - Grupo
                sTxt = sTxt & "|" & "" ' 11 - Genero
                sTxt = sTxt & "|" & cNull(Rst.Fields("CodigoBarras"))
                sTxt = sTxt & "|" & "" ' 13 - Reducao
                sTxt = sTxt & "|" & "" ' 14 - Codigo GAM57
                sTxt = sTxt & "|" & "" 'Rst.Fields("ICMSCST") 15- CST ICMS NFe Importacao e Entrada
                sTxt = sTxt & "|" & "" 'Rst.Fields("IPICST") 16- CST IPI NFe Importacao e Entrada
                sTxt = sTxt & "|" & "" 'Rst.Fields("PISCST")
                sTxt = sTxt & "|" & "" 'Rst.Fields("COFINS_CST")
                sTxt = sTxt & "|" & "" ' 19 - Codigo ANP
                sTxt = sTxt & "|" & "" ' 20 CST ICMS Simples Nacional
                sTxt = sTxt & "|" & "" ' 21 - CSOSN
                sTxt = sTxt & "|" & "" ' 22 - Produto Especifico
                sTxt = sTxt & "|" & "" ' 23 - Tipo de Medicamento
                l = l + 1
                grvFile nmFile, sTxt
    '            Rst.MoveNext
    '        Loop
    'End If
    'Rst.Close
      
    '##############################################################
    '### Registro Tipo OUM - Outras Unidades de Medida
    '##############################################################
    'sSQL = "SELECT * FROM EstoqueProduto"
    'Set Rst = RegistroBuscar(sSQL)
    'If Rst.BOF And Rst.EOF Then
    '    Else
    '        Rst.MoveFirst
    '        Do Until Rst.EOF
    '            status (Rst.RecordCount)
                sTxt = "OUM"
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("id")), "N", 9) ' 2 - Codigo do produto
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("Unidade")), "C", 9) '3 - Unidade de Medida
                sTxt = sTxt & "|" & Fortes_cvt("1", "V", 9) '4 - Unidade Equivalente Padrao
                sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("CodigoBarras")), "N", 20) '5 - Codigo de barras
                l = l + 1
                grvFile nmFile, sTxt
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    
'###################################################################################################################


    'Clientes


    '##############################################################
    '### Registro Tipo NFM - Notas Fiscais de Mercadoria
    '##############################################################
    
    Dim cancNFe As Boolean
    
    sSQL = "SELECT * " & _
           "FROM FaturamentoNFe " & _
           "WHERE ide_dEmi BETWEEN '" & Format(dtpDtIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & "' "
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            'Rst.MoveLast
            
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                cancNFe = IIf(IsNull(Rst.Fields("canc_nProt")), False, True)
                sTxt = "NFM"
                sTxt = sTxt & "|" & "0001" 'Fortes_cvt(cNull(Rst.Fields("dest_idDest")), "N", 4) ' 2 - Codigi do Estabelecimento
                sTxt = sTxt & "|" & IIf(Rst.Fields("ide_tpNF") = 0, "E", "S")
                
                'If Rst.Fields("ide_tpNF") = 0 Then
                '    MsgBox "KKSK"
                'End If
                
                sTxt = sTxt & "|" & "NFE"
                sTxt = sTxt & "|" & "S"
                sTxt = sTxt & "|" & "" ' 6 - AIDF
                sTxt = sTxt & "|" & Rst.Fields("ide_Serie")
                sTxt = sTxt & "|" & "" '8 - Sub SerieRst.Fields("ide_SSerie")
                sTxt = sTxt & "|" & Rst.Fields("ide_nNF")
                sTxt = sTxt & "|" & "" '10 - Formulario Inicial
                sTxt = sTxt & "|" & "" '11 - Formulario Final
                sTxt = sTxt & "|" & Fortes_cvt(Rst.Fields("ide_dEmi"), "D", 10)
                If cancNFe = True Then
                        sTxt = sTxt & "|" & "1" '13 - 0 Normal / 1 Cancelado
                    Else
                        If Rst.Fields("ide_tpNF") = 0 Then
                                'Entrada
                                sTxt = sTxt & "|" & ""
                                
                            Else
                                'Saida
                                sTxt = sTxt & "|" & "0"
                                
                        End If
                        
                End If
                
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", Fortes_cvt(cNull(Rst.Fields("ide_dEmi")), "D", 10)) '14 - Dt. Entr/Saida
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("dest_idDest")))
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", "N") '16 - Vinculo GNRE
                sTxt = sTxt & "|" & "" '17 - GNRE ICMS
                sTxt = sTxt & "|" & "" '18 - GNRE Mes/Ano
                sTxt = sTxt & "|" & "" '19 - GNRE Convenio
                sTxt = sTxt & "|" & "" '20 - GNRE Data Venc.
                sTxt = sTxt & "|" & "" '21 - GNRE Data Recolhimento
                sTxt = sTxt & "|" & "" '22 - GNRE Banco
                sTxt = sTxt & "|" & "" '23 - GNRE Agencia
                sTxt = sTxt & "|" & "" '24 - GNRE Agencia DV
                sTxt = sTxt & "|" & "" '25 - GNRE Autenticado
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vProd")))
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vFrete")))
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vSeg")))
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vOutro")))
                sTxt = sTxt & "|" & "" '30 - ICMS Importacao
                sTxt = sTxt & "|" & "" '31 - ICMS Importacao Diferimento
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vIPI")))
                sTxt = sTxt & "|" & "" '33 - Substituicao retido
                sTxt = sTxt & "|" & "" '34 - Servico ISS
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vDesc")))
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vNF")))
                sTxt = sTxt & "|" & "" '37 - Quantidade de Intes/Produtos
                sTxt = sTxt & "|" & "" '38 - ST Recolhes
                sTxt = sTxt & "|" & "" '39 - Antecipar Recolher
                sTxt = sTxt & "|" & "" '40 - Diferencial de Aliquota
                sTxt = sTxt & "|" & "" '41 - Valor Contabil ST
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vBCST"))) '42 - BC ICMS ST
                sTxt = sTxt & "|" & "" '43 - Valor Contabil Antecipado
                sTxt = sTxt & "|" & "" '44 - ISS Retido
                sTxt = sTxt & "|" & "" '45 - Data de Retencao do ISS
                sTxt = sTxt & "|" & "" '46 - Servico
                sTxt = sTxt & "|" & "" '47 - Data Entrada no Estado
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", IIf(Rst.Fields("transp_ModFrete") = 0, "R", "D")) '48 - Frete por conta 'Tab 11
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", "P") '49 - Fatura ' Tab 12
                sTxt = sTxt & "|" & "" '50 - Numero do EEC
                sTxt = sTxt & "|" & "" '51 - Numero do Cupom
                sTxt = sTxt & "|" & "" '52 - Receita tributavel COFINS
                sTxt = sTxt & "|" & "" '53 - Receita tributavel PIS
                sTxt = sTxt & "|" & "" '54 - Receita tributavel CSL 1
                sTxt = sTxt & "|" & "" '55 - Receita tributavel CSL 2
                sTxt = sTxt & "|" & "" '56 - Receita tributavel IRPJ 1
                sTxt = sTxt & "|" & "" '57 - Receita tributavel IRPJ 2
                sTxt = sTxt & "|" & "" '58 - Receita tributavel IRPJ 3
                sTxt = sTxt & "|" & "" '59 - Receita tributavel IRPJ 4
                sTxt = sTxt & "|" & "" '60 - COFINS Retido na fonte
                sTxt = sTxt & "|" & ""  '61 - PIS Retido na fonte
                sTxt = sTxt & "|" & ""  '62 - CSL Retido na fonte
                sTxt = sTxt & "|" & ""  '63 - IRPJ Retido na fonte
                sTxt = sTxt & "|" & "" '64 - Gera transferencia
                sTxt = sTxt & "|" & ""  '65 - Observacoes
                sTxt = sTxt & "|" & "" '66 - Aliquota ST
                
                sTxt = sTxt & "|" & Rst.Fields("idNFe") '67 - Chave Eletronica
                
                sTxt = sTxt & "|" & "" '68 - INSS Retido na Fonte
                sTxt = sTxt & "|" & "" '69 - BC COFINS / PIS nao cumulativo
                
                sTxt = sTxt & "|" & IIf(cancNFe = True, (Rst.Fields("canc_xJust")), "") '70 - Motivo de cancelamento
                
                sTxt = sTxt & "|" & cNull(Rst.Fields("ide_NatOp")) '71 - Natureza da Operacao
                sTxt = sTxt & "|" & ""  '72 - Cod. informacao complementar
                sTxt = sTxt & "|" & ""  '73 - Complemento das inf. complementares
                sTxt = sTxt & "|" & "" '74 - Hora da Saida
                sTxt = sTxt & "|" & cNull(Rst.Fields("emit_UF"))  '75 - UF de Embarque
                sTxt = sTxt & "|" & ""  '76 - Local de embarque
                grvFile nmFile, sTxt
                l = l + 1
'                Rst.MoveNext
'            Loop
'    End If
'    Rst.Close
    '##############################################################
    '### Registro Tipo PNM - Produtos(Notas Fiscais de Mercadoria)
    '##############################################################
    Dim Rst1 As Recordset
    'sSQL = "SELECT * FROM FaturamentoNFeItens"
    'Set Rst = RegistroBuscar(sSQL)
    'If Rst.BOF And Rst.EOF Then
    '    Else
            'Rst.MoveLast
    sSQL = "SELECT * " & _
           "FROM FaturamentoNFeItens " & _
           "WHERE FaturamentoNFeItens.idNFe = '" & Rst.Fields("idNFe") & "'"
    Set Rst1 = RegistroBuscar(sSQL)
    If Rst1.BOF And Rst1.EOF Then
        Else
            Rst1.MoveFirst
            Do Until Rst1.EOF
                'status (Rst1.RecordCount)
                'cancNFe = IIf(IsNull(Rst1.Fields("canc_nProt")), True, False)
                sTxt = "PNM"
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_cProd"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_CFOP"))
                sTxt = sTxt & "|" & ""  '4 - CFOP transferencia
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_ORIGEM")) '5 - CSTA
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_CST"))  '6 - CSTB
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_uCom"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_qCom"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_vProd"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("IPI_vIPI"))
                sTxt = sTxt & "|" & "3" '11 - Tipo Trib ICMS 'tab.13
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_vBC"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_pICMS"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_vBCST"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_vICMSST"))
                sTxt = sTxt & "|" & "" '16 - Tipo de recolhimento
                sTxt = sTxt & "|" & "" '17 - Tipo Substituicao
                sTxt = sTxt & "|" & "" '18 - Custo Aquisicao ST
                sTxt = sTxt & "|" & "" '19 - Perc. Agreg. Substituicao
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_vBCST"))
                sTxt = sTxt & "|" & "" '21 - Aliq ST
                sTxt = sTxt & "|" & "" '22 - Credito Origem
                sTxt = sTxt & "|" & "" '23 - Subst ja recolhido
                sTxt = sTxt & "|" & "" '24 - Custo da Aquisicao Antecip.
                sTxt = sTxt & "|" & "" '25 - Perc. Agregacao antecipada
                sTxt = sTxt & "|" & "" '26 - Aliquota Interna
                sTxt = sTxt & "|" & "" '27 - Credito de Origem
                sTxt = sTxt & "|" & "" '28 - Antec. ja Recolhido
                sTxt = sTxt & "|" & "" '29 - Base de Calc. Dif. Aliquota
                sTxt = sTxt & "|" & "" '30 - Aliquota de Origem
                sTxt = sTxt & "|" & "" '31 - Aliquota Interna
                sTxt = sTxt & "|" & "" '32 - Tipo Trib. IPI 'tab.13
                sTxt = sTxt & "|" & cNull(Rst1.Fields("IPI_vBC"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("IPI_pIPI"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("IPI_vIPI"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("IPI_CST")) '36 - CST IPI 'tab.17
                sTxt = sTxt & "|" & IIf(Rst.Fields("ide_tpNF") = 0, "70", cNull(Rst1.Fields("COFINS_CST"))) '37 - CST COFINS 'tab.18
                sTxt = sTxt & "|" & IIf(Rst.Fields("ide_tpNF") = 0, "70", cNull(Rst1.Fields("PIS_CST"))) '38 - CST PIS 'tab.18
                sTxt = sTxt & "|" & cNull(Rst1.Fields("COFINS_vBC"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("PIS_vBC"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_vFrete"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_vSeg"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_vDesc"))
                
                
                Dim vTotalSemImp As String
                
                vTotalSemImp = (Val(ChkVal(Rst1.Fields("det_vProd"), 0, cDecMoeda)) + Val(ChkVal(cNull(Rst1.Fields("det_vFrete")), 0, cDecMoeda)) + Val(ChkVal(cNull(Rst1.Fields("det_vSeg")), 0, cDecMoeda))) - Val(ChkVal(cNull(Rst1.Fields("det_vDesc")), 0, cDecMoeda))
                vTotalSemImp = ChkVal(vTotalSemImp, 0, cDecMoeda)
                sTxt = sTxt & "|" & cNull(vTotalSemImp) '44 - Valor Produto(Somatorio dos campos 9+41+42-43)
                
                sTxt = sTxt & "|" & "" '45 - Natureza da Receita COFINS
                sTxt = sTxt & "|" & "" '46 - Natureza da Receita PIS
                sTxt = sTxt & "|" & "" '47 - Indicador Especial - PRODEPE
                sTxt = sTxt & "|" & "" '48 - Codigo de Apuracao PRODEPE
                sTxt = sTxt & "|" & "" '49 - Cod. da ST do CSOSN
                sTxt = sTxt & "|" & "" '50 - CSOSN
                
                Dim cCofins As String
                If AliquotaEspecifica = "N" Then
                        cCofins = "1"
                    Else
                        If InStr(Rst1.Fields("COFINS_CST"), "03,04,06,50,51,52,53,54,55,56,60,67,70,71,72,73,74,75") Then
                                cCofins = "2"
                            Else
                                cCofins = "0"
                        End If
                End If
                sTxt = sTxt & "|" & cNull(cCofins)  '51 - Tipo Calc COFINS
                
                sTxt = sTxt & "|" & IIf(AliquotaEspecifica = "N", "", Fortes_cvt(cNull(Rst1.Fields("COFINS_pCOFINS")), "V", 7)) '52 - Aliquota COFINS(%)
                
                sTxt = sTxt & "|" & "" '53 - Aliquota COFINS(R$)
                
                sTxt = sTxt & "|" & IIf(AliquotaEspecifica = "N", "", Fortes_cvt(cNull(Rst1.Fields("COFINS_vCOFINS")), "V", 15)) '54 - Valor COFINS
                
                 Dim cPIS As String
                If AliquotaEspecifica = "N" Then
                        cPIS = "1"
                    Else
                        If InStr(Rst1.Fields("PIS_CST"), "03,04,06,50,51,52,53,54,55,56,60,67,70,71,72,73,74,75") Then
                                cPIS = "2"
                            Else
                                cPIS = "0"
                        End If
                End If
                sTxt = sTxt & "|" & cPIS '55 - Tipo Calc PIS
                sTxt = sTxt & "|" & IIf(AliquotaEspecifica = "N", "", Fortes_cvt(cNull(Rst1.Fields("PIS_pPIS")), "V", 15))  '56 - Aliquota PIS(%)
                sTxt = sTxt & "|" & "" '57 - Aliquota PIS(R$)
                sTxt = sTxt & "|" & IIf(AliquotaEspecifica = "N", "", Fortes_cvt(cNull(Rst1.Fields("PIS_vPIS")), "V", 15))   '58 - Valor PIS
                sTxt = sTxt & "|" & "" '59 - Codigo Ajuste Fiscal
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_xPed"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_nItemPed"))
                
                If cancNFe = False Then
                    grvFile nmFile, sTxt
                    l = l + 1
                End If
                Rst1.MoveNext
            Loop
    End If
    'Rst1.Close
    
        
    '##############################################################
    '### Registro Tipo INM - ICMS & IPI (Nota Fiscal de Mercadoria)
    '##############################################################
    Rst1.MoveFirst
    Do Until Rst1.EOF
        'status (Rst1.RecordCount)
        'Checa se a NFe Foi cancelada
        'cancNFe = IIf(IsNull(Rst1.Fields("canc_nProt")), True, False)
        
        sTxt = "INM"
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("det_vProd")), "V", 15) ' 2 - Valor
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("dest_UF")), "C", 2) ' 3 - Unidade Federal
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("det_CFOP")), "N", 4) ' 4 - CFOP
        sTxt = sTxt & "|" & Fortes_cvt("", "N", 4) ' 5 - CFOP Transferencia
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("ICMS_vBC")), "V", 15)
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("ICMS_pICMS")), "V", 5)
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("ICMS_vICMS")), "V", 15)
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '9 - Isenta do ICMS
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15)     '10 - Outras do ICMS
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("IPI_vBC")), "V", 15) '11 - BC do IPI
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("IPI_vIPI")), "V", 15)
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '13 - Isenta do IPI
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '14 - Outras do IPI
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '15 - ICMS ST / Somente no Simples
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '16 - IPI ST / Somente no Simples
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '17 - COFINS ST / Somente no Simples
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '18 - PIS ST / Somente no Simples
        
        If cancNFe = False Then
            grvFile nmFile, sTxt
            l = l + 1
        End If
        Rst1.MoveNext
    Loop
    Rst.MoveNext
    Loop
    
    Rst1.Close
    Rst.Close
    End If
    
    
    
    
    
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////                                                Fornecedores
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    '##############################################################
    '### Registro Tipo NFM - Notas Fiscais de Mercadoria
    '##############################################################
    
    'Dim cancNFe As Boolean
    
    sSQL = "SELECT * " & _
           "FROM FaturamentoNFeEntrada " & _
           "WHERE ide_dEmi BETWEEN '" & Format(dtpDtIni.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpDtFinal.Value, "YYYY-MM-DD") & "' "
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            'Rst.MoveLast
            
            Rst.MoveFirst
            Do Until Rst.EOF
                status (Rst.RecordCount)
                cancNFe = False 'IIf(IsNull(Rst.Fields("canc_nProt")), False, True)
                sTxt = "NFM"
                sTxt = sTxt & "|" & "0001" 'Fortes_cvt(cNull(Rst.Fields("dest_idDest")), "N", 4) ' 2 - Codigi do Estabelecimento
                sTxt = sTxt & "|" & "E" 'IIf(Rst.Fields("ide_tpNF") = 0, "E", "S")'3 - Operacao
                sTxt = sTxt & "|" & IIf(IsNull(Rst.Fields("idNFe")), "NF1", "NFE") '4 - Especie
                sTxt = sTxt & "|" & "N" '5 - Documento Proprio
                sTxt = sTxt & "|" & "" ' 6 - AIDF
                sTxt = sTxt & "|" & Rst.Fields("ide_Serie")
                sTxt = sTxt & "|" & "" '8 - Sub SerieRst.Fields("ide_SSerie")
                sTxt = sTxt & "|" & Rst.Fields("ide_nNF")
                sTxt = sTxt & "|" & "" '10 - Formulario Inicial
                sTxt = sTxt & "|" & "" '11 - Formulario Final
                sTxt = sTxt & "|" & Fortes_cvt(Rst.Fields("ide_dEmi"), "D", 10)
                sTxt = sTxt & "|" & "" 'IIf(cancNFe = True, "1", "0") '13 - 0 Normal / 1 Cancelado
                
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", Fortes_cvt(cNull(Rst.Fields("ide_dEmi")), "D", 10)) '14 - Dt. Entr/Saida
                
                cfor = Fortes_cvt(cNull(Rst.Fields("emit_id")), "N", 7)
                cfor = "6" & Mid(String(4, "0"), 1, 4 - Len(Trim(cfor))) & cfor
                
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cfor)
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", "N") '16 - Vinculo GNRE
                sTxt = sTxt & "|" & "" '17 - GNRE ICMS
                sTxt = sTxt & "|" & "" '18 - GNRE Mes/Ano
                sTxt = sTxt & "|" & "" '19 - GNRE Convenio
                sTxt = sTxt & "|" & "" '20 - GNRE Data Venc.
                sTxt = sTxt & "|" & "" '21 - GNRE Data Recolhimento
                sTxt = sTxt & "|" & "" '22 - GNRE Banco
                sTxt = sTxt & "|" & "" '23 - GNRE Agencia
                sTxt = sTxt & "|" & "" '24 - GNRE Agencia DV
                sTxt = sTxt & "|" & "" '25 - GNRE Autenticado
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vProd")))
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vFrete")))
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vSeg")))
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vOutro")))
                sTxt = sTxt & "|" & "" '30 - ICMS Importacao
                sTxt = sTxt & "|" & "" '31 - ICMS Importacao Diferimento
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vIPI")))
                sTxt = sTxt & "|" & "" '33 - Substituicao retido
                sTxt = sTxt & "|" & "" '34 - Servico ISS
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vDesc")))
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vNF")))
                sTxt = sTxt & "|" & "" '37 - Quantidade de Intes/Produtos
                sTxt = sTxt & "|" & "" '38 - ST Recolhes
                sTxt = sTxt & "|" & "" '39 - Antecipar Recolher
                sTxt = sTxt & "|" & "" '40 - Diferencial de Aliquota
                sTxt = sTxt & "|" & "" '41 - Valor Contabil ST
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", cNull(Rst.Fields("Total_vBCST"))) '42 - BC ICMS ST
                sTxt = sTxt & "|" & "" '43 - Valor Contabil Antecipado
                sTxt = sTxt & "|" & "" '44 - ISS Retido
                sTxt = sTxt & "|" & "" '45 - Data de Retencao do ISS
                sTxt = sTxt & "|" & "" '46 - Servico
                sTxt = sTxt & "|" & "" '47 - Data Entrada no Estado
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", IIf(Rst.Fields("transp_ModFrete") = 0, "R", "D")) '48 - Frete por conta 'Tab 11
                sTxt = sTxt & "|" & IIf(cancNFe = True, "", "P") '49 - Fatura ' Tab 12
                sTxt = sTxt & "|" & "" '50 - Numero do EEC
                sTxt = sTxt & "|" & "" '51 - Numero do Cupom
                sTxt = sTxt & "|" & "" '52 - Receita tributavel COFINS
                sTxt = sTxt & "|" & "" '53 - Receita tributavel PIS
                sTxt = sTxt & "|" & "" '54 - Receita tributavel CSL 1
                sTxt = sTxt & "|" & "" '55 - Receita tributavel CSL 2
                sTxt = sTxt & "|" & "" '56 - Receita tributavel IRPJ 1
                sTxt = sTxt & "|" & "" '57 - Receita tributavel IRPJ 2
                sTxt = sTxt & "|" & "" '58 - Receita tributavel IRPJ 3
                sTxt = sTxt & "|" & "" '59 - Receita tributavel IRPJ 4
                sTxt = sTxt & "|" & "" '60 - COFINS Retido na fonte
                sTxt = sTxt & "|" & ""  '61 - PIS Retido na fonte
                sTxt = sTxt & "|" & ""  '62 - CSL Retido na fonte
                sTxt = sTxt & "|" & ""  '63 - IRPJ Retido na fonte
                sTxt = sTxt & "|" & "" '64 - Gera transferencia
                sTxt = sTxt & "|" & ""  '65 - Observacoes
                sTxt = sTxt & "|" & "" '66 - Aliquota ST
                
                sTxt = sTxt & "|" & Rst.Fields("idNFe") '67 - Chave Eletronica
                
                sTxt = sTxt & "|" & "" '68 - INSS Retido na Fonte
                sTxt = sTxt & "|" & "" '69 - BC COFINS / PIS nao cumulativo
                
                sTxt = sTxt & "|" & "" 'IIf(cancNFe = True, (Rst.Fields("canc_xJust")), "") '70 - Motivo de cancelamento
                
                sTxt = sTxt & "|" & "" ' & cNull(Rst.Fields("ide_NatOp")) '71 - Natureza da Operacao
                sTxt = sTxt & "|" & ""  '72 - Cod. informacao complementar
                sTxt = sTxt & "|" & ""  '73 - Complemento das inf. complementares
                sTxt = sTxt & "|" & "" '74 - Hora da Saida
                sTxt = sTxt & "|" & cNull(Rst.Fields("emit_UF"))  '75 - UF de Embarque
                sTxt = sTxt & "|" & ""  '76 - Local de embarque
                grvFile nmFile, sTxt
                l = l + 1
'                Rst.MoveNext
'            Loop
'    End If
'    Rst.Close
    '##############################################################
    '### Registro Tipo PNM - Produtos(Notas Fiscais de Mercadoria)
    '##############################################################
    'Dim Rst1 As Recordset
    'sSQL = "SELECT * FROM FaturamentoNFeItens"
    'Set Rst = RegistroBuscar(sSQL)
    'If Rst.BOF And Rst.EOF Then
    '    Else
            'Rst.MoveLast
    sSQL = "SELECT * " & _
           "FROM FaturamentoNFeEntradaItens " & _
           "WHERE FaturamentoNFeEntradaItens.idNFe = '" & Rst.Fields("idNFe") & "'"
    Set Rst1 = RegistroBuscar(sSQL)
    If Rst1.BOF And Rst1.EOF Then
        Else
            Rst1.MoveFirst
            Do Until Rst1.EOF
                'status (Rst1.RecordCount)
                'cancNFe = IIf(IsNull(Rst1.Fields("canc_nProt")), True, False)
                sTxt = "PNM"
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_IdProduto")) 'cNull(Rst1.Fields("det_cProd"))
                sTxt = sTxt & "|" & Fortes_cvt(Fortes_convCFOP_ES(cNull(Rst1.Fields("det_CFOP"))), "N", 4)
                sTxt = sTxt & "|" & ""  '4 - CFOP transferencia
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_ORIGEM")) '5 - CSTA
                sTxt = sTxt & "|" & Fortes_convCST_ES(cNull(Rst1.Fields("ICMS_CST")))  '6 - CSTB
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_uCom"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_qCom"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_vProd"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("IPI_vIPI"))
                sTxt = sTxt & "|" & "3" '11 - Tipo Trib ICMS 'tab.13
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_vBC"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_pICMS"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_vBCST"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_vICMSST"))
                sTxt = sTxt & "|" & "" '16 - Tipo de recolhimento
                sTxt = sTxt & "|" & "" '17 - Tipo Substituicao
                sTxt = sTxt & "|" & "" '18 - Custo Aquisicao ST
                sTxt = sTxt & "|" & "" '19 - Perc. Agreg. Substituicao
                sTxt = sTxt & "|" & cNull(Rst1.Fields("ICMS_vBCST"))
                sTxt = sTxt & "|" & "" '21 - Aliq ST
                sTxt = sTxt & "|" & "" '22 - Credito Origem
                sTxt = sTxt & "|" & "" '23 - Subst ja recolhido
                sTxt = sTxt & "|" & "" '24 - Custo da Aquisicao Antecip.
                sTxt = sTxt & "|" & "" '25 - Perc. Agregacao antecipada
                sTxt = sTxt & "|" & "" '26 - Aliquota Interna
                sTxt = sTxt & "|" & "" '27 - Credito de Origem
                sTxt = sTxt & "|" & "" '28 - Antec. ja Recolhido
                sTxt = sTxt & "|" & "" '29 - Base de Calc. Dif. Aliquota
                sTxt = sTxt & "|" & "" '30 - Aliquota de Origem
                sTxt = sTxt & "|" & "" '31 - Aliquota Interna
                sTxt = sTxt & "|" & "" '32 - Tipo Trib. IPI 'tab.13
                sTxt = sTxt & "|" & cNull(Rst1.Fields("IPI_vBC"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("IPI_pIPI"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("IPI_vIPI"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("IPI_CST")) '36 - CST IPI 'tab.17
                sTxt = sTxt & "|" & "" 'cNull(Rst1.Fields("COFINS_CST")) '37 - CST COFINS 'tab.18
                sTxt = sTxt & "|" & "" 'cNull(Rst1.Fields("PIS_CST"))  '38 - CST PIS 'tab.18
                sTxt = sTxt & "|" & cNull(Rst1.Fields("COFINS_vBC"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("PIS_vBC"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_vFrete"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_vSeg"))
                sTxt = sTxt & "|" & cNull(Rst1.Fields("det_vDesc"))
                
                
                'Dim vTotalSemImp As String
                
                vTotalSemImp = (Val(ChkVal(Rst1.Fields("det_vProd"), 0, cDecMoeda)) + Val(ChkVal(cNull(Rst1.Fields("det_vFrete")), 0, cDecMoeda)) + Val(ChkVal(cNull(Rst1.Fields("det_vSeg")), 0, cDecMoeda))) - Val(ChkVal(cNull(Rst1.Fields("det_vDesc")), 0, cDecMoeda))
                vTotalSemImp = ChkVal(vTotalSemImp, 0, cDecMoeda)
                sTxt = sTxt & "|" & cNull(vTotalSemImp) '44 - Valor Produto(Somatorio dos campos 9+41+42-43)
                
                sTxt = sTxt & "|" & "" '45 - Natureza da Receita COFINS
                sTxt = sTxt & "|" & "" '46 - Natureza da Receita PIS
                sTxt = sTxt & "|" & "" '47 - Indicador Especial - PRODEPE
                sTxt = sTxt & "|" & "" '48 - Codigo de Apuracao PRODEPE
                sTxt = sTxt & "|" & "" '49 - Cod. da ST do CSOSN
                sTxt = sTxt & "|" & "" '50 - CSOSN
                
                'Dim cCofins As String
                If AliquotaEspecifica = "N" Then
                        cCofins = "1"
                    Else
                        If InStr(Rst1.Fields("COFINS_CST"), "03,04,06,50,51,52,53,54,55,56,60,67,70,71,72,73,74,75") Then
                                cCofins = "2"
                            Else
                                cCofins = "0"
                        End If
                End If
                sTxt = sTxt & "|" & cNull(cCofins)  '51 - Tipo Calc COFINS
                
                sTxt = sTxt & "|" & IIf(AliquotaEspecifica = "N", "", Fortes_cvt(cNull(Rst1.Fields("COFINS_pCOFINS")), "V", 7)) '52 - Aliquota COFINS(%)
                
                sTxt = sTxt & "|" & "" '53 - Aliquota COFINS(R$)
                
                sTxt = sTxt & "|" & IIf(AliquotaEspecifica = "N", "", Fortes_cvt(cNull(Rst1.Fields("COFINS_vCOFINS")), "V", 15)) '54 - Valor COFINS
                
                 'Dim cPIS As String
                If AliquotaEspecifica = "N" Then
                        cPIS = "1"
                    Else
                        If InStr(Rst1.Fields("PIS_CST"), "03,04,06,50,51,52,53,54,55,56,60,67,70,71,72,73,74,75") Then
                                cPIS = "2"
                            Else
                                cPIS = "0"
                        End If
                End If
                sTxt = sTxt & "|" & cPIS '55 - Tipo Calc PIS
                sTxt = sTxt & "|" & IIf(AliquotaEspecifica = "N", "", Fortes_cvt(cNull(Rst1.Fields("PIS_pPIS")), "V", 15))  '56 - Aliquota PIS(%)
                sTxt = sTxt & "|" & "" '57 - Aliquota PIS(R$)
                sTxt = sTxt & "|" & IIf(AliquotaEspecifica = "N", "", Fortes_cvt(cNull(Rst1.Fields("PIS_vPIS")), "V", 15))   '58 - Valor PIS
                sTxt = sTxt & "|" & "" '59 - Codigo Ajuste Fiscal
                sTxt = sTxt & "|" & "" 'cNull(Rst1.Fields("det_xPed"))
                sTxt = sTxt & "|" & "" 'cNull(Rst1.Fields("det_nItemPed"))
                
                If cancNFe = False Then
                    grvFile nmFile, sTxt
                    l = l + 1
                End If
                Rst1.MoveNext
            Loop
    End If
    'Rst1.Close
    
        
    '##############################################################
    '### Registro Tipo INM - ICMS & IPI (Nota Fiscal de Mercadoria)
    '##############################################################
    Rst1.MoveFirst
    Do Until Rst1.EOF
        'status (Rst1.RecordCount)
        'Checa se a NFe Foi cancelada
        'cancNFe = IIf(IsNull(Rst1.Fields("canc_nProt")), True, False)
        
        sTxt = "INM"
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("det_vProd")), "V", 15) ' 2 - Valor
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst.Fields("dest_UF")), "C", 2) ' 3 - Unidade Federal
        sTxt = sTxt & "|" & Fortes_cvt(Fortes_convCFOP_ES(cNull(Rst1.Fields("det_CFOP"))), "N", 4) ' 4 - CFOP
        sTxt = sTxt & "|" & Fortes_cvt("", "N", 4) ' 5 - CFOP Transferencia
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("ICMS_vBC")), "V", 15)
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("ICMS_pICMS")), "V", 5)
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("ICMS_vICMS")), "V", 15)
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '9 - Isenta do ICMS
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15)     '10 - Outras do ICMS
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("IPI_vBC")), "V", 15) '11 - BC do IPI
        sTxt = sTxt & "|" & Fortes_cvt(cNull(Rst1.Fields("IPI_vIPI")), "V", 15)
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '13 - Isenta do IPI
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '14 - Outras do IPI
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '15 - ICMS ST / Somente no Simples
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '16 - IPI ST / Somente no Simples
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '17 - COFINS ST / Somente no Simples
        sTxt = sTxt & "|" & Fortes_cvt("", "V", 15) '18 - PIS ST / Somente no Simples
        
        If cancNFe = False Then
            grvFile nmFile, sTxt
            l = l + 1
        End If
        Rst1.MoveNext
    Loop
    Rst.MoveNext
    Rst1.Close
    Loop
    
    'Rst1.Close
    Rst.Close
    End If
    
'###################################################################################################################
    
    
    '##############################################################
    '### Registro Tipo TRA - Trailler
    '##############################################################
    
    sTxt = "TRA"
    l = l + 1
    sTxt = sTxt & "|" & l
    grvFile nmFile, sTxt
    status (1)
    
    MsgBox "Arquivo gravado em " & nmFile & ".", vbInformation, App.EXEName
    
    Me.Enabled = True
    Exit Function
trtErroFortes:
    Me.Enabled = True
    MsgBox Err.Description, vbCritical, Err.Number
    RegLogDataBase 0, "Fortes_GerarArquivo", Err.Number, Err.Description
End Function
Private Sub status(Max As Long)
    On Error GoTo TrtStatus
    pb.min = 0
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
    Exit Sub
TrtStatus:
    pb.Visible = False
    pb.Value = 0
    Me.Enabled = True
    
End Sub



Private Sub cbcodFinEFD_DropDown()
    cbcodFinEFD.Clear
    cbcodFinEFD.AddItem "0 - Remessa do arquivo ORIGINAL"
    cbcodFinEFD.AddItem "1 - Remessa do arquivo SUBSTITUTO"

End Sub


Private Sub cboTpExportacao_Click()
    If Trim(cboTpExportacao.Text) = "" Then
            tpExp = 0
            Me.Height = 315
        Else
            tpExp = Left(Trim(cboTpExportacao.Text), 3)
            Me.Height = 3855
    End If
    Select Case tpExp
        Case 1 'Arquivo Fortes Fiscal
            verTela 1
        Case 2 'XML NFE E/S
            verTela 2
        Case 3 'EFD - ICMS/IPI
            verTela 3
        Case 4 'Sintegra
            verTela 4
        Case Else
            verTela 0
    End Select
    
    
End Sub
Private Sub verTela(opt As Integer)
    frmNFe.Visible = False
    frmFORTES.Visible = False
    frmNFe.Visible = False
    frmEfdIcmsIpi.Visible = False
    Select Case opt
        Case 0 'Else
            'frmFORTES.Visible = False
            'frmNFe.Visible = False
        Case 1 'Fortes
            frmFORTES.Top = 1380
            frmFORTES.Left = 60
            frmFORTES.Visible = True
            
            'frmNFe.Visible = False
        Case 2 'XML
            cd.Filter = "Compactado|*.zip"
            chkNFeEntrada.Visible = True
            chkNFeSaida.Visible = True
            'frmFORTES.Visible = False
            frmNFe.Caption = "Exportar XML da NFe"
            frmNFe.Top = 1380
            frmNFe.Left = 60
            frmNFe.Visible = True
        Case 3 'EFD ICMS IPI
            dtpDtIniEFD.Value = Date
            dtpDtFinEFD.Value = Date
            frmEfdIcmsIpi.Visible = True
            
            
        Case 4 'Sintegra
            cd.Filter = "Texto|*.txt"
            frmNFe.Top = 1380
            frmNFe.Left = 60
            
            frmNFe.Caption = "SINTEGRA"
            chkNFeEntrada.Visible = False
            chkNFeSaida.Visible = False
            frmNFe.Visible = True
    End Select
End Sub
Private Sub cboTpExportacao_DropDown()
    With cboTpExportacao
        .Clear
        .AddItem "001 - Registros Fiscais para aplicativo FORTES"
        .AddItem "002 - XML das Notas Fiscais de Entrada e Saida"
        .AddItem "003 - EFD - ICMS/IPI e Inventario"
        .AddItem "004 - SINTEGRA"
        
    End With
End Sub


Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Exportar"
            Exportar
    End Select
End Sub
Private Sub btoDestinoXMLNFe_Click()
   

    cd.DialogTitle = App.EXEName
    'cd.Filter = "Compactado|*.zip "
    cd.ShowSave
    txtDestXML.Text = cd.filename
    
End Sub


Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Height = cboTpExportacao.Top + cboTpExportacao.Height + 600
    LimpaFormulario Me
    
    frmFORTES.Visible = False
    dtpDtIni.Value = Date - 30
    dtpDtFinal.Value = Date
    
    
    dtpPeriodo.Value = Date
    frmNFe.Visible = False
End Sub
Private Sub XML_Exportar(Periodo As String)
    On Error GoTo trtErrorXMLexp
    Dim caminho             As String
    Dim arquivoCliente      As String
    Dim arquivoFornecedor   As String
    Dim zipDestino          As String
    
    'Valida Periodo
    If Trim(Periodo) = "" Then
            MsgBox "Selecione o periodo!", vbCritical, App.EXEName
            Exit Sub
        Else
            Periodo = Format(Periodo, "YYYYMM")
    End If
    
    'Valida Destino
    If Trim(txtDestXML.Text) = "" Then
            MsgBox "Selecione o local de destino do arquivo!", vbCritical, App.EXEName
            Exit Sub
        Else
            zipDestino = Trim(txtDestXML.Text)
    End If
    
    '#######################################################################################
    '### Saida : Clientes
    status (4)
    If chkNFeSaida.Value = 1 Then
        caminho = PgDadosConfig.pBackup & "\Autorizados\" & Periodo
        If Dir(caminho, vbDirectory) = "" Then
            MsgBox "Erro ao localizar a pasta com os dados!", vbCritical, App.EXEName
            Exit Sub
        End If
        arquivoCliente = PgDadosConfig.pFileArmazenamento & "\Saida-" & RS(PgDadosEmpresa(ID_Empresa).CNPJ) & Format(Periodo, "YYYYMM") & ".zip"
        Compacta arquivoCliente, caminho & "\*.*"
    End If
    '#######################################################################################
    
    '#######################################################################################
    '### Entrada : Fornecedores
    status (4)
    If chkNFeEntrada.Value = 1 Then
        caminho = PgDadosConfig.pXMLFornecedor & "\" & Periodo
        If Dir(caminho, vbDirectory) = "" Then
            MsgBox "Erro ao localizar a pasta com os dados"
            Exit Sub
        End If
        arquivoFornecedor = PgDadosConfig.pFileArmazenamento & "\Entrada-" & RS(PgDadosEmpresa(ID_Empresa).CNPJ) & Format(Periodo, "YYYYMM") & ".zip"
        Compacta arquivoFornecedor, caminho & "\*.*"
    End If
    '#######################################################################################
    status (4)
    Compacta zipDestino, PgDadosConfig.pFileArmazenamento & "\*-" & RS(PgDadosEmpresa(ID_Empresa).CNPJ) & Format(Periodo, "YYYYMM") & ".zip"
    
    '### Exclui os Arquivos pre compactados
    If chkNFeEntrada.Value = 1 Then
        Kill arquivoFornecedor
    End If
    If chkNFeSaida.Value = 1 Then
        Kill arquivoCliente
    End If
    
    status (4)
    If MsgBox("Arquivos gerados com sucesso!" & vbCrLf & "Deseja enviar por e-mail?", vbQuestion + vbYesNo, App.EXEName) = vbYes Then
        formSendMail.CarregarForm "gloria@argoscont.com.br", "Movimento Mensal", "Em Anexo", txtDestXML.Text
    End If
    Exit Sub
trtErrorXMLexp:
    MsgBox Err.Description, vbCritical, Err.Number
    RegLogDataBase 0, "0", "0", "XML_Exportar - Erro: " & Err.Number & " - " & Err.Description
    Resume Next
End Sub
Private Sub Sintegra(Periodo As String)

    Dim caminho             As String
    Dim zipDestino          As String
    
    'Valida Periodo
    If Trim(Periodo) = "" Then
            MsgBox "Selecione o periodo!", vbCritical, App.EXEName
            Exit Sub
        Else
            Periodo = Format(Periodo, "YYYYMM")
    End If
    
    'Valida Destino
    If Trim(txtDestXML.Text) = "" Then
            MsgBox "Selecione o local de destino do arquivo!", vbCritical, App.EXEName
            Exit Sub
        Else
            zipDestino = Trim(txtDestXML.Text)
    End If
    
    'Apaga destinho caso tenha um igual
    If Dir(zipDestino) <> "" Then
        Kill zipDestino
    End If
    
    '/Gerar sintegra
    
    gerarSintegra zipDestino, Periodo
    
End Sub
