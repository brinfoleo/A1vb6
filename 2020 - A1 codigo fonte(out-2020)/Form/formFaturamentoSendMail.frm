VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formFaturamentoSendMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento - Enviar Email"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   Icon            =   "formFaturamentoSendMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTo 
      Height          =   315
      Left            =   1140
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   120
      Width           =   9135
   End
   Begin VB.TextBox txtAnexo 
      Height          =   285
      Left            =   1140
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   900
      Width           =   9135
   End
   Begin MSComctlLib.StatusBar stbConexao 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4440
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18415
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtMsg 
      Height          =   2115
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1620
      Width           =   10155
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1140
      TabIndex        =   3
      Top             =   540
      Width           =   9135
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "&Enviar"
      Height          =   495
      Left            =   6900
      TabIndex        =   0
      Top             =   3900
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Anexo:"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Mensagem:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Assunto:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Para:"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "formFaturamentoSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim StatusEnvio        As Boolean
Dim mailTO              As String
'Dim idCli               As Integer

Dim msgGrvOK            As String
Dim msgGrvErr           As String
Dim msgRetorno          As String
Dim cNFe                As String 'Chave da NFe
Dim idCli               As Integer 'Identificadir do Clinete junto ao BD
Dim itpDoc              As Integer 'identifica o tipo de documento

Public Function enviarEmail(idCliente As Integer, tpDoc As Integer, NFe As String, email As String, Anexo As String, Assunto As String, texto As String) As Integer  'As Boolean
   
    mailTO = LCase(email)
    idCli = idCliente
    cboTo.Text = mailTO
    txtSubject.Text = Assunto
    txtAnexo.Text = Anexo
    txtMsg.Text = texto
'    StatusEnvio = 0
    cNFe = NFe
    idCli = idCliente
    itpDoc = tpDoc
    
    
    cmdEnviar_Click

End Function

Public Function ReceberDadosExternos(idCliente As Integer, tpDoc As Integer, NFe As String, email As String, Anexo As String, Assunto As String, texto As String) As Integer  'As Boolean
   
    
    mailTO = LCase(email)
    idCli = idCliente
    cboTo.Text = mailTO
    txtSubject.Text = Assunto
    txtAnexo.Text = Anexo
    txtMsg.Text = texto
'    StatusEnvio = 0
    cNFe = NFe
    idCli = idCliente
    itpDoc = tpDoc
    
    
    Me.Show 1
    
'    Select Case StatusEnvio
'        Case 0
'            ReceberDadosExternos = 0
'            msgRetorno = ""
'        Case 1
'            ReceberDadosExternos = 1
'            msgRetorno = msgGrvOK
'        Case -1
'            ReceberDadosExternos = -1
'            msgRetorno = msgGrvErr
'   End Select
'
'
'    '###############################################################
'    '### 08/03/2012
'    '### Grava o status do e-mail
'    '###############################################################
'    Dim vReg(10)    As Variant
'    Dim cReg        As Integer
'    If Trim(msgRetorno) <> "" Then
'        cReg = 0
'        vReg(cReg) = Array("idNFe", NFe, "S"): cReg = cReg + 1
'        vReg(cReg) = Array("idCliente", idCliente, "N"): cReg = cReg + 1
'        vReg(cReg) = Array("Status", msgRetorno, "S") ': cReg = cReg + 1
'        'If RegistroIncluir("FaturamentoNFeSendMail", vReg, cReg) = 0 Then
'        '    MsgBox "Erro ao incluir registro de envio de e-mail.", vbInformation, App.EXEName
'        'End If
'    End If

End Function



Private Sub cboTo_Click()
    If Trim(cboTo.Text) = "" Then Exit Sub
    cboTo.Text = Trim(Mid(cboTo.Text, InStr(cboTo.Text, ":") + 1, Len(cboTo.Text)))
    
End Sub

Private Sub cboTo_DropDown()
    With cboTo
        .Clear
        .AddItem "NFe: " & LCase(PgDadosCliente(idCli).emailnfe)
        .AddItem "Comercial: " & LCase(PgDadosCliente(idCli).emailcom)
        .AddItem "Financeiro: " & LCase(PgDadosCliente(idCli).emailfin)
        .AddItem "Contato: " & LCase(PgDadosCliente(idCli).Mail)
    End With
End Sub

Private Sub cmdEnviar_Click()
    Dim vReg(10)    As Variant
    Dim cReg        As Integer
    mailTO = Trim(LCase(cboTo.Text))
    
    If Len(Trim(mailTO)) = 0 Then
        MsgBox "Informe um e-mail, por favor!", vbInformation, App.EXEName
        Exit Sub
    End If
    
    
    '###############################################################
    '### 13/04/2012
    '### Seleciona o tipo de arquivo q esta sendo enviado
    '### 0 - XML Normal
    '### 1 - XML de Cancelamento
    '###############################################################
    If itpDoc = 0 Then
            msgGrvOK = "Enviado XML da NFe com sucesso para " & LCase(mailTO)
            msgGrvErr = "Falha no envio do XML da NFe para " & LCase(mailTO)
        Else
            msgGrvOK = "Enviado XML de CANCELAMENTO com sucesso para " & LCase(mailTO)
            msgGrvErr = "Falha no envio do XML de CANCELAMENTO para " & LCase(mailTO)
    End If
    If EnvioDeEmailDLL = True Then
            msgRetorno = msgGrvOK
        Else
            msgRetorno = msgGrvErr
    End If
    
    If Trim(msgRetorno) <> "" Then
        cReg = 0
        vReg(cReg) = Array("idNFe", cNFe, "S"): cReg = cReg + 1
        vReg(cReg) = Array("idCliente", idCli, "N"): cReg = cReg + 1
        vReg(cReg) = Array("Status", msgRetorno, "S") ': cReg = cReg + 1
        If RegistroIncluir("FaturamentoNFeSendMail", vReg, cReg) = 0 Then
            MsgBox "Erro ao incluir registro de envio de e-mail.", vbInformation, App.EXEName
        End If
    End If

   Unload Me
'  'Verificar se nenhuma conexão está em andamento
'    If Winsock1.Tag = "" Then
'        If Winsock1.State <> sckClosed Then Winsock1.Close
'        Winsock1.Connect PgDadosConfig.MailSMTP, PgDadosConfig.MailSMTPPorta ' 25
'    End If
End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub txtAnexo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboTo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub


'Private Sub Winsock1_Connect()
'    Winsock1.Tag = "conectado"
'End Sub

'Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'
'Dim strData  As String
'Dim MsgTexto As String
'Dim msg      As String
'Dim status   As String
'Dim Erro     As Boolean
'
'If Trim(Winsock1.Tag) <> "" Then
'  Winsock1.GetData strData
'  status = Left(strData, 3)
'
'  'Verifica de o servidor retornou alguma msg de erro
'  Select Case status
'     Case "250", "220", "354", "221", "334", "235": Erro = False
'     Case Else:
'       Erro = True
'       Winsock1.Tag = "fechar"
'       status = Mid(strData, 4)
'  End Select
'
'  Select Case Winsock1.Tag
'    Case "conectado":
'      If PgDadosConfig.MailAutenticacao = 1 Then ' chkAuth Then
'        msg = "ehlo " & Winsock1.LocalIP & vbCrLf
'        Winsock1.Tag = "autenticar"
'      Else
'        msg = "helo " & Winsock1.LocalIP & vbCrLf
'        Winsock1.Tag = "conectou"
'      End If
'
'      Winsock1.SendData msg
'      stbConexao.Panels(1).Text = "Conectado."
'
'    Case "autenticar":
'      msg = "auth login" & vbCrLf
'      Winsock1.SendData msg
'      Winsock1.Tag = "autenticar_usuario"
'
'    Case "autenticar_usuario":
'      msg = sBase64Encode(PgDadosConfig.MailLogin) & vbCrLf
'      Winsock1.SendData msg
'      Winsock1.Tag = "autenticar_senha"
'
'    Case "autenticar_senha":
'      msg = sBase64Encode(PgDadosConfig.MailSenha) & vbCrLf
'      Winsock1.SendData msg
'      Winsock1.Tag = "conectou"
'
'    Case "conectou":
'      stbConexao.Panels(1).Text = "Enviando..."
'      Winsock1.SendData "mail from:<" & Trim(PgDadosConfig.MailEndereco) & ">" & vbCrLf
'      Winsock1.Tag = "from"
'
'    Case "from":
'      Winsock1.SendData "rcpt to:<" & Trim(mailTO) & ">" & vbCrLf
'
'      'Com copia ***********************************
'        If PgDadosConfig.MailRecCopia = 1 Then
'            Winsock1.Tag = "to"
'            Winsock1.SendData "rcpt to:<" & Trim(PgDadosConfig.MailEndereco) & ">" & vbCrLf
'        End If
'      '*****************************************************
'      Winsock1.Tag = "to"
'
'    Case "to":
'      Winsock1.SendData "data" & vbCrLf
'      Winsock1.Tag = "data"
'
'    Case "data":
'      'A sequencia "." e quebra de linha deve ser substituida por ".." e quebra de linha
'      'para evitar que o servidor entenda fim de email antes do fim do texto
'      MsgTexto = txtMsg.Text & vbCrLf
'      While InStr(MsgTexto, vbCrLf & "." & vbCrLf) <> 0
'        MsgTexto = Replace(MsgTexto, vbCrLf & "." & vbCrLf, vbCrLf & ".." & vbCrLf)
'      Wend
'
'      msg = "subject: " & txtSubject & vbCrLf
'      '********************* Mensagem em HTML *************************************************************
'      'If chkHTML = vbChecked Then
'      '  Msg = Msg & "MIME-Version: 1.0" & vbCrLf & "Content-type: text/html; charset=iso-8859-1" & vbCrLf
'      'End If
'      '****************************************************************************************************
'
'      '********************* Enviar sem anexo ************************************
'      'msg = msg & MsgTexto & vbCrLf & "." & vbCrLf
'      '***************************************************************************
'      '********************* Enviar com anexo ************************************
'      msg = msg & MsgTexto & vbCrLf '& "." & vbCrLf
'      '***************************************************************************
'      '********************* Enviar um anexo  ************************************
'      msg = msg & UUEncodeFile(Trim(txtAnexo.Text)) & vbCrLf & "." & vbCrLf
'      '***************************************************************************
'      '********************* Enviar dois anexos **********************************
'      'msg = msg & UUEncodeFile(Trim(txtAnexo.Text)) & vbCrLf
'      'msg = msg & UUEncodeFile(Trim(txtAnexo.Text)) & vbCrLf & "." & vbCrLf
'      '***************************************************************************
'      Winsock1.SendData msg
'      Winsock1.Tag = "fim"
'
'    Case "fim":
'      stbConexao.Panels(1).Text = "Desconectando..."
'      Winsock1.SendData "quit" & vbCrLf
'      Winsock1.Tag = "fechar"
'
'    Case "fechar":
'      If Not Erro Then
'        stbConexao.Panels(1).Text = "Enviado com sucesso!"
'        StatusEnvio = 1 'StatusEnvio = True
'      Else
'        stbConexao.Panels(1).Text = "Erro ao enviar email!"
'        MsgBox status, vbCritical, "Erro"
'        StatusEnvio = -1 'StatusEnvio = False
'      End If
'
'      Winsock1.Close
'      Winsock1.Tag = ""
'      Unload Me
'  End Select
'
'End If
'
'End Sub

'Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    MsgBox "Erro ao conectar" & vbNewLine & "Verifique sua conexão ou o endereço do servidor", vbCritical, "Erro"
'End Sub
Private Function EnvioDeEmailDLL() As Boolean
    'On Error GoTo trtErroEmailDLL
    'Dim StatusEnvio As Integer '0 - Cancelado,1 - enviado , -1 - erro no envio
    Dim poSendMail As New vbSendMail.clsSendMail

    HDForm Me, False
    'stbConexao.Panels(1).Text = "Enviando email para " & LCase(Trim(txtTO.Text))
    With poSendMail
        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        '.SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        '.EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = PgDadosConfig.MailSMTP '"smtp.metalcentermetais.com.br" 'txtServer.Text                  ' Required the fist time, optional thereafter
        .from = PgDadosConfig.MailEndereco '"nfe@metalcentermetais.com.br"      'txtFrom.Text                        ' Required the fist time, optional thereafter
        .FromDisplayName = PgDadosEmpresa(ID_Empresa).Nome '"Metal Center - NF-e"             'txtFromName.Text         ' Optional, saved after first use
        
        .Recipient = Trim(mailTO)                 'txtTO.Text                     ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = "Destinatario"       'txtToName.Text      ' Optional, separate multiple entries with delimiter character
        '.CcRecipient = txtCc                        ' Optional, separate multiple entries with delimiter character
        '.CcDisplayName = txtCcName                  ' Optional, separate multiple entries with delimiter character
        '.BccRecipient = txtBcc                      ' Optional, separate multiple entries with delimiter character
        '.ReplyToAddress = txtFrom.Text              ' Optional, used when different than 'From' address
        .Subject = txtSubject.Text                  ' Optional
        .Message = txtMsg.Text                      ' Optional
        .Attachment = Trim(txtAnexo.Text)           'Trim(txtAttach.Text)          ' Optional, separate multiple entries with delimiter character

        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = False                             'bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MIME_ENCODE                   'MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = HIGH_PRIORITY ' NORMAL_PRIORITY                 ' etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = False                            ' bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = True                  ' bAuthLogin             ' Optional, default = FALSE
        '.UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .Username = PgDadosConfig.MailLogin  '"nfe@metalcentermetais.com.br"  'txtUserName                     ' Optional, default = Null String
        .Password = PgDadosConfig.MailSenha '"qwe123"                        'txtPassword                     ' Optional, default = Null String, value is NOT saved
        '.POP3Host = txtPopServer
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
        .SMTPPort = PgDadosConfig.MailSMTPPorta     ' Optional, default = 25

        ' **************************************************************************
        ' OK, all of the properties are set, send the email...
        ' **************************************************************************
        ' .Connect                                  ' Optional, use when sending bulk mail
        .Send                                       ' Required
        ' .Disconnect                               ' Optional, use when sending bulk mail
        'MsgBox .SMTPHost                  ' Optional, re-populate the Host in case
                                                     ' MX look up was used to find a host
    End With
    'Screen.MousePointer = vbDefault
    'cmdSend.Enabled = True
    'MsgBox "Email enviado com sucesso...", vbInformation, App.EXEName
    HDForm Me, True
    EnvioDeEmailDLL = True
    Exit Function
trtErroEmailDLL:
    MsgBox Err.Description, vbCritical, Err.Number
    EnvioDeEmailDLL = False
End Function
    
