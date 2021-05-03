VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formSendMail 
   Caption         =   "Enviar Email"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   Icon            =   "formSendMail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   1695
      Left            =   1080
      TabIndex        =   12
      Top             =   3420
      Width           =   5415
      ExtentX         =   9551
      ExtentY         =   2990
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton btoAnexar2 
      Caption         =   "&Anexar"
      Height          =   555
      Left            =   1680
      Picture         =   "formSendMail.frx":06EA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   60
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8100
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btoAnexar 
      Height          =   315
      Left            =   9900
      Picture         =   "formSendMail.frx":0DD4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtTO 
      Height          =   345
      Left            =   1140
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   780
      Width           =   9135
   End
   Begin VB.TextBox txtAnexo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1140
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1560
      Width           =   8655
   End
   Begin MSComctlLib.StatusBar stbConexao 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6795
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18018
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtMsg 
      Height          =   2115
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2280
      Width           =   10155
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1140
      TabIndex        =   3
      Top             =   1200
      Width           =   9135
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "&Enviar"
      Height          =   555
      Left            =   120
      Picture         =   "formSendMail.frx":14BE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   1515
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   9180
      Top             =   1860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Anexo:"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   1620
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Mensagem:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Assunto:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Para:"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "formSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim StatusEnvio As Boolean
Dim mailTO      As String
Dim idCli       As Integer
Dim StatusEnvio As Integer '0 - Cancelado,1 - enviado , -1 - erro no envio
Dim bHtml       As Boolean 'Informa se o email sera enviado em HTML ou TXT
Dim poSendMail As New vbSendMail.clsSendMail



Private Sub btoAnexar_Click()
    Anexar
End Sub
Private Sub Anexar()
    cd.ShowOpen
    txtAnexo.Text = cd.filename
End Sub

Private Sub cmdEnviar_Click()
    HDForm Me, False
    stbConexao.Panels(1).Text = "Enviando email para " & LCase(Trim(txtTO.Text))
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
        .FromDisplayName = "Email NF-e"             'txtFromName.Text         ' Optional, saved after first use
        
        .Recipient = Trim(txtTO.Text)                 'txtTO.Text                     ' Required, separate multiple entries with delimiter character
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
        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MIME_ENCODE                   'MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = NORMAL_PRIORITY                 ' etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = False                            ' bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = True                  ' bAuthLogin             ' Optional, default = FALSE
        '.UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .Username = PgDadosConfig.MailEndereco '"nfe@metalcentermetais.com.br"  'txtUserName                     ' Optional, default = Null String
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
        'txtServer.Text = .SMTPHost                  ' Optional, re-populate the Host in case
                                                    ' MX look up was used to find a host    End With
        
    End With
    'Screen.MousePointer = vbDefault
    'cmdSend.Enabled = True
    MsgBox "Email enviado com sucesso...", vbInformation, App.EXEName
    HDForm Me, True
    stbConexao.Panels(1).Text = ""
End Sub

Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub
Public Function CarregarForm(Optional sPara As String, Optional sAssunto As String, Optional sMensagem As String, Optional sAnexo As String, Optional mailHtml As Boolean)
    Dim htmlFile As String
    bHtml = mailHtml
    htmlFile = "c:\prop.html"
    If Dir(htmlFile) <> "" Then
        Kill htmlFile
    End If
    grvFile htmlFile, sMensagem
    
    txtTO.Text = sPara
    txtSubject.Text = sAssunto
    txtMsg.Text = sMensagem
    
    wb.Navigate htmlFile
    
    txtAnexo.Text = sAnexo
    
     Me.Show 1
    Unload Me
End Function
Private Sub Form_Load()
    LimpaFormulario Me
    'wb.Navigate "C:\Dropbox\Dropbox\Programas\A1\codigo fonte\html\Proposta_xxxx.html"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With txtMsg
        .Top = 2280
        .Left = 120
        .Width = Me.ScaleWidth - 250
        .Height = Me.ScaleHeight - 2850
        .Visible = IIf(bHtml = True, False, True)
    End With
    With wb
        .Top = 2280
        .Left = 120
        .Width = Me.ScaleWidth - 250
        .Height = Me.ScaleHeight - 2850
        .Visible = bHtml
    End With
End Sub

Private Sub txtAnexo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub


Private Sub Winsock1_Connect()

  Winsock1.Tag = "conectado"

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim strData  As String
Dim MsgTexto As String
Dim msg      As String
Dim status   As String
Dim Erro     As Boolean

If Trim(Winsock1.Tag) <> "" Then
  Winsock1.GetData strData
  status = Left(strData, 3)
  
  'Verifica de o servidor retornou alguma msg de erro
  Select Case status
     Case "250", "220", "354", "221", "334", "235": Erro = False
     Case Else:
       Erro = True
       Winsock1.Tag = "fechar"
       status = Mid(strData, 4)
  End Select
  
  Select Case Winsock1.Tag
    Case "conectado":
      If PgDadosConfig.MailAutenticacao = 1 Then ' chkAuth Then
        msg = "ehlo " & Winsock1.LocalIP & vbCrLf
        Winsock1.Tag = "autenticar"
      Else
        msg = "helo " & Winsock1.LocalIP & vbCrLf
        Winsock1.Tag = "conectou"
      End If
      
      Winsock1.SendData msg
      stbConexao.Panels(1).Text = "Conectado."
    
    Case "autenticar":
      msg = "auth login" & vbCrLf
      Winsock1.SendData msg
      Winsock1.Tag = "autenticar_usuario"
    
    Case "autenticar_usuario":
      msg = sBase64Encode(PgDadosConfig.MailLogin) & vbCrLf
      Winsock1.SendData msg
      Winsock1.Tag = "autenticar_senha"
    
    Case "autenticar_senha":
      msg = sBase64Encode(PgDadosConfig.MailSenha) & vbCrLf
      Winsock1.SendData msg
      Winsock1.Tag = "conectou"

    Case "conectou":
      stbConexao.Panels(1).Text = "Enviando..."
      Winsock1.SendData "mail from:<" & Trim(PgDadosConfig.MailEndereco) & ">" & vbCrLf
      Winsock1.Tag = "from"
    
    Case "from":
      Winsock1.SendData "rcpt to:<" & Trim(mailTO) & ">" & vbCrLf
      
      'Com copia ***********************************
      '  If PgDadosConfig.MailRecCopia = 1 Then
      '      Winsock1.Tag = "to"
      '      Winsock1.SendData "rcpt to:<" & Trim(PgDadosConfig.MailEndereco) & ">" & vbCrLf
      '  End If
      '*****************************************************
      Winsock1.Tag = "to"
    
    Case "to":
      Winsock1.SendData "data" & vbCrLf
      Winsock1.Tag = "data"
      
    Case "data":
      'A sequencia "." e quebra de linha deve ser substituida por ".." e quebra de linha
      'para evitar que o servidor entenda fim de email antes do fim do texto
      MsgTexto = txtMsg.Text & vbCrLf
      While InStr(MsgTexto, vbCrLf & "." & vbCrLf) <> 0
        MsgTexto = Replace(MsgTexto, vbCrLf & "." & vbCrLf, vbCrLf & ".." & vbCrLf)
      Wend
      
      msg = "subject: " & txtSubject & vbCrLf
      '********************* Mensagem em HTML *************************************************************
      'If chkHTML = vbChecked Then
      '  Msg = Msg & "MIME-Version: 1.0" & vbCrLf & "Content-type: text/html; charset=iso-8859-1" & vbCrLf
      'End If
      '****************************************************************************************************
      If Trim(txtAnexo.Text) = "" Then
                '********************* Enviar sem anexo ************************************
                msg = msg & MsgTexto & vbCrLf & "." & vbCrLf
                '***************************************************************************
                
            Else
                '********************* Enviar com anexo ************************************
                msg = msg & MsgTexto & vbCrLf '& "." & vbCrLf
                '***************************************************************************
                '********************* Enviar um anexo  ************************************
                msg = msg & UUEncodeFile(Trim(txtAnexo.Text)) & vbCrLf & "." & vbCrLf
                '***************************************************************************
                '********************* Enviar dois anexos **********************************
                'msg = msg & UUEncodeFile(Trim(txtAnexo.Text)) & vbCrLf
                'msg = msg & UUEncodeFile(Trim(txtAnexo.Text)) & vbCrLf & "." & vbCrLf
                '***************************************************************************
        End If
      Winsock1.SendData msg
      Winsock1.Tag = "fim"
      
    Case "fim":
      stbConexao.Panels(1).Text = "Desconectando..."
      Winsock1.SendData "quit" & vbCrLf
      Winsock1.Tag = "fechar"
      
    Case "fechar":
      If Not Erro Then
        stbConexao.Panels(1).Text = "Enviado com sucesso!"
        StatusEnvio = 1 'StatusEnvio = True
      Else
        stbConexao.Panels(1).Text = "Erro ao enviar email!"
        MsgBox status, vbCritical, "Erro"
        StatusEnvio = -1 'StatusEnvio = False
      End If
      
      Winsock1.Close
      Winsock1.Tag = ""
      Unload Me
  End Select
  
End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  MsgBox "Erro ao conectar" & vbNewLine & "Verifique sua conexão ou o endereço do servidor", vbCritical, "Erro"
End Sub
