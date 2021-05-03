VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form formSobre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre o Sistema"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   9315
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   7320
      TabIndex        =   16
      Top             =   7560
      Width           =   1875
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7275
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12832
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Licença de Uso"
      TabPicture(0)   =   "formSobre.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Text2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "formSobre.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Text1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Base de dados"
      TabPicture(2)   =   "formSobre.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Licença"
      TabPicture(3)   =   "formSobre.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.CommandButton Command2 
         Caption         =   "Atualizar Licença"
         Height          =   735
         Left            =   4080
         TabIndex        =   17
         Top             =   3420
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   6435
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "formSobre.frx":0070
         Top             =   660
         Width           =   8835
      End
      Begin VB.TextBox Text1 
         Height          =   6315
         Left            =   -74820
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "formSobre.frx":1283
         Top             =   720
         Width           =   8775
      End
      Begin VB.Frame Frame1 
         Caption         =   "Base de Dados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74820
         TabIndex        =   6
         Top             =   540
         Width           =   6435
         Begin VB.Label lblIP 
            Caption         =   "Label1"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   270
            Width           =   3495
         End
         Begin VB.Label lblPorta 
            Caption         =   "Label1"
            Height          =   195
            Left            =   3660
            TabIndex        =   10
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lblNomeBD 
            Caption         =   "Label1"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   560
            Width           =   4395
         End
         Begin VB.Label lblVersaobd 
            Caption         =   "Label1"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   850
            Width           =   4335
         End
         Begin VB.Label lblDeposito 
            Caption         =   "Label3"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   1140
            Width           =   4455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Aplicativo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         Left            =   180
         TabIndex        =   1
         Top             =   600
         Width           =   5895
         Begin VB.Label lblVersao 
            Caption         =   "Label1"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   4155
         End
         Begin VB.Label lblSerial 
            Caption         =   "Label1"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Width           =   4155
         End
         Begin VB.Label lblNomeExe 
            Caption         =   "Label3"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   900
            Width           =   2655
         End
         Begin VB.Label lblValidade 
            Caption         =   "lblValidade"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   1740
            Width           =   4215
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Licença de Uso:"
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
         Left            =   -74820
         TabIndex        =   15
         Top             =   420
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Historico:"
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
         Left            =   -74760
         TabIndex        =   13
         Top             =   480
         Width           =   2655
      End
   End
End
Attribute VB_Name = "formSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    licenca
    Form_Load
    
End Sub

'Private Sub Command1_Click()
'    On Error Resume Next
'    Dim sSQL As String
'    Dim Rst As Recordset
'    Dim Mail    As String
'
'    sSQL = "SELECT * FROM Clientes WHERE EmailNFe IS NOT NULL"
'    Set Rst = RegistroBuscar(sSQL)
'    If Rst.BOF And Rst.EOF Then
'
'        Else
'            Rst.MoveFirst
'            Do Until Rst.EOF
'                Mail = Rst.Fields("EmailNFe")
'                If Trim(Mail) <> "" Then
'                    If InStr(Mail, "@") = 0 Then
'                        Mail = Mid(Mail, 1, InStr(Mail, ".") - 1) & "@" & Mid(Mail, InStr(Mail, ".") + 1, Len(Mail))
'                        'Rst.Edit
'                        'Rst.Fields("EmailNFe") = LCase(Mail)
'                        BD.Execute "UPDATE Clientes SET EmailNFe = '" & LCase(Mail) & "', eMail = '" & LCase(Mail) & "' WHERE id=" & Rst.Fields("ID")
'                        'MsgBox Mail
'                    End If
'                End If
'                Rst.MoveNext
'            Loop
'    End If
'
'End Sub

Private Sub Form_Load()
  
    On Error Resume Next
    SSTab1.Tab = 0
    lblSerial.Caption = "Série: " & Numero_Serial
    lblIP.Caption = "IP: " & srv_IP
    lblPorta.Caption = "Porta: " & srv_Porta
    lblVersao = "Versão: " & sVersao & " rev." & cVersao
    lblNomeBD.Caption = "Database: " & nmDatabase
    lblVersaobd.Caption = "Versão Database: " & BD.Version
    lblNomeExe.Caption = "Nome executavel: " & App.EXEName
    lblDeposito.Caption = "Depósito: " & ID_Deposito & " - " & pgDescrDeposito(ID_Deposito)
    lblValidade.Caption = "VALIDADE: " & IIf(licencaValidade = "", "NÃO LICENCIADO", licencaValidade)
    
End Sub


