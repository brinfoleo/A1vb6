VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form formBaseDadosAtualizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualização da Base de Dados"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9180
   Icon            =   "formBaseDadosAtualizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9180
   Begin VB.Frame Frame1 
      Height          =   5325
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   9105
      Begin VB.Frame Frame3 
         Caption         =   "Arquivo:"
         Height          =   945
         Left            =   75
         TabIndex        =   3
         Top             =   120
         Width           =   8955
         Begin VB.Label Lb_Arq 
            Caption         =   "..."
            Height          =   585
            Left            =   180
            TabIndex        =   4
            Top             =   225
            Width           =   8640
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4110
         Left            =   75
         TabIndex        =   1
         Top             =   1080
         Width           =   8955
         Begin VB.ListBox Lst_Status 
            Height          =   3765
            Left            =   135
            TabIndex        =   2
            Top             =   225
            Width           =   8685
         End
      End
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Arquivo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Executar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin MSComDlg.CommonDialog CD_Conexao 
         Left            =   6180
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5220
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formBaseDadosAtualizar.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formBaseDadosAtualizar.frx":0B3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formBaseDadosAtualizar.frx":1236
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "formBaseDadosAtualizar.frx":1550
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "formBaseDadosAtualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim caminho As String
Dim cmd(100) As String
Dim cont        As Integer




Private Sub Form_Activate()
    If chkAcesso(Me, "c") = False Then
        Unload Me
    End If
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case tbMenu.Buttons(Button.Index).ToolTipText
        Case "Buscar Arquivo"
            BuscarArquivo
        Case "Executar"
            ExecutarSQL
    End Select
End Sub


Private Sub PgDados(iCaminho As String)
    On Error GoTo TratErro
    Dim Arquivo     As String
   
    cont = 1
    Arquivo = FreeFile
    Open iCaminho For Input As Arquivo
    Do
        
        Line Input #Arquivo, cmd(cont)
        If Trim(cmd(cont)) = "" Then Exit Do
        cont = cont + 1
    Loop
    
    
    Close #Arquivo
    Exit Sub
TratErro:
    If Err.Number = 62 Then
            If cont <> 0 Then cont = cont - 1
            Exit Sub
        Else
            MsgBox Err.Description, vbInformation, Err.Number
    End If
End Sub



Private Sub BuscarArquivo()
  With CD_Conexao
        .DialogTitle = "Atualiza Base de Dados"
        .InitDir = SistemPath
        .Filter = "SQL Execute|*.sql"
        .DefaultExt = "*.sql"
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .ShowOpen
        If Trim(.filename) = "" Then Exit Sub
        If Len(.filename) >= 200 Then
                MsgBox "Nome do local para o arquivo é muito extenso. Por favor modifique!", vbInformation, "Atualiza Base de Dados"
                'Bt_Executar.Enabled = False
            Else
                caminho = Trim(.filename)
                Lb_Arq.Caption = caminho
                PgDados (caminho)
                Lst_Status.Clear
                Lst_Status.AddItem "==============================================================="
                Lst_Status.AddItem " Número de Comandos: " & Left(String(3, "0"), 3 - Len(Trim(cont))) & cont
                Lst_Status.AddItem "==============================================================="
                Lst_Status.AddItem " "
                'Bt_Executar.Enabled = True
        
        End If
        
    End With

End Sub
Private Sub ExecutarSQL()

    On Error GoTo TratSQL
    'Call RegLog("0", "MODIFICANDO BASE DE DADOS")
    Dim xCont As Integer
    xCont = 1
    Lst_Status.AddItem " "
    Lst_Status.AddItem "****************** INICIANDO DO PROCESSO ******************"
    Lst_Status.AddItem " "
    Do Until xCont > cont
        Lst_Status.AddItem Now & " - Executando comando :" & Left(String(3, "0"), 3 - Len(Trim(xCont))) & xCont
        'Lst_Status.AddItem Cmd(xCont)
        BD.Execute cmd(xCont)
        xCont = xCont + 1
    Loop
    Lst_Status.AddItem " "
    Lst_Status.AddItem "********************* FIM DO PROCESSO *************************"
    Exit Sub
TratSQL:
    'Call RegLog(Err.Number, "ERRO MOD. BD: " & Err.Description)
    Lst_Status.AddItem Now & " - (Comando: " & Left(String(3, "0"), 3 - Len(Trim(xCont))) & xCont & " - ERRO Numero: " & Err.Number & " - " & Err.Description
    Lst_Status.AddItem cmd(xCont)
    Resume Next
    
End Sub



Private Sub Form_Load()
    HDMenu Me, True
    
End Sub
