VERSION 5.00
Begin VB.Form formUsuTrocarSenha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A1 - Trocar Senha"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btoGravar 
      Caption         =   "&Gravar"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4875
      Begin VB.TextBox txtSenhaAntiga 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "*"
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1260
         Width           =   2775
      End
      Begin VB.TextBox txtConfirmacao 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtNovaSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   420
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Senha Antiga:"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Repitir a Nova Senha:"
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nova Senha:"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1395
      End
   End
End
Attribute VB_Name = "formUsuTrocarSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btoGravar_Click()
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim vDados(100) As Variant
    Dim cReg        As Integer
    
    If Trim(txtNovaSenha.Text) <> Trim(txtConfirmacao.Text) Then
        MsgBox "Senhas informadas divergentes. Favor verificar!", vbInformation, "Aviso"
        Exit Sub
    End If
    sSQL = "SELECT * FROM UsuGerenciador WHERE ID_Empresa = " & ID_Empresa & " AND id = " & ID_Usuario
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar os dados do Ususario! Troca de senha Cancelada.", vbInformation, "Aviso"
            Exit Sub
        Else
            Rst.MoveFirst
            If Rst.Fields("Usu_Senha") = txtSenhaAntiga.Text Then
                    cReg = 0
                    vDados(cReg) = Array("Usu_Senha", txtNovaSenha.Text, "S"): cReg = cReg + 1
                    vDados(cReg) = Array("Usu_TrocaSenha", "0", "N") ': cReg = cReg + 1
                    
                    
                    If RegistroAlterar("UsuGerenciador", vDados, cReg, "Id = " & ID_Usuario) = False Then
                            MsgBox "Erro ao Alterar."
                        Else
                            MsgBox "Senha alterada com sucesso!", vbInformation, "Aviso"
                    End If

                    
                Else
                    MsgBox "Senha antiga invalida! Ação cancelada.", vbInformation, "Aviso"
            End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    LimpaFormulario Me
End Sub
