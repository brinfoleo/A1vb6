VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEFD_ICMSIPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim dtInicio    As String 'Reg.0000 item.04
Dim dtFinal     As String 'Reg.0000 item.05
Dim idPerfil    As String 'Reg.0000 item.14

Public Function getVersaoEFD() As String
    versaoEFD = "2.0.10"
End Function
Public Sub setDataInicial(sdate As String)
    dtInicio = sdate
End Sub
Public Sub setDataFinal(sdate As String)
    dtFinal = sdate
End Sub

Public Sub setPerfilArquivo(sPerfil As String)
    idPerfil = sPerfil
End Sub

