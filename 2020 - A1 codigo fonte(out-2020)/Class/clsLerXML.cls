VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLerXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function LoadXML(nmArquivo As String) As String
    Dim docNFe      As DOMDocument
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
Public Function LoadTagXML(tagI As String, tagF As String, sDoc As String) As String
    'sDoc = string do documento recuperado com LoadXML
    'tagI = Inicio da tag
    'tagF = Final da tag
    Dim str As String
    If InStr(sDoc, tagI) = 0 Then
        LoadTagXML = ""
        Exit Function
    End If
    str = Mid(sDoc, InStr(sDoc, tagI) + Len(tagI), Len(sDoc))
    str = Left(str, InStr(str, tagF) - 1)
    LoadTagXML = str
End Function
