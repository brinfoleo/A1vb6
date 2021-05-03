Attribute VB_Name = "ModuloUsuarios"
Option Explicit

Public Sub UsuariosConectados()
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim vReg(1) As Variant
    Dim sTexto  As String
    

    '#
    '# Leonardo Aquino
    '# 27/09/2012
    '#
    '# Mudanca de codigo para avalir todos os usuarios conectado ao
    '# invez de um por vez.
    '# Diminuição tempo de resposta para de 6 para 1 minuto
    '#
    '#
    
    Dim l As Integer
    sSQL = "SELECT * FROM ConexaoGerenciador WHERE ID_Empresa = " & ID_Empresa
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar Usuario na tabela. Feche o aplicativo e abra novamente!", vbInformation, App.EXEName
        Else
            Rst.MoveFirst
            l = 1
            Do Until Rst.EOF
                sTexto = "CHECAR - " & Time
                'msfgConec.TextMatrix(l, 7) = sTexto
                vReg(0) = Array("Status", sTexto, "S")
                RegistroAlterar "ConexaoGerenciador", vReg, 0, "id=" & Rst.fields("id")
                Rst.MoveNext
                'l = IIf(msfgConec.Rows < l, l, l + 1)
                
            Loop
    End If
    Rst.Close

End Sub

Public Sub apagarUsuariosOffline()
    On Error Resume Next
    Dim Rst     As Recordset
    Dim sSQL    As String
                    
    Dim Hr      As String
    Dim min     As Integer
    Dim status As String

    'UsuariosConectados

    'Checa se ja houve modificacao no status
    sSQL = "SELECT * FROM ConexaoGerenciador WHERE ID_Empresa = " & ID_Empresa & " ORDER BY IP"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                'verificarDadosStatus Rst.fields("id"), Rst.fields("Usuario"), Rst.fields("Status")
                status = Rst.fields("Status")
                If InStr(status, "CHECAR") <> 0 Then
                    Hr = Trim(Mid(status, InStr(status, "-") + 1, Len(status)))
                    min = DateDiff("n", Hr, Time)
                    If min > 1 Then
                        MsgBox "apagando"
                        'RegistroExcluir "ConexaoGerenciador", "id=" & Rst.fields("id")
                    End If
                End If
                
                Rst.MoveNext
            Loop
    End If
    Rst.Close
End Sub
