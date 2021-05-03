Attribute VB_Name = "Modulo_particularidadesMC"
Option Explicit

Public Sub MontarBaseDados_MMxPOL()
    '*
    '* 24/10/2012
    '* Monta a tabela na base de dados para conversao de
    '* polegada x milimetro
    '*
    Dim a(30)       As Variant
    Dim i           As Integer
    Dim pol         As Integer
    Dim sPol        As String
    Dim sMM         As String
    Dim Tabela      As String
    Dim sSQL        As String
    Dim vReg(5)     As Variant
    Dim cReg        As Integer
    
    Tabela = "conv_milimetro_polegada"
    
    a(0) = Array("1/32", "0,79")
    a(1) = Array("1/16", "1,58")
    a(2) = Array("3/32", "2,38")
    a(3) = Array("1/8", "3,18")
    a(4) = Array("5/32", "3,96")
    a(5) = Array("3/16", "4,76")
    a(6) = Array("7/32", "5,56")
    a(7) = Array("1/4", "6,35")
    a(8) = Array("9/32", "7,14")
    a(9) = Array("5/16", "7,94")
    a(10) = Array("11/32", "8,73")
    a(11) = Array("3/8", "9,53")
    a(12) = Array("13/32", "10,32")
    a(13) = Array("7/16", "11,11")
    a(14) = Array("15/32", "11,91")
    a(15) = Array("1/2", "12,70")
    a(16) = Array("17/32", "13,49")
    a(17) = Array("9/16", "14,29")
    a(18) = Array("19/32", "15,08")
    a(19) = Array("5/8", "15,87")
    a(20) = Array("21/32", "16,67")
    a(21) = Array("11/16", "17,46")
    a(22) = Array("23/32", "18,26")
    a(23) = Array("3/4", "19,05")
    a(24) = Array("25/32", "19,84")
    a(25) = Array("13/16", "20,64")
    a(26) = Array("27/32", "21,43")
    a(27) = Array("7/8", "22,22")
    a(28) = Array("29/32", "23,02")
    a(29) = Array("15/16", "23,81")
    a(30) = Array("31/32", "24,61")
    
    
    BD.Execute "DROP TABLE IF EXISTS " & LCase(Tabela)
    
    sSQL = "CREATE TABLE IF NOT EXISTS " & Tabela & _
           " (Id INT(11) NOT NULL AUTO_INCREMENT," & _
           "Id_Empresa INT default Null," & _
           "DtHr VARCHAR(20) default Null," & _
           "UsuID INT default Null," & _
           "polegada VARCHAR(200) default Null," & _
           "milimetro VARCHAR(200) default Null," & _
           "PRIMARY KEY (Id))"
    
    BD.Execute sSQL
        
    For pol = 0 To 25
        If pol > 0 Then
            sPol = pol
            sMM = 25.4 * pol
            cReg = 0
            vReg(cReg) = Array("polegada", sPol, "S"): cReg = cReg + 1
            vReg(cReg) = Array("milimetro", sMM, "S"): cReg = cReg + 1
            cReg = cReg - 1
            RegistroIncluir Tabela, vReg, cReg
        End If
        
        For i = 0 To 30
            sPol = IIf(pol = 0, a(i)(0), pol & "." & a(i)(0))
            sMM = pol * 25.4 + a(i)(1)
            cReg = 0
            vReg(cReg) = Array("polegada", sPol, "S"): cReg = cReg + 1
            vReg(cReg) = Array("milimetro", sMM, "S"): cReg = cReg + 1
            cReg = cReg - 1
            RegistroIncluir Tabela, vReg, cReg
        Next
    Next
        
        MsgBox "FIM"
End Sub
Public Function MontarReferenciaProduto(idProd As Integer) As String
    On Error GoTo MntErro
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim sTexto  As String
    
    sTexto = LCase(pgDadosEstoqueProduto(idProd).Descricao)
    Dim i As Integer
    Dim c As Integer
    Dim a(100, 100) As Variant
    c = 0
    For i = 1 To Len(sTexto)
        If Mid(sTexto, i, 1) = """" Then
            a(c, 0) = pgPol(Mid(sTexto, 1, i))
            sSQL = "SELECT * FROM conv_milimetro_polegada " & _
                 "WHERE polegada='" & a(c, 0) & "'"
            Set Rst = RegistroBuscar(sSQL)
            If Rst.BOF And Rst.EOF Then
                a(c, 1) = "0"
                Else
                a(c, 1) = cNull(Rst.fields("milimetro"))
            End If
            
            'MsgBox a(c, 0) & " - " & ZE(Replace(ChkVal(CStr(a(c, 1)), 0, 2), ".", ""), 6)
            c = c + 1
        End If
    Next
    Dim X As Integer
    Dim cod As String
    Dim y As String
    cod = ""
    For X = 0 To c - 1
        y = Replace(ChkVal(CStr(a(X, 1)), 0, 2), ".", "")
        y = Left(String(6, "0"), 6 - Len(Trim(y))) & Trim(y)
        cod = cod & y
        
    Next
    cod = ZE(CInt(pgDadosEstoqueProduto(idProd).Grupo), 6) & ZE(CInt(pgDadosEstoqueProduto(idProd).subGrupo), 6) & cod
    MontarReferenciaProduto = cod
    Exit Function
MntErro:
    MontarReferenciaProduto = "*** Erro ***"
End Function
Private Function pgPol(sTexto As String) As String
    Dim i As Integer
    Dim l As Integer
    l = Len(sTexto)
    For i = l To 0 Step -1
        If Mid(sTexto, i, 1) = " " Then
            Exit For
        End If
    Next
    pgPol = Mid(sTexto, i, l)
    pgPol = Trim(Replace(pgPol, """", ""))

End Function

