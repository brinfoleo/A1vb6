Attribute VB_Name = "ModuloNumeroSerial"

Option Explicit
Dim n(3, 99) As String
Dim i As Integer
Dim ix As Integer

Public Function Numero_Serial() As String
    Dim DtInstall1      As String
    Dim DtResult1       As String
    Dim NSerieDV1       As String
    Dim NInst1          As String
    Dim Grupo1          As String
    Dim NSerieHd        As String
    Dim DtBase          As Date
    
    NSerieHd = Mid(String(10, "0"), 1, 10 - Len(NumSerieHD)) & NumSerieHD

    'Calcula o Digito verificador do num do HD
    NSerieDV1 = CalcDV(NSerieHd)
    DtBase = "21/11/1978"
    DtInstall1 = "08/02/2011" 'DTPicker2.Value
    If CDate(DtInstall1) < CDate(DtBase) Then
        MsgBox "Data do sistema Invalida por favor verifique", vbInformation, "atençao"
        Exit Function
    End If
    DtResult1 = Trim(Val(Format(DtInstall1, "yyyymmdd")) - Val(Format(DtBase, "yyyymmdd")))
    
    If Len(DtResult1) > 6 Then
        MsgBox "Data de instalação invalida"
        Exit Function
    End If
    DtResult1 = Mid(String(6, "0"), 1, 6 - Len(DtResult1)) & DtResult1
    NInst1 = DtResult1 & NSerieHd & NSerieDV1
    'O grupo sera selecionado de acordo com o numero do dv do HD
    Grupo1 = NGrupo(NSerieDV1)
    Numero_Serial = ConvNum(Grupo1, NInst1) & NSerieDV1
End Function

Private Function NumSerieHD()
    Dim Unidade As String
    Dim lSerial As Long
    Dim fso As New FileSystemObject, drvDrive As Drive
    '06.-2.2017
    'Unidade = Left(CurDir, 2)
    Unidade = Left(App.Path, 2)

    Set drvDrive = fso.GetDrive(Left(fso.GetDriveName(Unidade), 2))
    lSerial = drvDrive.SerialNumber
    
    NumSerieHD = IIf(lSerial < 0, lSerial * -1, lSerial)
        

End Function

Private Function CalcDV(Num As String)
    'Dim i As Integer
    Dim fator As Integer
    Dim Result As Integer
    
    fator = 1
    
    For i = 1 To Len(Num)
        Result = Result + IIf(Len(fator * (Mid(Num, i, 1))) > 2, Right(fator * (Mid(Num, i, 1)), 1), fator * (Mid(Num, i, 1)))
        fator = IIf(fator = 1, 2, 1)
    Next
    CalcDV = Right(Result, 1)
End Function
Private Function ConvLetra(Grupo As String, Letra As String) As String
    Dim npeg As String
    Dim texto As String
    Select Case Grupo
        Case 1
            Call Grupo1
        Case 2
            Call Grupo2
        Case 3
            Call Grupo3
    End Select
    
    For i = 1 To Len(Letra)
        npeg = Mid(Letra, i, 1)
        If Not IsNumeric(npeg) Then
                For ix = 0 To 99
                    If n(Grupo, ix) = npeg Then
                        texto = texto & IIf(Len(Trim(ix)) = 1, "0" & ix, ix)
                        Exit For
                    End If
                Next
            Else
                If IsNumeric(Mid(Letra, IIf(i = 1, i, i - 1), 1)) Then
                        texto = texto & npeg
                    Else
                        If IsNumeric(Mid(Letra, i + 1, 1)) Or Mid(Letra, i + 1, 1) = "" Then
                                texto = texto & npeg
                            Else
                                texto = texto & IIf(Len(npeg) = 1, "0" & npeg, npeg)
                        End If
                End If
        End If
    Next
    ConvLetra = texto
End Function
Private Function ConvNum(Grupo As String, Num As String)
    Dim npeg As String
    Dim texto As String
    Select Case Grupo
        Case 1
            Call Grupo1
        Case 2
            Call Grupo2
        Case 3
            Call Grupo3
    End Select
    
    For i = 1 To Len(Num) Step 2
        npeg = Mid(Num, i, 2)
        If Not n(Grupo, CInt(npeg)) = "" Then
                If Len(npeg) = 1 Then
                        texto = texto & npeg
                    Else
                        texto = texto & n(Grupo, npeg)
                End If
            Else
                If Mid(Num, i + 1, 2) = "" Then
                        texto = texto & npeg
                    Else
                        texto = texto & IIf(Len(npeg) = 1, "0" & npeg, npeg)
                End If
        End If
    Next
    ConvNum = texto
End Function

Private Sub Grupo1()
    'Grupo I
    n(1, 0) = "B"
    n(1, 2) = "W"
    n(1, 4) = "s"
    n(1, 6) = "u"
    n(1, 8) = "F"
    n(1, 10) = "H"
    n(1, 12) = "r"
    n(1, 14) = "I"
    n(1, 16) = "L"
    n(1, 18) = "t"
    n(1, 20) = "M"
    n(1, 22) = "O"
    n(1, 24) = "A"
    n(1, 26) = "P"
    n(1, 28) = "b"
    n(1, 30) = "Q"
    n(1, 32) = "R"
    n(1, 34) = "S"
    n(1, 36) = "U"
    n(1, 38) = "z"
    n(1, 40) = "V"
    n(1, 42) = "X"
    n(1, 44) = "Y"
    n(1, 46) = "Z"
    n(1, 48) = "a"
    n(1, 50) = "G"
    n(1, 52) = "c"
    n(1, 54) = "d"
    n(1, 56) = "e"
    n(1, 58) = "f"
    n(1, 60) = "y"
    n(1, 62) = "g"
    n(1, 64) = "h"
    n(1, 66) = "i"
    n(1, 68) = "j"
    n(1, 70) = "k"
    n(1, 72) = "D"
    n(1, 74) = "l"
    n(1, 76) = "E"
    n(1, 78) = "m"
    n(1, 80) = "n"
    n(1, 82) = "C"
    n(1, 84) = "o"
    n(1, 86) = "p"
    n(1, 88) = "q"
    n(1, 90) = "J"
    n(1, 92) = "K"
    n(1, 94) = "v"
    n(1, 96) = "w"
    n(1, 98) = "T"
    n(1, 99) = "x"
    n(1, 3) = "N"
End Sub
Private Sub Grupo2()
    'Grupo II
    n(2, 0) = "Y"
    n(2, 1) = "a"
    n(2, 3) = "G"
    n(2, 5) = "c"
    n(2, 7) = "d"
    n(2, 9) = "e"
    n(2, 11) = "f"
    n(2, 13) = "y"
    n(2, 15) = "g"
    n(2, 17) = "h"
    n(2, 19) = "i"
    n(2, 21) = "j"
    n(2, 23) = "k"
    n(2, 25) = "D"
    n(2, 27) = "l"
    n(2, 29) = "E"
    n(2, 31) = "m"
    n(2, 33) = "n"
    n(2, 35) = "C"
    n(2, 37) = "o"
    n(2, 39) = "p"
    n(2, 41) = "q"
    n(2, 43) = "J"
    n(2, 45) = "K"
    n(2, 47) = "v"
    n(2, 49) = "w"
    n(2, 51) = "T"
    n(2, 53) = "x"
    n(2, 55) = "N"
    n(2, 57) = "B"
    n(2, 59) = "W"
    n(2, 61) = "s"
    n(2, 63) = "u"
    n(2, 65) = "F"
    n(2, 67) = "H"
    n(2, 69) = "r"
    n(2, 71) = "I"
    n(2, 73) = "L"
    n(2, 75) = "t"
    n(2, 77) = "M"
    n(2, 79) = "O"
    n(2, 81) = "A"
    n(2, 83) = "P"
    n(2, 85) = "b"
    n(2, 87) = "Q"
    n(2, 89) = "R"
    n(2, 91) = "S"
    n(2, 93) = "U"
    n(2, 95) = "z"
    n(2, 97) = "V"
    n(2, 99) = "X"
    n(2, 90) = "Z"
End Sub
Private Sub Grupo3()
    'Grupo III
    n(3, 0) = "B"
    n(3, 3) = "W"
    n(3, 6) = "s"
    n(3, 9) = "u"
    n(3, 12) = "F"
    n(3, 15) = "H"
    n(3, 18) = "r"
    n(3, 21) = "I"
    n(3, 24) = "L"
    n(3, 27) = "t"
    n(3, 30) = "M"
    n(3, 33) = "O"
    n(3, 36) = "A"
    n(3, 39) = "P"
    n(3, 42) = "b"
    n(3, 45) = "Q"
    n(3, 48) = "R"
    n(3, 51) = "S"
    n(3, 54) = "U"
    n(3, 57) = "E"
    n(3, 59) = "m"
    n(3, 62) = "n"
    n(3, 65) = "C"
    n(3, 68) = "o"
    n(3, 71) = "p"
    n(3, 74) = "q"
    n(3, 77) = "J"
    n(3, 80) = "K"
    n(3, 83) = "v"
    n(3, 86) = "w"
    n(3, 89) = "T"
    n(3, 91) = "x"
    n(3, 94) = "N"
    n(3, 97) = "z"
    n(3, 2) = "V"
    n(3, 4) = "X"
    n(3, 6) = "Y"
    n(3, 8) = "Z"
    n(3, 10) = "a"
    n(3, 13) = "G"
    n(3, 16) = "c"
    n(3, 96) = "d"
    n(3, 76) = "e"
    n(3, 63) = "f"
    n(3, 82) = "y"
    n(3, 29) = "g"
    n(3, 55) = "h"
    n(3, 81) = "i"
    n(3, 73) = "j"
    n(3, 37) = "k"
    n(3, 52) = "D"
    n(3, 99) = "l"
End Sub

Private Function NGrupo(Num As String)
    Select Case Num
        Case 0
            NGrupo = 3
        Case 1
            NGrupo = 2
        Case 2
            NGrupo = 1
        Case 3
            NGrupo = 3
        Case 4
            NGrupo = 2
        Case 5
            NGrupo = 1
        Case 6
            NGrupo = 3
        Case 7
            NGrupo = 2
        Case 8
            NGrupo = 1
        Case 9
            NGrupo = 3
    End Select
End Function

