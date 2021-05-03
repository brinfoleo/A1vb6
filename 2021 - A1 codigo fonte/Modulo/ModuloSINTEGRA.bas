Attribute VB_Name = "ModuloSINTEGRA"
Option Explicit
Dim pathFile As String
'*
'* Rio de Janeiro 18/02/2013
'* Modulo desenvolvido para atender as necessidades do SINTEGRA RJ
'* usando como base de dados o sistema de arquivos/registros da NFe
'* emitida pelo sistema
'*
Public Sub gerarSintegra(fDestino As String, Periodo As String)
    '
    'periodo deve obedecer o formato mm/aaaa
    '
    Dim DtIni As Date
    Dim DtFin As Date
    Dim totReg50 As Integer
    
    pathFile = fDestino
    DtIni = "01/" & Right(Periodo, 2) & "/" & Left(Periodo, 4)
    DtFin = CalcData("01", 3, DtIni)
    
    
    
    
    Reg10 DtIni, DtFin
    Reg11
    totReg50 = Reg50(DtIni, DtFin)
    Reg90 1, 1, totReg50
    'MsgBox "Arquivo gerado com sucesso!", vbInformation, App.EXEName

End Sub
Private Function MontarSequenciaRegistro50(sReg50 As String, chv As String) As Integer
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim cont    As Integer
    Dim Reg     As String
    
    Dim tICMS As String
    Dim tCFOP As String
    
    Dim vNF     As String
    Dim vBC     As String
    Dim vICMS   As String
    
    cont = 0
    vNF = 0
    vBC = 0
    vICMS = 0
    
    sSQL = "SELECT * FROM faturamentonfeitens " & _
           "WHERE idnfe='" & chv & "' " & _
           "ORDER BY det_CFOP, ICMS_pICMS"
         
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            
        Else
            
            Rst.MoveFirst
            tCFOP = Rst.fields("det_CFOP")
            tICMS = Rst.fields("ICMS_pICMS")
            
            Do Until Rst.EOF
                If InStr(chv, "06124") Then
                    MsgBox "ops"
                End If
                
                If tCFOP <> Rst.fields("det_cfop") Or tICMS <> Rst.fields("ICMS_pICMS") Then
                        cont = cont + 1
                        Reg = Left("0000", 4 - Len(Trim(tCFOP))) & Trim(tCFOP) & _
                              "P" & _
                              convValor(vNF, 13, 2) & _
                              convValor(vBC, 13, 2) & _
                              convValor(vICMS, 13, 2) & _
                              convValor("0", 13) & _
                              convValor("0", 13) & _
                              convValor(tICMS, 4, 2) & _
                              IIf(Trim(PgDadosNotaFiscal(Rst.fields("idnfe")).canc_nProt) = "", "N", "S")
                        
                            GrvArq sReg50 & Reg
                        tCFOP = Rst.fields("det_CFOP")
                        tICMS = Rst.fields("ICMS_pICMS")
                        vNF = ChkVal(Rst.fields("det_vProd"), 0, cDecMoeda)
                        vBC = ChkVal(Rst.fields("ICMS_vBC"), 0, cDecMoeda)
                        vICMS = ChkVal(Rst.fields("ICMS_vICMS"), 0, cDecMoeda)
                    Else
                        vNF = Val(ChkVal(vNF, 0, cDecMoeda)) + Val(ChkVal(Rst.fields("det_vProd"), 0, cDecMoeda))
                        vBC = Val(ChkVal(vBC, 0, cDecMoeda)) + Val(ChkVal(Rst.fields("ICMS_vBC"), 0, cDecMoeda))
                        vICMS = Val(ChkVal(vICMS, 0, cDecMoeda)) + Val(ChkVal(Rst.fields("ICMS_vICMS"), 0, cDecMoeda))
                End If
                Rst.MoveNext
            Loop
            cont = cont + 1
            Reg = Left("0000", 4 - Len(Trim(tCFOP))) & Trim(tCFOP) & _
                              "P" & _
                              convValor(vNF, 13, 2) & _
                              convValor(vBC, 13, 2) & _
                              convValor(vICMS, 13, 2) & _
                              convValor("0", 13) & _
                              convValor("0", 13) & _
                              convValor(tICMS, 4, 2) & _
                              IIf(Trim(PgDadosNotaFiscal(chv).canc_nProt) = "", "N", "S")
                        
            GrvArq sReg50 & Reg
            
    End If
    Rst.Close
    MontarSequenciaRegistro50 = cont
    'Left("0000", 4 - Len(Trim("0"))) & Trim("0") & _
    "P" & _
    convValor(Rst.fields("total_vNF"), 13) & _
    convValor(Rst.fields("total_vBC"), 13) & _
    convValor(Rst.fields("total_vICMS"), 13) & _
    convValor("0", 13) & _
    convValor(Rst.fields("total_vOutro"), 13) & _
    convValor("0", 4) & _
    IIf(Trim(PgDadosNotaFiscal(Rst.fields("idnfe")).canc_nProt) = "", "N", "S")
End Function

Private Sub Reg10(DtIni As Date, DtFin As Date)
'    On Error GoTo ErrReg10
    
    Dim Reg10   As String
    Dim tmp     As String
    '*********************************************************************************************
    '*********************************************************************************************
    '****** Registro 10
    '*********************************************************************************************
    '*********************************************************************************************
    
    
    Reg10 = "10"
            
    'CNPJ
    tmp = Mid(String(14, "0"), 1, 14 - Len(Trim(PgDadosEmpresa(ID_Empresa).CNPJ))) & PgDadosEmpresa(ID_Empresa).CNPJ
    Reg10 = Reg10 & tmp
    'IE
    tmp = Trim(PgDadosEmpresa(ID_Empresa).IE) & Mid(String(14, " "), 1, 14 - Len(Trim(PgDadosEmpresa(ID_Empresa).IE)))
    Reg10 = Reg10 & tmp
    'NOME
    tmp = Trim(PgDadosEmpresa(ID_Empresa).Nome) & Mid(String(35, " "), 1, 35 - Len(PgDadosEmpresa(ID_Empresa).Nome))
    Reg10 = Reg10 & tmp
         
    tmp = Trim(PgDadosEmpresa(ID_Empresa).Mun) & Mid(String(30, " "), 1, 30 - Len(PgDadosEmpresa(ID_Empresa).Mun))
    Reg10 = Reg10 & tmp
    'UF
    tmp = PgDadosEmpresa(ID_Empresa).UF
    Reg10 = Reg10 & tmp
         
    'Fone1
    tmp = Mid(String(10, "0"), 1, 10 - Len(RS(PgDadosEmpresa(ID_Empresa).Fone))) & Trim(RS(PgDadosEmpresa(ID_Empresa).Fone))
    Reg10 = Reg10 & tmp
            
    'DtInicial
    tmp = Format(DtIni, "YYYYMMDD") 'MesAno & "01"
    Reg10 = Reg10 & tmp
            
    'DtFinal
    tmp = Format(DtFin, "YYYYMMDD") 'MesAno & PgDiasdoMes(dtFin)
    Reg10 = Reg10 & tmp
            
    'Cod. Identificacao do Convenio
    Reg10 = Reg10 & "3"
            
    'Cod. Natureza da operacao
    Reg10 = Reg10 & "3"
    'Cod. Finalidade do arq magnetico
    Reg10 = Reg10 & "1"
            
    'Grava em arquivo o Registro 10
    GrvArq Reg10
    
    Exit Sub
ErrReg10:
    Reg10 = ""
    MsgBox Err.Description, vbCritical, Err.Number
    Exit Sub
End Sub
Private Sub Reg11()
    On Error GoTo ErrReg11
    Dim tmp     As String
    Dim Reg11   As String
    '*********************************************************************************************
    '*********************************************************************************************
    '****** Registro 11
    '*********************************************************************************************
    '*********************************************************************************************
            Reg11 = "11"
            
            'Lgr
            tmp = Trim(PgDadosEmpresa(ID_Empresa).Lgr) & Mid(String(34, " "), 1, 34 - Len(Trim(PgDadosEmpresa(ID_Empresa).Lgr)))
            Reg11 = Reg11 & tmp
            'Nro
            tmp = Left("00000", 5 - Len(Trim(PgDadosEmpresa(ID_Empresa).Nro))) & PgDadosEmpresa(ID_Empresa).Nro
            Reg11 = Reg11 & tmp
            'Cpl
            tmp = Trim(PgDadosEmpresa(ID_Empresa).Cpl) & Mid(String(34, " "), 1, 22 - Len(Trim(PgDadosEmpresa(ID_Empresa).Cpl)))
            Reg11 = Reg11 & tmp
            'Bairro
            tmp = Trim(PgDadosEmpresa(ID_Empresa).Bairro) & Mid(String(15, " "), 1, 15 - Len(Trim(PgDadosEmpresa(ID_Empresa).Bairro)))
            Reg11 = Reg11 & tmp
            'CEP
            tmp = RS(PgDadosEmpresa(ID_Empresa).CEP)
            Reg11 = Reg11 & tmp
            
            'DtInicial
            'DtFinal
            'Cod. Identificacao do Convenio
            'Cod. Natureza da operacao
            'Cod. Finalidade do arq magnetico
            
            'Resp
            tmp = Trim(PgDadosEmpresa(ID_Empresa).crNome) & Mid(String(28, " "), 1, 28 - Len(Trim(PgDadosEmpresa(ID_Empresa).crNome)))
            Reg11 = Reg11 & tmp
            'Fone2
            tmp = Mid(String(12, "0"), 1, 12 - Len(Trim(RS(PgDadosEmpresa(ID_Empresa).Fone)))) & Trim(RS(PgDadosEmpresa(ID_Empresa).Fone))
            Reg11 = Reg11 & tmp
            
            'Grava em arquivo o Registro 11
            GrvArq Reg11
    'Rst.Close
    Exit Sub
ErrReg11:
    Reg11 = ""
    
    Exit Sub
End Sub

Private Function Reg50(DtIni As Date, DtFin As Date) As Integer
    'Retorna o numero de registros gravados em arquivo
    Dim sSQL    As String
    Dim Rst     As New Recordset
    Dim Reg     As String
    Dim cont    As Integer
    
    Dim IE      As String 'Variavel criada para modificar o vazio para ISENTO
    Dim nNF     As String 'Variavel criada para modificar de 9 para 6 caracteres
    
    'sSQL = "SELECT * FROM faturamentonfe " & _
           "WHERE EmpresaID=" & ID_Empresa & " AND " & _
           "Emissao BETWEEN ide_dEmi '" & MesAno & "' " & _
           "ORDER BY nNF"
           
    sSQL = "SELECT * " & _
           "FROM faturamentonfe WHERE ID_Empresa = " & ID_Empresa & _
           " AND ide_tpNF=1" & _
           " AND ide_natOP = 'VENDA'" & _
           " AND ide_dEmi >='" & Format(DtIni, "YYYY-MM-DD") & "' AND ide_dEmi <= '" & Format(DtFin, "YYYY-MM-DD") & "'" & _
           " ORDER BY ide_dEmi, ide_nNF"
    
           
           
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            Reg50 = 0
            
        Else
            Rst.MoveFirst
            cont = 0
            Do Until Rst.EOF
                IE = IIf(Trim(Rst.fields("dest_IE")) = "", "ISENTO", Rst.fields("dest_IE"))
                IE = Trim(IE) & Left(String(14, " "), 14 - Len(Trim(IE)))
                
                nNF = CLng(Rst.fields("ide_nNF"))
                nNF = ZE(CInt(nNF), 6)
                
                Reg = "50" & _
                        Left(String(14, "0"), 14 - Len(Trim(Rst.fields("dest_CNPJ")))) & Rst.fields("dest_CNPJ") & _
                        IE & _
                        RS(Format(Rst.fields("ide_dEmi"), "YYYYMMDD")) & _
                        Rst.fields("dest_UF") & _
                        Left("00", 2 - Len(Trim(Rst.fields("ide_mod")))) & Trim(Rst.fields("ide_mod")) & _
                        Left("000", 3 - Len(Trim(Rst.fields("ide_serie")))) & Trim(Rst.fields("ide_serie")) & _
                        nNF
                        
                        cont = cont + MontarSequenciaRegistro50(Reg, Rst.fields("IdNFe"))
                        'Left("0000", 4 - Len(Trim("0"))) & Trim("0") & _
                        "P" & _
                        convValor(Rst.fields("total_vNF"), 13) & _
                        convValor(Rst.fields("total_vBC"), 13) & _
                        convValor(Rst.fields("total_vICMS"), 13) & _
                        convValor("0", 13) & _
                        convValor(Rst.fields("total_vOutro"), 13) & _
                        convValor("0", 4) & _
                        IIf(Trim(PgDadosNotaFiscal(Rst.fields("idnfe")).canc_nProt) = "", "N", "S")
                'cont = cont + 1
                'Grava em arquivo o Registro 50
                'Debug.Print Len(Reg) & " >> " & Reg
                'GrvArq Reg
                Rst.MoveNext
            Loop
            Reg50 = cont
    End If
    Rst.Close
End Function
Private Sub Reg90(tReg10 As Integer, tReg11 As Integer, tReg50 As Integer)
    'On Error GoTo ErrReg10
    Dim Reg90   As String
    Dim tReg    As Integer
    Dim sSQL    As String
    Dim Rst     As New Recordset
    Dim CNPJ    As String
    Dim IE      As String
    
    'sSQL = "SELECT * FROM Empresa"
    'Rst.Open sSQL, BD
    'If Rst.BOF And Rst.EOF Then
    '        Rst.Close
    '        MsgBox "Erro Registro 90!" & vbCrLf & "Nenhuma empresa cadastrada!", vbInformation, App.EXEName
    '        Exit Sub
    '    Else
     '       Rst.MoveFirst
            CNPJ = Left(String(14, "0"), 14 - Len(PgDadosEmpresa(ID_Empresa).CNPJ)) & PgDadosEmpresa(ID_Empresa).CNPJ
            IE = Trim(PgDadosEmpresa(ID_Empresa).IE) & Left(String(14, " "), 14 - Len(Trim(PgDadosEmpresa(ID_Empresa).IE)))
     '       Rst.Close
    'End If
    
    '1 reg 90
    Reg90 = "90" & CNPJ & IE & "50" & Left(String(8, "0"), 8 - Len(Trim(tReg50))) & Trim(tReg50) & String(85, " ") & "2"
    GrvArq Reg90
    'Totalizador do reg 90
    tReg = tReg10 + tReg11 + tReg50 + 2
    Reg90 = "90" & CNPJ & IE & "99" & Left(String(8, "0"), 8 - Len(Trim(tReg))) & Trim(tReg) & String(85, " ") & "2"
    GrvArq Reg90
    'Rst.Close
End Sub


Private Function PgDiasdoMes(MesAno As String) As Integer
    Dim Dia As Integer
    Dim Mes As Integer
    Dim ano As Long
    
    Mes = Right(MesAno, 2)
    ano = Left(MesAno, 4)
    Select Case Mes
        Case "01" '"JANEIRO"
            Dia = 31
        Case "02" '"FEVEREIRO"
            Mes = 2
            If (ano Mod 4 = 0 And ano Mod 100 <> 0) Or (ano Mod 400 = 0) Then
                Dia = 29
            Else
                Dia = 28
            End If
            
        Case "03" '"MARÇO"
            Dia = 31
        Case "04" '"ABRIL"
            Dia = 30
        Case "05" '"MAIO"
            Dia = 31
        Case "06" '"JUNHO"
            Dia = 30
        Case "07" '"JULHO"
            Dia = 31
        Case "08" '"AGOSTO"
            Dia = 31
        Case "09" '"SETEMBRO"
            Dia = 30
        Case "10" '"OUTUBRO"
            Dia = 31
        Case "11" '"NOVEMBRO"
            Dia = 30
        Case "12" '"DEZEMBRO"
            Dia = 31
    End Select
    PgDiasdoMes = Dia
End Function

Private Sub GrvArq(sTexto As String)
    
    
    grvFile pathFile, sTexto
    
    
End Sub
Private Function convValor(Valor As String, qTam As Integer, Optional iDec As Integer) As String
    
    If qTam <= 0 Then
        MsgBox "Tamano da string invalida!", vbInformation, App.EXEName
        convValor = ""
        Exit Function
    End If
    If iDec <> 0 Then
        Valor = ChkVal(Valor, 0, iDec)
    End If
    Valor = Replace(Valor, ",", "")
    Valor = Replace(Valor, ".", "")
    Valor = Replace(Valor, ";", "")
    Valor = Replace(Valor, "$", "")
    Valor = Replace(Valor, "R", "")
    Valor = Replace(Valor, "+", "")
    Valor = Replace(Valor, "-", "")
    
    convValor = Left(String(qTam, "0"), qTam - Len(Trim(Valor))) & Valor ' & Mid(String(qTam, "0"), 1, Len(Trim(valor)))
    
End Function
