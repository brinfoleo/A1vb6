Attribute VB_Name = "ModuloFuncoes"
 Option Explicit
'Declaracao que faz uma pausa no sistema
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Declaracao para validar Inscricao Estadual
Declare Function ConsisteInscricaoEstadual Lib "DllInscE32" (ByVal Insc As String, ByVal uf As String) As Integer

'Declaracao para abrir arquivos como o IExplore
'Chamar a declaracao: ShellExecute hwnd, "open", (App.Path & "BancoDeDados.mdb"), "", "", 1
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

     
Type CalcTitulo
    DiasVencidos    As Integer
    vMulta          As String
    vMora           As String
    vTotal          As String
    vCalcFin        As String 'Variavel que guarda a diferenca entre cerd e deb do boleto
End Type
Public Enum OpcoesConexao
    Conectar = 1
    Desconectar = 2
    
End Enum

Public Sub chkUsuariosConectado()
    On Error GoTo trtErroUsuCnc
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim IP      As String
    Dim Usuario As String
    Dim vReg(1) As Variant
        'Exit Sub
    IP = MDIFormA1.wsMain.LocalIP
    Usuario = ID_Usuario & " - " & Trim(Mid(PgDadosUsuario(ID_Usuario).Nome, 6, Len(PgDadosUsuario(ID_Usuario).Nome)))
    
    'msfgConec.Rows = 1'
    
    sSQL = "SELECT * FROM ConexaoGerenciador WHERE ID_Empresa = " & ID_Empresa & _
           " AND Usuario='" & Usuario & "'" & _
           " AND IP='" & IP & "'" & _
           " ORDER BY IP"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'MsgBox "Erro ao localizar Usuario na tabela. Feche o aplicativo e abra novamente!", vbInformation, App.EXEName
            MonitoramentoConexao Conectar
        Else
            Rst.MoveFirst
            If Trim(LCase(cNull(Rst.Fields("Status")))) <> "conectado" Then
                vReg(0) = Array("Status", "CONECTADO", "S")
                RegistroAlterar "ConexaoGerenciador", vReg, 0, "Usuario='" & Usuario & "' AND IP='" & IP & "'"
            End If
            
    End If
    Rst.Close
    Exit Sub
trtErroUsuCnc:
    RegLog "", Err.Number, Err.Description
End Sub

Public Sub grvFile(pathFile As String, LinhaTexto As String)
    'On Error GoTo TrtErro
    'define o ObjPreview filesystem e demais variaveis
    Dim fso As New FileSystemObject
    Dim Arquivo As File
    Dim arquivoLog As TextStream
    Dim msg As String
    'Dim caminho As String
 
 
    
    'se o arquivo não existir então cria
    If fso.FileExists(pathFile) Then
            Set Arquivo = fso.GetFile(pathFile)
        Else
            Set arquivoLog = fso.CreateTextFile(pathFile)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(pathFile)
    End If
    'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
    
    'monta informações para gerar a linha da mensagem
    msg = LinhaTexto

    'inclui linhas no arquivo texto
    arquivoLog.WriteLine msg
    
    'escreve uma linha em branco no arquivo - se voce quiser
    'arquivoLog.WriteBlankLines (1)
    'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing
    'Debug.Print Len(msg)
    Exit Sub
TrtErro:
        MsgBox "Modulo: ModuloFuncoes" & _
           vbCrLf & "Função: grvFile " & vbCrLf & vbCrLf & _
           "Erro n.: " & Err.Number & _
           vbCrLf & vbCrLf & _
           "Descrição: " & Err.Description & _
           vbCrLf & vbCrLf & _
           "Arquivo: " & pathFile & _
           vbCrLf

End Sub





Public Sub MonitoramentoConexao(op As OpcoesConexao)
    
    Dim vReg(100)   As Variant
    Dim cReg        As Integer
    Dim Nome        As String
    Dim IP          As String
    Dim dt          As Date
    Dim Hr          As String
    Dim status      As String
    Dim Usuario     As String
    
    Nome = MDIFormA1.wsMain.LocalHostName
    IP = MDIFormA1.wsMain.LocalIP
    dt = Date
    Hr = Time
    Usuario = ID_Usuario & " - " & Trim(Mid(PgDadosUsuario(ID_Usuario).Nome, 6, Len(PgDadosUsuario(ID_Usuario).Nome)))
    
    status = "Conectado"
    
    
    
    cReg = 0
    vReg(cReg) = Array("idPrg", Numero_Serial, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Nome", Nome, "S"): cReg = cReg + 1
    vReg(cReg) = Array("IP", IP, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Usuario", Usuario, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ID_Usuario", ID_Usuario, "N"): cReg = cReg + 1
    vReg(cReg) = Array("Data", dt, "D"): cReg = cReg + 1
    vReg(cReg) = Array("Hora", Hr, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Status", status, "S") ': cReg = cReg + 1
    
    
    
    If op = Conectar Then
            '##################################################################################################################
            '### 23/01/2012 - Fun apresenta inperfeicoes ainda. Linha de comando inclusa para evitar ambiguidade
            RegistroExcluir "ConexaoGerenciador", "ID_Usuario = '" & ID_Usuario & "' AND IP = '" & IP & "'"
            '##################################################################################################################
            RegistroIncluir "ConexaoGerenciador", vReg, cReg
        Else
            RegistroExcluir "ConexaoGerenciador", "ID_Usuario = '" & ID_Usuario & "'"
    End If
    
    
End Sub
Public Function LimpaFormulario(Formulario As Form)
    Dim Controle    As Control
    Dim i           As Integer
    DoEvents
    For i = 0 To Formulario.Controls.Count - 1
        Set Controle = Formulario.Controls(i)
        If TypeOf Controle Is TextBox Then
            Controle.Text = ""
        End If
        If TypeOf Controle Is ComboBox Then
            Controle.Clear
        End If
        
        If TypeOf Controle Is MSFlexGrid Then
            Controle.Rows = 1
            Controle.Rows = 2
        End If
        If TypeOf Controle Is CheckBox Then
            Controle.Value = 0
        End If
    Next i
    
End Function
Public Function ChkPvTemNFe(idPV As Integer) As Boolean
    '***********************************************************
    '***Data: 26/07/2011
    '*** Obj: Checa se a PV tem NFe Cadastrada
    '
    '***********************************************************
    '
    'True - possui NFe Cadastrada
    'False - Nao possui NFe Cadastrada
    
    'Caso o sistema libere multiplas NFe de uma unica pv o sistema filtra aqui
    If PgDadosConfig.EmissaoNFesPV = 1 Then
        ChkPvTemNFe = False
        Exit Function
    End If
    
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    sSQL = "SELECT * FROM FaturamentoNFe WHERE ID_Empresa = " & ID_Empresa & " AND ger_idPV = " & idPV
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            ChkPvTemNFe = False
        Else
            ChkPvTemNFe = True
    End If
    Rst.Close


End Function
Public Function HDForm(Formulario As Form, sModo As Boolean)
    Dim Controle    As Control
    Dim i           As Integer
    
    For i = 0 To Formulario.Controls.Count - 1
        Set Controle = Formulario.Controls(i)
        If TypeOf Controle Is TextBox Then
            Controle.Enabled = sModo
        End If
        If TypeOf Controle Is ComboBox Then
            Controle.Enabled = sModo
        End If
        If TypeOf Controle Is MSFlexGrid Then
            Controle.Enabled = sModo
        End If
        If TypeOf Controle Is CommandButton Then
            Controle.Enabled = sModo
        End If
        If TypeOf Controle Is CheckBox Then
            Controle.Enabled = sModo
        End If
        If TypeOf Controle Is DTPicker Then
            Controle.Enabled = sModo
        End If
        If TypeOf Controle Is OptionButton Then
            Controle.Enabled = sModo
        End If
        If TypeOf Controle Is TreeView Then
            Controle.Enabled = sModo
        End If
         If TypeOf Controle Is ListBox Then
            Controle.Enabled = sModo
        End If
    Next i
    
    
End Function

Public Sub HDMenu(Formulario As Form, sModo As Boolean)
    Dim i As Integer
    
    For i = 1 To Formulario.tbMenu.Buttons.Count
        
        Select Case Formulario.tbMenu.Buttons(i).ToolTipText
            Case "Salvar"
                Formulario.tbMenu.Buttons(i).Enabled = IIf(sModo = True, False, True)
            Case "Cancelar"
                Formulario.tbMenu.Buttons(i).Enabled = IIf(sModo = True, False, True)
            Case "Manutenção da Tabela"
                '********************************************************************
                'Habilita/Desabilita o botao de manutencao na base de dados/Tabela
                If PgDadosConfig.MenuManutencaoTabelas = 0 Then
                        Formulario.tbMenu.Buttons(i).Visible = False
                        Formulario.tbMenu.Buttons(i).Enabled = False
                    Else
                        Formulario.tbMenu.Buttons(i).Visible = True
                        Formulario.tbMenu.Buttons(i).Enabled = True
                    End If
                '********************************************************************
            Case "Importar NF-e"
                Formulario.tbMenu.Buttons(i).Enabled = IIf(sModo = True, False, True)
            Case Else
                Formulario.tbMenu.Buttons(i).Enabled = sModo
        End Select
    Next
    
    
End Sub
Private Sub FormsdoSistema(nForm As Form)
'Private Sub FormsdoSistema(strForm As String, strDesc As String)
    'Funcao checa se o form ja esta cadastrado no sistema e seus direitos de uso
    On Error Resume Next
    Dim Rst             As Recordset
    Dim sSQL            As String
    Dim vReg(100)       As Variant
    Dim cReg            As Integer
    Dim strForm         As String
    Dim strDesc         As String
    strForm = nForm.Name
    strDesc = nForm.Caption
    
    cReg = 0
    'sSQL = "SELECT * FROM UsuGerenciadorFormularios WHERE ID_Empresa = " & ID_Empresa & " AND Formulario = '" & strForm & "'"
    sSQL = "SELECT * FROM UsuGerenciadorFormularios WHERE Formulario = '" & strForm & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'Grava o form
            vReg(cReg) = Array("Formulario", strForm, "S"): cReg = cReg + 1
            vReg(cReg) = Array("Descricao", strDesc, "S")
            RegistroIncluir "UsuGerenciadorFormularios", vReg, cReg
        Else
            vReg(cReg) = Array("Descricao", strDesc, "S")
            RegistroAlterar "UsuGerenciadorFormularios", vReg, cReg, "Formulario = '" & strForm & "'"
    End If
    Rst.Close
End Sub

Public Sub RegLog(ByVal Chave As String, ByVal IDLog As String, ByVal Descricao As String)
    On Error GoTo TrtErro
    Dim ipOrigem    As String
    Dim cTMP        As String
    '********************************************************
    'Chave - Sera usado para codificar a mensagem
    'IDLog - Sera usado como identificador do erro/mensagem
    'Descrição - Mensagem Armazenada
    'IPOrigem - Sera o IP da maquina que gerou o log
    '********************************************************
    
    ipOrigem = MDIFormA1.wsMain.LocalIP
    
   ' cTMP = PgDadosConfig.pFileArmazenamento & "\Log"
    'define o ObjPreview filesystem e demais variaveis
    Dim fso As New FileSystemObject
    Dim Arquivo As File
    Dim arquivoLog As TextStream
    Dim msg As String
    Dim caminho As String

    If Dir(App.Path & "\Log", vbDirectory) = "" Then
        MkDir SistemPath & "\Log"
    End If
    'If Dir(PgDadosConfig.pFileArmazenamento, vbDirectory) = "" Then
    '    MkDir PgDadosConfig.pFileArmazenamento
    'End If
    'If Dir(PgDadosConfig.pFileArmazenamento & "\Log", vbDirectory) = "" Then
    '    MkDir PgDadosConfig.pFileArmazenamento & "\Log"
    'End If
    'caminho = SistemPath & "\Log\" & Format(Date, "yyyymm") & ".txt"
    
    'caminho = PgDadosConfig.pFileArmazenamento & "\Log\" & Format(Date, "yyyymm") & ".txt"
    caminho = App.Path & "\Log\" & Format(Date, "yyyymm") & ".txt"
    'se o arquivo não existir então cria
    If fso.FileExists(caminho) Then
            Set Arquivo = fso.GetFile(caminho)
        Else
            Set arquivoLog = fso.CreateTextFile(caminho)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(caminho)
    End If
    'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
    
    'monta informações para gerar a linha da mensagem
    msg = Format(Date, "DDMMYYYY") & "|" & _
          Time & "|" & _
          ID_Usuario & "|" & _
          Chave & "|" & _
          IDLog & "|" & _
          ipOrigem & "|" & _
          Descricao
          

    'inclui linhas no arquivo texto
    arquivoLog.WriteLine msg
    
    'escreve uma linha em branco no arquivo - se voce quiser
    'arquivoLog.WriteBlankLines (1)
    'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing
    Exit Sub
TrtErro:
    
    MsgBox "Erro ao gerar registro de log.                                   " & _
           vbCrLf & _
           "local: " & cTMP & _
           vbCrLf & vbCrLf & _
           "Erro n.: " & Err.Number & _
           vbCrLf & vbCrLf & _
           "Descrição: " & Err.Description & _
           vbCrLf

End Sub

Public Function ChkVal(sValor As String, Dig As Integer, CasasDecimais As Integer)
   'On Error GoTo trtErroChkVal
    'Funcao para ser usado junto ao KEYPRESS do TextBox para formatar textos
    'em valores conforme as casas decimais
    '
    'Sintax:
    '       Checar_Valor(Valor,Digito, Nº de casas decimais)
    '
    'onde:
    '       Valor         - Será o valor ja digitado na funcao text
    '       Digito        - Sera o KeyAscii
    '       CasasDecimais - O numero de casas apos o ponto
    '
    'Para converter virgulas em pontos no valor utilize Dig como "0" (Zero)
    '
     '04.05.21 -  checks the decimal separator
     Dim valor As String
    If InStr(Format("0.00", "#0.00"), ",") <> 0 Then
        'its have coma(,) BR
            'valor = Format(sValor, "###,###,###,##0." & String(CasasDecimais, "0"))
            valor = sValor
            
            'valor = Replace(sValor, ".", "")
            
        Else
            valor = sValor
            valor = Replace(valor, ",", "")
    End If
    'Valor = Format(Valor, "###,###,###,##0." & String(CasasDecimais, "0"))
    '
    '******************************************
    Dim tmpVl As String
    tmpVl = valor
    If Not IsNumeric(valor) And Len(Trim(valor)) <> 0 Then
        ChkVal = 0
        Exit Function
    End If
    
    Select Case Dig
        Case 0 'Converter  virgula em ponto em um determinado valor
            If Not InStr(valor, ",") = 0 Then
                    If Not InStr(valor, "$") = 0 Then
                        valor = Trim(Right(valor, Len(valor) - 2))
                    End If
                    valor = Replace(valor, ".", "")
                    valor = Replace(valor, ",", ".")
                    ChkVal = ChkVal(valor, 0, CasasDecimais)
                    'Valor = Format(Valor, "#." & String(CasasDecimais, "0"))
                    
                    'ChkVal = IIf(Trim(Mid(Valor, 1, IIf(InStr(Valor, ",") = 0, 1, InStr(Valor, ",")) - 1)) = "", "0", Mid(Valor, 1, IIf(InStr(Valor, ",") = 0, 1, InStr(Valor, ",")) - 1)) & "." _
                            & IIf(Len(Mid(Valor, InStr(Valor, ",") + 1, Len(Valor))) = 1, Mid(Valor, InStr(Valor, ",") + 1, Len(Valor)) & 0, Mid(Valor, InStr(Valor, ",") + 1, Len(Valor)))
                Else
                    If InStr(valor, ".") = 0 Then
                        ChkVal = IIf(CasasDecimais = 0, valor, valor & "." & String(CasasDecimais, "0"))
                        ChkVal = IIf(Mid(ChkVal, 1, InStr(ChkVal, ".")) = ".", "0" & ChkVal, ChkVal)
                        Exit Function
                    End If
                    
                    'ChkVal = Mid(Valor, 1, InStr(Valor, ".") - 1) & "." &  IIf(Len(Mid(Valor, InStr(Valor, ".") + 1, Len(Valor))) = 1, Mid(Valor, InStr(Valor, ".") + 1, Len(Valor)) & 0, Mid(Valor, InStr(Valor, ".") + 1, Len(Valor)))
                    Dim tmpValor As String
                    tmpValor = Replace(Format(Replace(valor, ".", ","), "#." & String(CasasDecimais, "0")), ",", ".")
                    'tmpValor = Replace(Format(Valor, "##,##0.00"), ",", "")
                    
                    
                    'tmpValor = Replace(tmpValor, ",", "")
                    'tmpValor = valor

                    ChkVal = IIf(Trim(Mid(tmpValor, 1, InStr(tmpValor, ".") - 1)) = "", "0", Mid(tmpValor, 1, InStr(tmpValor, ".") - 1)) & "." & _
                            IIf(Len(Mid(tmpValor, InStr(tmpValor, ".") + 1, Len(tmpValor))) < CasasDecimais, Mid(tmpValor, InStr(tmpValor, ".") + 1, Len(tmpValor)) & Left(String(CasasDecimais, "0"), Val(CasasDecimais) - Len(Mid(tmpValor, InStr(tmpValor, ".") + 1, Len(tmpValor)))), _
                                Mid(tmpValor, InStr(tmpValor, ".") + 1, Len(tmpValor)))
            End If
        Case 8 'Back Space
            ChkVal = Dig
        
        Case 44 'Virgula
            If InStr(valor, ".") = 0 Then
                    ChkVal = 46
                Else
                    ChkVal = 0
            End If
        Case 46 'Ponto
            If InStr(valor, ".") = 0 Then
                    ChkVal = Dig
                Else
                    ChkVal = 0
            End If
        
        Case Else 'Caso contrario
            If InStr(valor, ".") = 0 Then
                    If IsNumeric(Chr(Dig)) Then
                            ChkVal = Dig
                        Else
                            ChkVal = 0
                    End If
                Else
                    If IsNumeric(Chr(Dig)) Then
                            If Len(Mid(valor, InStr(valor, "."), Len(valor))) >= CasasDecimais + 1 Then
                                    ChkVal = 0
                                Else
                                    ChkVal = Dig
                            End If
                        Else
                            ChkVal = 0
                    End If
            End If
    End Select
    Exit Function
trtErroChkVal:
    RegLog "Err n." & Err.Number, "", "(chkVal) " & Err.Description
    MsgBox Err.Number, vbCritical, "(modulo chkval) " & Err.Description
    
    Resume Next
End Function

Public Function ConvMoeda(valor As String)
    'Converte qualquer valor em formato Dinheiro conforme conf. do computador
    Dim a As String
    Dim b As String
    Dim c As String 'Armazena o sinal antes do numero caso negativo -
    Dim m As String
    
    m = Trim(Mid(Format("1", "currency"), 1, InStr(Format("1", "currency"), " ")))
    valor = ChkVal(valor, 0, cDecMoeda)
    
    If InStr(valor, "-") <> 0 Then
            valor = Replace(valor, "-", "")
            c = "-"
        Else
            c = ""
    End If
    'If Not InStr(Valor, ".") = 0 Then
                    If Not InStr(valor, "$") = 0 Then
                        valor = Trim(Right(valor, Len(valor) - 2))
                    End If
                    If cDecMoeda = 0 Then
                        ConvMoeda = valor
                        Exit Function
                    End If
                        
                    'ConvMoeda = Format(Mid(Valor, 1, InStr(Valor, ".") - 1) & "," & Mid(Valor, InStr(Valor, ".") + 1, Len(Valor)), "Currency")
                    valor = ChkVal(valor, 0, cDecMoeda)
                    a = Mid(valor, 1, InStr(valor, ".") - 1)
                    b = Mid(valor, InStr(valor, ".") + 1, Len(valor))
                    
                    If Len(a) >= 4 Then
                        a = Format(a, "0,###")
                        a = Replace(a, ",", ".")
                    End If
                     ConvMoeda = Trim(c) & m & " " & a & "," & b
                     
                    
                    
     '           Else
     '               ConvMoeda = Format(Valor, "Currency")
     '       End If
End Function


Public Function Validar_CNPJ_CPF(strDoc As String) As Boolean
    Dim tstrDoc As Variant
' Rotina objetiva a testar o cnpj ou cpf de um determinado cliente/fornecedor

    Dim i As Integer, j As Integer, Soma As Integer, Dg1 As Integer, flag As Integer
    Dim Caracter As String, Resto As String, Resto1 As String

    ' Antes testar a veracidade do strDoc verifica se a número suficiente para o mesmo
    If Trim(strDoc) = Empty Then Exit Function

    ' Inicialização da variável usada na multiplicação

    flag = 2
    Dim VstrDoc As String, Tamanho As Integer, Tcaracter As Integer, JCaracter As String
    For i = (Len(strDoc) - 2) To 1 Step -1
        If IsNumeric(Mid(Left(Trim(strDoc), Len(Trim(strDoc))), i, 1)) Then
            JCaracter = JCaracter & Mid(Left(Trim(strDoc), Len(Trim(strDoc))), i, 1)
        End If
    Next
    If IsNumeric(JCaracter) Then
        Tamanho = (Len(Trim(JCaracter)))
    End If

    'Criar um loop para o cálculo do número do strDoc

    For i = 1 To Tamanho
        Caracter = Mid(Left(JCaracter, Tamanho), i, 1)
        If IsNumeric(Caracter) Then
            Soma = Soma + (Val(Caracter) * flag)
            If Tamanho = 12 Then
                    flag = IIf(flag = 9, 2, flag + 1)
                Else
                    flag = IIf(flag = 0, 2, flag + 1)
            End If
        End If
    Next
    If Soma = 0 Then
        'MsgBox "O CNPJ/CPF não foi digitado corretamente. Falta ou passa caracteres", vbInformation + vbOKOnly, "Atenção"
        Exit Function
    End If
    If (Len(JCaracter) <> 8 And Len(JCaracter) <> 9) And (Len(JCaracter) <> 11 And Len(JCaracter) <> 12) Then
        'MsgBox "O CNPJ/CPF não foi digitado corretamente. Falta ou passa caracteres", vbInformation + vbOKOnly, "Atenção"
        Exit Function
    End If

    ' Encontrado o primeiro digito do Cnpfj
    Resto = IIf(Soma Mod 11 = 0 Or Soma Mod 11 = 1, 0, 11 - Soma Mod 11)
    flag = 2
    Soma = 0

    ' Criar um loop para o cálculo do segundo número do strDoc

    Tamanho = Tamanho + 1
    For i = 1 To Tamanho
        Caracter = Mid(Resto & JCaracter, i, 1)
        If IsNumeric(Caracter) Then
            Soma = Soma + (Val(Caracter) * flag)
            If Tamanho = 13 Then
                flag = IIf(flag = 9, 2, flag + 1)
                ElseIf Tamanho = 10 Then
                flag = flag + 1
            End If
        End If
    Next
    Resto1 = IIf(Soma Mod 11 = 0 Or Soma Mod 11 = 1, 0, 11 - Soma Mod 11)
    Dim NstrDoc As String
    For i = Len(JCaracter) To 1 Step -1
        NstrDoc = NstrDoc & Mid(JCaracter, i, 1)
    Next
    NstrDoc = NstrDoc & Resto & Resto1

    ' Faça a verificação do digito verificador do strDoc

    If (Trim(Resto & Resto1)) <> (Right(Trim(strDoc), 2)) Then
        If Tamanho = 10 Then
                        'MsgBox "O CPF digitado não está correto. Digite novamente", vbInformation + vbOKOnly, "Atenção"
                    ElseIf Tamanho = 13 Then
                        'MsgBox "O CNPJ digitado não está correto. Digite novamente", vbInformation + vbOKOnly, "Atenção"
                Else
                    'MsgBox "O código digitado não é válido para o uso do CNPJ/CPF. Digite Novamente.", vbInformation + vbOKOnly, "Atenção"
            End If
            Exit Function
        Else
            tstrDoc = True
    End If
    If Len(NstrDoc) = 11 Then
        strDoc = Format(NstrDoc, "@@@.@@@.@@@-@@")
            'If UCase(TIPO) = "J" Then
            '        'MsgBox "Neste campo só é permitido código do CNPJ", vbInformation + vbOKOnly, "Atenção"
            '        Validar_CNPJ_CPF = False
            '    Else
                    Validar_CNPJ_CPF = True
            'End If
        ElseIf Len(NstrDoc) = 14 Then
            strDoc = Format(NstrDoc, "@@.@@@.@@@/@@@@-@@")
            'If UCase(TIPO) = "C" Then
            '        'MsgBox "Neste campo só é permitido código do CPF", vbInformation + vbOKOnly, "Atenção"
            '        Validar_CNPJ_CPF = False
            '    Else
                    Validar_CNPJ_CPF = True
            'End If
    End If
    If Not tstrDoc Then strDoc = Empty
End Function

Public Function Validar_IE(strDoc As String, strUF As String) As Boolean
    '### 10.01.2017 ####
    'Funcao retirada pois a propria SEFAZ esta checando a situação cadastral da IE
    'Retirado tbm p/ diminuir o numero de DLL instalada no cliente
    Validar_IE = True
'    On Error GoTo trtErroIE
'    Validar_IE = IIf(ConsisteInscricaoEstadual(strDoc, strUF) = 0, True, False)
'    Exit Function
'trtErroIE:
'    MsgBox "Erro na validação da Inscrição Estadual!", vbCritical, App.EXEName 'Err.Description, vbCritical, Err.Number
'    Validar_IE = False
End Function
Public Function rc(vCampo As String) As String
    '********************************************************************************
    '*** Criado em   : 00/00/0000
    '*** Objetivo    : Remover Caracteres ESPECIAIS
    '*** Revisado em : 14/07/2011
    '********************************************************************************
    Dim i As Integer
    Dim vTem As String
    Dim vRes As String

    vRes = ""
    
    vCampo = Replace(vCampo, "''", """")
    
    If Len(i) > 0 Then
        For i = 1 To Len(vCampo)
            vTem = UCase(Mid$(vCampo, i, 1))
            Select Case vTem
                Case "Á", "À", "Ã", "Â", "Ä"
                    vTem = "A"
                Case "É", "È", "Ê", "Ë"
                    vTem = "E"
                Case "Í", "Ì", "Î", "Ï"
                    vTem = "I"
                Case "Ó", "Ò", "Õ", "Ô", "Ö"
                    vTem = "O"
                Case "Ú", "Ù", "Û", "Ü"
                    vTem = "U"
                Case "Ç"
                    vTem = "C"
                'Case "@"
                '    vTem = "."
                Case "º", "°"
                    vTem = ".o"
                Case "ª"
                    vTem = ".a"
                Case "'"
                    vTem = "''"
                Case ";"
                    vTem = ","
                Case "´"
                    vTem = ""
                Case "|"
                    vTem = "-"
            End Select
            vRes = vRes + vTem
        Next
    End If
    rc = vRes



End Function
Public Function RS(sCampo As String) As String 'Remover Separadores
    sCampo = Replace(sCampo, ":", Empty)
    sCampo = Replace(sCampo, ";", Empty)
    sCampo = Replace(sCampo, ",", Empty)
    sCampo = Replace(sCampo, ".", Empty)
    sCampo = Replace(sCampo, "/", Empty)
    sCampo = Replace(sCampo, "-", Empty)
    sCampo = Replace(sCampo, " ", Empty)
    sCampo = Replace(sCampo, "\", Empty)
    RS = sCampo
End Function


Public Function MovimentarContasPagarReceber(ContaPR As String, DtEmissao As Date, _
                                      nFat As String, vFat As String, _
                                      Tabela As String, _
                                      idSac As Integer, nSac As String, _
                                      dSac As String, nConta As Integer, _
                                      nCentroCusto As Integer, tpDoc As Integer, _
                                      PlanoContas As Integer, LinhaDigitavel As String, _
                                      NossoNumero As String, Vencimento As Date, _
                                      nDupl As String, Multa As String, _
                                      Juros As String, IdBanco As String, _
                                      Acrescimo As String, Abatimento As String, _
                                      Deducoes As String, MultaMora As String, _
                                      vDupl As String, Obs As String, _
                                      Optional idNFe As String, _
                                      Optional ContaFV As String, _
                                      Optional idContaFV As String) As Boolean
    

    Dim vReg(1000)  As Variant
    Dim cReg        As Integer
    Dim vCob        As String 'Valor Cobrado
    cReg = 0
    vReg(cReg) = Array("Tabela", Tabela, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ContaPR", ContaPR, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Emissao", DtEmissao, "D"): cReg = cReg + 1
    vReg(cReg) = Array("NumFatura", nFat, "S"): cReg = cReg + 1
    vReg(cReg) = Array("VlFatura", vFat, "S"): cReg = cReg + 1
    vReg(cReg) = Array("IdSacado", idSac, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Nome", nSac, "S"): cReg = cReg + 1
    vReg(cReg) = Array("CNPJ", dSac, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Conta", nConta, "N"): cReg = cReg + 1
    vReg(cReg) = Array("CentroCusto", nCentroCusto, "N"): cReg = cReg + 1
    vReg(cReg) = Array("TpDocumento", tpDoc, "N"): cReg = cReg + 1
    vReg(cReg) = Array("PlanoContas", PlanoContas, "N"): cReg = cReg + 1
    vReg(cReg) = Array("LinhaDigitavel", LinhaDigitavel, "S"): cReg = cReg + 1
    vReg(cReg) = Array("NossoNumero", NossoNumero, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Vencimento", Vencimento, "D"): cReg = cReg + 1
    vReg(cReg) = Array("NumDuplicata", nDupl, "S"): cReg = cReg + 1
    vReg(cReg) = Array("VlDuplicata", vDupl, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Multa", Multa, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Juros", Juros, "S"): cReg = cReg + 1
    vReg(cReg) = Array("idBanco", IdBanco, "S"): cReg = cReg + 1
    
    vReg(cReg) = Array("Acrescimo", Acrescimo, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Abatimento", Abatimento, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Deducoes", Deducoes, "S"): cReg = cReg + 1
    vReg(cReg) = Array("MultaMora", MultaMora, "S"): cReg = cReg + 1
    
    vCob = Val(ChkVal(vDupl, 0, cDecMoeda))
    vCob = Val(ChkVal(vCob, 0, cDecMoeda)) + Val(ChkVal(MultaMora, 0, cDecMoeda)) + Val(ChkVal(Acrescimo, 0, cDecMoeda))
    vCob = Val(ChkVal(vCob, 0, cDecMoeda)) - (Val(ChkVal(Deducoes, 0, cDecMoeda)) + Val(ChkVal(Abatimento, 0, cDecMoeda)))
    vCob = ChkVal(vCob, 0, cDecMoeda)
    '04/01/12 - VlCobrado removidor pois deve ser a soma do vlDupl + Movimentos
    vReg(cReg) = Array("vlCobrado", vCob, "S"): cReg = cReg + 1
    vReg(cReg) = Array("DiasProtesto", pgDadosConta(nConta).DiasProtesto, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Obs", Obs, "S"): cReg = cReg + 1
    vReg(cReg) = Array("ide_NFe", idNFe, "S"): cReg = cReg + 1
    
    If Trim(ContaFV) = "" Then
        ContaFV = "V"
        idContaFV = 0
    End If
    vReg(cReg) = Array("FixoVariavel", ContaFV, "S"): cReg = cReg + 1
    vReg(cReg) = Array("IdFixa", idContaFV, "S"): cReg = cReg + 1
    
    cReg = cReg - 1
    If RegistroIncluir("FinanceiroContasPRCadastro", vReg, cReg) = 0 Then
            MsgBox "Erro ao Incluir"
            MovimentarContasPagarReceber = False
        Else
            MovimentarContasPagarReceber = True
    End If

End Function
Public Function Extenso(ByVal valor As _
       Double, ByVal MoedaPlural As _
       String, ByVal MoedaSingular As _
       String) As String
  Dim StrValor As String, Negativo As Boolean
  Dim Buf As String, Parcial As Integer
  Dim Posicao As Integer, Unidades
  Dim Dezenas, Centenas, PotenciasSingular
  Dim PotenciasPlural

  Negativo = (valor < 0)
  valor = Abs(CDec(valor))
  If valor Then
    Unidades = Array(vbNullString, "Um", "Dois", _
               "Três", "Quatro", "Cinco", _
               "Seis", "Sete", "Oito", "Nove", _
               "Dez", "Onze", "Doze", "Treze", _
               "Quatorze", "Quinze", "Dezesseis", _
               "Dezessete", "Dezoito", "Dezenove")
    Dezenas = Array(vbNullString, vbNullString, _
              "Vinte", "Trinta", "Quarenta", _
              "Cinqüenta", "Sessenta", "Setenta", _
              "Oitenta", "Noventa")
    Centenas = Array(vbNullString, "Cento", _
               "Duzentos", "Trezentos", _
               "Quatrocentos", "Quinhentos", _
               "Seiscentos", "Setecentos", _
               "Oitocentos", "Novecentos")
    PotenciasSingular = Array(vbNullString, " Mil", _
                        " Milhão", " Bilhão", _
                        " Trilhão", " Quatrilhão")
    PotenciasPlural = Array(vbNullString, " Mil", _
                      " Milhões", " Bilhões", _
                      " Trilhões", " Quatrilhões")

    StrValor = Left(Format(valor, String(18, "0") & _
               ".000"), 18)
    For Posicao = 1 To 18 Step 3
      Parcial = Val(Mid(StrValor, Posicao, 3))
      If Parcial Then
        If Parcial = 1 Then
          Buf = "Um" & PotenciasSingular((18 - _
                Posicao) \ 3)
        ElseIf Parcial = 100 Then
          Buf = "Cem" & PotenciasSingular((18 - _
                Posicao) \ 3)
        Else
          Buf = Centenas(Parcial \ 100)
          Parcial = Parcial Mod 100
          If Parcial <> 0 And Buf <> vbNullString Then
            Buf = Buf & " e "
          End If
          If Parcial < 20 Then
            Buf = Buf & Unidades(Parcial)
          Else
            Buf = Buf & Dezenas(Parcial \ 10)
            Parcial = Parcial Mod 10
            If Parcial <> 0 And Buf <> vbNullString Then
              Buf = Buf & " e "
            End If
            Buf = Buf & Unidades(Parcial)
          End If
          Buf = Buf & PotenciasPlural((18 - Posicao) \ 3)
        End If
        If Buf <> vbNullString Then
          If Extenso <> vbNullString Then
            Parcial = Val(Mid(StrValor, Posicao, 3))
            If Posicao = 16 And (Parcial < 100 Or _
                (Parcial Mod 100) = 0) Then
              Extenso = Extenso & " e "
            Else
              Extenso = Extenso & ", "
            End If
          End If
          Extenso = Extenso & Buf
        End If
      End If
    Next
    If Extenso <> vbNullString Then
      If Negativo Then
        Extenso = "Menos " & Extenso
      End If
      If Int(valor) = 1 Then
        Extenso = Extenso & " " & MoedaSingular
      Else
        Extenso = Extenso & " " & MoedaPlural
      End If
    End If
    Parcial = Int((valor - Int(valor)) * _
              100 + 0.1)
    If Parcial Then
    'MsgBox "X"
      Buf = Extenso(Parcial, "Centavos", _
            "Centavo")
      If Extenso <> vbNullString Then
        Extenso = Extenso & " e "
      End If
      Extenso = Extenso & Buf
    End If
  End If
End Function
Public Function MovimentarEstoque(Mov As String, _
                                    idProduto As Integer, _
                                    Data As Date, _
                                    sDoc As String, _
                                    Qtd As String, _
                                    vUnit As String, _
                                    vTot As String, _
                                    Obs As String, _
                                    Optional Nome As String, _
                                    Optional NFe As String, _
                                    Optional idNome As Integer, _
                                    Optional docNome As String) As Boolean

    
    'Mov= e- entrada  s- saida
                                    
    Dim Rst1            As Recordset
    'Dim Rst2            As Recordset
    'Dim Rst3            As Recordset
    Dim sSQL            As String
    Dim SaldoI          As String
    Dim saldoF          As String
    Dim vDados(1000)    As Variant
    Dim c               As Integer
    
    If idProduto = 0 Then
        MsgBox "Codigo do produto não informado! Movimento de Estoque cancelado!", vbInformation, "Aviso"
        MovimentarEstoque = False
        Exit Function
    End If
    '*****************************************************************************************************************************
    '*** Data: 05/07/2011
    '*** Acao: Pega o saldo na tabela principal (EstoqueProduto) para calculo e lanc no kardex
    '*****************************************************************************************************************************
    'sSQL = "SELECT * FROM EstoqueProduto WHERE ID_Empresa = " & ID_Empresa & " AND Deposito = " & ID_Deposito & " AND id = " & idProduto
    
    'Set Rst1 = RegistroBuscar(sSQL)
    'If Rst1.BOF And Rst1.EOF Then
    '        MsgBox "Erro ao localizar Produto no Estoque.", vbInformation, "Aviso"
    '        MovimentarEstoque = False
    '        Exit Function
    '    Else
    '        Rst1.MoveFirst
            SaldoI = pgDadosEstoqueProduto(idProduto).Saldo ' IIf(IsNull(Rst1.Fields("Saldo")), 0, Rst1.Fields("Saldo"))
    'End If
    'Rst1.Close
    
    '*****************************************************************************************************************************
    
    '*****************************************************************************************************************************
    '*** Data: 05/07/2011
    '*** Acao: Grava a Movimentacao na tabela EstoqueProduto
    '*****************************************************************************************************************************
    c = 0
    
    If LCase(Mov) = "e" Then
            'vDados(c) = Array("Movimento", "SOMAR (+)", "S"): c = c + 1
            saldoF = Val(ChkVal(SaldoI, 0, cDecQtd)) + Val(ChkVal(Qtd, 0, cDecQtd))
            saldoF = ChkVal(saldoF, 0, cDecQtd)
        Else
            'vDados(c) = Array("Movimento", "SUBTRAIR (-)", "S"): c = c + 1
            saldoF = Val(ChkVal(SaldoI, 0, cDecQtd)) - Val(ChkVal(Qtd, 0, cDecQtd))
            saldoF = ChkVal(saldoF, 0, cDecQtd)
    End If
    
    vDados(c) = Array("Saldo", saldoF, "S") ':c=c+1
    If RegistroAlterar("EstoqueProduto", vDados, c, "Id = " & idProduto) = False Then
            MovimentarEstoque = False
        Else
            MovimentarEstoque = True
    End If
    '*****************************************************************************************************************************
    
    '*****************************************************************************************************************************
    '*** Data: 05/07/2011
    '*** Acao: Grava a Movimentacao na tabela EstoqueKardex
    '*****************************************************************************************************************************
    c = 0
    vDados(c) = Array("idProduto", idProduto, "N"): c = c + 1
    vDados(c) = Array("Deposito", ID_Deposito, "N"): c = c + 1
    vDados(c) = Array("Data", Data, "D"): c = c + 1
    vDados(c) = Array("DataMov", Date, "D"): c = c + 1
    vDados(c) = Array("Documento", sDoc, "S"): c = c + 1
    vDados(c) = Array("Descricao", "Movimentação automatica pelo sistema", "S"): c = c + 1
    vDados(c) = Array("Quantidade", Qtd, "S"): c = c + 1
    
    vDados(c) = Array("Unidade", pgDadosEstoqueProduto(idProduto).Unidade, "S"): c = c + 1
    
    vDados(c) = Array("ValorUnitario", ChkVal(vUnit, 0, cDecMoeda), "S"): c = c + 1
    vDados(c) = Array("ValorTotal", vTot, "S"): c = c + 1
    If LCase(Mov) = "e" Then
            vDados(c) = Array("Movimento", "SOMAR (+)", "S"): c = c + 1
            'SaldoF = Val(ChkVal(SaldoI, 0, cDecQtd)) + Val(ChkVal(Qtd, 0, cDecQtd))
            'SaldoF = ChkVal(SaldoF, 0, cDecQtd)
        Else
            vDados(c) = Array("Movimento", "SUBTRAIR (-)", "S"): c = c + 1
            'SaldoF = Val(ChkVal(SaldoI, 0, cDecQtd)) - Val(ChkVal(Qtd, 0, cDecQtd))
            'SaldoF = ChkVal(SaldoF, 0, cDecQtd)
    End If
    vDados(c) = Array("Saldo", saldoF, "S"): c = c + 1
    vDados(c) = Array("Obs", Obs, "S"): c = c + 1
    vDados(c) = Array("idNome", idNome, "N"): c = c + 1
    vDados(c) = Array("Nome", Nome, "S"): c = c + 1
    vDados(c) = Array("docNome", docNome, "S"): c = c + 1
    vDados(c) = Array("NFe", NFe, "S") ': c = c + 1
    If RegistroIncluir("EstoqueKardex", vDados, c) = 0 Then
            MovimentarEstoque = False
        Else
            MovimentarEstoque = True
    End If
    
    
    'estoqueprodutokit
    '17/12/2017 - BAIXA DE KIT NO ESTOQUE
    'Verifica se o item e um kit caso nao seja sai do loop
    Dim totalQtd As String
    sSQL = "SELECT * FROM EstoqueProdutoKit WHERE idProduto=" & idProduto
    Set Rst1 = RegistroBuscar(sSQL)
    If Rst1 Is Nothing Then Exit Function
    If Rst1.BOF And Rst1.EOF Then
            'MsgBox "Erro ao localizar Produto no Estoque.", vbInformation, "Aviso"
            'MovimentarEstoque = False
            Exit Function
        Else
            Rst1.MoveFirst
            Do Until Rst1.EOF
                totalQtd = Val(ChkVal(Rst1.Fields("qtd"), 0, cDecQtd)) * Val(ChkVal(Qtd, 0, cDecQtd))
                totalQtd = ChkVal(totalQtd, 0, cDecQtd)
                MovimentarEstoque Mov, Rst1.Fields("IdItemKit"), Data, sDoc, totalQtd, vUnit, vTot, Obs, Nome, NFe, idNome, docNome
                Rst1.MoveNext
            Loop
            'SaldoI = pgDadosEstoqueProduto(idProduto).Saldo ' IIf(IsNull(Rst1.Fields("Saldo")), 0, Rst1.Fields("Saldo"))
    End If
    Rst1.Close
    
    
End Function
Public Function FormPermissao(sForm As String, TpAcesso As String, idGrupo As Integer) As Boolean
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim Permissoes  As String
    'Dim nmForm      As Form
    'Set nmForm = sForm
    sSQL = "SELECT * FROM UsuGerenciadorAcessos WHERE ID_Empresa = " & ID_Empresa & " AND Formulario = '" & sForm & "' AND GrupoId = " & idGrupo
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            'MsgBox "Erro ao localizar permissões de acesso.", vbInformation, "Aviso"
            'FormsdoSistema sForm, ""
            FormPermissao = False
            Rst.Close
            'FormPermissao = False
            Exit Function
        Else
            Rst.MoveFirst
            Permissoes = Rst.Fields("Permissao")
    End If
    Rst.Close
    Select Case LCase(TpAcesso)
        Case "n" 'ovo
            FormPermissao = IIf(Mid(Permissoes, 1, 1) = 1, True, False)
        Case "a" 'lterar
            FormPermissao = IIf(Mid(Permissoes, 2, 1) = 1, True, False)
        Case "e" 'xcluir
            FormPermissao = IIf(Mid(Permissoes, 3, 1) = 1, True, False)
        Case "i" 'mprimir
            FormPermissao = IIf(Mid(Permissoes, 4, 1) = 1, True, False)
        Case "c" 'onsultar
            FormPermissao = IIf(Mid(Permissoes, 5, 1) = 1, True, False)
        Case Else
            MsgBox "Solicitação de acesso [" & TpAcesso & "] não localizado.", vbInformation, "Aviso"
    End Select
End Function
'Public Sub addSkin(F As Form)
    'Exit Sub
    
    'DoEvents
    'Dim sIndex As String
    'Dim vSkin As ACTIVESKINLibCtl.Skin
    'Dim skinFile As String
    ''ADD SKIN NO FORM
    'If Not ActivexRegistrado("ActiveSkin4.SkinLabel.1") Then Exit Sub
    'If sIndex = "" Then Set vSkin = MDIFormA1.Skin1
    'If skinFile <> "" Then vSkin.LoadSkin skinFile
    'vSkin.LoadSkin App.Path & "\skin\skin.skn"
    'vSkin.ApplySkin F.Hwnd
    'If Err Then Err.Clear
    'Dim Controle    As Control
    'Dim i           As Integer
    'DoEvents
    'For i = 0 To F.Controls.Count - 1
    '    Set Controle = F.Controls(i)
    '    If TypeOf Controle Is SSTab Then
    '        Controle.BackColor = F.BackColor
    '    End If
    'Next i
'End Sub
Private Function ActivexRegistrado(ByVal T_ProgId As String) As Boolean
    On Error Resume Next
    Dim Obj As Object
    Set Obj = CreateObject(T_ProgId)
    ActivexRegistrado = (Err.Number = 0)
    Err.Clear
End Function
Public Function chkAcesso(Formulario As Form, TpAcesso As String) As Boolean
    Dim sSQL        As String
    Dim Rst         As Recordset
    Dim Permissoes  As String
    Dim sForm       As String
    sForm = Formulario.Name
    
    
    'Adiciona SKIN nos FORMS
    '06.02.2015 - Projeto adiado devido a nfce
    'addSkin Formulario
    
    'Envia o form para que seja averiguada mudancas em seu label
    FormsdoSistema Formulario
    
    
    sSQL = "SELECT * FROM UsuGerenciadorAcessos WHERE ID_Empresa = " & ID_Empresa & " AND Formulario = '" & sForm & "' AND GrupoId = " & PgDadosUsuario(ID_Usuario).Grupo
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar permissões de acesso ao formulario.", vbInformation, "Aviso"
            chkAcesso = False
            Rst.Close
            'FormPermissao = False
            Exit Function
        Else
            Rst.MoveFirst
            Permissoes = Rst.Fields("Permissao")
    End If
    Rst.Close
    Select Case LCase(TpAcesso)
        Case "n" 'ovo
            chkAcesso = IIf(Mid(Permissoes, 1, 1) = 1, True, False)
        Case "a" 'lterar
            chkAcesso = IIf(Mid(Permissoes, 2, 1) = 1, True, False)
        Case "e" 'xcluir
            chkAcesso = IIf(Mid(Permissoes, 3, 1) = 1, True, False)
        Case "i" 'mprimir
            chkAcesso = IIf(Mid(Permissoes, 4, 1) = 1, True, False)
        Case "c" 'onsultar
            chkAcesso = IIf(Mid(Permissoes, 5, 1) = 1, True, False)
        Case Else
            MsgBox "Solicitação de acesso [" & TpAcesso & "] não localizado.", vbInformation, "Aviso"
    End Select
    If chkAcesso = False Then
        MsgBox "Acesso NEGADO!", vbCritical, "Aviso"
    End If
End Function
Public Function ChkNFeTemCCe(chvNFe As String) As Integer
    '###############################################################################
    '### Funcao para checarse existe CC-e e retorna o num. da CC-e
    '###############################################################################
    
    Dim Rst     As Recordset
    Dim sSQL    As String
    sSQL = "SELECT * FROM FaturamentoNFeCartaCorrecao WHERE chvNFe = '" & chvNFe & "'"
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            ChkNFeTemCCe = 0
        Else
            Rst.MoveFirst
            ChkNFeTemCCe = Rst.Fields("ID")
    End If
    Rst.Close


End Function
Public Sub MovimentarConta(idConta As Integer, _
                           cd As String, _
                           IdRegDoc As Long, _
                           Data As String, _
                           nDoc As String, _
                           tDoc As Integer, _
                           Descricao As String, _
                           valor As String)
     
    'nDoc - Numero do Documeto
    'tDoc - Codigo interno do Tipo de Documento
    
    Dim Saldo   As String
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim vReg(20)    As Variant
    Dim cReg    As Integer
    
    sSQL = "SELECT * FROM FinanceiroConta WHERE id = " & idConta
    Set Rst = RegistroBuscar(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar conta.", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            Saldo = ChkVal(IIf(IsNull(Rst.Fields("Saldo")), 0, Rst.Fields("Saldo")), 0, cDecMoeda)
            If cd = "C" Then
                    Saldo = Val(Saldo) + Val(ChkVal(valor, 0, cDecMoeda))
                Else
                    Saldo = Val(Saldo) - Val(ChkVal(valor, 0, cDecMoeda))
            End If
            Saldo = ChkVal(Saldo, 0, cDecMoeda)
    End If
    Rst.Close
    cReg = 0
    vReg(cReg) = Array("IdConta", idConta, "N"): cReg = cReg + 1
    vReg(cReg) = Array("IdRegDoc", IdRegDoc, "N"): cReg = cReg + 1
    vReg(cReg) = Array("Data", Data, "D"): cReg = cReg + 1
    vReg(cReg) = Array("Documento", nDoc, "S"): cReg = cReg + 1
    vReg(cReg) = Array("TpDoc", tDoc, "N"): cReg = cReg + 1
    vReg(cReg) = Array("Descricao", Descricao, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Valor", ChkVal(valor, 0, cDecMoeda), "S"): cReg = cReg + 1
    vReg(cReg) = Array("CD", cd, "S"): cReg = cReg + 1
    vReg(cReg) = Array("Saldo", ChkVal(Saldo, 0, cDecMoeda), "S") ': cReg = cReg + 1
    
    RegistroIncluir "FinanceiroContaHistorico", vReg, cReg
    
    cReg = 0
    vReg(cReg) = Array("Saldo", Saldo, "S")
    RegistroAlterar "FinanceiroConta", vReg, cReg, "id = " & idConta
    
End Sub
Public Function ZE(Num As Integer, qtdZeros As Integer) As String
    On Error GoTo TrtErro
    '#######################################################################
    '### Funcao para colocar zeros a esquerda                            ###
    '#######################################################################
    ZE = Left(String(qtdZeros, "0"), qtdZeros - Len(Trim(Num))) & Trim(Num)
    Exit Function
TrtErro:
    ZE = Trim(Num)
End Function
Public Function cNull(sDados As Variant) As String
    'Retorna um vazio caso haja nulo
    cNull = IIf(IsNull(sDados), "", sDados)
End Function
Public Function SoNumeros(iKey As Integer) As Integer
    'Modificado em 09.12.2014 - incluir a instrucao or iKey=13
    '
    If iKey = 8 Or iKey = 13 Then
        SoNumeros = iKey
        Exit Function
    End If
    If Not IsNumeric(Chr(iKey)) Then
        SoNumeros = 0
        Exit Function
    End If
    SoNumeros = iKey
End Function

Public Sub CopyArray(SourceArray As Variant, DestArray As Variant)

'###########################################################################
'### Clonagem de Array
'### 22/03/2012
'### Obs: A array destino nao deve ter declarado seu tamanho ou seja,
'###      dim a(10) as variant -> tamanho definido
'###      dim a     as variant -> tamanho nao declarado
'###########################################################################
   
    Dim l As Long, lUBound As Long, lLBound As Long

    If (Not IsArray(SourceArray)) Or (Not IsArray(DestArray)) Then Exit Sub
        lLBound = LBound(SourceArray)
        lUBound = UBound(SourceArray)

        'ReDim DestArray(lLBound To lUBound)
        'ReDim DestArray(lUBound)
        For l = lLBound To lUBound
            DestArray(l) = SourceArray(l)
        Next
End Sub



Public Function Bissexto(intAno As Integer) As Boolean
'############################################################
'### 26/03/2012  verifica se um ano é bissexto
'############################################################
    Bissexto = False
    If intAno Mod 4 = 0 Then
        If intAno Mod 100 = 0 Then
                If intAno Mod 400 = 0 Then
                    Bissexto = True
                End If
            Else
                Bissexto = True
        End If
    End If

End Function
Public Function CalcData(vDia As Integer, iMov As Integer, sData As Date) As Date
    '##############################################################################################
    '### 26/03/2012
    '### Calcula a data (vDia) para o mes corrente (sData) e antecipa sabados e domingos
    '### iMov = Integer que ira movimentar para frente ou para traz a data dependendo do dia da
    '###        Semana
    '###        Sendo: 0 - Não Movimenta
    '###               1 - Antecipa a data
    '###               2 - Postergar a data
    '###               3 - Ultimo dia do Mes
    '##############################################################################################
    Dim Mes         As String
    Dim ano         As String
    Dim DataF       As Date 'Data Final
    Dim Semana      As String
    Mes = Format(sData, "MM")
    ano = Format(sData, "YYYY")
    
    'DataF = vDia & "/" & Mes & "/" & ano
    If vDia > 28 Then
        Select Case Mes
            Case "04", "06", "09", "11"
                If vDia > 30 Then vDia = "30"
            Case "02"
                If Bissexto(Val(ano)) = True Then
                        If vDia > 29 Then vDia = "29"
                    Else
                        If vDia > 28 Then vDia = "28"
                End If
            Case Else
                If vDia > 31 Then vDia = "31"
        End Select
    End If
    DataF = vDia & "/" & Mes & "/" & ano
    'Semana = LCase(Format(DataF, "Long Date"))
    'Semana = Mid(Semana, 1, InStr(Semana, ",") - 1)
    Semana = DatePart("w", DataF)
    
    'Movimenta a data
    Select Case iMov
        Case 0 'Nao Movimenta
            '
        Case 1 'Antecipa
            Select Case Semana
                Case 7 ' "sabado", "sábado"
                    DataF = DataF - 1
                Case 1 ' "domingo"
                    DataF = DataF - 2
            End Select
        Case 2 'Posterga
            Select Case Semana
                Case 7 ' "sabado", "sábado"
                    DataF = DataF + 2 'vDia + 2
                Case 1 ' "domingo"
                    DataF = DataF + 1
            End Select
        Case 3 'Ultimo dia do Mes
            Select Case Mes
                Case "04", "06", "09", "11"
                    vDia = "30"
                Case "02"
                    If Bissexto(Val(ano)) = True Then
                        vDia = "29"
                    Else
                        vDia = "28"
                End If
            Case Else
                vDia = "31"
        End Select
        DataF = vDia & "/" & Mes & "/" & ano
    
    End Select
    'DataF = CalcData(vDia, iMov, sData)
    CalcData = DataF 'vDia & "/" & Mes & "/" & ano
End Function

Public Sub ExcluirFile(nmArquivo As String)
    '* 16/10/2012
    '* Copia da funcao de brECF
    
    
    'On Error GoTo TrtErro
    Dim caminho As String
    
    'Checa se existe a pasta
    'If Dir(exportFolder, vbDirectory) = "" Then
    '    MkDir exportFolder
    'End If
    
    'Caminho = exportFolder & "\" & nmArquivo
    caminho = nmArquivo
    'se o arquivo não existir então cria
    If Dir(caminho) <> "" Then
        Kill caminho
    End If
End Sub
