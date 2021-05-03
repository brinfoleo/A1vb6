Attribute VB_Name = "ModuloLicenca"
Option Explicit
'Validade da licenca
Public licencaValidade  As String
Public licencaChave     As String
Private arqLicencaLocal As String
'Informa se o sistema sera aberto em:
' true - producao
' false - demonstracao
Public licencaAtiva      As Boolean
Public Sub licenca()
    On Error GoTo ErrorTrt
    
    
    arqLicencaLocal = App.Path & "\licenca.txt"
    
    'Tenta ler o arquivo de licença
    If LerArquivoLicenca(arqLicencaLocal) = False Then
            
            '**** NAO ACHOU ARQUIVO DE LICENCA
            
'            'Testa a conexao web para continuar a licenca
            If OpenCnxLicenca = False Then
                    licencaAtiva = False
                    licencaValidade = ""
                    closeCnxLicenca
                Else
                    closeCnxLicenca
            End If
            If pegaLicencaWeb = True Then
                    If CDate(licencaValidade) >= Date And Numero_Serial = licencaChave Then
                            licencaAtiva = True
                        Else
                            licencaAtiva = False
                    End If
                Else
                    licencaAtiva = False
                    licencaValidade = ""
                    grvArqLicenca
            End If
        Else
            '**** ACHOU ARQUIVO DE LICENCA
            If CDate(licencaValidade) >= Date And InStr(licencaChave, Numero_Serial) = 1 Then
                        
                        'Licença valida
                        licencaAtiva = True
                        OpenCnxLicenca
                        updateAcessSystem
                        closeCnxLicenca
                    Else
                    
                        'Licenca Invalida
                        'updateAcessSystem
                        licencaAtiva = False
                        pegaLicencaWeb 'Revisa a licenca na web
                End If
    End If
    
    
    licencaAtiva = True
    licencaValidade = "01/01/2025"
    Exit Sub
ErrorTrt:
    licencaAtiva = True
    licencaValidade = "01/01/3000"
End Sub
    Private Function OpenCnxLicenca() As Boolean
    On Error GoTo erroConect
    Dim cnString    As String
    
    'Debug.Print "OpenCnxLicenca"
    
    OpenCnxLicenca = False
    cnString = "driver={MySQL ODBC 5.1 Driver};Server=dbmy0048.whservidor.com;port=3306;uid=brinfo2_1;pwd=k3bw82Le@;database=brinfo2_1"
    
    'usamos um cursor do lado do cliente pois os dados
    'serao acessados na maquina do cliente e nao de um servidor
    'dbLicenca.CursorLocation = adUseClient
    dbLicenca.Open cnString
    OpenCnxLicenca = True
    Exit Function
erroConect:
    RegLog "0", "0", "OpenCnxLicenca - Erro: " & Err.Number & " - " & Err.Description
    OpenCnxLicenca = False
    End Function
Private Function closeCnxLicenca() As Boolean
    On Error GoTo erroConect
    
    'Debug.Print "closeCnxLicenca"
    closeCnxLicenca = False
    dbLicenca.Close
    closeCnxLicenca = True
    Exit Function
erroConect:
    RegLog "0", "0", "closeCnxLicenca - Erro: " & Err.Number & " - " & Err.Description
    closeCnxLicenca = False
    End Function
Private Function pegaLicencaWeb() As Boolean
    On Error GoTo TrtErroDB
    Dim Rst As ADODB.Recordset
    
    Dim sSQL        As String
    
    'Fax conexao com a licenca
    If OpenCnxLicenca = False Then
            pegaLicencaWeb = False
            Exit Function
        Else
            closeCnxLicenca
    End If

'     12.01.2017 - atualizado pois multiplas empresas o cnpj é capturado apos o login
'    sSQL = "SELECT * FROM licenca WHERE sistema='A1' AND cnpj='" & _
'            PgDadosEmpresa(ID_Empresa).CNPJ & "' AND nSerie='" & Numero_Serial & "'"

    sSQL = "SELECT * FROM licenca WHERE sistema='A1' AND nSerie='" & Numero_Serial & "'"

    
    Set Rst = New ADODB.Recordset
    OpenCnxLicenca
    Rst.Open sSQL, dbLicenca
    
    If Rst.BOF And Rst.EOF Then
            'Nao achou informação do sistema, ou seja, primeiro acesso
            MsgBox "Sistema não licenciado" & vbCrLf & _
                    "Favor entrar em contato com WhatsApp: (+44) 7404 447 791 ou brinfo.leo@gmail.com." & vbCrLf & vbCrLf & _
                    "chave do sistema: " & Numero_Serial, vbCritical, "Licença"
                    pegaLicencaWeb = False
                    
            'Registra o acesso
                 
            licencaValidade = ""
            
            'Registra o sistema caso haja CNPJ
            If Trim(PgDadosEmpresa(ID_Empresa).CNPJ) <> "" Then
                sSQL = "INSERT INTO licenca (dtCadastro,sistema,cnpj,nSerie,Empresa,versao)" & _
                       " VALUES ('" & Now & "','A1','" & PgDadosEmpresa(ID_Empresa).CNPJ & "','" & _
                       Numero_Serial & "','" & PgDadosEmpresa(ID_Empresa).Nome & "','" & _
                       sVersao & "." & cVersao & "')"
                dbLicenca.Execute sSQL
            End If
            
                    
                    
        Else
            'Achou informacao,nao e primeiro acesso
            'Verifica se esta registrado
            Rst.MoveFirst
            pegaLicencaWeb = True
            If Rst.Fields("registrado") = 1 Then
                    'Registrado
                    licencaValidade = Rst.Fields("validade")
                    licencaChave = Rst.Fields("nSerie")
                    grvArqLicenca
                Else
                    'Nao Registrado
                    licencaValidade = "01/01/1900"
                    licencaChave = ""
            End If
            
            'Atualiza o ultimo acesso feito para consulta de licenca
            'dbLicenca.Close
           updateAcessSystem
    End If

    Rst.Close
    closeCnxLicenca
    Exit Function
TrtErroDB:
    RegLog "0", "0", "pegaLicencaWeb - Erro: " & Err.Number & " " & Err.Description
    pegaLicencaWeb = False
    Resume Next
End Function
Private Sub updateAcessSystem()
    On Error GoTo trtErrAcess
    Dim sSQL As String
    
   ' OpenCnxLicenca
'   12.01.2017 - Atualizado devido o cnpj ser registrado apos o login
'    sSQL = "UPDATE licenca " & _
'           "SET ultacesso='" & Now() & "',versao='" & sVersao & "." & cVersao & "'" & _
'                 " WHERE sistema='A1' AND cnpj='" & _
'                 PgDadosEmpresa(ID_Empresa).CNPJ & "' AND nSerie='" & Numero_Serial & "'"
'
    
    sSQL = "UPDATE licenca " & _
           "SET ultacesso='" & Now() & "',versao='" & sVersao & "." & cVersao & "'" & _
                 " WHERE sistema='A1' AND nSerie='" & Numero_Serial & "'"
    dbLicenca.Execute sSQL
    Exit Sub
    
trtErrAcess:
    RegLog "0", "0", "updateAcessSystem - Erro: " & Err.Number & " " & Err.Description & " - [sql: " & sSQL & "]"
    Resume Next
End Sub
Public Sub grvArqLicenca()
    ExcluirFile arqLicencaLocal
    grvFile arqLicencaLocal, sBase64Encode("validade=" & IIf(Len(Trim(licencaValidade)) > 0, licencaValidade, "2000-01-01"))
    grvFile arqLicencaLocal, sBase64Encode("licencachave=" & Trim(Numero_Serial))
End Sub
Private Function LerArquivoLicenca(caminho As String) As Boolean
    '*********************************************************************************
    '*** Data: 27/01/2012
    '*** Obj.: Ler o arquivo para configuracao do sistema
    '*********************************************************************************
    On Error GoTo trtErroConexao
    Dim F           As Long
    Dim linha       As String
    Dim Campo       As String 'Le o campo do arquivo de config antes do sinal de =
    Dim parametro   As String 'Recebe as instrucoes a serem armazenadas nas variaveis
    'caminho - Armazena o caminho e nome do arquivo
    
    
    If Dir(caminho) = "" Then
        LerArquivoLicenca = False
        Exit Function
    End If
    
    
    F = FreeFile
    Open caminho For Input As F   'abre o arquivo texto
    
    Do While Not EOF(F)
        Line Input #F, linha 'lê uma linha do arquivo texto
        
        linha = Trim(sBase64Decode(linha))
        If Left(linha, 1) <> "#" And Trim(linha) <> "" Then 'Nao executa as funcoes abaixo pois esta linha é comentario

            'Separa os campos
            Campo = Trim(LCase(Mid(linha, 1, InStr(linha, "=") - 1)))
            parametro = Trim(Mid(linha, InStr(linha, "=") + 1, Len(linha)))
        
            'Testa se falta alguma instrucao
            'If Trim(campo) = "" Or Trim(parametro) = "" Then
            '    LerArquivoINI = False
            '    Close #f
            '    Exit Function
            'End If
        
            Select Case Campo
                Case "validade"
                    licencaValidade = Trim(parametro)
                Case "licencachave"
                    licencaChave = parametro
                
            End Select
        End If
    Loop
    Close #F
    LerArquivoLicenca = True
    Exit Function
trtErroConexao:
    MsgBox Err.Description, vbInformation, Err.Number
    LerArquivoLicenca = False
    'Resume Next
End Function
