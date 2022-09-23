Attribute VB_Name = "ModuloDatabase"
Option Explicit

Public BD           As New ADODB.Connection
Public dbLicenca    As New ADODB.Connection

Public nmDatabase   As String
Public srv_IP       As String
Public srv_Porta    As String
Public dbUsu        As String
Public dbSenha      As String
'**************************************
'*** Variaveis para analise de Trafego
Dim hIni As String
Dim hFin As String
'**************************************
Private Sub LerArquivoConfig()
    Dim F               As Long
    Dim linha           As String
    Dim localArq        As String
    Dim Campo           As String
    Dim parametro       As String
    
    
    
    localArq = App.Path & "\" & App.EXEName & ".cfg"
    
    
    F = FreeFile
        
    Open localArq For Input As F    'abre o arquivo texto
    
    Do While Not EOF(F)
        Line Input #F, linha 'lê uma linha do arquivo texto
        
        linha = Trim(linha)
        If Left(linha, 1) <> "#" And Trim(linha) <> "" Then 'Nao executa as funcoes abaixo pois esta linha é comentario
            'Separa os campos
            Campo = Trim(LCase(Mid(linha, 1, InStr(linha, "=") - 1)))
            parametro = Trim(Mid(linha, InStr(linha, "=") + 1, Len(linha)))
        
        Select Case LCase(Campo)
            Case "ip"
                srv_IP = parametro
            Case "port"
                srv_Porta = parametro
            Case "usu"
                dbUsu = parametro
            Case "senha"
                dbSenha = parametro
            Case "nmdatabase"
                    nmDatabase = parametro
        End Select
        End If
    Loop
    Close #F
    '###############################################

        

    
    
End Sub

Public Sub BloquearTabela(nmTab As String)
    '*** Data: 16/08/2011
    '*** Bolqueia a tabela para escrita
    'On Error Resume Next
    BD.Execute "LOCK TABLE " & nmTab & " WRITE"
End Sub
Public Sub DesbloquearTabela(nmTab As String)
    '*** Data: 16/08/2011
    '*** Desbolqueia a tabela para escrita
    'On Error Resume Next
    BD.Execute "UNLOCK TABLES" ' & nmTab & " WRITE"
End Sub

Public Function cnDatabase() As Boolean
    On Error GoTo TrtErroDB
    Dim cnString    As String
    
    'TestarConexaoServidor
    
    If TestarConexaoServidor = False Then
        If MsgBox("Deseja alterar as configurações de acesso ao servidor?", vbYesNo, App.EXEName) = vbYes Then
            formConexao.Show 1
        End If
        cnDatabase = False
        'End
        Exit Function
    End If
    
    'Access | cnString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & SistemPath & "\Dados.mdb;"
    'MySQL  |
    
    cnString = "driver={MySQL ODBC 5.1 Driver};Server=" & srv_IP & ";port=" & srv_Porta & ";uid=" & dbUsu & ";pwd=" & dbSenha & ";database=" & nmDatabase 'a1_database"


    BD.CursorLocation = adUseClient  'usamos um cursor do lado do cliente pois os dados 'serao acessados na maquina do cliente e nao de um servidor

    BD.Open cnString
    cnDatabase = True

    
    Exit Function
TrtErroDB:
    If BD.Errors.Item(0).Number = "-2147467259" Then
        If MsgBox("Servidor sem a base de dados. Deseja que o sistema crie uma nova base de dados!", vbCritical + vbYesNo, "Acesso a base de dados") = vbYes Then
            If CriaBancoDados = True Then
                MsgBox "Base de dados criada. Favor reiniciar o aplicativo.", vbInformation
                
                    
            End If
         
        End If
        
            If MsgBox("Deseja reconfigurar a conexão?", vbYesNo, App.EXEName) = vbYes Then
                formConexao.Show 1
                'FinalizandoSistema
                
            End If
        cnDatabase = False
        'End
        Exit Function
    End If
    MsgBox "Erro ao ABRIR a base de dados.                                   " & _
           vbCrLf & vbCrLf & _
           "Erro n.: " & Err.Number & _
           vbCrLf & vbCrLf & _
           "Descrição: " & Err.Description & _
           vbCrLf
    cnDatabase = False
    'End
     Exit Function
 End Function
'*************************************************************************
'*************************************************************************


Private Function CriaBancoDados() As Boolean
' AGREGAR ESTA FUNCAO PARA QUE O SISTEMA POSSA GERAR SEU PROPRIO DATABASE
    Dim pConexao As ADODB.Connection
    Dim strTemp As String

    On Error GoTo trata_erro

    On Error Resume Next
    Set pConexao = New ADODB.Connection
    pConexao.Open "DRIVER={MySQL ODBC 5.1 Driver};user=" & dbUsu & ";password=" & dbSenha & ";server=" & srv_IP & ";port=" & srv_Porta
'
    If pConexao.State = 1 Then
        strTemp = nmDatabase '"A1_Database"
            If Trim$(strTemp) <> "" Then
                pConexao.Execute "Create Database " & Trim$(strTemp), , adExecuteNoRecords
                
            End If
            CriaBancoDados = True
        Else
            MsgBox "Não foi possível estabelecer comunicação com o Servidor. Verifique seu Host e sua chave/Senha.", vbCritical, "Impossível criar Banco de dados."
            CriaBancoDados = False
        End If
    Exit Function
trata_erro:
    MsgBox Err.Number & " " & Err.Description, vbCritical, "Erro durante a criação do banco de dados."
End Function
Public Function RegistroAlterar(ByVal sTabela As String, ByVal vDados As Variant, nmReg As Integer, _
                               Optional ByVal sCondicao As String) As Boolean
    
'nmReg - Numero de registros passados
'vDados é um array de 3 dimensões, aonde a 1a. é o campo, a 2a. é o valor a ser gravado
'e a 3a. é o tipo de dado. "S" para string, "N" para Numero, "I" para inteiro,
'"D" para Data, "T" para tempo(time) e "V" para variável (qq. coisa)
    On Error GoTo TrataErro
    Dim sSQL       As String
    Dim sValues    As String
    Dim i          As Integer
    '   AnaliseTrafego True
    'nmReg = nmReg - 1
    'sFields = ",,"
    sTabela = LCase(sTabela)
    sValues = "DtHr = '" & Now() & "', Id_Empresa = '" & ID_Empresa & "', UsuID = " & ID_Usuario & ","
    '########### CUIDADO ###########################################
    If ValidarAplicativo(sTabela) = False Then Exit Function
    '###############################################################
    For i = 0 To nmReg 'UBound(vDados)
        
        If vDados(i)(2) = "S" Then
                vDados(i)(1) = Replace(vDados(i)(1), "'", "´") ' Subistitui  apostrofo por acento agudo evitando erro na string
                vDados(i)(1) = Replace(vDados(i)(1), "\", "\\")  'Subistitui \ por \\ para que o MySQL entenda que é uma barra
                sValues = sValues & vDados(i)(0) & "=" & IIf(Trim(vDados(i)(1)) = "", "Null,", "'" & vDados(i)(1) & "'" & ",")
                
            ElseIf vDados(i)(2) = "N" Then
                If vDados(i)(1) <> "" Then
                        sValues = sValues & vDados(i)(0) & "=" & vDados(i)(1) & ","
         
                    Else
                        sValues = sValues & vDados(i)(0) & "=" & 0 & ","
         
                End If
      
            ElseIf vDados(i)(2) = "I" Then
                If vDados(i)(1) <> "" Then
                        sValues = sValues & vDados(i)(0) & "=" & CStr(Val(vDados(i)(1))) & ","
         
                    Else
                        sValues = sValues & vDados(i)(0) & "=" & "0" & ","
         
                End If
      
            ElseIf vDados(i)(2) = "D" Then
                Dim sDt As String
                sDt = vDados(i)(1)
                'sDt = IIf(sDt = "", "Null", "")
                sDt = IIf(sDt = "", "Null", "'" & Format(sDt, "yyyy-mm-dd") & "'")
                
                sValues = sValues & vDados(i)(0) & "=" & sDt & "," 'vDados(i)(1) & "," ConverteData(vDados(i)(1)) & ","
      
            ElseIf vDados(i)(2) = "T" Then
                sValues = sValues & vDados(i)(1) & "," 'ConverteTempo(vDados(i)(1)) & ","
      
            ElseIf vDados(i)(2) = "V" Then
                sValues = sValues & vDados(i)(0) & "=" & vDados(i)(1) & ","
      
        End If
   
    Next i
   
    sValues = Left(sValues, Len(sValues) - 1)
    
    'If sCondicao = "" Then
    '        sSQL = "UPDATE " & sTabela & " " & _
                   "SET " & sValues
       
    '    Else
            sSQL = "UPDATE " & LCase(sTabela) & " " & _
                   "SET " & sValues & " " & _
                   "WHERE Id_Empresa = " & ID_Empresa & IIf(Trim(sCondicao) = "", "", " AND " & sCondicao)
          
    'End If
     
    
    
    'Dim Rst As ADODB.Recordset

    'Set Rst = New ADODB.Recordset
    
    'Rst.Open sSQL, BD
    BD.Execute sSQL
    RegistroAlterar = True
    'AnaliseTrafego False, "(A)" & sSQL
    Exit Function

TrataErro:
    RegLog "", Err.Number, Err.Description & " - SQL:[" & sSQL & "]"
    'MsgBox "Erro ao ALTERAR registro.                                   " & _
           vbCrLf & vbCrLf & _
           "Erro n.: " & Err.Number & _
           vbCrLf & vbCrLf & _
           "Descrição: " & Err.Description & _
           vbCrLf & vbCrLf & _
           "SQL: " & sSQL
    'MsgBox "Não foi possível alterar o registro em " & sTabela & "!" & vbCrLf & "ERRO: " & _
       CStr(Err.Number) & " - " & Err.Description, vbCritical & vbOKOnly, "ERRO!"
    RegistroAlterar = False

End Function
Public Function RegistroExcluir(sTabela As String, sFiltro As String) As Boolean
    On Error GoTo TrtErro
    Dim sSQL As String
    sTabela = LCase(sTabela)
    If Trim(sFiltro) = "" Then
            ''Usar com Access
            'sSQL = "DELETE * FROM " & sTabela
            sSQL = "DELETE FROM " & sTabela & " WHERE ID_Empresa = " & ID_Empresa
            
        Else
            'Usar com Access
            'sSQL = "DELETE * FROM " & sTabela & " " & _
                "WHERE " & sFiltro
            sSQL = "DELETE FROM " & sTabela & " " & _
                "WHERE ID_Empresa = " & ID_Empresa & " AND " & sFiltro
    End If
    sSQL = LCase(sSQL)
    BD.Execute sSQL
    RegistroExcluir = True
    'RegLog "0", "0", "Exclusao: " & sSQL
    Exit Function
TrtErro:
    RegLog "", Err.Number, Err.Description & " - SQL:[" & sSQL & "]"
    'RegLog "", Err.Number, Err.Description
    'MsgBox "Erro ao EXCLUIR registro.                                   " & _
           vbCrLf & vbCrLf & _
           "Erro n.: " & Err.Number & _
           vbCrLf & _
           "Descrição: " & Err.Description & _
           vbCrLf & vbCrLf & _
           "SQL: " & sSQL
        'Debug.Print Err.Description
    RegistroExcluir = False
End Function
Public Function RegistroBuscar(sSQL As String) As ADODB.Recordset
    On Error GoTo TrtErro
    Dim Rst As ADODB.Recordset
    
    'AnaliseTrafego True
    sSQL = LCase(sSQL)
    
    
    Set Rst = New ADODB.Recordset
    Rst.Open sSQL, BD ', adOpenForwardOnly
 
    Set RegistroBuscar = Rst
    'AnaliseTrafego False, "(B)" & sSQL
    Exit Function
TrtErro:
    Dim nErro As String
    Dim dErro As String
    
    nErro = Err.Number
    dErro = Err.Description
    RegLog "", nErro, dErro & " - SQL:[" & sSQL & "]"
    'RegLog "", nErro, dErro
      MsgBox "Erro ao BUSCAR registro.                                   " & _
           vbCrLf & vbCrLf & _
           "Erro n.: " & nErro & _
           vbCrLf & vbCrLf & _
           "Descrição: " & dErro & _
           vbCrLf & vbCrLf & _
           "SQL: " & sSQL
    
     Set RegistroBuscar = Nothing

End Function


Public Function RegistroIncluir(sTabela As String, vDados As Variant, nmReg As Integer) As Long
    'nmReg - Numero de registros passados
    'sTabela - Nome da Tabela
    'vDados - Contem os dados como (1) Campos, (2) Dados , (3) Tipo de dados
    '          Obs. no (3) os dados podem ser s - string, i - inerger, n - numero,
    '          d - data, t - tempo e v - variante
    '
    On Error GoTo TrataErro
    Dim sSQL       As String
    Dim sFields    As String
    Dim sValues    As String
    Dim i          As Integer
    'AnaliseTrafego True
    
    'nmReg = nmReg - 1
    sTabela = LCase(sTabela)
    
    sFields = "DtHr,Id_Empresa,UsuID,"
    sValues = "'" & Now() & "','" & ID_Empresa & "'," & ID_Usuario & ","
    '########### CUIDADO ######################################
    If ValidarAplicativo(sTabela) = False Then Exit Function
    '##########################################################
    For i = 0 To nmReg 'UBound(vDados)
        sFields = sFields & vDados(i)(0) & ","
        
        If vDados(i)(2) = "S" Then
                'vDados(i)(1) = Replace(vDados(i)(1), "'", "´") ' Subistitui  apostrofo por acento agudo evitando erro na string
                vDados(i)(1) = Replace(vDados(i)(1), "\", "\\")  'Subistitui \ por \\ para que o MySQL entenda que é uma barra
                vDados(i)(1) = rc(CStr(vDados(i)(1)))
                sValues = sValues & IIf(Trim(vDados(i)(1)) = "", "Null,", "'" & vDados(i)(1) & "'" & ",")
      
             ElseIf vDados(i)(2) = "N" Then
                If vDados(i)(1) <> "" Then
                        sValues = sValues & vDados(i)(1) & ","
                    Else
                        sValues = sValues & 0 & ","
                End If
      
            ElseIf vDados(i)(2) = "I" Then
                If vDados(i)(1) <> "" Then
                        sValues = sValues & CStr(Val(vDados(i)(1))) & ","
                    Else
                        sValues = sValues & "0" & ","
                End If
      
            ElseIf vDados(i)(2) = "D" Then
                Dim sDt As String
                sDt = vDados(i)(1)
                sDt = IIf(sDt = "", "Null", "'" & Format(sDt, "yyyy-mm-dd") & "'")
                
                sValues = sValues & sDt & "," 'ConverteData(vDados(I)(1)) & ","
      
            ElseIf vDados(i)(2) = "T" Then
                sValues = sValues & vDados(i)(1) & "," 'ConverteTempo(vDados(I)(1)) & ","
      
            ElseIf vDados(i)(2) = "V" Then
                sValues = sValues & vDados(i)(0) & "=" & vDados(i)(1) & ","
      
        End If
   
    Next i
   
    
    sFields = Left(sFields, Len(sFields) - 1)
    sValues = Left(sValues, Len(sValues) - 1)

    sSQL = "INSERT INTO " & LCase(sTabela) & _
           " (" & sFields & ") " & _
           " VALUES(" & sValues & ") "
     
     
    
    'sSQL = LCase(sSQL)
    
    Dim Rst        As ADODB.Recordset

    Set Rst = New ADODB.Recordset
    
    Rst.Open sSQL, BD
    Set Rst = RegistroBuscar("Select @@IDENTITY as Codigo")
    RegistroIncluir = Rst.Fields("codigo")
    'Rst.MoveLast
    'RegistroIncluir = Rst.Fields("ID")
    'AnaliseTrafego False, "(I)" & sSQL
    Exit Function

TrataErro:
    RegLog "", Err.Number, Err.Description & " - SQL:[" & sSQL & "]"
  ' Debug.Print Err.Description
   ' RegLog "", Err.Number, Err.Description
    'MsgBox "Erro ao INCLUIR registro.                                   " & _
           vbCrLf & vbCrLf & _
           "Erro n.: " & Err.Number & _
           vbCrLf & vbCrLf & _
           "Descrição: " & Err.Description & _
           vbCrLf & vbCrLf & _
           "SQL: " & sSQL
    'MsgBox "Não foi possível gravar o registro em " & sTabela & "!" & vbCrLf & "ERRO: " & _
            CStr(Err.Number) & " - " & Err.Description, vbCritical & vbOKOnly, "ERRO!"
    RegistroIncluir = 0
    Exit Function
End Function




Public Function TestarConexaoServidor() As Boolean
    On Error GoTo trtErroConexao
    
    Dim tstConexao      As New ADODB.Connection

    Dim cnString        As String
    
    LerArquivoConfig
    cnString = "driver={MySQL ODBC 5.1 Driver};Server=" & srv_IP & ";port=" & srv_Porta & ";uid=" & dbUsu & ";pwd=" & dbSenha
    'cnstring = "driver={MySQL ODBC 5.1 Driver};Server=" & srv_IP & ";port=" & srv_Porta & ";uid=root;pwd=newsys;database=a1_database"

    'tstConexao.CursorLocation = adUseClient  'usamos um cursor do lado do cliente pois os dados 'serao acessados na maquina do cliente e nao de um servidor

    tstConexao.Open cnString
    If tstConexao.State = 1 Then
        TestarConexaoServidor = True
        tstConexao.Close
    Else
        TestarConexaoServidor = False
        tstConexao.Close
    End If
    Exit Function
trtErroConexao:
    MsgBox Err.Description, vbInformation, Err.Number
    TestarConexaoServidor = False
End Function

Public Function RepararBD(Optional MySQLOpcao As String = "EXTENDED") As Boolean
    On Error GoTo errHandler
    Dim rsRepair As New ADODB.Recordset
    With rsRepair
        .Open "SHOW TABLE STATUS;", BD
        Do While Not (.BOF Or .EOF)
            BD.BeginTrans
            BD.Execute "REPAIR TABLE " & .Fields.Item("Name").Value & " " & MySQLOpcao & ";", , 0
            'Debug.Print .Fields.Item("name").Value
            BD.CommitTrans
            .MoveNext
        Loop
    End With
    rsRepair.Close: Set rsRepair = Nothing
    RepararBD = True
    Exit Function
errHandler:
    rsRepair.Close: Set rsRepair = Nothing

'As opções podem ser:
'
'Quick - Repara apenas os índices;
'Extended - Criar a fileira do índice pela fileira, em vez de criar um
'           índice de cada vez com a classificação. Isto pode ser melhor do que
'           classificando em chaves "fixed-length" se você tiver as chaves longas do
'           char () que comprimem;
'Use .FRM - Ideal para tabelas corrompidas. Nesta modalidade MySQL recriará
'           a tabela, usando a informação dos arquivos ".frm".
End Function

Public Function ValidarAplicativo(sTabela As String) As Boolean
    '#####################################################################
    '### 30/01/2012                                                    ###
    '### Função criada para limitar a quantidadede registros inclusos  ###
    '### no sistema.                                                   ###
    '### Objetivo: tornar o sistema com duas funções Teste e Completa  ###
    '#####################################################################
    On Error GoTo TrtErroValidador
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim MaxRecord   As Long
    If licencaAtiva = True Then
            '***********************************************
            '***           SISTEMA PARA PRODUCAO         ***
            '***********************************************
'            Select Case LCase(sTabela)
'                Case "clientes"
'                    MaxRecord = 500
'                Case "fornecedores"
'                    MaxRecord = 100
'                Case "transportadoras"
'                    MaxRecord = 5000
'                Case "faturamentonfe"
'                    MaxRecord = 10000
'                Case "financeirocontasprcadastro"
'                    MaxRecord = 20000
'                Case "faturamentonfesendmail"
'                    MaxRecord = 20000
'                Case "estoquekardex"
'                    MaxRecord = 40000
'                Case Else
'                    MaxRecord = 50000
'            End Select
            ValidarAplicativo = True
        Else
            '***********************************************
            '***        SISTEMA PARA DEMONSTRACAO        ***
            '***         COM LIMITE NAS TABELAS          ***
            '***********************************************
            Select Case LCase(sTabela)
                Case "clientes"
                    MaxRecord = 3
                Case "fornecedores"
                    MaxRecord = 3
                Case "transportadoras"
                    MaxRecord = 3
                Case "faturamentonfe"
                    MaxRecord = 2
                Case "financeirocontasprcadastro"
                    MaxRecord = 10
                Case "faturamentonfesendmail"
                    MaxRecord = 4
                Case "estoquekardex"
                    MaxRecord = 10
'                Case "usugerenciadorformularios"
'                    MaxRecord = 350
                Case Else
                    MaxRecord = 20000
            End Select
        '***********************************************
            
        sSQL = "SELECT * FROM " & sTabela
        Set Rst = RegistroBuscar(sSQL)
        
        If Rst.BOF And Rst.EOF Then
                ValidarAplicativo = True
            Else
                Rst.MoveLast
                If Rst.RecordCount >= MaxRecord Then
                        ValidarAplicativo = False
                    Else
                        ValidarAplicativo = True
                End If
        End If
        If ValidarAplicativo = False Then
    
            MsgBox "Warning: Table '" & sTabela & "' is marked as crashed and should be repaired query:" & vbCrLf & _
                   "Can’t open file: " & sTabela & ".myi" & vbCrLf & _
                   "Table-Corruption Issues", vbCritical, App.EXEName
        End If
        
    End If
   
    Exit Function
TrtErroValidador:
    ValidarAplicativo = False
    Resume Next
End Function
Private Sub AnaliseTrafego(start As Boolean, Optional sSQL As String)
    '*
    '* 26.10.2012
    '* Registra o tempo que leva uma instrução de gravação no servidor
    '*
    '* Removido em 31.10.2012 para nao haver delay na gravacao
    '*
    Exit Sub
    On Error Resume Next
    Dim sTexto As String
    Dim nFile As String
    nFile = App.Path & "\dbtime-" & Format(Date, "YYYY-MM-DD") & ".txt"
    Select Case start
        Case True
            hIni = Time
        Case False
            hFin = Time
            
             sTexto = MDIFormA1.wsMain.LocalIP & " [" & Date & " " & hIni & "|" & hFin & "] " & sSQL
            grvFile nFile, sTexto
    End Select
End Sub

