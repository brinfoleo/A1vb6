Attribute VB_Name = "ModuloPrincipal"
 Option Explicit

Public ID_Empresa       As Integer 'Empresa corrente
Public ID_Usuario       As Integer 'Usuario corrente
Public SistemPath       As String  'Local onde esta o programa
Public ID_Deposito      As Integer 'Deposito principal
Public cDecMoeda        As Integer 'Casas decimais usadas para o Moeda Corrente
Public cDecQtd          As Integer 'Casas decimais usadas para Quantidade
Public sVersao          As String  'Versao do sistema
Public cVersao          As String  'Compilacao da Versao do sistema


Public Sub Main()
    
    VersaoNFe = "4.00"
    sVersao = App.Major & "." & App.Minor
    cVersao = App.Revision
    SistemPath = App.Path
    
      
    
    'with formSplash
    With frmSplash
        
        .CarregarFormulario "Abrindo o sistema..."
        
        .CarregarFormulario "Lendo arquivos de inicialização..."
'        If LerArquivoINI = False Then
'            MsgBox "Erro ao localizar os paramentros do sistema!", vbCritical, App.EXEName
'            End
'        End If
        
        'Escrevendo arquivo de versao
        wFileVersion
        
        .CarregarFormulario "Conectando ao banco de dados..."
        
        If cnDatabase = False Then End
        '#############################################################
        'formdbDBGrid.Show 1
        'Exit Sub
        '#############################################################
        
        'Checando Licença
        .CarregarFormulario "Lendo licença de uso do sistema..."
        licenca
        
        .CarregarFormulario "Processando linha de comando..."
        
        If Trim(Command) <> "" Then
            Select Case LCase(Command)
                Case "backup"
                    Call formBackup.IniciarBackup
                    FinalizandoSistema
                    Exit Sub
                Case "reparar"
                    RepararBD
                Case "update"
                    updateSistema
                Case "?"
                    MsgBox "Digite:" & vbCrLf & _
                    "\A1.exe backup : para backup automatico" & vbCrLf & _
                    "\A1.exe update : realiza atualização no sistema" & vbCrLf & _
                    "\A1.exe reparar : para reparar o banco de dados", vbInformation, "Aviso"
                    FinalizandoSistema
            End Select
        End If
        
        
           
        
'        cDecMoeda = PgDadosConfig.cDecMoeda
'        cDecQtd = PgDadosConfig.cDecQtd
'
        .CarregarFormulario "Analisando data do equipamento..."
        If AnaliseDataEquipamento = False Then
            MsgBox "A data do seu equipamento pode estar errada favor verificar antes de continuar!", vbInformation, App.EXEName
            FinalizandoSistema
        End If
        
        
        'Validando usuario e empresa
        .CarregarFormulario "Entrada de Empresa/Usuário..."
        'If formLogin.EfetuarLogin = False Then
        If frmlogin.EfetuarLogin = False Then
            FinalizandoSistema
            Exit Sub
         End If
    
         cDecMoeda = PgDadosConfig.cDecMoeda
        cDecQtd = PgDadosConfig.cDecQtd
        
        
        
        
        '###############################################
        .CarregarFormulario "Implementando contas fixas..."
        cobrLancamentoAutomaticoContasFixas
        '###############################################
        
        .FecharFormulario
        
    End With
    
    '*****************************************************
    '            TESTE DO EFD
    '
    'MnFiscal_EFD 0, "01/02/2014", "28/02/2014", True
    '
    'FinalizandoSistema
    ' End
    '************************************************
    
    
    
    
'    'If formLogin.EfetuarLogin = False Then
'    If frmlogin.EfetuarLogin = False Then
'        FinalizandoSistema
'    End If
   
    MonitoramentoConexao Conectar
    
    
    MDIFormA1.Show
    
    
    
'************************************************************************************
'    If licencaAtiva = False Then
'        MDIFormA1.Caption = MDIFormA1.Caption & "    .:: Licença Expirada ::."
'    End If
'************************************************************************************
End Sub
Public Sub FinalizandoSistema()
    On Error GoTo TratErro
    MonitoramentoConexao Desconectar
    BD.Close
    End
    Exit Sub
TratErro:
    MsgBox "Erro ao fechar o sistema"
End Sub
Public Function LerArquivoINI() As Boolean
    '*********************************************************************************
    '*** Data: 27/01/2012
    '*** Obj.: Ler o arquivo para configuracao do sistema
    '*********************************************************************************
    On Error GoTo trtErroConexao
    Dim F           As Long
    Dim linha       As String
    Dim Campo       As String 'Le o campo do arquivo de config antes do sinal de =
    Dim parametro   As String 'Recebe as instrucoes a serem armazenadas nas variaveis
    Dim caminho     As String 'Armazena o caminho e nome do arquivo
    
    caminho = App.Path & "\" & App.EXEName & ".INI"
    If Dir(caminho) = "" Then
        LerArquivoINI = False
        Exit Function
    End If
    
    
    F = FreeFile
    Open caminho For Input As F   'abre o arquivo texto
    
    Do While Not EOF(F)
        Line Input #F, linha 'lê uma linha do arquivo texto
        
        linha = Trim(linha)
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
                Case "idempresa"
                    ID_Empresa = parametro
                Case "iddeposito"
                    '16.05.2017 - Substituido para cada
                    'empresa ter seu deposito padrao
                    'ID_Deposito = parametro
                'Case "nmdatabase"
                '    nmDatabase = parametro
                'Case "server"
                '    srv_Nome = parametro
                'Case "porta"
                '    srv_Porta = parametro
                'Case "ecf"
                '    nmECF = parametro
                'Case "user"
                '    nmUser = parametro
                'Case "pwd"
                '    sPWD = parametro
                'Case "cdecmoeda"
                '    cDecMoeda = parametro
                'Case "cdecqtd"
                '    cDecQtd = parametro
                'Case "logo"
                '    LocalLogo = parametro
                'Case "expcftxt"
                '    ExpCFtxt = parametro
                'Case "exportfolder"
                '    exportFolder = parametro
                'Case "tpdb"
                '    tipoDataBase = parametro
                'Case "pdv_fontsize"
                '    PDVFontSize = parametro
                'Case "backimage"
                '    BackImage = parametro
                
            End Select
        End If
    Loop
    Close #F
    LerArquivoINI = True
    Exit Function
trtErroConexao:
    MsgBox Err.Description, vbInformation, Err.Number
    LerArquivoINI = False
    'Resume Next
End Function


Private Function AnaliseDataEquipamento() As Boolean
    If PgDadosConfig.DtUltMov > Date And (PgDadosConfig.DtUltMov + 7) < Date Then
        AnaliseDataEquipamento = False
        Exit Function
    End If
    AnaliseDataEquipamento = True
    Dim vReg(1) As Variant
    
    vReg(0) = Array("DtUltMov", Date, "D")
    RegistroAlterar "Configuracoes", vReg, 0, "id=1"
    
End Function
Public Sub updateSistema()
    '#
    '# Leonardo Aquino - 27/09/2012
    '#
    '# Cria uma tabela com nome de updateSistema
    '# para registrar as ultimas atualizações
    '# Caso o sistema esteja com uma versao
    '# diferente da registrada iniciará o processo
    '# de atualização, iniciando com
    '# o fechamento dos micros da rede
    '#
    '#
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim vReg(5) As Variant
    Dim cReg    As Integer
    sSQL = "CREATE TABLE IF NOT EXISTS updatesistema " & _
              "(Id INT(11) NOT NULL AUTO_INCREMENT," & _
               "Id_Empresa INT default Null," & _
               "UsuID VARCHAR(10) default Null," & _
               "DtHr VARCHAR(20) default Null," & _
               "vAtual VARCHAR(30) default Null," & _
               " PRIMARY KEY (Id))"
    BD.Execute sSQL
    
    
    'sSQL = "SELECT * FROM updatesistema ORDER BY vAtual"
    'Set Rst = RegistroBuscar(sSQL)
    'If Rst.BOF And Rst.EOF Then
    '    Else
    'End If
    'FinalizandoSistema
    'End
    
End Sub
Private Sub install()

    formUsuGerenciador.MontarBaseDeDados
    formEmpresas.MontarBaseDeDados
    formConfiguracoes.MontarBaseDeDados
    
End Sub
Private Sub wFileVersion()
    Dim sTexto As String
    sTexto = rc(RS(sVersao & cVersao))
    If Dir(App.Path & "\versao.txt", vbArchive) <> "" Then
        Kill App.Path & "\versao.txt"
    End If
    grvFile App.Path & "\versao.txt", sTexto
End Sub

Public Sub escreverINIVazio()
    Dim sTexto As String
    Dim caminho As String
    caminho = App.Path & "\" & App.EXEName & ".INI"


    If Dir(caminho, vbArchive) <> "" Then
        Kill caminho
    End If
    sTexto = "idEmpresa=0"
    grvFile caminho, sTexto
End Sub
