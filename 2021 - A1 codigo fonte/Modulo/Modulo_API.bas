Attribute VB_Name = "Modulo_FuncaoShell"

' Declarações da API do Windows necessárias para as funções OpenProcess, WaitForSingleObject e CloseHandle
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' Constantes necessárias para as funções da API do Windows
Private Const SYNCHRONIZE = &H100000
Private Const WAIT_TIMEOUT = &H102


Public Function ExecutarAplicacaoExterna(CaminhoAplicacao As String) As Boolean
    Dim Retorno As Long
    Dim ProcessoID As Long

    ' Caminho completo para o executável da aplicação externa
    'Dim CaminhoAplicacao As String
    'CaminhoAplicacao = "C:\Caminho\Para\Sua\Aplicacao.exe"

    ' Inicia a aplicação externa e obtém o ID do processo
    
    ProcessoID = Shell(CaminhoAplicacao, vbNormalFocus)

    ' Verifica se a aplicação foi iniciada com sucesso
    If ProcessoID > 0 Then
        ' Aguarda a finalização da aplicação externa
        Do While IsProcessRunning(ProcessoID)
            DoEvents ' Permite que outros eventos sejam processados enquanto aguarda
            ' Você pode adicionar um pequeno atraso aqui, se necessário
            ' Sleep 100 ' Aguarda 100 milissegundos (requer declaração da API Sleep)
        Loop

        ' A aplicação externa foi finalizada
        'MsgBox "A aplicação externa foi finalizada."
        ExecutarAplicacaoExterna = True
    Else
        'MsgBox "Falha ao iniciar a aplicação externa."
        ExecutarAplicacaoExterna = False
    End If
End Function

' Função para verificar se um processo ainda está em execução
Private Function IsProcessRunning(ByVal ProcessoID As Long) As Boolean
    Dim hProcess As Long
    Dim Retorno As Long

    ' Abre um handle para o processo
    hProcess = OpenProcess(SYNCHRONIZE, False, ProcessoID)

    ' Verifica se o handle foi aberto com sucesso
    If hProcess > 0 Then
        ' Aguarda um curto período para verificar o status do processo
        Retorno = WaitForSingleObject(hProcess, 0)

        ' Verifica se o processo ainda está em execução
        If Retorno = WAIT_TIMEOUT Then
            IsProcessRunning = True
        Else
            IsProcessRunning = False
        End If

        ' Fecha o handle do processo
        CloseHandle hProcess
    Else
        ' Falha ao abrir o handle do processo (o processo pode não existir)
        IsProcessRunning = False
    End If
End Function
