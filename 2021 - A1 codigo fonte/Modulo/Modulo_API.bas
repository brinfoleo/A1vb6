Attribute VB_Name = "Modulo_FuncaoShell"

' Declara��es da API do Windows necess�rias para as fun��es OpenProcess, WaitForSingleObject e CloseHandle
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' Constantes necess�rias para as fun��es da API do Windows
Private Const SYNCHRONIZE = &H100000
Private Const WAIT_TIMEOUT = &H102


Public Function ExecutarAplicacaoExterna(CaminhoAplicacao As String) As Boolean
    Dim Retorno As Long
    Dim ProcessoID As Long

    ' Caminho completo para o execut�vel da aplica��o externa
    'Dim CaminhoAplicacao As String
    'CaminhoAplicacao = "C:\Caminho\Para\Sua\Aplicacao.exe"

    ' Inicia a aplica��o externa e obt�m o ID do processo
    
    ProcessoID = Shell(CaminhoAplicacao, vbNormalFocus)

    ' Verifica se a aplica��o foi iniciada com sucesso
    If ProcessoID > 0 Then
        ' Aguarda a finaliza��o da aplica��o externa
        Do While IsProcessRunning(ProcessoID)
            DoEvents ' Permite que outros eventos sejam processados enquanto aguarda
            ' Voc� pode adicionar um pequeno atraso aqui, se necess�rio
            ' Sleep 100 ' Aguarda 100 milissegundos (requer declara��o da API Sleep)
        Loop

        ' A aplica��o externa foi finalizada
        'MsgBox "A aplica��o externa foi finalizada."
        ExecutarAplicacaoExterna = True
    Else
        'MsgBox "Falha ao iniciar a aplica��o externa."
        ExecutarAplicacaoExterna = False
    End If
End Function

' Fun��o para verificar se um processo ainda est� em execu��o
Private Function IsProcessRunning(ByVal ProcessoID As Long) As Boolean
    Dim hProcess As Long
    Dim Retorno As Long

    ' Abre um handle para o processo
    hProcess = OpenProcess(SYNCHRONIZE, False, ProcessoID)

    ' Verifica se o handle foi aberto com sucesso
    If hProcess > 0 Then
        ' Aguarda um curto per�odo para verificar o status do processo
        Retorno = WaitForSingleObject(hProcess, 0)

        ' Verifica se o processo ainda est� em execu��o
        If Retorno = WAIT_TIMEOUT Then
            IsProcessRunning = True
        Else
            IsProcessRunning = False
        End If

        ' Fecha o handle do processo
        CloseHandle hProcess
    Else
        ' Falha ao abrir o handle do processo (o processo pode n�o existir)
        IsProcessRunning = False
    End If
End Function
