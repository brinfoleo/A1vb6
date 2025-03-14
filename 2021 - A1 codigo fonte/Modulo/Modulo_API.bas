Attribute VB_Name = "Modulo_API"
Option Explicit
Public Sub mGET()

    Dim http As Object
    Dim url As String
    Dim response As String
    
    url = "https://jsonplaceholder.typicode.com/posts" ' Substitua pela sua URL
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    'Configurando cabecalhos:
   ' http.setRequestHeader "Authorization", "Bearer token_aqui"
    
    http.Open "GET", url, False ' GET, URL, S�ncrono (True para ass�ncrono)
    http.Send
    
    response = http.responseText
    
    If http.status = 200 Then
        ' Requisi��o bem-sucedida
        Debug.Print response ' Exibe a resposta
        ' Ou processe a resposta como necess�rio
    Else
        ' Requisi��o falhou
        Debug.Print "Erro: " & http.status & " - " & http.StatusText
    End If
    
    Set http = Nothing
End Sub
Public Sub mPOST()
    Dim http As Object
    Dim url As String
    Dim jsonData As String
    
    url = "https://www.exemplo.com/api/post"
    jsonData = "{""chave"": ""valor"", ""outro"": 123}"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send jsonData
    
    If http.status = 200 Then
        Debug.Print http.responseText
    Else
        Debug.Print "Erro: " & http.status
    End If
    
    Set http = Nothing
End Sub
Public Sub cURL()
Dim cmd As String
Dim retVal As Long
    
    cmd = "curl https://jsonplaceholder.typicode.com/posts" ' Ajuste o caminho para curl.exe se necess�rio
    
    retVal = Shell(cmd, vbNormalFocus) ' Executa o comando
    ' A sa�da do curl ser� exibida no console.
    ' Capturar a sa�da diretamente � dif�cil.
    
    ' Para capturar a sa�da, voc� precisaria redirecionar a sa�da para um arquivo,
    ' e ent�o ler o conte�do do arquivo no VB6. Isso adiciona complexidade.
End Sub
