Attribute VB_Name = "ModuloDataLake"
Public dataLakeDb      As New ADODB.Connection
Public Function cnDatabaseLake() As Boolean
    On Error GoTo TrtErroDB
    
    Dim cnString        As String
    'MySQL  |
    
    cnString = "driver={MySQL ODBC 5.1 Driver};Server=" & srv_IP & ";port=" & srv_Porta & ";uid=" & dbUsu & ";pwd=" & dbSenha & ";database=datalake"


    dataLakeDb.CursorLocation = adUseClient  'usamos um cursor do lado do cliente pois os dados 'serao acessados na maquina do cliente e nao de um servidor

    dataLakeDb.Open cnString

    
    Exit Function
TrtErroDB:
     RegLog 0, 0, "cnDatabaseLake: " & Err.Number & " - " & Err.Description
     Exit Function
 End Function



Public Function dataLakeInputNFe(numNFe As String, chvA As String, strData As String) As Long
    
    On Error GoTo TrataErro
    Dim sSQL       As String
    Dim sFields    As String
    Dim sValues    As String

    'Open database
    cnDatabaseLake
    

    sTabela = LCase("nfe")
    
    sFields = "DtHr,Id_Empresa,UsuID"
    sValues = "'" & Now() & "','" & ID_Empresa & "'," & ID_Usuario
    
    
    sFields = sFields & ", tipo, numNFe, chv, txtNFe"
    sValues = sValues & ",1,'" & numNFe & "','" & chvA & "' , '" & strData & "'"
    

    sSQL = "INSERT INTO " & LCase(sTabela) & _
           " (" & sFields & ") " & _
           " VALUES(" & sValues & ") "
     
     
    
    'sSQL = LCase(sSQL)
    
    Dim Rst        As ADODB.Recordset

    Set Rst = New ADODB.Recordset
    
    Rst.Open sSQL, dataLakeDb
    
    'close database
    dataLakeDb.Close
    Exit Function

TrataErro:
    MsgBox "function dataLakeInputNFe: " & Err.Description, vbInformation, Err.Number
    RegLogDataBase 0, 0, 0, "function dataLakeInputNFe: " & Err.Number & "-" & Err.Description
    
    Exit Function
End Function

Public Function RegistroIncluirDataLake(sTabela As String, vDados As Variant, nmReg As Integer) As Long
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
    cnDatabaseLake
    Dim Rst        As ADODB.Recordset

    Set Rst = New ADODB.Recordset
    
    Rst.Open sSQL, dataLakeDb
    'Set Rst = RegistroBuscar("Select @@IDENTITY as Codigo")
    'RegistroIncluir = Rst.Fields("codigo")
    'Rst.MoveLast
    RegistroIncluirDataLake = 1  'Rst.Fields("ID")
    'AnaliseTrafego False, "(I)" & sSQL
    dataLakeDb.Close
    Exit Function

TrataErro:
    RegLog "", Err.Number, Err.Description & " - SQL:[" & sSQL & "]"
  
    RegistroIncluirDataLake = 0
    Exit Function
End Function






